//! `zlsx dbx` — the Databricks family: push / pull / genie.
//!
//! Talks to a Databricks workspace over plain REST (std.http, no SDK,
//! no third-party deps), so the static binary is the whole client:
//!
//!   zlsx dbx push local.xlsx /Volumes/cat/schema/vol/file.xlsx [--overwrite]
//!   zlsx dbx pull /Volumes/cat/schema/vol/file.xlsx local.xlsx
//!   zlsx dbx genie "question" [--space ID] [--timeout-secs N]
//!
//! Auth comes from the environment — `DATABRICKS_HOST` and
//! `DATABRICKS_TOKEN` (a PAT), plus `GENIE_SPACE_ID` as the default
//! space for `genie` — matching the `.env` convention the
//! integrations/databricks/ scripts use.
//!
//! The transfer commands apply the library's contract to the wire:
//! push OPENS the local workbook before uploading and pull PARSES the
//! downloaded bytes (via `Book.openBuffer`) before the atomic rename —
//! a corrupt workbook is refused on the way up and never replaces a
//! good file on the way down.
//!
//! Output is NDJSON records on stdout, same as every other `zlsx`
//! sub-command; errors are one human line on stderr.
//!
//! Exit codes: 0 success · 1 bad arguments / missing environment ·
//! 2 HTTP or API failure · 3 workbook verification refused · 5 local
//! file I/O error.

const std = @import("std");
const xlsx = @import("zlsx");
const pkg = @import("zlsx_pkg");

const Writer = std.Io.Writer;

const usage_text =
    \\usage: zlsx dbx <command> [args]
    \\
    \\  push <local.xlsx> </Volumes/...>   upload a workbook to a UC Volume
    \\        --overwrite                  replace the remote file if present
    \\  pull </Volumes/...> <local.xlsx>   download a workbook from a UC Volume
    \\  genie "<question>"                 ask a Genie space, print SQL + rows
    \\        --space <id>                 space id (default: $GENIE_SPACE_ID)
    \\        --timeout-secs <n>           poll budget (default 120)
    \\
    \\environment: DATABRICKS_HOST, DATABRICKS_TOKEN, GENIE_SPACE_ID
    \\
;

/// Poll cadence for the Genie conversation loop. 2s matches the
/// latency profile of a serverless warehouse without hammering the API.
const poll_interval_s: i64 = 2;
const default_timeout_s: u32 = 120;

/// Files-API responses are tiny JSON or empty; Genie result sets are
/// bounded by the space's row limit. 64 MiB caps the pull body instead:
/// a workbook bigger than that should not transit through memory here.
/// Known limit: the cap is checked after the fetch has buffered the
/// body, so a hostile endpoint could allocate past it first — accepted,
/// because the endpoint is the caller's own configured workspace and a
/// mid-stream cap needs a bounded Writer std.http doesn't offer yet.
const max_response_bytes: usize = 64 * 1024 * 1024;

pub fn run(
    alloc: std.mem.Allocator,
    io: std.Io,
    environ: std.process.Environ,
    argv: []const []const u8,
    out: *Writer,
    err_w: *Writer,
) !u8 {
    if (argv.len == 0) {
        try err_w.writeAll(usage_text);
        try err_w.flush();
        return 1;
    }
    const cmd = argv[0];
    if (std.mem.eql(u8, cmd, "-h") or std.mem.eql(u8, cmd, "--help")) {
        try out.writeAll(usage_text);
        try out.flush();
        return 0;
    }

    const rc: u8 = blk: {
        if (std.mem.eql(u8, cmd, "push")) break :blk try runPush(alloc, io, environ, argv[1..], out, err_w);
        if (std.mem.eql(u8, cmd, "pull")) break :blk try runPull(alloc, io, environ, argv[1..], out, err_w);
        if (std.mem.eql(u8, cmd, "genie")) break :blk try runGenie(alloc, io, environ, argv[1..], out, err_w);
        try err_w.print("zlsx dbx: unknown command '{s}'\n\n", .{cmd});
        try err_w.writeAll(usage_text);
        try err_w.flush();
        return 1;
    };
    try out.flush();
    try err_w.flush();
    return rc;
}

// ─── auth / environment ───────────────────────────────────────────────

const Auth = struct {
    host: []u8, // normalized: https://…, no trailing slash; owned
    token: []u8, // owned

    fn deinit(self: *Auth, alloc: std.mem.Allocator) void {
        alloc.free(self.host);
        alloc.free(self.token);
    }
};

fn loadAuth(alloc: std.mem.Allocator, environ: std.process.Environ, err_w: *Writer) !?Auth {
    const raw_host = environ.getAlloc(alloc, "DATABRICKS_HOST") catch |e| switch (e) {
        error.EnvironmentVariableMissing => {
            try err_w.writeAll("zlsx dbx: DATABRICKS_HOST is not set (source your .env)\n");
            return null;
        },
        else => return e,
    };
    defer alloc.free(raw_host);
    const token = environ.getAlloc(alloc, "DATABRICKS_TOKEN") catch |e| switch (e) {
        error.EnvironmentVariableMissing => {
            try err_w.writeAll("zlsx dbx: DATABRICKS_TOKEN is not set (source your .env)\n");
            return null;
        },
        else => return e,
    };
    errdefer alloc.free(token);
    const host = try normalizeHost(alloc, raw_host);
    return .{ .host = host, .token = token };
}

/// `dbc-x.cloud.databricks.com/` and `https://dbc-x.cloud.databricks.com`
/// both normalize to the latter. Rejects nothing — a wrong host shows up
/// as a connect error with the URL in the message, which is more useful
/// than second-guessing the workspace naming scheme here.
fn normalizeHost(alloc: std.mem.Allocator, raw: []const u8) ![]u8 {
    var h = std.mem.trim(u8, raw, " \t\r\n");
    while (h.len > 0 and h[h.len - 1] == '/') h = h[0 .. h.len - 1];
    if (std.mem.startsWith(u8, h, "https://") or std.mem.startsWith(u8, h, "http://")) {
        return alloc.dupe(u8, h);
    }
    return std.fmt.allocPrint(alloc, "https://{s}", .{h});
}

// ─── URL building ─────────────────────────────────────────────────────

const PathError = error{NotAVolumePath};

/// Files-API URL for a UC Volume path. Accepts `/Volumes/...` or the
/// `dbfs:/Volumes/...` spelling the databricks CLI prints, refuses
/// anything else: the Files API only serves Volumes, and a typo'd
/// path should fail here, not as a remote 404 after an upload attempt.
fn filesApiUrl(alloc: std.mem.Allocator, host: []const u8, volume_path: []const u8) (PathError || std.mem.Allocator.Error)![]u8 {
    var p = volume_path;
    if (std.mem.startsWith(u8, p, "dbfs:")) p = p["dbfs:".len..];
    if (!std.mem.startsWith(u8, p, "/Volumes/")) return error.NotAVolumePath;

    var aw: Writer.Allocating = .init(alloc);
    defer aw.deinit();
    const w = &aw.writer;
    w.writeAll(host) catch return error.OutOfMemory;
    w.writeAll("/api/2.0/fs/files") catch return error.OutOfMemory;
    percentEncodePath(w, p) catch return error.OutOfMemory;
    return alloc.dupe(u8, aw.written());
}

/// Percent-encode for URL embedding. RFC 3986 unreserved characters
/// pass through; everything else — spaces, unicode, `?`, `#` — is
/// %XX-encoded so it cannot be misparsed as query or fragment.
/// `keep_slash` distinguishes a path (separators intact) from a single
/// segment, where `/` must be encoded too.
fn percentEncode(w: *Writer, s: []const u8, comptime keep_slash: bool) !void {
    for (s) |c| {
        const keep = (c >= 'A' and c <= 'Z') or (c >= 'a' and c <= 'z') or
            (c >= '0' and c <= '9') or c == '-' or c == '_' or c == '.' or
            c == '~' or (keep_slash and c == '/');
        if (keep) {
            try w.writeByte(c);
        } else {
            try w.print("%{X:0>2}", .{c});
        }
    }
}

fn percentEncodePath(w: *Writer, path: []const u8) !void {
    return percentEncode(w, path, true);
}

/// Segment-encode a value that lands inside a URL path. The genie ids
/// (conversation/message/attachment) come back from the SERVER — a
/// hostile or proxied response must not be able to splice `../` or a
/// query string into the next authenticated request's URL.
fn encodeSegmentAlloc(alloc: std.mem.Allocator, s: []const u8) ![]u8 {
    var aw: Writer.Allocating = .init(alloc);
    defer aw.deinit();
    percentEncode(&aw.writer, s, false) catch return error.OutOfMemory;
    return alloc.dupe(u8, aw.written());
}

// ─── HTTP ─────────────────────────────────────────────────────────────

const HttpResult = struct {
    status: u16,
    body: []u8, // owned by caller's allocator

    fn deinit(self: *HttpResult, alloc: std.mem.Allocator) void {
        alloc.free(self.body);
    }
};

fn httpRequest(
    alloc: std.mem.Allocator,
    io: std.Io,
    method: std.http.Method,
    url: []const u8,
    token: []const u8,
    payload: ?[]const u8,
    content_type: ?[]const u8,
) !HttpResult {
    var client: std.http.Client = .{ .allocator = alloc, .io = io };
    defer client.deinit();

    const bearer = try std.fmt.allocPrint(alloc, "Bearer {s}", .{token});
    defer alloc.free(bearer);

    var aw: Writer.Allocating = .init(alloc);
    defer aw.deinit();

    const res = try client.fetch(.{
        .location = .{ .url = url },
        .method = method,
        .payload = payload,
        .response_writer = &aw.writer,
        // One-shot CLI: force `connection: close`. With keep-alive the
        // Files API's 204-No-Content response leaves the connection
        // parked and the response drain blocks until the load
        // balancer's idle timeout (observed as a 2-minute hang on
        // push). EOF-terminated bodies also make error-body reads
        // finite. The extra handshake per genie poll is noise next to
        // warehouse latency.
        .keep_alive = false,
        .headers = .{
            .authorization = .{ .override = bearer },
            .content_type = if (content_type) |ct| .{ .override = ct } else .default,
        },
    });
    if (aw.written().len > max_response_bytes) return error.ResponseTooLarge;
    return .{
        .status = @intFromEnum(res.status),
        .body = try alloc.dupe(u8, aw.written()),
    };
}

/// Copy `s` into `buf` (truncating) with every control byte replaced
/// by '.', so hostile response bytes cannot inject terminal escapes or
/// carriage-return tricks through a diagnostic line.
fn sanitizedCopy(buf: []u8, s: []const u8) []const u8 {
    const n = @min(buf.len, s.len);
    for (s[0..n], buf[0..n]) |c, *d| {
        d.* = if (c < 0x20 or c == 0x7f) '.' else c;
    }
    return buf[0..n];
}

/// One stderr line for a non-2xx API response: status plus the useful
/// prefix of the body (Databricks errors are small JSON objects).
fn reportApiError(err_w: *Writer, what: []const u8, status: u16, body: []const u8) !void {
    var excerpt_buf: [300]u8 = undefined;
    const excerpt = sanitizedCopy(&excerpt_buf, body);
    try err_w.print("zlsx dbx: {s} failed: HTTP {d} {s}\n", .{ what, status, excerpt });
    if (status == 401 or status == 403) {
        try err_w.writeAll("zlsx dbx: check DATABRICKS_TOKEN (expired or wrong workspace?)\n");
    }
}

// ─── push ─────────────────────────────────────────────────────────────

fn runPush(
    alloc: std.mem.Allocator,
    io: std.Io,
    environ: std.process.Environ,
    argv: []const []const u8,
    out: *Writer,
    err_w: *Writer,
) !u8 {
    var local: ?[]const u8 = null;
    var remote: ?[]const u8 = null;
    var overwrite = false;
    for (argv) |a| {
        if (std.mem.eql(u8, a, "--overwrite")) {
            overwrite = true;
        } else if (std.mem.startsWith(u8, a, "--")) {
            try err_w.print("zlsx dbx push: unknown flag '{s}'\n", .{a});
            return 1;
        } else if (local == null) {
            local = a;
        } else if (remote == null) {
            remote = a;
        } else {
            try err_w.print("zlsx dbx push: unexpected argument '{s}'\n", .{a});
            return 1;
        }
    }
    const local_path = local orelse {
        try err_w.writeAll("zlsx dbx push: missing <local.xlsx> and </Volumes/...>\n");
        return 1;
    };
    const remote_path = remote orelse {
        try err_w.writeAll("zlsx dbx push: missing </Volumes/...> destination\n");
        return 1;
    };

    var auth = (try loadAuth(alloc, environ, err_w)) orelse return 1;
    defer auth.deinit(alloc);

    const bytes = std.Io.Dir.cwd().readFileAlloc(io, local_path, alloc, .limited(max_response_bytes)) catch |e| {
        try err_w.print("zlsx dbx push: cannot read {s}: {s}\n", .{ local_path, @errorName(e) });
        return 5;
    };
    defer alloc.free(bytes);

    // Correct-or-refuse at the boundary: a file that does not parse as
    // a workbook is refused before any network traffic.
    const sheet_count = verifyWorkbook(alloc, io, bytes) catch |e| {
        try err_w.print(
            "zlsx dbx push: {s} is not a readable workbook ({s}); refusing to upload\n",
            .{ local_path, @errorName(e) },
        );
        return 3;
    };

    const base_url = filesApiUrl(alloc, auth.host, remote_path) catch |e| switch (e) {
        error.NotAVolumePath => {
            try err_w.print("zlsx dbx push: '{s}' is not a /Volumes/ path\n", .{remote_path});
            return 1;
        },
        else => return e,
    };
    defer alloc.free(base_url);
    const url = try std.fmt.allocPrint(
        alloc,
        "{s}?overwrite={s}",
        .{ base_url, if (overwrite) "true" else "false" },
    );
    defer alloc.free(url);

    var res = httpRequest(alloc, io, .PUT, url, auth.token, bytes, "application/octet-stream") catch |e| {
        try err_w.print("zlsx dbx push: request failed ({s})\n", .{@errorName(e)});
        return 2;
    };
    defer res.deinit(alloc);

    if (res.status == 409) {
        try err_w.print("zlsx dbx push: remote exists: {s} (pass --overwrite)\n", .{remote_path});
        return 2;
    }
    if (res.status < 200 or res.status >= 300) {
        try reportApiError(err_w, "push", res.status, res.body);
        return 2;
    }

    try out.print("{f}\n", .{std.json.fmt(.{
        .kind = "dbx_push",
        .local = local_path,
        .remote = remote_path,
        .bytes = bytes.len,
        .sheets = sheet_count,
        .overwrite = overwrite,
    }, .{})});
    return 0;
}

/// Parse `bytes` as a workbook; return the sheet count. Uses the
/// buffer-based open (v0.6.0's C-ABI ask) so no temp file is involved.
fn verifyWorkbook(alloc: std.mem.Allocator, io: std.Io, bytes: []const u8) !usize {
    var book = try xlsx.Book.openBuffer(alloc, io, bytes);
    defer book.deinit();
    return book.sheets.len;
}

// ─── pull ─────────────────────────────────────────────────────────────

fn runPull(
    alloc: std.mem.Allocator,
    io: std.Io,
    environ: std.process.Environ,
    argv: []const []const u8,
    out: *Writer,
    err_w: *Writer,
) !u8 {
    if (argv.len != 2) {
        try err_w.writeAll("zlsx dbx pull: expected </Volumes/...> <local.xlsx>\n");
        return 1;
    }
    const remote_path = argv[0];
    const local_path = argv[1];

    var auth = (try loadAuth(alloc, environ, err_w)) orelse return 1;
    defer auth.deinit(alloc);

    const url = filesApiUrl(alloc, auth.host, remote_path) catch |e| switch (e) {
        error.NotAVolumePath => {
            try err_w.print("zlsx dbx pull: '{s}' is not a /Volumes/ path\n", .{remote_path});
            return 1;
        },
        else => return e,
    };
    defer alloc.free(url);

    var res = httpRequest(alloc, io, .GET, url, auth.token, null, null) catch |e| {
        try err_w.print("zlsx dbx pull: request failed ({s})\n", .{@errorName(e)});
        return 2;
    };
    defer res.deinit(alloc);

    if (res.status != 200) {
        try reportApiError(err_w, "pull", res.status, res.body);
        return 2;
    }

    // Refuse to publish bytes that don't parse — the temp file is
    // dropped and any existing local file stays untouched.
    const sheet_count = verifyWorkbook(alloc, io, res.body) catch |e| {
        try err_w.print(
            "zlsx dbx pull: remote {s} is not a readable workbook ({s}); not writing {s}\n",
            .{ remote_path, @errorName(e), local_path },
        );
        return 3;
    };

    var write_buf: [64 * 1024]u8 = undefined;
    var af = pkg.AtomicFile.init(io, local_path, &write_buf) catch |e| {
        try err_w.print("zlsx dbx pull: cannot write {s}: {s}\n", .{ local_path, @errorName(e) });
        return 5;
    };
    defer af.deinit();
    af.file_writer.interface.writeAll(res.body) catch |e| {
        try err_w.print("zlsx dbx pull: write failed: {s}\n", .{@errorName(e)});
        return 5;
    };
    af.finish() catch |e| {
        try err_w.print("zlsx dbx pull: finalize failed: {s}\n", .{@errorName(e)});
        return 5;
    };

    try out.print("{f}\n", .{std.json.fmt(.{
        .kind = "dbx_pull",
        .remote = remote_path,
        .local = local_path,
        .bytes = res.body.len,
        .sheets = sheet_count,
    }, .{})});
    return 0;
}

// ─── genie ────────────────────────────────────────────────────────────

const StartConversation = struct {
    conversation_id: []const u8,
    message_id: []const u8,
};

const Attachment = struct {
    attachment_id: ?[]const u8 = null,
    text: ?struct { content: ?[]const u8 = null } = null,
    query: ?struct { query: ?[]const u8 = null } = null,
};

const MessagePoll = struct {
    status: []const u8 = "",
    attachments: ?[]Attachment = null,
};

const QueryResult = struct {
    statement_response: ?struct {
        manifest: ?struct {
            schema: ?struct {
                columns: ?[]struct { name: []const u8 } = null,
            } = null,
        } = null,
        result: ?struct {
            data_array: ?[][]?[]const u8 = null,
        } = null,
    } = null,
};

fn runGenie(
    alloc: std.mem.Allocator,
    io: std.Io,
    environ: std.process.Environ,
    argv: []const []const u8,
    out: *Writer,
    err_w: *Writer,
) !u8 {
    var question: ?[]const u8 = null;
    var space_flag: ?[]const u8 = null;
    var timeout_s: u32 = default_timeout_s;
    var i: usize = 0;
    while (i < argv.len) : (i += 1) {
        const a = argv[i];
        if (std.mem.eql(u8, a, "--space")) {
            i += 1;
            if (i >= argv.len) {
                try err_w.writeAll("zlsx dbx genie: --space needs a value\n");
                return 1;
            }
            space_flag = argv[i];
        } else if (std.mem.eql(u8, a, "--timeout-secs")) {
            i += 1;
            if (i >= argv.len) {
                try err_w.writeAll("zlsx dbx genie: --timeout-secs needs a value\n");
                return 1;
            }
            timeout_s = std.fmt.parseInt(u32, argv[i], 10) catch {
                try err_w.print("zlsx dbx genie: bad --timeout-secs '{s}'\n", .{argv[i]});
                return 1;
            };
        } else if (std.mem.startsWith(u8, a, "--")) {
            try err_w.print("zlsx dbx genie: unknown flag '{s}'\n", .{a});
            return 1;
        } else if (question == null) {
            question = a;
        } else {
            try err_w.print("zlsx dbx genie: unexpected argument '{s}'\n", .{a});
            return 1;
        }
    }
    const q = question orelse {
        try err_w.writeAll("zlsx dbx genie: missing \"<question>\"\n");
        return 1;
    };

    var auth = (try loadAuth(alloc, environ, err_w)) orelse return 1;
    defer auth.deinit(alloc);

    const space_owned: ?[]u8 = if (space_flag == null)
        environ.getAlloc(alloc, "GENIE_SPACE_ID") catch |e| switch (e) {
            error.EnvironmentVariableMissing => null,
            else => return e,
        }
    else
        null;
    defer if (space_owned) |s| alloc.free(s);
    const space = space_flag orelse (space_owned orelse {
        try err_w.writeAll("zlsx dbx genie: no space id (--space or GENIE_SPACE_ID)\n");
        return 1;
    });

    // start-conversation
    const space_enc = try encodeSegmentAlloc(alloc, space);
    defer alloc.free(space_enc);
    const start_url = try std.fmt.allocPrint(
        alloc,
        "{s}/api/2.0/genie/spaces/{s}/start-conversation",
        .{ auth.host, space_enc },
    );
    defer alloc.free(start_url);
    const body = try std.fmt.allocPrint(alloc, "{f}", .{std.json.fmt(.{ .content = q }, .{})});
    defer alloc.free(body);

    var start_res = httpRequest(alloc, io, .POST, start_url, auth.token, body, "application/json") catch |e| {
        try err_w.print("zlsx dbx genie: request failed ({s})\n", .{@errorName(e)});
        return 2;
    };
    defer start_res.deinit(alloc);
    if (start_res.status < 200 or start_res.status >= 300) {
        try reportApiError(err_w, "genie start-conversation", start_res.status, start_res.body);
        return 2;
    }
    const start = std.json.parseFromSlice(StartConversation, alloc, start_res.body, .{
        .ignore_unknown_fields = true,
        .allocate = .alloc_always,
    }) catch {
        try err_w.writeAll("zlsx dbx genie: unrecognized start-conversation response\n");
        return 2;
    };
    defer start.deinit();

    // poll until terminal status or timeout
    const conv_enc = try encodeSegmentAlloc(alloc, start.value.conversation_id);
    defer alloc.free(conv_enc);
    const msg_enc = try encodeSegmentAlloc(alloc, start.value.message_id);
    defer alloc.free(msg_enc);
    const msg_url = try std.fmt.allocPrint(
        alloc,
        "{s}/api/2.0/genie/spaces/{s}/conversations/{s}/messages/{s}",
        .{ auth.host, space_enc, conv_enc, msg_enc },
    );
    defer alloc.free(msg_url);

    // Wall-clock deadline, not an attempt count: each poll spends
    // request latency on top of the sleep, so counting attempts would
    // silently stretch a 120s budget toward 4 minutes. (Per-request
    // deadlines are still absent — std.http.Client.fetch has no
    // timeout hook in 0.16 — so a fully stalled TCP connection can
    // exceed the budget by one request's worth of hang.)
    const poll_start = std.Io.Timestamp.now(io, .awake);
    const poll = while (true) {
        var poll_res = httpRequest(alloc, io, .GET, msg_url, auth.token, null, null) catch |e| {
            try err_w.print("zlsx dbx genie: poll failed ({s})\n", .{@errorName(e)});
            return 2;
        };
        defer poll_res.deinit(alloc);
        if (poll_res.status != 200) {
            try reportApiError(err_w, "genie poll", poll_res.status, poll_res.body);
            return 2;
        }
        // alloc_always: the parsed value outlives poll_res.body (it is
        // broken out of the loop while the defer frees the body), so
        // every string must be copied into the parse arena, not
        // borrowed from the response buffer.
        const parsed = std.json.parseFromSlice(MessagePoll, alloc, poll_res.body, .{
            .ignore_unknown_fields = true,
            .allocate = .alloc_always,
        }) catch {
            try err_w.writeAll("zlsx dbx genie: unrecognized message response\n");
            return 2;
        };
        if (isTerminalStatus(parsed.value.status)) break parsed;
        parsed.deinit();
        const elapsed = poll_start.durationTo(std.Io.Timestamp.now(io, .awake));
        if (elapsed.nanoseconds >= @as(i96, timeout_s) * 1_000_000_000) {
            try err_w.print("zlsx dbx genie: timed out after {d}s (message still running)\n", .{timeout_s});
            return 2;
        }
        io.sleep(.fromSeconds(poll_interval_s), .awake) catch {};
    };
    defer poll.deinit();

    if (!std.mem.eql(u8, poll.value.status, "COMPLETED")) {
        var status_buf: [64]u8 = undefined;
        try err_w.print("zlsx dbx genie: message ended {s}\n", .{sanitizedCopy(&status_buf, poll.value.status)});
        return 2;
    }

    try out.print("{f}\n", .{std.json.fmt(.{
        .kind = "genie_status",
        .status = poll.value.status,
        .conversation_id = start.value.conversation_id,
    }, .{})});

    for (poll.value.attachments orelse &.{}) |att| {
        if (att.text) |txt| if (txt.content) |content| {
            try out.print("{f}\n", .{std.json.fmt(.{ .kind = "genie_text", .text = content }, .{})});
        };
        if (att.query) |qy| {
            if (qy.query) |sql| {
                try out.print("{f}\n", .{std.json.fmt(.{ .kind = "genie_sql", .sql = sql }, .{})});
            }
            const att_id = att.attachment_id orelse continue;
            const att_enc = try encodeSegmentAlloc(alloc, att_id);
            defer alloc.free(att_enc);
            const qr_url = try std.fmt.allocPrint(
                alloc,
                "{s}/attachments/{s}/query-result",
                .{ msg_url, att_enc },
            );
            defer alloc.free(qr_url);
            var qr_res = httpRequest(alloc, io, .GET, qr_url, auth.token, null, null) catch |e| {
                try err_w.print("zlsx dbx genie: query-result failed ({s})\n", .{@errorName(e)});
                return 2;
            };
            defer qr_res.deinit(alloc);
            if (qr_res.status != 200) {
                try reportApiError(err_w, "genie query-result", qr_res.status, qr_res.body);
                return 2;
            }
            emitQueryResult(alloc, out, qr_res.body, err_w) catch |e| switch (e) {
                error.UnrecognizedResponse => return 2,
                else => return e,
            };
        }
    }
    return 0;
}

fn isTerminalStatus(status: []const u8) bool {
    return std.mem.eql(u8, status, "COMPLETED") or
        std.mem.eql(u8, status, "FAILED") or
        std.mem.eql(u8, status, "CANCELLED") or
        std.mem.eql(u8, status, "QUERY_RESULT_EXPIRED");
}

/// Emit `genie_columns` + one `genie_row` per result row. Cells arrive
/// as strings-or-null from the statement API; they pass through as-is.
/// A body that doesn't parse is `error.UnrecognizedResponse` — the
/// caller maps it to exit 2, same as every other malformed API reply.
fn emitQueryResult(alloc: std.mem.Allocator, out: *Writer, body: []const u8, err_w: *Writer) !void {
    const parsed = std.json.parseFromSlice(QueryResult, alloc, body, .{
        .ignore_unknown_fields = true,
    }) catch {
        try err_w.writeAll("zlsx dbx genie: unrecognized query-result response\n");
        return error.UnrecognizedResponse;
    };
    defer parsed.deinit();
    const sr = parsed.value.statement_response orelse return;

    if (sr.manifest) |m| if (m.schema) |s| if (s.columns) |cols| {
        var names = try alloc.alloc([]const u8, cols.len);
        defer alloc.free(names);
        for (cols, 0..) |c, idx| names[idx] = c.name;
        try out.print("{f}\n", .{std.json.fmt(.{ .kind = "genie_columns", .columns = names }, .{})});
    };
    if (sr.result) |r| if (r.data_array) |rows| for (rows) |row| {
        try out.print("{f}\n", .{std.json.fmt(.{ .kind = "genie_row", .cells = row }, .{})});
    };
}

// ─── tests ────────────────────────────────────────────────────────────

const t = std.testing;

test "normalizeHost adds scheme and strips trailing slash" {
    const cases = [_][2][]const u8{
        .{ "dbc-x.cloud.databricks.com", "https://dbc-x.cloud.databricks.com" },
        .{ "dbc-x.cloud.databricks.com/", "https://dbc-x.cloud.databricks.com" },
        .{ "https://dbc-x.cloud.databricks.com", "https://dbc-x.cloud.databricks.com" },
        .{ "https://dbc-x.cloud.databricks.com//", "https://dbc-x.cloud.databricks.com" },
        .{ " https://h \n", "https://h" },
    };
    for (cases) |c| {
        const got = try normalizeHost(t.allocator, c[0]);
        defer t.allocator.free(got);
        try t.expectEqualStrings(c[1], got);
    }
}

test "filesApiUrl accepts Volumes paths and dbfs: spelling" {
    const url = try filesApiUrl(t.allocator, "https://h", "/Volumes/c/s/v/f.xlsx");
    defer t.allocator.free(url);
    try t.expectEqualStrings("https://h/api/2.0/fs/files/Volumes/c/s/v/f.xlsx", url);

    const url2 = try filesApiUrl(t.allocator, "https://h", "dbfs:/Volumes/c/s/v/f.xlsx");
    defer t.allocator.free(url2);
    try t.expectEqualStrings("https://h/api/2.0/fs/files/Volumes/c/s/v/f.xlsx", url2);
}

test "filesApiUrl refuses non-Volume paths" {
    try t.expectError(error.NotAVolumePath, filesApiUrl(t.allocator, "https://h", "/tmp/f.xlsx"));
    try t.expectError(error.NotAVolumePath, filesApiUrl(t.allocator, "https://h", "Volumes/c/s/v/f.xlsx"));
}

test "filesApiUrl percent-encodes spaces and unicode, keeps separators" {
    const url = try filesApiUrl(t.allocator, "https://h", "/Volumes/c/s/v/my file é.xlsx");
    defer t.allocator.free(url);
    try t.expectEqualStrings(
        "https://h/api/2.0/fs/files/Volumes/c/s/v/my%20file%20%C3%A9.xlsx",
        url,
    );
}

test "genie start body escapes the question" {
    const body = try std.fmt.allocPrint(
        t.allocator,
        "{f}",
        .{std.json.fmt(.{ .content = "what is \"total\"?\n" }, .{})},
    );
    defer t.allocator.free(body);
    try t.expectEqualStrings("{\"content\":\"what is \\\"total\\\"?\\n\"}", body);
}

test "start-conversation response parses" {
    const fixture =
        \\{"conversation_id":"c1","message_id":"m1","extra":{"ignored":true}}
    ;
    const parsed = try std.json.parseFromSlice(StartConversation, t.allocator, fixture, .{
        .ignore_unknown_fields = true,
    });
    defer parsed.deinit();
    try t.expectEqualStrings("c1", parsed.value.conversation_id);
    try t.expectEqualStrings("m1", parsed.value.message_id);
}

test "message poll parses attachments of both shapes" {
    const fixture =
        \\{"status":"COMPLETED","attachments":[
        \\ {"attachment_id":"a1","text":{"content":"hello"}},
        \\ {"attachment_id":"a2","query":{"query":"SELECT 1","description":"d"}}
        \\]}
    ;
    const parsed = try std.json.parseFromSlice(MessagePoll, t.allocator, fixture, .{
        .ignore_unknown_fields = true,
    });
    defer parsed.deinit();
    try t.expect(isTerminalStatus(parsed.value.status));
    const atts = parsed.value.attachments.?;
    try t.expectEqual(@as(usize, 2), atts.len);
    try t.expectEqualStrings("hello", atts[0].text.?.content.?);
    try t.expectEqualStrings("SELECT 1", atts[1].query.?.query.?);
}

test "query-result parses columns and null cells" {
    const fixture =
        \\{"statement_response":{"manifest":{"schema":{"columns":[{"name":"region","type_name":"STRING"}]}},
        \\ "result":{"data_array":[["AMER"],[null]]}}}
    ;
    const parsed = try std.json.parseFromSlice(QueryResult, t.allocator, fixture, .{
        .ignore_unknown_fields = true,
    });
    defer parsed.deinit();
    const sr = parsed.value.statement_response.?;
    try t.expectEqualStrings("region", sr.manifest.?.schema.?.columns.?[0].name);
    const rows = sr.result.?.data_array.?;
    try t.expectEqualStrings("AMER", rows[0][0].?);
    try t.expectEqual(@as(?[]const u8, null), rows[1][0]);
}

test "isTerminalStatus" {
    try t.expect(isTerminalStatus("COMPLETED"));
    try t.expect(isTerminalStatus("FAILED"));
    try t.expect(isTerminalStatus("CANCELLED"));
    try t.expect(isTerminalStatus("QUERY_RESULT_EXPIRED"));
    try t.expect(!isTerminalStatus("EXECUTING_QUERY"));
    try t.expect(!isTerminalStatus(""));
}

test "encodeSegmentAlloc neutralizes path and query splices" {
    const evil = try encodeSegmentAlloc(t.allocator, "../x?y=1&z=#f");
    defer t.allocator.free(evil);
    try t.expectEqualStrings("..%2Fx%3Fy%3D1%26z%3D%23f", evil);

    const uuid = try encodeSegmentAlloc(t.allocator, "01f18d9b85231023baffbe55");
    defer t.allocator.free(uuid);
    try t.expectEqualStrings("01f18d9b85231023baffbe55", uuid);
}

test "sanitizedCopy strips control bytes and truncates" {
    var buf: [8]u8 = undefined;
    try t.expectEqualStrings("a.b.c", sanitizedCopy(&buf, "a\x1bb\rc"));
    try t.expectEqualStrings("12345678", sanitizedCopy(&buf, "123456789abc"));
}
