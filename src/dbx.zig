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
    \\  audit </Volumes/dir/>              hash every workbook in a landing zone
    \\        --manifest <file>            compare against a prior run (drift)
    \\        --write-manifest <file>      record this run's hashes
    \\        --table <cat.sch.tbl>        compare against an ingestion record
    \\        --source-column <name>       provenance column (default _source_file)
    \\        --warehouse <id>             SQL warehouse (default: $DATABRICKS_WAREHOUSE_ID)
    \\        --timeout-secs <n>           poll budget (default 120)
    \\
    \\environment: DATABRICKS_HOST, DATABRICKS_TOKEN, GENIE_SPACE_ID,
    \\             DATABRICKS_WAREHOUSE_ID
    \\
    \\exit: 0 clean, 1 usage, 2 API failure, 3 findings (audit) or bad
    \\      content, 5 local I/O
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
        if (std.mem.eql(u8, cmd, "audit")) break :blk try runAudit(alloc, io, environ, argv[1..], out, err_w);
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

// ─── audit ────────────────────────────────────────────────────────────
//
// Idea F of the Databricks track: does what is in the landing zone match
// what was ingested? Three questions, three evidence sources:
//
//   drift    content hash now vs a manifest from a prior run
//   orphan   workbook in the zone that the ingestion record never saw
//   missing  ingestion record (or manifest) pointing at a file that is gone
//
// A workbook is identified by its CONTENT hash, not its mtime/size
// fingerprint. The streaming source keys its offsets on (mtime, size)
// because that is all a cheap directory listing gives it; an audit is
// allowed to be expensive and read the bytes, which is the only way to
// catch a rewrite that preserved size and touched mtime backwards.

/// The audit downloads and hashes every workbook in the zone, so a
/// single oversized file should fail that file, not the run. Kept well
/// under `max_response_bytes` so a zone of large workbooks still
/// terminates in bounded memory (one file at a time).
const max_audit_file_bytes: usize = 256 * 1024 * 1024;

/// Directory recursion bound. A landing zone is flat or nearly so;
/// anything deeper is a misconfigured root and would turn one audit into
/// an unbounded crawl of a catalog's Volumes.
const max_audit_depth: u8 = 8;

const AuditStatus = enum {
    /// Hash matches the manifest and the ingestion record has seen it.
    ok,
    /// Present and readable, but no prior manifest entry to compare to.
    new,
    /// Content hash differs from the manifest — the immutable-files
    /// convention the streaming source relies on has been broken.
    drift,
    /// In the zone, absent from the ingestion record.
    orphan,
    /// In the manifest or the ingestion record, absent from the zone.
    missing,
    /// Bytes are there but do not parse as a workbook.
    unreadable,

    fn isFinding(self: AuditStatus) bool {
        return self != .ok and self != .new;
    }
};

const AuditFile = struct {
    path: []const u8, // owned
    size: u64,
    sha256: [64]u8,
    sheets: usize,
    readable: bool,
};

/// On-disk manifest. An array rather than a path-keyed object so it
/// parses into a plain struct; the audit is O(n·m) over it, which is
/// noise next to downloading every workbook.
const Manifest = struct {
    version: u32 = 1,
    root: []const u8 = "",
    files: []const ManifestEntry = &.{},
};

const ManifestEntry = struct {
    path: []const u8,
    sha256: []const u8,
    size: u64 = 0,
    sheets: usize = 0,
};

// ─── Files API directory listing ─────────────────────────────────────

const DirEntry = struct {
    path: []const u8 = "",
    name: []const u8 = "",
    is_directory: bool = false,
    file_size: u64 = 0,
};

const DirListing = struct {
    contents: []const DirEntry = &.{},
    next_page_token: ?[]const u8 = null,
};

/// Files-API URL for a Volume *directory*. Same Volumes-only refusal as
/// `filesApiUrl`; different endpoint.
fn directoriesApiUrl(
    alloc: std.mem.Allocator,
    host: []const u8,
    volume_path: []const u8,
    page_token: ?[]const u8,
) (PathError || std.mem.Allocator.Error)![]u8 {
    var p = volume_path;
    if (std.mem.startsWith(u8, p, "dbfs:")) p = p["dbfs:".len..];
    if (!std.mem.startsWith(u8, p, "/Volumes/")) return error.NotAVolumePath;
    // A trailing slash yields an empty last segment and a 400 from the API.
    while (p.len > "/Volumes/".len and p[p.len - 1] == '/') p = p[0 .. p.len - 1];

    var aw: Writer.Allocating = .init(alloc);
    defer aw.deinit();
    const w = &aw.writer;
    w.writeAll(host) catch return error.OutOfMemory;
    w.writeAll("/api/2.0/fs/directories") catch return error.OutOfMemory;
    percentEncodePath(w, p) catch return error.OutOfMemory;
    if (page_token) |tok| {
        w.writeAll("?page_token=") catch return error.OutOfMemory;
        percentEncode(w, tok, false) catch return error.OutOfMemory;
    }
    return alloc.dupe(u8, aw.written());
}

fn isWorkbookName(name: []const u8) bool {
    return std.mem.endsWith(u8, name, ".xlsx") or std.mem.endsWith(u8, name, ".xlsm");
}

/// Recursively collect workbook paths under `dir`. Returns owned paths in
/// the caller's allocator, sorted, so the audit output is deterministic
/// regardless of the order the API pages them back.
fn listWorkbooks(
    alloc: std.mem.Allocator,
    io: std.Io,
    auth: Auth,
    dir: []const u8,
    depth: u8,
    acc: *std.ArrayListUnmanaged([]u8),
    err_w: *Writer,
) !void {
    if (depth > max_audit_depth) {
        try err_w.print("zlsx dbx audit: skipping {s} (deeper than {d} levels)\n", .{ dir, max_audit_depth });
        return;
    }

    var page_token: ?[]u8 = null;
    defer if (page_token) |tok| alloc.free(tok);

    while (true) {
        const url = try directoriesApiUrl(alloc, auth.host, dir, page_token);
        defer alloc.free(url);

        var res = try httpRequest(alloc, io, .GET, url, auth.token, null, null);
        defer res.deinit(alloc);
        if (res.status != 200) {
            try reportApiError(err_w, "audit list", res.status, res.body);
            return error.ListFailed;
        }

        // .alloc_always: the parsed strings are copied into `acc` and the
        // recursion below outlives `res.body`.
        const parsed = std.json.parseFromSlice(DirListing, alloc, res.body, .{
            .ignore_unknown_fields = true,
            .allocate = .alloc_always,
        }) catch {
            try err_w.print("zlsx dbx audit: unrecognized listing for {s}\n", .{dir});
            return error.UnrecognizedResponse;
        };
        defer parsed.deinit();

        for (parsed.value.contents) |e| {
            if (e.path.len == 0) continue;
            if (e.is_directory) {
                try listWorkbooks(alloc, io, auth, e.path, depth + 1, acc, err_w);
            } else if (isWorkbookName(e.name)) {
                try acc.append(alloc, try alloc.dupe(u8, e.path));
            }
        }

        const next = parsed.value.next_page_token orelse break;
        if (next.len == 0) break;
        if (page_token) |tok| alloc.free(tok);
        page_token = try alloc.dupe(u8, next);
    }
}

fn lessThanPath(_: void, a: []u8, b: []u8) bool {
    return std.mem.lessThan(u8, a, b);
}

fn sha256Hex(bytes: []const u8) [64]u8 {
    var digest: [32]u8 = undefined;
    std.crypto.hash.sha2.Sha256.hash(bytes, &digest, .{});
    var hex: [64]u8 = undefined;
    _ = std.fmt.bufPrint(&hex, "{x}", .{&digest}) catch unreachable;
    return hex;
}

// ─── ingestion record (SQL Statement Execution API) ──────────────────

const SqlError = struct {
    error_code: ?[]const u8 = null,
    message: ?[]const u8 = null,
};

const SqlStatus = struct {
    state: []const u8 = "",
    /// `error` is a Zig keyword; the wire name is not.
    @"error": ?SqlError = null,
};

const SqlChunk = struct {
    data_array: ?[]const []const ?[]const u8 = null,
};

const SqlManifest = struct {
    truncated: bool = false,
};

const SqlResponse = struct {
    statement_id: ?[]const u8 = null,
    status: ?SqlStatus = null,
    manifest: ?SqlManifest = null,
    result: ?SqlChunk = null,
};

/// Unity Catalog identifiers we are willing to splice into SQL. The
/// audit builds its own statement, so the CLI argument is an injection
/// vector unless it is constrained to something that cannot terminate an
/// identifier: letters, digits, underscore, and the dotted separator.
/// Backtick-quoting alone would not be enough — a backtick in the
/// argument would close the quote.
fn isSafeIdentifier(s: []const u8, allow_dots: bool) bool {
    if (s.len == 0 or s.len > 255) return false;
    var last_was_dot = true; // leading dot is invalid
    for (s) |c| {
        const ok = (c >= 'A' and c <= 'Z') or (c >= 'a' and c <= 'z') or
            (c >= '0' and c <= '9') or c == '_' or (allow_dots and c == '.');
        if (!ok) return false;
        if (c == '.') {
            if (last_was_dot) return false; // empty part
            last_was_dot = true;
        } else last_was_dot = false;
    }
    return !last_was_dot; // trailing dot is invalid
}

/// Quote a validated dotted identifier part-by-part:
/// `cat.sch.tbl` → `` `cat`.`sch`.`tbl` ``.
fn quoteIdentifier(alloc: std.mem.Allocator, dotted: []const u8) ![]u8 {
    var aw: Writer.Allocating = .init(alloc);
    defer aw.deinit();
    var it = std.mem.splitScalar(u8, dotted, '.');
    var first = true;
    while (it.next()) |part| {
        if (!first) aw.writer.writeByte('.') catch return error.OutOfMemory;
        first = false;
        aw.writer.print("`{s}`", .{part}) catch return error.OutOfMemory;
    }
    return alloc.dupe(u8, aw.written());
}

/// Distinct provenance values from the ingestion table. Returns owned
/// strings; caller frees each and the slice.
///
/// Polls rather than trusting one `wait_timeout`: a cold serverless
/// warehouse leaves the statement PENDING well past the wait, and
/// reading `result` off a PENDING response reports zero rows — which an
/// audit would render as "every workbook is an orphan".
fn fetchIngestedPaths(
    alloc: std.mem.Allocator,
    io: std.Io,
    auth: Auth,
    warehouse_id: []const u8,
    table: []const u8,
    source_column: []const u8,
    timeout_s: u32,
    out_truncated: *bool,
    err_w: *Writer,
) ![][]u8 {
    const qtable = try quoteIdentifier(alloc, table);
    defer alloc.free(qtable);
    const qcolumn = try quoteIdentifier(alloc, source_column);
    defer alloc.free(qcolumn);

    const statement = try std.fmt.allocPrint(
        alloc,
        "SELECT DISTINCT {s} FROM {s} WHERE {s} IS NOT NULL",
        .{ qcolumn, qtable, qcolumn },
    );
    defer alloc.free(statement);

    const body = try std.fmt.allocPrint(alloc, "{f}", .{std.json.fmt(.{
        .statement = statement,
        .warehouse_id = warehouse_id,
        .wait_timeout = "30s",
        .on_wait_timeout = "CONTINUE",
        .disposition = "INLINE",
        .format = "JSON_ARRAY",
    }, .{})});
    defer alloc.free(body);

    const url = try std.fmt.allocPrint(alloc, "{s}/api/2.0/sql/statements", .{auth.host});
    defer alloc.free(url);

    var res = try httpRequest(alloc, io, .POST, url, auth.token, body, "application/json");
    defer res.deinit(alloc);
    if (res.status != 200) {
        try reportApiError(err_w, "audit sql", res.status, res.body);
        return error.SqlFailed;
    }

    var owned_body: []u8 = try alloc.dupe(u8, res.body);
    defer alloc.free(owned_body);

    // Wall-clock budget, same shape as the genie poll loop: request
    // latency rides on top of every sleep, so an attempt count would
    // quietly overshoot the stated budget.
    const poll_start = std.Io.Timestamp.now(io, .awake);
    while (true) {
        const parsed = std.json.parseFromSlice(SqlResponse, alloc, owned_body, .{
            .ignore_unknown_fields = true,
            .allocate = .alloc_always,
        }) catch {
            try err_w.writeAll("zlsx dbx audit: unrecognized SQL response\n");
            return error.UnrecognizedResponse;
        };
        defer parsed.deinit();

        const state = if (parsed.value.status) |s| s.state else "";
        if (std.mem.eql(u8, state, "SUCCEEDED")) {
            if (parsed.value.manifest) |m| out_truncated.* = m.truncated;

            var paths: std.ArrayListUnmanaged([]u8) = .empty;
            errdefer {
                for (paths.items) |p| alloc.free(p);
                paths.deinit(alloc);
            }
            if (parsed.value.result) |r| if (r.data_array) |rows| for (rows) |row| {
                if (row.len == 0) continue;
                const cell = row[0] orelse continue;
                try paths.append(alloc, try alloc.dupe(u8, cell));
            };
            return paths.toOwnedSlice(alloc);
        }
        if (std.mem.eql(u8, state, "FAILED") or std.mem.eql(u8, state, "CANCELED") or
            std.mem.eql(u8, state, "CLOSED"))
        {
            // Carry the warehouse's own diagnosis through: "FAILED" alone
            // leaves the caller guessing between a typo'd table, a
            // missing column, and a permissions denial.
            var reason_buf: [400]u8 = undefined;
            const reason: []const u8 = if (parsed.value.status) |s|
                if (s.@"error") |e|
                    sanitizedCopy(&reason_buf, e.message orelse e.error_code orelse "")
                else
                    ""
            else
                "";
            if (reason.len > 0) {
                try err_w.print("zlsx dbx audit: SQL statement {s}: {s}\n", .{ state, reason });
            } else {
                try err_w.print("zlsx dbx audit: SQL statement {s}\n", .{state});
            }
            return error.SqlFailed;
        }
        const elapsed = poll_start.durationTo(std.Io.Timestamp.now(io, .awake));
        if (elapsed.nanoseconds >= @as(i96, timeout_s) * 1_000_000_000) {
            try err_w.print(
                "zlsx dbx audit: SQL statement still {s} after {d}s (warehouse cold?)\n",
                .{ state, timeout_s },
            );
            return error.SqlTimeout;
        }

        const id = parsed.value.statement_id orelse {
            try err_w.writeAll("zlsx dbx audit: SQL response carried no statement_id\n");
            return error.UnrecognizedResponse;
        };
        const id_enc = try encodeSegmentAlloc(alloc, id);
        defer alloc.free(id_enc);
        const poll_url = try std.fmt.allocPrint(
            alloc,
            "{s}/api/2.0/sql/statements/{s}",
            .{ auth.host, id_enc },
        );
        defer alloc.free(poll_url);

        io.sleep(.fromSeconds(poll_interval_s), .awake) catch {};

        var poll = try httpRequest(alloc, io, .GET, poll_url, auth.token, null, null);
        defer poll.deinit(alloc);
        if (poll.status != 200) {
            try reportApiError(err_w, "audit sql poll", poll.status, poll.body);
            return error.SqlFailed;
        }
        alloc.free(owned_body);
        owned_body = try alloc.dupe(u8, poll.body);
    }
}

// ─── audit orchestration ─────────────────────────────────────────────

fn runAudit(
    alloc: std.mem.Allocator,
    io: std.Io,
    environ: std.process.Environ,
    argv: []const []const u8,
    out: *Writer,
    err_w: *Writer,
) !u8 {
    var root: ?[]const u8 = null;
    var manifest_path: ?[]const u8 = null;
    var write_manifest_path: ?[]const u8 = null;
    var table: ?[]const u8 = null;
    var source_column: []const u8 = "_source_file";
    var warehouse_flag: ?[]const u8 = null;
    var timeout_s: u32 = default_timeout_s;

    var i: usize = 0;
    while (i < argv.len) : (i += 1) {
        const a = argv[i];
        const takes_value =
            std.mem.eql(u8, a, "--manifest") or
            std.mem.eql(u8, a, "--write-manifest") or
            std.mem.eql(u8, a, "--table") or
            std.mem.eql(u8, a, "--source-column") or
            std.mem.eql(u8, a, "--warehouse") or
            std.mem.eql(u8, a, "--timeout-secs");
        if (takes_value) {
            i += 1;
            if (i >= argv.len) {
                try err_w.print("zlsx dbx audit: {s} needs a value\n", .{a});
                return 1;
            }
            const v = argv[i];
            if (std.mem.eql(u8, a, "--manifest")) manifest_path = v;
            if (std.mem.eql(u8, a, "--write-manifest")) write_manifest_path = v;
            if (std.mem.eql(u8, a, "--table")) table = v;
            if (std.mem.eql(u8, a, "--source-column")) source_column = v;
            if (std.mem.eql(u8, a, "--warehouse")) warehouse_flag = v;
            if (std.mem.eql(u8, a, "--timeout-secs")) {
                timeout_s = std.fmt.parseInt(u32, v, 10) catch {
                    try err_w.print("zlsx dbx audit: bad --timeout-secs '{s}'\n", .{v});
                    return 1;
                };
            }
        } else if (std.mem.startsWith(u8, a, "--")) {
            try err_w.print("zlsx dbx audit: unknown flag '{s}'\n", .{a});
            return 1;
        } else if (root == null) {
            root = a;
        } else {
            try err_w.print("zlsx dbx audit: unexpected argument '{s}'\n", .{a});
            return 1;
        }
    }

    const zone = root orelse {
        try err_w.writeAll("zlsx dbx audit: missing </Volumes/dir/>\n");
        return 1;
    };
    if (table) |tbl| {
        if (!isSafeIdentifier(tbl, true)) {
            try err_w.print(
                "zlsx dbx audit: '{s}' is not a plain catalog.schema.table identifier\n",
                .{tbl},
            );
            return 1;
        }
        if (!isSafeIdentifier(source_column, false)) {
            try err_w.print(
                "zlsx dbx audit: '{s}' is not a plain column identifier\n",
                .{source_column},
            );
            return 1;
        }
    }

    var auth = (try loadAuth(alloc, environ, err_w)) orelse return 1;
    defer auth.deinit(alloc);

    // 1. Enumerate the zone.
    var found: std.ArrayListUnmanaged([]u8) = .empty;
    defer {
        for (found.items) |p| alloc.free(p);
        found.deinit(alloc);
    }
    listWorkbooks(alloc, io, auth, zone, 0, &found, err_w) catch |e| switch (e) {
        error.NotAVolumePath => {
            try err_w.print("zlsx dbx audit: '{s}' is not a /Volumes/ path\n", .{zone});
            return 1;
        },
        error.ListFailed, error.UnrecognizedResponse => return 2,
        else => {
            try err_w.print("zlsx dbx audit: listing failed ({s})\n", .{@errorName(e)});
            return 2;
        },
    };
    std.mem.sort([]u8, found.items, {}, lessThanPath);

    // 2. Prior manifest, if any.
    var manifest_json: ?[]u8 = null;
    defer if (manifest_json) |m| alloc.free(m);
    var prior: ?std.json.Parsed(Manifest) = null;
    defer if (prior) |*p| p.deinit();
    if (manifest_path) |mp| {
        manifest_json = std.Io.Dir.cwd().readFileAlloc(io, mp, alloc, .limited(64 * 1024 * 1024)) catch |e| {
            try err_w.print("zlsx dbx audit: cannot read manifest {s}: {s}\n", .{ mp, @errorName(e) });
            return 5;
        };
        prior = std.json.parseFromSlice(Manifest, alloc, manifest_json.?, .{
            .ignore_unknown_fields = true,
        }) catch {
            try err_w.print("zlsx dbx audit: {s} is not a zlsx audit manifest\n", .{mp});
            return 5;
        };
    }

    // 3. Ingestion record, if asked for.
    var ingested: ?[][]u8 = null;
    defer if (ingested) |paths| {
        for (paths) |p| alloc.free(p);
        alloc.free(paths);
    };
    var sql_truncated = false;
    if (table) |tbl| {
        const wh_owned: ?[]u8 = if (warehouse_flag == null)
            environ.getAlloc(alloc, "DATABRICKS_WAREHOUSE_ID") catch |e| switch (e) {
                error.EnvironmentVariableMissing => null,
                else => return e,
            }
        else
            null;
        defer if (wh_owned) |w| alloc.free(w);
        const warehouse = warehouse_flag orelse (wh_owned orelse {
            try err_w.writeAll(
                "zlsx dbx audit: --table needs a warehouse (--warehouse or DATABRICKS_WAREHOUSE_ID)\n",
            );
            return 1;
        });
        ingested = fetchIngestedPaths(
            alloc,
            io,
            auth,
            warehouse,
            tbl,
            source_column,
            timeout_s,
            &sql_truncated,
            err_w,
        ) catch |e| switch (e) {
            error.SqlFailed, error.SqlTimeout, error.UnrecognizedResponse => return 2,
            else => {
                try err_w.print("zlsx dbx audit: ingestion query failed ({s})\n", .{@errorName(e)});
                return 2;
            },
        };
        if (sql_truncated) {
            try err_w.writeAll(
                "zlsx dbx audit: ingestion result was TRUNCATED by the warehouse; " ++
                    "orphan findings below are not trustworthy\n",
            );
        }
    }

    // 4. Fetch, hash, verify, classify — one workbook at a time so peak
    //    memory is one file, not the whole zone.
    var seen: std.ArrayListUnmanaged(AuditFile) = .empty;
    defer {
        for (seen.items) |f| alloc.free(f.path);
        seen.deinit(alloc);
    }
    var findings: usize = 0;

    for (found.items) |path| {
        const url = filesApiUrl(alloc, auth.host, path) catch |e| switch (e) {
            error.NotAVolumePath => continue,
            else => return e,
        };
        defer alloc.free(url);

        var res = httpRequest(alloc, io, .GET, url, auth.token, null, null) catch |e| {
            try err_w.print("zlsx dbx audit: fetch {s} failed ({s})\n", .{ path, @errorName(e) });
            return 2;
        };
        defer res.deinit(alloc);
        if (res.status != 200) {
            try reportApiError(err_w, "audit fetch", res.status, res.body);
            return 2;
        }
        if (res.body.len > max_audit_file_bytes) {
            try err_w.print(
                "zlsx dbx audit: {s} is {d} bytes, over the {d} cap; skipped\n",
                .{ path, res.body.len, max_audit_file_bytes },
            );
            continue;
        }

        const hex = sha256Hex(res.body);
        var sheets: usize = 0;
        var readable = true;
        if (verifyWorkbook(alloc, io, res.body)) |n| {
            sheets = n;
        } else |_| {
            readable = false;
        }

        var status: AuditStatus = .ok;
        if (!readable) {
            status = .unreadable;
        } else if (prior) |p| blk: {
            for (p.value.files) |e| {
                if (std.mem.eql(u8, e.path, path)) {
                    status = if (std.mem.eql(u8, e.sha256, &hex)) .ok else .drift;
                    break :blk;
                }
            }
            status = .new;
        }
        // The ingestion record answers a different question than the
        // manifest, and an orphan matters more than "hash unchanged":
        // an unread workbook is data the table has never reflected.
        if (status == .ok or status == .new) {
            if (ingested) |paths| {
                var was_ingested = false;
                for (paths) |p| {
                    if (pathMatchesSource(path, p)) {
                        was_ingested = true;
                        break;
                    }
                }
                if (!was_ingested) status = .orphan;
            }
        }

        if (status.isFinding()) findings += 1;
        try out.print("{f}\n", .{std.json.fmt(.{
            .kind = "dbx_audit_file",
            .path = path,
            .status = @tagName(status),
            .bytes = res.body.len,
            .sha256 = &hex,
            .sheets = sheets,
        }, .{})});

        try seen.append(alloc, .{
            .path = try alloc.dupe(u8, path),
            .size = res.body.len,
            .sha256 = hex,
            .sheets = sheets,
            .readable = readable,
        });
    }

    // 5. Anything the manifest or the table knew about that is gone.
    if (prior) |p| {
        for (p.value.files) |e| {
            if (!containsPath(found.items, e.path)) {
                findings += 1;
                try out.print("{f}\n", .{std.json.fmt(.{
                    .kind = "dbx_audit_file",
                    .path = e.path,
                    .status = @tagName(AuditStatus.missing),
                    .source = "manifest",
                }, .{})});
            }
        }
    }
    if (ingested) |paths| {
        for (paths) |p| {
            if (!anyMatchesSource(found.items, p)) {
                findings += 1;
                try out.print("{f}\n", .{std.json.fmt(.{
                    .kind = "dbx_audit_file",
                    .path = p,
                    .status = @tagName(AuditStatus.missing),
                    .source = "table",
                }, .{})});
            }
        }
    }

    // 6. Record this run, if asked.
    if (write_manifest_path) |wp| {
        writeManifest(alloc, io, wp, zone, seen.items) catch |e| {
            try err_w.print("zlsx dbx audit: cannot write manifest {s}: {s}\n", .{ wp, @errorName(e) });
            return 5;
        };
    }

    try out.print("{f}\n", .{std.json.fmt(.{
        .kind = "dbx_audit_summary",
        .root = zone,
        .workbooks = seen.items.len,
        .findings = findings,
        .compared_to_manifest = manifest_path != null,
        .compared_to_table = table,
        .truncated = sql_truncated,
    }, .{})});

    return if (findings > 0) 3 else 0;
}

/// Does a Volume path match a provenance value recorded at ingest time?
///
/// Exact match first, then a suffix comparison: the same file is written
/// as `/Volumes/c/s/v/f.xlsx` by the Files API and as
/// `dbfs:/Volumes/c/s/v/f.xlsx` by the CLI, and Spark's `input_file_name()`
/// prefixes a scheme and authority. Suffix matching on a full path
/// component avoids `.../a/f.xlsx` matching `.../b/f.xlsx`.
fn pathMatchesSource(volume_path: []const u8, source: []const u8) bool {
    if (std.mem.eql(u8, volume_path, source)) return true;
    // Zone paths are absolute (`/Volumes/...`), so a suffix match already
    // starts on a '/' — the component boundary is free. Checking the
    // character *before* the suffix instead would reject the very
    // spellings this absorbs: `s3://bucket/Volumes/...` has 't' there.
    if (volume_path.len == 0 or volume_path[0] != '/') return false;
    return std.mem.endsWith(u8, source, volume_path);
}

fn anyMatchesSource(paths: []const []u8, source: []const u8) bool {
    for (paths) |p| if (pathMatchesSource(p, source)) return true;
    return false;
}

fn containsPath(paths: []const []u8, needle: []const u8) bool {
    for (paths) |p| if (std.mem.eql(u8, p, needle)) return true;
    return false;
}

fn writeManifest(
    alloc: std.mem.Allocator,
    io: std.Io,
    path: []const u8,
    root: []const u8,
    files: []const AuditFile,
) !void {
    var entries = try alloc.alloc(ManifestEntry, files.len);
    defer alloc.free(entries);
    // `|*f|`, not `|f|`: the hash is a fixed array stored inline, so a
    // by-value capture would put `&f.sha256` on this frame's loop slot
    // and every entry would alias it — the manifest then records the
    // LAST file's hash for all of them.
    for (files, 0..) |*f, idx| {
        entries[idx] = .{
            .path = f.path,
            .sha256 = &f.sha256,
            .size = f.size,
            .sheets = f.sheets,
        };
    }

    var aw: Writer.Allocating = .init(alloc);
    defer aw.deinit();
    try aw.writer.print("{f}\n", .{std.json.fmt(Manifest{
        .version = 1,
        .root = root,
        .files = entries,
    }, .{})});

    var write_buf: [64 * 1024]u8 = undefined;
    var af = try pkg.AtomicFile.init(io, path, &write_buf);
    defer af.deinit();
    try af.file_writer.interface.writeAll(aw.written());
    try af.finish();
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

// ─── audit ────────────────────────────────────────────────────────────

test "directoriesApiUrl targets the directories endpoint and trims trailing slashes" {
    const url = try directoriesApiUrl(t.allocator, "https://h", "/Volumes/c/s/v/zone/", null);
    defer t.allocator.free(url);
    try t.expectEqualStrings("https://h/api/2.0/fs/directories/Volumes/c/s/v/zone", url);

    const dbfs = try directoriesApiUrl(t.allocator, "https://h", "dbfs:/Volumes/c/s/v///", null);
    defer t.allocator.free(dbfs);
    try t.expectEqualStrings("https://h/api/2.0/fs/directories/Volumes/c/s/v", dbfs);

    // The Volumes root itself must survive the trailing-slash trim.
    const root = try directoriesApiUrl(t.allocator, "https://h", "/Volumes/", null);
    defer t.allocator.free(root);
    try t.expectEqualStrings("https://h/api/2.0/fs/directories/Volumes/", root);
}

test "directoriesApiUrl encodes the page token as a single segment" {
    const url = try directoriesApiUrl(t.allocator, "https://h", "/Volumes/c/s/v", "a/b+c=d");
    defer t.allocator.free(url);
    try t.expectEqualStrings(
        "https://h/api/2.0/fs/directories/Volumes/c/s/v?page_token=a%2Fb%2Bc%3Dd",
        url,
    );
}

test "directoriesApiUrl refuses non-Volume roots" {
    try t.expectError(error.NotAVolumePath, directoriesApiUrl(t.allocator, "https://h", "/tmp", null));
}

test "isWorkbookName accepts xlsx/xlsm only" {
    try t.expect(isWorkbookName("a.xlsx"));
    try t.expect(isWorkbookName("a.xlsm"));
    try t.expect(!isWorkbookName("a.xls"));
    try t.expect(!isWorkbookName("a.csv"));
    try t.expect(!isWorkbookName("xlsx"));
}

test "sha256Hex matches the known empty-input vector" {
    const hex = sha256Hex("");
    try t.expectEqualStrings(
        "e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855",
        &hex,
    );
    const abc = sha256Hex("abc");
    try t.expectEqualStrings(
        "ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad",
        &abc,
    );
}

test "isSafeIdentifier refuses anything that could escape a quoted identifier" {
    try t.expect(isSafeIdentifier("main", false));
    try t.expect(isSafeIdentifier("_source_file", false));
    try t.expect(isSafeIdentifier("cat.sch.tbl", true));

    // Dots only where allowed.
    try t.expect(!isSafeIdentifier("cat.sch", false));
    // Structural junk.
    try t.expect(!isSafeIdentifier("", true));
    try t.expect(!isSafeIdentifier(".lead", true));
    try t.expect(!isSafeIdentifier("trail.", true));
    try t.expect(!isSafeIdentifier("a..b", true));
    // Injection attempts — the backtick is the one that would close the
    // quote, the rest would terminate or extend the statement.
    try t.expect(!isSafeIdentifier("tbl`; DROP TABLE x; --", true));
    try t.expect(!isSafeIdentifier("tbl' OR '1'='1", true));
    try t.expect(!isSafeIdentifier("tbl x", true));
    try t.expect(!isSafeIdentifier("tbl\nUNION SELECT", true));
    try t.expect(!isSafeIdentifier("tbl-1", true));
}

test "quoteIdentifier backticks every part" {
    const q = try quoteIdentifier(t.allocator, "cat.sch.tbl");
    defer t.allocator.free(q);
    try t.expectEqualStrings("`cat`.`sch`.`tbl`", q);

    const one = try quoteIdentifier(t.allocator, "_source_file");
    defer t.allocator.free(one);
    try t.expectEqualStrings("`_source_file`", one);
}

test "pathMatchesSource spans the spellings a provenance column carries" {
    const vp = "/Volumes/c/s/v/f.xlsx";
    try t.expect(pathMatchesSource(vp, vp));
    try t.expect(pathMatchesSource(vp, "dbfs:/Volumes/c/s/v/f.xlsx"));
    try t.expect(pathMatchesSource(vp, "s3://bucket/Volumes/c/s/v/f.xlsx"));

    // Same basename, different directory, must NOT match.
    try t.expect(!pathMatchesSource("/Volumes/c/s/v/a/f.xlsx", "/Volumes/c/s/v/b/f.xlsx"));
    // Suffix that is not on a component boundary must NOT match.
    try t.expect(!pathMatchesSource("/Volumes/c/s/v/f.xlsx", "/Volumes/c/s/v/xf.xlsx"));
    try t.expect(!pathMatchesSource(vp, "/Volumes/c/s/v/other.xlsx"));
}

test "AuditStatus: ok and new are not findings, everything else is" {
    try t.expect(!AuditStatus.ok.isFinding());
    try t.expect(!AuditStatus.new.isFinding());
    try t.expect(AuditStatus.drift.isFinding());
    try t.expect(AuditStatus.orphan.isFinding());
    try t.expect(AuditStatus.missing.isFinding());
    try t.expect(AuditStatus.unreadable.isFinding());
}

test "directory listing parses, tolerating absent optional fields" {
    const fixture =
        \\{"contents":[
        \\  {"path":"/Volumes/c/s/v/a.xlsx","name":"a.xlsx","is_directory":false,"file_size":10},
        \\  {"path":"/Volumes/c/s/v/sub","name":"sub","is_directory":true},
        \\  {"path":"/Volumes/c/s/v/notes.txt","name":"notes.txt","is_directory":false}
        \\],"next_page_token":"tok","unknown":1}
    ;
    const parsed = try std.json.parseFromSlice(DirListing, t.allocator, fixture, .{
        .ignore_unknown_fields = true,
    });
    defer parsed.deinit();
    try t.expectEqual(@as(usize, 3), parsed.value.contents.len);
    try t.expectEqualStrings("a.xlsx", parsed.value.contents[0].name);
    try t.expect(parsed.value.contents[1].is_directory);
    try t.expectEqual(@as(u64, 0), parsed.value.contents[2].file_size);
    try t.expectEqualStrings("tok", parsed.value.next_page_token.?);
}

test "listing without a next_page_token ends pagination" {
    const parsed = try std.json.parseFromSlice(DirListing, t.allocator, "{\"contents\":[]}", .{
        .ignore_unknown_fields = true,
    });
    defer parsed.deinit();
    try t.expectEqual(@as(?[]const u8, null), parsed.value.next_page_token);
}

test "SQL response parses states, rows, and NULL cells" {
    const running = try std.json.parseFromSlice(
        SqlResponse,
        t.allocator,
        \\{"statement_id":"s1","status":{"state":"PENDING"}}
    ,
        .{ .ignore_unknown_fields = true },
    );
    defer running.deinit();
    try t.expectEqualStrings("PENDING", running.value.status.?.state);
    try t.expectEqual(@as(?SqlChunk, null), running.value.result);

    const done = try std.json.parseFromSlice(
        SqlResponse,
        t.allocator,
        \\{"statement_id":"s1","status":{"state":"SUCCEEDED"},
        \\ "manifest":{"truncated":true},
        \\ "result":{"data_array":[["/Volumes/c/s/v/a.xlsx"],[null]]}}
    ,
        .{ .ignore_unknown_fields = true },
    );
    defer done.deinit();
    try t.expectEqualStrings("SUCCEEDED", done.value.status.?.state);
    try t.expect(done.value.manifest.?.truncated);
    const rows = done.value.result.?.data_array.?;
    try t.expectEqual(@as(usize, 2), rows.len);
    try t.expectEqualStrings("/Volumes/c/s/v/a.xlsx", rows[0][0].?);
    // A NULL provenance cell must survive parsing so the collector can
    // skip it rather than crash mid-audit.
    try t.expectEqual(@as(?[]const u8, null), rows[1][0]);
}

test "SQL failure carries the warehouse's own reason" {
    const parsed = try std.json.parseFromSlice(
        SqlResponse,
        t.allocator,
        \\{"statement_id":"s1","status":{"state":"FAILED","error":
        \\ {"error_code":"TABLE_OR_VIEW_NOT_FOUND","message":"[TABLE_OR_VIEW_NOT_FOUND] cat.sch.nope"}}}
    ,
        .{ .ignore_unknown_fields = true },
    );
    defer parsed.deinit();
    try t.expectEqualStrings("FAILED", parsed.value.status.?.state);
    try t.expectEqualStrings(
        "[TABLE_OR_VIEW_NOT_FOUND] cat.sch.nope",
        parsed.value.status.?.@"error".?.message.?,
    );
}

test "SQL status without an error object still parses" {
    const parsed = try std.json.parseFromSlice(
        SqlResponse,
        t.allocator,
        \\{"status":{"state":"SUCCEEDED"}}
    ,
        .{ .ignore_unknown_fields = true },
    );
    defer parsed.deinit();
    try t.expectEqual(@as(?SqlError, null), parsed.value.status.?.@"error");
}

test "manifest round-trips through the on-disk shape" {
    const files = [_]AuditFile{
        .{ .path = "/Volumes/c/s/v/a.xlsx", .size = 10, .sha256 = sha256Hex("a"), .sheets = 2, .readable = true },
        .{ .path = "/Volumes/c/s/v/b.xlsx", .size = 20, .sha256 = sha256Hex("b"), .sheets = 1, .readable = true },
    };

    var tmp = std.testing.tmpDir(.{});
    defer tmp.cleanup();
    var threaded: std.Io.Threaded = .init(t.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();

    const dir = try tmp.dir.realPathFileAlloc(io, ".", t.allocator);
    defer t.allocator.free(dir);
    const path = try std.fs.path.joinZ(t.allocator, &.{ dir, "manifest.json" });
    defer t.allocator.free(path);

    try writeManifest(t.allocator, io, path, "/Volumes/c/s/v", &files);

    const bytes = try std.Io.Dir.cwd().readFileAlloc(io, path, t.allocator, .limited(1 << 20));
    defer t.allocator.free(bytes);
    const parsed = try std.json.parseFromSlice(Manifest, t.allocator, bytes, .{
        .ignore_unknown_fields = true,
    });
    defer parsed.deinit();

    try t.expectEqual(@as(u32, 1), parsed.value.version);
    try t.expectEqualStrings("/Volumes/c/s/v", parsed.value.root);
    try t.expectEqual(@as(usize, 2), parsed.value.files.len);
    try t.expectEqualStrings("/Volumes/c/s/v/a.xlsx", parsed.value.files[0].path);
    try t.expectEqualStrings(&sha256Hex("a"), parsed.value.files[0].sha256);
    try t.expectEqual(@as(u64, 20), parsed.value.files[1].size);
}

test "containsPath / anyMatchesSource over the collected zone" {
    var a = "/Volumes/c/s/v/a.xlsx".*;
    var b = "/Volumes/c/s/v/b.xlsx".*;
    const zone = [_][]u8{ &a, &b };

    try t.expect(containsPath(&zone, "/Volumes/c/s/v/a.xlsx"));
    try t.expect(!containsPath(&zone, "/Volumes/c/s/v/c.xlsx"));

    try t.expect(anyMatchesSource(&zone, "dbfs:/Volumes/c/s/v/b.xlsx"));
    try t.expect(!anyMatchesSource(&zone, "/Volumes/c/s/v/gone.xlsx"));
}
