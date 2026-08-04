# Third-party notices

zlsx has **zero third-party runtime dependencies**: everything it links
is the Zig standard library plus code in this repository. What it does
carry is *third-party data* — generated lookup tables derived from the
Unicode Character Database — and that data comes with an attribution
requirement. This file discharges it.

> **Scope note.** These notices cover the derived data files listed
> below. They do not license zlsx itself; see [`LICENSE`](LICENSE).
> `LICENSE` needs a carve-out saying so explicitly — flagged as an
> owner action in `goal_formula.md` (M-1 row) and still open.

Every distributed artifact ships this file: source archives, release
tarballs and zips, the Homebrew formula's `doc` install, and both the
Python wheel and sdist. `scripts/ci/check_third_party_notices.sh` is
the gate; it fails the build if any of them stops carrying it.

---

## Unicode Character Database

**Used by**

| Generated file | Property data | Consumer |
|---|---|---|
| `unicode/tables/xid_data.zig` | `XID_Start`, `XID_Continue` from `DerivedCoreProperties.txt` | `unicode/xid.zig` — formula identifier grammar |
| `unicode/tables/nfc_data.zig` | canonical decompositions, combining classes, composition exclusions from `UnicodeData.txt` + `CompositionExclusions.txt` | `unicode/nfc.zig` — NFC normalization |
| `unicode/tables/casefold_data.zig` | full case folding (statuses C + F) from `CaseFolding.txt` | `unicode/casefold.zig` — sheet-name comparison, `collation_v1` |
| `unicode/tables/casing_data.zig` | full upper/lower/title mappings from `UnicodeData.txt` + unconditional `SpecialCasing.txt`, plus `Cased` and `Case_Ignorable` from `DerivedCoreProperties.txt` | `unicode/casing.zig` — `casing_v1`, the `UPPER`/`LOWER` functions |

Each generated file pins its Unicode version and the SHA-256 of every
input in its header. `scripts/gen_unicode_tables.py` regenerates them;
`scripts/ci/check_unicode_tables.sh` re-derives from the pinned inputs
and fails on any diff.

**License**

<!-- Verbatim copy of https://www.unicode.org/license.txt — do not reflow. -->

```
UNICODE LICENSE V3

COPYRIGHT AND PERMISSION NOTICE

Copyright © 1991-2026 Unicode, Inc.

NOTICE TO USER: Carefully read the following legal agreement. BY
DOWNLOADING, INSTALLING, COPYING OR OTHERWISE USING DATA FILES, AND/OR
SOFTWARE, YOU UNEQUIVOCALLY ACCEPT, AND AGREE TO BE BOUND BY, ALL OF THE
TERMS AND CONDITIONS OF THIS AGREEMENT. IF YOU DO NOT AGREE, DO NOT
DOWNLOAD, INSTALL, COPY, DISTRIBUTE OR USE THE DATA FILES OR SOFTWARE.

Permission is hereby granted, free of charge, to any person obtaining a
copy of data files and any associated documentation (the "Data Files") or
software and any associated documentation (the "Software") to deal in the
Data Files or Software without restriction, including without limitation
the rights to use, copy, modify, merge, publish, distribute, and/or sell
copies of the Data Files or Software, and to permit persons to whom the
Data Files or Software are furnished to do so, provided that either (a)
this copyright and permission notice appear with all copies of the Data
Files or Software, or (b) this copyright and permission notice appear in
associated Documentation.

THE DATA FILES AND SOFTWARE ARE PROVIDED "AS IS", WITHOUT WARRANTY OF ANY
KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT OF
THIRD PARTY RIGHTS.

IN NO EVENT SHALL THE COPYRIGHT HOLDER OR HOLDERS INCLUDED IN THIS NOTICE
BE LIABLE FOR ANY CLAIM, OR ANY SPECIAL INDIRECT OR CONSEQUENTIAL DAMAGES,
OR ANY DAMAGES WHATSOEVER RESULTING FROM LOSS OF USE, DATA OR PROFITS,
WHETHER IN AN ACTION OF CONTRACT, NEGLIGENCE OR OTHER TORTIOUS ACTION,
ARISING OUT OF OR IN CONNECTION WITH THE USE OR PERFORMANCE OF THE DATA
FILES OR SOFTWARE.

Except as contained in this notice, the name of a copyright holder shall
not be used in advertising or otherwise to promote the sale, use or other
dealings in these Data Files or Software without prior written
authorization of the copyright holder.
```

Unicode and the Unicode Logo are registered trademarks of Unicode, Inc.
in the U.S. and other countries.

---

## Zig test runner (`vendor/zig-test-runner/`)

`test_runner.zig` is a copy of Zig 0.16.0's
`lib/compiler/test_runner.zig` with one hunk changed, vendored so
coverage-guided fuzzing works at all: Zig 0.16.0 cannot compile its own
test runner under `-ffuzz`. It is licensed under the MIT license as
part of the Zig project (Copyright © Zig contributors); see
`vendor/zig-test-runner/README.md` for the provenance and the exact
diff.

It is a **build-time** artifact — it links into the two `-ffuzz` test
binaries and into no released artifact — so it is listed here for
completeness rather than because a distributed binary carries it.
