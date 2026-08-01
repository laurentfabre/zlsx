-- read_xlsx_json: query an xlsx workbook from PURE DBSQL — no Spark session,
-- no Delta copy. Verified on a serverless PRO warehouse 2026-07-31.
--
-- How it works: read_files(format => 'binaryFile') surfaces the workbook
-- bytes in SQL; this UC Python UDF parses them with py-zlsx (installed from
-- a UC Volume wheel via ENVIRONMENT) and returns the sheet as a JSON array
-- of objects keyed by the header row. A view then explodes it into columns.
--
-- Placeholders to adapt: catalog.schema (here workspace.default) and the
-- Volume paths. Sandbox facts, verified: the native-lib wheel loads
-- (ctypes/dlopen allowed), tempfile writes are allowed.
--
-- py-zlsx >= 0.6.0 has zlsx.open_bytes() (the buffer-based C ABI), so the
-- workbook bytes are parsed directly — no temp file. The hasattr fallback
-- keeps the function working on a 0.5.0 wheel, where a tempfile shim is
-- the only option.

CREATE OR REPLACE FUNCTION workspace.default.read_xlsx_json(content BINARY, sheet STRING)
RETURNS STRING
LANGUAGE PYTHON
ENVIRONMENT (
  dependencies = '["/Volumes/workspace/default/zlsx_smoke/py_zlsx-0.6.0-py3-none-manylinux_2_28_aarch64.whl"]',
  environment_version = '3'
)
AS $$
import json
import zlsx

sel = int(sheet) if sheet.isdigit() else sheet
if hasattr(zlsx, "open_bytes"):
    with zlsx.open_bytes(bytes(content)) as book:
        rows = list(book.sheet(sel).rows())
else:  # py-zlsx 0.5.0: no buffer API — fall back to a tempfile shim
    import tempfile, os
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
        f.write(bytes(content))
        path = f.name
    try:
        with zlsx.open(path) as book:
            rows = list(book.sheet(sel).rows())
    finally:
        os.unlink(path)
header = [str(h) for h in rows[0]]
return json.dumps([dict(zip(header, r)) for r in rows[1:]])
$$;

-- A LIVE view over the workbook: edit the file in the Volume and the next
-- query reflects it. Genie can sit on this view — a Genie room over a file.
CREATE OR REPLACE VIEW workspace.default.wb_sales_live
COMMENT 'LIVE view over the Excel workbook itself — parsed at query time by the zlsx UDF; no Delta copy. Source: /Volumes/workspace/default/zlsx_smoke/fy2026_sales.xlsx'
AS SELECT inline(from_json(workspace.default.read_xlsx_json(content, 'Sales'),
  'array<struct<region:string,product:string,month:string,units:bigint,revenue:double>>'))
FROM read_files('/Volumes/workspace/default/zlsx_smoke/fy2026_sales.xlsx', format => 'binaryFile');

-- Smoke query. First call pays the UDF environment build (~10 s); after
-- that the environment is cached by the warehouse.
SELECT region, ROUND(SUM(revenue), 2) AS revenue, SUM(units) AS units
FROM workspace.default.wb_sales_live
GROUP BY region ORDER BY region;
