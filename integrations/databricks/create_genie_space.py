"""Create a Genie space over zlsx-landed Delta tables, via the REST API.

Reads DATABRICKS_HOST / DATABRICKS_TOKEN / DATABRICKS_WAREHOUSE_ID from the
environment (source your .env first) and shells out to `databricks api`.

Verified 2026-07-30. Two API facts learned the hard way:

  - serialized_space is version 2 and the validator REJECTS unsorted
    column_configs — they must be ordered by column_name.
  - POST /api/2.0/genie/spaces takes warehouse_id + title + description at
    the top level and the rest as a JSON-encoded serialized_space string.

The space this created ("zlsx Excel landing zone — FY2026 sales") answered a
cross-table attainment question and a _source_file provenance follow-up
correctly on first try — the instructions block below is what made the
month→quarter mapping and the join work.
"""
import json
import os
import subprocess
import uuid


def hid():
    return uuid.uuid4().hex


def cols(cc):
    # The serialized_space validator requires column_configs sorted by name.
    return sorted(cc, key=lambda c: c["column_name"])


sales_cols = cols([
    {"column_name": "region", "synonyms": ["geo", "territory", "market"], "enable_format_assistance": True},
    {"column_name": "product", "synonyms": ["item", "sku", "product line"], "enable_format_assistance": True},
    {"column_name": "month", "synonyms": ["period"], "enable_format_assistance": True},
    {"column_name": "units", "synonyms": ["quantity", "volume", "units sold"], "enable_format_assistance": True},
    {"column_name": "revenue", "synonyms": ["sales", "turnover", "income"], "enable_format_assistance": True},
    {"column_name": "_source_file", "synonyms": ["workbook", "source spreadsheet", "excel file", "provenance"], "enable_format_assistance": True},
])
target_cols = cols([
    {"column_name": "region", "synonyms": ["geo", "territory", "market"], "enable_format_assistance": True},
    {"column_name": "quarter", "synonyms": ["fiscal quarter", "Q"], "enable_format_assistance": True},
    {"column_name": "target_revenue", "synonyms": ["target", "quota", "goal", "plan"], "enable_format_assistance": True},
    {"column_name": "_source_file", "synonyms": ["workbook", "source spreadsheet", "excel file"], "enable_format_assistance": True},
])

serialized = {
    "version": 2,
    "config": {"sample_questions": [
        {"id": hid(), "question": ["What was total revenue by region in Q1 2026?"]},
        {"id": hid(), "question": ["Which region is furthest from its Q2 target?"]},
        {"id": hid(), "question": ["Show the monthly revenue trend by product"]},
        {"id": hid(), "question": ["Which Excel workbook did this data come from?"]},
    ]},
    "data_sources": {"tables": [
        {"identifier": "workspace.default.zlsx_sales", "column_configs": sales_cols},
        {"identifier": "workspace.default.zlsx_targets", "column_configs": target_cols},
    ]},
    "instructions": {"text_instructions": [{"id": hid(), "content": [
        "# zlsx Excel landing zone — FY2026 sales\n", "\n",
        "This data arrived as an Excel workbook and was ingested into Delta by the\n",
        "zlsx landing zone. `_source_file` on every row is the Volume path of the\n",
        "originating workbook — use it to answer any provenance question.\n", "\n",
        "- `zlsx_sales` is monthly (2026-01 .. 2026-06); `zlsx_targets` is quarterly.\n",
        "- Map months to quarters: 2026-01..2026-03 = 2026-Q1, 2026-04..2026-06 = 2026-Q2.\n",
        "- Join the two tables on region (plus the derived quarter).\n",
        "- 'Attainment' means SUM(revenue) / target_revenue for the matching region+quarter.\n",
        "- Revenue and targets are in USD. Format currency with 2 decimals.\n",
    ]}]},
}

body = {
    "warehouse_id": os.environ["DATABRICKS_WAREHOUSE_ID"],
    "title": "zlsx Excel landing zone — FY2026 sales",
    "description": "Ask questions about FY2026 sales and targets that landed as Excel workbooks via the zlsx ingestion path. Provenance (_source_file) preserved on every row.",
    "serialized_space": json.dumps(serialized),
}
out = subprocess.run(["databricks", "api", "post", "/api/2.0/genie/spaces", "--json", json.dumps(body)],
                     capture_output=True, text=True)
d = json.loads(out.stdout or "{}")
if d.get("space_id"):
    print("space_id:", d["space_id"])
    print("title:", d.get("title"))
else:
    print("FAILED")
    print(out.stdout[:600])
    print(out.stderr[:600])
    raise SystemExit(1)
