#!/usr/bin/env python3
"""Regenerate tests/corpus/openpyxl_chart.xlsx — the openpyxl-written chart
witness for the chart `<c:f>` sweep.

openpyxl spells its chart part with the chart namespace as the DEFAULT
namespace (`<chartSpace xmlns:a="…/main" xmlns="…/drawingml/2006/chart">`
with unprefixed `<f>` carriers) and its drawing likewise
(`<wsDr … xmlns="…/spreadsheetDrawing">`). The workbook test
"openpyxl's default-namespace chart part" pins that a rename and a row
insert respell and shift the three carriers and that a save persists
them (in-house CF-REL-401: the shape had been documented as unproduced
and left unwalked). One sheet `Data` — Region / Qty over three rows —
and one clustered bar chart over B1 (title) / A2:A4 (categories) /
B2:B4 (values), anchored at D2.

    python3 scripts/gen_openpyxl_chart_fixture.py   # needs openpyxl (3.1.x)
"""
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference

out = Path(__file__).resolve().parent.parent / "tests" / "corpus" / "openpyxl_chart.xlsx"

wb = Workbook()
ws = wb.active
ws.title = "Data"
ws.append(["Region", "Qty"])
ws.append(["East", 3])
ws.append(["West", 4])
ws.append(["East", 5])

chart = BarChart()
chart.title = "Qty by region"
chart.add_data(Reference(ws, min_col=2, min_row=1, max_row=4), titles_from_data=True)
chart.set_categories(Reference(ws, min_col=1, min_row=2, max_row=4))
ws.add_chart(chart, "D2")

wb.save(out)
print(f"wrote {out}")
