"""Generate the FY2026 demo workbook the other scripts consume.

Two sheets so downstream demos have a join to reason about:
  Sales   — region x product x month (2026-01..06), 72 rows
  Targets — region x quarter, 6 rows

Deterministic (seeded); requires py-zlsx.

Usage: python3 make_fixture.py [out.xlsx]
"""
import random
import sys

import zlsx

OUT = sys.argv[1] if len(sys.argv) > 1 else "fy2026_sales.xlsx"

random.seed(2026)
regions = ["EMEA", "APAC", "AMER"]
products = [("widget", 40.0), ("gadget", 202.5), ("sprocket", 62.0), ("flange", 118.0)]
months = [f"2026-{m:02d}" for m in range(1, 7)]

sales = []
for rg in regions:
    lift = {"EMEA": 1.0, "APAC": 1.25, "AMER": 1.6}[rg]
    for p, price in products:
        for i, mo in enumerate(months):
            units = int(random.gauss(150, 40) * lift * (1 + 0.06 * i))
            units = max(units, 5)
            sales.append([rg, p, mo, units, round(units * price * random.uniform(0.92, 1.05), 2)])

targets = [[rg, q, round(sum(r[4] for r in sales if r[0] == rg) / 2 * f, -3)]
           for rg in regions for q, f in [("2026-Q1", 0.95), ("2026-Q2", 1.10)]]

with zlsx.write(OUT) as w:
    hdr = w.add_style(zlsx.Style(font_bold=True, font_color_argb=0xFFFFFFFF,
                                 fill_pattern="solid", fill_fg_argb=0xFF1E3A8A))
    s = w.add_sheet("Sales")
    s.freeze_panes(rows=1, cols=0)
    s.set_auto_filter("A1:E1")
    s.write_row(["region", "product", "month", "units", "revenue"], styles=[hdr] * 5)
    for r in sales:
        s.write_row(r)
    t = w.add_sheet("Targets")
    t.freeze_panes(rows=1, cols=0)
    t.write_row(["region", "quarter", "target_revenue"], styles=[hdr] * 3)
    for r in targets:
        t.write_row(r)

print(f"{OUT}: {len(sales)} sales rows, {len(targets)} target rows")
