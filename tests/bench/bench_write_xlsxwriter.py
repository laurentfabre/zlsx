"""Matching xlsxwriter benchmark — 1000 rows × 10 cols + header style."""
import sys
import xlsxwriter


def main() -> None:
    out = sys.argv[1]
    wb = xlsxwriter.Workbook(out, {"constant_memory": True})
    sheet = wb.add_worksheet("Bench")

    header_fmt = wb.add_format({
        "bold": True,
        "font_color": "#FFFFFF",
        "bg_color": "#1E3A8A",
        "align": "center",
    })
    sheet.set_column(0, 0, 20)
    sheet.freeze_panes(1, 0)

    header = ["Name", "Amount", "Share", "Qty", "Active", "Code",
              "A", "B", "C", "D"]
    for c, v in enumerate(header):
        sheet.write(0, c, v, header_fmt)

    money_fmt = wb.add_format({"num_format": "$#,##0.00"})
    pct_fmt = wb.add_format({"num_format": "0.00%"})

    for i in range(1000):
        r = i + 1
        sheet.write_string(r, 0, f"row_{i}")
        sheet.write_number(r, 1, 100.0 + i * 1.5, money_fmt)
        sheet.write_number(r, 2, i / 1000.0, pct_fmt)
        sheet.write_number(r, 3, i)
        sheet.write_boolean(r, 4, i % 2 == 0)
        sheet.write_string(r, 5, f"CODE_{i:x}")
        sheet.write_number(r, 6, i * 7)
        sheet.write_number(r, 7, i * 0.1)
        sheet.write_string(r, 8, "x")

    wb.close()


if __name__ == "__main__":
    main()
