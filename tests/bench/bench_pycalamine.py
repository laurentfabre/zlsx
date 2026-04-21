"""Same benchmark workload for python-calamine."""
import sys
from python_calamine import CalamineWorkbook

def main() -> None:
    path = sys.argv[1]
    wb = CalamineWorkbook.from_path(path)
    sheet = wb.get_sheet_by_index(0)
    rows = sheet.to_python()
    n_rows = n_str = n_int = n_num = n_bool = n_empty = 0
    for row in rows:
        n_rows += 1
        for c in row:
            if c is None or c == "":
                n_empty += 1
            elif isinstance(c, bool):
                n_bool += 1
            elif isinstance(c, int):
                n_int += 1
            elif isinstance(c, float):
                n_num += 1
            else:
                n_str += 1
    print(f"rows={n_rows} str={n_str} int={n_int} num={n_num} bool={n_bool} empty={n_empty}")

if __name__ == "__main__":
    main()
