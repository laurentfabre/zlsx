use calamine::{open_workbook_auto, Data, Reader};
use std::env;

fn main() {
    let path = env::args().nth(1).expect("usage: bench <xlsx>");
    let mut wb = open_workbook_auto(&path).expect("open");
    let sheet_name = wb.sheet_names().first().cloned().expect("no sheets");
    let range = wb.worksheet_range(&sheet_name).expect("read range");
    let (mut n_rows, mut n_str, mut n_int, mut n_num, mut n_bool, mut n_empty) = (0usize, 0, 0, 0, 0, 0);
    for row in range.rows() {
        n_rows += 1;
        for c in row {
            match c {
                Data::Empty => n_empty += 1,
                Data::String(_) | Data::DateTimeIso(_) | Data::DurationIso(_) => n_str += 1,
                Data::Int(_) => n_int += 1,
                Data::Float(_) => n_num += 1,
                Data::Bool(_) => n_bool += 1,
                Data::DateTime(_) => n_num += 1,
                Data::Error(_) => n_str += 1,
            }
        }
    }
    println!("rows={n_rows} str={n_str} int={n_int} num={n_num} bool={n_bool} empty={n_empty}");
}
