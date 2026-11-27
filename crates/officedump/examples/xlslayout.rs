//! Write a one-sheet workbook whose column widths and row heights come from the
//! command line, so `ColumnWidth` / `RowHeight` read-back can be measured:
//!
//! ```text
//! cargo run -q -p officedump --example xlslayout -- /tmp/w.xls 10 20 50
//! ```
use office_xls::write::Layout;
use office_xls::{Cell, CellValue};

fn main() {
    let args: Vec<String> = std::env::args().skip(1).collect();
    if args.len() < 2 {
        eprintln!("usage: xlslayout OUT.xls WIDTH...");
        std::process::exit(2);
    }
    let out = &args[0];
    let widths: Vec<f64> = args[1..].iter().filter_map(|s| s.parse().ok()).collect();
    let mut layout = Layout::default();
    let mut cells = Vec::new();
    for (i, width) in widths.iter().enumerate() {
        let col = i as u16;
        layout.columns.push((col, *width));
        layout.rows.push((col, *width)); // same number as a row height, in points
        cells.push(Cell { row: col, col, xf: 15, value: CellValue::Text(format!("{width}")) });
    }
    let bytes = office_xls::write::write_with_layout("Widths", &cells, &[], &layout);
    std::fs::write(out, &bytes).unwrap();
    println!("{out} ({} bytes)", bytes.len());
}
