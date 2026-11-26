//! Regenerate the files that the *writers* produce.
//!
//! Usage: `cargo run --example mkfiles`
//!
//! `samples/written.doc` and `samples/written.xls` are produced from scratch by
//! `office_doc::write_minimal` and `office_xls::write_minimal`; `tools/roundtrip.ps1`
//! opens them in Word and Excel to prove the bytes are valid, not merely
//! self-consistent.
fn main() {
    let text = "WRITER-DOC-001\rSecond line from the writer.\rEND-WRITER-9\r";
    let doc = office_doc::write_minimal(text);
    std::fs::write("samples/written.doc", &doc).unwrap();
    println!("samples/written.doc ({} bytes)", doc.len());

    use office_xls::{Cell, CellValue};
    let cells = vec![
        Cell { row: 0, col: 0, xf: 15, value: CellValue::Text("WRITER-XLS-001".into()) },
        Cell { row: 1, col: 0, xf: 15, value: CellValue::Number(1234.5) },
        Cell { row: 2, col: 0, xf: 15, value: CellValue::Number(42.0) },
        Cell { row: 3, col: 0, xf: 15, value: CellValue::Bool(true) },
    ];
    let xls = office_xls::write_minimal("Writer", &cells);
    std::fs::write("samples/written.xls", &xls).unwrap();
    println!("samples/written.xls ({} bytes)", xls.len());
}
