//! Regenerate the files that the *writers* produce.
//!
//! Usage: `cargo run --example mkfiles`
//!
//! The files are produced from scratch by `office_doc::write_document` and
//! `office_xls::write_minimal`; `tools/roundtrip.ps1` opens them in Word and
//! Excel to prove the bytes are valid, not merely self-consistent.
use office_doc::fmt::{CharProps, ParaProps};
use office_doc::write::{write_document, Paragraph, Run};

fn main() {
    // Plain text.
    let text = "WRITER-DOC-001\rSecond line from the writer.\rEND-WRITER-9\r";
    let doc = office_doc::write_minimal(text);
    std::fs::write("samples/written.doc", &doc).unwrap();
    println!("samples/written.doc ({} bytes)", doc.len());

    // The same, with character and paragraph formatting.
    let red_bold = CharProps {
        bold: Some(true),
        size_half_points: Some(36),
        color: Some(0x0000_00FF),
        ..CharProps::default()
    };
    let italic = CharProps {
        italic: Some(true),
        underline: Some(1),
        ..CharProps::default()
    };
    let paras = vec![
        Paragraph::styled(
            vec![Run::styled("Formatted by our writer", red_bold)],
            ParaProps {
                justify: Some(1),
                ..ParaProps::default()
            },
        ),
        Paragraph::styled(
            vec![
                Run::new("plain "),
                Run::styled("italic", italic),
                Run::new(" plain"),
            ],
            ParaProps {
                dxa_left: Some(720),
                dxa_left_first: Some(360),
                ..ParaProps::default()
            },
        ),
        Paragraph::text("WRITER-FMT-END"),
    ];
    let doc = write_document(&paras);
    std::fs::write("samples/written-formatted.doc", &doc).unwrap();
    println!("samples/written-formatted.doc ({} bytes)", doc.len());

    // Enough paragraphs to need several formatting pages, alternating bold.
    let bold = CharProps { bold: Some(true), ..CharProps::default() };
    let mut long: Vec<Paragraph> = Vec::new();
    for i in 0..120 {
        let text = format!("paragraph {i:03}");
        long.push(if i % 3 == 0 {
            Paragraph::styled(vec![Run::styled(text, bold.clone())], ParaProps::default())
        } else {
            Paragraph::text(text)
        });
    }
    let doc = write_document(&long);
    std::fs::write("samples/written-long.doc", &doc).unwrap();
    println!("samples/written-long.doc ({} bytes)", doc.len());

    use office_ppt::write::{write_from_slide_template, SlideSpec};
    let deck = vec![
        SlideSpec::new("WRITER-PPT-001").with_body(&["first body line", "second line"]),
        SlideSpec::new("WRITER-PPT-002").with_body(&["第二页正文"]),
    ];
    // The deck is cloned from a template PowerPoint itself wrote: from-scratch
    // output passes our own reader but PowerPoint rejects it.
    let template = std::fs::read("samples/ms/minimal.ppt").unwrap();
    let ppt = write_from_slide_template(&template, &deck).unwrap();
    std::fs::write("samples/written.ppt", &ppt).unwrap();
    println!("samples/written.ppt ({} bytes)", ppt.len());

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
