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

    use office_ppt::write::{write_from_slide_template, write_presentation, SlideSpec};
    let deck = vec![
        SlideSpec::new("WRITER-PPT-001").with_body(&["first body line", "second line"]),
        SlideSpec::new("WRITER-PPT-002").with_body(&["第二页正文"]),
    ];
    // Self-contained: the template is baked into the crate, and each slide's
    // shape tree is generated (docs/ppt.md §9.24).
    let scratch = write_presentation(&deck).unwrap();
    std::fs::write("samples/written-scratch.ppt", &scratch).unwrap();
    println!("samples/written-scratch.ppt ({} bytes)", scratch.len());

    // Speaker notes: one notes page per slide that has them (docs §9.25).
    let noted = vec![
        SlideSpec::new("WRITER-PPT-NOTES")
            .with_body(&["a slide with notes"])
            .with_notes(&["remember to mention the shape tree"]),
        SlideSpec::new("WRITER-PPT-002").with_body(&["no notes on this one"]),
        SlideSpec::new("WRITER-PPT-003")
            .with_body(&["two note paragraphs"])
            .with_notes(&["first note paragraph", "second note paragraph"]),
    ];
    let notes = write_presentation(&noted).unwrap();
    std::fs::write("samples/written-notes.ppt", &notes).unwrap();
    println!("samples/written-notes.ppt ({} bytes)", notes.len());

    if let Ok(template) = std::fs::read("samples/ms/minimal.ppt") {
        let ppt = write_from_slide_template(&template, &deck).unwrap();
        std::fs::write("samples/written.ppt", &ppt).unwrap();
        println!("samples/written.ppt ({} bytes)", ppt.len());
    }

    if let Ok(png) = std::fs::read("samples/pixel.png") {
        let deck = vec![
            SlideSpec::new("WRITER-PPT-PIC")
                .with_body(&["a picture below"])
                .with_picture(png.clone(), 96, 96),
            SlideSpec::new("Second page").with_picture(png, 48, 48),
        ];
        let ppt = write_presentation(&deck).unwrap();
        std::fs::write("samples/written-picture.ppt", &ppt).unwrap();
        println!("samples/written-picture.ppt ({} bytes)", ppt.len());
    }

    use office_xls::{Cell, CellValue};
    let cells = vec![
        Cell { row: 0, col: 0, xf: 15, value: CellValue::Text("WRITER-XLS-001".into()) },
        Cell { row: 1, col: 0, xf: 15, value: CellValue::Number(1234.5) },
        Cell { row: 2, col: 0, xf: 15, value: CellValue::Number(42.0) },
        Cell { row: 3, col: 0, xf: 15, value: CellValue::Bool(true) },
    ];
    // A table: paragraphs whose mark is the cell mark, plus a row-end mark
    // carrying TDefTable.
    let table = office_doc::Table::new(&[&["H1", "H2", "H3"], &["a", "b", "c"]]);
    let doc = office_doc::write_table_document("Before table.", &table, "After table.");
    std::fs::write("samples/written-table.doc", &doc).unwrap();
    println!("samples/written-table.doc ({} bytes)", doc.len());

    // A document with an inline picture: the PNG goes into the `Data` stream and
    // the text carries a picture mark.
    if let Ok(png) = std::fs::read("samples/pixel.png") {
        let paragraphs = vec![
            office_doc::write::Paragraph::text("PICTURE-DOC-777 before the image"),
            office_doc::write::Paragraph::styled(
                vec![office_doc::write::Run::picture(
                    office_doc::write::InlinePicture::new(png.clone(), 480, 480),
                )],
                office_doc::fmt::ParaProps::default(),
            ),
            office_doc::write::Paragraph::text("after the image"),
        ];
        let doc = office_doc::write::write_document(&paragraphs);
        std::fs::write("samples/written-picture.doc", &doc).unwrap();
        println!("samples/written-picture.doc ({} bytes)", doc.len());
    }

    let xls = office_xls::write_minimal("Writer", &cells);
    std::fs::write("samples/written.xls", &xls).unwrap();
    println!("samples/written.xls ({} bytes)", xls.len());

    // Formulas: the text is compiled into a real `rgce` token array, so Excel
    // recalculates instead of showing a frozen number.
    use office_xls::CachedValue;
    let formula = |row: u16, text: &str, cached: f64| Cell {
        row,
        col: 1,
        xf: 15,
        value: CellValue::Formula {
            cached: CachedValue::Number(cached),
            rgce: Vec::new(),
            text: text.into(),
        },
    };
    let cells = vec![
        Cell { row: 0, col: 0, xf: 15, value: CellValue::Number(10.0) },
        Cell { row: 1, col: 0, xf: 15, value: CellValue::Number(20.0) },
        Cell { row: 2, col: 0, xf: 15, value: CellValue::Number(1234.5678) },
        formula(0, "A1+A2", 30.0),
        formula(1, "A1*A2", 200.0),
        formula(2, "ROUND(A3,1)", 1234.6),
        formula(3, "IF(A1>5,\"big\",\"small\")", 0.0),
        formula(4, "SUM(A1:A2)/2", 15.0),
    ];
    let xls = office_xls::write_minimal("Formulas", &cells);
    std::fs::write("samples/written-formulas.xls", &xls).unwrap();
    println!("samples/written-formulas.xls ({} bytes)", xls.len());

    let cells = [
        Cell { row: 0, col: 0, xf: 0, value: CellValue::Number(45356.0) },
        Cell { row: 1, col: 0, xf: 1, value: CellValue::Number(0.25) },
        Cell { row: 2, col: 0, xf: 2, value: CellValue::Number(1234.5) },
        Cell { row: 3, col: 0, xf: 3, value: CellValue::Number(-42.0) },
    ];
    let xls = office_xls::write_with_formats(
        "Formats",
        &cells,
        &["yyyy-mm-dd", "0.00%", "#,##0.00", "0.00"],
    );
    std::fs::write("samples/written-formats.xls", &xls).unwrap();
    println!("samples/written-formats.xls ({} bytes)", xls.len());

    // Sheet furniture: a merged title, two column widths, two row heights and
    // frozen panes (docs/xls.md §15.4).
    let cells = [
        Cell { row: 0, col: 0, xf: 15, value: CellValue::Text("WRITER-XLS-LAYOUT".into()) },
        Cell { row: 2, col: 0, xf: 15, value: CellValue::Text("left".into()) },
        Cell { row: 2, col: 1, xf: 15, value: CellValue::Number(42.0) },
        Cell { row: 2, col: 2, xf: 15, value: CellValue::Text("wide".into()) },
    ];
    let layout = office_xls::write::Layout::default()
        .merge(0, 0, 0, 2)
        .column(0, 20.0)
        .column(1, 5.0)
        .column(2, 30.0)
        .row(0, 40.0)
        .row(2, 22.5)
        .freeze(1, 1);
    let xls = office_xls::write::write_with_layout("Layout", &cells, &[], &layout);
    std::fs::write("samples/written-layout.xls", &xls).unwrap();
    println!("samples/written-layout.xls ({} bytes)", xls.len());
}
