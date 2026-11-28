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

    // A document with a repeating header and footer.
    {
        let body = vec![
            office_doc::write::Paragraph::text("HEADER-DOC-123 body text."),
            office_doc::write::Paragraph::text("Second body paragraph."),
        ];
        let header = vec![office_doc::write::Paragraph::text("Draft header")];
        let footer = vec![office_doc::write::Paragraph::text("Draft footer")];
        let doc = office_doc::write::write_document_with(
            &body,
            &office_doc::write::HeaderFooter { header: &header, footer: &footer },
        );
        std::fs::write("samples/written-header.doc", &doc).unwrap();
        println!("samples/written-header.doc ({} bytes)", doc.len());
    }

    // A document with footnotes: the body carries `0x02` marks where the notes
    // are referenced, and the notes themselves ride in the footnote story.
    {
        use office_doc::write::{write_document_with_footnotes, Documents, Paragraph, FOOTNOTE_MARK};
        let body = vec![
            Paragraph::text(format!("WRITER-DOC-001 carries a note.{FOOTNOTE_MARK}")),
            Paragraph::text(format!("WRITER-DOC-002 has the second one.{FOOTNOTE_MARK}")),
            Paragraph::text("No note on this paragraph."),
        ];
        let notes = vec![
            vec![Paragraph::text("note one")],
            vec![Paragraph::text("note two, with more words")],
        ];
        let doc = write_document_with_footnotes(&body, &Documents::default().notes(&notes));
        std::fs::write("samples/written-footnotes.doc", &doc).unwrap();
        println!("samples/written-footnotes.doc ({} bytes)", doc.len());
    }

    // A document with both kinds of note: `ENDNOTE_MARK` in the source becomes
    // the same `0x02` in the stream, and only the PLC pairs tell the two apart.
    {
        use office_doc::write::{
            write_document_with_footnotes, Documents, Paragraph, ENDNOTE_MARK, FOOTNOTE_MARK,
        };
        let body = vec![
            Paragraph::text(format!("WRITER-DOC-003 has a footnote.{FOOTNOTE_MARK}")),
            Paragraph::text(format!("WRITER-DOC-004 has an endnote.{ENDNOTE_MARK}")),
            Paragraph::text("No note on this paragraph."),
        ];
        let notes = vec![vec![Paragraph::text("footnote text")]];
        let endnotes = vec![vec![Paragraph::text("endnote text")]];
        let documents = Documents::default().notes(&notes).endnotes(&endnotes);
        let doc = write_document_with_footnotes(&body, &documents);
        std::fs::write("samples/written-endnotes.doc", &doc).unwrap();
        println!("samples/written-endnotes.doc ({} bytes)", doc.len());

        // The same document with a header and a footer as well: the header
        // story sits between the footnote and endnote stories, so this is the
        // combination where the `ccp*` arithmetic has to be exactly right.
        let header = vec![Paragraph::text("Draft header")];
        let footer = vec![Paragraph::text("Draft footer")];
        let documents = Documents::new(&header, &footer).notes(&notes).endnotes(&endnotes);
        let doc = write_document_with_footnotes(&body, &documents);
        std::fs::write("samples/written-notes-header.doc", &doc).unwrap();
        println!("samples/written-notes-header.doc ({} bytes)", doc.len());
    }

    // A workbook with a print setup: A4 landscape, custom margins and a
    // centred header and footer.
    {
        use office_xls::write::{write_workbook, SheetData};
        let cells = vec![
            office_xls::Cell {
                row: 0,
                col: 0,
                xf: 15,
                value: office_xls::CellValue::Text("PRINT-XLS-001".into()),
            },
            office_xls::Cell { row: 1, col: 0, xf: 15, value: office_xls::CellValue::Number(42.0) },
        ];
        let mut print = office_xls::PrintSetup::a4();
        print.landscape = true;
        print.left_margin = 0.5;
        print.right_margin = 0.75;
        print.top_margin = 0.25;
        print.bottom_margin = 1.5;
        print.header = "&CPRINT-HEADER".into();
        print.footer = "&CPRINT-FOOTER".into();
        let sheet = SheetData::new("Print", &cells).with_print(&print);
        let bytes = write_workbook(&[sheet]);
        std::fs::write("samples/written-print.xls", &bytes).unwrap();
        println!("samples/written-print.xls ({} bytes)", bytes.len());
    }

    // A document on a Letter page with one-inch margins all round: the page
    // setup lives in the section properties (`SEPX`), which is why this needs
    // `write_document_with_footnotes` rather than `write_document`.
    {
        use office_doc::write::{write_document_with_footnotes, Documents};
        let paragraphs = vec![
            office_doc::write::Paragraph::text("PAGE-DOC-001 sits on a Letter page."),
            office_doc::write::Paragraph::text("Margins are one inch all round."),
        ];
        let page = office_doc::PageSetup::letter().margins(1440, 1440, 1440, 1440);
        let documents = Documents::default().page_setup(page);
        let doc = write_document_with_footnotes(&paragraphs, &documents);
        std::fs::write("samples/written-page.doc", &doc).unwrap();
        println!("samples/written-page.doc ({} bytes)", doc.len());
    }

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

    // Several sheets in one workbook, each with its own cells, formats and
    // furniture (docs/xls.md §17).
    let first = [
        // `xf` indexes the sheet's `formats`; 15 is past the end, so the
        // title stays General (docs/xls.md §17).
        Cell { row: 0, col: 0, xf: 15, value: CellValue::Text("WRITER-XLS-SHEETS".into()) },
        Cell { row: 1, col: 0, xf: 0, value: CellValue::Number(45356.0) },
        Cell { row: 2, col: 0, xf: 1, value: CellValue::Number(0.25) },
    ];
    let second = [
        Cell { row: 0, col: 0, xf: 0, value: CellValue::Text("second sheet".into()) },
        Cell { row: 1, col: 0, xf: 0, value: CellValue::Bool(true) },
    ];
    let third = [Cell { row: 0, col: 0, xf: 0, value: CellValue::Number(3.0) }];
    let xls = office_xls::write::write_workbook(&[
        office_xls::write::SheetData::new("First", &first)
            .with_formats(&["yyyy-mm-dd", "0.00%"]),
        office_xls::write::SheetData::new("Second", &second),
        office_xls::write::SheetData::new("Third", &third)
            .with_layout(office_xls::write::Layout::default().column(0, 24.0)),
    ]);
    std::fs::write("samples/written-sheets.xls", &xls).unwrap();
    println!("samples/written-sheets.xls ({} bytes)", xls.len());

    // Fonts: a bold header, an italic line, a larger face and a red one
    // (docs/xls.md §18).
    use office_xls::write::XfSpec;
    use office_xls::Font;
    let cells = [
        Cell { row: 0, col: 0, xf: 0, value: CellValue::Text("WRITER-XLS-STYLES".into()) },
        Cell { row: 1, col: 0, xf: 1, value: CellValue::Text("italic 14pt".into()) },
        Cell { row: 2, col: 0, xf: 2, value: CellValue::Text("red".into()) },
        Cell { row: 3, col: 0, xf: 3, value: CellValue::Number(0.25) },
    ];
    let xfs = [
        XfSpec::plain().font(Font { bold: true, size: 14.0, ..Font::default_face() }),
        XfSpec::plain().font(Font { italic: true, size: 14.0, ..Font::default_face() }),
        XfSpec::plain().font(Font { colour: Some((0xC0, 0x00, 0x00)), ..Font::default_face() }),
        XfSpec::format("0.00%").font(Font { bold: true, ..Font::default_face() }),
    ];
    let xls = office_xls::write::write_workbook(&[office_xls::write::SheetData::new("Styles", &cells)
        .with_xfs(&xfs)
        .with_layout(office_xls::write::Layout::default().column(0, 22.0))]);
    std::fs::write("samples/written-styles.xls", &xls).unwrap();
    println!("samples/written-styles.xls ({} bytes)", xls.len());

    // Borders and alignment (docs/xls.md §19).
    use office_xls::{Align, Border, BorderStyle, BoxBorder, VAlign};
    let cells = [
        Cell { row: 0, col: 0, xf: 0, value: CellValue::Text("boxed".into()) },
        Cell { row: 2, col: 0, xf: 1, value: CellValue::Text("thick".into()) },
        Cell { row: 4, col: 0, xf: 2, value: CellValue::Text("underline".into()) },
        Cell { row: 6, col: 0, xf: 3, value: CellValue::Text("centred and wrapped text".into()) },
        Cell { row: 8, col: 0, xf: 4, value: CellValue::Text("indented".into()) },
        Cell { row: 10, col: 0, xf: 5, value: CellValue::Text("red edge".into()) },
        Cell { row: 0, col: 2, xf: 6, value: CellValue::Text("diagonal".into()) },
    ];
    let xfs = [
        XfSpec::plain().border(BoxBorder::thin_box()),
        XfSpec::plain().border(BoxBorder::all(Border::new(BorderStyle::Thick))),
        XfSpec::plain().border(BoxBorder {
            bottom: Border::new(BorderStyle::Double),
            ..BoxBorder::default()
        }),
        XfSpec::plain().align(Align::Centre).valign(VAlign::Centre).wrap(true),
        XfSpec::plain().indent(2),
        XfSpec::plain().border(BoxBorder::all(Border::thin().colour((0xFF, 0x00, 0x00)))),
        XfSpec::plain().border(BoxBorder {
            diagonal: Border::new(BorderStyle::Thin),
            diagonal_down: true,
            ..BoxBorder::default()
        }),
    ];
    let xls = office_xls::write::write_workbook(&[office_xls::write::SheetData::new(
        "Borders", &cells,
    )
    .with_xfs(&xfs)
    .with_layout(office_xls::write::Layout::default().column(0, 26.0).column(2, 12.0))]);
    std::fs::write("samples/written-borders.xls", &xls).unwrap();
    println!("samples/written-borders.xls ({} bytes)", xls.len());

    // ---- .xls cell fills ---------------------------------------------------
    use office_xls::write::{Fill, Pattern};
    let cells = [
        Cell { row: 0, col: 0, xf: 0, value: CellValue::Text("solid red".into()) },
        Cell { row: 2, col: 0, xf: 1, value: CellValue::Text("grey 50%".into()) },
        Cell { row: 4, col: 0, xf: 2, value: CellValue::Text("dark grid".into()) },
        Cell { row: 6, col: 0, xf: 3, value: CellValue::Text("light trellis".into()) },
        Cell { row: 8, col: 0, xf: 4, value: CellValue::Text("no fill".into()) },
    ];
    let xfs = [
        XfSpec::plain().fill(Fill::solid((0xFF, 0x00, 0x00))),
        XfSpec::plain().fill(Fill {
            pattern: Pattern::Gray50,
            fore: Some((0xC0, 0xC0, 0xC0)),
            back: Some((0xFF, 0xFF, 0xFF)),
        }),
        XfSpec::plain().fill(Fill {
            pattern: Pattern::DarkGrid,
            fore: Some((0x00, 0x00, 0x80)),
            back: None,
        }),
        XfSpec::plain().fill(Fill {
            pattern: Pattern::LightTrellis,
            fore: Some((0xFF, 0xFF, 0x00)),
            back: Some((0x00, 0x80, 0x00)),
        }),
        XfSpec::plain(),
    ];
    let xls = office_xls::write::write_workbook(&[office_xls::write::SheetData::new(
        "Fills", &cells,
    )
    .with_xfs(&xfs)
    .with_layout(office_xls::write::Layout::default().column(0, 26.0))]);
    std::fs::write("samples/written-fills.xls", &xls).unwrap();
    println!("samples/written-fills.xls ({} bytes)", xls.len());

    // ---- .xls hyperlinks ---------------------------------------------------
    use office_xls::write::Link;
    let cells = [
        Cell { row: 0, col: 0, xf: 0, value: CellValue::Text("links below".into()) },
        Cell { row: 1, col: 1, xf: 0, value: CellValue::Text("example.com".into()) },
        Cell { row: 3, col: 1, xf: 0, value: CellValue::Text("mail us".into()) },
    ];
    let links = [
        Link { row: 1, col: 1, url: "https://example.com/page", text: "example.com" },
        Link { row: 3, col: 1, url: "mailto:someone@example.com", text: "mail us" },
    ];
    let xls = office_xls::write::write_workbook(&[office_xls::write::SheetData::new(
        "Links", &cells,
    )
    .with_links(&links)]);
    std::fs::write("samples/written-links.xls", &xls).unwrap();
    println!("samples/written-links.xls ({} bytes)", xls.len());

    // ---- .xls data validation ---------------------------------------------
    use office_xls::write::ValidationSpec;
    let cells = [
        Cell { row: 0, col: 0, xf: 0, value: CellValue::Text("green".into()) },
        Cell { row: 2, col: 0, xf: 0, value: CellValue::Number(4.0) },
        Cell { row: 4, col: 0, xf: 0, value: CellValue::Text("beta".into()) },
        Cell { row: 0, col: 3, xf: 0, value: CellValue::Text("alpha".into()) },
        Cell { row: 1, col: 3, xf: 0, value: CellValue::Text("beta".into()) },
        Cell { row: 2, col: 3, xf: 0, value: CellValue::Text("gamma".into()) },
    ];
    let choices = ["red", "green", "blue"];
    let rules = [
        ValidationSpec::list((0, 0), (0, 0), &choices)
            .prompt("Colour", "pick one of three")
            .error("Not a colour", "red, green or blue"),
        ValidationSpec::whole_between((2, 0), (2, 0), "1", "10"),
        ValidationSpec::list_from_range((4, 0), (4, 0), "D1:D3"),
    ];
    let xls = office_xls::write::write_workbook(&[office_xls::write::SheetData::new(
        "Validation", &cells,
    )
    .with_validations(&rules)]);
    std::fs::write("samples/written-validation.xls", &xls).unwrap();
    println!("samples/written-validation.xls ({} bytes)", xls.len());

    // ---- .xls AutoFilter ---------------------------------------------------
    use office_xls::AutoFilter;
    let cells = [
        Cell { row: 0, col: 0, xf: 0, value: CellValue::Text("name".into()) },
        Cell { row: 0, col: 1, xf: 0, value: CellValue::Text("kind".into()) },
        Cell { row: 0, col: 2, xf: 0, value: CellValue::Text("size".into()) },
        Cell { row: 1, col: 0, xf: 0, value: CellValue::Text("alpha".into()) },
        Cell { row: 1, col: 1, xf: 0, value: CellValue::Text("red".into()) },
        Cell { row: 1, col: 2, xf: 0, value: CellValue::Number(10.0) },
        Cell { row: 2, col: 0, xf: 0, value: CellValue::Text("beta".into()) },
        Cell { row: 2, col: 1, xf: 0, value: CellValue::Text("green".into()) },
        Cell { row: 2, col: 2, xf: 0, value: CellValue::Number(20.0) },
        Cell { row: 3, col: 0, xf: 0, value: CellValue::Text("gamma".into()) },
        Cell { row: 3, col: 1, xf: 0, value: CellValue::Text("blue".into()) },
        Cell { row: 3, col: 2, xf: 0, value: CellValue::Number(30.0) },
    ];
    let xls = office_xls::write::write_workbook(&[office_xls::write::SheetData::new(
        "Filtered", &cells,
    )
    .with_filter(AutoFilter { first: (0, 0), last: (3, 2) })]);
    std::fs::write("samples/written-filter.xls", &xls).unwrap();
    println!("samples/written-filter.xls ({} bytes)", xls.len());
}
