//! End-to-end checks against real files: our own WPS-written samples plus, when
//! present, the Microsoft-authored files shipped with Office.
//!
//! Every test skips silently when the file is missing, so the suite still works
//! on a machine without Office installed.

use std::path::PathBuf;

fn sample(name: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../samples")
        .join(name)
}

fn office16(rel: &str) -> PathBuf {
    PathBuf::from(r"C:\Program Files\Microsoft Office\root\Office16").join(rel)
}

/// The snippet from docs/README.md: one entry point for all three formats.
fn read_any(path: &std::path::Path) -> Result<String, Box<dyn std::error::Error>> {
    let bytes = std::fs::read(path)?;
    let cfb = office_cfb::Cfb::parse(bytes.clone())?;
    if cfb.stream("WordDocument").is_ok() {
        Ok(office_doc::parse(&bytes)?.plain_text())
    } else if cfb.stream("Workbook").is_ok() || cfb.stream("Book").is_ok() {
        Ok(office_xls::parse(&bytes)?.plain_text())
    } else {
        let doc = cfb.stream("PowerPoint Document")?;
        let user = cfb.stream("Current User").ok();
        Ok(office_ppt::parse_stream(&doc, user.as_deref())?.plain_text())
    }
}

#[test]
fn doc_sample_roundtrips() {
    let p = sample("hello.doc");
    if !p.exists() {
        return;
    }
    let doc = office_doc::parse(&std::fs::read(&p).unwrap()).unwrap();
    assert_eq!(doc.fib.n_fib, 0x00C1);
    assert_eq!(doc.fib.table_stream, "0Table");
    assert_eq!(doc.fib.ccp_text, 82);
    assert_eq!(doc.pieces.len(), 1);
    assert!(!doc.pieces[0].compressed, "hello.doc stores UTF-16");
    let text = doc.plain_text();
    assert!(text.contains("HELLO-DOC-001"), "{text}");
    assert!(text.contains("中文测试 Ω≈ç√"), "{text}");
    assert!(text.contains("END-MARKER-42"), "{text}");
}

#[test]
fn xls_sample_cells_and_formula_cache() {
    use office_xls::CellValue;

    let p = sample("hello.xls");
    if !p.exists() {
        return;
    }
    let wb = office_xls::parse(&std::fs::read(&p).unwrap()).unwrap();
    assert_eq!(wb.biff_version, 0x0600);
    assert_eq!(wb.sst.len(), 3);

    let data = wb.sheet("Data").expect("Data sheet");
    let a1 = data.cells.iter().find(|c| (c.row, c.col) == (0, 0)).unwrap();
    assert_eq!(a1.value, CellValue::Text("HELLO-XLS-001".into()));
    let a3 = data.cells.iter().find(|c| (c.row, c.col) == (2, 0)).unwrap();
    assert_eq!(a3.value, CellValue::Number(1234.5));
    let b3 = data.cells.iter().find(|c| (c.row, c.col) == (2, 1)).unwrap();
    match &b3.value {
        CellValue::Formula { cached, .. } => {
            assert_eq!(*cached, office_xls::CachedValue::Number(3.0))
        }
        other => panic!("B3 should be a formula, got {other:?}"),
    }
    // BOUNDSHEET order is tab order, not physical order.
    assert_eq!(wb.sheets[0].name, "Sheet2");
    assert_eq!(wb.sheets[1].name, "Data");
}

#[test]
fn ppt_sample_slides_in_persist_order() {
    use office_ppt::TextKind;

    let p = sample("hello.ppt");
    if !p.exists() {
        return;
    }
    let bytes = std::fs::read(&p).unwrap();
    let cfb = office_cfb::Cfb::parse(bytes).unwrap();
    let data = cfb.stream("PowerPoint Document").unwrap();
    let user = cfb.stream("Current User").ok();
    let deck = office_ppt::parse_stream(&data, user.as_deref()).unwrap();

    assert!(deck.ordered_by_persist_directory);
    assert_eq!(deck.slides.len(), 2);
    assert_eq!(deck.slides[0].texts[0].kind, TextKind::Title);
    assert_eq!(deck.slides[0].texts[0].text, "HELLO-PPT-001");
    assert_eq!(deck.slides[1].texts[0].text, "Second slide");
    assert!(deck.plain_text().contains("中文幻灯片内容"));
}

#[test]
fn read_any_handles_all_three() {
    for name in ["hello.doc", "hello.xls", "hello.ppt"] {
        let p = sample(name);
        if !p.exists() {
            continue;
        }
        let text = read_any(&p).unwrap_or_else(|e| panic!("{name}: {e}"));
        assert!(!text.trim().is_empty(), "{name} produced no text");
    }
}

#[test]
fn microsoft_word_template() {
    let p = office16(r"2052\PROTTPLN.DOC");
    if !p.exists() {
        return;
    }
    let doc = office_doc::parse(&std::fs::read(&p).unwrap()).unwrap();
    // Microsoft writes this one with the 1Table stream, WPS uses 0Table.
    assert_eq!(doc.fib.table_stream, "1Table");
    assert!(doc.plain_text().contains("权限"), "{}", doc.plain_text());
}

#[test]
fn microsoft_excel_samples() {
    let p = office16(r"SAMPLES\SOLVSAMP.XLS");
    if !p.exists() {
        return;
    }
    let wb = office_xls::parse(&std::fs::read(&p).unwrap()).unwrap();
    assert_eq!(wb.sheets.len(), 6);
    assert!(wb.sst.len() > 100, "sst={}", wb.sst.len());
    let mix = wb.sheet("Product Mix").unwrap();
    assert!(mix.cells.len() > 50);
}

#[test]
fn microsoft_powerpoint_template() {
    let p = office16(r"2052\PROTTPLN.PPT");
    if !p.exists() {
        return;
    }
    let bytes = std::fs::read(&p).unwrap();
    let cfb = office_cfb::Cfb::parse(bytes).unwrap();
    let data = cfb.stream("PowerPoint Document").unwrap();
    let user = cfb.stream("Current User").ok();
    let deck = office_ppt::parse_stream(&data, user.as_deref()).unwrap();
    assert_eq!(deck.slides.len(), 1);
    assert!(deck.ordered_by_persist_directory);
}

// --- formatting, formulas, stories, pictures -------------------------------

#[test]
fn doc_character_formatting_survives_the_fkp_walk() {
    let p = sample("formatted.doc");
    if !p.exists() {
        return;
    }
    let doc = office_doc::parse(&std::fs::read(&p).unwrap()).unwrap();
    let paras = doc.paragraphs_rich();
    assert_eq!(paras.len(), 5, "five paragraphs, the trailing mark dropped");

    let first = &paras[0];
    assert_eq!(first.text, "Bold red heading");
    let run = first.runs.first().expect("a formatted run");
    assert_eq!(run.props.bold, Some(true));
    assert_eq!(run.props.size_half_points, Some(36)); // 18 pt
    assert_eq!(run.props.color, Some(0x0000_00FF)); // RGB(255,0,0) as 0x00BBGGRR

    let second = &paras[1];
    assert!(second
        .runs
        .iter()
        .any(|r| r.props.italic == Some(true) && r.props.underline == Some(1)));
    // The "mixed " span is bold on top of italic+underline.
    assert!(second
        .runs
        .iter()
        .any(|r| r.props.bold == Some(true)
            && r.props.italic == Some(true)
            && r.text(&second.text) == "mixed "));

    // Paragraph properties live in the PAPX page: one run per paragraph mark.
    // Every run must have its *own* `BX`, which is a 13-byte entry — reading
    // the offset array with a stride of 1 gives every run the first run's
    // properties and silently loses all of these.
    assert_eq!(doc.para_runs.len(), 6);
    let paras = doc.paragraphs_rich();
    assert_eq!(paras.len(), 5);
    assert_eq!(paras[2].text, "Centered paragraph");
    assert_eq!(paras[2].props.justify, Some(1), "paragraph 3 is centred");
    assert_eq!(paras[3].text, "Indented paragraph");
    assert_eq!(paras[3].props.dxa_left, Some(720), "0.5 inch left indent");
    assert_eq!(paras[3].props.dxa_left_first, Some(360));
}

#[test]
fn doc_style_names_come_from_the_style_sheet() {
    let p = sample("formatted.doc");
    if !p.exists() {
        return;
    }
    let bytes = std::fs::read(&p).unwrap();
    let doc = office_doc::parse(&bytes).unwrap();
    // The style sheet is `cstd` `STD` records; Word omits the definition of a
    // built-in style that is still at its default, so only some have names.
    assert_eq!(doc.styles.len(), 15);
    // The style itself carries properties: WPS stores the paragraph
    // justification and the body size in `STD.grupe`, not in the `PAPX`.
    let body = doc.style_props(0).expect("style 0");
    assert_eq!(body.para.justify, Some(3), "正文 is justified");
    assert_eq!(body.char.size_half_points, Some(21), "正文 is 10.5 pt");
    assert_eq!(doc.style_name(0), Some("正文"));
    assert_eq!(doc.style_name(10), Some("默认段落字体"));
    assert_eq!(doc.style_name(11), Some("普通表格"));
    assert_eq!(doc.style_name(1), None);
    // Every paragraph of the sample uses the default paragraph style.
    let paras = doc.paragraphs_rich();
    assert_eq!(paras.len(), 5);
    for para in &paras {
        assert_eq!(para.style_name.as_deref(), Some("正文"), "{:?}", para.text);
    }
}

#[test]
fn microsoft_word_style_sheet_is_parsed_too() {
    let p = office16("PROTTPLN.DOC");
    if !p.exists() {
        return;
    }
    let doc = office_doc::parse(&std::fs::read(&p).unwrap()).unwrap();
    assert_eq!(doc.style_name(0), Some("Normal"));
    assert_eq!(doc.style_name(10), Some("Default Paragraph Font"));
    assert_eq!(doc.style_name(16), Some("Table Grid"));
    assert_eq!(doc.style_name(17), Some("Balloon Text"));
}

#[test]
fn doc_text_stream_splits_into_stories() {
    let p = sample("stories.doc");
    if !p.exists() {
        return;
    }
    let doc = office_doc::parse(&std::fs::read(&p).unwrap()).unwrap();
    assert_eq!(doc.fib.ccp_counts[0..3], [57, 19, 33]);
    let kinds: Vec<&str> = doc.stories().iter().map(|(k, _)| k.name()).collect();
    assert_eq!(kinds, ["main", "footnote", "header"]);
    assert!(doc.story(office_doc::Story::Header).contains("HEADER-MARK-11"));
}

#[test]
fn xls_formulas_and_number_formats() {
    let p = sample("formulas.xls");
    if !p.exists() {
        return;
    }
    let wb = office_xls::parse(&std::fs::read(&p).unwrap()).unwrap();
    let f = wb.sheet("F").expect("sheet F");
    let formula = |row: u16| {
        f.cells
            .iter()
            .find(|c| c.row == row && c.col == 1)
            .and_then(|c| match &c.value {
                office_xls::CellValue::Formula { text, .. } => Some(text.clone()),
                _ => None,
            })
            .unwrap_or_default()
    };
    assert_eq!(formula(0), "A1+A2");
    assert_eq!(formula(2), "SUM(A1:A2)");
    assert_eq!(formula(5), r#"IF(A1>5,"big","small")"#);
    assert_eq!(formula(11), "ROUND(A3,1)");
    assert_eq!(formula(12), "SUM(A1:A2)*2");
    assert_eq!(formula(14), "A1<>A2");
    // A cross-sheet reference in column C: =SUM(G!A1,3)
    let c2 = f.cells.iter().find(|c| c.row == 1 && c.col == 2).unwrap();
    assert_eq!(wb.formula_text(c2), Some("SUM(G!A1,3)"));

    let a3 = f.cells.iter().find(|c| c.row == 2 && c.col == 0).unwrap();
    assert_eq!(wb.cell_text(a3), "1234.57"); // number format "0.00"
    let a5 = f.cells.iter().find(|c| c.row == 4 && c.col == 0).unwrap();
    assert_eq!(wb.cell_text(a5), "2024-03-05"); // "yyyy-mm-dd"
}

#[test]
fn ppt_pictures_are_extracted_from_the_pictures_stream() {
    let p = sample("image.ppt");
    if !p.exists() {
        return;
    }
    let bytes = std::fs::read(&p).unwrap();
    let cfb = office_cfb::Cfb::parse(bytes).unwrap();
    let pics = office_ppt::pictures(&cfb.stream("Pictures").unwrap());
    assert_eq!(pics.len(), 1);
    assert_eq!(pics[0].index, 1);
    assert_eq!(pics[0].kind, office_ppt::PictureKind::Png);
    assert!(pics[0].data.starts_with(&[0x89, b'P', b'N', b'G']));

    let deck = office_ppt::parse_stream(
        &cfb.stream("PowerPoint Document").unwrap(),
        cfb.stream("Current User").ok().as_deref(),
    )
    .unwrap();
    assert_eq!(deck.slides[0].pictures, [1]); // the pib property points at picture 1
}

#[test]
fn encrypted_files_open_only_with_the_right_password() {
    let dir = concat!(env!("CARGO_MANIFEST_DIR"), "/../../samples");

    // .doc — legacy RC4 v1.1, header in the table stream.
    let doc_path = format!("{dir}/encrypted.doc");
    if let Ok(bytes) = std::fs::read(&doc_path) {
        assert!(office_doc::parse(&bytes).is_err(), "must not parse without a password");
        assert!(
            office_doc::parse_with_password(&bytes, "wrong").is_err(),
            "a wrong password must be rejected"
        );
        let doc = office_doc::parse_with_password(&bytes, "secret123").expect("decrypt .doc");
        assert!(doc.plain_text().contains("ENCRYPTED-DOC-777"));
        assert_eq!(doc.encryption.as_ref().map(|s| s.name()), Some("RC4 (legacy, v1.1)"));
    }

    // .xls — FILEPASS record, record headers stay in the clear.
    let xls_path = format!("{dir}/encrypted.xls");
    if let Ok(bytes) = std::fs::read(&xls_path) {
        assert!(office_xls::parse(&bytes).is_err(), "must not parse without a password");
        let wb = office_xls::parse_with_password(&bytes, "secret123").expect("decrypt .xls");
        let sheet = &wb.sheets[0];
        assert_eq!(sheet.name, "Secret");
        let text: String = sheet.cells.iter().map(|c| wb.cell_text(c)).collect::<Vec<_>>().join(" ");
        assert!(text.contains("ENCRYPTED-XLS-888"), "text was {text}");
    }

    // .ppt — per-persist-object RC4 CryptoAPI, header in a CryptSession10Container.
    let ppt_path = format!("{dir}/encrypted.ppt");
    if let Ok(bytes) = std::fs::read(&ppt_path) {
        let cfb = office_cfb::Cfb::parse(bytes).expect("cfb");
        let stream = cfb.stream("PowerPoint Document").expect("stream");
        let user = cfb.stream("Current User").ok();
        assert!(office_ppt::parse_stream(&stream, user.as_deref()).is_err());
        let deck =
            office_ppt::parse_stream_with_password(&stream, user.as_deref(), "secret123")
                .expect("decrypt .ppt");
        assert!(deck.plain_text().contains("ENCRYPTED-PPT-999"));
        assert_eq!(deck.encryption.as_ref().map(|s| s.name()), Some("RC4 CryptoAPI (128-bit)"));
    }
}

#[test]
fn doc_pictures_are_extracted_from_the_data_stream() {
    let path = concat!(env!("CARGO_MANIFEST_DIR"), "/../../samples/picture.doc");
    let Ok(bytes) = std::fs::read(path) else { return };
    let doc = office_doc::parse(&bytes).expect("parse .doc");
    assert!(!doc.data.is_empty(), "the Data stream must not be empty");

    let pics = doc.pictures();
    assert_eq!(pics.len(), 1, "one inline picture");
    assert_eq!(pics[0].kind, office_blip::PictureKind::Png);
    // The stored bytes are the original file, unchanged.
    let png = std::fs::read(concat!(env!("CARGO_MANIFEST_DIR"), "/../../samples/pixel.png"))
        .expect("pixel.png");
    assert_eq!(pics[0].data, png);
}

#[test]
fn doc_tables_are_read_from_the_cell_marks() {
    let p = sample("table.doc");
    if !p.exists() {
        return;
    }
    let bytes = std::fs::read(&p).unwrap();
    let doc = office_doc::parse(&bytes).unwrap();
    let paras = doc.paragraphs_rich();
    // Word ends every cell with 0x07 and every row with one extra 0x07 whose
    // properties carry `sprmPFTtp`.
    assert_eq!(paras.len(), 10);
    assert_eq!(paras[0].text, "Before table.");
    assert_eq!(paras[0].props.in_table, None);
    let cells: Vec<&str> = paras
        .iter()
        .filter(|p| p.props.in_table == Some(true) && p.props.ttp != Some(true))
        .map(|p| p.text.as_str())
        .collect();
    assert_eq!(cells, ["H1", "H2", "H3", "a", "b", "c"]);
    let rows = paras.iter().filter(|p| p.props.ttp == Some(true)).count();
    assert_eq!(rows, 2, "two rows, one terminator each");
    // The row-end paragraph carries `TDefTable`: three ~2841-twip columns.
    let terminator = paras.iter().find(|p| p.props.ttp == Some(true)).unwrap();
    assert_eq!(terminator.props.col_widths, vec![-108, 2732, 5573, 8414]);
    assert_eq!(paras[9].text, "After table.");
    assert_eq!(paras[9].props.in_table, None);
}

#[test]
fn ppt_speaker_notes_come_from_the_notes_persist_objects() {
    let p = sample("notes.ppt");
    if !p.exists() {
        return;
    }
    let bytes = std::fs::read(&p).unwrap();
    let cfb = office_cfb::Cfb::parse(bytes).unwrap();
    let stream = cfb.stream("PowerPoint Document").unwrap();
    let user = cfb.stream("Current User").ok();
    let deck = office_ppt::parse_stream(&stream, user.as_deref()).unwrap();
    assert_eq!(deck.slides.len(), 2);
    assert_eq!(deck.slides[0].texts[0].text, "NOTES-SLIDE-1");
    // Notes are a separate `Notes` persist object listed by the
    // SlideListWithText of instance 2, in slide order.
    assert_eq!(deck.slides[0].notes_plain_text(), "SPEAKER-NOTE-1 hidden text");
    assert_eq!(deck.slides[1].notes_plain_text(), "SPEAKER-NOTE-2 hidden text");
}

#[test]
fn xls_sst_continue_records_survive() {
    let p = sample("longstr.xls");
    if !p.exists() {
        return;
    }
    let bytes = std::fs::read(&p).unwrap();
    let wb = office_xls::parse(&bytes).unwrap();
    let sheet = wb.sheet("Long").unwrap();
    let cell = sheet.cell(0, 0).unwrap();
    let text = wb.cell_text(cell);
    // A 9000-character string cannot fit in one record: BIFF8 continues it in
    // CONTINUE records, each of which repeats the high-byte flag.
    assert_eq!(text.len(), 9000);
    assert!(text.starts_with("ABCDEFGHIJ"));
    assert_eq!(wb.cell_text(sheet.cell(1, 0).unwrap()), "tail-after-continue");
}

#[test]
fn doc_multi_page_text_is_complete() {
    let p = sample("long.doc");
    if !p.exists() {
        return;
    }
    let bytes = std::fs::read(&p).unwrap();
    let doc = office_doc::parse(&bytes).unwrap();
    let paras = doc.paragraphs();
    assert_eq!(paras.len(), 60);
    assert!(paras[0].starts_with("Paragraph 1 of the long document."));
    assert!(paras[59].starts_with("Paragraph 60 of the long document."));
    assert_eq!(doc.text.chars().count(), doc.fib.ccp_text as usize);
}

/// `.ppt` character formatting comes from `StyleTextPropAtom`.
///
/// The sample is written by WPS: the title is bold 32 pt red, one four-character
/// run in the body is bold, and the second slide's title is italic.
#[test]
fn ppt_text_formatting_comes_from_the_style_atom() {
    let p = sample("pptfmt.ppt");
    if !p.exists() {
        return;
    }
    let bytes = std::fs::read(&p).unwrap();
    let cfb = office_cfb::Cfb::parse(bytes).unwrap();
    let doc = cfb.stream("PowerPoint Document").unwrap();
    let user = cfb.stream("Current User").ok();
    let deck = office_ppt::parse_stream(&doc, user.as_deref()).unwrap();
    assert_eq!(deck.slides.len(), 2);

    let title = &deck.slides[0].texts[0];
    assert_eq!(title.kind, office_ppt::TextKind::Title);
    assert_eq!(title.text, "FMT-TITLE");
    assert_eq!(title.format.bold, Some(true));
    assert_eq!(title.format.size, Some(32));
    assert_eq!(title.format.colour, Some((0xFF, 0x00, 0x00)));
    assert_eq!(title.format.align, Some(1), "centred title");

    // The body is split at the bold run.
    let body: Vec<&office_ppt::TextRun> = deck.slides[0]
        .texts
        .iter()
        .filter(|r| r.kind == office_ppt::TextKind::Body)
        .collect();
    let bold = body
        .iter()
        .find(|r| r.format.bold == Some(true))
        .expect("one bold run in the body");
    assert_eq!(bold.text, "BOLD");

    // Formatting splits the body into three runs, but they are one paragraph.
    assert_eq!(deck.slides[0].paragraphs()[1], "plain then BOLD run");

    let second = &deck.slides[1].texts[0];
    assert_eq!(second.text, "FMT-2");
    assert_eq!(second.format.italic, Some(true));
    assert_eq!(second.format.size, Some(60));
}

/// A deck written by real Microsoft PowerPoint, not WPS.
///
/// `samples/你好.ppt` has a 60 pt red centred title and a 24 pt subtitle; it is
/// the file that confirmed the `StyleTextPropAtom` layout in the first place.
#[test]
fn microsoft_powerpoint_text_formatting() {
    let p = sample("你好.ppt");
    if !p.exists() {
        return;
    }
    let bytes = std::fs::read(&p).unwrap();
    let cfb = office_cfb::Cfb::parse(bytes).unwrap();
    let doc = cfb.stream("PowerPoint Document").unwrap();
    let user = cfb.stream("Current User").ok();
    let deck = office_ppt::parse_stream(&doc, user.as_deref()).unwrap();
    assert_eq!(deck.slides.len(), 1);
    let title = &deck.slides[0].texts[0];
    assert_eq!(title.text, "你好");
    assert_eq!(title.format.size, Some(60));
    assert_eq!(title.format.colour, Some((0xFF, 0x00, 0x00)));
    assert_eq!(title.format.align, Some(1));
    let sub = &deck.slides[0].texts[1];
    assert_eq!(sub.text, "这是副标题");
    assert_eq!(sub.format.size, Some(24));
}

/// The master's `TextMasterStyleAtom`s give the placeholder defaults: a
/// Microsoft deck says 44 pt for the title and a 28/24/20/18 ladder for body
/// text, which is what PowerPoint shows before any run states its own size.
#[test]
fn ppt_master_styles_are_parsed() {
    let p = sample("你好.ppt");
    if !p.exists() {
        return;
    }
    let bytes = std::fs::read(&p).unwrap();
    let cfb = office_cfb::Cfb::parse(bytes).unwrap();
    let doc = cfb.stream("PowerPoint Document").unwrap();
    let user = cfb.stream("Current User").ok();
    let deck = office_ppt::parse_stream(&doc, user.as_deref()).unwrap();

    let title = deck
        .master_styles
        .get(&office_ppt::TextKind::Title)
        .expect("title master style");
    assert_eq!(title[0].size, Some(44));
    assert_eq!(title[0].colour, Some((0, 0, 0)));

    let body = deck
        .master_styles
        .get(&office_ppt::TextKind::Body)
        .expect("body master style");
    let sizes: Vec<Option<u16>> = body.iter().take(4).map(|f| f.size).collect();
    assert_eq!(sizes, vec![Some(28), Some(24), Some(20), Some(18)]);

    // What the run states itself wins over the master.
    assert_eq!(deck.slides[0].texts[0].format.size, Some(60));
}

/// A run that states no size inherits it from the master's body level 0.
#[test]
fn ppt_runs_inherit_the_master_size() {
    let p = sample("hello.ppt");
    if !p.exists() {
        return;
    }
    let bytes = std::fs::read(&p).unwrap();
    let cfb = office_cfb::Cfb::parse(bytes).unwrap();
    let doc = cfb.stream("PowerPoint Document").unwrap();
    let user = cfb.stream("Current User").ok();
    let deck = office_ppt::parse_stream(&doc, user.as_deref()).unwrap();
    let master = deck
        .master_styles
        .get(&office_ppt::TextKind::Body)
        .expect("body master style");
    let default_size = master[0].size.expect("master body size");
    let second = deck.slides[1]
        .texts
        .iter()
        .find(|r| r.kind == office_ppt::TextKind::Body)
        .expect("body run on slide 2");
    assert_eq!(second.format.size, Some(default_size));
}

// ---------------------------------------------------------------------------
// Files written by *Microsoft* Office (see tools/ms_samples.ps1), which pin the
// parts of each format that WPS gets wrong or never exercises.
// ---------------------------------------------------------------------------

fn ms_sample(name: &str) -> PathBuf {
    sample("ms").join(name)
}

/// Excel 95 really does write BIFF5 when asked for `xlExcel9795` — this file is
/// the only real BIFF5 workbook the reader has ever seen.
#[test]
fn microsoft_excel95_is_biff5() {
    let p = ms_sample("excel95.xls");
    if !p.exists() {
        return;
    }
    let wb = office_xls::parse(&std::fs::read(&p).unwrap()).unwrap();
    assert_eq!(wb.biff_version, 0x0500);
    let sheet = wb.sheet("Old").expect("sheet Old");
    assert!(sheet.cells.iter().any(|c| matches!(
        &c.value,
        office_xls::CellValue::Text(t) if t == "BIFF5-TEXT"
    )));
    // BIFF5 keeps the column in one byte and the relative flags in the row word.
    assert!(wb
        .sheets
        .iter()
        .flat_map(|s| &s.cells)
        .any(|c| wb.formula_text(c) == Some("A2*2")));
}

/// A write-protected Excel 95 workbook is XOR-obfuscated, not RC4-encrypted.
#[test]
fn microsoft_xor_workbook_opens_with_the_password() {
    let p = ms_sample("xor.xls");
    if !p.exists() {
        return;
    }
    let bytes = std::fs::read(&p).unwrap();
    assert!(office_xls::parse(&bytes).is_err(), "must refuse without a password");
    assert!(office_xls::parse_with_password(&bytes, "wrong").is_err());
    let wb = office_xls::parse_with_password(&bytes, "secret123").expect("xor workbook");
    assert_eq!(wb.encryption.as_ref().map(|s| s.name()), Some("XOR obfuscation"));
    assert_eq!(wb.sheets[0].name, "Xor");
    assert!(wb
        .sheets[0]
        .cells
        .iter()
        .any(|c| matches!(&c.value, office_xls::CellValue::Text(t) if t == "XOR-SECRET")));
}

/// Word 97 with a password written by Microsoft uses RC4 CryptoAPI at 128 bits,
/// so the CryptoAPI key schedule is now checked against a Microsoft header.
#[test]
fn microsoft_cryptoapi_document_opens() {
    let p = ms_sample("locked.doc");
    if !p.exists() {
        return;
    }
    let bytes = std::fs::read(&p).unwrap();
    assert!(office_doc::parse(&bytes).is_err());
    let doc = office_doc::parse_with_password(&bytes, "secret123").expect("locked doc");
    assert_eq!(
        doc.encryption.as_ref().map(|s| s.name()),
        Some("RC4 CryptoAPI (128-bit)")
    );
    assert!(doc.plain_text().contains("MS-LOCKED-DOC-777"));
}

/// One pass over the plain Microsoft-written files: text, formulas, notes.
#[test]
fn microsoft_written_samples_parse() {
    let doc = ms_sample("hello.doc");
    if doc.exists() {
        let d = office_doc::parse(&std::fs::read(&doc).unwrap()).unwrap();
        assert_eq!(d.fib.table_stream, "1Table");
        assert!(d.plain_text().contains("MS-DOC-001"));
    }
    let xls = ms_sample("formulas.xls");
    if xls.exists() {
        let wb = office_xls::parse(&std::fs::read(&xls).unwrap()).unwrap();
        let sheet = wb.sheet("F").expect("sheet F");
        let texts: Vec<&str> = sheet
            .cells
            .iter()
            .filter_map(|c| wb.formula_text(c))
            .collect();
        for want in [
            "A1+A2",
            "A1*A2",
            "SUM(A1:A2)",
            "IF(A1>5,\"big\",\"small\")",
            "ROUND(A3,1)",
            "SUM(A1:A2)*2",
            "G!A1",
            "SUM(G!A1,3)",
        ] {
            assert!(texts.contains(&want), "missing {want} in {texts:?}");
        }
    }
    let ppt = ms_sample("formatted.ppt");
    if ppt.exists() {
        let bytes = std::fs::read(&ppt).unwrap();
        let cfb = office_cfb::Cfb::parse(bytes).unwrap();
        let doc = cfb.stream("PowerPoint Document").unwrap();
        let user = cfb.stream("Current User").ok();
        let deck = office_ppt::parse_stream(&doc, user.as_deref()).unwrap();
        assert_eq!(deck.slides.len(), 2);
        assert!(deck.slides.iter().any(|s| !s.notes.is_empty()));
        assert!(deck.slides.iter().any(|s| !s.pictures.is_empty()));
    }
}

/// Microsoft-written `.doc` puts the piece offset in `fc` **unhalved** for
/// compressed text (`0x40001000` where WPS writes `0x800`) and writes every
/// bold run as `ToggleOperand 0x81`. Both used to lose all formatting and the
/// whole table, so this test pins the rendering.
#[test]
fn microsoft_document_formatting_and_tables() {
    let p = ms_sample("formatted.doc");
    if !p.exists() {
        return;
    }
    let doc = office_doc::parse(&std::fs::read(&p).unwrap()).unwrap();
    let paras = doc.paragraphs_rich();

    let heading = &paras[0];
    assert!(heading.text.ends_with("Bold red heading"), "{}", heading.text);
    let span = heading.runs.last().expect("a run over the heading text");
    assert_eq!(span.props.bold, Some(true));
    assert_eq!(span.props.size_half_points, Some(36));
    assert_eq!(span.props.color, Some(0x0000_00FF));

    assert_eq!(paras[1].runs[0].props.italic, Some(true));
    assert_eq!(paras[1].runs[0].props.underline, Some(1));
    assert_eq!(paras[2].props.justify, Some(1));

    let cells: Vec<&str> = paras
        .iter()
        .filter(|p| p.props.in_table == Some(true) && p.props.ttp != Some(true))
        .map(|p| p.text.as_str())
        .collect();
    assert_eq!(cells, ["H1", "H2", "H3", "a", "b", "c"]);
    let rows: Vec<&office_doc::fmt::Paragraph> = paras
        .iter()
        .filter(|p| p.props.ttp == Some(true))
        .collect();
    assert_eq!(rows.len(), 2);
    assert_eq!(rows[0].props.col_widths.len(), 4);
}

/// A chart substream states its type by *which record is present*, not by a
/// field (MS-XLS 2.2.3.4): `Line` 0x1018, `Pie` 0x1019 with `pcDonut`,
/// `Bar` 0x1017 with `fTranspose`, `Scatter` 0x103A. Verified against Excel
/// writing five known `ChartType` values.
#[test]
fn microsoft_charts_are_typed_from_their_records() {
    let p = ms_sample("charts.xls");
    if !p.exists() {
        return;
    }
    let wb = office_xls::parse(&std::fs::read(&p).unwrap()).unwrap();
    let kinds: Vec<&str> = wb.charts.iter().map(|c| c.kind_name()).collect();
    // BOUNDSHEET order when Excel saved: Scatter, Bar, Pie, Column, Line.
    assert_eq!(kinds, ["scatter", "bar", "pie", "column", "line"]);
    for c in &wb.charts {
        assert_eq!(c.series, ["Sales"]);
        assert!(c.title.as_deref().is_some_and(|t| t.starts_with("T-")));
    }
}

/// A chart sheet carries the title, the axis titles and the series names in
/// `SeriesText` records; only the nesting depth tells them apart.
#[test]
fn microsoft_chart_sheet_title_and_axes() {
    let p = ms_sample("chart.xls");
    if !p.exists() {
        return;
    }
    let wb = office_xls::parse(&std::fs::read(&p).unwrap()).unwrap();
    let chart = wb.charts.first().expect("one chart");
    assert_eq!(chart.kind_name(), "line");
    assert_eq!(chart.title.as_deref(), Some("CHART-TITLE-XLS"));
    assert_eq!(chart.axis_titles, ["月份", "Sales"]);
    assert_eq!(chart.series, ["Sales"]);
    // This deck's chart is *embedded* in the worksheet (no BOUNDSHEET with
    // dt = 2), so it is found by scanning every `BOF dt = 0x0020` substream.
    assert!(wb.sheets.iter().all(|s| s.kind != office_xls::SheetKind::Chart));
    assert_eq!(wb.sheets.len(), 1);
}

/// `samples/obfuscated.doc` is `hello.doc` with Word's XOR obfuscation turned
/// on (`fEncrypted | fObfuscated`, verifier in `FibBase.lKey`). It is built by
/// `crates/officedump/examples/obf.rs`, so this test proves the *plumbing*
/// (flags → verifier → deobfuscation → parse) rather than the key schedule,
/// which is pinned separately against a Microsoft-written XOR workbook.
#[test]
fn xor_obfuscated_document_opens_with_the_password() {
    let bytes = std::fs::read(sample("obfuscated.doc")).unwrap();
    match office_doc::parse(&bytes) {
        Err(office_doc::Error::Encrypted) => {}
        other => panic!("expected Encrypted, got {other:?}"),
    }
    assert!(matches!(
        office_doc::parse_with_password(&bytes, "nope"),
        Err(office_doc::Error::WrongPassword)
    ));

    let d = office_doc::parse_with_password(&bytes, "secret123").unwrap();
    let scheme = d.encryption.as_ref().expect("scheme");
    assert_eq!(scheme.name(), "XOR obfuscation");
    assert_eq!(d.fib.flags & 0x8000, 0x8000, "fObfuscated stays set");
    assert!(d.text.contains("HELLO-DOC-001"));
    assert!(d.text.contains("中文测试"));
    assert!(d.text.contains("END-MARKER-42"));
}

#[test]
fn xor_obfuscation_round_trips_at_the_stream_level() {
    // Independent of any file: obfuscate then deobfuscate, with the FIB prefix
    // kept readable.
    let plain: Vec<u8> = (0..700u32).map(|i| (i * 11 + 5) as u8).collect();
    let arr = office_crypto::xor::array("secret123");
    let mut enc = office_crypto::xor::encrypt_run(&plain, &arr, 0, false);
    enc[..0x44].copy_from_slice(&plain[..0x44]);
    let back = office_doc::deobfuscate_stream(&enc, "secret123", 0x44);
    assert_eq!(back, plain);
}
