//! Randomised round-trip tests: writer -> reader -> same model.
//!
//! The hand-written tests cover the cases we thought of. These cover the ones we
//! did not: every value is derived from a small deterministic PRNG, so a failure
//! is reproducible from the seed printed in the assertion message.

use office_cfb::{Cfb, OutputStream};

/// xorshift64*: tiny, deterministic, and good enough to shuffle test data.
struct Rng(u64);

impl Rng {
    fn new(seed: u64) -> Rng {
        Rng(seed | 1)
    }

    fn next(&mut self) -> u64 {
        let mut x = self.0;
        x ^= x >> 12;
        x ^= x << 25;
        x ^= x >> 27;
        self.0 = x;
        x.wrapping_mul(0x2545_F491_4F6C_DD1D)
    }

    fn below(&mut self, n: u64) -> u64 {
        self.next() % n
    }

    fn text(&mut self, max_chars: usize) -> String {
        let alphabet: Vec<char> = "abcXYZ019 -_中文测试Ω".chars().collect();
        let n = 1 + self.below(max_chars as u64) as usize;
        (0..n).map(|_| alphabet[self.below(alphabet.len() as u64) as usize]).collect()
    }
}

#[test]
fn cfb_round_trips_random_stream_sets() {
    for seed in 1..=8u64 {
        let mut rng = Rng::new(seed);
        let count = 1 + rng.below(5) as usize;
        let mut payloads: Vec<(String, Vec<u8>)> = Vec::new();
        for i in 0..count {
            // Mix sizes around the 4096-byte mini-stream cutoff, both sides.
            let len = match rng.below(4) {
                0 => rng.below(64) as usize,
                1 => 4095 - rng.below(8) as usize,
                2 => 4096 + rng.below(8) as usize,
                _ => rng.below(20000) as usize,
            };
            let data: Vec<u8> = (0..len).map(|_| rng.next() as u8).collect();
            payloads.push((format!("stream{i}"), data));
        }
        let streams: Vec<OutputStream<'_>> = payloads
            .iter()
            .map(|(name, data)| OutputStream::new(name, data))
            .collect();
        let bytes = office_cfb::write(&streams);
        let cfb = Cfb::parse(bytes).unwrap_or_else(|e| panic!("seed {seed}: {e}"));
        for (name, data) in &payloads {
            let got = cfb.stream(name).unwrap_or_else(|e| panic!("seed {seed}: {name}: {e}"));
            assert_eq!(got, data.as_slice(), "seed {seed}: stream {name} came back changed");
        }
    }
}

#[test]
fn xls_round_trips_random_cells() {
    for seed in 1..=12u64 {
        let mut rng = Rng::new(seed);
        let mut cells = Vec::new();
        for row in 0..rng.below(12) as u16 {
            let col = rng.below(4) as u16;
            let value = match rng.below(4) {
                0 => office_xls::CellValue::Text(rng.text(12)),
                1 => office_xls::CellValue::Number((rng.next() % 100000) as f64 / 8.0),
                2 => office_xls::CellValue::Bool(rng.below(2) == 0),
                _ => office_xls::CellValue::Number(rng.below(1000) as f64),
            };
            cells.push(office_xls::Cell { row, col, xf: 15, value });
        }
        if cells.is_empty() {
            continue;
        }
        let bytes = office_xls::write_minimal("Random", &cells);
        let wb = office_xls::parse(&bytes).unwrap_or_else(|e| panic!("seed {seed}: {e}"));
        let sheet = &wb.sheets[0];
        assert_eq!(sheet.cells.len(), cells.len(), "seed {seed}: cell count");
        for (got, want) in sheet.cells.iter().zip(cells.iter()) {
            assert_eq!((got.row, got.col), (want.row, want.col), "seed {seed}: position");
            assert_eq!(got.value.to_string(), want.value.to_string(), "seed {seed}: value");
        }
    }
}

#[test]
fn xls_formulas_are_stable_under_compile_decompile() {
    // `compile` then `decompile` must return the input text: if the token widths
    // were wrong, the reader would resynchronise onto garbage and the text would
    // drift. Re-compiling the decompiled text must then produce the same bytes.
    let formulas = [
        "A1+A2",
        "A1*A2-A3/A4",
        "(A1+A2)*3",
        "-A1+2",
        "SUM(A1:A5)",
        "SUM(A1:A2,A3)",
        "ROUND(A1,2)*2",
        "IF(A1>5,\"yes\",\"no\")",
        "AVERAGE(B1:B4)+MAX(C1:C3)",
        "A1&\"x\"",
    ];
    let ctx = office_xls::FormulaCtx {
        row: 0,
        col: 0,
        sheet: "",
        names: &[],
        extern_sheets: &[],
        biff8: true,
    };
    for text in formulas {
        let tokens = office_xls::compile::compile(text).unwrap_or_else(|e| panic!("{text}: {e}"));
        let back = office_xls::formula::decompile(&tokens, &ctx);
        assert_eq!(back, text, "decompiled text drifted");
        let again = office_xls::compile::compile(&back).unwrap_or_else(|e| panic!("{back}: {e}"));
        assert_eq!(again, tokens, "recompiling {back} produced different tokens");
    }
}

#[test]
fn doc_round_trips_random_formatting() {
    for seed in 1..=10u64 {
        let mut rng = Rng::new(seed);
        let mut paragraphs = Vec::new();
        for _ in 0..1 + rng.below(6) {
            let mut props = office_doc::fmt::CharProps::default();
            if rng.below(2) == 0 {
                props.bold = Some(true);
            }
            if rng.below(3) == 0 {
                props.italic = Some(true);
            }
            if rng.below(3) == 0 {
                props.size_half_points = Some(16 + rng.below(16) as u16);
            }
            if rng.below(3) == 0 {
                props.color = Some((rng.below(256) << 16 | rng.below(256) << 8 | rng.below(256)) as u32);
            }
            let mut pprops = office_doc::fmt::ParaProps::default();
            if rng.below(2) == 0 {
                pprops.justify = Some(rng.below(4) as u8);
            }
            if rng.below(3) == 0 {
                pprops.dxa_left = Some(rng.below(2000) as i16);
            }
            let text = rng.text(20);
            paragraphs.push(office_doc::Paragraph::styled(
                vec![office_doc::Run::styled(text, props.clone())],
                pprops.clone(),
            ));
        }
        let bytes = office_doc::write_document(&paragraphs);
        let doc = office_doc::parse(&bytes).unwrap_or_else(|e| panic!("seed {seed}: {e}"));
        let rich = doc.paragraphs_rich();
        assert_eq!(rich.len(), paragraphs.len(), "seed {seed}: paragraph count");
        for (got, want) in rich.iter().zip(paragraphs.iter()) {
            let want_text: String = want.runs.iter().map(|r| r.text.as_str()).collect();
            assert_eq!(got.text, want_text, "seed {seed}: text");
            // The writer always states the alignment (Word needs it), so a
            // `None` in the model comes back as `Some(0)` = left.
            assert_eq!(
                got.props.justify.unwrap_or(0),
                want.props.justify.unwrap_or(0),
                "seed {seed}: justify"
            );
            assert_eq!(got.props.dxa_left, want.props.dxa_left, "seed {seed}: indent");
            let span = got.runs.first().unwrap_or_else(|| panic!("seed {seed}: no spans"));
            let want_props = &want.runs[0].props;
            assert_eq!(span.props.bold, want_props.bold, "seed {seed}: bold");
            assert_eq!(span.props.italic, want_props.italic, "seed {seed}: italic");
            assert_eq!(
                span.props.size_half_points, want_props.size_half_points,
                "seed {seed}: size"
            );
            assert_eq!(span.props.color, want_props.color, "seed {seed}: colour");
        }
    }
}

#[test]
fn doc_round_trips_random_tables() {
    for seed in 1..=6u64 {
        let mut rng = Rng::new(seed);
        let rows = 1 + rng.below(4) as usize;
        let cols = 1 + rng.below(4) as usize;
        let grid: Vec<Vec<String>> = (0..rows)
            .map(|_| (0..cols).map(|_| rng.text(6)).collect())
            .collect();
        let refs: Vec<Vec<&str>> = grid.iter().map(|r| r.iter().map(|s| s.as_str()).collect()).collect();
        let table = office_doc::Table::new(&refs.iter().map(|r| r.as_slice()).collect::<Vec<_>>());
        let bytes = office_doc::write_table_document("head", &table, "tail");
        let doc = office_doc::parse(&bytes).unwrap_or_else(|e| panic!("seed {seed}: {e}"));
        let paras = doc.paragraphs_rich();
        let cells: Vec<&str> = paras
            .iter()
            .filter(|p| p.props.in_table == Some(true) && p.props.ttp != Some(true))
            .map(|p| p.text.as_str())
            .collect();
        let want: Vec<&str> = grid.iter().flatten().map(|s| s.as_str()).collect();
        assert_eq!(cells, want, "seed {seed}: cell contents");
        let ends = paras.iter().filter(|p| p.props.ttp == Some(true)).count();
        assert_eq!(ends, rows, "seed {seed}: one row-end mark per row");
    }
}

#[test]
fn ppt_round_trips_random_decks() {
    use office_ppt::write::{write_presentation, SlideSpec};

    for seed in 1..=4u64 {
        let mut rng = Rng::new(seed);
        let count = 1 + rng.below(3) as usize;
        let specs: Vec<SlideSpec> = (0..count)
            .map(|i| {
                let mut spec = SlideSpec::new(format!("T{seed}-{i}-{}", rng.text(6)));
                let lines = rng.below(3) as usize;
                let body: Vec<String> = (0..lines).map(|_| rng.text(10)).collect();
                if !body.is_empty() {
                    let refs: Vec<&str> = body.iter().map(|s| s.as_str()).collect();
                    spec = spec.with_body(&refs);
                }
                spec
            })
            .collect();

        let bytes = write_presentation(&specs).unwrap_or_else(|e| panic!("seed {seed}: {e}"));
        let cfb = Cfb::parse(bytes).unwrap_or_else(|e| panic!("seed {seed}: {e}"));
        let stream = cfb.stream("PowerPoint Document").unwrap();
        let user = cfb.stream("Current User").unwrap_or_default();
        let deck = office_ppt::parse_stream(&stream, Some(&user))
            .unwrap_or_else(|e| panic!("seed {seed}: {e}"));
        assert_eq!(deck.slides.len(), count, "seed {seed}: slide count");
        for (slide, spec) in deck.slides.iter().zip(specs.iter()) {
            let texts: Vec<&str> = slide.texts.iter().map(|t| t.text.as_str()).collect();
            assert_eq!(texts.first().copied(), Some(spec.title.as_str()), "seed {seed}: title");
            // The body lines go into one text atom, joined by CR.
            if !spec.body.is_empty() {
                let want = spec.body.join("\r");
                assert_eq!(texts.get(1).copied(), Some(want.as_str()), "seed {seed}: body");
            }
        }
    }
}
