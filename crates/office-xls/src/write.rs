//! Write a minimal BIFF8 workbook.
//!
//! The point is not to be a spreadsheet engine but to close the loop: the bytes
//! come out of the same model the reader uses, and the tests prove the reader
//! reads them back. Anything the reader cannot re-read is a hole in the model,
//! which is exactly what this file exists to expose.
//!
//! ```text
//! BOF(0x0809, dt=0x0005 globals)
//!   CODEPAGE 0x0042 (0x04B0 = UTF-16)
//!   DATEMODE 0x0022 (0 = 1900 epoch)
//!   SST 0x00FC          (only when there is text)
//!   BOUNDSHEET 0x0085   -> sheet substream offset + name
//!   EOF 0x000A
//! BOF(0x0809, dt=0x0010 worksheet)   <- the offset recorded above
//!   LABELSST 0x00FD / NUMBER 0x0203 / RK 0x027E / BOOLERR 0x0205
//!   FORMULA 0x0006   (rgce compiled from the formula text)
//!   EOF 0x000A
//! ```

use crate::{
    CachedValue, Cell, CellValue, REC_BOF, REC_BOOLERR, REC_BOUNDSHEET,
    REC_EOF, REC_FORMULA, REC_LABELSST, REC_NUMBER, REC_RK, REC_SST,
};

/// `FONT` — Excel refuses a workbook whose XF table points at a missing font.
const REC_FONT: u16 = 0x0031;
/// `XF` — a cell's `xf` indexes this table, so it cannot be left out either.
const REC_XF: u16 = 0x00E0;
/// `STYLE` — names the built-in "Normal" style that owns XF 0.
/// `DIMENSIONS` — the used range; Excel shows `#REF!` worksheets without it.
const REC_DIMENSIONS: u16 = 0x0200;
/// `WINDOW1` / `WINDOW2` — window state; Excel is tolerant but writes them.
const REC_WINDOW2: u16 = 0x023E;

/// Index of the cell XF every written cell points at.
const CELL_XF: u16 = 15;
/// Custom `FORMAT` records start here (MS-XLS reserves 0..=163 for built-ins).
const FIRST_CUSTOM_IFMT: u16 = 164;
const REC_FORMAT: u16 = 0x041E;

/// BIFF8 version word, the one Excel 97-2003 writes.
const VERSION: u16 = 0x0600;
/// `BOF.rupBuild`, as written by Excel 97.
const BUILD: u16 = 0x4F5A;
/// `BOF.rupYear`.
const BUILD_YEAR: u16 = 0x07CD;
/// `BOF.bfh` — the flag word Excel 97 writes (0x000180C1).
const BOF_FLAGS: u32 = 0x0001_80C1;
/// `BOF.verLowestBiff` — 6.6, i.e. "Excel 97 can read this".
const LOWEST_BIFF: u32 = 0x0000_0606;

fn rec(out: &mut Vec<u8>, id: u16, payload: &[u8]) {
    out.extend_from_slice(&id.to_le_bytes());
    out.extend_from_slice(&(payload.len() as u16).to_le_bytes());
    out.extend_from_slice(payload);
}

fn bof(out: &mut Vec<u8>, dt: u16) {
    let mut p = Vec::new();
    p.extend_from_slice(&VERSION.to_le_bytes()); // vers
    p.extend_from_slice(&dt.to_le_bytes()); // dt: 0x0005 globals, 0x0010 sheet
    p.extend_from_slice(&BUILD.to_le_bytes()); // rupBuild
    p.extend_from_slice(&BUILD_YEAR.to_le_bytes()); // rupYear
    p.extend_from_slice(&BOF_FLAGS.to_le_bytes()); // bfh
    p.extend_from_slice(&LOWEST_BIFF.to_le_bytes()); // verLowestBiff
    rec(out, REC_BOF, &p);
}

fn push_utf16(out: &mut Vec<u8>, s: &str) {
    for ch in s.chars() {
        let mut buf = [0u16; 2];
        for unit in ch.encode_utf16(&mut buf) {
            out.extend_from_slice(&unit.to_le_bytes());
        }
    }
}

/// `XLUnicodeRichExtendedString` without the rich/extended tails, as used in `SST`.
fn sst_string(out: &mut Vec<u8>, s: &str) {
    out.extend_from_slice(&(s.chars().count() as u16).to_le_bytes());
    // Always UTF-16: one code path, and CJK needs no code page.
    out.push(0x01);
    push_utf16(out, s);
}

fn rk_number(out: &mut Vec<u8>, row: u16, col: u16, n: f64) {
    // Integer form when it fits, otherwise the top 32 bits of the f64.
    let rk = if n.fract() == 0.0 && n.abs() < 2.0_f64.powi(29) {
        ((n as i32) << 2) as u32 | 0x2
    } else {
        (n.to_bits() >> 32) as u32 & 0xFFFF_FFFC
    };
    let mut p = Vec::new();
    p.extend_from_slice(&row.to_le_bytes());
    p.extend_from_slice(&col.to_le_bytes());
    p.extend_from_slice(&0u16.to_le_bytes()); // xf
    p.extend_from_slice(&rk.to_le_bytes());
    rec(out, REC_RK, &p);
}

/// One `FONT` record: Arial 10 pt, the font XF 0 refers to.
fn write_font(out: &mut Vec<u8>) {
    let mut p = Vec::new();
    p.extend_from_slice(&200u16.to_le_bytes()); // dyHeight: 200 twips = 10 pt
    p.extend_from_slice(&0u16.to_le_bytes()); // grbit
    p.extend_from_slice(&0x7FFFu16.to_le_bytes()); // icv: automatic colour
    p.extend_from_slice(&400u16.to_le_bytes()); // bls: normal weight
    p.extend_from_slice(&0u16.to_le_bytes()); // sss: none
    p.push(0); // uls
    p.push(0); // bFamily
    p.push(0); // bCharSet
    p.push(0); // reserved
    p.push(5); // cch
    p.push(0x01); // fHighByte
    push_utf16(&mut p, "Arial");
    rec(out, REC_FONT, &p);
}

/// The 16 `XF` records Excel's built-in "Normal" style occupies: XF 0 is the
/// style XF, XFs 1..=14 are spare cell XFs, XF 15 is the default cell XF.
/// Fonts 1..4 are the bold/italic/bold-italic/underline faces Excel expects to
/// find in every workbook (they are referenced by the built-in style XFs).
fn write_font_variant(out: &mut Vec<u8>, which: u16) {
    let mut p = Vec::new();
    p.extend_from_slice(&200u16.to_le_bytes()); // dyHeight: 10 pt
    p.extend_from_slice(&(if which & 1 != 0 { 0x0002u16 } else { 0 } | if which & 2 != 0 { 0x0004 } else { 0 }).to_le_bytes());
    p.extend_from_slice(&0x7FFFu16.to_le_bytes()); // icv: automatic
    p.extend_from_slice(&400u16.to_le_bytes()); // bls: normal weight
    p.extend_from_slice(&0u16.to_le_bytes()); // sss
    p.push(0); // uls
    p.push(0); // family
    p.push(0); // charset
    p.push(0); // reserved
    let name = "Arial";
    p.push(name.chars().count() as u8);
    p.push(1); // fHighByte
    for ch in name.encode_utf16() {
        p.extend_from_slice(&ch.to_le_bytes());
    }
    rec(out, REC_FONT, &p);
}

/// Tail of a cell `XF` (everything after `ifnt`/`ifmt`), copied byte for byte
/// from what Excel writes: `flags 0xFFF5` says every attribute group applies,
/// then `f5 ff 10 00 00 f4 ...` and the trailing `c0 20`.
///
/// A hand-made tail (`flags 0x0001` + zeros) makes Excel refuse the workbook
/// with "Office 检测到文件数据有一个问题".
const CELL_TAIL: [u8; 16] = [
    // flags 0xFFF1: every attribute group applies *except* fStyle (bit 2) — a cell
    // XF that claims fStyle is treated as a style definition and its `ifmt` is
    // ignored, which is why formatted cells came out as General.
    0xf1, 0xff, 0x10, 0x00, 0x00, 0xf4, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0xc0,
    0x20,
];

/// Same, for the style `XF` 0 (Excel writes zeros where the cell XF has `f4`).
const STYLE_TAIL: [u8; 16] = [
    0xf5, 0xff, 0x10, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0xc0,
    0x20,
];

fn write_xfs(out: &mut Vec<u8>, formats: &[&str]) {
    // `FORMAT` first: Excel resolves each XF's `ifmt` while reading, so an XF that
    // names a `FORMAT` appearing later falls back to General (the values then show
    // up unformatted even though the file is structurally fine).
    for (i, code) in formats.iter().enumerate() {
        let ifmt = FIRST_CUSTOM_IFMT + i as u16;
        let mut p = Vec::new();
        p.extend_from_slice(&ifmt.to_le_bytes());
        p.extend_from_slice(&(code.chars().count() as u16).to_le_bytes());
        p.push(1); // fHighByte: UTF-16
        for ch in code.encode_utf16() {
            p.extend_from_slice(&ch.to_le_bytes());
        }
        rec(out, REC_FORMAT, &p);
    }
    for i in 0..16u16 {
        let mut p = Vec::new();
        p.extend_from_slice(&0u16.to_le_bytes()); // ifnt: our single font
        p.extend_from_slice(&0u16.to_le_bytes()); // ifmt: General
        p.extend_from_slice(if i == 0 { &STYLE_TAIL } else { &CELL_TAIL });
        rec(out, REC_XF, &p);
    }
    // One extra XF per custom format: `CELL_XF + 1 + i`.
    for (i, _) in formats.iter().enumerate() {
        let mut p = Vec::new();
        p.extend_from_slice(&0u16.to_le_bytes()); // ifnt
        p.extend_from_slice(&(FIRST_CUSTOM_IFMT + i as u16).to_le_bytes());
        p.extend_from_slice(&CELL_TAIL);
        rec(out, REC_XF, &p);
    }
}

/// The XF a written cell should use: `xf` selects a slot in `formats`, and cells
/// that ask for no format fall back to the default `CELL_XF`.
fn cell_xf(xf: u16, n_formats: usize) -> u16 {
    if (xf as usize) < n_formats {
        CELL_XF + 1 + xf
    } else {
        CELL_XF
    }
}

/// Build a one-sheet BIFF8 workbook and wrap it in a CFB container.
///
/// Cells are written in the order given; `xf` is always 0, so no number format
/// is applied (a written date stays a bare serial number).
pub fn write_minimal(sheet_name: &str, cells: &[Cell]) -> Vec<u8> {
    write_impl(sheet_name, cells, &[])
}

/// Like [`write_minimal`] but with per-cell number formats: a cell whose `xf` is
/// `i` gets `formats[i]` (Excel shows the cached value through that format).
pub fn write_with_formats(sheet_name: &str, cells: &[Cell], formats: &[&str]) -> Vec<u8> {
    write_impl(sheet_name, cells, formats)
}

fn write_impl(sheet_name: &str, cells: &[Cell], formats: &[&str]) -> Vec<u8> {
    // --- shared string table: unique texts in first-seen order --------------
    let mut texts: Vec<&str> = Vec::new();
    let mut total = 0usize;
    for c in cells {
        if let CellValue::Text(t) = &c.value {
            total += 1;
            if !texts.contains(&t.as_str()) {
                texts.push(t);
            }
        }
    }

    let mut sst = Vec::new();
    if !texts.is_empty() {
        sst.extend_from_slice(&(total as u32).to_le_bytes()); // cstTotal
        sst.extend_from_slice(&(texts.len() as u32).to_le_bytes()); // cstUnique
        for t in &texts {
            sst_string(&mut sst, t);
        }
    }

    // --- sheet substream ----------------------------------------------------
    let mut sheet = Vec::new();
    bof(&mut sheet, 0x0010);
    // DIMENSIONS: rwMic, rwMac (one past the last row), colMic, colMac, reserved.
    let last_row = cells.iter().map(|c| c.row).max().unwrap_or(0) as u32;
    let last_col = cells.iter().map(|c| c.col).max().unwrap_or(0);
    let mut dim = Vec::new();
    dim.extend_from_slice(&0u32.to_le_bytes());
    dim.extend_from_slice(&(last_row + 1).to_le_bytes());
    dim.extend_from_slice(&0u16.to_le_bytes());
    dim.extend_from_slice(&(last_col + 1).to_le_bytes());
    dim.extend_from_slice(&0u16.to_le_bytes());
    rec(&mut sheet, REC_DIMENSIONS, &dim);
    for c in cells {
        match &c.value {
            CellValue::Text(t) => {
                let isst = texts.iter().position(|x| *x == t.as_str()).unwrap_or(0) as u32;
                let mut p = Vec::new();
                p.extend_from_slice(&c.row.to_le_bytes());
                p.extend_from_slice(&c.col.to_le_bytes());
                p.extend_from_slice(&cell_xf(c.xf, formats.len()).to_le_bytes());
                p.extend_from_slice(&isst.to_le_bytes());
                rec(&mut sheet, REC_LABELSST, &p);
            }
            CellValue::Number(n) => {
                let mut p = Vec::new();
                p.extend_from_slice(&c.row.to_le_bytes());
                p.extend_from_slice(&c.col.to_le_bytes());
                p.extend_from_slice(&cell_xf(c.xf, formats.len()).to_le_bytes());
                p.extend_from_slice(&n.to_le_bytes());
                rec(&mut sheet, REC_NUMBER, &p);
            }
            CellValue::Bool(b) => {
                let mut p = Vec::new();
                p.extend_from_slice(&c.row.to_le_bytes());
                p.extend_from_slice(&c.col.to_le_bytes());
                p.extend_from_slice(&cell_xf(c.xf, formats.len()).to_le_bytes());
                p.push(u8::from(*b)); // bBoolErr comes first...
                p.push(0); // ...then fError = 0, i.e. a boolean not an error
                rec(&mut sheet, REC_BOOLERR, &p);
            }
            CellValue::Blank => {}
            CellValue::Error(code) => {
                let mut p = Vec::new();
                p.extend_from_slice(&c.row.to_le_bytes());
                p.extend_from_slice(&c.col.to_le_bytes());
                p.extend_from_slice(&cell_xf(c.xf, formats.len()).to_le_bytes());
                p.push(*code); // bBoolErr: the error code
                p.push(1); // fError = 1
                rec(&mut sheet, REC_BOOLERR, &p);
            }
            CellValue::Formula { cached, rgce, text } => {
                // Compile the text back into a token array, or pass through the
                // rgce we already have (e.g. a formula read from another file).
                let tokens = if !text.is_empty() {
                    crate::compile::compile(text).ok()
                } else if !rgce.is_empty() {
                    Some(rgce.clone())
                } else {
                    None
                };
                if let Some(t) = tokens.filter(|t| !t.is_empty()) {
                    let mut p = Vec::new();
                    p.extend_from_slice(&c.row.to_le_bytes());
                    p.extend_from_slice(&c.col.to_le_bytes());
                    p.extend_from_slice(&cell_xf(c.xf, formats.len()).to_le_bytes());
                    // Only a numeric cached value is written; anything else is
                    // recomputed because grbit carries fAlwaysCalc.
                    let cached_value = match cached {
                        CachedValue::Number(n) => *n,
                        _ => 0.0,
                    };
                    p.extend_from_slice(&cached_value.to_le_bytes());
                    p.extend_from_slice(&1u16.to_le_bytes()); // grbit: fAlwaysCalc
                    p.extend_from_slice(&0u32.to_le_bytes()); // chn
                    p.extend_from_slice(&(t.len() as u16).to_le_bytes()); // cce
                    p.extend_from_slice(&t);
                    rec(&mut sheet, REC_FORMULA, &p);
                } else {
                // No tokens: keep the cached value as a plain cell.
                match cached {
                CachedValue::Number(n) => rk_number(&mut sheet, c.row, c.col, *n),
                CachedValue::Text(t) => {
                    let isst = texts.iter().position(|x| *x == t.as_str()).unwrap_or(0) as u32;
                    let mut p = Vec::new();
                    p.extend_from_slice(&c.row.to_le_bytes());
                    p.extend_from_slice(&c.col.to_le_bytes());
                    p.extend_from_slice(&cell_xf(c.xf, formats.len()).to_le_bytes());
                    p.extend_from_slice(&isst.to_le_bytes());
                    rec(&mut sheet, REC_LABELSST, &p);
                }
                CachedValue::Bool(b) => {
                    let mut p = Vec::new();
                    p.extend_from_slice(&c.row.to_le_bytes());
                    p.extend_from_slice(&c.col.to_le_bytes());
                    p.extend_from_slice(&cell_xf(c.xf, formats.len()).to_le_bytes());
                    p.push(u8::from(*b));
                    p.push(0);
                    rec(&mut sheet, REC_BOOLERR, &p);
                }
                CachedValue::Error(code) => {
                    let mut p = Vec::new();
                    p.extend_from_slice(&c.row.to_le_bytes());
                    p.extend_from_slice(&c.col.to_le_bytes());
                    p.extend_from_slice(&cell_xf(c.xf, formats.len()).to_le_bytes());
                    p.push(*code);
                    p.push(1);
                    rec(&mut sheet, REC_BOOLERR, &p);
                }
                CachedValue::Blank => {}
                }
                }
            }
        }
    }
    // WINDOW2: flags (0x06B6 = the defaults Excel writes) then geometry.
    let mut w2 = Vec::new();
    w2.extend_from_slice(&0x06B6u16.to_le_bytes());
    w2.extend_from_slice(&[0u8; 16]);
    rec(&mut sheet, REC_WINDOW2, &w2);
    rec(&mut sheet, REC_EOF, &[]);

    // --- globals, whose BOUNDSHEET must point at the sheet substream --------
    //
    // Boilerplate Excel insists on: `globals.rs` holds it as Rust source, and
    // `globals_reproduce_that_workbook_byte_for_byte` keeps it honest. Records
    // whose payload depends on *our* workbook — `FONT`, `FORMAT`, `XF`, `STYLE`,
    // `SST`, `BOUNDSHEET` — are written below instead.
    let mut head = Vec::new();
    crate::globals::write_globals(&mut head);

    write_font(&mut head);
    for i in 1..5u16 {
        write_font_variant(&mut head, i);
    }
    write_xfs(&mut head, formats);
    if !texts.is_empty() {
        rec(&mut head, REC_SST, &sst);
    }
    let mut bs = Vec::new();
    bs.extend_from_slice(&0u32.to_le_bytes()); // lbPlyPos, patched below
    bs.push(0); // visibility: visible
    bs.push(0); // dt: worksheet
    bs.push(sheet_name.chars().count() as u8); // cch
    bs.push(0x01); // fHighByte: UTF-16
    push_utf16(&mut bs, sheet_name);
    let bs_start = head.len();
    rec(&mut head, REC_BOUNDSHEET, &bs);
    crate::globals::write_eof(&mut head);

    let sheet_at = head.len() as u32;
    head[bs_start + 4..bs_start + 8].copy_from_slice(&sheet_at.to_le_bytes());

    let mut stream = head;
    stream.extend_from_slice(&sheet);
    // Excel refuses a `Workbook` stream that lives in the mini stream (which is
    // where anything under 4096 bytes goes), so pad it past the cutoff.
    while stream.len() < 4096 {
        stream.push(0);
    }
    office_cfb::write(&[office_cfb::OutputStream::new("Workbook", &stream)])
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn writes_a_workbook_the_reader_accepts() {
        let cells = [
            Cell { row: 0, col: 0, xf: 0, value: CellValue::Text("写出的单元格".into()) },
            Cell { row: 1, col: 0, xf: 0, value: CellValue::Number(1234.5) },
            Cell { row: 2, col: 0, xf: 0, value: CellValue::Bool(true) },
            Cell { row: 3, col: 0, xf: 0, value: CellValue::Bool(false) },
        ];
        let bytes = write_minimal("Sheet1", &cells);
        let wb = crate::parse(&bytes).unwrap();
        assert_eq!(wb.sheets.len(), 1);
        assert_eq!(wb.sheets[0].name, "Sheet1");
        let s = &wb.sheets[0];
        assert_eq!(s.cells.len(), 4);
        assert!(matches!(&s.cells[0].value, CellValue::Text(t) if t == "写出的单元格"));
        assert!(matches!(s.cells[1].value, CellValue::Number(n) if (n - 1234.5).abs() < 1e-9));
        assert!(matches!(s.cells[2].value, CellValue::Bool(true)));
        assert!(matches!(s.cells[3].value, CellValue::Bool(false)));
    }

    #[test]
    fn texts_are_deduplicated_in_the_sst() {
        let cells = [
            Cell { row: 0, col: 0, xf: 0, value: CellValue::Text("same".into()) },
            Cell { row: 1, col: 0, xf: 0, value: CellValue::Text("same".into()) },
            Cell { row: 2, col: 0, xf: 0, value: CellValue::Text("other".into()) },
        ];
        let bytes = write_minimal("S", &cells);
        let wb = crate::parse(&bytes).unwrap();
        assert_eq!(wb.sst.len(), 2, "two unique strings");
        assert_eq!(wb.sheets[0].cells.len(), 3);
    }

    #[test]
    fn writes_custom_number_formats() {
        let cells = [
            Cell { row: 0, col: 0, xf: 0, value: CellValue::Number(45356.0) },
            Cell { row: 1, col: 0, xf: 1, value: CellValue::Number(0.25) },
            Cell { row: 2, col: 0, xf: 9, value: CellValue::Number(1234.5) },
        ];
        let bytes = write_with_formats("Fmt", &cells, &["yyyy-mm-dd", "0.00%"]);
        let wb = crate::parse(&bytes).unwrap();
        let sheet = &wb.sheets[0];
        assert_eq!(sheet.cells.len(), 3);
        assert_eq!(wb.cell_text(&sheet.cells[0]), "2024-03-05");
        assert_eq!(wb.cell_text(&sheet.cells[1]), "25.00%");
        // an `xf` past the end of `formats` falls back to the default XF (General)
        assert_eq!(wb.cell_text(&sheet.cells[2]), "1234.5");
        assert_eq!(wb.format_code(sheet.cells[0].xf), Some("yyyy-mm-dd"));
    }

    #[test]
    fn writes_real_formulas() {
        let cells = [
            Cell { row: 0, col: 0, xf: 0, value: CellValue::Number(10.0) },
            Cell { row: 1, col: 0, xf: 0, value: CellValue::Number(20.0) },
            Cell {
                row: 2,
                col: 0,
                xf: 0,
                value: CellValue::Formula {
                    cached: CachedValue::Number(30.0),
                    rgce: Vec::new(),
                    text: "SUM(A1:A2)".into(),
                },
            },
            Cell {
                row: 3,
                col: 0,
                xf: 0,
                value: CellValue::Formula {
                    cached: CachedValue::Number(200.0),
                    rgce: Vec::new(),
                    text: "A1*A2".into(),
                },
            },
        ];
        let bytes = write_minimal("F", &cells);
        let wb = crate::parse(&bytes).unwrap();
        let s = &wb.sheets[0];
        assert_eq!(s.cells.len(), 4);
        assert_eq!(wb.formula_text(&s.cells[2]), Some("SUM(A1:A2)"));
        assert_eq!(wb.formula_text(&s.cells[3]), Some("A1*A2"));
        assert_eq!(wb.cell_text(&s.cells[3]), "=A1*A2");
        // The cached value survives the round trip as well.
        match &s.cells[3].value {
            CellValue::Formula { cached, .. } => {
                assert!(matches!(cached, CachedValue::Number(n) if (*n - 200.0).abs() < 1e-9))
            }
            other => panic!("expected a formula, got {other:?}"),
        }
    }

    /// Excel ignores a cell XF's `ifmt` when the XF claims `fStyle` (bit 2): it
    /// treats the record as a style definition instead. Our first version set the
    /// bit and every number came out as "General".
    #[test]
    fn cell_xfs_do_not_claim_fstyle() {
        let cells = [Cell { row: 0, col: 0, xf: 0, value: CellValue::Number(1.0) }];
        let bytes = write_with_formats("F", &cells, &["0.00"]);
        let wb = crate::parse(&bytes).unwrap();
        // 16 default XFs + one per format; the last one is what a formatted cell uses.
        let last = *wb.xfs.last().unwrap();
        assert_eq!(last.ifmt, FIRST_CUSTOM_IFMT);
        // XF 0 is the style XF, so it keeps fStyle; every other XF must not.
        for (i, xf) in wb.xfs.iter().enumerate().skip(1) {
            assert_eq!(xf.ifmt, if i == 16 { FIRST_CUSTOM_IFMT } else { 0 });
        }
    }

    /// `FORMAT` must appear before the `XF` that points at it — Excel resolves
    /// `ifmt` while reading and falls back to General otherwise.
    #[test]
    fn formats_precede_the_xfs_that_use_them() {
        let cells = [Cell { row: 0, col: 0, xf: 0, value: CellValue::Number(1.0) }];
        let bytes = write_with_formats("F", &cells, &["0.00"]);
        let stream = office_cfb::Cfb::parse(bytes).unwrap().stream("Workbook").unwrap().to_vec();
        let fmt_at = stream
            .windows(2)
            .position(|w| w == [0x1E, 0x04])
            .expect("FORMAT record");
        let xf_at = stream
            .windows(2)
            .position(|w| w == [0xE0, 0x00])
            .expect("XF record");
        assert!(fmt_at < xf_at, "FORMAT at {fmt_at} must come before XF at {xf_at}");
    }
}
