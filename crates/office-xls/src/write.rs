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
/// `DEFCOLWIDTH` / `COLINFO` — the default and per-column widths.
const REC_DEFCOLWIDTH: u16 = 0x0055;
const REC_COLINFO: u16 = 0x007D;
/// `ROW` — one record per row that has a non-default height.
const REC_ROW: u16 = 0x0208;
/// `PANE` — the frozen/split panes of a sheet.
const REC_PANE: u16 = 0x0041;
/// `MERGEDCELLS` — the list of merged ranges.
const REC_MERGEDCELLS: u16 = 0x00E5;

/// `WINDOW2.grbit` as Excel writes it for a plain sheet, and the two extra
/// bits it sets when panes are frozen (`docs/xls.md` §15.4).
const WINDOW2_FLAGS: u16 = 0x06B6;
const WINDOW2_FROZEN: u16 = 0x0108;
/// `COLINFO.grbit`: the width was set by the user rather than by Excel.
const COLINFO_USER_SET: u16 = 0x0002;
/// `ROW.grbit` for a row whose height was set by the user.
const ROW_CUSTOM_HEIGHT: u16 = 0x0140;

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

/// Build a one-sheet BIFF8 workbook and wrap it in a CFB container.
///
/// Cells are written in the order given; `xf` is always 0, so no number format
/// is applied (a written date stays a bare serial number).
pub fn write_minimal(sheet_name: &str, cells: &[Cell]) -> Vec<u8> {
    write_impl(sheet_name, cells, &[], &Layout::default())
}

/// Like [`write_minimal`] but with per-cell number formats: a cell whose `xf` is
/// `i` gets `formats[i]` (Excel shows the cached value through that format).
pub fn write_with_formats(sheet_name: &str, cells: &[Cell], formats: &[&str]) -> Vec<u8> {
    write_impl(sheet_name, cells, formats, &Layout::default())
}

/// Sheet furniture that is not a cell: merged ranges, column widths, row
/// heights and frozen panes (`docs/xls.md` §15.4).
///
/// Rows and columns are zero-based like [`Cell::row`] / [`Cell::col`]; a merged
/// range is inclusive on both ends, so `(0, 0, 0, 2)` merges `A1:C1`.
#[derive(Clone, Debug, Default)]
pub struct Layout {
    /// Merged ranges as `(first_row, last_row, first_col, last_col)`.
    pub merged: Vec<(u16, u16, u16, u16)>,
    /// `(col, width)` in characters, i.e. what Excel's `ColumnWidth` shows.
    pub columns: Vec<(u16, f64)>,
    /// `(row, height)` in points, i.e. what Excel's `RowHeight` shows.
    pub rows: Vec<(u16, f64)>,
    /// `(rows, cols)` to freeze at the top-left corner.
    pub frozen: Option<(u16, u16)>,
}

impl Layout {
    /// Merge an inclusive rectangle, e.g. `.merge(0, 0, 0, 2)` for `A1:C1`.
    pub fn merge(mut self, first_row: u16, last_row: u16, first_col: u16, last_col: u16) -> Self {
        self.merged.push((first_row, last_row, first_col, last_col));
        self
    }

    /// Set a column's width in characters.
    pub fn column(mut self, col: u16, width: f64) -> Self {
        self.columns.push((col, width));
        self
    }

    /// Set a row's height in points.
    pub fn row(mut self, row: u16, height: f64) -> Self {
        self.rows.push((row, height));
        self
    }

    /// Freeze `rows` rows and `cols` columns at the top-left corner.
    pub fn freeze(mut self, rows: u16, cols: u16) -> Self {
        self.frozen = Some((rows, cols));
        self
    }
}

/// Like [`write_with_formats`], plus merged ranges, column widths, row heights
/// and frozen panes.
pub fn write_with_layout(
    sheet_name: &str,
    cells: &[Cell],
    formats: &[&str],
    layout: &Layout,
) -> Vec<u8> {
    write_impl(sheet_name, cells, formats, layout)
}

/// `coldx` from a width in characters.
///
/// Excel's own files store `width * 256 + 164`, but that padding is a function of
/// the workbook's default font: our `FONT` record is Arial 10, and Excel then
/// reads back `coldx / 256 - 0.71875`, so the constant is 184. Measured
/// (`docs/xls.md` §15.4): 10 → 10.00, 20 → 20.00, 50 → 50.00; Excel quantises
/// widths to whole pixels, so an odd width can read back a few hundredths off.
fn coldx(width: f64) -> u16 {
    (width * 256.0 + 184.0).round().clamp(0.0, 65535.0) as u16
}

/// `miyRw` from a height in points: 20 twips per point.
fn miyrw(height: f64) -> u16 {
    (height * 20.0).round().clamp(0.0, 65535.0) as u16
}

fn write_impl(sheet_name: &str, cells: &[Cell], formats: &[&str], layout: &Layout) -> Vec<u8> {
    write_workbook(&[SheetData {
        name: sheet_name,
        cells,
        formats,
        layout: layout.clone(),
    }])
}

/// One worksheet of a workbook written by [`write_workbook`].
pub struct SheetData<'a> {
    pub name: &'a str,
    pub cells: &'a [Cell],
    /// Per-cell number formats: a cell's `xf` indexes this slice.
    pub formats: &'a [&'a str],
    pub layout: Layout,
}

impl<'a> SheetData<'a> {
    pub fn new(name: &'a str, cells: &'a [Cell]) -> Self {
        SheetData { name, cells, formats: &[], layout: Layout::default() }
    }

    pub fn with_formats(mut self, formats: &'a [&'a str]) -> Self {
        self.formats = formats;
        self
    }

    pub fn with_layout(mut self, layout: Layout) -> Self {
        self.layout = layout;
        self
    }
}

/// The XF a cell should use: `xf` indexes this sheet's `formats`, which map
/// into the workbook-wide format table.
fn xf_for(xf: u16, map: &[u16]) -> u16 {
    match map.get(xf as usize) {
        Some(at) => CELL_XF + 1 + at,
        None => CELL_XF,
    }
}

pub fn write_workbook<'a>(sheets: &[SheetData<'a>]) -> Vec<u8> {
    // --- shared string table: unique texts in first-seen order --------------
    // One SST serves every sheet, so its indexes are global.
    let mut texts: Vec<&str> = Vec::new();
    let mut total = 0usize;
    for sheet in sheets {
        for c in sheet.cells {
            if let CellValue::Text(t) = &c.value {
                total += 1;
                if !texts.contains(&t.as_str()) {
                    texts.push(t);
                }
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

    // --- number formats: one XF per unique format, shared by all sheets -----
    let mut all_formats: Vec<&'a str> = Vec::new();
    let mut maps: Vec<Vec<u16>> = Vec::new();
    for sheet in sheets {
        let mut map = Vec::new();
        for f in sheet.formats {
            let at = match all_formats.iter().position(|x| x == f) {
                Some(at) => at,
                None => {
                    all_formats.push(f);
                    all_formats.len() - 1
                }
            };
            map.push(at as u16);
        }
        maps.push(map);
    }

    // --- one substream per sheet --------------------------------------------
    let mut substreams = Vec::new();
    for (i, sheet) in sheets.iter().enumerate() {
        substreams.push(sheet_substream(sheet, &texts, &maps[i]));
    }

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
    write_xfs(&mut head, &all_formats);
    if !texts.is_empty() {
        rec(&mut head, REC_SST, &sst);
    }
    // One BOUNDSHEET per sheet; `lbPlyPos` is patched once the substreams are
    // appended, because the globals' own length depends on these records.
    let mut sheet_at = Vec::new();
    for sheet in sheets {
        let mut bs = Vec::new();
        bs.extend_from_slice(&0u32.to_le_bytes()); // lbPlyPos, patched below
        bs.push(0); // visibility: visible
        bs.push(0); // dt: worksheet
        bs.push(sheet.name.chars().count() as u8); // cch
        bs.push(0x01); // fHighByte: UTF-16
        push_utf16(&mut bs, sheet.name);
        let at = head.len();
        rec(&mut head, REC_BOUNDSHEET, &bs);
        sheet_at.push(at);
    }
    crate::globals::write_eof(&mut head);

    let mut stream = head;
    let mut offset = stream.len() as u32;
    for (i, sub) in substreams.iter().enumerate() {
        stream[sheet_at[i] + 4..sheet_at[i] + 8].copy_from_slice(&offset.to_le_bytes());
        stream.extend_from_slice(sub);
        offset += sub.len() as u32;
    }
    // Excel refuses a `Workbook` stream that lives in the mini stream (which is
    // where anything under 4096 bytes goes), so pad it past the cutoff.
    while stream.len() < 4096 {
        stream.push(0);
    }
    office_cfb::write(&[office_cfb::OutputStream::new("Workbook", &stream)])
}

/// The records of one sheet substream: `BOF` … `EOF`.
fn sheet_substream<'a>(sheet: &SheetData<'a>, texts: &[&'a str], map: &[u16]) -> Vec<u8> {
    // --- sheet substream ----------------------------------------------------
    let mut out = Vec::new();
    bof(&mut out, 0x0010);
    // Column widths: Excel writes the default first, then one COLINFO per
    // column whose width the user set.
    if !sheet.layout.columns.is_empty() {
        rec(&mut out, REC_DEFCOLWIDTH, &8u16.to_le_bytes());
        for &(col, width) in &sheet.layout.columns {
            let mut p = Vec::new();
            p.extend_from_slice(&col.to_le_bytes());
            p.extend_from_slice(&col.to_le_bytes()); // colLast: one column
            p.extend_from_slice(&coldx(width).to_le_bytes());
            p.extend_from_slice(&CELL_XF.to_le_bytes()); // ixfe
            p.extend_from_slice(&COLINFO_USER_SET.to_le_bytes());
            p.extend_from_slice(&0u16.to_le_bytes()); // reserved
            rec(&mut out, REC_COLINFO, &p);
        }
    }
    // DIMENSIONS: rwMic, rwMac (one past the last row), colMic, colMac, reserved.
    let last_row = sheet.cells.iter().map(|c| c.row).max().unwrap_or(0) as u32;
    let last_col = sheet.cells.iter().map(|c| c.col).max().unwrap_or(0);
    let mut dim = Vec::new();
    dim.extend_from_slice(&0u32.to_le_bytes());
    dim.extend_from_slice(&(last_row + 1).to_le_bytes());
    dim.extend_from_slice(&0u16.to_le_bytes());
    dim.extend_from_slice(&(last_col + 1).to_le_bytes());
    dim.extend_from_slice(&0u16.to_le_bytes());
    rec(&mut out, REC_DIMENSIONS, &dim);
    // Row heights, before the cells they apply to.
    let mut heights = sheet.layout.rows.clone();
    heights.sort_by_key(|&(row, _)| row);
    for (row, height) in heights {
        let mut p = Vec::new();
        p.extend_from_slice(&row.to_le_bytes());
        p.extend_from_slice(&0u16.to_le_bytes()); // colMic
        p.extend_from_slice(&(last_col + 1).to_le_bytes()); // colMac
        p.extend_from_slice(&miyrw(height).to_le_bytes()); // miyRw, twips
        p.extend_from_slice(&0u16.to_le_bytes()); // irwMac: unused in BIFF8
        p.extend_from_slice(&0u16.to_le_bytes()); // reserved
        p.extend_from_slice(&ROW_CUSTOM_HEIGHT.to_le_bytes());
        p.extend_from_slice(&CELL_XF.to_le_bytes()); // ixfe
        rec(&mut out, REC_ROW, &p);
    }
    for c in sheet.cells {
        match &c.value {
            CellValue::Text(t) => {
                let isst = texts.iter().position(|x| *x == t.as_str()).unwrap_or(0) as u32;
                let mut p = Vec::new();
                p.extend_from_slice(&c.row.to_le_bytes());
                p.extend_from_slice(&c.col.to_le_bytes());
                p.extend_from_slice(&xf_for(c.xf, map).to_le_bytes());
                p.extend_from_slice(&isst.to_le_bytes());
                rec(&mut out, REC_LABELSST, &p);
            }
            CellValue::Number(n) => {
                let mut p = Vec::new();
                p.extend_from_slice(&c.row.to_le_bytes());
                p.extend_from_slice(&c.col.to_le_bytes());
                p.extend_from_slice(&xf_for(c.xf, map).to_le_bytes());
                p.extend_from_slice(&n.to_le_bytes());
                rec(&mut out, REC_NUMBER, &p);
            }
            CellValue::Bool(b) => {
                let mut p = Vec::new();
                p.extend_from_slice(&c.row.to_le_bytes());
                p.extend_from_slice(&c.col.to_le_bytes());
                p.extend_from_slice(&xf_for(c.xf, map).to_le_bytes());
                p.push(u8::from(*b)); // bBoolErr comes first...
                p.push(0); // ...then fError = 0, i.e. a boolean not an error
                rec(&mut out, REC_BOOLERR, &p);
            }
            CellValue::Blank => {}
            CellValue::Error(code) => {
                let mut p = Vec::new();
                p.extend_from_slice(&c.row.to_le_bytes());
                p.extend_from_slice(&c.col.to_le_bytes());
                p.extend_from_slice(&xf_for(c.xf, map).to_le_bytes());
                p.push(*code); // bBoolErr: the error code
                p.push(1); // fError = 1
                rec(&mut out, REC_BOOLERR, &p);
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
                    p.extend_from_slice(&xf_for(c.xf, map).to_le_bytes());
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
                    rec(&mut out, REC_FORMULA, &p);
                } else {
                // No tokens: keep the cached value as a plain cell.
                match cached {
                CachedValue::Number(n) => rk_number(&mut out, c.row, c.col, *n),
                CachedValue::Text(t) => {
                    let isst = texts.iter().position(|x| *x == t.as_str()).unwrap_or(0) as u32;
                    let mut p = Vec::new();
                    p.extend_from_slice(&c.row.to_le_bytes());
                    p.extend_from_slice(&c.col.to_le_bytes());
                    p.extend_from_slice(&xf_for(c.xf, map).to_le_bytes());
                    p.extend_from_slice(&isst.to_le_bytes());
                    rec(&mut out, REC_LABELSST, &p);
                }
                CachedValue::Bool(b) => {
                    let mut p = Vec::new();
                    p.extend_from_slice(&c.row.to_le_bytes());
                    p.extend_from_slice(&c.col.to_le_bytes());
                    p.extend_from_slice(&xf_for(c.xf, map).to_le_bytes());
                    p.push(u8::from(*b));
                    p.push(0);
                    rec(&mut out, REC_BOOLERR, &p);
                }
                CachedValue::Error(code) => {
                    let mut p = Vec::new();
                    p.extend_from_slice(&c.row.to_le_bytes());
                    p.extend_from_slice(&c.col.to_le_bytes());
                    p.extend_from_slice(&xf_for(c.xf, map).to_le_bytes());
                    p.push(*code);
                    p.push(1);
                    rec(&mut out, REC_BOOLERR, &p);
                }
                CachedValue::Blank => {}
                }
                }
            }
        }
    }
    // WINDOW2: flags then geometry; frozen panes add two flag bits and the
    // trailing counter Excel writes with them (measured, `docs/xls.md` §15.4).
    let frozen = sheet.layout.frozen.filter(|&(rows, cols)| rows > 0 || cols > 0);
    let mut w2 = Vec::new();
    let flags = match frozen {
        Some(_) => WINDOW2_FLAGS | WINDOW2_FROZEN,
        None => WINDOW2_FLAGS,
    };
    w2.extend_from_slice(&flags.to_le_bytes());
    w2.extend_from_slice(&[0u8; 12]);
    w2.extend_from_slice(&(u32::from(frozen.is_some()) * 2).to_le_bytes());
    rec(&mut out, REC_WINDOW2, &w2);
    if let Some((rows, cols)) = frozen {
        let mut p = Vec::new();
        p.extend_from_slice(&cols.to_le_bytes()); // xSplit: frozen columns
        p.extend_from_slice(&rows.to_le_bytes()); // ySplit: frozen rows
        p.extend_from_slice(&rows.to_le_bytes()); // rwTop
        p.extend_from_slice(&cols.to_le_bytes()); // colLeft
        p.extend_from_slice(&0u16.to_le_bytes()); // pnnAct: bottom-right pane
        rec(&mut out, REC_PANE, &p);
    }
    if !sheet.layout.merged.is_empty() {
        let mut p = Vec::new();
        p.extend_from_slice(&(sheet.layout.merged.len() as u16).to_le_bytes());
        for &(first_row, last_row, first_col, last_col) in &sheet.layout.merged {
            p.extend_from_slice(&first_row.to_le_bytes());
            p.extend_from_slice(&last_row.to_le_bytes());
            p.extend_from_slice(&first_col.to_le_bytes());
            p.extend_from_slice(&last_col.to_le_bytes());
        }
        rec(&mut out, REC_MERGEDCELLS, &p);
    }
    rec(&mut out, REC_EOF, &[]);
    out
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

#[cfg(test)]
mod layout_tests {
    use super::*;

    fn cells() -> Vec<Cell> {
        vec![
            Cell { row: 0, col: 0, xf: 15, value: CellValue::Text("title".into()) },
            Cell { row: 2, col: 0, xf: 15, value: CellValue::Text("left".into()) },
            Cell { row: 2, col: 2, xf: 15, value: CellValue::Number(42.0) },
        ]
    }

    #[test]
    fn layout_round_trips_through_the_reader() {
        let layout = Layout::default()
            .merge(0, 0, 0, 2)
            .merge(1, 2, 1, 1)
            .column(0, 20.0)
            .column(2, 30.0)
            .row(0, 40.0)
            .row(2, 22.5)
            .freeze(1, 1);
        let bytes = write_with_layout("Layout", &cells(), &[], &layout);
        let wb = crate::parse(&bytes).expect("parse back");
        let f = &wb.sheets[0].furniture;
        assert_eq!(f.merged, vec![(0, 0, 0, 2), (1, 2, 1, 1)]);
        // coldx = width * 256 + 184, measured against Excel's read-back.
        assert_eq!(f.columns, vec![(0, 5304), (2, 7864)]);
        // miyRw = height * 20 twips.
        assert_eq!(f.rows, vec![(0, 800), (2, 450)]);
        assert_eq!(f.frozen, Some((1, 1)));
    }

    #[test]
    fn a_plain_workbook_has_no_furniture() {
        let bytes = write_minimal("Plain", &cells());
        let wb = crate::parse(&bytes).expect("parse back");
        assert_eq!(wb.sheets[0].furniture, crate::Furniture::default());
    }
}

#[cfg(test)]
mod sheets_tests {
    use super::*;

    fn text(row: u16, s: &str) -> Cell {
        Cell { row, col: 0, xf: 0, value: CellValue::Text(s.into()) }
    }

    #[test]
    fn writes_every_sheet() {
        let a = [text(0, "first"), text(1, "shared")];
        let b = [text(0, "shared"), text(1, "second")];
        let bytes = write_workbook(&[SheetData::new("One", &a), SheetData::new("Two", &b)]);
        let wb = crate::parse(&bytes).expect("parse back");
        assert_eq!(wb.sheets.len(), 2);
        assert_eq!(wb.sheets[0].name, "One");
        assert_eq!(wb.sheets[1].name, "Two");
        assert_eq!(wb.cell_text(&wb.sheets[0].cells[0]), "first");
        assert_eq!(wb.cell_text(&wb.sheets[1].cells[1]), "second");
        // Both sheets index one shared string table.
        assert_eq!(wb.sst.len(), 3, "{:?}", wb.sst);
    }

    #[test]
    fn each_sheet_keeps_its_own_number_formats() {
        const DATES: &[&str] = &["yyyy-mm-dd"];
        const PCT: &[&str] = &["0.00%"];
        let dates = [Cell { row: 0, col: 0, xf: 0, value: CellValue::Number(45356.0) }];
        let pct = [Cell { row: 0, col: 0, xf: 0, value: CellValue::Number(0.25) }];
        let bytes = write_workbook(&[
            SheetData::new("Dates", &dates).with_formats(DATES),
            SheetData::new("Pct", &pct).with_formats(PCT),
        ]);
        let wb = crate::parse(&bytes).expect("parse back");
        assert_eq!(wb.cell_text(&wb.sheets[0].cells[0]), "2024-03-05");
        assert_eq!(wb.cell_text(&wb.sheets[1].cells[0]), "25.00%");
        // `xf` indexes the sheet's own `formats`, not a workbook-wide table.
        assert_eq!(wb.format_code(wb.sheets[0].cells[0].xf), Some("yyyy-mm-dd"));
        assert_eq!(wb.format_code(wb.sheets[1].cells[0].xf), Some("0.00%"));
    }

    #[test]
    fn furniture_belongs_to_its_sheet() {
        let a = [text(0, "wide")];
        let b = [text(0, "narrow")];
        let bytes = write_workbook(&[
            SheetData::new("Wide", &a)
                .with_layout(Layout::default().column(0, 24.0).freeze(0, 1)),
            SheetData::new("Narrow", &b),
        ]);
        let wb = crate::parse(&bytes).expect("parse back");
        assert_eq!(wb.sheets[0].furniture.columns, vec![(0, 6328)]);
        assert_eq!(wb.sheets[0].furniture.frozen, Some((0, 1)));
        assert_eq!(wb.sheets[1].furniture, crate::Furniture::default());
    }

    #[test]
    fn a_single_sheet_workbook_matches_write_minimal() {
        let cells = [text(0, "same")];
        let one = write_workbook(&[SheetData::new("Sheet1", &cells)]);
        let direct = write_minimal("Sheet1", &cells);
        assert_eq!(one, direct, "the one-sheet path must stay byte-identical");
    }
}
