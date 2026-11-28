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
    AutoFilter, CachedValue, Cell, CellValue, ErrorStyle, Font, REC_AUTOFILTERINFO, REC_BOF,
    REC_BOOLERR, REC_BOUNDSHEET, REC_NAME,
    PrintSetup, REC_BOTTOMMARGIN, REC_DV, REC_DVAL, REC_EOF, REC_FOOTER, REC_HEADER, REC_LEFTMARGIN,
    REC_RIGHTMARGIN, REC_SETUP, REC_TOPMARGIN, ValidationKind, ValidationOperator,
    REC_FORMULA, REC_LABELSST, REC_NUMBER, REC_RK, REC_SST,
};

// Re-exported so callers of `XfSpec` can name the types it takes without
// reaching into the crate root.
pub use crate::{Align, Border, BorderStyle, BoxBorder, Fill, Pattern, VAlign};

/// `FONT` — Excel refuses a workbook whose XF table points at a missing font.
const REC_FONT: u16 = 0x0031;
/// `PALETTE` — only written when a font asks for a colour the default table
/// cannot name (`docs/xls.md` §18).
const REC_PALETTE: u16 = 0x0092;
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

/// `HLINK`: one hyperlink on a cell range (`docs/xls.md` §20.2).
const REC_HLINK: u16 = 0x01B8;
/// `CLSID_Hyperlink`, the first moniker in an `HLINK` payload.
const CLSID_HYPERLINK: [u8; 16] = [
    0xD0, 0xC9, 0xEA, 0x79, 0xF9, 0xBA, 0xCE, 0x11, 0x8C, 0x82, 0x00, 0xAA, 0x00, 0x4B, 0xA9, 0x0B,
];
/// `CLSID_StdURLMoniker`, the second one.
const CLSID_URL_MONIKER: [u8; 16] = [
    0xE0, 0xC9, 0xEA, 0x79, 0xF9, 0xBA, 0xCE, 0x11, 0x8C, 0x82, 0x00, 0xAA, 0x00, 0x4B, 0xA9, 0x0B,
];
/// `Dv`: one data validation rule (`docs/xls.md` section 22).
const REC_DV_WRITE: u16 = REC_DV;
/// `DVal`: the header of the `Dv` list, and the number of rules.
const REC_DVAL_WRITE: u16 = REC_DVAL;
/// The 26 bytes Excel writes after the address: a NUL, a GUID, and two
/// counters. Identical in every `HLINK` we have measured.
const HLINK_TAIL: [u8; 26] = [
    0x00, 0x00, 0x79, 0x58, 0x81, 0xF4, 0x3B, 0x1D, 0x7F, 0x48, 0xAF, 0x2C, 0x82, 0x5D, 0xC4, 0x85,
    0x27, 0x63, 0x00, 0x00, 0x00, 0x00, 0xA5, 0xAB, 0x00, 0x03,
];
/// `ROW.grbit` for a row whose height was set by the user.
const ROW_CUSTOM_HEIGHT: u16 = 0x0140;

/// Index of the cell XF every written cell points at.
const CELL_XF: u16 = 15;
/// Custom `FORMAT` records start here (MS-XLS reserves 0..=163 for built-ins).
const FIRST_CUSTOM_IFMT: u16 = 164;
/// The faces every workbook starts with: the default one plus the
/// bold/italic/bold-italic variants Excel's built-in styles expect.
const BASE_FONTS: u16 = 4;
/// Records per caller-supplied face. Excel renumbers the font table from index
/// 4 on — it inserts a copy of the default face at 4 — so an index that is
/// correct by the specification shows the *previous* face in Excel. Writing
/// every custom face twice and pointing the `XF` at the second copy lands on
/// the same face under both readings (`docs/xls.md` §18).
const FONT_PAIR: u16 = 2;
/// `FONT.icv` for "let Excel pick the text colour".
const ICV_AUTOMATIC: u16 = 0x7FFF;
/// `PALETTE` entries replace the default table from this index on.
const FIRST_PALETTE_INDEX: u16 = 8;
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

/// The `FONT` records every workbook starts with. They must all differ:
/// Excel de-duplicates identical `FONT` records and then every `ifnt` in the
/// `XF` table is off by one, which silently applies the wrong face
/// (`docs/xls.md` §18).
fn base_fonts() -> [Font; BASE_FONTS as usize] {
    let face = Font::default_face;
    [
        face(),
        Font { bold: true, ..face() },
        Font { italic: true, ..face() },
        Font { bold: true, italic: true, ..face() },
    ]
}

/// One `FONT` record for a caller-supplied face (`docs/xls.md` §18).
fn write_custom_font(out: &mut Vec<u8>, font: &Font, icv: u16) {
    let mut p = Vec::new();
    p.extend_from_slice(&((font.size * 20.0).round() as u16).to_le_bytes());
    let mut grbit = 0u16;
    if font.bold {
        grbit |= 0x0001;
    }
    if font.italic {
        grbit |= 0x0002;
    }
    if font.underline {
        grbit |= 0x0004;
    }
    if font.strikeout {
        grbit |= 0x0008;
    }
    p.extend_from_slice(&grbit.to_le_bytes());
    p.extend_from_slice(&icv.to_le_bytes()); // icv: palette index or automatic
    p.extend_from_slice(&(if font.bold { 700u16 } else { 400 }).to_le_bytes()); // bls
    p.extend_from_slice(&0u16.to_le_bytes()); // sss: no super/subscript
    p.push(0); // uls: the underline style, unused (grbit bit 2 is enough)
    p.push(0); // bFamily
    p.push(0); // bCharSet
    p.push(0); // reserved
    let name = if font.name.is_empty() { "Arial" } else { font.name.as_str() };
    p.push(name.chars().count() as u8);
    p.push(1); // fHighByte
    for ch in name.encode_utf16() {
        p.extend_from_slice(&ch.to_le_bytes());
    }
    rec(out, REC_FONT, &p);
}

/// `PALETTE`: our custom colours, from index 8 on. `ccv` may be shorter than the
/// 56 entries the format allows — the ones we do not name keep their default.
fn write_palette(out: &mut Vec<u8>, colours: &[(u8, u8, u8)]) {
    let mut p = Vec::new();
    p.extend_from_slice(&(colours.len() as u16).to_le_bytes());
    for &(r, g, b) in colours {
        p.push(r); // red first: the bytes of the little-endian u32 `0x00BBGGRR`
        p.push(g);
        p.push(b);
        p.push(0);
    }
    rec(out, REC_PALETTE, &p);
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

/// The 16 `XF` records Excel's built-in "Normal" style occupies: XF 0 is the
/// style XF, XFs 1..=14 are spare cell XFs, XF 15 is the default cell XF.
fn write_xfs(
    out: &mut Vec<u8>,
    formats: &[&str],
    combos: &[XfSpec],
    pairs: &[(u16, u16)],
    palette: &[(u8, u8, u8)],
) {
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
    // One extra XF per (number format, font) pair: `CELL_XF + 1 + i`.
    for (spec, &(ifmt, ifnt)) in combos.iter().zip(pairs) {
        let mut p = Vec::new();
        p.extend_from_slice(&ifnt.to_le_bytes());
        p.extend_from_slice(&ifmt.to_le_bytes());
        p.extend_from_slice(&xf_tail(spec, palette));
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
        xfs: &[],
        layout: layout.clone(),
        links: &[],
        validations: &[],
        filter: None,
        print: None,
    }])
}

/// One entry of a sheet's XF table: the number format, the font and the cell
/// appearance a cell's `xf` index selects (`docs/xls.md` §18, §19).
#[derive(Clone, Debug, Default, PartialEq)]
pub struct XfSpec<'a> {
    /// `None` keeps the General format.
    pub format: Option<&'a str>,
    pub font: Font,
    /// The four edges and two diagonals.
    pub border: BoxBorder,
    pub align: Align,
    pub valign: VAlign,
    /// Wrap text inside the cell.
    pub wrap: bool,
    /// Indent level, 0..=15.
    pub indent: u8,
    /// Background pattern and colours.
    pub fill: Fill,
}

impl<'a> XfSpec<'a> {
    /// A General-format entry with the default face and no borders.
    pub fn plain() -> Self {
        XfSpec { format: None, font: Font::default_face(), ..XfSpec::default() }
    }

    /// An entry whose number format is `code`.
    pub fn format(code: &'a str) -> Self {
        XfSpec { format: Some(code), ..XfSpec::plain() }
    }

    /// Use `font` for this entry.
    pub fn font(mut self, font: Font) -> Self {
        self.font = font;
        self
    }

    /// Draw `border` around (or through) the cell.
    pub fn border(mut self, border: BoxBorder) -> Self {
        self.border = border;
        self
    }

    /// Align the text horizontally.
    pub fn align(mut self, align: Align) -> Self {
        self.align = align;
        self
    }

    /// Align the text vertically.
    pub fn valign(mut self, valign: VAlign) -> Self {
        self.valign = valign;
        self
    }

    /// Wrap long text instead of letting it spill into the next cell.
    pub fn wrap(mut self, wrap: bool) -> Self {
        self.wrap = wrap;
        self
    }

    /// Indent the text by `level` character widths (0..=15).
    pub fn indent(mut self, level: u8) -> Self {
        self.indent = level.min(15);
        self
    }

    /// Paint the background with `fill`.
    pub fn fill(mut self, fill: Fill) -> Self {
        self.fill = fill;
        self
    }
}

/// The 16 bytes of a cell `XF` that follow `ifnt`/`ifmt` (`docs/xls.md` §19):
/// alignment, borders and the palette indices that colour them.
fn xf_tail(xf: &XfSpec, palette: &[(u8, u8, u8)]) -> [u8; 16] {
    let icv = |colour: Option<(u8, u8, u8)>| -> u16 {
        match colour {
            // 0x40 is Excel's "automatic" for borders (`FONT.icv` uses 0x7FFF).
            None => 0x40,
            Some(rgb) => {
                let at = palette.iter().position(|p| *p == rgb).unwrap_or(0);
                8 + at as u16
            }
        }
    };
    let mut tail = CELL_TAIL;
    tail[2] = xf.align.code() | (xf.wrap as u8) << 3 | xf.valign.code() << 4;
    tail[4] = xf.indent & 0x0F;
    let dg = u16::from(xf.border.left.style.code())
        | u16::from(xf.border.right.style.code()) << 4
        | u16::from(xf.border.top.style.code()) << 8
        | u16::from(xf.border.bottom.style.code()) << 12;
    tail[6..8].copy_from_slice(&dg.to_le_bytes());
    let mut edges = (icv(xf.border.left.colour) & 0x7F) | (icv(xf.border.right.colour) & 0x7F) << 7;
    if xf.border.diagonal_down {
        edges |= 1 << 14;
    }
    if xf.border.diagonal_up {
        edges |= 1 << 15;
    }
    tail[8..10].copy_from_slice(&edges.to_le_bytes());
    let fill_icv = |colour: Option<(u8, u8, u8)>| -> u16 {
        match colour {
            // 0x40 / 0x41 are the two "automatic" values for fills.
            None => 0x40,
            Some(rgb) => {
                let at = palette.iter().position(|p| *p == rgb).unwrap_or(0);
                8 + at as u16
            }
        }
    };
    tail[14..16].copy_from_slice(
        &((fill_icv(xf.fill.fore) & 0x7F) | (fill_icv(xf.fill.back) & 0x7F) << 7).to_le_bytes(),
    );
    let rest = u32::from(icv(xf.border.top.colour) & 0x7F)
        | u32::from(icv(xf.border.bottom.colour) & 0x7F) << 7
        | u32::from(icv(xf.border.diagonal.colour) & 0x7F) << 14
        | u32::from(xf.border.diagonal.style.code()) << 21
        | u32::from(xf.fill.pattern.code()) << 26;
    tail[10..14].copy_from_slice(&rest.to_le_bytes());
    tail
}

/// One worksheet of a workbook written by [`write_workbook`].
/// One hyperlink: a URL attached to a single cell.
#[derive(Clone, Copy, Debug)]
pub struct Link<'a> {
    /// 0-based row and column of the cell the link hangs on.
    pub row: u16,
    pub col: u16,
    /// The address, e.g. `https://example.com/` or `mailto:someone@example.com`.
    pub url: &'a str,
    /// The text shown in the cell. An empty string shows `url`.
    pub text: &'a str,
}

/// One bound of a validation rule.
#[derive(Clone, Copy, Debug)]
pub enum Rule<'a> {
    /// A fixed list of choices, written as a `ptgStr` with NUL separators.
    List(&'a [&'a str]),
    /// Any other formula: `"1"`, `"=D1:D3"`, `"$A$1>5"`.
    Formula(&'a str),
}

/// One data validation rule (`Dv`, MS-XLS 2.4.95, `docs/xls.md` section 22).
#[derive(Clone, Debug)]
pub struct ValidationSpec<'a> {
    /// The range the rule covers, 0-based.
    pub first: (u16, u16),
    pub last: (u16, u16),
    pub kind: ValidationKind,
    pub operator: ValidationOperator,
    pub error_style: ErrorStyle,
    pub allow_blank: bool,
    pub show_dropdown: bool,
    pub prompt_title: &'a str,
    pub prompt: &'a str,
    pub error_title: &'a str,
    pub error: &'a str,
    pub formula1: Option<Rule<'a>>,
    pub formula2: Option<Rule<'a>>,
}

impl<'a> ValidationSpec<'a> {
    /// A rule that only accepts one of `choices`, with the in-cell drop-down.
    pub fn list(first: (u16, u16), last: (u16, u16), choices: &'a [&'a str]) -> Self {
        ValidationSpec {
            first,
            last,
            kind: ValidationKind::List,
            show_dropdown: true,
            allow_blank: true,
            formula1: Some(Rule::List(choices)),
            ..ValidationSpec::any(first, last)
        }
    }

    /// A rule that only accepts one of the values in `range` (e.g. `"D1:D3"`).
    pub fn list_from_range(first: (u16, u16), last: (u16, u16), range: &'a str) -> Self {
        ValidationSpec {
            kind: ValidationKind::List,
            show_dropdown: true,
            allow_blank: true,
            formula1: Some(Rule::Formula(range)),
            ..ValidationSpec::any(first, last)
        }
    }

    /// A whole number between `min` and `max` inclusive.
    pub fn whole_between(
        first: (u16, u16),
        last: (u16, u16),
        min: &'a str,
        max: &'a str,
    ) -> Self {
        ValidationSpec {
            kind: ValidationKind::Whole,
            operator: ValidationOperator::Between,
            allow_blank: true,
            formula1: Some(Rule::Formula(min)),
            formula2: Some(Rule::Formula(max)),
            ..ValidationSpec::any(first, last)
        }
    }

    /// A rule that accepts anything (useful as the base of the helpers above).
    pub fn any(first: (u16, u16), last: (u16, u16)) -> Self {
        ValidationSpec {
            first,
            last,
            kind: ValidationKind::Any,
            operator: ValidationOperator::Between,
            error_style: ErrorStyle::Stop,
            allow_blank: true,
            show_dropdown: true,
            prompt_title: "",
            prompt: "",
            error_title: "",
            error: "",
            formula1: None,
            formula2: None,
        }
    }

    /// The input prompt Excel shows when the cell is selected.
    pub fn prompt(mut self, title: &'a str, message: &'a str) -> Self {
        self.prompt_title = title;
        self.prompt = message;
        self
    }

    /// The alert Excel shows when a value is rejected.
    pub fn error(mut self, title: &'a str, message: &'a str) -> Self {
        self.error_title = title;
        self.error = message;
        self
    }

    /// The alert icon.
    pub fn error_style(mut self, style: ErrorStyle) -> Self {
        self.error_style = style;
        self
    }
}

pub struct SheetData<'a> {
    pub name: &'a str,
    pub cells: &'a [Cell],
    /// Per-cell number formats: a cell's `xf` indexes this slice.
    pub formats: &'a [&'a str],
    /// A richer XF table; when non-empty it replaces `formats`.
    pub xfs: &'a [XfSpec<'a>],
    pub layout: Layout,
    /// Hyperlinks; written after the cell records (`docs/xls.md` §20).
    pub links: &'a [Link<'a>],
    /// Print setup: paper, orientation, margins, header and footer
    /// (`docs/xls.md` §24). `None` leaves the sheet with no `SETUP` record.
    pub print: Option<&'a PrintSetup>,
    /// Data validation rules (`docs/xls.md` §22).
    pub validations: &'a [ValidationSpec<'a>],
    /// The AutoFilter range, if the sheet should show filter arrows
    /// (`docs/xls.md` §23).
    pub filter: Option<AutoFilter>,
}

impl<'a> SheetData<'a> {
    pub fn new(name: &'a str, cells: &'a [Cell]) -> Self {
        SheetData {
            name,
            cells,
            formats: &[],
            xfs: &[],
            layout: Layout::default(),
            links: &[],
            validations: &[],
            filter: None,
            print: None,
        }
    }

    /// Put an AutoFilter over `first..=last`, so every column gets a dropdown.
    pub fn with_filter(mut self, filter: AutoFilter) -> Self {
        self.filter = Some(filter);
        self
    }

    /// Add data validation rules to this sheet.
    pub fn with_validations(mut self, validations: &'a [ValidationSpec<'a>]) -> Self {
        self.validations = validations;
        self
    }

    /// Attach a print setup to this sheet.
    pub fn with_print(mut self, print: &'a PrintSetup) -> Self {
        self.print = Some(print);
        self
    }

    /// Attach `links` to this sheet.
    pub fn with_links(mut self, links: &'a [Link<'a>]) -> Self {
        self.links = links;
        self
    }

    pub fn with_formats(mut self, formats: &'a [&'a str]) -> Self {
        self.formats = formats;
        self
    }

    /// Cells pick a number format *and* a font through `Cell::xf`.
    pub fn with_xfs(mut self, xfs: &'a [XfSpec<'a>]) -> Self {
        self.xfs = xfs;
        self
    }

    /// The XF table this sheet actually uses.
    fn table(&self) -> Vec<XfSpec<'a>> {
        if self.xfs.is_empty() {
            self.formats.iter().map(|f| XfSpec::format(f)).collect()
        } else {
            self.xfs.to_vec()
        }
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

    // --- XF table: one entry per unique (number format, font) pair ----------
    let mut combos: Vec<XfSpec<'a>> = Vec::new();
    let mut maps: Vec<Vec<u16>> = Vec::new();
    for sheet in sheets {
        let mut map = Vec::new();
        for xf in sheet.table() {
            let at = match combos.iter().position(|c| *c == xf) {
                Some(at) => at,
                None => {
                    combos.push(xf);
                    combos.len() - 1
                }
            };
            map.push(at as u16);
        }
        maps.push(map);
    }
    // Number formats and fonts are numbered separately; `pairs` ties them back
    // together as the `(ifmt, ifnt)` of one XF record.
    let mut formats: Vec<&'a str> = Vec::new();
    let mut fonts: Vec<Font> = Vec::new();
    let mut pairs: Vec<(u16, u16)> = Vec::new();
    for spec in &combos {
        let code = &spec.format;
        let font = &spec.font;
        let ifmt = match code {
            Some(code) => {
                let at = match formats.iter().position(|f| f == code) {
                    Some(at) => at,
                    None => {
                        formats.push(code);
                        formats.len() - 1
                    }
                };
                FIRST_CUSTOM_IFMT + at as u16
            }
            None => 0,
        };
        let ifnt = match base_fonts().iter().position(|f| f == font) {
            Some(at) => at as u16,
            None => {
                let at = match fonts.iter().position(|f| f == font) {
                    Some(at) => at,
                    None => {
                        fonts.push(font.clone());
                        fonts.len() - 1
                    }
                };
                BASE_FONTS + FONT_PAIR * at as u16 + 1
            }
        };
        pairs.push((ifmt, ifnt));
    }
    // Colours the default palette cannot name become `PALETTE` entries from
    // index 8 on, in first-seen order.
    let mut palette: Vec<(u8, u8, u8)> = Vec::new();
    let add_colour = |c: (u8, u8, u8), palette: &mut Vec<(u8, u8, u8)>| {
        if !palette.contains(&c) {
            palette.push(c);
        }
    };
    for font in &fonts {
        if let Some(c) = font.colour {
            add_colour(c, &mut palette);
        }
    }
    // Border colours share the same palette (`docs/xls.md` §19).
    for spec in &combos {
        for border in [
            spec.border.left,
            spec.border.right,
            spec.border.top,
            spec.border.bottom,
            spec.border.diagonal,
        ] {
            if let Some(c) = border.colour {
                add_colour(c, &mut palette);
            }
        }
        // Fill colours live in the same palette (`docs/xls.md` §20).
        for c in [spec.fill.fore, spec.fill.back].into_iter().flatten() {
            add_colour(c, &mut palette);
        }
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

    for font in base_fonts().iter().chain(fonts.iter().flat_map(|f| [f, f])) {
        let icv = match font.colour {
            Some(c) => FIRST_PALETTE_INDEX + palette.iter().position(|p| *p == c).unwrap() as u16,
            None => ICV_AUTOMATIC,
        };
        write_custom_font(&mut head, font, icv);
    }
    if !palette.is_empty() {
        write_palette(&mut head, &palette);
    }
    write_xfs(&mut head, &formats, &combos, &pairs, &palette);
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
    // One `_FilterDatabase` name per filtered sheet, pointing at its range.
    for (i, sheet) in sheets.iter().enumerate() {
        if let Some(filter) = sheet.filter {
            let mut p = Vec::new();
            // fBuiltin | fHidden, one-character name, an 11-byte `ptgArea3d`.
            p.extend_from_slice(&0x0021u16.to_le_bytes());
            p.push(0); // chKey
            p.push(1); // cch
            p.extend_from_slice(&11u16.to_le_bytes()); // cce
            p.extend_from_slice(&0u16.to_le_bytes()); // ixals
            p.extend_from_slice(&((i + 1) as u16).to_le_bytes()); // itab
            p.extend_from_slice(&[0, 0, 0, 0]); // cchCustMenu/Description/Helptopic/Statusbar
            p.push(0x00); // the name: not wide
            p.push(0x0D); // `_FilterDatabase`
            p.push(0x3B); // ptgArea3d
            p.extend_from_slice(&0u16.to_le_bytes()); // ixti: this workbook
            p.extend_from_slice(&filter.first.0.to_le_bytes());
            p.extend_from_slice(&filter.last.0.to_le_bytes());
            p.extend_from_slice(&filter.first.1.to_le_bytes());
            p.extend_from_slice(&filter.last.1.to_le_bytes());
            rec(&mut head, REC_NAME, &p);
        }
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
    // `AutoFilterInfo` sits with the dimensions, before the cells; Excel derives
    // the range from the `_FilterDatabase` name in the globals.
    if let Some(filter) = sheet.filter {
        let columns = filter.last.1.saturating_sub(filter.first.1) + 1;
        rec(&mut out, REC_AUTOFILTERINFO, &columns.to_le_bytes());
    }
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
    for link in sheet.links {
        write_hlink(&mut out, link);
    }
    if !sheet.validations.is_empty() {
        // The header comes first and counts the rules that follow.
        let mut header = vec![0u8; 8];
        header.extend_from_slice(&[0, 0, 1, 0, 0, 0]);
        header.extend_from_slice(&(sheet.validations.len() as u16).to_le_bytes());
        header.extend_from_slice(&[0, 0]);
        rec(&mut out, REC_DVAL_WRITE, &header);
        for rule in sheet.validations {
            write_validation(&mut out, rule);
        }
    }
    if let Some(print) = sheet.print {
        write_print(&mut out, print);
    }
    rec(&mut out, REC_EOF, &[]);
    out
}

/// `SETUP`, the four margin records and the header/footer.
///
/// Measured against Excel 2016's own `samples/ms/print.xls` and
/// `samples/ms/orient-{portrait,landscape}.xls`: portrait sets bit 1 of
/// `SETUP.grbit`, landscape clears it, and the margins are `f64` inches.
fn write_print(out: &mut Vec<u8>, print: &PrintSetup) {
    let mut p = Vec::with_capacity(34);
    p.extend_from_slice(&print.paper_size.to_le_bytes());
    p.extend_from_slice(&100u16.to_le_bytes()); // iScale
    p.extend_from_slice(&1u16.to_le_bytes()); // iPageStart
    p.extend_from_slice(&1u16.to_le_bytes()); // iFitWidth
    p.extend_from_slice(&1u16.to_le_bytes()); // iFitHeight
    p.extend_from_slice(&if print.landscape { 0u16 } else { 2u16 }.to_le_bytes());
    p.extend_from_slice(&600u16.to_le_bytes()); // iRes
    p.extend_from_slice(&600u16.to_le_bytes()); // iVRes
    p.extend_from_slice(&0.3f64.to_le_bytes()); // numHdr
    p.extend_from_slice(&0.3f64.to_le_bytes()); // numFtr
    p.extend_from_slice(&1u16.to_le_bytes()); // iCopies
    rec(out, REC_SETUP, &p);
    rec(out, REC_LEFTMARGIN, &print.left_margin.to_le_bytes());
    rec(out, REC_RIGHTMARGIN, &print.right_margin.to_le_bytes());
    rec(out, REC_TOPMARGIN, &print.top_margin.to_le_bytes());
    rec(out, REC_BOTTOMMARGIN, &print.bottom_margin.to_le_bytes());
    write_header_footer(out, REC_HEADER, &print.header);
    write_header_footer(out, REC_FOOTER, &print.footer);
}

/// `HEADER`/`FOOTER`: `cch` (2), flags (1), then cp1252 or UTF-16 characters.
fn write_header_footer(out: &mut Vec<u8>, kind: u16, text: &str) {
    if text.is_empty() {
        return;
    }
    let ascii = text.chars().all(|ch| (ch as u32) < 0x80);
    let mut p = Vec::new();
    p.extend_from_slice(&(text.chars().count() as u16).to_le_bytes());
    p.push(u8::from(!ascii));
    if ascii {
        p.extend(text.chars().map(|ch| ch as u8));
    } else {
        for ch in text.encode_utf16() {
            p.extend_from_slice(&ch.to_le_bytes());
        }
    }
    rec(out, kind, &p);
}

/// `HLINK`: the reference range, the display string, and the URL moniker.
/// Measured layout, `docs/xls.md` §20.2 — the second length field counts bytes
/// (the address plus the 26-byte tail), not characters.
fn write_hlink(out: &mut Vec<u8>, link: &Link<'_>) {
    let text = if link.text.is_empty() { link.url } else { link.text };
    let mut p = Vec::new();
    // ref8: rwFirst, rwLast, colFirst, colLast — zero-based, unlike almost
    // every other BIFF row/column field.
    for _ in 0..2 {
        p.extend_from_slice(&link.row.to_le_bytes());
    }
    for _ in 0..2 {
        p.extend_from_slice(&link.col.to_le_bytes());
    }
    p.extend_from_slice(&CLSID_HYPERLINK);
    p.extend_from_slice(&2u32.to_le_bytes());
    p.extend_from_slice(&0x17u32.to_le_bytes());
    p.extend_from_slice(&((text.chars().count() as u32 + 1).to_le_bytes()));
    for unit in text.encode_utf16() {
        p.extend_from_slice(&unit.to_le_bytes());
    }
    p.extend_from_slice(&0u16.to_le_bytes());
    p.extend_from_slice(&CLSID_URL_MONIKER);
    let bytes = (link.url.encode_utf16().count() * 2 + HLINK_TAIL.len()) as u32;
    p.extend_from_slice(&bytes.to_le_bytes());
    for unit in link.url.encode_utf16() {
        p.extend_from_slice(&unit.to_le_bytes());
    }
    p.extend_from_slice(&HLINK_TAIL);
    rec(out, REC_HLINK, &p);
}

/// `Dv`: the option word, four strings, two bounds and one range.
fn write_validation(out: &mut Vec<u8>, rule: &ValidationSpec<'_>) {
    let mut flags = rule.kind.code() | rule.error_style.code() << 4;
    if matches!(rule.formula1, Some(Rule::List(_))) {
        flags |= 1 << 7; // fStrLookup: the list is a literal string
    }
    if rule.allow_blank {
        flags |= 1 << 8;
    }
    if !rule.show_dropdown {
        flags |= 1 << 9;
    }
    flags |= 1 << 18 | 1 << 19; // show the prompt and the error box
    flags |= rule.operator.code() << 20;
    let mut p = flags.to_le_bytes().to_vec();
    for text in [rule.prompt_title, rule.error_title, rule.prompt, rule.error] {
        push_xl_string(&mut p, text);
    }
    push_dv_formula(&mut p, rule.formula1, rule.first);
    push_dv_formula(&mut p, rule.formula2, rule.first);
    // sqref: one range, `(rwFirst, rwLast, colFirst, colLast)`.
    p.extend_from_slice(&1u16.to_le_bytes());
    p.extend_from_slice(&rule.first.0.to_le_bytes());
    p.extend_from_slice(&rule.last.0.to_le_bytes());
    p.extend_from_slice(&rule.first.1.to_le_bytes());
    p.extend_from_slice(&rule.last.1.to_le_bytes());
    rec(out, REC_DV_WRITE, &p);
}

/// `XLUnicodeString`: `cch`, a flag byte, then the characters. Excel writes an
/// empty string as one NUL character, so we do too.
fn push_xl_string(out: &mut Vec<u8>, text: &str) {
    if text.is_empty() {
        out.extend_from_slice(&1u16.to_le_bytes());
        out.push(0);
        out.push(0);
        return;
    }
    let wide = !text.is_ascii();
    out.extend_from_slice(&(text.chars().count() as u16).to_le_bytes());
    out.push(u8::from(wide));
    if wide {
        for unit in text.encode_utf16() {
            out.extend_from_slice(&unit.to_le_bytes());
        }
    } else {
        out.extend_from_slice(text.as_bytes());
    }
}

/// `DVParsedFormula`: `cce`, two unused bytes (Excel leaves garbage there, we
/// write zeros), then the token stream.
fn push_dv_formula(out: &mut Vec<u8>, rule: Option<Rule<'_>>, anchor: (u16, u16)) {
    let rgce = match rule {
        Some(Rule::List(choices)) => {
            // A list is one string constant, choices NUL-separated and the
            // whole thing wrapped in the formula's own quotes.
            let joined = format!("\"{}\"", choices.join("\u{0}"));
            let wide = !joined.is_ascii();
            let mut rgce = vec![0x17u8];
            rgce.push(joined.chars().count() as u8);
            rgce.push(u8::from(wide));
            if wide {
                for unit in joined.encode_utf16() {
                    rgce.extend_from_slice(&unit.to_le_bytes());
                }
            } else {
                rgce.extend_from_slice(joined.as_bytes());
            }
            rgce
        }
        Some(Rule::Formula(text)) => relative_rgce(text, anchor),
        None => Vec::new(),
    };
    out.extend_from_slice(&(rgce.len() as u16).to_le_bytes());
    out.extend_from_slice(&0u16.to_le_bytes());
    out.extend_from_slice(&rgce);
}

/// A `Dv` bound is evaluated relative to the validated cell, so a plain
/// `ptgRef`/`ptgArea` (`D1:D3`) must be stored as the `N` variant with offsets
/// (`docs/xls.md` section 22.3).
fn relative_rgce(text: &str, anchor: (u16, u16)) -> Vec<u8> {
    let rgce = crate::compile::compile(text).unwrap_or_default();
    let row_off = |row: u16| -> u16 { (row as i32 - anchor.0 as i32) as i16 as u16 };
    let col_off = |col: u16| -> u16 { ((col & 0xFF) as i32 - anchor.1 as i32) as i8 as u16 };
    let mut out = rgce.clone();
    match rgce.first().copied() {
        // `ptgRef` (row, col) → `ptgRefN` (row offset, col offset | flags).
        Some(kind @ (0x24 | 0x44 | 0x64)) if rgce.len() == 5 => {
            let row = u16::from_le_bytes([rgce[1], rgce[2]]);
            let col = u16::from_le_bytes([rgce[3], rgce[4]]);
            out[0] = kind | 0x08;
            out[1..3].copy_from_slice(&row_off(row).to_le_bytes());
            out[3..5].copy_from_slice(&(col_off(col) | 0xC000).to_le_bytes());
        }
        // `ptgArea` (rwFirst, rwLast, colFirst, colLast) → `ptgAreaN`.
        Some(kind @ (0x25 | 0x45 | 0x65)) if rgce.len() == 9 => {
            let rw_first = u16::from_le_bytes([rgce[1], rgce[2]]);
            let rw_last = u16::from_le_bytes([rgce[3], rgce[4]]);
            let col_first = u16::from_le_bytes([rgce[5], rgce[6]]);
            let col_last = u16::from_le_bytes([rgce[7], rgce[8]]);
            out[0] = kind | 0x08;
            out[1..3].copy_from_slice(&row_off(rw_first).to_le_bytes());
            out[3..5].copy_from_slice(&row_off(rw_last).to_le_bytes());
            // The relative flags of the whole range live in `colFirst`.
            out[5..7].copy_from_slice(&(col_off(col_first) | 0xC000).to_le_bytes());
            out[7..9].copy_from_slice(&(col_off(col_last) | 0xC000).to_le_bytes());
        }
        _ => {}
    }
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

#[cfg(test)]
mod style_tests {
    use super::*;

    #[test]
    fn a_solid_fill_round_trips() {
        let xfs = [XfSpec::plain().fill(Fill::solid((0xFF, 0, 0))), XfSpec::plain()];
        let bytes = write_workbook(&[SheetData::new("S", &cells()).with_xfs(&xfs)]);
        let wb = crate::parse(&bytes).expect("parse back");
        let fill = wb.xfs[wb.sheets[0].cells[0].xf as usize].fill;
        assert_eq!(fill.pattern, Pattern::Solid);
        assert_eq!(fill.fore, Some((0xFF, 0, 0)));
        assert_eq!(fill.back, None);
        assert!(!wb.xfs[wb.sheets[0].cells[1].xf as usize].fill.is_set());
    }

    #[test]
    fn a_pattern_keeps_both_colours() {
        let fill = Fill {
            pattern: Pattern::DarkGrid,
            fore: Some((0xFF, 0xFF, 0)),
            back: Some((0, 0, 0xFF)),
        };
        let xfs = [XfSpec::plain().fill(fill)];
        let bytes = write_workbook(&[SheetData::new("S", &cells()[..1]).with_xfs(&xfs)]);
        let wb = crate::parse(&bytes).expect("parse back");
        assert_eq!(wb.xfs[wb.sheets[0].cells[0].xf as usize].fill, fill);
        assert_eq!(wb.palette.len(), 2, "{:?}", wb.palette);
    }

    #[test]
    fn every_pattern_keeps_its_code() {
        let all = [
            Pattern::Solid,
            Pattern::Gray50,
            Pattern::Gray75,
            Pattern::Gray25,
            Pattern::DarkHorizontal,
            Pattern::DarkVertical,
            Pattern::DarkDown,
            Pattern::DarkUp,
            Pattern::DarkGrid,
            Pattern::DarkTrellis,
            Pattern::LightHorizontal,
            Pattern::LightVertical,
            Pattern::LightDown,
            Pattern::LightUp,
            Pattern::LightGrid,
            Pattern::LightTrellis,
            Pattern::Gray125,
            Pattern::Gray625,
        ];
        let xfs: Vec<XfSpec> =
            all.iter().map(|p| XfSpec::plain().fill(Fill { pattern: *p, ..Fill::default() })).collect();
        let cells: Vec<Cell> = (0..all.len() as u16)
            .map(|i| Cell { row: 0, col: i, xf: i, value: CellValue::Text("x".into()) })
            .collect();
        let bytes = write_workbook(&[SheetData::new("S", &cells).with_xfs(&xfs)]);
        let wb = crate::parse(&bytes).expect("parse back");
        for (i, want) in all.iter().enumerate() {
            let got = wb.xfs[wb.sheets[0].cells[i].xf as usize].fill.pattern;
            assert_eq!(got, *want, "cell {i}");
        }
    }

    fn cells() -> Vec<Cell> {
        vec![
            Cell { row: 0, col: 0, xf: 0, value: CellValue::Text("header".into()) },
            Cell { row: 1, col: 0, xf: 1, value: CellValue::Text("body".into()) },
        ]
    }

    #[test]
    fn a_bold_font_round_trips() {
        let xfs = [
            XfSpec::plain().font(Font { bold: true, ..Font::default_face() }),
            XfSpec::plain(),
        ];
        let bytes = write_workbook(&[SheetData::new("S", &cells()).with_xfs(&xfs)]);
        let wb = crate::parse(&bytes).expect("parse back");
        let header = wb.font(wb.sheets[0].cells[0].xf).expect("font");
        assert!(header.bold, "{header:?}");
        assert_eq!(header.name, "Arial");
        assert_eq!(header.size, 10.0);
        let body = wb.font(wb.sheets[0].cells[1].xf).expect("font");
        assert!(!body.bold);
    }

    #[test]
    fn size_italic_and_underline_round_trip() {
        let xfs = [
            XfSpec::plain().font(Font { size: 14.0, italic: true, underline: true, ..Font::default_face() }),
            XfSpec::plain().font(Font { size: 7.5, ..Font::default_face() }),
        ];
        let bytes = write_workbook(&[SheetData::new("S", &cells()).with_xfs(&xfs)]);
        let wb = crate::parse(&bytes).expect("parse back");
        let f = wb.font(wb.sheets[0].cells[0].xf).expect("font");
        assert_eq!(f.size, 14.0);
        assert!(f.italic && f.underline && !f.bold);
        assert_eq!(wb.font(wb.sheets[0].cells[1].xf).unwrap().size, 7.5);
    }

    #[test]
    fn a_custom_colour_becomes_a_palette_entry() {
        let red = Font { colour: Some((0xFF, 0x00, 0x00)), ..Font::default_face() };
        let blue = Font { colour: Some((0x00, 0x00, 0xFF)), ..Font::default_face() };
        let xfs = [XfSpec::plain().font(red), XfSpec::plain().font(blue)];
        let bytes = write_workbook(&[SheetData::new("S", &cells()).with_xfs(&xfs)]);
        let wb = crate::parse(&bytes).expect("parse back");
        assert_eq!(wb.palette.len(), 2, "{:?}", wb.palette);
        assert_eq!(wb.font(wb.sheets[0].cells[0].xf).unwrap().colour, Some((0xFF, 0, 0)));
        assert_eq!(wb.font(wb.sheets[0].cells[1].xf).unwrap().colour, Some((0, 0, 0xFF)));
    }

    #[test]
    fn the_palette_stores_red_first() {
        // Excel writes a colour entry as red, green, blue, 0x00 — the bytes of
        // the little-endian u32 `0x00BBGGRR`. Getting this backwards used to
        // pass every round-trip test (both sides were flipped) while Excel
        // painted the "red" cell blue, so pin the bytes themselves.
        let red = Font { colour: Some((0xC0, 0x00, 0x00)), ..Font::default_face() };
        let xfs = [XfSpec::plain().font(red)];
        let bytes = write_workbook(&[SheetData::new("S", &cells()).with_xfs(&xfs)]);
        let wanted = [0x92, 0x00, 0x06, 0x00, 0x01, 0x00, 0xC0, 0x00, 0x00, 0x00];
        assert!(
            bytes.windows(wanted.len()).any(|w| w == wanted),
            "PALETTE should hold c0 00 00 00 for #C00000"
        );
    }

    #[test]
    fn one_font_record_serves_every_sheet() {
        let bold = XfSpec::plain().font(Font { bold: true, ..Font::default_face() });
        let a = [Cell { row: 0, col: 0, xf: 0, value: CellValue::Text("a".into()) }];
        let b = [Cell { row: 0, col: 0, xf: 0, value: CellValue::Text("b".into()) }];
        let xfs = [bold];
        let bytes = write_workbook(&[
            SheetData::new("A", &a).with_xfs(&xfs),
            SheetData::new("B", &b).with_xfs(&xfs),
        ]);
        let wb = crate::parse(&bytes).expect("parse back");
        // A bold Arial 10 face is one of the base faces, so the workbook still
        // carries four `FONT` records and both sheets share one `ifnt`
        // (`docs/xls.md` §18).
        assert_eq!(wb.fonts.len(), 4, "{:?}", wb.fonts);
        let first = wb.font(wb.sheets[0].cells[0].xf).unwrap();
        let second = wb.font(wb.sheets[1].cells[0].xf).unwrap();
        assert!(first.bold, "{first:?}");
        assert_eq!(first, second);
    }

    #[test]
    fn print_setup_round_trips() {
        let cells = [Cell { row: 0, col: 0, xf: 0, value: CellValue::Text("print".into()) }];
        let mut print = PrintSetup::a4();
        print.landscape = true;
        print.left_margin = 0.5;
        print.right_margin = 0.75;
        print.top_margin = 0.25;
        print.bottom_margin = 1.5;
        print.header = "&CPRINT-HEADER".into();
        print.footer = "&CPRINT-FOOTER".into();
        let bytes = write_workbook(&[SheetData::new("Print", &cells).with_print(&print)]);
        let wb = crate::parse(&bytes).expect("parse back");
        let found = wb.sheets[0].print.as_ref().expect("a print setup");
        assert_eq!(found.paper_size, 9);
        assert!(found.landscape, "landscape clears the bit Word sets for portrait");
        assert_eq!(found.left_margin, 0.5);
        assert_eq!(found.right_margin, 0.75);
        assert_eq!(found.top_margin, 0.25);
        assert_eq!(found.bottom_margin, 1.5);
        assert_eq!(found.header, "&CPRINT-HEADER");
        assert_eq!(found.footer, "&CPRINT-FOOTER");
        // A sheet without a setup keeps none.
        let plain = write_minimal("Plain", &cells);
        assert!(crate::parse(&plain).unwrap().sheets[0].print.is_none());
    }

    #[test]
    fn hyperlinks_round_trip() {
        let cells = [Cell { row: 1, col: 1, xf: 0, value: CellValue::Text("click me".into()) }];
        let links = [
            Link { row: 1, col: 1, url: "https://example.com/page", text: "click me" },
            Link { row: 4, col: 1, url: "mailto:someone@example.com", text: "" },
        ];
        let bytes = write_workbook(&[SheetData::new("Links", &cells).with_links(&links)]);
        let wb = crate::parse(&bytes).expect("parse back");
        let found = &wb.sheets[0].hyperlinks;
        assert_eq!(found.len(), 2, "{found:?}");
        assert_eq!(found[0].first, (1, 1));
        assert_eq!(found[0].last, (1, 1));
        assert_eq!(found[0].url, "https://example.com/page");
        assert_eq!(found[0].text, "click me");
        assert_eq!(found[1].first, (4, 1));
        assert_eq!(found[1].url, "mailto:someone@example.com");
        assert_eq!(found[1].text, "mailto:someone@example.com", "empty text falls back to the url");
    }

    #[test]
    fn autofilter_round_trips() {
        let cells = [
            Cell { row: 0, col: 0, xf: 0, value: CellValue::Text("name".into()) },
            Cell { row: 1, col: 0, xf: 0, value: CellValue::Text("alpha".into()) },
        ];
        let filter = AutoFilter { first: (0, 0), last: (1, 2) };
        let bytes = write_workbook(&[SheetData::new("F", &cells).with_filter(filter)]);
        let wb = crate::parse(&bytes).expect("parse back");
        assert_eq!(wb.sheets[0].filter, Some(filter));
    }

    #[test]
    fn validations_round_trip() {
        let cells = [
            Cell { row: 0, col: 0, xf: 0, value: CellValue::Text("red".into()) },
            Cell { row: 2, col: 0, xf: 0, value: CellValue::Number(4.0) },
        ];
        let choices = ["red", "green", "blue"];
        let rules = [
            ValidationSpec::list((0, 0), (0, 0), &choices)
                .prompt("Colour", "pick one of three")
                .error("Not a colour", "red, green or blue"),
            ValidationSpec::whole_between((2, 0), (2, 0), "1", "10"),
        ];
        let bytes = write_workbook(&[SheetData::new("V", &cells).with_validations(&rules)]);
        let wb = crate::parse(&bytes).expect("parse back");
        let found = &wb.sheets[0].validations;
        assert_eq!(found.len(), 2, "{found:?}");
        assert_eq!(found[0].kind, ValidationKind::List);
        assert_eq!(found[0].first, (0, 0));
        assert_eq!(found[0].formula1, "red, green, blue");
        assert!(found[0].show_dropdown);
        assert!(found[0].allow_blank);
        assert_eq!(found[0].prompt_title, "Colour");
        assert_eq!(found[0].prompt, "pick one of three");
        assert_eq!(found[0].error_title, "Not a colour");
        assert_eq!(found[0].error, "red, green or blue");
        assert_eq!(found[1].kind, ValidationKind::Whole);
        assert_eq!(found[1].operator, ValidationOperator::Between);
        assert_eq!(found[1].first, (2, 0));
        assert_eq!(found[1].formula1, "1");
        assert_eq!(found[1].formula2, "10");
    }

    #[test]
    fn borders_round_trip() {
        let red = Border::thin().colour((0xFF, 0x00, 0x00));
        let xfs = [XfSpec::plain().border(BoxBorder {
            left: red,
            top: Border::new(BorderStyle::Thick),
            bottom: Border::new(BorderStyle::Double),
            ..BoxBorder::default()
        })];
        let cells = [Cell { row: 0, col: 0, xf: 0, value: CellValue::Text("b".into()) }];
        let bytes = write_workbook(&[SheetData::new("S", &cells).with_xfs(&xfs)]);
        let wb = crate::parse(&bytes).expect("parse back");
        let b = wb.xfs[wb.sheets[0].cells[0].xf as usize].border;
        assert_eq!(b.left.style, BorderStyle::Thin);
        assert_eq!(b.left.colour, Some((0xFF, 0x00, 0x00)));
        assert_eq!(b.top.style, BorderStyle::Thick);
        assert_eq!(b.top.colour, None, "automatic border colour");
        assert_eq!(b.bottom.style, BorderStyle::Double);
        assert_eq!(b.right.style, BorderStyle::None);
        assert_eq!(wb.palette.len(), 1, "{:?}", wb.palette);
    }

    #[test]
    fn every_border_style_keeps_its_code() {
        let styles = [
            BorderStyle::Thin,
            BorderStyle::Medium,
            BorderStyle::Dashed,
            BorderStyle::Dotted,
            BorderStyle::Thick,
            BorderStyle::Double,
            BorderStyle::Hair,
        ];
        let xfs: Vec<XfSpec> = styles
            .iter()
            .map(|s| XfSpec::plain().border(BoxBorder::all(Border::new(*s))))
            .collect();
        let cells: Vec<Cell> = (0..styles.len() as u16)
            .map(|i| Cell { row: i, col: 0, xf: i, value: CellValue::Text("x".into()) })
            .collect();
        let bytes = write_workbook(&[SheetData::new("S", &cells).with_xfs(&xfs)]);
        let wb = crate::parse(&bytes).expect("parse back");
        for (i, style) in styles.iter().enumerate() {
            let b = wb.xfs[wb.sheets[0].cells[i].xf as usize].border;
            assert_eq!(b.left.style, *style, "left of row {i}");
            assert_eq!(b.bottom.style, *style, "bottom of row {i}");
        }
    }

    #[test]
    fn alignment_and_diagonal_round_trip() {
        let xfs = [XfSpec::plain()
            .align(Align::Centre)
            .valign(VAlign::Centre)
            .wrap(true)
            .indent(2)
            .border(BoxBorder {
                diagonal: Border::new(BorderStyle::Thin),
                diagonal_down: true,
                ..BoxBorder::default()
            })];
        let cells = [Cell { row: 0, col: 0, xf: 0, value: CellValue::Text("a".into()) }];
        let bytes = write_workbook(&[SheetData::new("S", &cells).with_xfs(&xfs)]);
        let wb = crate::parse(&bytes).expect("parse back");
        let xf = wb.xfs[wb.sheets[0].cells[0].xf as usize];
        assert_eq!(xf.align, Align::Centre);
        assert_eq!(xf.valign, VAlign::Centre);
        assert!(xf.wrap);
        assert_eq!(xf.indent, 2);
        assert_eq!(xf.border.diagonal.style, BorderStyle::Thin);
        assert!(xf.border.diagonal_down);
        assert!(!xf.border.diagonal_up);
    }

    #[test]
    fn custom_faces_occupy_two_records() {
        let xfs = [XfSpec::plain().font(Font { size: 14.0, ..Font::default_face() })];
        let cells = [Cell { row: 0, col: 0, xf: 0, value: CellValue::Text("big".into()) }];
        let bytes = write_workbook(&[SheetData::new("S", &cells).with_xfs(&xfs)]);
        let wb = crate::parse(&bytes).expect("parse back");
        // Four base faces, then the custom face twice (`FONT_PAIR`, §18).
        assert_eq!(wb.fonts.len(), 4 + FONT_PAIR as usize, "{:?}", wb.fonts);
        assert_eq!(wb.fonts[4], wb.fonts[5]);
        assert_eq!(wb.fonts[4].size, 14.0);
        let cell = &wb.sheets[0].cells[0];
        assert_eq!(wb.font(cell.xf).unwrap().size, 14.0);
        assert_eq!(wb.xfs[cell.xf as usize].font, 5);
    }

    #[test]
    fn format_and_font_combine_in_one_xf() {
        const CODES: &[&str] = &["0.00%"];
        let xfs = [XfSpec::format(CODES[0]).font(Font { bold: true, ..Font::default_face() })];
        let cells = [Cell { row: 0, col: 0, xf: 0, value: CellValue::Number(0.25) }];
        let bytes = write_workbook(&[SheetData::new("S", &cells).with_xfs(&xfs)]);
        let wb = crate::parse(&bytes).expect("parse back");
        let cell = &wb.sheets[0].cells[0];
        assert_eq!(wb.cell_text(cell), "25.00%");
        assert!(wb.font(cell.xf).unwrap().bold);
    }
}
