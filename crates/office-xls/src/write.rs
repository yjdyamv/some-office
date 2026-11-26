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
//!   EOF 0x000A
//! ```

use crate::{
    CachedValue, Cell, CellValue, REC_BOF, REC_BOOLERR, REC_BOUNDSHEET, REC_CODEPAGE, REC_DATEMODE,
    REC_EOF, REC_LABELSST, REC_NUMBER, REC_RK, REC_SST,
};

/// `FONT` — Excel refuses a workbook whose XF table points at a missing font.
const REC_FONT: u16 = 0x0031;
/// `XF` — a cell's `xf` indexes this table, so it cannot be left out either.
const REC_XF: u16 = 0x00E0;
/// `STYLE` — names the built-in "Normal" style that owns XF 0.
const REC_STYLE: u16 = 0x0293;
/// `DIMENSIONS` — the used range; Excel shows `#REF!` worksheets without it.
const REC_DIMENSIONS: u16 = 0x0200;
/// `WINDOW1` / `WINDOW2` — window state; Excel is tolerant but writes them.
const REC_WINDOW1: u16 = 0x003D;
const REC_WINDOW2: u16 = 0x023E;

/// Index of the cell XF every written cell points at.
const CELL_XF: u16 = 15;

/// BIFF8 version word, the one Excel 97-2003 writes.
const VERSION: u16 = 0x0600;
/// `BOF.rupBuild`, as written by Excel 97.
const BUILD: u16 = 0x0DBB;
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
fn write_xfs(out: &mut Vec<u8>) {
    for i in 0..16u16 {
        let mut p = Vec::new();
        p.extend_from_slice(&0u16.to_le_bytes()); // ifnt: our single font
        p.extend_from_slice(&0u16.to_le_bytes()); // ifmt: General
        // fLocked (bit 0) always; fStyle (bit 2) only for the style XF 0.
        let flags: u16 = if i == 0 { 0x0005 } else { 0x0001 };
        p.extend_from_slice(&flags.to_le_bytes());
        p.extend_from_slice(&[0u8; 14]); // alc/rot/txt/used + borders + fill
        rec(out, REC_XF, &p);
    }
}

/// Build a one-sheet BIFF8 workbook and wrap it in a CFB container.
///
/// Cells are written in the order given; `xf` is always 0, so no number format
/// is applied (a written date stays a bare serial number).
pub fn write_minimal(sheet_name: &str, cells: &[Cell]) -> Vec<u8> {
    // --- shared string table: unique texts in first-seen order --------------
    let mut texts: Vec<&str> = Vec::new();
    let mut total = 0usize;
    for c in cells {
        if let CellValue::Text(t) = &c.value {
            total += 1;
            if !texts.iter().any(|x| *x == t.as_str()) {
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
    let last_col = cells.iter().map(|c| c.col).max().unwrap_or(0) as u16;
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
                let _ = c.xf; // written cells all use CELL_XF
                p.extend_from_slice(&CELL_XF.to_le_bytes());
                p.extend_from_slice(&isst.to_le_bytes());
                rec(&mut sheet, REC_LABELSST, &p);
            }
            CellValue::Number(n) => {
                let mut p = Vec::new();
                p.extend_from_slice(&c.row.to_le_bytes());
                p.extend_from_slice(&c.col.to_le_bytes());
                let _ = c.xf; // written cells all use CELL_XF
                p.extend_from_slice(&CELL_XF.to_le_bytes());
                p.extend_from_slice(&n.to_le_bytes());
                rec(&mut sheet, REC_NUMBER, &p);
            }
            CellValue::Bool(b) => {
                let mut p = Vec::new();
                p.extend_from_slice(&c.row.to_le_bytes());
                p.extend_from_slice(&c.col.to_le_bytes());
                let _ = c.xf; // written cells all use CELL_XF
                p.extend_from_slice(&CELL_XF.to_le_bytes());
                p.push(u8::from(*b)); // bBoolErr comes first...
                p.push(0); // ...then fError = 0, i.e. a boolean not an error
                rec(&mut sheet, REC_BOOLERR, &p);
            }
            CellValue::Blank => {}
            CellValue::Error(code) => {
                let mut p = Vec::new();
                p.extend_from_slice(&c.row.to_le_bytes());
                p.extend_from_slice(&c.col.to_le_bytes());
                let _ = c.xf; // written cells all use CELL_XF
                p.extend_from_slice(&CELL_XF.to_le_bytes());
                p.push(*code); // bBoolErr: the error code
                p.push(1); // fError = 1
                rec(&mut sheet, REC_BOOLERR, &p);
            }
            // A written formula keeps only its cached value: rgce would have to
            // be re-emitted token by token, and the reader never invents one.
            CellValue::Formula { cached, .. } => match cached {
                CachedValue::Number(n) => rk_number(&mut sheet, c.row, c.col, *n),
                CachedValue::Text(t) => {
                    let isst = texts.iter().position(|x| *x == t.as_str()).unwrap_or(0) as u32;
                    let mut p = Vec::new();
                    p.extend_from_slice(&c.row.to_le_bytes());
                    p.extend_from_slice(&c.col.to_le_bytes());
                    let _ = c.xf; // written cells all use CELL_XF
                p.extend_from_slice(&CELL_XF.to_le_bytes());
                    p.extend_from_slice(&isst.to_le_bytes());
                    rec(&mut sheet, REC_LABELSST, &p);
                }
                CachedValue::Bool(b) => {
                    let mut p = Vec::new();
                    p.extend_from_slice(&c.row.to_le_bytes());
                    p.extend_from_slice(&c.col.to_le_bytes());
                    let _ = c.xf; // written cells all use CELL_XF
                p.extend_from_slice(&CELL_XF.to_le_bytes());
                    p.push(u8::from(*b));
                    p.push(0);
                    rec(&mut sheet, REC_BOOLERR, &p);
                }
                CachedValue::Error(code) => {
                    let mut p = Vec::new();
                    p.extend_from_slice(&c.row.to_le_bytes());
                    p.extend_from_slice(&c.col.to_le_bytes());
                    let _ = c.xf; // written cells all use CELL_XF
                p.extend_from_slice(&CELL_XF.to_le_bytes());
                    p.push(*code);
                    p.push(1);
                    rec(&mut sheet, REC_BOOLERR, &p);
                }
                CachedValue::Blank => {}
            },
        }
    }
    // WINDOW2: flags (0x06B6 = the defaults Excel writes) then geometry.
    let mut w2 = Vec::new();
    w2.extend_from_slice(&0x06B6u16.to_le_bytes());
    w2.extend_from_slice(&[0u8; 16]);
    rec(&mut sheet, REC_WINDOW2, &w2);
    rec(&mut sheet, REC_EOF, &[]);

    // --- globals, whose BOUNDSHEET must point at the sheet substream --------
    let mut head = Vec::new();
    bof(&mut head, 0x0005);
    rec(&mut head, REC_CODEPAGE, &0x04B0u16.to_le_bytes());
    rec(&mut head, REC_DATEMODE, &0u16.to_le_bytes());
    write_font(&mut head);
    write_xfs(&mut head);
    // BIFF8 STYLE: ixfe, flags (bit7 fBuiltin), builtinId (0 = Normal).
    rec(&mut head, REC_STYLE, &[0x00, 0x00, 0x80, 0x00]);
    // BIFF8 WINDOW1: 18 bytes of window geometry; zeros = one default window.
    rec(&mut head, REC_WINDOW1, &[0u8; 18]);
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
    rec(&mut head, REC_EOF, &[]);

    let sheet_at = head.len() as u32;
    head[bs_start + 4..bs_start + 8].copy_from_slice(&sheet_at.to_le_bytes());

    let mut stream = head;
    stream.extend_from_slice(&sheet);
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
}
