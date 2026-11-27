//! # office-xls — Excel 97-2003 binary format (BIFF8, `.xls`)
//!
//! The `Workbook` stream is a flat sequence of **BIFF records**:
//!
//! ```text
//! u16 record id | u16 payload length | payload
//! ```
//!
//! A workbook is split into substreams, each starting with a `BOF` record
//! whose `dt` field says what it is (0x0005 globals, 0x0010 worksheet,
//! 0x0020 chart, 0x0040 macro sheet) and ending with `EOF` (0x000A):
//!
//! ```text
//! BOF(globals) CODEPAGE DATEMODE BOUNDSHEET* SST EXTSST ... EOF
//! BOF(worksheet) DIMENSIONS ROW* cell* ... EOF
//! BOF(worksheet) ... EOF
//! ```
//!
//! Shared strings live once in the `SST` record and cells reference them by
//! index (`LABELSST`); numbers are either IEEE doubles (`NUMBER`) or a packed
//! `RK` value. `SST` strings can straddle `CONTINUE` records, where the
//! character encoding flag byte is repeated at each continuation boundary.
//!
//! Verified against Excel/WPS output; see `docs/xls.md`.

use office_cfb::{Cfb, Error as CfbError};
use std::collections::HashMap;

use office_enc::decode_cp1252;

pub mod chart;
pub mod compile;
pub mod globals;
pub mod write;
pub use write::{write_minimal, write_with_formats};
pub mod formula;
pub mod numfmt;

pub use formula::Ctx as FormulaCtx;

pub type Result<T> = std::result::Result<T, Error>;

#[derive(Debug)]
pub enum Error {
    Cfb(CfbError),
    NoWorkbookStream,
    Truncated(&'static str),
    BadRecord(&'static str),
    UnsupportedBiff(u16),
    /// `FILEPASS` was found but no password was supplied.
    Encrypted,
    /// The supplied password does not match the stored verifier.
    WrongPassword,
    Crypto(office_crypto::Error),
}

impl From<office_crypto::Error> for Error {
    fn from(e: office_crypto::Error) -> Self {
        Error::Crypto(e)
    }
}

impl std::fmt::Display for Error {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Error::Cfb(e) => write!(f, "cfb: {e}"),
            Error::NoWorkbookStream => write!(f, "no Workbook/Book stream"),
            Error::Truncated(w) => write!(f, "truncated {w}"),
            Error::BadRecord(w) => write!(f, "bad record: {w}"),
            Error::UnsupportedBiff(v) => write!(f, "unsupported BIFF version 0x{v:04X}"),
            Error::Encrypted => write!(f, "workbook is encrypted; a password is required"),
            Error::WrongPassword => write!(f, "wrong password"),
            Error::Crypto(e) => write!(f, "crypto: {e}"),
        }
    }
}

impl std::error::Error for Error {}

impl From<CfbError> for Error {
    fn from(e: CfbError) -> Self {
        Error::Cfb(e)
    }
}

// ---- record ids used below (MS-XLS §2.1.4 / [MS-XLS] record enumeration) ----
pub const REC_BOF: u16 = 0x0809;
pub const REC_EOF: u16 = 0x000A;
pub const REC_CONTINUE: u16 = 0x003C;
pub const REC_BOUNDSHEET: u16 = 0x0085;
pub const REC_SST: u16 = 0x00FC;
pub const REC_CODEPAGE: u16 = 0x0042;
pub const REC_DATEMODE: u16 = 0x0022;
pub const REC_DIMENSIONS: u16 = 0x0200;
pub const REC_LABELSST: u16 = 0x00FD;
pub const REC_LABEL: u16 = 0x0204;
pub const REC_NUMBER: u16 = 0x0203;
pub const REC_RK: u16 = 0x027E;
pub const REC_MULRK: u16 = 0x00BD;
pub const REC_FORMULA: u16 = 0x0006;
pub const REC_STRING: u16 = 0x0207;
pub const REC_BOOLERR: u16 = 0x0205;
pub const REC_BLANK: u16 = 0x0201;
pub const REC_MULBLANK: u16 = 0x00BE;
pub const REC_FORMAT: u16 = 0x041E;
pub const REC_XF: u16 = 0x00E0;
pub const REC_NAME: u16 = 0x0018;
pub const REC_EXTERNSHEET: u16 = 0x0017;
pub const REC_SUPBOOK: u16 = 0x01AE;
/// `FILEPASS` — present exactly when the workbook is encrypted.
pub const REC_FILEPASS: u16 = 0x002F;
/// Records whose payload is never encrypted.
pub const REC_USREXCL: u16 = 0x0194;
pub const REC_FILELOCK: u16 = 0x0195;
pub const REC_INTERFACEHDR: u16 = 0x00E1;
pub const REC_RRDINFO: u16 = 0x0196;
pub const REC_RRDHEAD: u16 = 0x0138;
/// RC4 blocks in a `.xls` stream are 1024 bytes.
pub const XLS_BLOCK_SIZE: usize = 1024;

/// Substream type from the `BOF.dt` field.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum SheetKind {
    Worksheet,
    Chart,
    MacroSheet,
    Workspace,
    Unknown(u16),
}

impl SheetKind {
    fn from_dt(dt: u16) -> SheetKind {
        match dt {
            0x0005 => SheetKind::Workspace, // globals
            0x0010 => SheetKind::Worksheet,
            0x0020 => SheetKind::Chart,
            0x0040 => SheetKind::MacroSheet,
            other => SheetKind::Unknown(other),
        }
    }
}

/// A cached formula result (the value Excel shows before recalculation).
#[derive(Clone, Debug, PartialEq)]
pub enum CachedValue {
    Blank,
    Number(f64),
    Text(String),
    Bool(bool),
    Error(u8),
}

#[derive(Clone, Debug, PartialEq)]
pub enum CellValue {
    Blank,
    Number(f64),
    Text(String),
    Bool(bool),
    Error(u8),
    Formula {
        cached: CachedValue,
        /// Raw `rgce` tokens (formula RPN), undecoded.
        rgce: Vec<u8>,
        /// Decompiled infix text, e.g. `SUM(A1:A2)*2`. Empty when the tokens
        /// could not be decoded.
        text: String,
    },
}

/// One `XF` record: the formatting a cell points at with its `xf` index.
#[derive(Clone, Copy, Debug, Default)]
pub struct Xf {
    pub font: u16,
    /// Index into the workbook's `FORMAT` table (or the built-in table).
    pub ifmt: u16,
}

impl std::fmt::Display for CachedValue {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            CachedValue::Blank => Ok(()),
            CachedValue::Number(n) => write!(f, "{}", format_number(*n)),
            CachedValue::Text(s) => f.write_str(s),
            CachedValue::Bool(b) => write!(f, "{}", if *b { "TRUE" } else { "FALSE" }),
            CachedValue::Error(c) => write!(f, "#ERR{}", c),
        }
    }
}

impl std::fmt::Display for CellValue {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            CellValue::Blank => Ok(()),
            CellValue::Number(n) => write!(f, "{}", format_number(*n)),
            CellValue::Text(s) => f.write_str(s),
            CellValue::Bool(b) => write!(f, "{}", if *b { "TRUE" } else { "FALSE" }),
            CellValue::Error(c) => write!(f, "#ERR{}", c),
            CellValue::Formula { cached, text, .. } => {
                if text.is_empty() {
                    write!(f, "{cached}")
                } else {
                    write!(f, "={text}")
                }
            }
        }
    }
}

/// Print 3.0 as `3` but keep fractional digits, the way Excel's grid does.
fn format_number(n: f64) -> String {
    if n.fract() == 0.0 && n.abs() < 1e15 {
        format!("{}", n as i64)
    } else {
        format!("{n}")
    }
}

#[derive(Clone, Debug)]
pub struct Cell {
    pub row: u16,
    pub col: u16,
    pub xf: u16,
    pub value: CellValue,
}

#[derive(Clone, Debug)]
pub struct Sheet {
    pub name: String,
    pub kind: SheetKind,
    /// Absolute offset of the sheet's `BOF` record in the workbook stream.
    pub offset: u32,
    pub cells: Vec<Cell>,
}

impl Sheet {
    pub fn cell(&self, row: u16, col: u16) -> Option<&Cell> {
        self.cells.iter().find(|c| c.row == row && c.col == col)
    }
    /// Cells in row-major order.
    pub fn sorted_cells(&self) -> Vec<&Cell> {
        let mut v: Vec<&Cell> = self.cells.iter().collect();
        v.sort_by_key(|c| (c.row, c.col));
        v
    }

    /// Cells grouped by row (rows and cells ascending), for grid printing.
    pub fn rows(&self) -> Vec<Vec<&Cell>> {
        let mut rows: Vec<Vec<&Cell>> = Vec::new();
        for cell in self.sorted_cells() {
            match rows.last_mut() {
                Some(r) if r[0].row == cell.row => r.push(cell),
                _ => rows.push(vec![cell]),
            }
        }
        rows
    }
}

#[derive(Clone, Debug)]
pub struct Workbook {
    pub biff_version: u16,
    pub codepage: u16,
    /// true = 1904 date system (Mac), false = 1900.
    pub date_1904: bool,
    pub sst: Vec<String>,
    /// `FORMAT` records: `ifmt` → format code. Overrides [`numfmt::builtin`].
    pub formats: HashMap<u16, String>,
    /// `XF` records, indexed by the `xf` field of a cell.
    pub xfs: Vec<Xf>,
    /// Defined names (`NAME` records), 1-based as `ptgName` indexes them.
    pub names: Vec<String>,
    /// `EXTERNSHEET` entries resolved to text (`Sheet2`, `[Book]Sheet1`, …).
    pub extern_sheets: Vec<String>,
    pub sheets: Vec<Sheet>,
    /// Charts, in stream order. A chart sheet also has an entry in `sheets`
    /// with `kind == SheetKind::Chart` and the same `offset`.
    pub charts: Vec<chart::Chart>,
    /// The encryption scheme, when the workbook was password protected.
    pub encryption: Option<office_crypto::Scheme>,
}

impl Workbook {
    pub fn sheet(&self, name: &str) -> Option<&Sheet> {
        self.sheets.iter().find(|s| s.name == name)
    }

    /// The format code a cell's `xf` points at, custom records first.
    pub fn format_code(&self, xf: u16) -> Option<&str> {
        let xf = self.xfs.get(xf as usize)?;
        self.formats
            .get(&xf.ifmt)
            .map(|s| s.as_str())
            .or_else(|| numfmt::builtin(xf.ifmt))
    }

    /// What the grid shows for a cell: dates rendered as dates, percentages as
    /// percentages, formula cells as their cached result.
    pub fn cell_text(&self, cell: &Cell) -> String {
        match &cell.value {
            CellValue::Number(n) => match self.format_code(cell.xf) {
                Some(code) => numfmt::format_value(*n, code, self.date_1904),
                None => format_number(*n),
            },
            // A decompiled formula is more informative than its cached value,
            // and that is what Excel's "show formulas" mode displays too.
            CellValue::Formula { text, .. } if !text.is_empty() => format!("={text}"),
            CellValue::Formula { cached, .. } => match cached {
                CachedValue::Number(n) => match self.format_code(cell.xf) {
                    Some(code) => numfmt::format_value(*n, code, self.date_1904),
                    None => format_number(*n),
                },
                other => other.to_string(),
            },
            other => other.to_string(),
        }
    }

    /// The decompiled formula of a cell, if it has one.
    pub fn formula_text<'a>(&self, cell: &'a Cell) -> Option<&'a str> {
        match &cell.value {
            CellValue::Formula { text, .. } if !text.is_empty() => Some(text),
            _ => None,
        }
    }

    /// All sheets as tab-separated text, each prefixed by `== name ==`.
    pub fn plain_text(&self) -> String {
        let mut out = String::new();
        for sheet in &self.sheets {
            out.push_str(&format!("== {} ==\n", sheet.name));
            for row in sheet.rows() {
                let mut first = true;
                for cell in row {
                    if !first {
                        out.push('\t');
                    }
                    first = false;
                    out.push_str(&self.cell_text(cell));
                }
                out.push('\n');
            }
        }
        out
    }
}

/// One BIFF record.
#[derive(Clone, Copy, Debug)]
pub struct Record<'a> {
    pub offset: usize,
    pub kind: u16,
    pub data: &'a [u8],
}

/// Iterate records; stops at a truncated header rather than panicking.
pub fn records(data: &[u8]) -> Vec<Record<'_>> {
    let mut out = Vec::new();
    let mut pos = 0usize;
    while pos + 4 <= data.len() {
        let kind = u16::from_le_bytes([data[pos], data[pos + 1]]);
        let len = u16::from_le_bytes([data[pos + 2], data[pos + 3]]) as usize;
        let start = pos + 4;
        let end = start + len;
        if end > data.len() {
            break;
        }
        out.push(Record {
            offset: pos,
            kind,
            data: &data[start..end],
        });
        pos = end;
    }
    out
}

/// Open and parse a `.xls` file from disk.
pub fn open<P: AsRef<std::path::Path>>(path: P) -> Result<Workbook> {
    let bytes = std::fs::read(path).map_err(CfbError::Io)?;
    parse(&bytes)
}

/// Parse a `.xls` file from its raw bytes (the CFB container).
pub fn parse(bytes: &[u8]) -> Result<Workbook> {
    let cfb = Cfb::parse(bytes.to_vec())?;
    let stream = cfb
        .stream("Workbook")
        .or_else(|_| cfb.stream("Book"))
        .map_err(|_| Error::NoWorkbookStream)?;
    parse_stream(&stream)
}

/// Open a password protected `.xls` file from disk.
pub fn open_with_password<P: AsRef<std::path::Path>>(path: P, password: &str) -> Result<Workbook> {
    let bytes = std::fs::read(path).map_err(CfbError::Io)?;
    parse_with_password(&bytes, password)
}

/// Parse a password protected `.xls` file from its raw bytes.
pub fn parse_with_password(bytes: &[u8], password: &str) -> Result<Workbook> {
    let cfb = Cfb::parse(bytes.to_vec())?;
    let stream = cfb
        .stream("Workbook")
        .or_else(|_| cfb.stream("Book"))
        .map_err(|_| Error::NoWorkbookStream)?;
    parse_stream_with_password(&stream, password)
}

/// Parse a raw BIFF stream, decrypting it with `password` when it is protected.
pub fn parse_stream_with_password(stream: &[u8], password: &str) -> Result<Workbook> {
    // A password supplied for an unprotected workbook is not an error: the file
    // is simply read as-is (the CLI always passes `--password` through).
    if !records(stream).iter().any(|r| r.kind == REC_FILEPASS) {
        return parse_stream_inner(stream, true);
    }
    let (plain, scheme) = decrypt_stream(stream, password)?;
    let mut wb = parse_stream_inner(&plain, true)?;
    wb.encryption = Some(scheme);
    Ok(wb)
}

/// Decrypt a `Workbook` stream and drop the `FILEPASS` record.
///
/// The whole stream is encrypted in 1024-byte RC4 blocks numbered from offset
/// 0, headers included, and then a handful of records are written back in the
/// clear so a reader can still walk the file: `BOF`, `FILEPASS`, `UsrExcl`,
/// `FileLock`, `InterfaceHdr`, `RRDInfo`, `RRDHead`, plus the first four
/// payload bytes of every `BoundSheet8` (`lbPlyPos`).
fn decrypt_stream(stream: &[u8], password: &str) -> Result<(Vec<u8>, office_crypto::Scheme)> {
    let recs = records(stream);
    let filepass = recs
        .iter()
        .find(|r| r.kind == REC_FILEPASS)
        .ok_or(Error::BadRecord("no FILEPASS record"))?;
    // FILEPASS payload: wEncryptionType (u16), then the version pair and the
    // rest of the header exactly as `.doc` stores it.
    let payload = filepass.data;
    // BIFF5 (Excel 5.0/95) writes a four-byte payload — the key and the
    // verifier, with no `wEncryptionType` field at all, because XOR was the
    // only scheme those versions had. Measured on a Microsoft-written file.
    if payload.len() == 4 {
        return decrypt_xor_stream(stream, password, &recs, filepass);
    }
    if payload.len() < 6 {
        return Err(Error::Truncated("FILEPASS"));
    }
    // `wEncryptionType` 0 selects XOR obfuscation, whose header is just the key
    // and the verifier — it does not share the RC4 layout.
    if payload[0] == 0 && payload[1] == 0 {
        return decrypt_xor_stream(stream, password, &recs, filepass);
    }
    let header = office_crypto::parse_header(&payload[2..])?;
    if !header.scheme.verify(password) {
        return Err(Error::WrongPassword);
    }

    let mut plain = office_crypto::decrypt(&header.scheme, password, stream, XLS_BLOCK_SIZE)?;

    // Restore the parts that are stored in the clear, in place: every record
    // header is plaintext (otherwise the stream could not be walked), and a
    // handful of payloads are too.
    for r in &recs {
        let body = r.offset + 4;
        let len = r.data.len();
        plain[r.offset..body].copy_from_slice(&stream[r.offset..body]);
        let keep = match r.kind {
            REC_BOF
            | REC_FILEPASS
            | REC_USREXCL
            | REC_FILELOCK
            | REC_INTERFACEHDR
            | REC_RRDINFO
            | REC_RRDHEAD => len,
            REC_BOUNDSHEET => len.min(4),
            _ => 0,
        };
        if keep > 0 {
            plain[body..body + keep].copy_from_slice(&stream[body..body + keep]);
        }
    }
    Ok((plain, header.scheme))
}

/// Decrypt an XOR-obfuscated stream (MS-OFFCRYPTO "method 1").
///
/// As with RC4 the 4-byte record headers stay readable — without them the
/// stream could not be walked at all — but XOR also keeps whole payloads in the
/// clear: `BOF`, `FilePass`, `UsrExcl`, `FileLock`, `InterfaceHdr`, `RRDInfo`
/// and `RRDHead`. `BoundSheet8` keeps its first four payload bytes (`lbPlyPos`),
/// which is why its keystream index is shifted by 4.
fn decrypt_xor_stream(
    stream: &[u8],
    password: &str,
    recs: &[Record<'_>],
    filepass: &Record<'_>,
) -> Result<(Vec<u8>, office_crypto::Scheme)> {
    let payload = filepass.data;
    // Four-byte payloads are BIFF5 and carry no type field; six-byte ones are
    // BIFF8 and start with `wEncryptionType == 0`.
    let (key, verifier) = if payload.len() == 4 {
        (
            u16::from_le_bytes([payload[0], payload[1]]),
            u16::from_le_bytes([payload[2], payload[3]]),
        )
    } else {
        (
            u16::from_le_bytes([payload[2], payload[3]]),
            u16::from_le_bytes([payload[4], payload[5]]),
        )
    };
    if !office_crypto::xor::verify(password, verifier) {
        return Err(Error::WrongPassword);
    }
    let arr = office_crypto::xor::array(password);
    let mut plain = stream.to_vec();
    for r in recs {
        let total = 4 + r.data.len();
        match r.kind {
            // The record survives as a zeroed stub so the stream length and all
            // following offsets stay valid; only its type and payload go away.
            REC_FILEPASS => {
                plain[r.offset] = 0;
                plain[r.offset + 1] = 0;
                for b in plain[r.offset + 4..r.offset + total].iter_mut() {
                    *b = 0;
                }
            }
            REC_BOF | REC_USREXCL | REC_FILELOCK | REC_INTERFACEHDR | REC_RRDINFO
            | REC_RRDHEAD => {}
            REC_BOUNDSHEET => {
                let keep = r.data.len().min(4);
                let start = r.offset + 4 + keep;
                let end = r.offset + total;
                let run = office_crypto::xor::decrypt_run(&stream[start..end], &arr, start, true);
                plain[start..end].copy_from_slice(&run);
            }
            _ => {
                // Only the payload is obfuscated; the index is anchored at the
                // payload's end offset, so it is `offset + len`.
                let start = r.offset + 4;
                let run = office_crypto::xor::decrypt_run(&stream[start..r.offset + total], &arr, start, false);
                plain[start..r.offset + total].copy_from_slice(&run);
            }
        }
    }
    Ok((plain, office_crypto::Scheme::Xor { key }))
}

/// Parse the raw BIFF stream (no CFB involved).
pub fn parse_stream(stream: &[u8]) -> Result<Workbook> {
    parse_stream_inner(stream, false)
}

fn parse_stream_inner(stream: &[u8], decrypted: bool) -> Result<Workbook> {
    let recs = records(stream);
    if recs.is_empty() || recs[0].kind != REC_BOF {
        return Err(Error::BadRecord("stream does not start with BOF"));
    }
    let bof = recs[0].data;
    if bof.len() < 4 {
        return Err(Error::Truncated("BOF"));
    }
    // Every encrypted workbook carries a FILEPASS record before any content
    // record; the payload is unreadable without the password.
    if !decrypted && recs.iter().any(|r| r.kind == REC_FILEPASS) {
        return Err(Error::Encrypted);
    }
    let biff_version = u16::from_le_bytes([bof[0], bof[1]]);
    let biff8 = biff_version >= 0x0600;
    if !(0x0500..=0x0600).contains(&biff_version) {
        return Err(Error::UnsupportedBiff(biff_version));
    }

    let mut codepage = 1252u16;
    let mut date_1904 = false;
    let mut sst: Vec<String> = Vec::new();
    let mut boundsheets: Vec<(u32, SheetKind, String)> = Vec::new();
    let mut formats: HashMap<u16, String> = HashMap::new();
    let mut xfs: Vec<Xf> = Vec::new();
    let mut names: Vec<String> = Vec::new();
    // EXTERNSHEET entries as (iSupBook, itabFirst, itabLast), resolved below.
    let mut extern_raw: Vec<(u16, u16, u16)> = Vec::new();
    // SUPBOOK: Some(name) for an external workbook, None for the internal one.
    let mut supbooks: Vec<Option<String>> = Vec::new();

    // ---- globals substream: everything up to the first EOF ----------------
    let mut i = 0usize;
    while i < recs.len() {
        let r = recs[i];
        match r.kind {
            REC_EOF => break,
            REC_CODEPAGE if r.data.len() >= 2 => {
                codepage = u16::from_le_bytes([r.data[0], r.data[1]]);
            }
            REC_DATEMODE if r.data.len() >= 2 => {
                date_1904 = u16::from_le_bytes([r.data[0], r.data[1]]) == 1;
            }
            REC_BOUNDSHEET => {
                if let Some(b) = parse_boundsheet(r.data, biff8) {
                    boundsheets.push(b);
                }
            }
            REC_FORMAT if r.data.len() >= 3 => {
                if let Some((ifmt, code)) = parse_format(r.data, biff8) {
                    formats.insert(ifmt, code);
                }
            }
            REC_XF if r.data.len() >= 4 => {
                // font u16, ifmt u16, then alignment/border/fill bytes.
                xfs.push(Xf {
                    font: u16::from_le_bytes([r.data[0], r.data[1]]),
                    ifmt: u16::from_le_bytes([r.data[2], r.data[3]]),
                });
            }
            REC_NAME => {
                if let Some(name) = parse_name(r.data, biff8) {
                    names.push(name);
                }
            }
            REC_EXTERNSHEET if r.data.len() >= 2 => {
                let n = u16::from_le_bytes([r.data[0], r.data[1]]) as usize;
                for k in 0..n {
                    let o = 2 + k * 6;
                    if o + 6 > r.data.len() {
                        break;
                    }
                    extern_raw.push((
                        u16::from_le_bytes([r.data[o], r.data[o + 1]]),
                        u16::from_le_bytes([r.data[o + 2], r.data[o + 3]]),
                        u16::from_le_bytes([r.data[o + 4], r.data[o + 5]]),
                    ));
                }
            }
            REC_SUPBOOK if r.data.len() >= 2 => {
                let ctab = u16::from_le_bytes([r.data[0], r.data[1]]);
                if ctab == 0x0401 || ctab == 0x3A01 {
                    supbooks.push(None); // internal / add-in
                } else {
                    // An internal SUPBOOK is written two ways: `ctab, 0x0401` or
                    // just the sheet count with no workbook name (WPS). Both mean
                    // "this workbook", so an empty name counts as internal.
                    let name = if biff8 {
                        parse_unicode_string(&r.data[2..])
                    } else {
                        let cch = *r.data.get(2).unwrap_or(&0) as usize;
                        decode_cp1252(r.data.get(3..3 + cch).unwrap_or(&[]))
                    };
                    if name.is_empty() {
                        supbooks.push(None);
                    } else {
                        supbooks.push(Some(name));
                    }
                }
            }
            REC_SST => {
                // SST payload + following CONTINUE payloads form one string pool.
                let mut segs: Vec<&[u8]> = vec![r.data];
                let mut j = i + 1;
                while j < recs.len() && recs[j].kind == REC_CONTINUE {
                    segs.push(recs[j].data);
                    j += 1;
                }
                sst = parse_sst(&segs)?;
                i = j;
                continue;
            }
            _ => {}
        }
        i += 1;
    }

    // `EXTERNSHEET` indexes into the BOUNDSHEET list; turn each entry into the
    // text a formula needs (`Sheet2`, `Sheet1:Sheet3`, `[Book.xls]Sheet1`).
    let sheet_names: Vec<String> = boundsheets.iter().map(|b| b.2.clone()).collect();
    let extern_sheets: Vec<String> = extern_raw
        .iter()
        .map(|(isup, first, last)| {
            let book = supbooks
                .get(*isup as usize)
                .and_then(|b| b.clone())
                .map(|b| format!("[{b}]"))
                .unwrap_or_default();
            let range = if first == last {
                sheet_names.get(*first as usize).cloned().unwrap_or_default()
            } else {
                format!(
                    "{}:{}",
                    sheet_names.get(*first as usize).cloned().unwrap_or_default(),
                    sheet_names.get(*last as usize).cloned().unwrap_or_default()
                )
            };
            format!("{book}{range}")
        })
        .collect();

    // offset -> record index, for BOUNDSHEET-directed parsing
    let index_of: HashMap<usize, usize> =
        recs.iter().enumerate().map(|(k, r)| (r.offset, k)).collect();

    let mut sheets = Vec::new();
    for (off, kind, name) in boundsheets {
        let start = index_of.get(&(off as usize)).copied().unwrap_or(0);
        let ctx = FormulaCtx {
            row: 0,
            col: 0,
            sheet: &name,
            names: &names,
            extern_sheets: &extern_sheets,
            biff8,
        };
        let cells = parse_sheet_cells(&recs[start..], &sst, biff8, &ctx)?;
        sheets.push(Sheet {
            name,
            kind,
            offset: off,
            cells,
        });
    }

    // Fallback for files whose BOUNDSHEET offsets are unusable: walk every
    // worksheet substream in file order.
    if sheets.is_empty() {
        for (k, r) in recs.iter().enumerate() {
            if r.kind == REC_BOF && r.data.len() >= 4 {
                let dt = u16::from_le_bytes([r.data[2], r.data[3]]);
                if dt == 0x0010 {
                    let name = format!("Sheet{}", sheets.len() + 1);
                    let ctx = FormulaCtx {
                        row: 0,
                        col: 0,
                        sheet: &name,
                        names: &names,
                        extern_sheets: &extern_sheets,
                        biff8,
                    };
                    let cells = parse_sheet_cells(&recs[k..], &sst, biff8, &ctx)?;
                    sheets.push(Sheet {
                        name,
                        kind: SheetKind::from_dt(dt),
                        offset: r.offset as u32,
                        cells,
                    });
                }
            }
        }
    }

    // Every chart lives in its own substream: a `BOF` whose `dt` is 0x0020.
    // A chart sheet's BOUNDSHEET offset points at it; an embedded chart's
    // substream sits at the end of the sheet stream, so scan all records.
    let mut charts: Vec<chart::Chart> = Vec::new();
    for (k, r) in recs.iter().enumerate() {
        if r.kind == REC_BOF && r.data.len() >= 4 {
            let dt = u16::from_le_bytes([r.data[2], r.data[3]]);
            if dt == chart::BOF_DT_CHART {
                charts.push(chart::parse(&recs, k));
            }
        }
    }

    Ok(Workbook {
        biff_version,
        codepage,
        date_1904,
        sst,
        formats,
        xfs,
        names,
        extern_sheets,
        sheets,
        charts,
        encryption: None,
    })
}

fn parse_boundsheet(data: &[u8], biff8: bool) -> Option<(u32, SheetKind, String)> {
    if data.len() < 7 {
        return None;
    }
    let offset = u32::from_le_bytes([data[0], data[1], data[2], data[3]]);
    // byte 4 = visibility, byte 5 = sheet type (0x00 worksheet, 0x01 macro,
    // 0x02 chart, 0x06 VBA module)
    let kind = match data[5] {
        0x00 => SheetKind::Worksheet,
        0x01 => SheetKind::MacroSheet,
        0x02 => SheetKind::Chart,
        other => SheetKind::Unknown(other as u16),
    };
    let name = if biff8 {
        // XLUnicodeString: cch u8, flags u8, chars
        let cch = *data.get(6)? as usize;
        let flags = *data.get(7)?;
        let high = flags & 1 != 0;
        let body = data.get(8..)?;
        if high {
            let bytes = body.get(..cch * 2)?;
            Some(String::from_utf16_lossy(
                &bytes
                    .chunks_exact(2)
                    .map(|c| u16::from_le_bytes([c[0], c[1]]))
                    .collect::<Vec<_>>(),
            ))
        } else {
            Some(decode_cp1252(body.get(..cch)?))
        }
    } else {
        // BIFF5: cch u8, then bytes
        let cch = *data.get(6)? as usize;
        Some(decode_cp1252(data.get(7..7 + cch)?))
    };
    Some((offset, kind, name?))
}

/// Reads a BIFF8 string pool spread over SST + CONTINUE segments. When a
/// string's characters straddle a segment boundary, the continuation starts
/// with a fresh option-flags byte (MS-XLS §2.5.293).
struct SegReader<'a> {
    segs: Vec<&'a [u8]>,
    seg: usize,
    pos: usize,
}

impl<'a> SegReader<'a> {
    fn new(segs: Vec<&'a [u8]>) -> Self {
        SegReader { segs, seg: 0, pos: 0 }
    }

    fn advance(&mut self) -> bool {
        if self.seg + 1 < self.segs.len() {
            self.seg += 1;
            self.pos = 0;
            true
        } else {
            false
        }
    }

    fn u8(&mut self) -> Option<u8> {
        while self.seg < self.segs.len() {
            if let Some(&b) = self.segs[self.seg].get(self.pos) {
                self.pos += 1;
                return Some(b);
            }
            if !self.advance() {
                return None;
            }
        }
        None
    }

    fn u16(&mut self) -> Option<u16> {
        Some(u16::from_le_bytes([self.u8()?, self.u8()?]))
    }

    fn u32(&mut self) -> Option<u32> {
        Some(u32::from_le_bytes([self.u8()?, self.u8()?, self.u8()?, self.u8()?]))
    }

    fn take(&mut self, n: usize) -> Option<Vec<u8>> {
        let mut out = Vec::with_capacity(n);
        while out.len() < n {
            if self.seg >= self.segs.len() {
                return None;
            }
            let cur = self.segs[self.seg];
            if self.pos >= cur.len() {
                if !self.advance() {
                    return None;
                }
                continue;
            }
            let take = (n - out.len()).min(cur.len() - self.pos);
            out.extend_from_slice(&cur[self.pos..self.pos + take]);
            self.pos += take;
        }
        Some(out)
    }

    fn string(&mut self) -> Option<String> {
        let cch = self.u16()? as usize;
        let flags = self.u8()?;
        let mut high = flags & 0x01 != 0;
        let rich = flags & 0x08 != 0;
        let ext = flags & 0x04 != 0;
        let c_run = if rich { self.u16()? as usize } else { 0 };
        let cb_ext = if ext { self.u32()? as usize } else { 0 };

        let mut chars: Vec<char> = Vec::with_capacity(cch);
        while chars.len() < cch {
            // At a segment boundary the continuation repeats the flags byte.
            if self.pos >= self.segs[self.seg].len() {
                if !self.advance() {
                    return None;
                }
                let f = *self.segs[self.seg].first()?;
                self.pos = 1;
                high = f & 0x01 != 0;
            }
            let avail = self.segs[self.seg].len() - self.pos;
            let can = if high { avail / 2 } else { avail };
            if can == 0 {
                if !self.advance() {
                    return None;
                }
                let f = *self.segs[self.seg].first()?;
                self.pos = 1;
                high = f & 0x01 != 0;
                continue;
            }
            let take = (cch - chars.len()).min(can);
            let cur = self.segs[self.seg];
            if high {
                for k in 0..take {
                    let o = self.pos + k * 2;
                    let unit = u16::from_le_bytes([cur[o], cur[o + 1]]);
                    chars.push(char::from_u32(unit as u32).unwrap_or('\u{FFFD}'));
                }
            } else {
                for k in 0..take {
                    let b = cur[self.pos + k];
                    chars.push(if (0x80..=0x9F).contains(&b) {
                        const HIGH: [char; 32] = [
                            '\u{20AC}', '\u{0081}', '\u{201A}', '\u{0192}', '\u{201E}', '\u{2026}',
                            '\u{2020}', '\u{2021}', '\u{02C6}', '\u{2030}', '\u{0160}', '\u{2039}',
                            '\u{0152}', '\u{008D}', '\u{017D}', '\u{008F}', '\u{0090}', '\u{2018}',
                            '\u{2019}', '\u{201C}', '\u{201D}', '\u{2022}', '\u{2013}', '\u{2014}',
                            '\u{02DC}', '\u{2122}', '\u{0161}', '\u{203A}', '\u{0153}', '\u{009D}',
                            '\u{017E}', '\u{0178}',
                        ];
                        HIGH[(b - 0x80) as usize]
                    } else {
                        b as char
                    });
                }
            }
            self.pos += take * if high { 2 } else { 1 };
        }
        if c_run > 0 {
            self.take(c_run * 4)?;
        }
        if cb_ext > 0 {
            self.take(cb_ext)?;
        }
        Some(chars.into_iter().collect())
    }
}

fn parse_sst(segs: &[&[u8]]) -> Result<Vec<String>> {
    let mut r = SegReader::new(segs.to_vec());
    // Both counters are 4 bytes wide (MS-XLS 2.4.265), not 2 — a 2-byte read
    // makes every BIFF8 pool look empty because cstUnique comes out 0.
    let _total = r.u32().ok_or(Error::Truncated("SST cstTotal"))?;
    let unique = r.u32().ok_or(Error::Truncated("SST cstUnique"))?;
    let unique = unique.min(1_000_000); // sanity cap against corrupt pools
    let mut out = Vec::with_capacity(unique as usize);
    for _ in 0..unique {
        match r.string() {
            Some(s) => out.push(s),
            None => break, // tolerate a truncated pool rather than failing the file
        }
    }
    Ok(out)
}

/// Parse cells of one substream starting at its `BOF`.
fn parse_sheet_cells(
    recs: &[Record<'_>],
    sst: &[String],
    biff8: bool,
    ctx: &FormulaCtx<'_>,
) -> Result<Vec<Cell>> {
    let mut cells = Vec::new();
    let mut pending_formula: Option<usize> = None;
    for (k, r) in recs.iter().enumerate() {
        if k == 0 {
            continue; // BOF
        }
        match r.kind {
            REC_EOF => break,
            REC_LABELSST => {
                if r.data.len() < 10 {
                    continue;
                }
                let row = u16::from_le_bytes([r.data[0], r.data[1]]);
                let col = u16::from_le_bytes([r.data[2], r.data[3]]);
                let xf = u16::from_le_bytes([r.data[4], r.data[5]]);
                let isst = u32::from_le_bytes([r.data[6], r.data[7], r.data[8], r.data[9]]);
                let text = sst.get(isst as usize).cloned().unwrap_or_default();
                cells.push(Cell {
                    row,
                    col,
                    xf,
                    value: CellValue::Text(text),
                });
            }
            REC_LABEL => {
                if r.data.len() < 6 {
                    continue;
                }
                let row = u16::from_le_bytes([r.data[0], r.data[1]]);
                let col = u16::from_le_bytes([r.data[2], r.data[3]]);
                let xf = u16::from_le_bytes([r.data[4], r.data[5]]);
                let text = if biff8 {
                    parse_unicode_string(&r.data[6..])
                } else {
                    let cch = if r.data.len() >= 8 {
                        u16::from_le_bytes([r.data[6], r.data[7]]) as usize
                    } else {
                        0
                    };
                    decode_cp1252(r.data.get(8..8 + cch).unwrap_or(&[]))
                };
                cells.push(Cell {
                    row,
                    col,
                    xf,
                    value: CellValue::Text(text),
                });
            }
            REC_NUMBER => {
                if r.data.len() < 14 {
                    continue;
                }
                let row = u16::from_le_bytes([r.data[0], r.data[1]]);
                let col = u16::from_le_bytes([r.data[2], r.data[3]]);
                let xf = u16::from_le_bytes([r.data[4], r.data[5]]);
                let v = f64::from_le_bytes(r.data[6..14].try_into().unwrap());
                cells.push(Cell {
                    row,
                    col,
                    xf,
                    value: CellValue::Number(v),
                });
            }
            REC_RK => {
                if r.data.len() < 10 {
                    continue;
                }
                let row = u16::from_le_bytes([r.data[0], r.data[1]]);
                let col = u16::from_le_bytes([r.data[2], r.data[3]]);
                let xf = u16::from_le_bytes([r.data[4], r.data[5]]);
                let rk = u32::from_le_bytes([r.data[6], r.data[7], r.data[8], r.data[9]]);
                cells.push(Cell {
                    row,
                    col,
                    xf,
                    value: CellValue::Number(decode_rk(rk)),
                });
            }
            REC_MULRK => {
                if r.data.len() < 6 {
                    continue;
                }
                let row = u16::from_le_bytes([r.data[0], r.data[1]]);
                let first = u16::from_le_bytes([r.data[2], r.data[3]]);
                let n = (r.data.len().saturating_sub(6)) / 6;
                for j in 0..n {
                    let o = 4 + j * 6;
                    let xf = u16::from_le_bytes([r.data[o], r.data[o + 1]]);
                    let rk = u32::from_le_bytes([
                        r.data[o + 2],
                        r.data[o + 3],
                        r.data[o + 4],
                        r.data[o + 5],
                    ]);
                    cells.push(Cell {
                        row,
                        col: first + j as u16,
                        xf,
                        value: CellValue::Number(decode_rk(rk)),
                    });
                }
            }
            REC_BOOLERR => {
                if r.data.len() < 8 {
                    continue;
                }
                let row = u16::from_le_bytes([r.data[0], r.data[1]]);
                let col = u16::from_le_bytes([r.data[2], r.data[3]]);
                let xf = u16::from_le_bytes([r.data[4], r.data[5]]);
                let val = r.data[6];
                let is_err = r.data[7] != 0;
                cells.push(Cell {
                    row,
                    col,
                    xf,
                    value: if is_err {
                        CellValue::Error(val)
                    } else {
                        CellValue::Bool(val != 0)
                    },
                });
            }
            REC_BLANK => {
                if r.data.len() < 6 {
                    continue;
                }
                cells.push(Cell {
                    row: u16::from_le_bytes([r.data[0], r.data[1]]),
                    col: u16::from_le_bytes([r.data[2], r.data[3]]),
                    xf: u16::from_le_bytes([r.data[4], r.data[5]]),
                    value: CellValue::Blank,
                });
            }
            REC_MULBLANK => {
                if r.data.len() < 6 {
                    continue;
                }
                let row = u16::from_le_bytes([r.data[0], r.data[1]]);
                let first = u16::from_le_bytes([r.data[2], r.data[3]]);
                let n = (r.data.len().saturating_sub(6)) / 2;
                for j in 0..n {
                    let o = 4 + j * 2;
                    cells.push(Cell {
                        row,
                        col: first + j as u16,
                        xf: u16::from_le_bytes([r.data[o], r.data[o + 1]]),
                        value: CellValue::Blank,
                    });
                }
            }
            REC_FORMULA => {
                // row u16, col u16, xf u16, result 8 B, grbit u16, chn u32,
                // cce u16, then the token array itself.
                if r.data.len() < 22 {
                    continue;
                }
                let row = u16::from_le_bytes([r.data[0], r.data[1]]);
                let col = u16::from_le_bytes([r.data[2], r.data[3]]);
                let xf = u16::from_le_bytes([r.data[4], r.data[5]]);
                let res = &r.data[6..14];
                let cce = u16::from_le_bytes([r.data[20], r.data[21]]) as usize;
                let rgce = r.data.get(22..22 + cce).unwrap_or(&[]).to_vec();
                let cached = if res[6] == 0xFF && res[7] == 0xFF {
                    match res[0] {
                        0x00 => CachedValue::Text(String::new()), // value in following STRING
                        0x01 => CachedValue::Bool(res[2] != 0),
                        0x02 => CachedValue::Error(res[2]),
                        0x03 => CachedValue::Blank,
                        _ => CachedValue::Blank,
                    }
                } else {
                    CachedValue::Number(f64::from_le_bytes(res.try_into().unwrap()))
                };
                cells.push(Cell {
                    row,
                    col,
                    xf,
                    value: CellValue::Formula {
                        cached,
                        rgce,
                        text: String::new(),
                    },
                });
                pending_formula = Some(cells.len() - 1);
            }
            REC_STRING => {
                // Only valid immediately after a FORMULA with a text result.
                if let Some(idx) = pending_formula.take() {
                    let s = if biff8 {
                        parse_unicode_string(r.data)
                    } else {
                        let cch = if r.data.len() >= 2 {
                            u16::from_le_bytes([r.data[0], r.data[1]]) as usize
                        } else {
                            0
                        };
                        decode_cp1252(r.data.get(2..2 + cch).unwrap_or(&[]))
                    };
                    if let CellValue::Formula { cached, .. } = &mut cells[idx].value {
                        *cached = CachedValue::Text(s);
                    }
                }
            }
            _ => {}
        }
    }

    // Formulas are decompiled in a second pass: a shared formula's followers
    // hold only a `ptgExp` pointer at the master cell, which may come later in
    // the record stream.
    let mut masters: HashMap<(u16, u16), Vec<u8>> = HashMap::new();
    for cell in &cells {
        if let CellValue::Formula { rgce, .. } = &cell.value {
            if formula::shared_master(rgce).is_none() && !rgce.is_empty() {
                masters.insert((cell.row, cell.col), rgce.clone());
            }
        }
    }
    for cell in &mut cells {
        let (row, col) = (cell.row, cell.col);
        if let CellValue::Formula { rgce, text, .. } = &mut cell.value {
            let effective = match formula::shared_master(rgce) {
                Some((mr, mc)) => masters
                    .get(&(mr, mc))
                    .map(|m| {
                        formula::shift_relative(
                            m,
                            row as i32 - mr as i32,
                            col as i32 - mc as i32,
                        )
                    })
                    .unwrap_or_default(),
                None => rgce.clone(),
            };
            if effective.is_empty() {
                continue;
            }
            let local = FormulaCtx {
                row,
                col,
                sheet: ctx.sheet,
                names: ctx.names,
                extern_sheets: ctx.extern_sheets,
                biff8: ctx.biff8,
            };
            *text = formula::decompile(&effective, &local);
        }
    }
    Ok(cells)
}

/// `FORMAT` record: `ifmt` u16 then the format code as a string.
fn parse_format(data: &[u8], biff8: bool) -> Option<(u16, String)> {
    if data.len() < 3 {
        return None;
    }
    let ifmt = u16::from_le_bytes([data[0], data[1]]);
    let code = if biff8 {
        parse_unicode_string(&data[2..])
    } else {
        let cch = data[2] as usize;
        decode_cp1252(data.get(3..3 + cch).unwrap_or(&[]))
    };
    Some((ifmt, code))
}

/// `NAME` record → the defined name text (the parsed formula is skipped).
fn parse_name(data: &[u8], _biff8: bool) -> Option<String> {
    // grbit u16, chKey u8, cch u8, cce u16, ixals u16, itab u16, then four
    // u8 counts (custom menu / description / help topic / status text).
    if data.len() < 14 {
        return None;
    }
    let cch = data[3] as usize;
    let body = data.get(14..14 + cch)?;
    if cch == 0 {
        return None;
    }
    // BIFF8 stores the name as UTF-16 when the high-byte flag is set; the
    // reliable signal is the byte pattern, since ASCII names encode as
    // `P\0r\0i\0n\0t\0`.
    let looks_wide = body.len() >= 2 && body.chunks(2).all(|c| c.len() < 2 || c[1] == 0);
    Some(if looks_wide && cch > 1 {
        office_enc::decode_utf16le(body)
    } else {
        decode_cp1252(body)
    })
}

/// BIFF8 `XLUnicodeString` (no CONTINUE): cch u16, flags u8, chars, runs, ext.
pub fn parse_unicode_string(data: &[u8]) -> String {
    if data.len() < 3 {
        return String::new();
    }
    let cch = u16::from_le_bytes([data[0], data[1]]) as usize;
    let flags = data[2];
    let high = flags & 0x01 != 0;
    let rich = flags & 0x08 != 0;
    let ext = flags & 0x04 != 0;
    let mut off = 3usize;
    let c_run = if rich && data.len() >= off + 2 {
        let n = u16::from_le_bytes([data[off], data[off + 1]]) as usize;
        off += 2;
        n
    } else {
        0
    };
    let cb_ext = if ext && data.len() >= off + 4 {
        let n = u32::from_le_bytes([data[off], data[off + 1], data[off + 2], data[off + 3]]) as usize;
        off += 4;
        n
    } else {
        0
    };
    let s = if high {
        let bytes = data.get(off..off + cch * 2).unwrap_or(&[]);
        String::from_utf16_lossy(
            &bytes
                .chunks_exact(2)
                .map(|c| u16::from_le_bytes([c[0], c[1]]))
                .collect::<Vec<_>>(),
        )
    } else {
        decode_cp1252(data.get(off..off + cch).unwrap_or(&[]))
    };
    let _ = (c_run, cb_ext);
    s
}

/// Decode a 32-bit RK value: bit 1 = integer, bit 0 = divide by 100.
///
/// * integer form: `(rk as i32) >> 2`
/// * double form: the top 32 bits of an IEEE-754 double whose low 32 bits are
///   zero, i.e. `f64::from_bits((rk & 0xFFFF_FFFC) as u64) << 32)`
pub fn decode_rk(rk: u32) -> f64 {
    let mut v = if rk & 0x02 != 0 {
        ((rk as i32) >> 2) as f64
    } else {
        f64::from_bits(((rk & 0xFFFF_FFFC) as u64) << 32)
    };
    if rk & 0x01 != 0 {
        v /= 100.0;
    }
    v
}

#[cfg(test)]
mod biff5_tests {
    //! A hand-built BIFF5 (Excel 5.0/95) workbook stream.
    //!
    //! `samples/ms/excel95.xls` (written by Microsoft Excel via COM) is the real
    //! thing and is covered by the integration tests; this synthetic stream
    //! exists so the reader can be exercised without that file.
    use super::*;

    fn rec(kind: u16, body: &[u8], out: &mut Vec<u8>) {
        out.extend_from_slice(&kind.to_le_bytes());
        out.extend_from_slice(&(body.len() as u16).to_le_bytes());
        out.extend_from_slice(body);
    }

    fn bof(version: u16, dt: u16, out: &mut Vec<u8>) {
        let mut b = Vec::new();
        b.extend_from_slice(&version.to_le_bytes());
        b.extend_from_slice(&dt.to_le_bytes());
        b.extend_from_slice(&0x0E5Bu16.to_le_bytes()); // rupBuild
        b.extend_from_slice(&0x07C9u16.to_le_bytes()); // rupYear
        b.extend_from_slice(&0u32.to_le_bytes()); // flags
        b.extend_from_slice(&0x0500u32.to_le_bytes()); // lowest version
        rec(REC_BOF, &b, out);
    }

    fn label(row: u16, col: u16, text: &[u8], out: &mut Vec<u8>) {
        let mut b = Vec::new();
        b.extend_from_slice(&row.to_le_bytes());
        b.extend_from_slice(&col.to_le_bytes());
        b.extend_from_slice(&0u16.to_le_bytes()); // xf
        b.extend_from_slice(&(text.len() as u16).to_le_bytes());
        b.extend_from_slice(text); // BIFF5: no grbit, code page bytes
        rec(REC_LABEL, &b, out);
    }

    fn number(row: u16, col: u16, v: f64, out: &mut Vec<u8>) {
        let mut b = Vec::new();
        b.extend_from_slice(&row.to_le_bytes());
        b.extend_from_slice(&col.to_le_bytes());
        b.extend_from_slice(&0u16.to_le_bytes());
        b.extend_from_slice(&v.to_le_bytes());
        rec(REC_NUMBER, &b, out);
    }

    fn workbook_stream() -> Vec<u8> {
        let mut sheet = Vec::new();
        bof(0x0500, 0x0010, &mut sheet);
        label(0, 0, b"BIFF5-TEXT", &mut sheet);
        // 0xE9 is 'é' in CP1252, so the byte must be spelled out.
        label(1, 0, b"caf\xE9", &mut sheet);
        number(2, 0, 1234.5, &mut sheet);
        rec(REC_EOF, &[], &mut sheet);

        let mut s = Vec::new();
        bof(0x0500, 0x0005, &mut s);
        rec(REC_CODEPAGE, &1252u16.to_le_bytes(), &mut s);
        rec(REC_DATEMODE, &0u16.to_le_bytes(), &mut s);

        // BoundSheet8 needs the sheet offset, which depends on its own length:
        // build it with a placeholder, then patch the offset in.
        let mut boundsheet = Vec::new();
        boundsheet.extend_from_slice(&0u32.to_le_bytes());
        boundsheet.push(0); // grbit: visible
        boundsheet.push(0); // dt: worksheet
        boundsheet.push(4); // cch
        boundsheet.extend_from_slice(b"Old5");
        let sheet_at = s.len() + 4 + boundsheet.len();
        boundsheet[0..4].copy_from_slice(&(sheet_at as u32).to_le_bytes());
        rec(REC_BOUNDSHEET, &boundsheet, &mut s);

        rec(REC_EOF, &[], &mut s);
        s.extend_from_slice(&sheet);
        s
    }

    /// A chart sheet is a `BOUNDSHEET` with `dt == 2` pointing at a substream
    /// whose `BOF` has `dt == 0x0020`. Its records are chart-specific, so the
    /// sheet must still be *enumerated* (name + kind) with no cells rather than
    /// being dropped or parsed as a worksheet.
    #[test]
    fn enumerates_a_chart_sheet_without_cells() {
        let mut chart = Vec::new();
        bof(0x0600, 0x0020, &mut chart);
        rec(0x1001, &[0u8; 8], &mut chart); // Chart
        rec(0x1002, &[0u8; 4], &mut chart); // chart format-ish
        rec(REC_EOF, &[], &mut chart);

        let mut sheet = Vec::new();
        bof(0x0600, 0x0010, &mut sheet);
        label(0, 0, b"ChartSrc", &mut sheet);
        rec(REC_EOF, &[], &mut sheet);

        let mut s = Vec::new();
        bof(0x0600, 0x0005, &mut s);
        rec(REC_CODEPAGE, &1200u16.to_le_bytes(), &mut s);
        // Two BoundSheet8 records; offsets are patched once both are laid out.
        let mut bs = Vec::new();
        for (name, dt) in [("Data", 0u8), ("Chart1", 2u8)] {
            bs.push((name, dt));
        }
        let mut patches = Vec::new();
        for (name, dt) in &bs {
            let mut b = Vec::new();
            b.extend_from_slice(&0u32.to_le_bytes());
            b.push(0);
            b.push(*dt);
            b.push(name.len() as u8);
            b.push(1); // grbit: UTF-16 name
            for ch in name.encode_utf16() {
                b.extend_from_slice(&ch.to_le_bytes());
            }
            patches.push((s.len() + 4, b.len()));
            rec(REC_BOUNDSHEET, &b, &mut s);
        }
        rec(REC_EOF, &[], &mut s);
        let sheet_at = s.len() as u32;
        s.extend_from_slice(&sheet);
        let chart_at = s.len() as u32;
        s.extend_from_slice(&chart);
        // First BoundSheet8 points at the worksheet, second at the chart.
        for (k, (at, len)) in patches.iter().enumerate() {
            let off = if k == 0 { sheet_at } else { chart_at };
            s[*at..*at + 4].copy_from_slice(&off.to_le_bytes());
            let _ = len;
        }

        let wb = parse_stream(&s).expect("workbook with a chart sheet");
        assert_eq!(wb.sheets.len(), 2);
        assert_eq!(wb.sheets[0].name, "Data");
        assert_eq!(wb.sheets[0].kind, SheetKind::Worksheet);
        assert_eq!(wb.sheets[0].cells.len(), 1);
        assert_eq!(wb.sheets[1].name, "Chart1");
        assert_eq!(wb.sheets[1].kind, SheetKind::Chart);
        assert!(wb.sheets[1].cells.is_empty());
    }

    #[test]
    fn parses_a_biff5_workbook() {
        let stream = workbook_stream();
        let wb = parse_stream(&stream).expect("BIFF5 stream");
        assert_eq!(wb.biff_version, 0x0500);
        assert_eq!(wb.codepage, 1252);
        assert_eq!(wb.sheets.len(), 1);
        let sheet = &wb.sheets[0];
        assert_eq!(sheet.name, "Old5");
        assert_eq!(sheet.cells.len(), 3);
        match &sheet.cells[0].value {
            CellValue::Text(t) => assert_eq!(t, "BIFF5-TEXT"),
            other => panic!("expected text, got {other:?}"),
        }
        match &sheet.cells[1].value {
            CellValue::Text(t) => assert_eq!(t, "café"),
            other => panic!("expected text, got {other:?}"),
        }
        match &sheet.cells[2].value {
            CellValue::Number(v) => assert_eq!(*v, 1234.5),
            other => panic!("expected number, got {other:?}"),
        }
    }

    /// The token arrays below were copied byte-for-byte out of
    /// `samples/ms/excel95.xls` and `samples/ms/biff5mix.xls`, which Microsoft
    /// Excel 95 wrote. They pin the BIFF5 reference encoding: the column is one
    /// byte and the two relative flags live in the top of the row word.
    #[test]
    fn decompiles_real_biff5_token_arrays() {
        use crate::formula::{decompile, Ctx};
        let ctx = Ctx {
            row: 0,
            col: 1,
            sheet: "Sheet1",
            names: &[],
            extern_sheets: &[],
            biff8: false,
        };
        let cases: &[(&[u8], &str)] = &[
            // =A2  (row word 0xC001 = fColRel|fRwRel + row 1, column 0)
            (&[0x44, 0x01, 0xC0, 0x00], "A2"),
            // =A2+1
            (&[0x44, 0x01, 0xC0, 0x00, 0x1E, 0x01, 0x00, 0x03], "A2+1"),
            // =A2*2
            (&[0x44, 0x01, 0xC0, 0x00, 0x1E, 0x02, 0x00, 0x05], "A2*2"),
            // =B1+C1
            (&[0x44, 0x00, 0xC0, 0x01, 0x44, 0x00, 0xC0, 0x02, 0x03], "B1+C1"),
            // =SUM(A1:A3) — ptgArea then the 4-byte ptgAttr with grbit 0x10
            (
                &[0x25, 0x00, 0xC0, 0x02, 0xC0, 0x00, 0x00, 0x19, 0x10, 0x00, 0x00],
                "SUM(A1:A3)",
            ),
        ];
        for (rgce, want) in cases {
            assert_eq!(decompile(rgce, &ctx), *want, "rgce {rgce:02x?}");
        }
        // The same bytes read as BIFF8 give nonsense, which is the point: the
        // two encodings are not interchangeable.
        let biff8 = Ctx { biff8: true, ..ctx };
        assert_ne!(decompile(&[0x44, 0x01, 0xC0, 0x00], &biff8), "A2");
    }
}

#[cfg(test)]
mod xor_tests {
    //! XOR obfuscation (MS-OFFCRYPTO "method 1"), tested by *encrypting* a
    //! synthetic workbook with the inverse transform and reading it back.
    //!
    //! `samples/ms/xor.xls` is a real Microsoft-written XOR workbook (Excel 95
    //! with a write password) and is covered by the integration tests; this
    //! synthetic stream keeps the unit test independent of that file.
    use super::*;

    const PW: &str = "secret123";

    fn rec(kind: u16, body: &[u8], out: &mut Vec<u8>) {
        out.extend_from_slice(&kind.to_le_bytes());
        out.extend_from_slice(&(body.len() as u16).to_le_bytes());
        out.extend_from_slice(body);
    }

    fn bof(version: u16, dt: u16, out: &mut Vec<u8>) {
        let mut b = Vec::new();
        b.extend_from_slice(&version.to_le_bytes());
        b.extend_from_slice(&dt.to_le_bytes());
        b.extend_from_slice(&0x0E5Bu16.to_le_bytes());
        b.extend_from_slice(&0x07C9u16.to_le_bytes());
        b.extend_from_slice(&0u32.to_le_bytes());
        b.extend_from_slice(&0x0600u32.to_le_bytes());
        rec(REC_BOF, &b, out);
    }

    /// Inverse of `office_crypto::xor::decrypt_run`.
    fn encrypt_run(data: &[u8], arr: &[u8; 16], offset: usize, boundsheet: bool) -> Vec<u8> {
        let mut index = (offset + data.len()) % 16;
        if boundsheet {
            index = (index + 4) % 16;
        }
        let mut out = Vec::with_capacity(data.len());
        for byte in data {
            out.push(byte.rotate_left(5) ^ arr[index]);
            index = (index + 1) % 16;
        }
        out
    }

    /// Build a workbook whose `Workbook` stream is XOR obfuscated.
    fn xor_workbook() -> Vec<u8> {
        // 1. Plain layout, with the sheet offset left for later.
        let mut sheet = Vec::new();
        bof(0x0600, 0x0010, &mut sheet);
        let mut label = Vec::new();
        label.extend_from_slice(&0u16.to_le_bytes()); // row
        label.extend_from_slice(&0u16.to_le_bytes()); // col
        label.extend_from_slice(&0u16.to_le_bytes()); // xf
        label.extend_from_slice(&5u16.to_le_bytes()); // cch: 5 characters
        label.push(1); // grbit: UTF-16
        label.extend_from_slice(&[b'O', 0, b'B', 0, b'F', 0, b'U', 0, b'S', 0]);
        rec(REC_LABEL, &label, &mut sheet);
        let mut num = Vec::new();
        num.extend_from_slice(&1u16.to_le_bytes());
        num.extend_from_slice(&0u16.to_le_bytes());
        num.extend_from_slice(&0u16.to_le_bytes());
        num.extend_from_slice(&42.5f64.to_le_bytes());
        rec(REC_NUMBER, &num, &mut sheet);
        rec(REC_EOF, &[], &mut sheet);

        let mut s = Vec::new();
        bof(0x0600, 0x0005, &mut s);
        // The FILEPASS record is 10 bytes and goes right after the 20-byte BOF,
        // so every later offset moves by 10.
        let mut filepass = Vec::new();
        filepass.extend_from_slice(&0u16.to_le_bytes()); // wEncryptionType: XOR
        filepass.extend_from_slice(&office_crypto::xor::key(PW).to_le_bytes());
        filepass.extend_from_slice(&office_crypto::xor::verifier(PW).to_le_bytes());
        rec(REC_FILEPASS, &filepass, &mut s);
        rec(REC_CODEPAGE, &1200u16.to_le_bytes(), &mut s);
        rec(REC_DATEMODE, &0u16.to_le_bytes(), &mut s);

        let mut boundsheet = Vec::new();
        boundsheet.extend_from_slice(&0u32.to_le_bytes());
        boundsheet.push(0);
        boundsheet.push(0);
        boundsheet.push(2); // cch
        boundsheet.push(1); // grbit: UTF-16
        boundsheet.extend_from_slice(&[b'O', 0, b'B', 0]);
        let sheet_at = s.len() + 4 + boundsheet.len();
        boundsheet[0..4].copy_from_slice(&(sheet_at as u32).to_le_bytes());
        rec(REC_BOUNDSHEET, &boundsheet, &mut s);
        rec(REC_EOF, &[], &mut s);
        s.extend_from_slice(&sheet);

        // 2. Obfuscate in place, honouring the plaintext exceptions.
        let arr = office_crypto::xor::array(PW);
        let recs = records(&s);
        let mut out = s.clone();
        for r in &recs {
            let total = 4 + r.data.len();
            match r.kind {
                REC_FILEPASS
                | REC_BOF
                | REC_USREXCL
                | REC_FILELOCK
                | REC_INTERFACEHDR
                | REC_RRDINFO
                | REC_RRDHEAD => {}
                REC_BOUNDSHEET => {
                    let keep = r.data.len().min(4);
                    let start = r.offset + 4 + keep;
                    let end = r.offset + total;
                    let run = encrypt_run(&s[start..end], &arr, start, true);
                    out[start..end].copy_from_slice(&run);
                }
                _ => {
                    let start = r.offset + 4;
                    let run = encrypt_run(&s[start..r.offset + total], &arr, start, false);
                    out[start..r.offset + total].copy_from_slice(&run);
                }
            }
        }
        out
    }

    #[test]
    fn decrypts_an_xor_obfuscated_workbook() {
        let encrypted = xor_workbook();
        // The obfuscated stream must not look like a readable workbook...
        assert!(parse_stream(&encrypted).is_err());
        // ...but opens with the password.
        let wb = parse_stream_with_password(&encrypted, PW).expect("XOR workbook");
        assert_eq!(wb.biff_version, 0x0600);
        match &wb.encryption {
            Some(office_crypto::Scheme::Xor { key }) => {
                assert_eq!(*key, office_crypto::xor::key(PW))
            }
            other => panic!("expected XOR scheme, got {other:?}"),
        }
        assert_eq!(wb.sheets.len(), 1);
        let sheet = &wb.sheets[0];
        assert_eq!(sheet.name, "OB");
        assert_eq!(sheet.cells.len(), 2);
        match &sheet.cells[0].value {
            CellValue::Text(t) => assert_eq!(t, "OBFUS"),
            other => panic!("expected text, got {other:?}"),
        }
        match &sheet.cells[1].value {
            CellValue::Number(v) => assert_eq!(*v, 42.5),
            other => panic!("expected number, got {other:?}"),
        }
    }

    #[test]
    fn rejects_a_wrong_password_on_an_xor_workbook() {
        let encrypted = xor_workbook();
        assert!(matches!(
            parse_stream_with_password(&encrypted, "nope"),
            Err(Error::WrongPassword)
        ));
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn refuses_an_encrypted_workbook_without_a_password() {
        let p = concat!(env!("CARGO_MANIFEST_DIR"), "/../../samples/encrypted.xls");
        if !std::path::Path::new(p).exists() {
            return;
        }
        let bytes = std::fs::read(p).unwrap();
        assert!(matches!(parse(&bytes), Err(Error::Encrypted)));
        assert!(matches!(
            parse_with_password(&bytes, "wrong-password"),
            Err(Error::WrongPassword)
        ));
    }

    #[test]
    fn decrypts_the_sample_workbook() {
        let p = concat!(env!("CARGO_MANIFEST_DIR"), "/../../samples/encrypted.xls");
        if !std::path::Path::new(p).exists() {
            return;
        }
        let bytes = std::fs::read(p).unwrap();
        let wb = parse_with_password(&bytes, "secret123").expect("decrypt");
        assert_eq!(wb.encryption.as_ref().map(|s| s.name()), Some("RC4 (legacy, v1.1)"));
        assert_eq!(wb.biff_version, 0x0600);
        let s = wb.sheets.first().expect("one sheet");
        assert!(s.cells.iter().any(|c| matches!(&c.value, CellValue::Text(t) if t.contains("ENCRYPTED"))));
    }

    #[test]
    fn rk_integer_and_double_forms() {
        assert_eq!(decode_rk(0x0000_0006), 1.0); // 1<<2 | fInt
        assert_eq!(decode_rk(0x4000_0000), 2.0); // top half of 2.0f64
        assert_eq!(decode_rk(0x4093_4A00), 1234.5);
        assert_eq!(decode_rk(0xC01C_0000), -7.0);
        // 12345 with fX100 flag -> 123.45
        assert_eq!(decode_rk((12345 << 2) | 3), 123.45);
    }

    #[test]
    fn unicode_string_roundtrip() {
        // "AB" as UTF-16LE
        let mut d = vec![2, 0, 0x01, b'A', 0, b'B', 0];
        assert_eq!(parse_unicode_string(&d), "AB");
        // "AB" compressed
        d = vec![2, 0, 0x00, b'A', b'B'];
        assert_eq!(parse_unicode_string(&d), "AB");
    }

    /// `"B12"` → `(11, 1)`.
    fn addr(a: &str) -> (u16, u16) {
        let split = a.find(|c: char| c.is_ascii_digit()).unwrap();
        let (letters, digits) = a.split_at(split);
        let mut col = 0u16;
        for c in letters.chars() {
            col = col * 26 + (c.to_ascii_uppercase() as u16 - b'A' as u16 + 1);
        }
        (digits.parse::<u16>().unwrap() - 1, col - 1)
    }

    fn sample_file(name: &str) -> Option<Workbook> {
        let p = format!("{}/../../samples/{name}", env!("CARGO_MANIFEST_DIR"));
        if !std::path::Path::new(&p).exists() {
            return None;
        }
        Some(open(&p).unwrap_or_else(|e| panic!("parse {name}: {e}")))
    }

    #[test]
    fn decompiles_real_formulas_from_the_sample() {
        let Some(wb) = sample_file("formulas.xls") else { return };
        let f = wb.sheet("F").expect("sheet F");
        // The exact text WPS/Excel put in each cell when the file was written.
        let want = [
            ("B1", "A1+A2"),
            ("B2", "A1*A2"),
            ("B3", "SUM(A1:A2)"),
            ("B4", "A1/A2"),
            ("B5", "A1^2"),
            ("B6", "IF(A1>5,\"big\",\"small\")"),
            ("B7", "AVERAGE(A1:A2)"),
            ("B8", "MAX(A1:A2)"),
            ("B9", "-A1"),
            ("B10", "A1&\"x\""),
            ("B11", "A1>5"),
            ("B12", "ROUND(A3,1)"),
            ("B13", "SUM(A1:A2)*2"),
            ("B14", "ABS(A1-A2)"),
            ("B15", "A1<>A2"),
            ("B16", "SUM(A1,A2)"),
            ("C1", "G!A1"),
            ("C2", "SUM(G!A1,3)"),
        ];
        for (a, text) in want {
            let (row, col) = addr(a);
            let cell = f.cell(row, col).unwrap_or_else(|| panic!("{a} missing"));
            assert_eq!(wb.formula_text(cell), Some(text), "formula in {a}");
        }
    }

    #[test]
    fn applies_number_formats_from_the_sample() {
        let Some(wb) = sample_file("formulas.xls") else { return };
        let f = wb.sheet("F").expect("sheet F");
        let show = |a: &str| {
            let (row, col) = addr(a);
            wb.cell_text(f.cell(row, col).expect(a))
        };
        assert_eq!(show("A3"), "1234.57");          // 0.00
        assert_eq!(show("A4"), "25.00%");           // 0.00%
        assert_eq!(show("A5"), "2024-03-05");       // yyyy-mm-dd
        assert_eq!(show("A6"), "$45,356.00");       // $#,##0.00
        assert_eq!(show("A1"), "10");               // General
        // The custom formats really came out of FORMAT records.
        assert!(wb.formats.values().any(|c| c.contains("yyyy")));
    }

    fn sample() -> Option<Workbook> {
        let p = concat!(env!("CARGO_MANIFEST_DIR"), "/../../samples/hello.xls");
        if !std::path::Path::new(p).exists() {
            return None;
        }
        Some(open(p).expect("parse hello.xls"))
    }

    #[test]
    fn reads_hello_xls() {
        let Some(wb) = sample() else { return };
        assert_eq!(wb.sheets.len(), 2);
        // BOUNDSHEET order is tab order; this file lists Sheet2 first.
        let data = wb.sheet("Data").expect("Data sheet");
        let a1 = data.cell(0, 0).unwrap();
        assert_eq!(a1.value, CellValue::Text("HELLO-XLS-001".into()));
        let a3 = data.cell(2, 0).unwrap();
        assert_eq!(a3.value, CellValue::Number(1234.5));
        let b3 = data.cell(2, 1).unwrap();
        match &b3.value {
            CellValue::Formula { cached, .. } => assert_eq!(*cached, CachedValue::Number(3.0)),
            other => panic!("expected formula, got {other:?}"),
        }
        assert!(wb.sst.iter().any(|s| s == "中文单元格"));
    }
}
