//! # office-doc — Word 97-2003 binary format (`.doc`)
//!
//! A `.doc` file is a CFB container holding a `WordDocument` stream plus a
//! table stream (`0Table` or `1Table`, selected by FIB flag `fWhichTblStm`).
//! The `WordDocument` stream starts with the **FIB** (File Information Block),
//! which points at the **CLX** (a piece table) inside the table stream. Text
//! is not stored contiguously: the piece table maps character positions (CPs)
//! to file offsets, and each piece is either compressed (one byte per char,
//! CP1252) or UTF-16LE.
//!
//! ```text
//! WordDocument stream                    1Table stream
//! +----------------------+               +--------------------+
//! | FIB (>= 0x01A2 bytes)|--fcClx/lcbClx->| CLX: Prc*, Pcdt    |
//! | ...                  |               |  PlcPcd: CP[] PCD[]|
//! | text bytes           |<--fc of PCD---|                    |
//! +----------------------+               +--------------------+
//! ```
//!
//! Verified against Word 97+ files and WPS Office output; see `docs/doc.md`.

use office_cfb::{Cfb, Error as CfbError};

pub type Result<T> = std::result::Result<T, Error>;

#[derive(Debug)]
pub enum Error {
    Cfb(CfbError),
    NotWordDocument { w_ident: u16 },
    UnsupportedFib(u16),
    Truncated(&'static str),
    BadClx(&'static str),
    /// The file is password protected and no password was supplied.
    Encrypted,
    /// A password was supplied but the stored verifier rejected it.
    WrongPassword,
    Crypto(office_crypto::Error),
}

impl std::fmt::Display for Error {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Error::Cfb(e) => write!(f, "cfb: {e}"),
            Error::NotWordDocument { w_ident } => {
                write!(f, "not a Word document: wIdent=0x{w_ident:04X} (expected 0xA5EC)")
            }
            Error::UnsupportedFib(n) => write!(f, "unsupported Word nFib 0x{n:04X}"),
            Error::Truncated(what) => write!(f, "truncated {what}"),
            Error::BadClx(why) => write!(f, "bad CLX/piece table: {why}"),
            Error::Encrypted => write!(f, "document is encrypted; a password is required"),
            Error::WrongPassword => write!(f, "wrong password"),
            Error::Crypto(e) => write!(f, "crypto: {e}"),
        }
    }
}

impl std::error::Error for Error {}

impl From<office_crypto::Error> for Error {
    fn from(e: office_crypto::Error) -> Self {
        Error::Crypto(e)
    }
}

impl From<CfbError> for Error {
    fn from(e: CfbError) -> Self {
        Error::Cfb(e)
    }
}

/// `wIdent` of every Word binary document.
pub const W_IDENT: u16 = 0xA5EC;

pub mod fmt;
pub mod write;

pub use write::{
    write_document, write_minimal, write_table_document, Paragraph, Run, Table,
};

/// Word's paragraph mark, the separator between paragraphs in `Doc::text`.
pub const PARA_MARK: char = '\u{000D}';

/// Word's table cell mark: it ends a cell exactly like a paragraph mark ends a
/// paragraph, and a row ends with one extra mark whose properties carry
/// `sprmPFTtp`.
pub const CELL_MARK: char = '\u{0007}';

/// Parsed File Information Block (the fields this crate needs).
#[derive(Clone, Debug)]
pub struct Fib {
    pub w_ident: u16,
    pub n_fib: u16,
    /// Language id (`lid`), e.g. 0x0804 = Chinese (PRC).
    pub lid: u16,
    pub flags: u16,
    pub fc_min: u32,
    pub fc_mac: u32,
    /// Number of characters in the main document text.
    pub ccp_text: u32,
    /// `"1Table"` when `fWhichTblStm` (flags bit 9) is set, else `"0Table"`.
    pub table_stream: &'static str,
    pub fc_clx: u32,
    pub lcb_clx: u32,
    /// `fibRgLw97` character counts, in text-stream order: main text,
    /// footnotes, headers, macros, annotations, endnotes, textboxes, header
    /// textboxes.
    pub ccp_counts: [u32; 8],
}

/// One piece of the piece table.
#[derive(Clone, Debug)]
pub struct Piece {
    pub cp_start: u32,
    pub cp_end: u32,
    /// Raw `fc` from the PCD (bit 30 = `fCompressed`).
    pub fc: u32,
    pub compressed: bool,
    pub text: String,
}

/// The page a section is printed on.
///
/// Every length is in twips (1/20 pt), the unit `SEPX` uses. The writer emits
/// `sprmSXaPage`, `sprmSYaPage`, `sprmSDxaLeft`, `sprmSDxaRight` and, when the
/// margins are given, `sprmSDyaTop` / `sprmSDyaBottom`; Word's own files also
/// carry a handful of printer defaults that a document reads fine without.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct PageSetup {
    /// Page width in twips (A4 is 11906).
    pub width: u16,
    /// Page height in twips (A4 is 16838).
    pub height: u16,
    pub left: i16,
    pub right: i16,
    /// Top margin, or `None` to leave Word's default (1440 twips).
    pub top: Option<i16>,
    /// Bottom margin, or `None` to leave Word's default (1440 twips).
    pub bottom: Option<i16>,
}

impl PageSetup {
    /// A4 portrait with 1800-twip (1.25") left and right margins.
    pub fn a4() -> Self {
        PageSetup { width: 11906, height: 16838, left: 1800, right: 1800, top: None, bottom: None }
    }

    /// US Letter portrait with 1800-twip left and right margins.
    pub fn letter() -> Self {
        PageSetup { width: 12240, height: 15840, left: 1800, right: 1800, top: None, bottom: None }
    }

    pub fn margins(mut self, left: i16, right: i16, top: i16, bottom: i16) -> Self {
        self.left = left;
        self.right = right;
        self.top = Some(top);
        self.bottom = Some(bottom);
        self
    }

    /// Swap width and height: Word decides landscape from the page shape.
    pub fn landscape(mut self) -> Self {
        std::mem::swap(&mut self.width, &mut self.height);
        self
    }
}
/// A parsed `.doc` file.
#[derive(Clone, Debug)]
pub struct Doc {
    pub fib: Fib,
    pub pieces: Vec<Piece>,
    /// Main document text, exactly `fib.ccp_text` characters.
    pub text: String,
    /// The whole text stream: main text followed by the footnote, header,
    /// annotation, endnote and textbox stories (see [`Story`]).
    pub full_text: String,
    /// Character formatting runs, in stream order (see [`fmt`]).
    pub char_runs: Vec<fmt::CharRun>,
    /// Paragraph formatting runs, in stream order (see [`fmt`]).
    pub para_runs: Vec<fmt::ParaRun>,
    /// The encryption scheme, when the document was password protected.
    pub encryption: Option<office_crypto::Scheme>,
    /// The `Data` stream: pictures (`PICF` + OfficeArt) and form fields.
    pub data: Vec<u8>,
    /// The style sheet indexed by `istd` (see [`Doc::style_name`]); an entry
    /// with no name is a style whose definition the file never wrote.
    pub styles: Vec<fmt::StyleDef>,
    /// `PlcfHdd` cut points, relative to the start of the header story: six
    /// document-wide separators, then six slots per section (even-page header,
    /// primary header, even-page footer, primary footer, first-page header,
    /// first-page footer), then a closing slot. Empty when the file has no
    /// header story.
    pub hdd_slots: Vec<u32>,
    /// The page the first section is printed on, from its `SEPX`. `None` when
    /// the file carries no section properties.
    pub page: Option<PageSetup>,
    /// Every section's page setup, in document order, from `Plcfsed`. A section
    /// that inherits its properties (`fcSepx` is `0xFFFF_0000`) repeats the one
    /// before it.
    pub sections: Vec<PageSetup>,
    /// `PlcffndRef` cut points: where each footnote's reference mark sits in the
    /// main text (`0x02`, one character each), followed by one CP past every
    /// story -- the same value `Plcfsed` closes on. Empty when the file has no
    /// footnotes.
    pub footnote_slots: Vec<u32>,
    /// One `FRD` word per footnote, from `PlcffndRef`: bit 0 is `fAuto`, set for
    /// an auto-numbered note and clear for a custom mark. Empty when the file
    /// has no footnotes.
    pub footnote_auto: Vec<bool>,
    /// `PlcffndTxt` cut points, relative to the start of the footnote story:
    /// where each note's own text starts, followed by a closing CP three past
    /// the last note. Empty when the file has no footnotes.
    pub footnote_text_slots: Vec<u32>,
    /// `PlcfendRef` cut points: the same convention as [`Doc::footnote_slots`],
    /// for the endnote marks in the main text. Empty when the file has no
    /// endnotes.
    pub endnote_slots: Vec<u32>,
    /// `PlcfendTxt` cut points, relative to the start of the endnote story.
    /// Empty when the file has no endnotes.
    pub endnote_text_slots: Vec<u32>,
    /// The list definitions from `PlfLst` (FIB pair 73), in `ilfo`-addressable
    /// order after [`Doc::list_overrides`]. Empty when the file has no lists.
    pub lists: Vec<ListDef>,
    /// The `lsid` of each override in `PlfLfo` (FIB pair 74); a paragraph's
    /// `sprmPIlfo` indexes this one-based.
    pub list_overrides: Vec<i32>,
    /// One entry per annotation (Word comment), in document order, from
    /// `PlcfandRef`/`PlcfandTxt` (FIB pairs 4 and 5). Empty when the file has
    /// no comments.
    pub comments: Vec<DocComment>,
}

/// One Word comment: where it anchors in the main text, who wrote it, and what
/// it says.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DocComment {
    /// Character offset of the comment's reference mark in the main text (the
    /// `U+0005` character Word stores there). `PlcfandRef` records this one CP
    /// per comment; the commented text sits just before it.
    pub anchor: u32,
    /// The author's initials, from the comment's `ATRDPre10`; empty when none.
    pub initials: String,
    /// The comment text, without the leading `U+0005` mark Word stores.
    pub text: String,
}

/// One picture stored in the `Data` stream.
#[derive(Debug, Clone)]
pub struct DocPicture {
    /// Offset of the `PICF` structure inside the `Data` stream
    /// (`sprmCPicLocation`).
    pub fc: u32,
    pub kind: office_blip::PictureKind,
    /// The image bytes exactly as stored (a complete PNG/JPEG/… file).
    pub data: Vec<u8>,
}

/// Open and parse a `.doc` file from disk.
pub fn open<P: AsRef<std::path::Path>>(path: P) -> Result<Doc> {
    let bytes = std::fs::read(path).map_err(CfbError::Io)?;
    parse(&bytes)
}

/// Open and parse a password protected `.doc` file from disk.
pub fn open_with_password<P: AsRef<std::path::Path>>(path: P, password: &str) -> Result<Doc> {
    let bytes = std::fs::read(path).map_err(CfbError::Io)?;
    parse_with_password(&bytes, password)
}

/// Parse a `.doc` file from its raw bytes (the CFB container).
pub fn parse(bytes: &[u8]) -> Result<Doc> {
    parse_impl(bytes, None)
}

/// Parse a password protected `.doc` file from its raw bytes.
pub fn parse_with_password(bytes: &[u8], password: &str) -> Result<Doc> {
    parse_impl(bytes, Some(password))
}

/// Bytes of the `WordDocument` stream that are never encrypted (FibBase plus
/// the first part of `fibRgW97`).
const FIB_PLAIN_LEN: usize = 0x44;

/// `FibBase.flags` bit 8: the document is protected.
const FLAG_ENCRYPTED: u16 = 0x0100;
/// `FibBase.flags` bit 15: the protection is XOR obfuscation, not RC4.
/// `FibBase.lKey` (u32 at 0x0E) then holds the XOR password verifier in its low
/// 16 bits (MS-DOC 2.5.1).
const FLAG_OBFUSCATED: u16 = 0x8000;
/// Offset of `FibBase.lKey`.
const FIB_LKEY: usize = 0x0E;

/// Undo Word's XOR obfuscation on one stream.
///
/// Unlike RC4, obfuscation has no blocks: every byte is `ror8(byte ^ array[i], 5)`
/// with `i` cycling over 16 positions anchored at the *end* of the stream, so
/// the stream must be passed whole (position 0..len). The first
/// `plain_prefix` bytes were written back in the clear and are restored.
pub fn deobfuscate_stream(data: &[u8], password: &str, plain_prefix: usize) -> Vec<u8> {
    let arr = office_crypto::xor::array(password);
    let mut out = office_crypto::xor::decrypt_run(data, &arr, 0, false);
    let keep = plain_prefix.min(out.len()).min(data.len());
    out[..keep].copy_from_slice(&data[..keep]);
    out
}

/// Decrypt a Word stream.
///
/// Word encrypts the *whole* stream in 512-byte RC4 blocks numbered from offset
/// 0, then writes a few leading bytes back in plaintext (`FIB_PLAIN_LEN` for
/// `WordDocument`, the encryption header for the table stream). Those bytes are
/// restored here, which is why the prefix length is a parameter.
fn decrypt_stream(
    data: &[u8],
    scheme: &office_crypto::Scheme,
    password: &str,
    plain_prefix: usize,
) -> Result<Vec<u8>> {
    let mut out = office_crypto::decrypt(scheme, password, data, 512)?;
    let keep = plain_prefix.min(out.len()).min(data.len());
    out[..keep].copy_from_slice(&data[..keep]);
    Ok(out)
}

fn parse_impl(bytes: &[u8], password: Option<&str>) -> Result<Doc> {
    let cfb = Cfb::parse(bytes.to_vec())?;
    let wd_raw = cfb.stream("WordDocument")?;
    if wd_raw.len() < 0x20 {
        return Err(Error::Truncated("FIB"));
    }
    let raw_flags = u16::from_le_bytes([wd_raw[10], wd_raw[11]]);
    let table_stream = if raw_flags & 0x0200 != 0 { "1Table" } else { "0Table" };

    // A protected document keeps its FIB in the clear (so the stream names are
    // readable) and sets fEncrypted; the encryption header then sits at the
    // start of the table stream.
    // Two protection schemes live under the same flag: Word 97 XOR obfuscation
    // (`fObfuscated`, verifier in `lKey`, no header anywhere) and RC4/RC4
    // CryptoAPI (verifier in an encryption header at the start of the table
    // stream).
    let (wd, table, encryption) = if raw_flags & FLAG_ENCRYPTED != 0 && raw_flags & FLAG_OBFUSCATED != 0 {
        let password = password.ok_or(Error::Encrypted)?;
        let verifier = u32::from_le_bytes([
            wd_raw[FIB_LKEY],
            wd_raw[FIB_LKEY + 1],
            wd_raw[FIB_LKEY + 2],
            wd_raw[FIB_LKEY + 3],
        ]) as u16;
        if !office_crypto::xor::verify(password, verifier) {
            return Err(Error::WrongPassword);
        }
        let scheme = office_crypto::Scheme::Xor {
            key: office_crypto::xor::key(password),
        };
        // The FIB stays readable; everything after it is obfuscated. The table
        // and `Data` streams are obfuscated from their first byte.
        let wd = deobfuscate_stream(&wd_raw, password, FIB_PLAIN_LEN);
        let table = deobfuscate_stream(&cfb.stream(table_stream)?, password, 0);
        (wd, table, Some(scheme))
    } else if raw_flags & FLAG_ENCRYPTED != 0 {
        let password = password.ok_or(Error::Encrypted)?;
        let table_raw = cfb.stream(table_stream)?;
        let header = office_crypto::parse_header(&table_raw)?;
        if !header.scheme.verify(password) {
            return Err(Error::WrongPassword);
        }
        let wd = decrypt_stream(&wd_raw, &header.scheme, password, FIB_PLAIN_LEN)?;
        let table = decrypt_stream(&table_raw, &header.scheme, password, header.consumed)?;
        (wd, table, Some(header.scheme))
    } else {
        (wd_raw, cfb.stream(table_stream)?, None)
    };

    // The `Data` stream holds pictures and form-field data; it is empty in a
    // document without pictures.
    let data = cfb.stream("Data").unwrap_or_default();

    let u16_at = |o: usize| u16::from_le_bytes([wd[o], wd[o + 1]]);
    let u32_at = |o: usize| u32::from_le_bytes([wd[o], wd[o + 1], wd[o + 2], wd[o + 3]]);

    let w_ident = u16_at(0);
    if w_ident != W_IDENT {
        return Err(Error::NotWordDocument { w_ident });
    }
    let n_fib = u16_at(2);
    let lid = u16_at(6);
    let flags = u16_at(10);
    let fc_min = u32_at(24);
    let fc_mac = u32_at(28);
    let ccp_text = u32_at(76);

    // Word 6/95 (nFib <= 0x0068) has no piece table: text is the raw byte run
    // [fcMin, fcMac) in the WordDocument stream.
    if n_fib < 0x00C1 {
        if fc_min as usize > wd.len() || fc_mac as usize > wd.len() || fc_min > fc_mac {
            return Err(Error::Truncated("Word 6/95 text range"));
        }
        let text = decode_cp1252(&wd[fc_min as usize..fc_mac as usize]);
        let fib = Fib {
            w_ident,
            n_fib,
            lid,
            flags,
            fc_min,
            fc_mac,
            ccp_text,
            table_stream,
            fc_clx: 0,
            lcb_clx: 0,
            ccp_counts: [text.chars().count() as u32, 0, 0, 0, 0, 0, 0, 0],
        };
        return Ok(Doc {
            fib,
            pieces: vec![Piece {
                cp_start: 0,
                cp_end: text.chars().count() as u32,
                fc: fc_min,
                compressed: true,
                text: text.clone(),
            }],
            full_text: text.clone(),
            text,
            page: None,
            sections: Vec::new(),
            footnote_auto: Vec::new(),
            char_runs: Vec::new(),
            para_runs: Vec::new(),
            encryption: encryption.clone(),
            data,
            styles: Vec::new(),
            hdd_slots: Vec::new(),
            footnote_slots: Vec::new(),
            footnote_text_slots: Vec::new(),
            endnote_slots: Vec::new(),
            endnote_text_slots: Vec::new(),
            lists: Vec::new(),
            list_overrides: Vec::new(),
            comments: Vec::new(),
        });
    }

    if wd.len() < 0x01AA {
        return Err(Error::Truncated("FibRgFcLcb97"));
    }
    let fc_clx = u32_at(0x01A2);
    let lcb_clx = u32_at(0x01A6);
    let fib = Fib {
        w_ident,
        n_fib,
        lid,
        flags,
        fc_min,
        fc_mac,
        ccp_text,
        table_stream,
        fc_clx,
        lcb_clx,
        ccp_counts: std::array::from_fn(|i| u32_at(0x4C + i * 4)),
    };

    let clx = table
        .get(fc_clx as usize..fc_clx as usize + lcb_clx as usize)
        .ok_or(Error::Truncated("CLX range in table stream"))?;

    let pieces = parse_clx(clx, &wd)?;
    let mut text = String::new();
    for p in &pieces {
        if text.chars().count() as u32 >= ccp_text {
            break;
        }
        let take = (ccp_text - text.chars().count() as u32) as usize;
        let slice: String = p.text.chars().take(take).collect();
        text.push_str(&slice);
    }
    let full_text: String = pieces.iter().map(|p| p.text.as_str()).collect();
    let char_runs = formatting_pages(&wd, &table, 12, fmt::chpx_page);
    let para_runs = formatting_pages(&wd, &table, 13, fmt::papx_page);
    // fibRgFcLcb97 pair 1 is `fcStshf`/`lcbStshf`.
    let styles = parse_stsh(&table, u32_at(0x9A + 8), u32_at(0x9A + 12));
    // Pair 6 is `fcPlcfsed`/`lcbPlcfsed`: a `Plc` of `n + 1` CPs followed by
    // `n` twelve-byte `SED`s, each `fn` (2), `fcSepx` (4), `fnMpr` (2),
    // `fcMpr` (4). `fcSepx` points into the WordDocument stream, not the table.
    let sections = section_pages(&wd, &table, u32_at(0x9A + 6 * 8), u32_at(0x9A + 6 * 8 + 4));
    let page = sections.first().copied();
    // Pair 11 is `fcPlcfhdd`/`lcbPlcfhdd`: bare CPs, no trailing data.
    let hdd_slots = plcfhdd(&table, u32_at(0x9A + 11 * 8), u32_at(0x9A + 11 * 8 + 4));
    // Pair 2 is `fcPlcffndRef`/`lcbPlcffndRef`, pair 3 `fcPlcffndTxt`/`lcb`.
    // Both carry CP arrays with no trailing data, like `PlcfHdd`.
    let footnote_slots = plcfndref(&table, u32_at(0x9A + 2 * 8), u32_at(0x9A + 2 * 8 + 4));
    let footnote_auto: Vec<bool> = plcfndref_frd(
        &table,
        u32_at(0x9A + 2 * 8),
        u32_at(0x9A + 2 * 8 + 4),
    )
    .into_iter()
    .map(|frd| frd & 1 != 0)
    .collect();
    let footnote_text_slots = plcfhdd(&table, u32_at(0x9A + 3 * 8), u32_at(0x9A + 3 * 8 + 4));
    // Pairs 4 and 5 are `fcPlcfendRef`/`fcPlcfendTxt`: the footnote pair again,
    // this time cutting the endnote story.
    // Pairs 46 and 47 hold the endnote cut points; pairs 4 and 5 are the
    // annotation PLCs and reading those made the endnote story come out empty.
    let endnote_slots = plcfndref(&table, u32_at(0x9A + 46 * 8), u32_at(0x9A + 46 * 8 + 4));
    let endnote_text_slots = plcfhdd(&table, u32_at(0x9A + 47 * 8), u32_at(0x9A + 47 * 8 + 4));
    // Pairs 73 and 74 are `PlfLst` (the definitions, with the `LVL` array right
    // after the PLC) and `PlfLfo` (the override table that `sprmPIlfo` indexes).
    let lists = parse_plflst(&table, u32_at(0x9A + 73 * 8), u32_at(0x9A + 73 * 8 + 4));
    let list_overrides = parse_plflfo(&table, u32_at(0x9A + 74 * 8), u32_at(0x9A + 74 * 8 + 4));
    // Pairs 4 and 5 are `PlcfandRef`/`PlcfandTxt`: where each Word comment
    // anchors and the story that holds its text. `PlcfandTxt` cut points are
    // relative to the annotation story, so pass that slice on its own.
    let ann_start = (fib.ccp_counts[0] + fib.ccp_counts[1] + fib.ccp_counts[2] + fib.ccp_counts[3])
        as usize;
    let ann_len = fib.ccp_counts[4] as usize;
    let ann_story = full_text.get(ann_start..ann_start + ann_len).unwrap_or("");
    let comments = parse_comments(
        &table,
        u32_at(0x9A + 4 * 8),
        u32_at(0x9A + 4 * 8 + 4),
        u32_at(0x9A + 5 * 8),
        u32_at(0x9A + 5 * 8 + 4),
        ann_story,
    );

    Ok(Doc {
        fib,
        pieces,
        text,
        full_text,
        char_runs,
        para_runs,
        encryption,
        data,
        styles,
        hdd_slots,
        page,
        sections,
        footnote_slots,
        footnote_auto,
        footnote_text_slots,
        endnote_slots,
        endnote_text_slots,
        lists,
        list_overrides,
        comments,
    })
}

/// A bare CP array: `PlcfHdd` and `PlcffndTxt` are just lists of cut points.
/// Read every section's `SEPX` and pull the page setups out of them.
///
/// A section with `fcSepx = 0xFFFF_0000` has no properties of its own and
/// inherits the previous one, which is how Word stores a section the user never
/// changed (`samples/ms/hdd-twosections.doc`).
fn section_pages(wd: &[u8], table: &[u8], fc: u32, lcb: u32) -> Vec<PageSetup> {
    let fc = fc as usize;
    let lcb = lcb as usize;
    if lcb < 16 || fc + lcb > table.len() {
        return Vec::new();
    }
    let sections = (lcb - 4) / 12;
    let mut pages: Vec<PageSetup> = Vec::new();
    for i in 0..sections {
        let sed = fc + (sections + 1) * 4 + i * 12;
        if sed + 12 > table.len() {
            break;
        }
        let fc_sepx = u32_at_in(table, sed + 2) as usize;
        if fc_sepx == 0xFFFF_FFFF || fc_sepx == 0xFFFF_0000 || fc_sepx + 2 > wd.len() {
            if let Some(previous) = pages.last().copied() {
                pages.push(previous);
            }
            continue;
        }
        let cb = u16::from_le_bytes([wd[fc_sepx], wd[fc_sepx + 1]]) as usize;
        let body = &wd[fc_sepx + 2..(fc_sepx + 2 + cb).min(wd.len())];
        if let Some(page) = sepx_page(body) {
            pages.push(page);
        }
    }
    pages
}

/// Pull the page setup out of one `SEPX` body.
///
/// Word writes far more sprms than these (`sprmSFpc`, `sprmSBkc`, printer
/// defaults, ...); the five below are the ones `PageSetup` reports.
fn sepx_page(body: &[u8]) -> Option<PageSetup> {
    let mut page = PageSetup { width: 0, height: 0, left: 0, right: 0, top: None, bottom: None };
    let mut seen = false;
    let mut at = 0;
    while at + 2 <= body.len() {
        let sprm = u16::from_le_bytes([body[at], body[at + 1]]);
        let fixed = match (sprm >> 13) & 0x7 {
            0 | 1 => 1,
            2 | 4 | 5 => 2,
            3 => 4,
            6 => 0,
            _ => 3,
        };
        let (size, start) = if fixed == 0 {
            (body.get(at + 2).copied().unwrap_or(0) as usize, at + 3)
        } else {
            (fixed, at + 2)
        };
        if start + size > body.len() {
            break;
        }
        let operand = &body[start..start + size];
        let word = || u16::from_le_bytes([operand[0], operand[1]]);
        match sprm {
            0xB01F => {
                page.width = word();
                seen = true;
            }
            0xB020 => page.height = word(),
            0xB021 => page.left = word() as i16,
            0xB022 => page.right = word() as i16,
            0x9023 => page.top = Some(word() as i16),
            0x9024 => page.bottom = Some(word() as i16),
            _ => {}
        }
        at = start + size;
    }
    seen.then_some(page)
}

fn u32_at_in(data: &[u8], at: usize) -> u32 {
    u32::from_le_bytes([data[at], data[at + 1], data[at + 2], data[at + 3]])
}

fn plcfhdd(table: &[u8], fc: u32, lcb: u32) -> Vec<u32> {
    if lcb < 8 {
        return Vec::new();
    }
    let Some(plc) = table.get(fc as usize..(fc + lcb) as usize) else {
        return Vec::new();
    };
    (0..plc.len() / 4)
        .map(|i| u32::from_le_bytes([plc[i * 4], plc[i * 4 + 1], plc[i * 4 + 2], plc[i * 4 + 3]]))
        .collect()
}

impl Doc {
    /// The style name for an `istd` (a paragraph style index from `PAPX`, or a
    /// character style index from `sprmCIstd`).
    pub fn style_name(&self, istd: u16) -> Option<&str> {
        self.styles
            .get(istd as usize)
            .and_then(|s| s.name.as_deref())
    }

    /// The properties a style itself carries (`STD.grupe`), if it was defined.
    pub fn style_props(&self, istd: u16) -> Option<&fmt::StyleDef> {
        self.styles.get(istd as usize)
    }

    /// Every picture reachable from a `sprmCPicLocation` run.
    ///
    /// Word stores pictures in the `Data` stream, not in `WordDocument`: the
    /// character properties of the picture character (0x01) carry the offset of
    /// a `PICF` structure, whose 68-byte header is followed by OfficeArt
    /// records — an `OfficeArtSpContainer` with the shape and an
    /// `OfficeArtFBSE` (0xF007) holding the BLIP.
    pub fn pictures(&self) -> Vec<DocPicture> {
        let mut out = Vec::new();
        let mut seen: Vec<u32> = Vec::new();
        for run in &self.char_runs {
            let Some(fc) = run.props.pic_location else {
                continue;
            };
            if seen.contains(&fc) {
                continue;
            }
            seen.push(fc);
            if let Some(pic) = self.picture_at(fc) {
                out.push(pic);
            }
        }
        out
    }

    fn picture_at(&self, fc: u32) -> Option<DocPicture> {
        let data = &self.data;
        let off = fc as usize;
        let lcb = u32::from_le_bytes(data.get(off..off + 4)?.try_into().ok()?) as usize;
        let cb_header = u16::from_le_bytes(data.get(off + 4..off + 6)?.try_into().ok()?) as usize;
        // PICF: lcb (4) cbHeader (2) mfHeader (2) bm_rcWinMF (8) dxaGoal (2)
        // dyaGoal (2) mx (2) my (2) dxaCropLeft … (8×2) brcl (2) …
        let start = off + cb_header.max(0x44);
        let end = (off + lcb).min(data.len());
        if start >= end {
            return None;
        }
        let blip = find_blip(data, start, end)?;
        Some(DocPicture {
            fc,
            kind: blip.kind,
            data: blip.data.to_vec(),
        })
    }
}

/// Depth-first search for a BLIP in an OfficeArt record range.
fn find_blip(data: &[u8], start: usize, end: usize) -> Option<office_blip::Blip<'_>> {
    let mut pos = start;
    while pos + 8 <= end {
        let ver_inst = u16::from_le_bytes([data[pos], data[pos + 1]]);
        let kind = u16::from_le_bytes([data[pos + 2], data[pos + 3]]);
        let len = u32::from_le_bytes([
            data[pos + 4],
            data[pos + 5],
            data[pos + 6],
            data[pos + 7],
        ]) as usize;
        let body = data.get(pos + 8..pos + 8 + len).unwrap_or(&data[pos + 8..]);
        if kind == office_blip::RT_FBSE {
            if let Some(blip) = office_blip::from_bse(body) {
                return Some(blip);
            }
        } else if office_blip::is_blip(kind) {
            if let Some(blip) = office_blip::from_blip(body) {
                return Some(blip);
            }
        } else if ver_inst & 0x000F == 0x000F {
            // 0xF ⇒ container (0xF000..0xF006): recurse into it.
            if let Some(blip) = find_blip(data, pos + 8, (pos + 8 + len).min(end)) {
                return Some(blip);
            }
        }
        pos += 8 + len;
    }
    None
}

/// Follow one `PlcfBte*` pair (`PlcfBteChpx` = pair 12, `PlcfBtePapx` = 13)
/// to its 512-byte FKP pages and decode each of them.
fn formatting_pages<T>(
    word: &[u8],
    table: &[u8],
    pair_index: usize,
    decode: fn(&[u8]) -> Vec<T>,
) -> Vec<T> {
    let o = 0x9A + pair_index * 8;
    let (Some(fc), Some(lcb)) = (
        word.get(o..o + 4).map(|b| u32::from_le_bytes([b[0], b[1], b[2], b[3]])),
        word.get(o + 4..o + 8).map(|b| u32::from_le_bytes([b[0], b[1], b[2], b[3]])),
    ) else {
        return Vec::new();
    };
    if lcb == 0 {
        return Vec::new();
    }
    let Some(plc) = table.get(fc as usize..(fc + lcb) as usize) else {
        return Vec::new();
    };
    let mut out = Vec::new();
    for pn in fmt::bte_pages(plc) {
        let start = pn as usize * 512;
        if let Some(page) = word.get(start..start + 512) {
            out.extend(decode(page));
        }
    }
    out
}

/// Parse the style sheet (`STSH`) out of the table stream.
///
/// The layout was measured on real files, and WPS and Microsoft agree on it:
/// `cbStshi u16`, `cstd u16`, `cbSTDBaseInFile u16`, … , then `cstd` `STD`
/// records starting at `2 + cbStshi`, each a `cbStd u16` length followed by
/// that many bytes. The style name (`Xstz`) sits at `STD + cbSTDBaseInFile`:
/// a `cch u16` count followed by UTF-16 characters.
///
/// `cbStd == 0` means the style exists in the index space but its definition
/// was not written — Word omits built-in styles that are still at their
/// defaults — so it gets no name.
///
/// Note the leading `cbStshi`: it is not in MS-DOC 2.9.272's `STSHI`, but both
/// `PROTTPLN.DOC` (1118) and the WPS samples (1062) carry it, and the `STD`
/// walk only lands exactly on the end of the block when it is honoured.
fn parse_stsh(table: &[u8], fc: u32, lcb: u32) -> Vec<fmt::StyleDef> {
    let Some(block) = table.get(fc as usize..(fc + lcb) as usize) else {
        return Vec::new();
    };
    let u16_at = |o: usize| -> Option<u16> {
        block
            .get(o..o + 2)
            .map(|b| u16::from_le_bytes([b[0], b[1]]))
    };
    let (Some(cb_stshi), Some(cstd), Some(cb_base)) = (u16_at(0), u16_at(2), u16_at(4)) else {
        return Vec::new();
    };
    let mut out = Vec::with_capacity(cstd as usize);
    let mut pos = 2 + cb_stshi as usize;
    for _ in 0..cstd {
        let Some(cb_std) = u16_at(pos) else { break };
        let body = pos + 2;
        let mut style = fmt::StyleDef::default();
        if cb_std > 0 {
            let at = body + cb_base as usize;
            let limit = (body + cb_std as usize).min(block.len());
            let mut rest = limit;
            if let Some(cch) = u16_at(at) {
                if let Some(chars) = block.get(at + 2..limit) {
                    let take = (cch as usize * 2).min(chars.len());
                    let name = decode_utf16le(&chars[..take]);
                    if !name.is_empty() {
                        style.name = Some(name);
                    }
                    // The name is NUL-terminated, then the `grupe` array
                    // follows: `cbUPX u16` + that many bytes of `grpprl`.
                    rest = (at + 2 + take + 2).min(limit);
                }
            }
            if rest < limit {
                for upx in fmt::style_upx(&block[rest..limit]) {
                    // The array holds a paragraph block and a character block;
                    // feeding both to both walkers is safe because each only
                    // reacts to the `sprm` ids it knows.
                    fmt::apply_para_sprms(&mut style.para, upx);
                    fmt::apply_char_sprms(&mut style.char, upx);
                }
            }
        }
        out.push(style);
        pos = body + cb_std as usize;
    }
    out
}

/// Walk the CLX: zero or more `Prc` (0x01) records followed by one `Pcdt`
/// (0x02) that carries the `PlcPcd` (CP array + PCD array).
fn parse_clx(clx: &[u8], word: &[u8]) -> Result<Vec<Piece>> {
    let mut i = 0usize;
    loop {
        let kind = *clx.get(i).ok_or(Error::BadClx("empty CLX"))?;
        match kind {
            0x01 => {
                // Prc: cbGrpprl (u16) + grpprl — property modifiers, skipped
                let cb = u16::from_le_bytes([
                    *clx.get(i + 1).ok_or(Error::BadClx("Prc length"))?,
                    *clx.get(i + 2).ok_or(Error::BadClx("Prc length"))?,
                ]) as usize;
                i += 3 + cb;
            }
            0x02 => {
                let lcb = u32::from_le_bytes([
                    *clx.get(i + 1).ok_or(Error::BadClx("Pcdt length"))?,
                    *clx.get(i + 2).ok_or(Error::BadClx("Pcdt length"))?,
                    *clx.get(i + 3).ok_or(Error::BadClx("Pcdt length"))?,
                    *clx.get(i + 4).ok_or(Error::BadClx("Pcdt length"))?,
                ]) as usize;
                let plc = clx
                    .get(i + 5..i + 5 + lcb)
                    .ok_or(Error::BadClx("PlcPcd out of range"))?;
                return parse_plcpcd(plc, word);
            }
            other => return Err(Error::BadClx(if other == 0x00 { "zero byte" } else { "unknown CLX tag" })),
        }
    }
}

fn parse_plcpcd(plc: &[u8], word: &[u8]) -> Result<Vec<Piece>> {
    if plc.len() < 4 {
        return Err(Error::BadClx("PlcPcd too short"));
    }
    // A PlcPcd of n pieces is 4*(n+1) bytes of CPs followed by 8*n bytes of
    // PCDs: 4*(n+1) + 8*n = lcb  =>  n = (lcb - 4) / 12.
    if (plc.len() - 4) % 12 != 0 {
        return Err(Error::BadClx("PlcPcd length not of form 4+12n"));
    }
    let n = (plc.len() - 4) / 12;
    let mut cps = Vec::with_capacity(n + 1);
    for k in 0..=n {
        let o = k * 4;
        cps.push(u32::from_le_bytes([plc[o], plc[o + 1], plc[o + 2], plc[o + 3]]));
    }
    let pcd_base = 4 * (n + 1);
    let mut pieces = Vec::with_capacity(n);
    for k in 0..n {
        let pcd = &plc[pcd_base + k * 8..pcd_base + k * 8 + 8];
        // PCD: flags(u16) fc(u32) prm(u16). fc bit 30 = fCompressed.
        let fc_raw = u32::from_le_bytes([pcd[2], pcd[3], pcd[4], pcd[5]]);
        let compressed = fc_raw & 0x4000_0000 != 0;
        let fc = fc_raw & 0x3FFF_FFFF;
        let cp_start = cps[k];
        let cp_end = cps[k + 1];
        let n_chars = cp_end.saturating_sub(cp_start) as usize;
        // Compressed pieces are stored at half the byte offset.
        let byte_off = if compressed { (fc as usize) / 2 } else { fc as usize };
        let byte_len = if compressed { n_chars } else { n_chars * 2 };
        let raw = word
            .get(byte_off..byte_off + byte_len)
            .ok_or(Error::Truncated("piece text in WordDocument"))?;
        let text = if compressed {
            decode_cp1252(raw)
        } else {
            decode_utf16le(raw)
        };
        pieces.push(Piece {
            cp_start,
            cp_end,
            fc: fc_raw,
            compressed,
            text,
        });
    }
    Ok(pieces)
}


/// The parts of a Word document's text stream.
///
/// Word concatenates every "story" into one stream: the main text first, then
/// footnotes, headers and footers, macros, annotations, endnotes and textbox
/// text. The FIB's `fibRgLw97` counts say where each one starts, so no extra
/// structures are needed to split them apart.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Story {
    Main,
    Footnote,
    Header,
    Annotation,
    Endnote,
    Textbox,
    HeaderTextbox,
}

impl Story {
    /// Order of the stories inside the text stream.
    pub const ALL: [Story; 7] = [
        Story::Main,
        Story::Footnote,
        Story::Header,
        Story::Annotation,
        Story::Endnote,
        Story::Textbox,
        Story::HeaderTextbox,
    ];

    pub fn name(&self) -> &'static str {
        match self {
            Story::Main => "main",
            Story::Footnote => "footnote",
            Story::Header => "header",
            Story::Annotation => "annotation",
            Story::Endnote => "endnote",
            Story::Textbox => "textbox",
            Story::HeaderTextbox => "header-textbox",
        }
    }
}

/// `PlcffndRef` cut points. Unlike the bare arrays this PLC trails one 2-byte
/// `FRD` per note after the cut points, so `lcb` is `4k + 2(k-1)` for k cut
/// points: `samples/ms/footnote2.doc` reads `lcb = 10` for `[17, 63]` plus one
/// `FRD`, and `samples/ms/footnote.doc` `lcb = 16` for `[31, 75, 166]` plus two.
/// Reading the whole `lcb` as u32s turns each `FRD` into a bogus cut point.
fn plcfndref(table: &[u8], fc: u32, lcb: u32) -> Vec<u32> {
    let Some(count) = lcb.checked_add(2).map(|n| (n / 6) as usize) else {
        return Vec::new();
    };
    let Some(plc) = table.get(fc as usize..(fc + lcb) as usize) else {
        return Vec::new();
    };
    (0..count)
        .filter_map(|i| plc.get(i * 4..i * 4 + 4))
        .map(|b| u32::from_le_bytes([b[0], b[1], b[2], b[3]]))
        .collect()
}
/// The `FRD` words that `PlcffndRef` stores after its cut points, one per note.
fn plcfndref_frd(table: &[u8], fc: u32, lcb: u32) -> Vec<u16> {
    let Some(count) = lcb.checked_add(2).map(|n| (n / 6) as usize) else {
        return Vec::new();
    };
    if count == 0 {
        return Vec::new();
    }
    let Some(plc) = table.get(fc as usize..(fc + lcb) as usize) else {
        return Vec::new();
    };
    (0..count - 1)
        .filter_map(|i| plc.get(count * 4 + i * 2..count * 4 + i * 2 + 2))
        .map(|b| u16::from_le_bytes([b[0], b[1]]))
        .collect()
}

/// One list definition parsed from `PlfLst`; `levels` are outermost first.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct ListDef {
    pub lsid: i32,
    pub levels: Vec<fmt::ListFormat>,
}

/// Parse `PlfLst` (FIB pair 73) and the `LVL` array that follows it.
///
/// `PlfLst` is `cLst` (u16) then `cLst` 28-byte `LSTF`s; the `LVL`s live right
/// after the PLC (`fc + lcb`), one per level, with a simple list carrying a
/// single level and a multi-level list nine.
fn parse_plflst(table: &[u8], fc: u32, lcb: u32) -> Vec<ListDef> {
    // Word leaves the pair at 0/0 for a document with no lists.
    if lcb == 0 {
        return Vec::new();
    }
    let base = fc as usize;
    let Some(head) = table.get(base..base + 2) else { return Vec::new() };
    let clst = u16::from_le_bytes([head[0], head[1]]) as usize;
    let mut defs = Vec::with_capacity(clst);
    // The LVL array starts right after the PlfLst body (`fc + lcb`).
    let mut cursor = base + lcb as usize;
    for i in 0..clst {
        let at = base + 2 + i * 28;
        let Some(lstf) = table.get(at..at + 28) else { break };
        let lsid = i32::from_le_bytes([lstf[0], lstf[1], lstf[2], lstf[3]]);
        let simple = lstf[26] & 0x01 != 0;
        let count = if simple { 1 } else { 9 };
        let mut levels = Vec::with_capacity(count);
        for lvl in 0..count {
            let Some((def, next)) = parse_lvl(table, cursor, lvl as u8) else { break };
            levels.push(def);
            cursor = next;
        }
        defs.push(ListDef { lsid, levels });
    }
    defs
}

/// One `LVL` record at `at`: `LVLF` (28) + `grpprlPapx` + `grpprlChpx` + `Xst`.
fn parse_lvl(table: &[u8], at: usize, level: u8) -> Option<(fmt::ListFormat, usize)> {
    let lvlf = table.get(at..at + 28)?;
    let start = i32::from_le_bytes([lvlf[0], lvlf[1], lvlf[2], lvlf[3]]);
    let nfc = lvlf[4];
    // `rgbxchNums[9]` at 6..15 marks which template characters are placeholders.
    let placeholders: Vec<usize> = lvlf[6..15]
        .iter()
        .take_while(|b| **b != 0)
        .map(|b| (*b as usize).saturating_sub(1))
        .collect();
    let cb_chpx = lvlf[24] as usize;
    let cb_papx = lvlf[25] as usize;
    let xst_at = at + 28 + cb_papx + cb_chpx;
    let cch = u16::from_le_bytes([*table.get(xst_at)?, *table.get(xst_at + 1)?]) as usize;
    let mut template = String::new();
    for i in 0..cch {
        let off = xst_at + 2 + i * 2;
        let code = u16::from_le_bytes([*table.get(off)?, *table.get(off + 1)?]);
        let ch = char::from_u32(code as u32).unwrap_or('\u{FFFD}');
        // A placeholder becomes `1` so the caller sees `1.` rather than `\u{0}.`.
        template.push(if placeholders.contains(&i) { '1' } else { ch });
    }
    let kind = match nfc {
        0x17 => fmt::ListKind::Bullet,
        0xFF => fmt::ListKind::None,
        _ => fmt::ListKind::Number,
    };
    let end = xst_at + 2 + cch * 2;
    Some(
        (
            fmt::ListFormat { kind, level, text: template, start },
            end,
        ),
    )
}

/// The `lsid` of every list override in `PlfLfo` (FIB pair 74), in `ilfo` order.
fn parse_plflfo(table: &[u8], fc: u32, lcb: u32) -> Vec<i32> {
    // Word leaves the pair at 0/0 for a document with no lists.
    if lcb == 0 {
        return Vec::new();
    }
    let base = fc as usize;
    let Some(head) = table.get(base..base + 4) else { return Vec::new() };
    let mac = u32::from_le_bytes([head[0], head[1], head[2], head[3]]) as usize;
    let mut out = Vec::with_capacity(mac);
    for i in 0..mac {
        let at = base + 4 + i * 16;
        let Some(lfo) = table.get(at..at + 4) else { break };
        out.push(i32::from_le_bytes([lfo[0], lfo[1], lfo[2], lfo[3]]));
    }
    out
}

/// Parse the annotations (Word comments) from `PlcfandRef` (pair 4) and
/// `PlcfandTxt` (pair 5).
///
/// `PlcfandRef` is a `Plc` whose cut points bracket each comment's reference in
/// the main text; after them trail one 30-byte `ATRDPre10` per comment, whose
/// first two bytes are `xstUsrInitl` (the author's initials stored as a
/// `cch`-prefixed UTF-16 string). `PlcfandTxt` slices the annotation story, one
/// `U+0005`-prefixed block per comment. `samples/ms/comments.doc` reads
/// `PlcfandRef lcb = 72` for `[27, 58, 98]` plus two `ATRDPre10`, and
/// `PlcfandTxt lcb = 16` for `[0, 18, 33, 36]`.
fn parse_comments(
    table: &[u8],
    ref_fc: u32,
    ref_lcb: u32,
    txt_fc: u32,
    txt_lcb: u32,
    story: &str,
) -> Vec<DocComment> {
    // Each comment costs 4 bytes of cut-point plus 30 bytes of `ATRDPre10`, so
    // `lcb = 4(n+1) + 30n = 34n + 4`.
    if ref_lcb < 4 {
        return Vec::new();
    }
    let count = ((ref_lcb - 4) / 34) as usize;
    let Some(plc) = table.get(ref_fc as usize..(ref_fc + ref_lcb) as usize) else {
        return Vec::new();
    };
    let cut = |i: usize| -> u32 {
        let at = i * 4;
        u32::from_le_bytes([plc[at], plc[at + 1], plc[at + 2], plc[at + 3]])
    };
    let atrds = &plc[(count + 1) * 4..];
    // `PlcfandTxt` cut points slice the annotation story, two per comment (its
    // own characters, then the paragraph that follows).
    let txt = plcfhdd(table, txt_fc, txt_lcb);

    let mut out = Vec::with_capacity(count);
    for i in 0..count {
        let anchor = cut(i);
        // The initials ride in the `ATRDPre10` at 30 bytes per comment.
        let initials = atrds
            .get(i * 30..i * 30 + 30)
            .map(|a| {
                let cch = u16::from_le_bytes([a[0], a[1]]) as usize;
                let mut s = String::new();
                for k in 0..cch {
                    let at = 2 + k * 2;
                    if let Some(code) = a.get(at..at + 2) {
                        s.push(char::from_u32(u16::from_le_bytes([code[0], code[1]]) as u32).unwrap_or('?'));
                    }
                }
                s
            })
            .unwrap_or_default();
        // The comment's own text lives between consecutive `PlcfandTxt` cut
        // points, one `U+0005` block each; drop the leading mark.
        let text = txt
            .get(i)
            .zip(txt.get(i + 1))
            .and_then(|(&from, &to)| {
                let from = from as usize;
                let to = to as usize;
                story.get(from..to).map(|slice| {
                    slice
                        .trim_start_matches(['\u{5}', '\u{2}'])
                        .trim_end_matches(['\r', '\n'])
                        .to_string()
                })
            })
            .unwrap_or_default();
        out.push(DocComment { anchor, initials, text });
    }
    out
}

impl Doc {
    /// Character range of one story inside [`Doc::full_text`].
    pub fn story_range(&self, story: Story) -> (u32, u32) {
        // fibRgLw97: cbMac, reserved, reserved, ccpText, ccpFtn, ccpHdd,
        // ccpMcr, ccpAtn, ccpEdn, ccpTxbx, ccpHdrTxbx at 0x40 + 4*i.
        let counts = |i: usize| -> u32 {
            // The counts are cached on the Fib at parse time for the main text;
            // the rest are read from the same offsets on demand.
            self.fib.ccp_counts.get(i).copied().unwrap_or(0)
        };
        let start = match story {
            Story::Main => 0,
            Story::Footnote => counts(0),
            Story::Header => counts(0) + counts(1),
            // The macro story sits between the headers and the annotations;
            // it is almost always empty, but skipping it would shift every
            // later story by the macros' length.
            Story::Annotation => counts(0) + counts(1) + counts(2) + counts(3),
            Story::Endnote => counts(0) + counts(1) + counts(2) + counts(3) + counts(4),
            Story::Textbox => {
                counts(0) + counts(1) + counts(2) + counts(3) + counts(4) + counts(5)
            }
            Story::HeaderTextbox => {
                counts(0) + counts(1) + counts(2) + counts(3) + counts(4) + counts(5) + counts(6)
            }
        };
        let len = match story {
            Story::Main => counts(0),
            Story::Footnote => counts(1),
            Story::Header => counts(2),
            Story::Annotation => counts(4),
            Story::Endnote => counts(5),
            Story::Textbox => counts(6),
            Story::HeaderTextbox => counts(7),
        };
        (start, start + len)
    }

    /// The text of one story, control characters untouched.
    pub fn story(&self, story: Story) -> String {
        let (a, b) = self.story_range(story);
        self.full_text
            .chars()
            .skip(a as usize)
            .take((b - a) as usize)
            .collect()
    }

    /// The header story's slots as text, one entry per `PlcfHdd` slot. Only the
    /// per-section slots are named by [`Doc::header_footer`]; this returns the
    /// raw separators too, which is what the CP layout actually contains.
    pub fn hdd_story_slots(&self) -> Vec<String> {
        let story = self.story(Story::Header);
        let chars: Vec<char> = story.chars().collect();
        self.hdd_slots
            .windows(2)
            .map(|w| {
                let (a, b) = (w[0] as usize, w[1] as usize);
                chars.get(a..b).map(|s| s.iter().collect()).unwrap_or_default()
            })
            .collect()
    }

    /// The primary header and footer text for one section (0-based), as Word's
    /// `Headers.Item(1)` and `Footers.Item(1)` report them. The six
    /// document-wide separators come first, then six slots per section in the
    /// order even-page header, primary header, even-page footer, primary footer,
    /// first-page header, first-page footer.
    pub fn header_footer(&self, section: usize) -> Option<(String, String)> {
        self.header_footer_slots(section, 1, 3)
    }

    /// The header and footer of the first page of `section`, if the section
    /// carries a separate pair.
    ///
    /// Word only shows these when the section has `sprmSFTitlePage` set; the
    /// slots exist either way, so an empty pair means "not used".
    pub fn first_page_header_footer(&self, section: usize) -> Option<(String, String)> {
        self.header_footer_slots(section, 4, 5)
    }

    /// The header and footer of even-numbered pages of `section`.
    ///
    /// Word only shows these when the document's `DOP` has `fFacingPages` set.
    pub fn even_page_header_footer(&self, section: usize) -> Option<(String, String)> {
        self.header_footer_slots(section, 0, 2)
    }

    /// Read one pair of `PlcfHdd` slots, trimming the paragraph marks Word
    /// keeps in them.
    fn header_footer_slots(
        &self,
        section: usize,
        header_slot: usize,
        footer_slot: usize,
    ) -> Option<(String, String)> {
        let slots = self.hdd_story_slots();
        let base = 6 + section * 6;
        let header = slots.get(base + header_slot)?.trim_end_matches('\r').to_string();
        let footer = slots.get(base + footer_slot)?.trim_end_matches('\r').to_string();
        Some((header, footer))
    }

    /// Each footnote's reference mark, in order: the auto-numbered ones are
    /// `None`, custom-marked ones carry their literal character. The mark sits
    /// in the main text exactly where `footnote_slots[i]` points, so this reads
    /// that character back when the corresponding `FRD` has `fAuto` clear.
    pub fn footnote_marks(&self) -> Vec<Option<char>> {
        let main: Vec<char> = self.story(Story::Main).chars().collect();
        let auto = &self.footnote_auto;
        (0..auto.len())
            .map(|i| {
                if auto[i] {
                    None
                } else {
                    self.footnote_slots
                        .get(i)
                        .and_then(|cp| main.get(*cp as usize).copied())
                }
            })
            .collect()
    }

    /// Every `HYPERLINK` field in the main text, as `(url, display text)`.
    ///
    /// Word stores a hyperlink as a field, not a record: the instruction sits
    /// between `0x13` and `0x14`, the visible text between `0x14` and `0x15`
    /// (`docs/doc.md` §14.14). `.xls` links are a separate `HLINK` record, but
    /// here the field is the whole thing.
    pub fn hyperlinks(&self) -> Vec<(String, String)> {
        let text: Vec<char> = self.story(Story::Main).chars().collect();
        let mut out = Vec::new();
        let mut i = 0;
        while i < text.len() {
            if text[i] != '\u{13}' {
                i += 1;
                continue;
            }
            // Gather the instruction up to the separator.
            let mut j = i + 1;
            let mut instr = String::new();
            while j < text.len() && text[j] != '\u{14}' && text[j] != '\u{15}' {
                instr.push(text[j]);
                j += 1;
            }
            let mut shown = String::new();
            if j < text.len() && text[j] == '\u{14}' {
                j += 1;
                while j < text.len() && text[j] != '\u{15}' {
                    shown.push(text[j]);
                    j += 1;
                }
            }
            let instr = instr.trim();
            if let Some(rest) = instr.strip_prefix("HYPERLINK") {
                if let Some(url) = rest.trim().trim_matches('"').split('"').next() {
                    if !url.is_empty() {
                        // Word pads the visible text with a trailing space
                        // before the field end; strip it like it does.
                        out.push((url.to_string(), shown.trim_end().to_string()));
                    }
                }
            }
            i = j + 1;
        }
        out
    }

    /// Every non-empty story, main text included.
    pub fn stories(&self) -> Vec<(Story, String)> {
        Story::ALL
            .iter()
            .map(|s| (*s, self.story(*s)))
            .filter(|(_, t)| !t.is_empty())
            .collect()
    }
}

/// Text decoders live in `office-enc` so .doc/.xls/.ppt agree on CP1252.
pub use office_enc::{decode_cp1252, decode_utf16le};

impl Doc {
    /// Map a byte offset in the WordDocument stream to a character position in
    /// [`Doc::text`]. FKPs store byte offsets, the piece table stores
    /// characters, so every formatting run has to cross this bridge.
    pub fn fc_to_cp(&self, fc: u32) -> Option<u32> {
        for p in &self.pieces {
            // A compressed piece addresses two characters per byte, so the FKP
            // pages (which always speak in file offsets) start at `fc / 2`.
            // Microsoft Word relies on this: its `fc` is `0x40001000` for a
            // piece whose first byte is 0x1000, while WPS writes 0x800 (an
            // already-halved offset) — comparing the raw value works for one
            // writer and fails for the other.
            let raw = p.fc & 0x3FFF_FFFF;
            let scale = if p.compressed { 1 } else { 2 };
            let piece_fc = if p.compressed { raw / 2 } else { raw };
            let len = (p.cp_end - p.cp_start) * scale;
            // `fc` may be an *end* position (a formatting run ends at the
            // first byte it does not cover), so the upper bound is inclusive:
            // without this every PAPX/CHPX run that ends exactly at `fcMac`
            // silently disappears.
            if fc >= piece_fc && fc <= piece_fc + len {
                return Some(p.cp_start + (fc - piece_fc) / scale);
            }
        }
        None
    }

    /// Resolve a paragraph's `sprmPIlfo`/`sprmPIlvl` into a concrete marker.
    ///
    /// `ilfo` indexes [`Doc::list_overrides`] one-based; the override's `lsid`
    /// names the definition in [`Doc::lists`]. When the override carries its
    /// own level formatting we would use it, but Word's simplest files (such as
    /// `samples/ms/lists.doc`) leave `clfolvl = 0` and rely on the definition's
    /// levels, which is all this reader follows.
    fn list_format(&self, ilfo: Option<u16>, ilvl: Option<u8>) -> Option<fmt::ListFormat> {
        let ilfo = ilfo?;
        if ilfo == 0 {
            return None;
        }
        let lsid = *self.list_overrides.get(ilfo as usize - 1)?;
        let def = self.lists.iter().find(|d| d.lsid == lsid)?;
        let level = ilvl.unwrap_or(0);
        def.levels.get(level as usize).cloned()
    }

    /// Paragraphs with their formatting, in document order.
    ///
    /// Character runs are clipped to each paragraph and expressed as character
    /// offsets into `paragraph.text`, so a caller can render bold/italic runs
    /// without touching the FKP pages.
    pub fn paragraphs_rich(&self) -> Vec<fmt::Paragraph> {
        // Character runs as CP ranges, merged when neighbours agree.
        let mut cp_runs: Vec<(u32, u32, fmt::CharProps)> = Vec::new();
        for r in &self.char_runs {
            let (Some(a), Some(b)) = (self.fc_to_cp(r.fc_start), self.fc_to_cp(r.fc_end))
            else {
                continue;
            };
            match cp_runs.last_mut() {
                Some((_, end, props)) if *end == a && *props == r.props => *end = b,
                _ => cp_runs.push((a, b, r.props.clone())),
            }
        }
        let para_props: Vec<(u32, u32, fmt::ParaProps)> = self
            .para_runs
            .iter()
            .filter_map(|r| {
                Some((
                    self.fc_to_cp(r.fc_start)?,
                    self.fc_to_cp(r.fc_end)?,
                    r.props.clone(),
                ))
            })
            .collect();

        let chars: Vec<char> = self.text.chars().collect();
        let mut out = Vec::new();
        let mut start = 0u32;
        for (i, ch) in chars.iter().enumerate() {
            if start as usize > i {
                continue; // inside a paragraph already emitted
            }
            let last = i + 1 == chars.len();
            if *ch != PARA_MARK && *ch != CELL_MARK && !last {
                continue;
            }
            // A cell mark is a boundary too: the mark itself is not part of the
            // cell's text, so the paragraph ends *before* it.
            let end = if *ch == PARA_MARK || *ch == CELL_MARK {
                i as u32
            } else {
                i as u32 + 1
            };
            // A paragraph mark that no PAPX run covers inherits the previous
            // paragraph's properties, exactly like a run with `bOffset == 0`.
            let mut props = out
                .last()
                .map(|p: &fmt::Paragraph| p.props.clone())
                .unwrap_or_default();
            for (a, b, p) in &para_props {
                if *a <= start && start < *b {
                    props = p.clone();
                    break;
                }
            }
            // The style's own properties sit underneath the direct formatting
            // that `PAPX`/`CHPX` carry: start from the style, then let the
            // run's sprms win.
            if let Some(id) = props.style_id {
                if let Some(style) = self.style_props(id) {
                    let mut base = style.para.clone();
                    base.merge(&props);
                    props = base;
                }
            }
            let mut runs = Vec::new();
            for (a, b, cp) in &cp_runs {
                let (lo, hi) = ((*a).max(start), (*b).min(end));
                if lo < hi {
                    let mut cprops = cp.clone();
                    if let Some(id) = cprops.style_id {
                        if let Some(style) = self.style_props(id) {
                            let mut base = style.char.clone();
                            base.merge(&cprops);
                            cprops = base;
                        }
                    }
                    runs.push(fmt::CharSpan {
                        start: (lo - start) as usize,
                        end: (hi - start) as usize,
                        props: cprops,
                    });
                }
            }
            let style_name = props
                .style_id
                .and_then(|id| self.style_name(id))
                .map(str::to_owned);
            let list = self.list_format(props.ilfo, props.ilvl);
            out.push(fmt::Paragraph {
                text: chars[start as usize..end as usize].iter().collect(),
                props,
                runs,
                style_name,
                list,
            });
            start = end + 1;
        }
        // A trailing paragraph mark terminates the last paragraph rather than
        // starting a new one; drop the empty paragraph it would otherwise add.
        if out.last().is_some_and(|p: &fmt::Paragraph| p.text.is_empty()) {
            out.pop();
        }
        out
    }

    /// Text with Word's in-line control characters translated to something a
    /// terminal can show: `\r` -> `\n`, cell/row marks -> `\t`, field
    /// delimiters (0x13/0x14/0x15) and object anchors removed.
    pub fn plain_text(&self) -> String {
        let mut out = String::with_capacity(self.text.len());
        for ch in self.text.chars() {
            match ch {
                '\r' => out.push('\n'),
                '\u{0007}' => out.push('\t'),
                '\u{000B}' => out.push('\n'),
                '\u{0013}' | '\u{0014}' | '\u{0015}' => {}
                '\u{0001}' | '\u{0002}' | '\u{0008}' => {}
                c => out.push(c),
            }
        }
        out
    }

    /// Paragraph strings (split on Word's `\r` paragraph mark).
    pub fn paragraphs(&self) -> Vec<String> {
        self.text
            .split('\r')
            .map(|p| p.replace('\u{0007}', "\t"))
            .filter(|p| !p.is_empty())
            .collect()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn cp1252_special_block() {
        assert_eq!(decode_cp1252(&[0x93, b'h', b'i', 0x94]), "\u{201C}hi\u{201D}");
        assert_eq!(decode_cp1252(b"abc"), "abc");
        assert_eq!(decode_cp1252(&[0xE9]), "é");
    }

    #[test]
    fn utf16_surrogates_are_lossy() {
        assert_eq!(decode_utf16le(&[0x2D, 0x4E]), "中");
    }

    fn sample() -> Option<Doc> {
        let p = concat!(env!("CARGO_MANIFEST_DIR"), "/../../samples/hello.doc");
        if !std::path::Path::new(p).exists() {
            return None;
        }
        Some(open(p).expect("parse hello.doc"))
    }

    fn formatted() -> Option<Doc> {
        let p = concat!(env!("CARGO_MANIFEST_DIR"), "/../../samples/formatted.doc");
        if !std::path::Path::new(p).exists() {
            return None;
        }
        Some(open(p).expect("parse formatted.doc"))
    }

    fn encrypted() -> Option<Vec<u8>> {
        let p = concat!(env!("CARGO_MANIFEST_DIR"), "/../../samples/encrypted.doc");
        if !std::path::Path::new(p).exists() {
            return None;
        }
        Some(std::fs::read(p).expect("read encrypted.doc"))
    }

    #[test]
    fn refuses_an_encrypted_document_without_a_password() {
        let Some(bytes) = encrypted() else { return };
        assert!(matches!(parse(&bytes), Err(Error::Encrypted)));
        assert!(matches!(
            parse_with_password(&bytes, "wrong-password"),
            Err(Error::WrongPassword)
        ));
    }

    #[test]
    fn decrypts_the_sample_document() {
        let Some(bytes) = encrypted() else { return };
        let d = parse_with_password(&bytes, "secret123").expect("decrypt");
        let scheme = d.encryption.expect("scheme recorded");
        assert_eq!(scheme.name(), "RC4 (legacy, v1.1)");
        assert_eq!(d.fib.n_fib, 0x00C1);
        assert_eq!(d.fib.ccp_text, 39);
        assert_eq!(d.fib.flags & 0x0100, 0x0100, "fEncrypted stays set in the FIB");
        assert!(d.text.starts_with("ENCRYPTED-DOC-777"));
        assert!(d.text.contains("Second secret line."));
    }

    #[test]
    fn splits_the_text_stream_into_stories() {
        let p = concat!(env!("CARGO_MANIFEST_DIR"), "/../../samples/stories.doc");
        if !std::path::Path::new(p).exists() {
            return;
        }
        let d = open(p).expect("parse stories.doc");
        // fibRgLw97: ccpText 57, ccpFtn 19, ccpHdd 33, the rest empty.
        assert_eq!(d.fib.ccp_counts, [57, 19, 33, 0, 0, 0, 0, 0]);
        let main = d.story(Story::Main);
        assert!(main.starts_with("Body paragraph one.\rBody paragraph two.\r"));
        assert_eq!(main.chars().count(), 57);
        let ftn = d.story(Story::Footnote);
        assert_eq!(ftn.chars().count(), 19);
        assert!(ftn.contains("FOOTNOTE-MARK-33"));
        let hdd = d.story(Story::Header);
        assert_eq!(hdd.chars().count(), 33);
        assert!(hdd.contains("HEADER-MARK-11"));
        assert!(hdd.contains("FOOTER-MARK-22"));
        // The stories partition the stream exactly.
        let total: usize = d.stories().iter().map(|(_, t)| t.chars().count()).sum();
        assert_eq!(total, d.full_text.chars().count() - 1);
        assert!(d.story(Story::Endnote).is_empty());
    }

    #[test]
    fn reads_comments_from_the_sample() {
        let dir = concat!(env!("CARGO_MANIFEST_DIR"), "/../../samples/ms");
        let path = format!("{dir}/comments.doc");
        if !std::path::Path::new(&path).exists() {
            return;
        }
        let d = open(&path).expect("parse comments sample");
        assert_eq!(d.comments.len(), 2);
        assert_eq!(d.comments[0].initials, "jy");
        assert_eq!(d.comments[0].text, "needs a citation");
        assert_eq!(d.comments[1].text, "second remark");
        // The reference mark sits in the main text at each anchor.
        let marks: Vec<u32> = d
            .text
            .chars()
            .enumerate()
            .filter(|(_, c)| *c == '\u{5}')
            .map(|(i, _)| i as u32)
            .collect();
        assert_eq!(marks, vec![d.comments[0].anchor, d.comments[1].anchor]);
    }

    #[test]
    fn reads_the_primary_header_and_footer_from_the_sample() {
        let dir = concat!(env!("CARGO_MANIFEST_DIR"), "/../../samples/ms");
        let header = format!("{dir}/hdd-hdr-primary.doc");
        if !std::path::Path::new(&header).exists() {
            return;
        }
        let d = open(&header).expect("parse header sample");
        assert_eq!(d.hdd_slots.len(), 14, "six separators, six slots, closing");
        assert_eq!(d.header_footer(0).unwrap(), ("MARK-hdr-primary".into(), String::new()));
        let footer = open(format!("{dir}/hdd-ftr-primary.doc")).expect("footer sample");
        assert_eq!(footer.header_footer(0).unwrap(), (String::new(), "MARK-ftr-primary".into()));
        // A document without a header story has no slots at all.
        let plain = open(concat!(env!("CARGO_MANIFEST_DIR"), "/../../samples/written.doc"))
            .expect("plain sample");
        assert!(plain.hdd_slots.is_empty());
        assert!(plain.header_footer(0).is_none());
    }

    #[test]
    fn reads_a_word_hyperlink_field_from_the_sample() {
        let path = concat!(env!("CARGO_MANIFEST_DIR"), "/../../samples/ms/hyperlink.doc");
        if !std::path::Path::new(path).exists() {
            return;
        }
        let d = open(path).expect("parse the hyperlink sample");
        assert_eq!(
            d.hyperlinks(),
            vec![("https://example.com/guide".to_string(), "the site".to_string())]
        );
    }

    #[test]
    fn reads_hyperlinks_as_fields() {
        let body = vec![
            Paragraph::text("Intro line."),
            Paragraph::text(format!(
                "Go to {} now.",
                crate::write::hyperlink("https://example.com/page", "the site")
            )),
        ];
        let bytes = crate::write::write_document(&body);
        let d = parse(&bytes).unwrap();
        assert_eq!(d.hyperlinks(), vec![("https://example.com/page".to_string(), "the site".to_string())]);
    }

    #[test]
    fn reads_custom_footnote_marks_from_the_sample() {
        let dir = concat!(env!("CARGO_MANIFEST_DIR"), "/../../samples/ms");
        let path = format!("{dir}/footnote-custom.doc");
        if !std::path::Path::new(&path).exists() {
            return;
        }
        let d = open(&path).expect("parse the custom-mark sample");
        // Two notes: a starred one (custom) and an auto-numbered one.
        assert_eq!(d.footnote_auto, vec![false, true]);
        assert_eq!(d.footnote_marks(), vec![Some('*'), None]);
        let story = d.story(Story::Footnote);
        assert!(story.starts_with("* star note"), "story was {story:?}");
    }

    #[test]
    fn reads_character_formatting_from_fkp_pages() {
        let Some(d) = formatted() else { return };
        let paras = d.paragraphs_rich();
        assert_eq!(paras.len(), 5, "one entry per paragraph");

        // 1: bold, 18 pt, red
        let p0 = &paras[0];
        assert_eq!(p0.text, "Bold red heading");
        let r0 = &p0.runs[0];
        assert_eq!(r0.props.bold, Some(true));
        assert_eq!(r0.props.size_half_points, Some(36));
        assert_eq!(r0.props.color, Some(0xFF)); // sprmCCv: 0x000000FF

        // 2: italic+underline, with a bold word in the middle
        let p1 = &paras[1];
        assert_eq!(p1.runs.len(), 3);
        assert_eq!(p1.runs[0].props.italic, Some(true));
        assert_eq!(p1.runs[0].props.underline, Some(1));
        assert_eq!(p1.runs[0].props.bold, None);
        let mid: String = p1.text.chars().skip(p1.runs[1].start).take(p1.runs[1].end - p1.runs[1].start).collect();
        assert_eq!(mid, "mixed ");
        assert_eq!(p1.runs[1].props.bold, Some(true));

        // 3..5: no character properties at all
        for p in &paras[2..] {
            assert!(p.runs.iter().all(|r| r.props == fmt::CharProps::default()));
        }
    }

    #[test]
    fn formatting_runs_cover_the_text() {
        let Some(d) = formatted() else { return };
        for p in d.paragraphs_rich() {
            let mut at = 0usize;
            for r in &p.runs {
                assert_eq!(r.start, at, "runs are contiguous in {:?}", p.text);
                assert!(r.end > r.start);
                at = r.end;
            }
            assert_eq!(at, p.text.chars().count());
        }
    }

    #[test]
    fn reads_hello_doc() {
        let Some(doc) = sample() else { return };
        assert_eq!(doc.fib.w_ident, W_IDENT);
        assert_eq!(doc.fib.n_fib, 0x00C1);
        assert!(doc.text.contains("HELLO-DOC-001"));
        assert!(doc.text.contains("中文测试"));
        assert!(doc.text.contains("END-MARKER-42"));
        assert!(doc.text.chars().count() >= doc.fib.ccp_text as usize);
    }
}
