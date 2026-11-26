//! Writing Word 97 documents, with character and paragraph formatting.
//!
//! This is the reverse of the reader: the same FIB pairs, the same CLX, the same
//! CHPX/PAPX pages — only the direction differs. Keeping both sides in one crate
//! means a wrong field width shows up as a failed round-trip instead of a
//! mystery in Word.
//!
//! What Word actually insists on was established by experiment, not by reading
//! MS-DOC (see `docs/doc.md` §14): a document whose `PlcfBteChpx`/`PlcfBtePapx`
//! pairs are empty hangs the reader, and a missing stylesheet or section table
//! makes it report "the document could not be opened". So the writer always
//! emits one CHPX page, one PAPX page, a one-style `STSH` and a one-section
//! `PlcfSed`.

use crate::fmt::{CharProps, ParaProps};

/// Byte offset where the text starts inside the `WordDocument` stream.
const TEXT_AT: usize = 0x800;
/// FIB object size: `FibBase` + `fibRgW97` + `fibRgLw97` + `FibRgFcLcb97` + `cswNew`.
const FIB_LEN: usize = 0x380;
/// A formatting page is one sector; the run count lives in the last byte.
const PAGE: usize = 512;
/// Formatting data starts here so the run arrays never collide with it.
const DATA_AT: usize = 0x100;

/// A stretch of text that shares one set of character properties.
#[derive(Clone, Debug, Default)]
pub struct Run {
    pub text: String,
    pub props: CharProps,
}

impl Run {
    pub fn new(text: impl Into<String>) -> Self {
        Run {
            text: text.into(),
            props: CharProps::default(),
        }
    }

    pub fn styled(text: impl Into<String>, props: CharProps) -> Self {
        Run {
            text: text.into(),
            props,
        }
    }
}

/// One paragraph: one or more runs plus its paragraph properties.
///
/// The paragraph mark is written for you; `text()` splits on `\n` so the common
/// case needs no run bookkeeping at all.
#[derive(Clone, Debug, Default)]
pub struct Paragraph {
    pub runs: Vec<Run>,
    pub props: ParaProps,
    /// The character that ends the paragraph: `\r` for a normal one, `\x07`
    /// for a table cell (and for the row-end mark).
    pub mark: char,
}

impl Paragraph {
    pub fn text(text: impl Into<String>) -> Self {
        Paragraph { mark: '\r', ..Paragraph::raw(text) }
    }

    fn raw(text: impl Into<String>) -> Self {
        Paragraph {
            runs: vec![Run::new(text)],
            props: ParaProps::default(),
            mark: char::from(13),
        }
    }

    pub fn styled(runs: Vec<Run>, props: ParaProps) -> Self {
        Paragraph { runs, props, mark: char::from(13) }
    }

    /// One table cell: the paragraph mark becomes the cell mark (``).
    pub fn cell(runs: Vec<Run>, props: ParaProps) -> Self {
        Paragraph { runs, props, mark: char::from(7) }
    }

    /// The mark that ends a table row: empty text, `fTtp` set, and the column
    /// widths that `TDefTable` carries.
    pub fn row_end(widths: &[i16]) -> Self {
        let mut props = ParaProps::default();
        props.in_table = Some(true);
        props.ttp = Some(true);
        props.col_widths = widths.to_vec();
        Paragraph { runs: Vec::new(), props, mark: char::from(7) }
    }
}

/// Build a `.doc` from plain text, one paragraph per line.
pub fn write_minimal(text: &str) -> Vec<u8> {
    // Both separators are accepted: a caller may hand us text copied out of a
    // Word stream (CR marks) or a plain file (LF).
    let paragraphs: Vec<Paragraph> = text
        .split(['\n', '\r'])
        .filter(|line| !line.is_empty())
        .map(Paragraph::text)
        .collect();
    write_document(&paragraphs)
}

/// Build a `.doc` from paragraphs that carry character and paragraph properties.
pub fn write_document(paragraphs: &[Paragraph]) -> Vec<u8> {
    // ---- lay the text out and record where every run starts and ends --------
    let mut utf16: Vec<u8> = Vec::new();
    let mut spans: Vec<(u32, u32, CharProps)> = Vec::new();
    let mut para_ends: Vec<(u32, u32, ParaProps)> = Vec::new();
    for paragraph in paragraphs {
        let para_start = TEXT_AT as u32 + utf16.len() as u32;
        let mut runs: Vec<(u32, u32, CharProps)> = Vec::new();
        for run in &paragraph.runs {
            let start = TEXT_AT as u32 + utf16.len() as u32;
            for ch in run.text.chars() {
                utf16.extend_from_slice(&(ch as u16).to_le_bytes());
            }
            runs.push((start, TEXT_AT as u32 + utf16.len() as u32, run.props.clone()));
        }
        if runs.is_empty() {
            runs.push((
                TEXT_AT as u32 + utf16.len() as u32,
                TEXT_AT as u32 + utf16.len() as u32,
                CharProps::default(),
            ));
        }
        // The paragraph mark belongs to the last run.
        utf16.extend_from_slice(&(paragraph.mark as u16).to_le_bytes());
        let end = TEXT_AT as u32 + utf16.len() as u32;
        if let Some(last) = runs.last_mut() {
            last.1 = end;
        }
        spans.extend(runs);
        para_ends.push((para_start, end, paragraph.props.clone()));
    }
    let ccp_text = utf16.len() as u32 / 2;
    let fc_min = TEXT_AT as u32;
    let fc_mac = TEXT_AT as u32 + utf16.len() as u32;

    // ---- formatting pages --------------------------------------------------
    // Each page holds a run array, a per-run offset array and the sprms, all
    // inside one 512-byte sector, so the page split depends on how big the
    // sprms are — not just on how many runs there are.
    let chpx_bodies: Vec<Vec<u8>> = spans.iter().map(|(_, _, p)| char_grpprl(p)).collect();
    let chpx_sizes: Vec<usize> = chpx_bodies
        .iter()
        .map(|b| if b.is_empty() { 0 } else { (1 + b.len() + 1) & !1 })
        .collect();
    let papx_bodies: Vec<Vec<u8>> = para_ends
        .iter()
        .map(|(_, _, props)| {
            // A PAPX body is `istd u16` followed by the paragraph sprms. The
            // `cb == 0` form stores the length in two-byte units, which is what
            // Word itself writes, so the body must be even. Padding with a
            // stray zero byte would leave Word reading a truncated `sprm` id,
            // so pad with a real no-op sprm instead (`sprmPFWidowControl`,
            // whose default is already "on").
            let mut body: Vec<u8> = Vec::new();
            body.extend_from_slice(&props.style_id.unwrap_or(0).to_le_bytes());
            body.extend_from_slice(&para_grpprl(props));
            if body.len() % 2 == 1 {
                body.extend_from_slice(&[0x31, 0x24, 0x01]);
            }
            body
        })
        .collect();
    let papx_sizes: Vec<usize> = papx_bodies.iter().map(|b| 2 + b.len()).collect();
    let chpx_pages = paginate_fit(&chpx_sizes, |n| 4 * (n + 1) + n);
    let papx_pages = paginate_fit(&papx_sizes, |n| 4 * (n + 1) + 13 * n);
    let page_count = chpx_pages.len() + papx_pages.len();
    // A `SEPX` (section properties) sits between the text and the first
    // formatting page. Word reads the section through `PlcfSed.fcSepx`, and a
    // `0xFFFFFFFF` "no SEPX" makes it treat the body as empty.
    let mut sepx: Vec<u8> = Vec::new();
    push_sprm(&mut sepx, 0xB01F, &12240u16.to_le_bytes()); // sprmSXaPage
    push_sprm(&mut sepx, 0xB020, &15840u16.to_le_bytes()); // sprmSXaPgHt
    push_sprm(&mut sepx, 0xB021, &1800i16.to_le_bytes()); // sprmSDxaLeft
    push_sprm(&mut sepx, 0xB022, &1800i16.to_le_bytes()); // sprmSDxaRight
    let first_page = fc_mac.div_ceil(PAGE as u32) as usize;
    // Word puts the section descriptor after the formatting pages, at the very
    // end of the stream; `PlcfSed.fcSepx` points at it.
    let fc_sepx = ((first_page + page_count) * PAGE) as u32;
    let stream_len = fc_sepx as usize + 2 + sepx.len();

    let mut wd = vec![0u8; stream_len];
    wd[TEXT_AT..TEXT_AT + utf16.len()].copy_from_slice(&utf16);
    let sepx_at = fc_sepx as usize;
    wd[sepx_at..sepx_at + 2].copy_from_slice(&(sepx.len() as u16).to_le_bytes());
    wd[sepx_at + 2..sepx_at + 2 + sepx.len()].copy_from_slice(&sepx);

    // CHPX pages: `rgfc`, then one byte per run, then `cb`-prefixed sprms.
    let mut chpx_plc: Vec<(u32, u32)> = Vec::new();
    let mut next_page = first_page;
    for chunk in chpx_pages {
        let page = &mut wd[next_page * PAGE..(next_page + 1) * PAGE];
        let runs = &spans[chunk.clone()];
        page[0..4].copy_from_slice(&runs[0].0.to_le_bytes());
        for (i, (_, end, _)) in runs.iter().enumerate() {
            page[(i + 1) * 4..(i + 2) * 4].copy_from_slice(&end.to_le_bytes());
        }
        let mut data = DATA_AT;
        for (i, grpprl) in chpx_bodies[chunk.clone()].iter().enumerate() {
            if grpprl.is_empty() {
                continue;
            }
            page[(runs.len() + 1) * 4 + i] = (data / 2) as u8;
            page[data] = grpprl.len() as u8;
            page[data + 1..data + 1 + grpprl.len()].copy_from_slice(grpprl);
            data += 1 + grpprl.len();
            if data % 2 == 1 {
                data += 1;
            }
        }
        page[511] = runs.len() as u8;
        chpx_plc.push((runs[0].0, next_page as u32));
        next_page += 1;
    }
    chpx_plc.push((fc_mac, 0));

    // PAPX pages: `rgfc`, then 13-byte `BX` entries, then the PAPX bodies.
    let mut papx_plc: Vec<(u32, u32)> = Vec::new();
    for chunk in papx_pages {
        let page = &mut wd[next_page * PAGE..(next_page + 1) * PAGE];
        let paras = &para_ends[chunk.clone()];
        page[0..4].copy_from_slice(&paras[0].0.to_le_bytes());
        for (i, (_, end, _)) in paras.iter().enumerate() {
            page[(i + 1) * 4..(i + 2) * 4].copy_from_slice(&end.to_le_bytes());
        }
        let mut data = DATA_AT;
        for (i, body) in papx_bodies[chunk.clone()].iter().enumerate() {
            page[(paras.len() + 1) * 4 + i * 13] = (data / 2) as u8;
            page[data] = 0;
            page[data + 1] = (body.len() / 2) as u8;
            page[data + 2..data + 2 + body.len()].copy_from_slice(&body);
            data += 2 + body.len();
            if data % 2 == 1 {
                data += 1;
            }
        }
        page[511] = paras.len() as u8;
        papx_plc.push((paras[0].0, next_page as u32));
        next_page += 1;
    }
    papx_plc.push((fc_mac, 0));

    // ---- 1Table: CLX, stylesheet, section table, BTE tables -----------------
    let mut table: Vec<u8> = Vec::new();

    // CLX: one piece covering the whole text, stored as UTF-16.
    table.push(0x02); // clxt = Pcdt
    table.extend_from_slice(&16u32.to_le_bytes()); // lcb: 2 CPs + 1 PCD
    table.extend_from_slice(&0u32.to_le_bytes()); // CP[0]
    table.extend_from_slice(&ccp_text.to_le_bytes()); // CP[1]
    table.extend_from_slice(&0u16.to_le_bytes()); // PCD flags
    table.extend_from_slice(&fc_min.to_le_bytes()); // fc, bit 30 clear: UTF-16
    table.extend_from_slice(&0u16.to_le_bytes()); // prm
    let clx_len: u32 = 5 + 16;

    // Stylesheet: `cbStshi` + `STSHI` + one `Normal` style.
    let stsh_at = table.len() as u32;
    let cb_base: u16 = 10;
    table.extend_from_slice(&18u16.to_le_bytes()); // cbStshi
    table.extend_from_slice(&1u16.to_le_bytes()); // cstd
    table.extend_from_slice(&cb_base.to_le_bytes()); // cbSTDBaseInFile
    table.extend_from_slice(&1u16.to_le_bytes()); // flags: names written
    table.extend_from_slice(&1u16.to_le_bytes()); // stiMaxWhenSaved
    table.extend_from_slice(&15u16.to_le_bytes()); // istdMaxFixedWhenSaved
    table.extend_from_slice(&1u16.to_le_bytes()); // nVerBuiltInNamesWhenSaved
    table.extend_from_slice(&[0u8; 6]); // rgftc[3]
    let name = "Normal";
    let cb_std = cb_base as usize + 2 + name.len() * 2 + 2;
    table.extend_from_slice(&(cb_std as u16).to_le_bytes());
    table.extend_from_slice(&[0u8; 10]); // STDBase: defaults
    table.extend_from_slice(&(name.len() as u16).to_le_bytes());
    for ch in name.chars() {
        table.extend_from_slice(&(ch as u16).to_le_bytes());
    }
    table.extend_from_slice(&0u16.to_le_bytes()); // NUL terminator
    let stsh_len = table.len() as u32 - stsh_at;

    // One section covering the whole text.
    let sed_at = table.len() as u32;
    table.extend_from_slice(&0u32.to_le_bytes()); // CP[0]
    table.extend_from_slice(&ccp_text.to_le_bytes()); // CP[1]
    table.extend_from_slice(&0u16.to_le_bytes()); // fn
    table.extend_from_slice(&fc_sepx.to_le_bytes()); // fcSepx
    table.extend_from_slice(&0u16.to_le_bytes()); // fnMpr
    table.extend_from_slice(&0xFFFF_FFFFu32.to_le_bytes()); // fcMpr
    let sed_len = table.len() as u32 - sed_at;

    let chpx_plc_at = table.len() as u32;
    for (fc, _) in &chpx_plc {
        table.extend_from_slice(&fc.to_le_bytes());
    }
    for (_, pn) in &chpx_plc {
        if *pn != 0 {
            table.extend_from_slice(&pn.to_le_bytes());
        }
    }
    let chpx_plc_len = table.len() as u32 - chpx_plc_at;

    let papx_plc_at = table.len() as u32;
    for (fc, _) in &papx_plc {
        table.extend_from_slice(&fc.to_le_bytes());
    }
    for (_, pn) in &papx_plc {
        if *pn != 0 {
            table.extend_from_slice(&pn.to_le_bytes());
        }
    }
    let papx_plc_len = table.len() as u32 - papx_plc_at;

    // ---- FIB ---------------------------------------------------------------
    let mut fib = vec![0u8; FIB_LEN];
    fib[0x00..0x02].copy_from_slice(&0xA5ECu16.to_le_bytes()); // wIdent
    fib[0x02..0x04].copy_from_slice(&0x00C1u16.to_le_bytes()); // nFib (Word 97)
    fib[0x06..0x08].copy_from_slice(&0x0409u16.to_le_bytes()); // lid (en-US)
    // fComplex (bit 2) is required as soon as a piece table (CLX) exists:
    // without it Word ignores the CLX and reads fcMin..fcMac as one run.
    // fExtChar (bit 12) | fWhichTblStm (bit 9) => UTF-16 text, auxiliary 1Table.
    fib[0x0A..0x0C].copy_from_slice(&0x1204u16.to_le_bytes());
    fib[0x0C..0x0E].copy_from_slice(&0x00BFu16.to_le_bytes()); // nFibBack
    fib[0x18..0x1C].copy_from_slice(&fc_min.to_le_bytes()); // fcMin
    fib[0x1C..0x20].copy_from_slice(&fc_mac.to_le_bytes()); // fcMac
    fib[0x20..0x22].copy_from_slice(&0x000Eu16.to_le_bytes()); // csw
    fib[0x3E..0x40].copy_from_slice(&0x0016u16.to_le_bytes()); // cslw
    fib[0x40..0x44].copy_from_slice(&(stream_len as u32).to_le_bytes()); // cbMac
    fib[0x4C..0x50].copy_from_slice(&ccp_text.to_le_bytes()); // ccpText
    fib[0x98..0x9A].copy_from_slice(&0x005Du16.to_le_bytes()); // cbRgFcLcb

    // FibRgFcLcb97 pairs: 0/1 stylesheet, 6 section table, 12/13 BTE, 33 CLX.
    fib[0x9A..0x9E].copy_from_slice(&stsh_at.to_le_bytes());
    fib[0x9E..0xA2].copy_from_slice(&stsh_len.to_le_bytes());
    fib[0xA2..0xA6].copy_from_slice(&stsh_at.to_le_bytes());
    fib[0xA6..0xAA].copy_from_slice(&stsh_len.to_le_bytes());
    fib[0xCA..0xCE].copy_from_slice(&sed_at.to_le_bytes());
    fib[0xCE..0xD2].copy_from_slice(&sed_len.to_le_bytes());
    fib[0x0FA..0x0FE].copy_from_slice(&chpx_plc_at.to_le_bytes());
    fib[0x0FE..0x102].copy_from_slice(&chpx_plc_len.to_le_bytes());
    fib[0x102..0x106].copy_from_slice(&papx_plc_at.to_le_bytes());
    fib[0x106..0x10A].copy_from_slice(&papx_plc_len.to_le_bytes());
    let clx_pair = 0x9A + 33 * 8;
    fib[clx_pair..clx_pair + 4].copy_from_slice(&0u32.to_le_bytes()); // fcClx
    fib[clx_pair + 4..clx_pair + 8].copy_from_slice(&clx_len.to_le_bytes());

    wd[..FIB_LEN].copy_from_slice(&fib);

    office_cfb::write(&[
        office_cfb::OutputStream::new("WordDocument", &wd),
        office_cfb::OutputStream::new("1Table", &table),
    ])
}

/// Split items into formatting pages: `overhead(n)` is the space the run array
/// and the offset array take for `n` runs, `sizes` is each item's data size.
fn paginate_fit(sizes: &[usize], overhead: impl Fn(usize) -> usize) -> Vec<std::ops::Range<usize>> {
    let mut out = Vec::new();
    let mut start = 0usize;
    while start < sizes.len() {
        let mut end = start;
        let mut data = 0usize;
        while end < sizes.len() {
            let count = end - start + 1;
            let next = data + sizes[end];
            if overhead(count) > DATA_AT || DATA_AT + next > 511 {
                break;
            }
            data = next;
            end += 1;
        }
        if end == start {
            end = start + 1; // a single oversized item still gets its own page
        }
        out.push(start..end);
        start = end;
    }
    if out.is_empty() {
        out.push(0..0);
    }
    out
}

fn push_sprm(out: &mut Vec<u8>, id: u16, operand: &[u8]) {
    out.extend_from_slice(&id.to_le_bytes());
    out.extend_from_slice(operand);
}

/// Encode character properties the way Word stores them in a CHPX.
fn char_grpprl(props: &CharProps) -> Vec<u8> {
    let mut out = Vec::new();
    if let Some(v) = props.bold {
        push_sprm(&mut out, 0x0835, &[v as u8]);
    }
    if let Some(v) = props.italic {
        push_sprm(&mut out, 0x0836, &[v as u8]);
    }
    if let Some(v) = props.strike {
        push_sprm(&mut out, 0x0837, &[v as u8]);
    }
    if let Some(v) = props.double_strike {
        push_sprm(&mut out, 0x2A53, &[v as u8]);
    }
    if let Some(v) = props.caps {
        push_sprm(&mut out, 0x083B, &[v as u8]);
    }
    if let Some(v) = props.small_caps {
        push_sprm(&mut out, 0x083A, &[v as u8]);
    }
    if let Some(v) = props.underline {
        push_sprm(&mut out, 0x2A3E, &[v]);
    }
    if let Some(v) = props.size_half_points {
        push_sprm(&mut out, 0x4A43, &v.to_le_bytes());
    }
    // Word only honours `sprmCCv` when `sprmCIco` says "use the RGB value".
    if let Some(v) = props.color {
        push_sprm(&mut out, 0x2A42, &[props.color_index.unwrap_or(0)]);
        push_sprm(&mut out, 0x6870, &v.to_le_bytes());
    } else if let Some(v) = props.color_index {
        push_sprm(&mut out, 0x2A42, &[v]);
    }
    if let Some(v) = props.font_id {
        push_sprm(&mut out, 0x4A4F, &v.to_le_bytes());
    }
    if let Some(v) = props.iss {
        push_sprm(&mut out, 0x2A48, &[v]);
    }
    if let Some(v) = props.highlight {
        push_sprm(&mut out, 0x2A0C, &[v]);
    }
    if let Some(v) = props.style_id {
        push_sprm(&mut out, 0x4A30, &v.to_le_bytes());
    }
    out
}

/// Encode paragraph properties (the `istd` is written separately).
fn para_grpprl(props: &ParaProps) -> Vec<u8> {
    let mut out = Vec::new();
    // Always state the alignment. A document whose stylesheet carries no
    // alignment (ours does not) otherwise makes Word fall back to justified
    // text, which is not what the caller asked for.
    push_sprm(&mut out, 0x2403, &[props.justify.unwrap_or(0)]);
    if let Some(v) = props.dxa_right {
        push_sprm(&mut out, 0x840E, &v.to_le_bytes());
    }
    if let Some(v) = props.dxa_left {
        push_sprm(&mut out, 0x840F, &v.to_le_bytes());
    }
    if let Some(v) = props.dxa_left_first {
        push_sprm(&mut out, 0x8411, &v.to_le_bytes());
    }
    if let Some(v) = props.dya_before {
        push_sprm(&mut out, 0xA413, &v.to_le_bytes());
    }
    if let Some(v) = props.dya_after {
        push_sprm(&mut out, 0xA414, &v.to_le_bytes());
    }
    if let Some(v) = props.in_table {
        push_sprm(&mut out, 0x2416, &[v as u8]);
        // Table depth. Word refuses a table whose paragraphs do not carry it.
        push_sprm(&mut out, 0x6649, &1u32.to_le_bytes());
    }
    if let Some(v) = props.ttp {
        push_sprm(&mut out, 0x2417, &[v as u8]);
    }
    if let Some(v) = props.ilvl {
        push_sprm(&mut out, 0x260A, &[v]);
    }
    if let Some(v) = props.ilfo {
        push_sprm(&mut out, 0x460B, &v.to_le_bytes());
    }
    if let Some(v) = props.outline_level {
        push_sprm(&mut out, 0x2640, &[v]);
    }
    if props.col_widths.len() >= 2 {
        // spra 6: the operand is preceded by its own byte length.
        let op = tdef_table(&props.col_widths);
        out.extend_from_slice(&0xD608u16.to_le_bytes());
        out.push(op.len() as u8);
        out.extend_from_slice(&op);
    }
    out
}

/// `TDefTableOperand`: a zero lead byte, `itcMac`, then `itcMac + 1` cell
/// centres in twips followed by `itcMac` 20-byte `TC` entries.
fn tdef_table(centres: &[i16]) -> Vec<u8> {
    let itc = (centres.len() - 1).min(63) as u8;
    let mut out = Vec::with_capacity(2 + (itc as usize + 1) * 2 + itc as usize * 20);
    out.push(0); // undocumented lead byte, zero in every file we have seen
    out.push(itc);
    for c in &centres[..itc as usize + 1] {
        out.extend_from_slice(&c.to_le_bytes());
    }
    out.extend(std::iter::repeat(0u8).take(itc as usize * 20)); // plain cells
    out
}

/// A table to be written: the cells plus the column boundaries.
///
/// Word stores a table as ordinary paragraphs whose mark is the cell mark
/// (`U+0007`); the row ends with one more such paragraph carrying `fTtp` and a
/// `TDefTable` operand with the cell centres. `Table` is just a convenient way
/// to build those paragraphs.
#[derive(Clone, Debug, Default)]
pub struct Table {
    pub rows: Vec<Vec<String>>,
    /// `itcMac + 1` cell centres in twips, measured from the left margin.
    pub centres: Vec<i16>,
}

impl Table {
    /// Build a table with equally wide columns (1440 twips = 1 inch each).
    pub fn new(rows: &[&[&str]]) -> Table {
        let columns = rows.iter().map(|r| r.len()).max().unwrap_or(0);
        Table {
            rows: rows
                .iter()
                .map(|r| r.iter().map(|c| (*c).to_string()).collect())
                .collect(),
            centres: (0..=columns).map(|i| (i as i16) * 1440).collect(),
        }
    }

    /// The paragraphs this table turns into: one per cell, one per row end.
    pub fn paragraphs(&self) -> Vec<Paragraph> {
        let mut out = Vec::new();
        let mut props = ParaProps::default();
        props.in_table = Some(true);
        for row in &self.rows {
            for cell in row {
                out.push(Paragraph::cell(vec![Run::new(cell.clone())], props.clone()));
            }
            out.push(Paragraph::row_end(&self.centres));
        }
        out
    }
}

/// A document made of some text, a table, and some more text.
pub fn write_table_document(before: &str, table: &Table, after: &str) -> Vec<u8> {
    let mut paragraphs = Vec::new();
    if !before.is_empty() {
        paragraphs.push(Paragraph::text(before));
    }
    paragraphs.extend(table.paragraphs());
    if !after.is_empty() {
        paragraphs.push(Paragraph::text(after));
    }
    write_document(&paragraphs)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn writes_a_document_the_reader_accepts() {
        let bytes = write_minimal("Hello from the writer.\nSecond paragraph 中文.");
        let doc = crate::parse(&bytes).unwrap();
        assert_eq!(doc.fib.n_fib, 0x00C1);
        assert_eq!(doc.fib.table_stream, "1Table");
        assert_eq!(doc.text, "Hello from the writer.\rSecond paragraph 中文.\r");
        assert_eq!(doc.pieces.len(), 1);
        assert!(!doc.pieces[0].compressed);
        let paras = doc.paragraphs();
        assert_eq!(paras, vec!["Hello from the writer.", "Second paragraph 中文."]);
    }

    #[test]
    fn writes_the_formatting_pages_word_insists_on() {
        // Word hangs on a document whose PlcfBteChpx/PlcfBtePapx pairs are
        // empty, so the writer always emits both pages.
        let bytes = write_minimal("One\nTwo\n\u{4e09}");
        let doc = crate::parse(&bytes).unwrap();
        assert_eq!(doc.para_runs.len(), 3);
        assert_eq!(doc.paragraphs(), vec!["One", "Two", "\u{4e09}"]);
        // The paragraph boundaries are UTF-16 code units, not UTF-8 bytes.
        let ends: Vec<u32> = doc.para_runs.iter().map(|r| r.fc_end).collect();
        assert_eq!(ends, vec![0x808, 0x810, 0x814]);
    }

    #[test]
    fn characters_are_counted_not_bytes() {
        let bytes = write_minimal("中文测试字符");
        let doc = crate::parse(&bytes).unwrap();
        // Six CJK characters plus the paragraph mark the writer always adds.
        assert_eq!(doc.text.chars().count(), 7);
        assert_eq!(doc.fib.ccp_text, 7);
        assert_eq!(doc.fib.fc_mac - doc.fib.fc_min, 14);
    }

    #[test]
    fn writes_character_formatting() {
        let bold = CharProps {
            bold: Some(true),
            size_half_points: Some(36),
            color: Some(0x0000_00FF),
            ..CharProps::default()
        };
        let paragraph = Paragraph::styled(
            vec![
                Run::new("plain "),
                Run::styled("bold", bold),
                Run::new(" plain again"),
            ],
            ParaProps {
                justify: Some(1),
                dxa_left: Some(720),
                ..ParaProps::default()
            },
        );
        let bytes = write_document(&[paragraph]);
        let doc = crate::parse(&bytes).unwrap();
        let paras = doc.paragraphs_rich();
        assert_eq!(paras.len(), 1);
        let p = &paras[0];
        assert_eq!(p.text, "plain bold plain again");
        assert_eq!(p.props.justify, Some(1));
        assert_eq!(p.props.dxa_left, Some(720));
        assert_eq!(p.runs.len(), 3);
        assert_eq!(p.runs[0].props.bold, None);
        assert_eq!(p.runs[1].props.bold, Some(true));
        assert_eq!(p.runs[1].props.size_half_points, Some(36));
        assert_eq!(p.runs[1].props.color, Some(0x0000_00FF));
        // The spans must line up with the text the reader sliced.
        let bold_text: String = p
            .runs
            .iter()
            .map(|s| s.text(&p.text))
            .collect::<Vec<_>>()
            .join("|");
        assert_eq!(bold_text, "plain |bold| plain again");
    }

    #[test]
    fn writes_many_paragraphs_across_pages() {
        // 40 paragraphs need two PAPX pages (29 runs each), and the reader has
        // to walk both to see every paragraph.
        let paragraphs: Vec<Paragraph> = (0..40)
            .map(|i| Paragraph::text(format!("paragraph {i}")))
            .collect();
        let bytes = write_document(&paragraphs);
        let doc = crate::parse(&bytes).unwrap();
        assert_eq!(doc.para_runs.len(), 40);
        assert_eq!(doc.paragraphs().len(), 40);
        assert_eq!(doc.paragraphs()[39], "paragraph 39");
    }

    #[test]
    fn writes_a_table_word_can_see() {
        let table = Table::new(&[&["H1", "H2", "H3"], &["a", "b", "c"]]);
        let bytes = write_table_document("Before", &table, "After");
        let doc = crate::parse(&bytes).unwrap();
        let paras = doc.paragraphs_rich();
        // 1 intro + (3 cells + 1 row end) * 2 rows + 1 outro
        assert_eq!(paras.len(), 10);
        let cells: Vec<&str> = paras
            .iter()
            .filter(|p| p.props.in_table == Some(true) && p.props.ttp != Some(true))
            .map(|p| p.text.as_str())
            .collect();
        assert_eq!(cells, vec!["H1", "H2", "H3", "a", "b", "c"]);
        let rows: Vec<&crate::fmt::Paragraph> =
            paras.iter().filter(|p| p.props.ttp == Some(true)).collect();
        assert_eq!(rows.len(), 2);
        assert_eq!(rows[0].props.col_widths, vec![0, 1440, 2880, 4320]);
        assert_eq!(paras[9].text, "After");
    }
}
