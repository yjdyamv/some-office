//! Writing a minimal Word 97 document.
//!
//! This is deliberately the smallest file that is still *valid*: a `WordDocument`
//! stream holding the FIB, the text and the two formatting pages (one CHPX FKP
//! and one PAPX FKP, each with a single run), and a `1Table` stream holding a
//! single-piece CLX, a stylesheet with just `Normal`, and a one-section
//! `PlcfSed`.
//!
//! Word is strict about the formatting pages: a document whose
//! `PlcfBteChpx`/`PlcfBtePapx` pairs are empty hangs the reader, so they are
//! always written even when no character or paragraph property is set.
//!
//! Its purpose is to test the reader against bytes we produced on purpose: if
//! `office-doc` and this writer agree, the FIB/CLX model is at least
//! self-consistent, and `tools/roundtrip.ps1` additionally opens the result in
//! Word so the bytes are not merely self-consistent but *valid*.

/// Byte offset where the text starts inside the `WordDocument` stream.
const TEXT_AT: usize = 0x800;
/// FIB object size: `FibBase` + `fibRgW97` + `fibRgLw97` + `FibRgFcLcb97` + `cswNew`.
const FIB_LEN: usize = 0x380;

/// Build a `.doc` from plain text. `\n` becomes a paragraph mark.
pub fn write_minimal(text: &str) -> Vec<u8> {
    // Paragraph marks in a Word text stream are CR, not LF.
    let text: String = text.replace('\n', "\r");
    let utf16: Vec<u8> = text.encode_utf16().flat_map(|c| c.to_le_bytes()).collect();
    let ccp_text = text.chars().count() as u32;

    // ---- WordDocument: FIB + text -------------------------------------------
    let mut wd = vec![0u8; TEXT_AT];
    let mut fib = vec![0u8; FIB_LEN];
    fib[0x00..0x02].copy_from_slice(&0xA5ECu16.to_le_bytes()); // wIdent
    fib[0x02..0x04].copy_from_slice(&0x00C1u16.to_le_bytes()); // nFib (Word 97)
    fib[0x06..0x08].copy_from_slice(&0x0409u16.to_le_bytes()); // lid (en-US)
    // fComplex (bit 2) is required as soon as a piece table (CLX) exists:
    // without it Word ignores the CLX and reads fcMin..fcMac as one run.
    // fExtChar (bit 12) | fWhichTblStm (bit 9) => UTF-16 text, auxiliary 1Table.
    fib[0x0A..0x0C].copy_from_slice(&0x1204u16.to_le_bytes());
    fib[0x0C..0x0E].copy_from_slice(&0x00BFu16.to_le_bytes()); // nFibBack
    fib[0x18..0x1C].copy_from_slice(&(TEXT_AT as u32).to_le_bytes()); // fcMin
    fib[0x1C..0x20].copy_from_slice(&((TEXT_AT + utf16.len()) as u32).to_le_bytes()); // fcMac
    fib[0x20..0x22].copy_from_slice(&0x000Eu16.to_le_bytes()); // csw
    fib[0x3E..0x40].copy_from_slice(&0x0016u16.to_le_bytes()); // cslw
    let cb_mac = (TEXT_AT + utf16.len()) as u32;
    fib[0x40..0x44].copy_from_slice(&cb_mac.to_le_bytes()); // cbMac
    fib[0x4C..0x50].copy_from_slice(&ccp_text.to_le_bytes()); // ccpText
    fib[0x98..0x9A].copy_from_slice(&0x005Du16.to_le_bytes()); // cbRgFcLcb

    // ---- 1Table: one piece covering the whole text --------------------------
    let mut table: Vec<u8> = Vec::new();
    table.push(0x02); // clxt = Pcdt
    table.extend_from_slice(&16u32.to_le_bytes()); // lcb: 2 CPs + 1 PCD
    table.extend_from_slice(&0u32.to_le_bytes()); // CP[0]
    table.extend_from_slice(&ccp_text.to_le_bytes()); // CP[1]
    table.extend_from_slice(&0u16.to_le_bytes()); // PCD flags
    // fc: bit 30 clear means the piece is UTF-16.
    table.extend_from_slice(&(TEXT_AT as u32).to_le_bytes());
    table.extend_from_slice(&0u16.to_le_bytes()); // prm

    // ---- 1Table: a stylesheet holding the built-in "Normal" style ----------
    // Real files start the STSH with `cbStshi`; the STD records begin at
    // 2 + cbStshi and their names sit at STD + cbSTDBaseInFile.
    let stsh_at = table.len() as u32;
    let cb_stshi: u16 = 18;
    let cb_base: u16 = 10;
    table.extend_from_slice(&cb_stshi.to_le_bytes());
    table.extend_from_slice(&1u16.to_le_bytes()); // cstd
    table.extend_from_slice(&cb_base.to_le_bytes()); // cbSTDBaseInFile
    table.extend_from_slice(&1u16.to_le_bytes()); // flags: names written
    table.extend_from_slice(&1u16.to_le_bytes()); // stiMaxWhenSaved
    table.extend_from_slice(&15u16.to_le_bytes()); // istdMaxFixedWhenSaved
    table.extend_from_slice(&1u16.to_le_bytes()); // nVerBuiltInNamesWhenSaved
    table.extend_from_slice(&[0u8; 6]); // rgftc[3]
    // One STD: cbStd, STDBase (istdBase/cupx/istdNext/bchUpe), then the name.
    let std_name = "Normal";
    let cb_std = (cb_base as usize) + 2 + std_name.len() * 2 + 2;
    table.extend_from_slice(&(cb_std as u16).to_le_bytes());
    table.extend_from_slice(&[0u8; 10]); // STDBase: all zeroes = defaults
    table.extend_from_slice(&(std_name.chars().count() as u16).to_le_bytes());
    for ch in std_name.chars() {
        table.extend_from_slice(&(ch as u16).to_le_bytes());
    }
    table.extend_from_slice(&0u16.to_le_bytes()); // NUL terminator of Xstz
    let stsh_len = table.len() as u32 - stsh_at;

    // ---- 1Table: one section, so Word has a page to lay the text out on ----
    // PlcfSed = CPs (0, ccp) + one SED { fn, fcSepx, fnMpr, fcMpr }.
    let sed_at = table.len() as u32;
    table.extend_from_slice(&0u32.to_le_bytes()); // CP[0]
    table.extend_from_slice(&ccp_text.to_le_bytes()); // CP[1]
    table.extend_from_slice(&0u16.to_le_bytes()); // fn
    table.extend_from_slice(&0xFFFF_FFFFu32.to_le_bytes()); // fcSepx: no SEPX
    table.extend_from_slice(&0u16.to_le_bytes()); // fnMpr
    table.extend_from_slice(&0xFFFF_FFFFu32.to_le_bytes()); // fcMpr
    let sed_len = table.len() as u32 - sed_at;

    // FibRgFcLcb97 index 33 is the CLX; its pair starts at 0x9A + 33 * 8.
    let clx_len: u32 = 5 + 16;
    let clx_pair = 0x9A + 33 * 8;
    fib[clx_pair..clx_pair + 4].copy_from_slice(&0u32.to_le_bytes()); // fcClx
    fib[clx_pair + 4..clx_pair + 8].copy_from_slice(&clx_len.to_le_bytes());
    // Index 0/1 are StshfOrig/Stshf: the stylesheet, written once and shared.
    fib[0x9A..0x9E].copy_from_slice(&stsh_at.to_le_bytes());
    fib[0x9E..0xA2].copy_from_slice(&stsh_len.to_le_bytes());
    fib[0xA2..0xA6].copy_from_slice(&stsh_at.to_le_bytes());
    fib[0xA6..0xAA].copy_from_slice(&stsh_len.to_le_bytes());
    // Index 6 is PlcfSed: the section table Word needs to paginate the text.
    let sed_pair = 0x9A + 6 * 8;
    fib[sed_pair..sed_pair + 4].copy_from_slice(&sed_at.to_le_bytes());
    fib[sed_pair + 4..sed_pair + 8].copy_from_slice(&sed_len.to_le_bytes());

    wd[..FIB_LEN].copy_from_slice(&fib);
    wd.extend_from_slice(&utf16);

    // ---- WordDocument: one CHPX FKP and one PAPX FKP -----------------------
    // A page is 512 bytes and its number is simply `offset / 512`.
    let text_end = (TEXT_AT + utf16.len()) as u32;
    let first_page = text_end.div_ceil(512) as usize;
    let papx_page = first_page;
    let chpx_page = first_page + 1;
    let fc_min = TEXT_AT as u32;
    let fc_mac = text_end;
    wd.resize((chpx_page + 1) * 512, 0);
    // cbMac is the size of the whole WordDocument stream, formatting pages included.
    fib[0x40..0x44].copy_from_slice(&(wd.len() as u32).to_le_bytes());

    // PAPX FKP: one run per paragraph, each PAPX holding `istd = 0` and no
    // sprms. A run ends at its paragraph mark, so the last `rgfc` is `fcMac`.
    // A page holds at most 29 runs (4 bytes of `rgfc` each + 13 bytes of `BX`).
    // `fc` is a byte offset into the UTF-16 stream, so count UTF-16 code
    // units: a CJK character is 3 bytes in UTF-8 but 2 bytes here.
    let mut marks: Vec<u32> = Vec::new();
    let mut units: u32 = 0;
    for ch in text.chars() {
        units += ch.len_utf16() as u32;
        if ch == '\r' {
            marks.push(TEXT_AT as u32 + units * 2);
        }
    }
    if marks.last() != Some(&fc_mac) {
        marks.push(fc_mac);
    }
    if marks.len() > 29 {
        // Too many paragraphs for a single page: fall back to one run so the
        // file stays valid, and let Word split them on the marks itself.
        marks = vec![fc_mac];
    }
    let runs = marks.len();
    let papx = papx_page * 512;
    let page = &mut wd[papx..papx + 512];
    page[0..4].copy_from_slice(&fc_min.to_le_bytes());
    for (i, mark) in marks.iter().enumerate() {
        let at = (i + 1) * 4;
        page[at..at + 4].copy_from_slice(&mark.to_le_bytes());
    }
    for i in 0..runs {
        // BX[i] starts right after the `rgfc` array and is 13 bytes wide:
        // a one-byte offset (in two-byte units) plus a 12-byte `PHE`.
        page[(runs + 1) * 4 + i * 13] = 0x80 + (i as u8) * 2;
    }
    page[511] = runs as u8;
    for i in 0..runs {
        let at = 0x100 + i * 4;
        page[at] = 1; // cb: one two-byte word follows
        page[at + 1] = 0; // istd = 0 (Normal)
    }

    // CHPX FKP: one run with no character properties (rgb[0] == 0 means none).
    let chpx = chpx_page * 512;
    let page = &mut wd[chpx..chpx + 512];
    page[0..4].copy_from_slice(&fc_min.to_le_bytes());
    page[4..8].copy_from_slice(&fc_mac.to_le_bytes());
    page[511] = 1;

    // ---- 1Table: the BTE tables that point at those pages ------------------
    let chpx_plc_at = table.len() as u32;
    table.extend_from_slice(&fc_min.to_le_bytes());
    table.extend_from_slice(&fc_mac.to_le_bytes());
    table.extend_from_slice(&(chpx_page as u32).to_le_bytes());
    let chpx_plc_len = table.len() as u32 - chpx_plc_at;
    let papx_plc_at = table.len() as u32;
    table.extend_from_slice(&fc_min.to_le_bytes());
    table.extend_from_slice(&fc_mac.to_le_bytes());
    table.extend_from_slice(&(papx_page as u32).to_le_bytes());
    let papx_plc_len = table.len() as u32 - papx_plc_at;

    // Index 12 is PlcfBteChpx, 13 is PlcfBtePapx.
    fib[0x0FA..0x0FE].copy_from_slice(&chpx_plc_at.to_le_bytes());
    fib[0x0FE..0x102].copy_from_slice(&chpx_plc_len.to_le_bytes());
    fib[0x102..0x106].copy_from_slice(&papx_plc_at.to_le_bytes());
    fib[0x106..0x10A].copy_from_slice(&papx_plc_len.to_le_bytes());

    wd[..FIB_LEN].copy_from_slice(&fib);

    office_cfb::write(&[
        office_cfb::OutputStream::new("WordDocument", &wd),
        office_cfb::OutputStream::new("1Table", &table),
    ])
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn writes_a_document_the_reader_accepts() {
        let text = "Hello from the writer.\nSecond paragraph 中文.";
        let bytes = write_minimal(text);
        let doc = crate::parse(&bytes).unwrap();
        assert_eq!(doc.fib.n_fib, 0x00C1);
        assert_eq!(doc.fib.table_stream, "1Table");
        assert_eq!(doc.text, text.replace('\n', "\r"));
        assert_eq!(doc.pieces.len(), 1);
        assert!(!doc.pieces[0].compressed);
        let paras = doc.paragraphs();
        assert_eq!(paras, vec!["Hello from the writer.", "Second paragraph 中文."]);
    }

    #[test]
    fn writes_the_formatting_pages_word_insists_on() {
        // Word hangs on a document whose PlcfBteChpx/PlcfBtePapx pairs are
        // empty, so the writer always emits one CHPX page and one PAPX page
        // with a run per paragraph. Walking those pages back is the only way to
        // tell "the file has them" from "the file claims to have them".
        let bytes = write_minimal("One
Two
\u{4e09}
");
        let doc = crate::parse(&bytes).unwrap();
        assert_eq!(doc.para_runs.len(), 3);
        assert_eq!(doc.char_runs.len(), 1);
        assert_eq!(doc.paragraphs(), vec!["One", "Two", "\u{4e09}"]);
        // The paragraph boundaries are UTF-16 code units, not UTF-8 bytes.
        let ends: Vec<u32> = doc.para_runs.iter().map(|r| r.fc_end).collect();
        assert_eq!(ends, vec![0x808, 0x810, 0x814]);
    }

    #[test]
    fn characters_are_counted_not_bytes() {
        // 6 CJK chars are 12 UTF-16 bytes but 6 characters.
        let bytes = write_minimal("中文测试字符");
        let doc = crate::parse(&bytes).unwrap();
        assert_eq!(doc.text.chars().count(), 6);
        assert_eq!(doc.fib.ccp_text, 6);
        assert_eq!(doc.fib.fc_mac - doc.fib.fc_min, 12);
    }
}
