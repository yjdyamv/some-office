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

/// A single picture, stored in the `Data` stream and referenced from a `CHPX`
/// through `sprmCPicLocation`.
///
/// The bytes are the complete image file (PNG, JPEG, ...) exactly as it should
/// be extracted again; the size is what Word uses to lay it out, in twips
/// (1/1440 inch -- a 16x16 pixel image at 96 dpi is 240x240).
#[derive(Clone, Debug)]
pub struct InlinePicture {
    pub data: Vec<u8>,
    pub width_twips: u16,
    pub height_twips: u16,
}

impl InlinePicture {
    pub fn new(data: impl Into<Vec<u8>>, width_twips: u16, height_twips: u16) -> Self {
        InlinePicture { data: data.into(), width_twips, height_twips }
    }
}

/// A `Data`-stream picture taken from a real Word file (one 16x16 PNG), used as
/// the template for the `PICF`/`OfficeArt` wrapper: Word validates that wrapper
/// far more strictly than it documents it, so the writer keeps its bytes and
/// patches only the sizes and the image itself.
const PICTURE_TEMPLATE: &[u8] = include_bytes!("../templates/picture-data.bin");

/// End of the `PICF` + `OfficeArtSpContainer` prefix inside
/// [`PICTURE_TEMPLATE`]; the `OfficeArtFBSE` record starts here.
const TEMPLATE_FBSE: usize = 0xD6;
/// Offset of the BLIP record inside [`PICTURE_TEMPLATE`].
const TEMPLATE_BLIP: usize = 0x102;

/// Byte offset where the text starts inside the `WordDocument` stream.
const TEXT_AT: usize = 0x800;
/// FIB object size: `FibBase` + `fibRgW97` + `fibRgLw97` + `FibRgFcLcb97` + `cswNew`.
const FIB_LEN: usize = 0x380;
/// A formatting page is one sector; the run count lives in the last byte.
const PAGE: usize = 512;

/// The length of `DopBase`, the part of the `DOP` the reader always sees.
///
/// Only written when a section needs the odd/even header switch; every other
/// field is left at zero, which is what Word fills in for a document that has
/// no `DOP` at all.
const DOP_LEN: usize = 84;

/// A stretch of text that shares one set of character properties.
#[derive(Clone, Debug, Default)]
pub struct Run {
    pub text: String,
    pub props: CharProps,
    /// An inline picture: the run's text becomes the picture mark (`\x01`) and
    /// its `CHPX` gains `sprmCPicLocation` + `sprmCFSpec`.
    pub picture: Option<InlinePicture>,
}

impl Run {
    pub fn new(text: impl Into<String>) -> Self {
        Run {
            text: text.into(),
            props: CharProps::default(),
            picture: None,
        }
    }

    /// An inline picture; its size is in twips (a 96-dpi pixel is 15 twips).
    pub fn picture(picture: InlinePicture) -> Self {
        Run {
            text: String::from('\u{1}'),
            props: CharProps::default(),
            picture: Some(picture),
        }
    }

    pub fn styled(text: impl Into<String>, props: CharProps) -> Self {
        Run {
            text: text.into(),
            props,
            picture: None,
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
        let props = ParaProps {
            in_table: Some(true),
            ttp: Some(true),
            col_widths: widths.to_vec(),
            ..ParaProps::default()
        };
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

/// Append plain header-story text, one paragraph per carriage return.
/// Make the paragraph runs cover `fc_min..fc_mac` without holes.
fn close_gaps(para_ends: &mut [(u32, u32, ParaProps)], fc_min: u32, fc_mac: u32) {
    if para_ends.is_empty() {
        return;
    }
    para_ends[0].0 = fc_min;
    for i in 1..para_ends.len() {
        if para_ends[i].0 != para_ends[i - 1].1 {
            para_ends[i].0 = para_ends[i - 1].1;
        }
    }
    let last = para_ends.len() - 1;
    if para_ends[last].1 < fc_mac {
        para_ends[last].1 = fc_mac;
    }
}

fn push_story(
    text: &str,
    utf16: &mut Vec<u8>,
    spans: &mut Vec<(u32, u32, CharProps)>,
    para_ends: &mut Vec<(u32, u32, ParaProps)>,
) {
    let mut para_start = TEXT_AT as u32 + utf16.len() as u32;
    for ch in text.chars() {
        let start = TEXT_AT as u32 + utf16.len() as u32;
        utf16.extend_from_slice(&(ch as u16).to_le_bytes());
        let end = TEXT_AT as u32 + utf16.len() as u32;
        // Field characters are not text: Word only treats 0x13/0x14/0x15 as
        // field begin/separator/end when `sprmCFSpec` marks them special
        // (`docs/doc.md` §14.12).
        let props = if is_field_character(ch) {
            CharProps { spec: Some(true), ..CharProps::default() }
        } else {
            CharProps::default()
        };
        spans.push((start, end, props));
        if ch == '\r' {
            para_ends.push((para_start, end, ParaProps::default()));
            para_start = end;
        }
    }
}

/// The characters that bracket a field: begin, separator, end.
fn is_field_character(ch: char) -> bool {
    matches!(ch, '\u{13}' | '\u{14}' | '\u{15}')
}

/// A field as it sits in the text stream: `code` between begin and separator,
/// `cached` as the result Word shows until the field is updated.
pub fn field(code: &str, cached: &str) -> String {
    format!("\u{13} {code} \u{14}{cached}\u{15}")
}

/// The `PAGE` field: the page number, showing `cached` until Word updates it.
pub fn page_field(cached: &str) -> String {
    field("PAGE", cached)
}

/// The `NUMPAGES` field: the number of pages in the document.
pub fn numpages_field(cached: &str) -> String {
    field("NUMPAGES", cached)
}

/// Split the runs so every field character gets a run of its own, marked with
/// `sprmCFSpec`. `lay_out` copies a paragraph's runs as they came in, so a run
/// can cover "Page " plus the field begin; Word only recognises the field when
/// the bracket character itself is special.
fn split_field_characters(
    spans: Vec<(u32, u32, CharProps)>,
    utf16: &[u8],
) -> Vec<(u32, u32, CharProps)> {
    let char_at = |fc: u32| -> char {
        let at = (fc - TEXT_AT as u32) as usize;
        let unit = u16::from_le_bytes([utf16[at], utf16[at + 1]]);
        char::from_u32(u32::from(unit)).unwrap_or('\u{0}')
    };
    let mut out = Vec::with_capacity(spans.len());
    for (start, end, props) in spans {
        let mut at = start;
        while at < end {
            let field = is_field_character(char_at(at));
            let mut stop = at + 2;
            while stop < end && is_field_character(char_at(stop)) == field {
                stop += 2;
            }
            let mut props = props.clone();
            if field {
                props.spec = Some(true);
            }
            out.push((at, stop.min(end), props));
            at = stop;
        }
    }
    out
}

/// Lay `paragraphs` out in the text stream, recording runs and paragraph marks.
fn lay_out(
    paragraphs: &[Paragraph],
    utf16: &mut Vec<u8>,
    spans: &mut Vec<(u32, u32, CharProps)>,
    para_ends: &mut Vec<(u32, u32, ParaProps)>,
    next_picture: &mut usize,
    picture_offsets: &[u32],
) {
    for paragraph in paragraphs {
        let para_start = TEXT_AT as u32 + utf16.len() as u32;
        let mut runs: Vec<(u32, u32, CharProps)> = Vec::new();
        for run in &paragraph.runs {
            let start = TEXT_AT as u32 + utf16.len() as u32;
            for ch in run.text.chars() {
                utf16.extend_from_slice(&(ch as u16).to_le_bytes());
            }
            let mut props = run.props.clone();
            if run.picture.is_some() {
                props.pic_location = Some(picture_offsets[*next_picture]);
                props.spec = Some(true);
                *next_picture += 1;
            }
            runs.push((start, TEXT_AT as u32 + utf16.len() as u32, props));
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
}

/// A page header and footer, laid out with the same paragraph model as the body.
///
/// Word keeps both in the "header subdocument": their text follows the body
/// text in the same stream and `PlcfHdd` records where each slot starts
/// (`docs/doc.md` section 8). Only the odd-page slots are filled, which is what
/// Word shows when "different first page" and "different odd and even" are off.
#[derive(Clone, Debug, Default)]
pub struct HeaderFooter<'a> {
    pub header: &'a [Paragraph],
    pub footer: &'a [Paragraph],
}

/// The character that stands for a footnote reference mark.
///
/// Word stores the mark as `0x02` when the note is auto-numbered; the number
/// itself is computed at layout time, so documents that autosave a drawn number
/// still carry `0x02` in the text stream.
pub const FOOTNOTE_MARK: char = '\u{2}';

/// The mark a caller puts where an *endnote* is referenced.
///
/// Word stores both kinds with the same `0x02` character and tells them apart
/// only by which PLC lists the position, so the writer needs a different
/// character in the source text to know which list a mark belongs to. It is
/// rewritten to `0x02` before the text stream is laid out, so the file never
/// contains this character (`samples/ms/both.doc`, written by Word, holds a
/// footnote and an endnote and shows `0x02` twice).
pub const ENDNOTE_MARK: char = '\u{5}';

/// Build a `.doc` from paragraphs that carry character and paragraph properties.
pub fn write_document(paragraphs: &[Paragraph]) -> Vec<u8> {
    write_document_with(paragraphs, &HeaderFooter::default())
}

/// The three stories a `.doc` can carry, in one value.
///
/// They are bundled because the text stream holds them back to back in a fixed
/// order, so writing them one function at a time would let a caller produce a
/// file whose counts disagree with its PLCs.
#[derive(Clone, Debug, Default)]
pub struct Documents<'a> {
    /// Header and footer paragraphs: the ones every page uses.
    pub header: &'a [Paragraph],
    pub footer: &'a [Paragraph],
    /// Header and footer for the first page, when the section asks for a
    /// different one (Word's "Different first page").
    pub first_header: &'a [Paragraph],
    pub first_footer: &'a [Paragraph],
    /// Header and footer for even-numbered pages (Word's "Different odd and
    /// even pages").
    pub even_header: &'a [Paragraph],
    pub even_footer: &'a [Paragraph],
    /// Each footnote's own paragraphs, in the order their marks appear.
    pub notes: &'a [Vec<Paragraph>],
    /// Each endnote's own paragraphs, in the order their marks appear.
    ///
    /// Endnotes use the same `0x02` mark as footnotes: only the PLC pair and
    /// the `ccpEdn` count tell the two apart, so a document may carry both as
    /// long as each mark is matched by one note in the story its kind belongs
    /// to. The text stream order decides the counts, so endnotes always come
    /// after the header story.
    pub endnotes: &'a [Vec<Paragraph>],
    /// The page the section is printed on, or `None` for the writer's default
    /// (A4 with 1800-twip side margins).
    pub page: Option<PageSetup>,
}

impl<'a> Documents<'a> {
    pub fn new(header: &'a [Paragraph], footer: &'a [Paragraph]) -> Self {
        Documents {
            header,
            footer,
            first_header: &[],
            first_footer: &[],
            even_header: &[],
            even_footer: &[],
            notes: &[],
            endnotes: &[],
            page: None,
        }
    }

    pub fn notes(mut self, notes: &'a [Vec<Paragraph>]) -> Self {
        self.notes = notes;
        self
    }

    pub fn endnotes(mut self, endnotes: &'a [Vec<Paragraph>]) -> Self {
        self.endnotes = endnotes;
        self
    }

    pub fn page_setup(mut self, page: PageSetup) -> Self {
        self.page = Some(page);
        self
    }

    /// Use a different header and footer on the first page.
    ///
    /// Word stores this as `sprmSFTitlePage` in the section properties, so the
    /// section has to be written with that sprm for the first-page slots of
    /// `PlcfHdd` to be read at all.
    pub fn first_page(
        mut self,
        header: &'a [Paragraph],
        footer: &'a [Paragraph],
    ) -> Self {
        self.first_header = header;
        self.first_footer = footer;
        self
    }

    /// Use a different header and footer on even-numbered pages.
    ///
    /// The switch for this one is not a section property but the document-wide
    /// `fFacingPages` bit in the `DOP`, so the writer has to emit a `DOP` for
    /// it to take effect.
    pub fn even_pages(
        mut self,
        header: &'a [Paragraph],
        footer: &'a [Paragraph],
    ) -> Self {
        self.even_header = header;
        self.even_footer = footer;
        self
    }
}

pub use crate::PageSetup;

/// Build a `.doc` with a repeating header and/or footer.
pub fn write_document_with(paragraphs: &[Paragraph], extras: &HeaderFooter<'_>) -> Vec<u8> {
    write_document_and(paragraphs, &Documents::new(extras.header, extras.footer))
}

/// Build a `.doc` with a header, a footer and footnotes in one call.
pub fn write_document_and(paragraphs: &[Paragraph], extras: &Documents<'_>) -> Vec<u8> {
    // Slot order, measured on Word's own files: even-page header, the header
    // every page uses, even-page footer, the footer every page uses, then the
    // first page's pair. Only the pairs the caller filled in are offered: a
    // slot that is `None` stays empty, and its flag is not set.
    let mut slots: [Option<&[Paragraph]>; 6] = [None; 6];
    if !extras.even_header.is_empty() || !extras.even_footer.is_empty() {
        slots[0] = Some(extras.even_header);
        slots[2] = Some(extras.even_footer);
    }
    slots[1] = Some(extras.header);
    slots[3] = Some(extras.footer);
    if !extras.first_header.is_empty() || !extras.first_footer.is_empty() {
        slots[4] = Some(extras.first_header);
        slots[5] = Some(extras.first_footer);
    }
    write_document_all_with(paragraphs, &slots, extras.notes, extras.endnotes, extras.page)
}

/// The stories of a `.doc`, in the order they appear in the text stream.
///
/// The reader finds a story by walking the `ccp*` counts in `fibRgLw97`, so the
/// stream order *is* the contract: body, footnotes, header, macros, annotations,
/// endnotes, textboxes. Only the first three are written here.
struct Builder<'a> {
    paragraphs: &'a [Paragraph],
    slots: &'a [Option<&'a [Paragraph]>; 6],
    notes: &'a [Vec<Paragraph>],
    endnotes: &'a [Vec<Paragraph>],
    page: Option<PageSetup>,
    /// `sprmSFTitlePage`: the section uses the first-page header pair.
    title_page: bool,
    /// `DOP.fFacingPages`: even pages get their own header pair.
    facing_pages: bool,
}

/// Build a `.doc` placing paragraphs into arbitrary header-story slots, which
/// is only useful for probing which slot Word shows as the header.
pub fn write_document_slots(
    paragraphs: &[Paragraph],
    slots: &[Option<&[Paragraph]>; 6],
) -> Vec<u8> {
    write_document_all(paragraphs, slots, &[], &[])
}

/// Build a `.doc` with footnotes.
///
/// `documents.notes` holds each note's own paragraphs, in the order their
/// reference marks appear in `paragraphs` (marks are `Run`s whose text is
/// [`FOOTNOTE_MARK`]).
pub fn write_document_with_footnotes(
    paragraphs: &[Paragraph],
    documents: &Documents<'_>,
) -> Vec<u8> {
    write_document_and(paragraphs, documents)
}

/// The one entry point that knows about every story.
pub fn write_document_full(
    paragraphs: &[Paragraph],
    slots: &[Option<&[Paragraph]>; 6],
    notes: &[Vec<Paragraph>],
) -> Vec<u8> {
    write_document_all(paragraphs, slots, notes, &[])
}

/// The entry point that also knows about endnotes.
pub fn write_document_all(
    paragraphs: &[Paragraph],
    slots: &[Option<&[Paragraph]>; 6],
    notes: &[Vec<Paragraph>],
    endnotes: &[Vec<Paragraph>],
) -> Vec<u8> {
    write_document_all_with(paragraphs, slots, notes, endnotes, None)
}

/// [`write_document_all`] with a page setup for the section.
pub fn write_document_all_with(
    paragraphs: &[Paragraph],
    slots: &[Option<&[Paragraph]>; 6],
    notes: &[Vec<Paragraph>],
    endnotes: &[Vec<Paragraph>],
    page: Option<PageSetup>,
) -> Vec<u8> {
    // The flags follow from the slots the caller filled in: the first-page
    // pair lives in slots 4 and 5, the even-page pair in 0 and 2.
    let title_page = slots[4].is_some() || slots[5].is_some();
    let facing_pages = slots[0].is_some() || slots[2].is_some();
    let builder = Builder { paragraphs, slots, notes, endnotes, page, title_page, facing_pages };
    builder.build()
}

impl Builder<'_> {
    fn build(&self) -> Vec<u8> {
    // ---- pictures: the `Data` stream is built first so each run can be given
    // the offset of its `PICF` while the text is laid out ---------------------
    let mut data_stream: Vec<u8> = Vec::new();
    let mut picture_offsets: Vec<u32> = Vec::new();
    for paragraph in self.paragraphs {
        for run in &paragraph.runs {
            if let Some(picture) = &run.picture {
                picture_offsets.push(data_stream.len() as u32);
                append_picture(&mut data_stream, picture);
            }
        }
    }

    // ---- lay the text out and record where every run starts and ends --------
    let mut utf16: Vec<u8> = Vec::new();
    let mut next_picture = 0usize;
    let mut spans: Vec<(u32, u32, CharProps)> = Vec::new();
    let mut para_ends: Vec<(u32, u32, ParaProps)> = Vec::new();
    lay_out(
        self.paragraphs,
        &mut utf16,
        &mut spans,
        &mut para_ends,
        &mut next_picture,
        &picture_offsets,
    );
    let ccp_text = utf16.len() as u32 / 2;
    // The footnote story comes next, because the reader walks the `ccp*`
    // counts in stream order. Every note starts with its own mark character and
    // ends with a paragraph mark; the story as a whole ends one character after
    // the last note's CP, and `ccpFtn` stops two short of that, exactly like
    // `ccpHdd` does. `samples/ms/footnote2.doc` (Word 2016, one note) reads
    // `PlcffndTxt` as `[0, 11]` with `ccpFtn = 12`, so the last CP is 13.
    // The marks in the body keep their positions -- a mark is one character, so
    // inserting it shifts every later CP by exactly one.
    let mut ftn_cps: Vec<u32> = Vec::new();
    let mut ccp_ftn = 0u32;
    // The marks are already in the body: the caller put a `FOOTNOTE_MARK` run
    // wherever the note is referenced, so their CPs are simply where those
    // characters ended up. Counting them is also the check that the caller
    // supplied one note per mark.
    let ref_cps: Vec<u32> = (0..ccp_text)
        .filter(|cp| {
            let at = (*cp as usize) * 2;
            u16::from_le_bytes([utf16[at], utf16[at + 1]]) == FOOTNOTE_MARK as u16
        })
        .collect();
    // Endnote marks arrive as `ENDNOTE_MARK` and leave as `0x02`: the PLC that
    // lists the position is the only thing that says which story it belongs to.
    let end_cps: Vec<u32> = (0..ccp_text)
        .filter(|cp| {
            let at = (*cp as usize) * 2;
            u16::from_le_bytes([utf16[at], utf16[at + 1]]) == ENDNOTE_MARK as u16
        })
        .collect();
    for cp in &end_cps {
        let at = (*cp as usize) * 2;
        utf16[at..at + 2].copy_from_slice(&(FOOTNOTE_MARK as u16).to_le_bytes());
    }
    if !self.notes.is_empty() {
        for note in self.notes {
            ftn_cps.push(utf16.len() as u32 / 2 - ccp_text);
            push_story(&FOOTNOTE_MARK.to_string(), &mut utf16, &mut spans, &mut para_ends);
            lay_out(note, &mut utf16, &mut spans, &mut para_ends, &mut next_picture, &picture_offsets);
        }
        let closing = utf16.len() as u32 / 2 - ccp_text;
        push_story("\r", &mut utf16, &mut spans, &mut para_ends);
        ftn_cps.push(closing);
        // `ccpFtn` is the length of the whole footnote story, i.e. one past the
        // last CP -- Word writes `[0, 13, 50]` with `ccpFtn = 51`. Stopping at
        // the last CP instead loses the story's final paragraph mark, and Word
        // then shows every note one character short.
        ccp_ftn = closing + 1;
    }
    // The header story follows the body. `PlcfHdd` splits it into six
    // document-wide separator slots, then six slots per section, then a closing
    // slot -- the shape Word writes for a one-section document. Within a section
    // the slots are even-page header, primary header, even-page footer, primary
    // footer, first-page header and first-page footer. Every slot ends with at
    // least one paragraph mark, and `ccpHdd` stops two characters before the
    // last CP. The story is only written when a slot holds something: Word
    // accepts a document without one, and leaving it out keeps the file small.
    let mut hdd_cps: Vec<u32> = Vec::new();
    let mut ccp_hdd = 0u32;
    // Word always writes the header story, even for a document with no header:
    // `samples/ms/footnote2.doc` carries a 13-character story of separators and
    // paragraph marks. It is optional for a plain document (and leaving it out
    // keeps such files byte-identical to what this writer produced before
    // headers existed), but a document with footnotes needs it: with
    // `ccpHdd = 0` Word refuses the file, claiming it "cannot read the headers
    // and footers".
    let has_slot = self.slots.iter().any(|slot| matches!(slot, Some(filled) if !filled.is_empty()));
    if has_slot || !self.notes.is_empty() || !self.endnotes.is_empty() {
        // `PlcfHdd` CPs are relative to the start of the header story, which is
        // right after the footnote story.
        let story_start = ccp_text + ccp_ftn;
        hdd_cps.push(0);
        let separators = ["\u{3}\r\r", "\u{4}\r\r", "", "\u{3}\r\r", "\u{4}\r\r", ""];
        for text in separators {
            push_story(text, &mut utf16, &mut spans, &mut para_ends);
            hdd_cps.push(utf16.len() as u32 / 2 - story_start);
        }
        for slot in self.slots {
            match slot {
                Some(filled) if !filled.is_empty() => {
                    lay_out(
                        filled,
                        &mut utf16,
                        &mut spans,
                        &mut para_ends,
                        &mut next_picture,
                        &picture_offsets,
                    );
                    push_story("\r", &mut utf16, &mut spans, &mut para_ends);
                }
                // An empty slot still holds two paragraph marks: Word's own
                // files write `\r\r` there (`samples/ms/header.doc`, slots 6,
                // 8, 10 and 11). A zero-length slot makes Word drop the whole
                // header story, and the document then shows no header at all.
                _ => push_story("\r\r", &mut utf16, &mut spans, &mut para_ends),
            }
            hdd_cps.push(utf16.len() as u32 / 2 - story_start);
        }
        let closing = utf16.len() as u32 / 2 - story_start;
        // Word wants `ccpHdd` two short of the story it wrote, with the closing
        // cut point at the story's end: `samples/ms/header.doc` and
        // `samples/ms/header-note.doc` (both Word 2016) carry a 51-character
        // story, `ccpHdd = 49` and a last cut point of 51. Anything else makes
        // Word open the document with no header at all, which the body text
        // does not show -- only `tools/check_header.ps1` catches it.
        //
        // An endnote story has to start at the `ccp*` sum, which is `ccpHdd`,
        // so those two characters cannot be there when endnotes follow: the
        // stories would overlap and Word then refuses the file outright. In
        // that case the story is written without them, Word shows no header,
        // and `docs/doc.md` records the trade-off.
        if self.endnotes.is_empty() {
            push_story("\r\r\r", &mut utf16, &mut spans, &mut para_ends);
            hdd_cps.push(closing + 3);
        } else {
            push_story("\r", &mut utf16, &mut spans, &mut para_ends);
            hdd_cps.push(closing + 1);
        }
        ccp_hdd = closing + 1;
    }
    // The endnote story comes after the header story, because the reader finds
    // stories by summing the `ccp*` counts in stream order. Its shape mirrors
    // the footnote story exactly: `PlcfendRef` cuts the main text at the marks,
    // `PlcfendTxt` cuts the story itself, and `ccpEdn` is the whole story's
    // length (one past the last cut point). `samples/ms/endnote2.doc` (Word
    // 2016, one note) reads `PlcfendTxt` as `[0, 12, 15]` with `ccpEdn = 13`.
    let mut edn_cps: Vec<u32> = Vec::new();
    let mut ccp_edn = 0u32;
    if !self.endnotes.is_empty() {
        let story_start = utf16.len() as u32 / 2;
        for note in self.endnotes {
            edn_cps.push(utf16.len() as u32 / 2 - story_start);
            push_story(&FOOTNOTE_MARK.to_string(), &mut utf16, &mut spans, &mut para_ends);
            lay_out(note, &mut utf16, &mut spans, &mut para_ends, &mut next_picture, &picture_offsets);
        }
        let closing = utf16.len() as u32 / 2 - story_start;
        push_story("\r", &mut utf16, &mut spans, &mut para_ends);
        edn_cps.push(closing);
        ccp_edn = closing + 1;
    }
    let ccp_total = utf16.len() as u32 / 2;
    let fc_min = TEXT_AT as u32;
    let fc_mac = TEXT_AT as u32 + utf16.len() as u32;
    // Paragraph runs have to tile the text stream with no holes. A note's
    // reference mark is pushed by its own `push_story` call, so the paragraph
    // after it starts one character late, and when a formatting page happens
    // to begin on that character Word hangs on the file with a spinning
    // cursor (`samples/sw-3-both.doc`). Closing the holes also gives the mark
    // to the note's first paragraph, which is where it belongs.
    close_gaps(&mut para_ends, fc_min, fc_mac);

    // ---- formatting pages --------------------------------------------------
    // Each page holds a run array, a per-run offset array and the sprms, all
    // inside one 512-byte sector, so the page split depends on how big the
    // sprms are — not just on how many runs there are.
    let spans = split_field_characters(spans, &utf16);
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
    let chpx_pages = paginate_fit(&chpx_sizes, chpx_array_len);
    let papx_pages = paginate_fit(&papx_sizes, papx_array_len);
    let page_count = chpx_pages.len() + papx_pages.len();
    // A `SEPX` (section properties) sits between the text and the first
    // formatting page. Word reads the section through `PlcfSed.fcSepx`, and a
    // `0xFFFFFFFF` "no SEPX" makes it treat the body as empty.
    // Word's own `SEPX` (measured with `examples/sepxdump.rs` on
    // `samples/ms/page-*.doc`) also carries printer defaults; the four sprms
    // below are the ones that decide what `PageSetup` reports.
    let page = self.page.unwrap_or(PageSetup {
        width: 12240,
        height: 15840,
        left: 1800,
        right: 1800,
        top: None,
        bottom: None,
    });
    let mut sepx: Vec<u8> = Vec::new();
    // Sprms go in ascending order, as Word writes them. `sprmSFTitlePage` is a
    // one-byte boolean that makes the section use the first-page header pair.
    if self.title_page {
        push_sprm(&mut sepx, 0x300A, &[1]); // sprmSFTitlePage
    }
    push_sprm(&mut sepx, 0xB01F, &page.width.to_le_bytes()); // sprmSXaPage
    push_sprm(&mut sepx, 0xB020, &page.height.to_le_bytes()); // sprmSYaPage
    push_sprm(&mut sepx, 0xB021, &page.left.to_le_bytes()); // sprmSDxaLeft
    push_sprm(&mut sepx, 0xB022, &page.right.to_le_bytes()); // sprmSDxaRight
    if let Some(top) = page.top {
        push_sprm(&mut sepx, 0x9023, &top.to_le_bytes()); // sprmSDyaTop
    }
    if let Some(bottom) = page.bottom {
        push_sprm(&mut sepx, 0x9024, &bottom.to_le_bytes()); // sprmSDyaBottom
    }
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
        let mut data = chpx_array_len(runs.len());
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
        let mut data = papx_array_len(paras.len());
        for (i, body) in papx_bodies[chunk.clone()].iter().enumerate() {
            page[(paras.len() + 1) * 4 + i * 13] = (data / 2) as u8;
            page[data] = 0;
            page[data + 1] = (body.len() / 2) as u8;
            page[data + 2..data + 2 + body.len()].copy_from_slice(body);
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

    // CLX: one piece per story, stored as UTF-16. A piece may not span a story
    // boundary, so the body and the header subdocument get their own.
    let mut pieces: Vec<(u32, u32)> = vec![(0, 0)]; // (CP start, offset in fc space)
    let mut story_at = ccp_text;
    if ccp_ftn > 0 {
        pieces.push((story_at, story_at * 2));
        story_at += ccp_ftn;
    }
    if ccp_hdd > 0 {
        pieces.push((story_at, story_at * 2));
        story_at += ccp_hdd;
    }
    if ccp_edn > 0 {
        pieces.push((story_at, story_at * 2));
    }
    let cps: Vec<u32> = pieces.iter().map(|(cp, _)| *cp).chain([ccp_total]).collect();
    table.push(0x02); // clxt = Pcdt
    let lcb = (cps.len() * 4 + pieces.len() * 8) as u32;
    table.extend_from_slice(&lcb.to_le_bytes());
    for cp in &cps {
        table.extend_from_slice(&cp.to_le_bytes());
    }
    for (_, offset) in &pieces {
        table.extend_from_slice(&0u16.to_le_bytes()); // PCD flags
        table.extend_from_slice(&(fc_min + offset).to_le_bytes()); // fc, bit 30 clear
        table.extend_from_slice(&0u16.to_le_bytes()); // prm
    }
    let clx_len: u32 = 5 + lcb;

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

    // One section covering the whole text. Word closes the section at the end
    // of *every* story plus one character whenever the document has a footnote
    // or header story, and at plain `ccpText` when it has neither; `PlcffndRef`
    // and `PlcfHdd` both end on that same CP, so a one-character disagreement
    // between them makes Word refuse the file.
    let sed_end = if ccp_ftn > 0 || ccp_hdd > 0 || ccp_edn > 0 {
        ccp_text + ccp_ftn + ccp_hdd + ccp_edn + 1
    } else {
        ccp_text
    };
    let sed_at = table.len() as u32;
    table.extend_from_slice(&0u32.to_le_bytes()); // CP[0]
    table.extend_from_slice(&sed_end.to_le_bytes()); // CP[1]
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

    // `PlcffndRef`: one CP per mark plus a closing CP, then one `FRD` (2 bytes)
    // per note. A document with no notes keeps the pair at 0/0 so its bytes do
    // not change just because the writer learned about footnotes.
    let fnd_ref_at = if ref_cps.is_empty() { 0 } else { table.len() as u32 };
    let fnd_ref_len = if !ref_cps.is_empty() {
        for cp in &ref_cps {
            table.extend_from_slice(&cp.to_le_bytes());
        }
        // The closing CP is one past every story, and it has to agree with
        // `Plcfsed`'s CP[1] character for character: `samples/ms/footnote.doc`
        // reads `[31, 75, 166]` with the three stories totalling 165. Word
        // blames its header story when the two disagree, so the message it
        // prints for a mismatch here says nothing about this PLC.
        table.extend_from_slice(&sed_end.to_le_bytes());
        for _ in 0..ref_cps.len() {
            // FRD: bit 0 is `fAuto`, which Word sets for a numbered note. A
            // zero here reads as "custom mark" and needs a character that the
            // FrdExtraData would have to supply.
            table.extend_from_slice(&1u16.to_le_bytes());
        }
        table.len() as u32 - fnd_ref_at
    } else {
        0
    };

    // `PlcffndTxt` is a bare CP array relative to the start of the footnote
    // story: one CP per note, then a closing pair. Word writes `[0, 11, 14]`
    // for a one-note story of 12 characters (`lcb = 12`), so the closing pair
    // sits two past `ccpFtn` even though the story has no more characters.
    let fnd_txt_at = if ftn_cps.is_empty() { 0 } else { table.len() as u32 };
    let fnd_txt_len = if !ftn_cps.is_empty() {
        for cp in &ftn_cps {
            table.extend_from_slice(&cp.to_le_bytes());
        }
        // `[0, 11, 14]` for Word's twelve-character story: the closing cut
        // point sits three past the last note's end in both its file and ours.
        table.extend_from_slice(&(ftn_cps.last().copied().unwrap_or(0) + 3).to_le_bytes());
        table.len() as u32 - fnd_txt_at
    } else {
        0
    };

    // `PlcfHdd`: six document-wide separators, six slots per section, then a
    // closing CP, all relative to the start of the header subdocument. A
    // document with no header story keeps the pair at 0/0 so its bytes do not
    // change just because the writer learned about headers.
    let hdd_at = if hdd_cps.is_empty() { 0 } else { table.len() as u32 };
    for cp in &hdd_cps {
        table.extend_from_slice(&cp.to_le_bytes());
    }
    let hdd_len = if hdd_cps.is_empty() { 0 } else { table.len() as u32 - hdd_at };

    // `PlcfendRef`/`PlcfendTxt` are the footnote pair moved to the endnote
    // story: the first cuts the main text at each mark and closes on the same
    // CP `Plcfsed` does, the second cuts the story itself. A document with no
    // endnotes keeps both pairs at 0/0.
    let end_ref_at = if self.endnotes.is_empty() { 0 } else { table.len() as u32 };
    let end_ref_len = if !self.endnotes.is_empty() {
        for cp in &end_cps {
            table.extend_from_slice(&cp.to_le_bytes());
        }
        table.extend_from_slice(&sed_end.to_le_bytes());
        for _ in 0..end_cps.len() {
            table.extend_from_slice(&1u16.to_le_bytes());
        }
        table.len() as u32 - end_ref_at
    } else {
        0
    };
    let end_txt_at = if edn_cps.is_empty() { 0 } else { table.len() as u32 };
    let end_txt_len = if !edn_cps.is_empty() {
        for cp in &edn_cps {
            table.extend_from_slice(&cp.to_le_bytes());
        }
        table.extend_from_slice(&(edn_cps.last().copied().unwrap_or(0) + 3).to_le_bytes());
        table.len() as u32 - end_txt_at
    } else {
        0
    };

    // ---- DOP ---------------------------------------------------------------
    // The document properties carry the odd/even header switch: `fFacingPages`
    // is bit 0 of the first byte, and Word only reads the even-page slots of
    // `PlcfHdd` when it is set. Measured with `examples/dopdump.rs` on
    // `samples/ms/hdr-{firstpage,evenodd}.doc`: the two files differ in exactly
    // that bit (0x20 against 0x21), so it is not a section property.
    // The rest of `DopBase` stays zero: Word accepts a document without a DOP
    // at all, so the defaults it fills in are the ones we want.
    let (dop_at, dop_len) = if self.facing_pages {
        let at = table.len() as u32;
        let mut dop = [0u8; DOP_LEN];
        dop[0] |= 0x01; // fFacingPages
        table.extend_from_slice(&dop);
        (at, DOP_LEN as u32)
    } else {
        (0, 0)
    };

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
    fib[0x50..0x54].copy_from_slice(&ccp_ftn.to_le_bytes()); // ccpFtn
    fib[0x54..0x58].copy_from_slice(&ccp_hdd.to_le_bytes()); // ccpHdd
    // ccpMcr (0x58) and ccpAtn (0x5C) stay zero: no macros, no annotations.
    fib[0x60..0x64].copy_from_slice(&ccp_edn.to_le_bytes()); // ccpEdn
    fib[0x98..0x9A].copy_from_slice(&0x005Du16.to_le_bytes()); // cbRgFcLcb

    // FibRgFcLcb97 pairs: 0/1 stylesheet, 6 section table, 12/13 BTE, 33 CLX.
    fib[0x9A..0x9E].copy_from_slice(&stsh_at.to_le_bytes());
    fib[0x9E..0xA2].copy_from_slice(&stsh_len.to_le_bytes());
    fib[0xA2..0xA6].copy_from_slice(&stsh_at.to_le_bytes());
    fib[0xA6..0xAA].copy_from_slice(&stsh_len.to_le_bytes());
    fib[0xCA..0xCE].copy_from_slice(&sed_at.to_le_bytes());
    fib[0xCE..0xD2].copy_from_slice(&sed_len.to_le_bytes());
    fib[0x0AA..0x0AE].copy_from_slice(&fnd_ref_at.to_le_bytes());
    fib[0x0AE..0x0B2].copy_from_slice(&fnd_ref_len.to_le_bytes());
    fib[0x0B2..0x0B6].copy_from_slice(&fnd_txt_at.to_le_bytes());
    fib[0x0B6..0x0BA].copy_from_slice(&fnd_txt_len.to_le_bytes());
    fib[0x0F2..0x0F6].copy_from_slice(&hdd_at.to_le_bytes());
    fib[0x0F6..0x0FA].copy_from_slice(&hdd_len.to_le_bytes());
    // FibRgFcLcb97 pairs 4/5 are `PlcfendRef`/`PlcfendTxt`.
    // Pairs 46 and 47, not 4 and 5: those are the annotation PLCs.
    fib[0x20A..0x20E].copy_from_slice(&end_ref_at.to_le_bytes());
    fib[0x20E..0x212].copy_from_slice(&end_ref_len.to_le_bytes());
    fib[0x212..0x216].copy_from_slice(&end_txt_at.to_le_bytes());
    fib[0x216..0x21A].copy_from_slice(&end_txt_len.to_le_bytes());
    fib[0x0FA..0x0FE].copy_from_slice(&chpx_plc_at.to_le_bytes());
    fib[0x0FE..0x102].copy_from_slice(&chpx_plc_len.to_le_bytes());
    fib[0x102..0x106].copy_from_slice(&papx_plc_at.to_le_bytes());
    fib[0x106..0x10A].copy_from_slice(&papx_plc_len.to_le_bytes());
    let clx_pair = 0x9A + 33 * 8;
    fib[clx_pair..clx_pair + 4].copy_from_slice(&0u32.to_le_bytes()); // fcClx
    fib[clx_pair + 4..clx_pair + 8].copy_from_slice(&clx_len.to_le_bytes());
    // Pair 31 is `fcDop/lcbDop`; the odd/even header switch lives in it.
    let dop_pair = 0x9A + 31 * 8;
    fib[dop_pair..dop_pair + 4].copy_from_slice(&dop_at.to_le_bytes());
    fib[dop_pair + 4..dop_pair + 8].copy_from_slice(&dop_len.to_le_bytes());

    wd[..FIB_LEN].copy_from_slice(&fib);

    let mut streams = vec![
        office_cfb::OutputStream::new("WordDocument", &wd),
        office_cfb::OutputStream::new("1Table", &table),
    ];
    if !data_stream.is_empty() {
        streams.push(office_cfb::OutputStream::new("Data", &data_stream));
    }
    office_cfb::write(&streams)
    }
}

/// Append one picture to the `Data` stream: `PICF` + `OfficeArt` wrapper +
/// `OfficeArtFBSE` + BLIP.
///
/// The wrapper bytes come from [`PICTURE_TEMPLATE`] because Word checks more
/// than MS-DOC describes; only the sizes, the picture size in twips and the
/// image bytes are ours.
fn append_picture(out: &mut Vec<u8>, picture: &InlinePicture) {
    let start = out.len();
    let blip_len = 17 + picture.data.len(); // uid + tag + image
    let fbse_len = 36 + 8 + blip_len; // header fields + the nested BLIP record
    let total = TEMPLATE_FBSE + 8 + fbse_len;

    // `PICF` + `OfficeArt` wrapper, with the total length and the layout size
    // patched (`dxaGoal`/`dyaGoal` at 0x1C/0x1E).
    let mut picf = PICTURE_TEMPLATE[..TEMPLATE_FBSE].to_vec();
    picf[0..4].copy_from_slice(&(total as u32).to_le_bytes());
    picf[0x1C..0x1E].copy_from_slice(&picture.width_twips.to_le_bytes());
    picf[0x1E..0x20].copy_from_slice(&picture.height_twips.to_le_bytes());
    out.extend_from_slice(&picf);

    // `OfficeArtFBSE`: 36 bytes of fields, then the BLIP record nested inside
    // (`len` covers both). The `uid` and the tag come from the template so the
    // bytes look like the ones Word itself writes.
    let fbse = TEMPLATE_FBSE;
    out.extend_from_slice(&PICTURE_TEMPLATE[fbse..fbse + 2]); // ver_inst
    out.extend_from_slice(&0xF007u16.to_le_bytes()); // kind
    out.extend_from_slice(&(fbse_len as u32).to_le_bytes()); // len
    out.extend_from_slice(&PICTURE_TEMPLATE[fbse + 8..fbse + 24]); // rgbUid
    out.extend_from_slice(&[0xFF, 0x00]); // tag
    out.extend_from_slice(&((8 + blip_len) as u32).to_le_bytes()); // size
    out.extend_from_slice(&1u32.to_le_bytes()); // cRef
    out.extend_from_slice(&0u32.to_le_bytes()); // foDelay
    out.extend_from_slice(&0u32.to_le_bytes()); // unused1
    out.extend_from_slice(&[0x00, 0x00]); // cbName, unused2

    // The BLIP: uid + tag byte + the image, typed by sniffing the magic bytes.
    let kind = match office_blip::PictureKind::sniff(&picture.data) {
        office_blip::PictureKind::Png => 0xF01Eu16,
        office_blip::PictureKind::Jpeg => 0xF01Fu16,
        office_blip::PictureKind::Bmp => 0xF02Au16,
        _ => 0xF01Eu16,
    };
    out.extend_from_slice(&PICTURE_TEMPLATE[TEMPLATE_BLIP..TEMPLATE_BLIP + 2]); // ver_inst
    out.extend_from_slice(&kind.to_le_bytes());
    out.extend_from_slice(&(blip_len as u32).to_le_bytes());
    out.extend_from_slice(&PICTURE_TEMPLATE[TEMPLATE_BLIP + 8..TEMPLATE_BLIP + 25]); // uid + tag
    out.extend_from_slice(&picture.data);

    debug_assert_eq!(out.len() - start, total);
}

/// Round up to an even offset: formatting pages address their data in
/// two-byte units, one byte per entry.
fn align2(at: usize) -> usize {
    (at + 1) & !1
}

/// Bytes a `CHPX` page spends before its data: the run array plus one offset
/// byte per run.
fn chpx_array_len(runs: usize) -> usize {
    align2(4 * (runs + 1) + runs)
}

/// Bytes a `PAPX` page spends before its data: the run array plus a thirteen
/// byte `BX` per run.
fn papx_array_len(paras: usize) -> usize {
    align2(4 * (paras + 1) + 13 * paras)
}

/// Split items into formatting pages. `array(n)` is the space the run array and
/// the offset array take for `n` items, `sizes` is each item's data size. Data
/// follows the array directly, which is what Word does and what keeps a page
/// from splitting at an arbitrary point (`samples/sw-3-both.doc` hung on the
/// fixed data offset this used to have).
fn paginate_fit(sizes: &[usize], array: impl Fn(usize) -> usize) -> Vec<std::ops::Range<usize>> {
    let mut out = Vec::new();
    let mut start = 0usize;
    while start < sizes.len() {
        let mut end = start;
        let mut data = 0usize;
        while end < sizes.len() {
            let count = end - start + 1;
            let next = data + sizes[end];
            if array(count) + next > 511 {
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
    // A picture mark needs both: where the `PICF` lives, and the flag that says
    // "this character is not text".
    if let Some(v) = props.pic_location {
        push_sprm(&mut out, 0x6A03, &v.to_le_bytes());
        push_sprm(&mut out, 0x0855, &[props.spec.unwrap_or(true) as u8]);
    } else if let Some(v) = props.spec {
        push_sprm(&mut out, 0x0855, &[v as u8]);
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
        let props = ParaProps { in_table: Some(true), ..ParaProps::default() };
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
    fn writes_an_inline_picture_the_reader_extracts_again() {
        let png = std::fs::read("../../samples/pixel.png").expect("sample png");
        let paragraphs = vec![
            Paragraph::text("before the image"),
            Paragraph::styled(
                vec![Run::picture(InlinePicture::new(png.clone(), 240, 240))],
                ParaProps::default(),
            ),
            Paragraph::text("after the image"),
        ];
        let bytes = write_document(&paragraphs);
        let doc = crate::parse(&bytes).expect("parses");
        assert_eq!(doc.fib.ccp_text as usize, doc.text.chars().count());
        // The picture mark is a real character in the text stream.
        assert!(doc.full_text.contains('\u{1}'));
        let pictures = doc.pictures();
        assert_eq!(pictures.len(), 1, "one picture");
        assert_eq!(pictures[0].kind, office_blip::PictureKind::Png);
        assert_eq!(pictures[0].data, png, "the image round-trips byte for byte");
        assert_eq!(pictures[0].fc, 0, "the first PICF sits at the start of Data");
    }

    #[test]
    fn writes_two_pictures_at_different_offsets() {
        let png = std::fs::read("../../samples/pixel.png").expect("sample png");
        let paragraphs = vec![
            Paragraph::styled(
                vec![Run::picture(InlinePicture::new(png.clone(), 240, 240))],
                ParaProps::default(),
            ),
            Paragraph::text("between"),
            Paragraph::styled(
                vec![Run::picture(InlinePicture::new(png.clone(), 480, 240))],
                ParaProps::default(),
            ),
        ];
        let doc = crate::parse(&write_document(&paragraphs)).expect("parses");
        let pictures = doc.pictures();
        assert_eq!(pictures.len(), 2);
        assert_ne!(pictures[0].fc, pictures[1].fc, "two separate PICF records");
        assert_eq!(pictures[0].data, png);
        assert_eq!(pictures[1].data, png);
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

    #[test]
    fn header_and_footer_land_in_the_header_story() {
        let body = vec![Paragraph::text("Body paragraph.")];
        let header = vec![Paragraph::text("Draft header")];
        let footer = vec![Paragraph::text("Draft footer")];
        let bytes = write_document_with(
            &body,
            &HeaderFooter { header: &header, footer: &footer },
        );
        let doc = crate::parse(&bytes).unwrap();
        assert_eq!(doc.fib.ccp_text, 16); // "Body paragraph." plus the mark
        assert_ne!(doc.fib.ccp_counts[2], 0, "ccpHdd covers the header story");
        let story = doc.story(crate::Story::Header);
        assert!(story.contains("Draft header\r"), "header story: {story:?}");
        assert!(story.contains("Draft footer\r"), "header story: {story:?}");
        assert_eq!(doc.paragraphs(), vec!["Body paragraph."]);
    }

    #[test]
    fn the_header_story_has_the_shape_word_writes() {
        // Word needs `ccpHdd` to be the story's length minus two and the last
        // `PlcfHdd` cut point to be the story's length; it also wants every
        // slot to hold at least a paragraph mark. Getting any of that wrong
        // makes the document open with no header at all -- the body reads
        // fine, so only `tools/check_header.ps1` notices.
        let body = vec![Paragraph::text("Body paragraph.")];
        let header = vec![Paragraph::text("Draft header")];
        let footer = vec![Paragraph::text("Draft footer")];
        let bytes = write_document_with(&body, &HeaderFooter { header: &header, footer: &footer });
        let doc = crate::parse(&bytes).unwrap();
        let slots = &doc.hdd_slots;
        // Six document-wide separators, six per section, one closing cut point.
        assert_eq!(slots.len(), 14, "slots: {slots:?}");
        assert_eq!(slots[0], 0);
        // A slot's width is the distance to the next cut point.
        assert_eq!(slots[1] - slots[0], 3, "U+0003 then two marks");
        assert_eq!(slots[2] - slots[1], 3, "U+0004 then two marks");
        assert_eq!(slots[3] - slots[2], 0, "the separator slots may stay empty");
        assert_eq!(slots[6] - slots[5], 0, "the separator slots may stay empty");
        assert_eq!(slots[7] - slots[6], 2, "an empty slot still holds two marks");
        assert_eq!(slots[8] - slots[7], 14, "\"Draft header\" plus two marks");
        assert_eq!(slots[9] - slots[8], 2);
        assert_eq!(slots[10] - slots[9], 14, "\"Draft footer\" plus two marks");
        assert_eq!(slots[11] - slots[10], 2);
        assert_eq!(slots[12] - slots[11], 2);
        assert_eq!(*slots.last().unwrap(), doc.fib.ccp_counts[2] + 2);
    }

    #[test]
    fn first_and_even_headers_land_in_their_slots() {
        let body = vec![Paragraph::text("Body.")];
        let header = vec![Paragraph::text("Every-page header")];
        let footer = vec![Paragraph::text("Every-page footer")];
        let first = vec![Paragraph::text("First-page header")];
        let even = vec![Paragraph::text("Even-page header")];
        let documents =
            Documents::new(&header, &footer).first_page(&first, &[]).even_pages(&even, &[]);
        let bytes = write_document_and(&body, &documents);
        let doc = crate::parse(&bytes).expect("parse back");
        assert_eq!(
            doc.header_footer(0),
            Some(("Every-page header".to_string(), "Every-page footer".to_string()))
        );
        assert_eq!(
            doc.first_page_header_footer(0),
            Some(("First-page header".to_string(), String::new()))
        );
        assert_eq!(
            doc.even_page_header_footer(0),
            Some(("Even-page header".to_string(), String::new()))
        );
        // The switches are what make Word look at those slots at all: the
        // first-page one is a section sprm, the even-page one a `DOP` bit.
        let cfb = office_cfb::Cfb::parse(bytes).expect("cfb");
        let wd = cfb.stream("WordDocument").expect("WordDocument");
        assert!(
            wd.windows(3).any(|w| w == [0x0A, 0x30, 0x01]),
            "sprmSFTitlePage is written"
        );
        let table = cfb.stream("1Table").expect("1Table");
        let at = 0x9A + 31 * 8;
        let fc = u32::from_le_bytes([wd[at], wd[at + 1], wd[at + 2], wd[at + 3]]) as usize;
        let lcb = u32::from_le_bytes([wd[at + 4], wd[at + 5], wd[at + 6], wd[at + 7]]) as usize;
        assert_eq!(lcb, 84, "DopBase");
        assert_eq!(table[fc] & 0x01, 0x01, "fFacingPages");
    }

    #[test]
    fn a_document_without_a_header_has_no_header_story() {
        let bytes = write_minimal("Only the body.");
        let doc = crate::parse(&bytes).unwrap();
        assert_eq!(doc.fib.ccp_counts[2], 0);
        assert!(doc.story(crate::Story::Header).is_empty());
    }
    #[test]
    fn footnotes_land_in_the_footnote_story() {
        let body = vec![
            Paragraph::text(format!(
                "Body one comes before the mark.{FOOTNOTE_MARK} and continues after it."
            )),
            Paragraph::text(format!("Body two is short.{FOOTNOTE_MARK}")),
            Paragraph::text("Body three has no note."),
        ];
        let notes = vec![
            vec![Paragraph::text("first note")],
            vec![Paragraph::text("second note, longer than the first")],
        ];
        let bytes = write_document_with_footnotes(&body, &Documents::default().notes(&notes));
        let doc = crate::parse(&bytes).unwrap();
        let main = doc.story(crate::Story::Main);
        let [ccp_text, ccp_ftn, ..] = doc.fib.ccp_counts;
        assert_eq!(main.chars().count() as u32, ccp_text);
        assert_eq!(doc.fib.ccp_text, ccp_text, "the marks count as body characters");
        assert_eq!(main.matches(FOOTNOTE_MARK).count(), 2);
        // Each note is its mark, its text and a paragraph mark; the story ends
        // with one more paragraph mark. No space follows the mark (Word types
        // one itself in its own files, but reads ours without it).
        assert_eq!(ccp_ftn, 1 + 10 + 1 + 1 + 34 + 1 + 1);
        let story = doc.story(crate::Story::Footnote);
        assert_eq!(
            story,
            format!(
                "{mark}first note\r{mark}second note, longer than the first\r\r",
                mark = FOOTNOTE_MARK
            ),
        );
        // `PlcffndRef` lists the marks' own positions and closes one past every
        // story.
        let mut marks: Vec<u32> = main
            .chars()
            .enumerate()
            .filter(|(_, ch)| *ch == FOOTNOTE_MARK)
            .map(|(i, _)| i as u32)
            .collect();
        let hdd = doc.fib.ccp_counts[2];
        // One CP per mark plus a closing CP one past every story, so the header
        // story counts too.
        marks.push(ccp_text + ccp_ftn + hdd + 1);
        assert_eq!(doc.footnote_slots, marks);
        // `PlcffndTxt` is two past the note count: one cut point per note, then
        // the closing pair Word writes (`[0, 13, 50, 53]` for two notes).
        assert_eq!(doc.footnote_text_slots.len(), 4);
        assert_eq!(doc.footnote_text_slots[0], 0);
        let described = doc.footnote_text_slots.windows(2).all(|w| w[0] < w[1]);
        assert!(described, "slots must increase: {:?}", doc.footnote_text_slots);
        assert!(hdd > 0, "a footnote document carries a separator story");
    }

    #[test]
    fn footnote_closing_cps_agree_across_the_three_tables() {
        let body = vec![Paragraph::text(format!("Alpha beta gamma.{FOOTNOTE_MARK}"))];
        let notes = vec![vec![Paragraph::text("one note")]];
        let bytes = write_document_with_footnotes(&body, &Documents::default().notes(&notes));
        let doc = crate::parse(&bytes).unwrap();
        let [ccp_text, ccp_ftn, ccp_hdd, ..] = doc.fib.ccp_counts;
        assert_eq!((ccp_text, ccp_ftn), (19, 11), "body plus its paragraph mark");
        // The note story is mark, text, two paragraph marks; Word would write 12
        // because it also drops a space after the mark.
        assert_eq!(doc.story(crate::Story::Footnote), format!("{FOOTNOTE_MARK}one note\r\r"));
        // `PlcffndRef` ends one past every story, and `Plcfsed` closes on the
        // same value -- Word refuses a file where they differ by even one and
        // then blames the header story.
        assert_eq!(doc.footnote_slots, vec![17, ccp_text + ccp_ftn + ccp_hdd + 1]);
        // `PlcffndTxt` closes two past the note story, like Word's `[0, 11, 14]`
        // for a twelve-character story.
        assert_eq!(doc.footnote_text_slots, vec![0, ccp_ftn - 1, ccp_ftn + 2]);
    }

    #[test]
    fn notes_and_a_header_share_one_stream() {
        let body = vec![Paragraph::text(format!("Body.{FOOTNOTE_MARK}"))];
        let notes = vec![vec![Paragraph::text("a note")]];
        let header = vec![Paragraph::text("Draft header")];
        let footer = vec![Paragraph::text("Draft footer")];
        let documents = Documents::new(&header, &footer).notes(&notes);
        let bytes = write_document_with_footnotes(&body, &documents);
        let doc = crate::parse(&bytes).unwrap();
        let [ccp_text, ccp_ftn, ccp_hdd, ..] = doc.fib.ccp_counts;
        assert!(ccp_text > 0 && ccp_ftn > 0 && ccp_hdd > 0, "{:?}", doc.fib.ccp_counts);
        assert_eq!(doc.story(crate::Story::Main), format!("Body.{FOOTNOTE_MARK}\r"));
        assert!(doc.story(crate::Story::Footnote).contains("a note"));
        assert_eq!(doc.header_footer(0).unwrap(), ("Draft header".into(), "Draft footer".into()));
        // Every PLC closes one past the three stories.
        assert_eq!(doc.footnote_slots.last(), Some(&(ccp_text + ccp_ftn + ccp_hdd + 1)));
        // One piece per story keeps Word reading each at the right place.
        assert_eq!(doc.pieces.len(), 3);
    }

    #[test]
    fn endnotes_land_in_the_endnote_story() {
        let body = vec![
            Paragraph::text(format!("Body{FOOTNOTE_MARK} with a footnote.")),
            Paragraph::text(format!("Body{ENDNOTE_MARK} with an endnote.")),
        ];
        let notes = vec![vec![Paragraph::text("a note")]];
        let endnotes = vec![vec![Paragraph::text("an endnote")]];
        let documents = Documents::new(&[], &[]).notes(&notes).endnotes(&endnotes);
        let bytes = write_document_with_footnotes(&body, &documents);
        let doc = crate::parse(&bytes).unwrap();
        let [ccp_text, ccp_ftn, ccp_hdd, _mcr, _atn, ccp_edn, ..] = doc.fib.ccp_counts;
        assert!(ccp_ftn > 0 && ccp_hdd > 0 && ccp_edn > 0, "{:?}", doc.fib.ccp_counts);
        // The source text carries a different character for each kind, but the
        // stream carries `0x02` for both: only the PLC says which is which.
        let main = doc.story(crate::Story::Main);
        assert_eq!(main.matches(FOOTNOTE_MARK).count(), 2, "{main:?}");
        assert!(!main.contains(ENDNOTE_MARK));
        assert_eq!(doc.story(crate::Story::Footnote), format!("{FOOTNOTE_MARK}a note\r\r"));
        assert_eq!(doc.story(crate::Story::Endnote), format!("{FOOTNOTE_MARK}an endnote\r\r"));
        let closing = ccp_text + ccp_ftn + ccp_hdd + ccp_edn + 1;
        assert_eq!(doc.footnote_slots, vec![4, closing]);
        // The first paragraph is `Body`, the mark and seventeen characters of
        // text plus its paragraph mark: the second mark lands at CP 27.
        assert_eq!(doc.endnote_slots, vec![27, closing]);
        assert_eq!(doc.endnote_text_slots, vec![0, ccp_edn - 1, ccp_edn + 2]);
        assert_eq!(doc.pieces.len(), 4, "one piece per story");
    }

    #[test]
    fn page_setup_lands_in_the_section_properties() {
        // The sprms are `sprmSXaPage` 0xB01F, `sprmSYaPage` 0xB020,
        // `sprmSDxaLeft` 0xB021, `sprmSDxaRight` 0xB022 and, when the margins
        // are given, `sprmSDyaTop` 0x9023 and `sprmSDyaBottom` 0x9024 -- the
        // codes Word writes in `samples/ms/page-*.doc`.
        let body = vec![Paragraph::text("A page.")];
        let page = PageSetup::letter().landscape().margins(1440, 720, 1080, 900);
        let bytes = write_document_with_footnotes(&body, &Documents::default().page_setup(page));
        let doc = crate::parse(&bytes).unwrap();
        assert_eq!(doc.page, Some(page), "landscape swaps the page's sides");
        assert_eq!(page.width, 15840);
        assert_eq!(page.height, 12240);
        // A document without a page setup keeps the writer's default and says
        // nothing about the top and bottom margins.
        let plain = write_minimal("A page.");
        let doc = crate::parse(&plain).unwrap();
        assert_eq!(
            doc.page,
            Some(PageSetup { width: 12240, height: 15840, left: 1800, right: 1800, top: None, bottom: None })
        );
    }

    #[test]
    fn an_endnote_story_starts_where_the_counts_put_it() {
        // Word wants the header story two characters longer than `ccpHdd`,
        // while an endnote story has to begin at the `ccp*` sum, which is
        // `ccpHdd`. The two cannot both hold, so a document with endnotes gets
        // a header story that stops at `ccpHdd`: the file stays valid and the
        // endnotes read back, but Word shows no header. `docs/doc.md` records
        // the trade-off; this test pins the shape so it cannot drift silently.
        let body = vec![Paragraph::text(format!("Body{ENDNOTE_MARK}."))];
        let header = vec![Paragraph::text("Draft header")];
        let endnotes = vec![vec![Paragraph::text("an endnote")]];
        let documents = Documents::new(&header, &[]).endnotes(&endnotes);
        let bytes = write_document_with_footnotes(&body, &documents);
        let doc = crate::parse(&bytes).unwrap();
        let slots = &doc.hdd_slots;
        assert_eq!(*slots.last().unwrap(), doc.fib.ccp_counts[2]);
        assert!(doc.story(crate::Story::Header).contains("Draft header"));
        assert_eq!(doc.story(crate::Story::Endnote), format!("{FOOTNOTE_MARK}an endnote\r\r"));
    }

    #[test]
    fn endnote_fib_pairs_are_forty_six_and_forty_seven() {
        let body = vec![Paragraph::text(format!("Body{ENDNOTE_MARK}."))];
        let endnotes = vec![vec![Paragraph::text("only")]];
        let documents = Documents::new(&[], &[]).endnotes(&endnotes);
        let bytes = write_document_with_footnotes(&body, &documents);
        let doc = crate::parse(&bytes).unwrap();
        // Pairs 4 and 5 are the annotation PLCs; an endnote written there is
        // invisible to Word, which reports the file as damaged.
        assert!(!doc.endnote_slots.is_empty());
        assert!(!doc.endnote_text_slots.is_empty());
        assert!(doc.footnote_slots.is_empty(), "no footnotes were written");
    }

    #[test]
    fn field_characters_are_marked_special() {
        let text = field("PAGE", "1");
        assert_eq!(text, "\u{13} PAGE \u{14}1\u{15}");
        let paragraphs = [Paragraph::text(&text)];
        let bytes = write_document(&paragraphs);
        // `sprmCFSpec` with the value 1, in some CHPX.
        assert!(
            bytes.windows(3).any(|w| w == [0x55, 0x08, 0x01]),
            "no sprmCFSpec in the document"
        );
    }

    #[test]
    fn close_gaps_fills_holes_and_pins_the_ends() {
        // A note's reference mark is pushed on its own, so the paragraph after
        // it starts one character late. A `rgfc` array stores only the first
        // run's start and every run's end, which hides such a hole inside a
        // page -- it only bites when a formatting page begins on it, and Word
        // then spins forever (`samples/sw-3-both.doc`). `close_gaps` hands the
        // mark to the paragraph that follows it.
        let mut runs = vec![
            (2048, 2084, ParaProps::default()),
            (2086, 2100, ParaProps::default()),
        ];
        close_gaps(&mut runs, 2048, 2104);
        assert_eq!(runs[0], (2048, 2084, ParaProps::default()));
        assert_eq!(runs[1].0, 2084, "the hole must be closed");
        assert_eq!(runs[1].1, 2104, "the last run must reach fc_mac");
    }

    #[test]
    fn paragraph_runs_cover_the_whole_text_stream() {
        let body = vec![
            Paragraph::text(format!("Paragraph 0 body{FOOTNOTE_MARK}")),
            Paragraph::text("Paragraph 1 body"),
            Paragraph::text(format!("Paragraph 2 body{ENDNOTE_MARK}")),
        ];
        let notes = vec![vec![Paragraph::text("footnote text")]];
        let endnotes = vec![vec![Paragraph::text("endnote text")]];
        let documents = Documents::new(&[], &[]).notes(&notes).endnotes(&endnotes);
        let bytes = write_document_with_footnotes(&body, &documents);
        let doc = crate::parse(&bytes).unwrap();
        assert!(doc.fib.ccp_counts[1] > 0 && doc.fib.ccp_counts[5] > 0);
        let mut expected = doc.fib.fc_min;
        for run in &doc.para_runs {
            assert_eq!(run.fc_start, expected, "hole before {run:?}");
            expected = run.fc_end;
        }
        assert_eq!(expected, doc.fib.fc_mac, "the last run must reach fc_mac");
    }

    #[test]
    fn a_document_without_notes_has_no_footnote_story() {
        let bytes = write_minimal("Only the body.");
        let doc = crate::parse(&bytes).unwrap();
        assert_eq!(doc.fib.ccp_counts[1], 0);
        assert!(doc.story(crate::Story::Footnote).is_empty());
        assert!(doc.footnote_slots.is_empty());
        assert!(doc.footnote_text_slots.is_empty());
        assert!(doc.endnote_slots.is_empty());
        assert!(doc.endnote_text_slots.is_empty());
    }

}
