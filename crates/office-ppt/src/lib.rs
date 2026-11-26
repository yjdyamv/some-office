//! # office-ppt — the PowerPoint 97-2003 binary format
//!
//! A `.ppt` file is a CFB container whose `PowerPoint Document` stream is a
//! tree of records. Every record is `[ver/instance: u16][type: u16][len: u32]`
//! followed by `len` bytes of payload; the payload of a *container* record is
//! itself a sequence of records. Two id spaces are mixed in that tree:
//!
//! * **PPT records** — ids such as `0x03E8` (Document), `0x03EE` (Slide),
//!   `0x0FA0` (TextCharsAtom).
//! * **OfficeArt (Escher) records** — ids `0xF000..=0xF004` for the drawing
//!   containers and shapes, `0xF00D` for a shape's text box.
//!
//! Slide order is *not* stream order. The last edit is described by
//! `UserEditAtom`, which points at a `PersistDirectoryAtom` mapping
//! persist-ids to stream offsets, and the Document's `SlideListWithText`
//! (instance 0) lists the slides in presentation order as `SlidePersistAtom`s
//! holding those persist-ids. That indirection exists because PowerPoint saves
//! incrementally: a slide can be rewritten at the end of the stream without
//! moving the other slides.
//!
//! ```no_run
//! let cfb = office_cfb::Cfb::open("deck.ppt")?;
//! let doc = cfb.stream("PowerPoint Document")?;
//! let user = cfb.stream("Current User").ok();
//! let deck = office_ppt::parse_stream(&doc, user.as_deref())?;
//! println!("{}", deck.plain_text());
//! # Ok::<(), Box<dyn std::error::Error>>(())
//! ```

use std::collections::HashMap;

use office_enc::{decode_cp1252, decode_utf16le};

// --- record ids -----------------------------------------------------------

pub const RT_DOCUMENT: u16 = 0x03E8;
pub const RT_DOCUMENT_ATOM: u16 = 0x03E9;
pub const RT_END_DOCUMENT: u16 = 0x03EA;
pub const RT_SLIDE: u16 = 0x03EE;
pub const RT_SLIDE_ATOM: u16 = 0x03EF;
pub const RT_MAIN_MASTER: u16 = 0x03F8;
pub const RT_SLIDE_PERSIST_ATOM: u16 = 0x03F3;
/// `Notes` container: one persist object per slide's speaker notes.
pub const RT_NOTES: u16 = 0x03F0;
pub const RT_NOTES_ATOM: u16 = 0x03F1;
pub const RT_PPDRAWING: u16 = 0x040C;
pub const RT_PPDRAWING_GROUP: u16 = 0x040D;
pub const RT_SLIDE_LIST_WITH_TEXT: u16 = 0x0FF0;
pub const RT_USER_EDIT_ATOM: u16 = 0x0FF5;
pub const RT_TEXT_HEADER_ATOM: u16 = 0x0F9F;
pub const RT_TEXT_CHARS_ATOM: u16 = 0x0FA0;
pub const RT_TEXT_BYTES_ATOM: u16 = 0x0FA8;
/// Character- and paragraph-level formatting of the text atom before it.
pub const RT_STYLE_TEXT_PROP_ATOM: u16 = 0x0FA1;
/// `TextMasterStyleAtom`: the placeholder styles on the main master.
pub const RT_TEXT_MASTER_STYLE_ATOM: u16 = 0x0FA3;
pub const RT_PERSIST_DIRECTORY_ATOM: u16 = 0x1772;
/// `CryptSession10Container` — holds the encryption header of a protected file.
pub const RT_CRYPT_SESSION: u16 = 0x2F14;

pub const RT_OFFICEART_DGG_CONTAINER: u16 = 0xF000;
pub const RT_OFFICEART_DG_CONTAINER: u16 = 0xF002;
pub const RT_OFFICEART_SPGR_CONTAINER: u16 = 0xF003;
pub const RT_OFFICEART_SP_CONTAINER: u16 = 0xF004;
pub const RT_OFFICEART_CLIENT_TEXTBOX: u16 = 0xF00D;
/// OfficeArt property table of a shape (`pib`, fill, line, ...).
pub const RT_OFFICEART_FOPT: u16 = 0xF00B;
/// The BLIP store at the top of the `Pictures` stream.
pub const RT_OFFICEART_BSTORE_CONTAINER: u16 = 0xF001;
/// One entry of that store (a picture plus its reference count).
pub const RT_OFFICEART_BSE: u16 = 0xF007;
/// `pib` property id inside an `OfficeArtFOPT`: 1-based index into `Pictures`.
pub const PID_PIB: u16 = 0x0104;

/// `CurrentUserAtom.headerToken`; PowerPoint refuses the stream without it.
pub const CURRENT_USER_TOKEN: u32 = 0xE391_C05F;

const MAX_DEPTH: usize = 32;

// --- errors ---------------------------------------------------------------

#[derive(Debug)]
pub enum Error {
    /// A record header or payload ran past the end of the stream.
    Truncated,
    /// The `Current User` stream did not start with `CURRENT_USER_TOKEN`.
    BadCurrentUserToken(u32),
    /// No `Document` record (0x03E8) in the stream.
    NoDocument,
    /// The presentation is not encrypted (no `CryptSession10Container`).
    NotEncrypted,
    /// The presentation is encrypted but no password was supplied.
    Encrypted,
    /// The password did not match the stored verifier.
    WrongPassword,
    Crypto(office_crypto::Error),
    /// The compound file could not be opened.
    Cfb(office_cfb::Error),
}

impl From<office_cfb::Error> for Error {
    fn from(e: office_cfb::Error) -> Self {
        Error::Cfb(e)
    }
}

impl From<office_crypto::Error> for Error {
    fn from(e: office_crypto::Error) -> Self {
        Error::Crypto(e)
    }
}

impl std::fmt::Display for Error {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Error::Truncated => write!(f, "record runs past the end of the stream"),
            Error::BadCurrentUserToken(t) => {
                write!(f, "Current User stream has token {t:#010x}, expected {CURRENT_USER_TOKEN:#010x}")
            }
            Error::NoDocument => write!(f, "no Document record in PowerPoint Document stream"),
            Error::NotEncrypted => write!(f, "presentation is not encrypted"),
            Error::Encrypted => write!(f, "presentation is encrypted; a password is required"),
            Error::WrongPassword => write!(f, "wrong password"),
            Error::Crypto(e) => write!(f, "crypto: {e}"),
            Error::Cfb(e) => write!(f, "compound file: {e}"),
        }
    }
}

impl std::error::Error for Error {}

// --- records --------------------------------------------------------------

/// One record header. `payload` lives at `offset + 8 .. offset + 8 + len`.
#[derive(Debug, Clone, Copy)]
pub struct Record {
    pub offset: usize,
    pub ver: u8,
    pub instance: u16,
    pub kind: u16,
    pub len: usize,
}

impl Record {
    pub fn payload<'a>(&self, data: &'a [u8]) -> &'a [u8] {
        &data[self.offset + 8..self.offset + 8 + self.len]
    }

    pub fn is_container(&self) -> bool {
        matches!(
            self.kind,
            RT_DOCUMENT
                | RT_SLIDE
                | RT_MAIN_MASTER
                | RT_PPDRAWING
                | RT_PPDRAWING_GROUP
                | RT_SLIDE_LIST_WITH_TEXT
                | RT_OFFICEART_DGG_CONTAINER
                | RT_OFFICEART_DG_CONTAINER
                | RT_OFFICEART_SPGR_CONTAINER
                | RT_OFFICEART_SP_CONTAINER
        )
    }
}

/// Read the record at `offset`, clamped to `end`. `None` if there is no room
/// for a header.
pub fn read_record(data: &[u8], offset: usize, end: usize) -> Option<Record> {
    if offset + 8 > end {
        return None;
    }
    let ver_inst = u16::from_le_bytes([data[offset], data[offset + 1]]);
    let kind = u16::from_le_bytes([data[offset + 2], data[offset + 3]]);
    let len = u32::from_le_bytes([
        data[offset + 4],
        data[offset + 5],
        data[offset + 6],
        data[offset + 7],
    ]) as usize;
    let len = len.min(end - offset - 8);
    Some(Record {
        offset,
        ver: (ver_inst & 0x0F) as u8,
        instance: ver_inst >> 4,
        kind,
        len,
    })
}

/// Iterate the sibling records in `data[start..end]`.
pub fn records(data: &[u8], start: usize, end: usize) -> Vec<Record> {
    let mut out = Vec::new();
    let mut off = start;
    while let Some(r) = read_record(data, off, end) {
        out.push(r);
        off += 8 + r.len;
    }
    out
}

// --- text -----------------------------------------------------------------

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum TextKind {
    Title,
    Body,
    Notes,
    Other,
}

impl TextKind {
    /// `TextHeaderAtom.textType`, see MS-PPT 2.13.33 `TextTypeEnum`.
    ///
    /// 0 title, 1 body, 2 notes, 3 unused, 4 other, 5 center body,
    /// 6 center title, 7 half body, 8 quarter body.
    fn from_atom(v: u32) -> TextKind {
        match v {
            0 | 6 => TextKind::Title,
            1 | 5 | 7 | 8 => TextKind::Body,
            2 => TextKind::Notes,
            _ => TextKind::Other,
        }
    }
}

/// Character- and paragraph-level formatting of a run of text.
///
/// `None` means "not specified by this run" — the value is inherited from the
/// placeholder's `TextMasterStyleAtom` on the main master, which this crate
/// does not resolve.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct TextFormat {
    pub bold: Option<bool>,
    pub italic: Option<bool>,
    pub underline: Option<bool>,
    /// Font size in points (`fontSize` is in points already, not half-points).
    pub size: Option<u16>,
    /// sRGB colour; the fourth byte of `ColorIndexStruct` is the palette index
    /// and `0xFE` means "use these RGB bytes".
    pub colour: Option<(u8, u8, u8)>,
    /// Index into the master's font table (`fontRef`).
    pub typeface: Option<u16>,
    /// `TextAlignmentEnum` from the paragraph run: 0 left, 1 centred, 2 right,
    /// 3 justified.
    pub align: Option<u16>,
    /// Outline level (`indentLevel`) of the paragraph run. It selects the entry
    /// of the placeholder's master style.
    pub level: Option<u16>,
}

impl TextFormat {
    pub fn is_empty(&self) -> bool {
        *self == TextFormat::default()
    }

    fn absorb(&mut self, other: &TextFormat) {
        self.bold = other.bold.or(self.bold);
        self.italic = other.italic.or(self.italic);
        self.underline = other.underline.or(self.underline);
        self.size = other.size.or(self.size);
        self.colour = other.colour.or(self.colour);
        self.typeface = other.typeface.or(self.typeface);
        self.align = other.align.or(self.align);
        self.level = other.level.or(self.level);
    }

    /// Inline CSS for the properties that are set.
    pub fn css(&self) -> String {
        let mut out = String::new();
        if self.bold == Some(true) {
            out.push_str("font-weight:bold;");
        }
        if self.italic == Some(true) {
            out.push_str("font-style:italic;");
        }
        if self.underline == Some(true) {
            out.push_str("text-decoration:underline;");
        }
        if let Some(pt) = self.size {
            out.push_str(&format!("font-size:{pt}pt;"));
        }
        if let Some((r, g, b)) = self.colour {
            out.push_str(&format!("color:#{r:02X}{g:02X}{b:02X};"));
        }
        out
    }
}

#[derive(Debug, Clone)]
pub struct TextRun {
    pub kind: TextKind,
    pub text: String,
    pub format: TextFormat,
    /// Runs that came from the same text atom share a group; formatting splits
    /// one text atom into several runs, and `plain_text` joins those back
    /// without a separator.
    pub group: usize,
}

#[derive(Debug, Clone)]
pub struct Slide {
    /// 0-based position in presentation order.
    pub index: usize,
    /// Persist id when the slide was reached through the persist directory.
    pub persist_id: Option<u32>,
    /// Offset of the `Slide` record inside the stream.
    pub offset: usize,
    pub texts: Vec<TextRun>,
    /// `pib` values of the pictures on this slide: 1-based indices into the
    /// `Pictures` stream (see [`pictures`]).
    pub pictures: Vec<u32>,
    /// Speaker notes. They live in a separate `Notes` persist object, which the
    /// `SlideListWithText` of instance 2 lists in slide order.
    pub notes: Vec<TextRun>,
}

impl Slide {
    /// One string per text atom (i.e. per shape), with formatting splits joined
    /// back together.
    pub fn paragraphs(&self) -> Vec<String> {
        join_groups(&self.texts)
    }

    pub fn plain_text(&self) -> String {
        let mut out = self
            .paragraphs()
            .iter()
            .map(|t| t.replace('\r', "\n"))
            .collect::<Vec<_>>()
            .join("\n");
        if !self.notes.is_empty() {
            out.push_str("\n[notes] ");
            out.push_str(&self.notes_plain_text());
        }
        out
    }

    /// Speaker notes only, without the slide text.
    pub fn notes_plain_text(&self) -> String {
        join_groups(&self.notes)
            .iter()
            .map(|t| t.replace('\r', "\n"))
            .collect::<Vec<_>>()
            .join("\n")
    }
}

/// PowerPoint stores a *field* placeholder (slide number, date, footer) as the
/// literal text `*` — the reader sees a stray asterisk in every notes page and
/// in any master that carries one. Treat such a run as a field, not as text.
fn is_field_placeholder(text: &str) -> bool {
    text.trim() == "*"
}

/// Concatenate runs that came from the same text atom; separate atoms stay
/// separate. A `StyleTextPropAtom` splits one atom into many runs, and those
/// must not gain a newline between them.
fn join_groups(runs: &[TextRun]) -> Vec<String> {
    let mut out: Vec<String> = Vec::new();
    let mut current: Option<usize> = None;
    for run in runs {
        if current == Some(run.group) {
            if let Some(last) = out.last_mut() {
                last.push_str(&run.text);
            }
        } else {
            current = Some(run.group);
            out.push(run.text.clone());
        }
    }
    out
}

#[derive(Debug, Default)]
pub struct Presentation {
    pub slides: Vec<Slide>,
    /// Placeholder styles from the main master, keyed by placeholder kind; the
    /// vector is indexed by outline level. Slide runs that do not state a
    /// property inherit it from here.
    pub master_styles: HashMap<TextKind, Vec<TextFormat>>,
    /// Encryption scheme when the file was protected and opened with a password.
    pub encryption: Option<office_crypto::Scheme>,
    /// True when slide order came from `SlideListWithText` rather than from
    /// stream order (i.e. the persist directory was usable).
    pub ordered_by_persist_directory: bool,
}

impl Presentation {
    pub fn plain_text(&self) -> String {
        self.slides
            .iter()
            .enumerate()
            .map(|(i, s)| format!("--- slide {} ---
{}
", i + 1, s.plain_text()))
            .collect::<Vec<_>>()
            .join("\n")
    }
}

// --- parsing --------------------------------------------------------------

/// Parse the `PowerPoint Document` stream. `current_user` is the optional
/// `Current User` stream; without it the last `UserEditAtom` in the stream is
/// used instead.
pub fn parse_stream(data: &[u8], current_user: Option<&[u8]>) -> Result<Presentation, Error> {
    let top = records(data, 0, data.len());

    // 1. Locate the newest UserEditAtom and its persist directory.
    let persist = persist_directory(data, current_user, &top);

    // A protected file keeps a CryptSession10Container at one of the persist
    // offsets; every object is encrypted, so nothing else is readable.
    if let Some(persist) = &persist {
        let protected = persist.values().any(|off| {
            read_record(data, *off as usize, data.len())
                .is_some_and(|r| r.kind == RT_CRYPT_SESSION)
        });
        if protected {
            return Err(Error::Encrypted);
        }
    }

    // The Document record normally sits at offset 0, but incremental saves can
    // leave dead records before the live one; take the first one reachable.
    let doc = top
        .iter()
        .find(|r| r.kind == RT_DOCUMENT)
        .copied()
        .ok_or(Error::NoDocument)?;

    // 2. SlideListWithText instance 0 lists slide persist-ids in presentation order.
    let mut slides = Vec::new();
    let mut ordered = false;
    if let Some(persist) = &persist {
        for list in records(data, doc.offset + 8, doc.offset + 8 + doc.len)
            .iter()
            .filter(|r| r.kind == RT_SLIDE_LIST_WITH_TEXT && r.instance == 0)
        {
            for atom in records(data, list.offset + 8, list.offset + 8 + list.len)
                .iter()
                .filter(|r| r.kind == RT_SLIDE_PERSIST_ATOM)
            {
                let p = atom.payload(data);
                if p.len() < 4 {
                    continue;
                }
                let pid = u32::from_le_bytes([p[0], p[1], p[2], p[3]]);
                if let Some(&off) = persist.get(&pid) {
                    if let Some(rec) = read_record(data, off as usize, data.len()) {
                        if rec.kind == RT_SLIDE {
                            slides.push(build_slide(data, rec, slides.len(), Some(pid)));
                            ordered = true;
                        }
                    }
                }
            }
        }
    }

    // 3. Fallback: every Slide record in stream order.
    if slides.is_empty() {
        for rec in top.iter().filter(|r| r.kind == RT_SLIDE) {
            slides.push(build_slide(data, *rec, slides.len(), None));
        }
    }

    attach_notes(data, doc, &persist, &mut slides);

    // 4. Placeholder styles live on the main master and are inherited by every
    //    run that does not state the property itself.
    let master_styles = main_master_styles(data, &persist, &top);
    for slide in &mut slides {
        inherit_master_styles(&master_styles, &mut slide.texts);
    }

    Ok(Presentation {
        encryption: None,
        slides,
        master_styles,
        ordered_by_persist_directory: ordered,
    })
}

/// Placeholder styles of the main master, keyed by placeholder kind.
fn main_master_styles(
    data: &[u8],
    persist: &Option<HashMap<u32, u32>>,
    top: &[Record],
) -> HashMap<TextKind, Vec<TextFormat>> {
    let mut master = persist.as_ref().and_then(|p| {
        p.values().find_map(|off| {
            let rec = read_record(data, *off as usize, data.len())?;
            (rec.kind == RT_MAIN_MASTER).then_some(rec)
        })
    });
    if master.is_none() {
        master = top.iter().find(|r| r.kind == RT_MAIN_MASTER).copied();
    }
    let Some(master) = master else {
        return HashMap::new();
    };
    let mut out = HashMap::new();
    let mut kind = TextKind::Other;
    let mut fallback = 0usize;
    collect_master_styles(
        data,
        master.offset + 8,
        master.offset + 8 + master.len,
        &mut out,
        &mut kind,
        &mut fallback,
        0,
    );
    out
}

/// Which placeholder a master style atom belongs to, from its `recInstance`.
fn master_style_kind(instance: u16) -> TextKind {
    match instance {
        0 | 4 => TextKind::Title,
        1 | 5 => TextKind::Body,
        _ => TextKind::Other,
    }
}

fn collect_master_styles(
    data: &[u8],
    start: usize,
    end: usize,
    out: &mut HashMap<TextKind, Vec<TextFormat>>,
    kind: &mut TextKind,
    fallback: &mut usize,
    depth: usize,
) {
    if depth > MAX_DEPTH {
        return;
    }
    for rec in records(data, start, end) {
        match rec.kind {
            RT_TEXT_HEADER_ATOM => {
                let p = rec.payload(data);
                if p.len() >= 4 {
                    *kind = TextKind::from_atom(u32::from_le_bytes([p[0], p[1], p[2], p[3]]));
                }
            }
            RT_TEXT_MASTER_STYLE_ATOM => {
                let levels = master_style_levels(rec.payload(data));
                if levels.is_empty() {
                    continue;
                }
                // `recInstance` is the style type (MS-PPT 2.9.46): 0/4 title,
                // 1/5 body, everything else "other". The lower numbers are the
                // master's own placeholders, so they win when both are present.
                let slot = master_style_kind(rec.instance);
                out.entry(slot).or_insert(levels);
                *fallback += 1;
            }
            k if rec.is_container() || k == RT_OFFICEART_CLIENT_TEXTBOX => {
                if k == RT_SLIDE_LIST_WITH_TEXT {
                    continue;
                }
                collect_master_styles(
                    data,
                    rec.offset + 8,
                    rec.offset + 8 + rec.len,
                    out,
                    kind,
                    fallback,
                    depth + 1,
                );
            }
            _ => {}
        }
    }
}

/// Decode a `TextMasterStyleAtom` payload into one [`TextFormat`] per level.
///
/// Layout (MS-PPT 2.9.46 `TextMasterStyleAtom`): `cLevels u16`, then for every
/// level a `TextPFException` immediately followed by a `TextCFException` — with
/// no run counts, unlike [`style_runs`], because a level applies to the whole
/// level.
pub fn master_style_levels(payload: &[u8]) -> Vec<TextFormat> {
    let Some(levels) = u16_at(payload, 0) else {
        return Vec::new();
    };
    let mut pos = 2usize;
    let mut out = Vec::new();
    for _ in 0..levels.min(9) {
        let Some(pf_masks) = u32_at(payload, pos) else { break };
        let Some((next, align)) = pf_exception(payload, pos + 4, pf_masks) else {
            break;
        };
        if next <= pos {
            break;
        }
        pos = next;
        let Some(cf_masks) = u32_at(payload, pos) else { break };
        let Some((next, mut fmt)) = cf_exception(payload, pos + 4, cf_masks) else {
            break;
        };
        if next <= pos {
            break;
        }
        pos = next;
        fmt.align = align;
        // Reject nonsense so a misparse never leaks into the output.
        if fmt.size == Some(0) || fmt.size.is_some_and(|s| s > 400) {
            fmt.size = None;
        }
        out.push(fmt);
    }
    out
}

/// Fill in properties a slide run leaves unset from the master style of the same
/// placeholder kind and outline level.
fn inherit_master_styles(master: &HashMap<TextKind, Vec<TextFormat>>, runs: &mut [TextRun]) {
    for run in runs {
        let Some(levels) = master.get(&run.kind) else {
            continue;
        };
        let level = run.format.level.unwrap_or(0) as usize;
        let Some(base) = levels.get(level).or_else(|| levels.first()) else {
            continue;
        };
        let own = run.format.clone();
        let mut merged = base.clone();
        merged.absorb(&own); // what the run states itself always wins
        merged.level = own.level;
        run.format = merged;
    }
}

/// Fill `Slide::notes` from the `Notes` persist objects.
///
/// The notes are *not* inside the slide: PowerPoint keeps one `Notes` object per
/// slide, listed by the `SlideListWithText` whose instance is 2, in the same
/// order as the slides. When that list is missing (older or trimmed files) the
/// `Notes` objects are matched by persist offset instead.
fn attach_notes(
    data: &[u8],
    doc: Record,
    persist: &Option<HashMap<u32, u32>>,
    slides: &mut [Slide],
) {
    if slides.is_empty() {
        return;
    }
    // 1. Every Notes object we can reach.
    let mut by_pid: HashMap<u32, Vec<TextRun>> = HashMap::new();
    if let Some(persist) = persist {
        for (&pid, &off) in persist.iter() {
            if let Some(rec) = read_record(data, off as usize, data.len()) {
                if rec.kind == RT_NOTES {
                    by_pid.insert(pid, notes_text(data, rec));
                }
            }
        }
    }
    // 2. The notes list gives the slide order.
    if let Some(persist) = persist {
        let mut order = Vec::new();
        for list in records(data, doc.offset + 8, doc.offset + 8 + doc.len)
            .iter()
            .filter(|r| r.kind == RT_SLIDE_LIST_WITH_TEXT && r.instance == 2)
        {
            for atom in records(data, list.offset + 8, list.offset + 8 + list.len) {
                if atom.kind != RT_SLIDE_PERSIST_ATOM && atom.kind != RT_NOTES_ATOM {
                    continue;
                }
                let p = atom.payload(data);
                if p.len() < 4 {
                    continue;
                }
                order.push(u32::from_le_bytes([p[0], p[1], p[2], p[3]]));
            }
        }
        if !order.is_empty() {
            for (i, slide) in slides.iter_mut().enumerate() {
                if let Some(pid) = order.get(i) {
                    if let Some(notes) = by_pid.get(pid) {
                        slide.notes = notes.clone();
                    }
                }
            }
            if slides.iter().any(|s| !s.notes.is_empty()) {
                return;
            }
        }
        // 3. No usable list: fall back to persist order.
        let mut ids: Vec<u32> = by_pid.keys().copied().collect();
        ids.sort_by_key(|pid| persist.get(pid).copied().unwrap_or(0));
        for (i, slide) in slides.iter_mut().enumerate() {
            if let Some(notes) = ids.get(i).and_then(|pid| by_pid.get(pid)) {
                slide.notes = notes.clone();
            }
        }
    }
}

fn build_slide(data: &[u8], rec: Record, index: usize, persist_id: Option<u32>) -> Slide {
    let mut texts = Vec::new();
    let mut kind = TextKind::Other;
    let mut pictures = Vec::new();
    walk(
        data,
        rec.offset + 8,
        rec.offset + 8 + rec.len,
        &mut texts,
        &mut kind,
        &mut pictures,
        &mut 0usize,
        0,
    );
    Slide {
        index,
        persist_id,
        offset: rec.offset,
        texts,
        pictures,
        notes: Vec::new(),
    }
}

/// The text of a `Notes` object: same shape as a slide, without pictures.
fn notes_text(data: &[u8], rec: Record) -> Vec<TextRun> {
    let mut texts = Vec::new();
    let mut kind = TextKind::Other;
    let mut pictures = Vec::new();
    walk(
        data,
        rec.offset + 8,
        rec.offset + 8 + rec.len,
        &mut texts,
        &mut kind,
        &mut pictures,
        &mut 0usize,
        0,
    );
    texts
}

// --- StyleTextPropAtom ----------------------------------------------------

fn u16_at(d: &[u8], at: usize) -> Option<u16> {
    Some(u16::from_le_bytes([*d.get(at)?, *d.get(at + 1)?]))
}

fn u32_at(d: &[u8], at: usize) -> Option<u32> {
    Some(u32::from_le_bytes([
        *d.get(at)?,
        *d.get(at + 1)?,
        *d.get(at + 2)?,
        *d.get(at + 3)?,
    ]))
}

/// Skip a `TextPFException` (`PFMasks` + the fields they select) and report the
/// alignment when it carries one.
///
/// Field widths follow MS-PPT 2.9.38 `TextPFException`; the mask bits are
/// MS-PPT 2.9.29 `PFMasks` (bit 11 = `align`, bit 20 = `tabStops`).
fn pf_exception(d: &[u8], mut pos: usize, masks: u32) -> Option<(usize, Option<u16>)> {
    let mut align = None;
    // bulletFlags (2) covers bits 0..=3.
    if masks & 0x0000_000F != 0 {
        pos += 2;
    }
    // bulletChar (2), bulletFontRef (2), bulletSize (2), bulletColor (4).
    if masks & 0x0000_0080 != 0 {
        pos += 2;
    }
    if masks & 0x0000_0010 != 0 {
        pos += 2;
    }
    if masks & 0x0000_0040 != 0 {
        pos += 2;
    }
    if masks & 0x0000_0020 != 0 {
        pos += 4;
    }
    if masks & 0x0000_0800 != 0 {
        align = Some(u16_at(d, pos)?);
        pos += 2;
    }
    // lineSpacing, spaceBefore, spaceAfter, leftMargin, indent, defaultTabSize.
    for bit in [0x1000u32, 0x2000, 0x4000, 0x0100, 0x0400, 0x8000] {
        if masks & bit != 0 {
            pos += 2;
        }
    }
    if masks & 0x0010_0000 != 0 {
        // tabStops: count (2) then count 4-byte stops.
        let count = u16_at(d, pos)? as usize;
        pos += 2 + count * 4;
    }
    if masks & 0x0001_0000 != 0 {
        pos += 2; // fontAlign
    }
    if masks & 0x000E_0000 != 0 {
        pos += 2; // wrapFlags
    }
    if masks & 0x0020_0000 != 0 {
        pos += 2; // textDirection
    }
    if pos > d.len() {
        return None;
    }
    Some((pos, align))
}

/// Read a `TextCFException` (`CFMasks` + selected fields).
///
/// Mask bits are MS-PPT 2.9.6 `CFMasks`: 0 bold, 1 italic, 2 underline,
/// 16 typeface, 17 size, 18 color, 19 position, 21..=23 other typefaces.
fn cf_exception(d: &[u8], mut pos: usize, masks: u32) -> Option<(usize, TextFormat)> {
    let mut f = TextFormat::default();
    if masks & 0x0000_0437 != 0 {
        let style = u16_at(d, pos)?;
        pos += 2;
        if masks & 0x01 != 0 {
            f.bold = Some(style & 0x01 != 0);
        }
        if masks & 0x02 != 0 {
            f.italic = Some(style & 0x02 != 0);
        }
        if masks & 0x04 != 0 {
            f.underline = Some(style & 0x04 != 0);
        }
    }
    if masks & 0x0001_0000 != 0 {
        f.typeface = Some(u16_at(d, pos)?);
        pos += 2;
    }
    for bit in [0x0020_0000u32, 0x0040_0000, 0x0080_0000] {
        if masks & bit != 0 {
            pos += 2;
        }
    }
    if masks & 0x0002_0000 != 0 {
        f.size = Some(u16_at(d, pos)?);
        pos += 2;
    }
    if masks & 0x0004_0000 != 0 {
        let r = *d.get(pos)?;
        let g = *d.get(pos + 1)?;
        let b = *d.get(pos + 2)?;
        f.colour = Some((r, g, b));
        pos += 4;
    }
    if masks & 0x0008_0000 != 0 {
        pos += 2; // position
    }
    if pos > d.len() {
        return None;
    }
    Some((pos, f))
}

/// Split a `StyleTextPropAtom` payload into `(start, end, format)` character
/// ranges over `cch` characters.
///
/// The payload is a `rgTextPFRun` array followed by a `rgTextCFRun` array with
/// no counts of their own (MS-PPT 2.9.35): a paragraph run is `count`,
/// `indentLevel` and a `TextPFException`; a character run is `count` and a
/// `TextCFException`. Both arrays end where the next one begins, and the sum of
/// their counts equals the length of the text the atom belongs to.
///
/// Returns an empty vector when nothing can be decoded, so callers keep the
/// unformatted run.
pub fn style_runs(payload: &[u8], cch: usize) -> Vec<(usize, usize, TextFormat)> {
    parse_style(payload, cch).0
}

/// How many bytes of `payload` the run arrays actually covered.
///
/// A fully decoded `StyleTextPropAtom` covers the whole payload; a smaller
/// number means the walk stopped early and the formatting is only partial.
/// Exposed so tests can assert the field widths are right on real files.
pub fn style_runs_len(payload: &[u8], cch: usize) -> usize {
    parse_style(payload, cch).1
}

/// Parse the `rgTextCFRun` array that starts at `start`; it must end exactly at
/// the end of the payload.
///
/// `None` means "this is not where the character runs begin", which is how the
/// caller finds the boundary between the paragraph runs and the character runs:
/// the payload has no length fields for either array, so adjacency plus an exact
/// end is the only signal.
fn parse_cf_runs(
    payload: &[u8],
    start: usize,
    cch: usize,
) -> Option<Vec<(usize, usize, TextFormat)>> {
    let mut pos = start;
    let mut at = 0usize;
    let mut runs = Vec::new();
    while pos < payload.len() {
        let count = u32_at(payload, pos)? as usize;
        if count == 0 || at + count > cch + 1 {
            return None;
        }
        let masks = u32_at(payload, pos + 4)?;
        let (next, fmt) = cf_exception(payload, pos + 8, masks)?;
        if next <= pos {
            return None;
        }
        pos = next;
        runs.push((at, (at + count).min(cch), fmt));
        at += count;
    }
    if pos != payload.len() || at == 0 {
        return None;
    }
    Some(runs)
}

fn parse_style(payload: &[u8], cch: usize) -> (Vec<(usize, usize, TextFormat)>, usize) {
    if cch == 0 || payload.len() < 8 {
        return (Vec::new(), 0);
    }

    // Paragraph runs. Record every boundary we could stop at, together with the
    // alignment each run contributes.
    let mut bounds: Vec<(usize, Vec<(usize, Option<u16>, u16)>)> = vec![(0, Vec::new())];
    let mut pos = 0usize;
    let mut at = 0usize;
    let mut aligns: Vec<(usize, Option<u16>, u16)> = Vec::new();
    while at < cch {
        let Some(count) = u32_at(payload, pos) else { break };
        let count = count as usize;
        if count == 0 || at + count > cch + 1 {
            break;
        }
        let Some(level) = u16_at(payload, pos + 4) else { break };
        let Some(masks) = u32_at(payload, pos + 6) else { break };
        let Some((next, align)) = pf_exception(payload, pos + 10, masks) else {
            break;
        };
        if next <= pos {
            break;
        }
        pos = next;
        at += count;
        aligns.push((at, align, level));
        bounds.push((pos, aligns.clone()));
    }

    let build = |aligns: &[(usize, Option<u16>, u16)], runs: &[(usize, usize, TextFormat)]| {
        let mut per_char = vec![TextFormat::default(); cch];
        let mut from = 0usize;
        for (to, align, level) in aligns {
            let to = (*to).min(cch);
            for slot in &mut per_char[from..to] {
                slot.level = Some(*level);
                if let Some(a) = align {
                    slot.align = Some(*a);
                }
            }
            from = to;
        }
        for (start, end, fmt) in runs {
            for slot in &mut per_char[*start..*end] {
                slot.absorb(fmt);
            }
        }
        let mut out: Vec<(usize, usize, TextFormat)> = Vec::new();
        let mut start = 0usize;
        for i in 1..=cch {
            if i == cch || per_char[i] != per_char[start] {
                out.push((start, i, per_char[start].clone()));
                start = i;
            }
        }
        out
    };

    // The smallest boundary whose character runs consume the rest of the atom.
    for (start, aligns) in &bounds {
        if let Some(runs) = parse_cf_runs(payload, *start, cch) {
            return (build(aligns, &runs), payload.len());
        }
    }

    // No character runs at all: only legal when the paragraph runs fill the atom.
    if let Some((start, aligns)) = bounds.last() {
        if *start == payload.len() {
            return (build(aligns, &[]), payload.len());
        }
    }
    (Vec::new(), 0)
}

/// Depth-first walk over PPT + OfficeArt records collecting text.
fn walk(
    data: &[u8],
    start: usize,
    end: usize,
    out: &mut Vec<TextRun>,
    kind: &mut TextKind,
    pictures: &mut Vec<u32>,
    groups: &mut usize,
    depth: usize,
) {
    if depth > MAX_DEPTH {
        return;
    }
    // Index in `out` of the run pushed by the most recent text atom, so a
    // `StyleTextPropAtom` only ever formats the text atom directly before it.
    // An empty text box still has a style atom; without this the atom would
    // steal the previous shape's run.
    let mut pending: Option<usize> = None;
    for rec in records(data, start, end) {
        match rec.kind {
            RT_TEXT_HEADER_ATOM => {
                pending = None;
                let p = rec.payload(data);
                if p.len() >= 4 {
                    *kind = TextKind::from_atom(u32::from_le_bytes([p[0], p[1], p[2], p[3]]));
                }
            }
            RT_TEXT_CHARS_ATOM => {
                let text = decode_utf16le(rec.payload(data));
                if is_field_placeholder(&text) {
                    continue;
                }
                *groups += 1;
                out.push(TextRun {
                    kind: *kind,
                    text,
                    format: TextFormat::default(),
                    group: *groups,
                });
                pending = Some(out.len() - 1);
            }
            RT_TEXT_BYTES_ATOM => {
                let text = decode_cp1252(rec.payload(data));
                if is_field_placeholder(&text) {
                    continue;
                }
                *groups += 1;
                out.push(TextRun {
                    kind: *kind,
                    text,
                    format: TextFormat::default(),
                    group: *groups,
                });
                pending = Some(out.len() - 1);
            }
            RT_STYLE_TEXT_PROP_ATOM => {
                let Some(idx) = pending.take() else { continue };
                if idx + 1 != out.len() {
                    continue;
                }
                let last = out.pop().unwrap();
                let chars: Vec<char> = last.text.chars().collect();
                let spans = style_runs(rec.payload(data), chars.len());
                if spans.is_empty() {
                    out.push(last);
                    continue;
                }
                let mut at = 0usize;
                for (start, end, format) in spans {
                    if start > at {
                        out.push(TextRun {
                            kind: last.kind,
                            text: chars[at..start].iter().collect(),
                            format: TextFormat::default(),
                            group: last.group,
                        });
                    }
                    out.push(TextRun {
                        kind: last.kind,
                        text: chars[start..end].iter().collect(),
                        format,
                        group: last.group,
                    });
                    at = end;
                }
                if at < chars.len() {
                    out.push(TextRun {
                        kind: last.kind,
                        text: chars[at..].iter().collect(),
                        format: TextFormat::default(),
                        group: last.group,
                    });
                }
            }
            k if rec.is_container() || k == RT_OFFICEART_CLIENT_TEXTBOX => {
                // ClientTextbox carries the shape's text; everything else here
                // is a container whose payload is more records.
                pending = None;
                if k == RT_SLIDE_LIST_WITH_TEXT {
                    continue; // outline text of *all* slides lives here
                }
                walk(
                    data,
                    rec.offset + 8,
                    rec.offset + 8 + rec.len,
                    out,
                    kind,
                    pictures,
                    groups,
                    depth + 1,
                );
            }
            RT_OFFICEART_FOPT => {
                pending = None;
                if let Some(pib) = fopt_picture(rec.instance as usize, rec.payload(data)) {
                    pictures.push(pib);
                }
            }
            _ => pending = None,
        }
    }
}

/// Pull the `pib` (picture index) out of an `OfficeArtFOPT` property table.
///
/// The instance field is the number of properties; each is a 6-byte
/// `(pid u16, value u32)` pair. The id needs masking because `pib` is written
/// with the `fBid` flag set (`0x4104`, not `0x0104`).
fn fopt_picture(count: usize, payload: &[u8]) -> Option<u32> {
    let n = count.min(payload.len() / 6);
    for k in 0..n {
        let o = k * 6;
        // OfficeArtFOPTE packs the id into 14 bits: bit 14 = fBid (the value is
        // a BLIP id), bit 15 = fComplex. Real files set fBid on `pib`, so mask.
        let pid = u16::from_le_bytes([payload[o], payload[o + 1]]) & 0x3FFF;
        if pid == PID_PIB {
            return Some(u32::from_le_bytes([
                payload[o + 2],
                payload[o + 3],
                payload[o + 4],
                payload[o + 5],
            ]));
        }
    }
    None
}

/// Resolve persist-id → stream offset for the newest edit.
fn persist_directory(
    data: &[u8],
    current_user: Option<&[u8]>,
    top: &[Record],
) -> Option<HashMap<u32, u32>> {
    let user_edit = user_edit_atom(data, current_user, top)?;
    let p = user_edit.payload(data);
    if p.len() < 16 {
        return None;
    }
    let dir_off = u32::from_le_bytes([p[12], p[13], p[14], p[15]]) as usize;
    let rec = read_record(data, dir_off, data.len())?;
    if rec.kind != RT_PERSIST_DIRECTORY_ATOM {
        return None;
    }
    Some(parse_persist_directory(rec.payload(data)))
}

fn user_edit_atom(data: &[u8], current_user: Option<&[u8]>, top: &[Record]) -> Option<Record> {
    if let Some(cu) = current_user {
        if cu.len() >= 20 {
            let token = u32::from_le_bytes([cu[12], cu[13], cu[14], cu[15]]);
            if token == CURRENT_USER_TOKEN || token == office_crypto::TOKEN_ENCRYPTED {
                let off = u32::from_le_bytes([cu[16], cu[17], cu[18], cu[19]]) as usize;
                if let Some(rec) = read_record(data, off, data.len()) {
                    if rec.kind == RT_USER_EDIT_ATOM {
                        return Some(rec);
                    }
                }
            }
        }
    }
    top.iter().rev().find(|r| r.kind == RT_USER_EDIT_ATOM).copied()
}

fn parse_persist_directory(p: &[u8]) -> HashMap<u32, u32> {
    let mut map = HashMap::new();
    let mut i = 0usize;
    while i + 4 <= p.len() {
        let w = u32::from_le_bytes([p[i], p[i + 1], p[i + 2], p[i + 3]]);
        let first_id = w & 0x000F_FFFF;
        // The 12-bit field is the number of offsets that follow, not "count
        // minus 1": measured 4 for the 32-byte directory of samples/hello.ppt
        // (4 objects), 7 for samples/notes.ppt (7 objects) and 4 for
        // samples/encrypted.ppt. The clamp only guards a truncated record,
        // which incremental saves can leave behind.
        let mut n = ((w >> 20) & 0xFFF) as usize;
        let avail = (p.len() - i - 4) / 4;
        if n > avail {
            n = avail;
        }
        for k in 0..n {
            let o = i + 4 + k * 4;
            let off = u32::from_le_bytes([p[o], p[o + 1], p[o + 2], p[o + 3]]);
            map.insert(first_id + k as u32, off);
        }
        i += 4 + n * 4;
    }
    map
}


// --- encryption -----------------------------------------------------------

/// Decrypt a `PowerPoint Document` stream in place and return the plaintext.
///
/// PowerPoint does not use one long RC4 stream. Every persist object (the
/// `Document`, the `MainMaster`, each `Slide`, …) is encrypted on its own with
/// `block = persistId`, and the encryption header lives in a
/// `CryptSession10Container` record (0x2F14) whose offset the persist directory
/// gives under `UserEditAtom.encryptSessionPersistIdRef`.
///
/// Returns [`Error::NotEncrypted`] when the stream carries no container, so
/// callers can use it as a probe.
pub fn decrypt_stream(
    data: &[u8],
    current_user: Option<&[u8]>,
    password: &str,
) -> Result<Vec<u8>, Error> {
    let top = records(data, 0, data.len());
    let persist = persist_directory(data, current_user, &top).ok_or(Error::Truncated)?;
    let user_edit = user_edit_atom(data, current_user, &top).ok_or(Error::NoDocument)?;
    let p = user_edit.payload(data);
    // UserEditAtom: … lastView u16 (24), unused u16 (26),
    // encryptSessionPersistIdRef u32 (28). The 0x1C-byte form has no such
    // field at all, so a short record means "not encrypted".
    if p.len() < 32 {
        return Err(Error::NotEncrypted);
    }
    let enc_ref = u32::from_le_bytes([p[28], p[29], p[30], p[31]]);
    if enc_ref == 0 {
        return Err(Error::NotEncrypted);
    }
    let Some(container_off) = persist.get(&enc_ref).copied() else {
        return Err(Error::NotEncrypted);
    };
    let container = read_record(data, container_off as usize, data.len()).ok_or(Error::Truncated)?;
    if container.kind != RT_CRYPT_SESSION {
        return Err(Error::NotEncrypted);
    }
    let header = office_crypto::parse_header(container.payload(data))?;
    if !header.scheme.verify(password) {
        return Err(Error::WrongPassword);
    }

    // Objects are laid out in stream order; an object ends where the next one
    // begins (its own length field is encrypted too).
    let mut items: Vec<(u32, u32)> = persist.iter().map(|(id, off)| (*id, *off)).collect();
    items.sort_by_key(|(_, off)| *off);
    let key_size = header.scheme.key_size().max(8) as usize;

    let mut out = data.to_vec();
    // The container itself is not encrypted; blank it so the decrypted stream
    // no longer looks protected.
    let cstart = container_off as usize;
    let cend = (cstart + 8 + container.len).min(out.len());
    out[cstart..cend].fill(0);
    for (i, (persist_id, offset)) in items.iter().enumerate() {
        let offset = *offset as usize;
        if *persist_id == enc_ref {
            continue;
        }
        let end = items
            .get(i + 1)
            .map(|(_, o)| *o as usize)
            .unwrap_or(data.len());
        if end <= offset + 8 || end > data.len() {
            continue;
        }
        let total = end - offset;
        // PowerPoint rounds the RC4 block size up to a multiple of the key
        // length in bits, not bytes — reproduce that quirk.
        let block_size = key_size * (total / key_size + 1);
        let plain = match &header.scheme {
            office_crypto::Scheme::LegacyRc4 { salt, .. } => {
                office_crypto::decrypt_legacy_blocks_from(
                    &data[offset..end],
                    password,
                    salt,
                    block_size,
                    *persist_id,
                )
            }
            office_crypto::Scheme::Rc4CryptoApi {
                salt, key_size: ks, ..
            } => office_crypto::decrypt_cryptoapi_blocks(
                &data[offset..end],
                password,
                salt,
                *ks,
                block_size,
                *persist_id,
            ),
            office_crypto::Scheme::Xor { .. } => return Err(Error::WrongPassword),
        };
        out[offset..offset + plain.len()].copy_from_slice(&plain);
    }
    Ok(out)
}

/// The encryption scheme of a protected stream, without decrypting anything.
pub fn encryption_scheme(
    data: &[u8],
    current_user: Option<&[u8]>,
) -> Option<office_crypto::Scheme> {
    let top = records(data, 0, data.len());
    let persist = persist_directory(data, current_user, &top)?;
    let user_edit = user_edit_atom(data, current_user, &top)?;
    let p = user_edit.payload(data);
    if p.len() < 32 {
        return None;
    }
    let enc_ref = u32::from_le_bytes([p[28], p[29], p[30], p[31]]);
    let offset = *persist.get(&enc_ref)? as usize;
    let container = read_record(data, offset, data.len())?;
    if container.kind != RT_CRYPT_SESSION {
        return None;
    }
    office_crypto::parse_header(container.payload(data))
        .ok()
        .map(|h| h.scheme)
}

/// Parse a `PowerPoint Document` stream, decrypting it when it is protected.
pub fn parse_stream_with_password(
    data: &[u8],
    current_user: Option<&[u8]>,
    password: &str,
) -> Result<Presentation, Error> {
    match decrypt_stream(data, current_user, password) {
        Ok(plain) => {
            let mut deck = parse_stream(&plain, current_user)?;
            deck.encryption = encryption_scheme(data, current_user);
            Ok(deck)
        }
        Err(Error::NotEncrypted) => parse_stream(data, current_user),
        Err(e) => Err(e),
    }
}

/// Open a `.ppt` file from disk.
pub fn open<P: AsRef<std::path::Path>>(path: P) -> Result<Presentation, Error> {
    let bytes = std::fs::read(path).map_err(|_| Error::Truncated)?;
    let cfb = office_cfb::Cfb::parse(bytes).map_err(|_| Error::Truncated)?;
    let doc = cfb.stream("PowerPoint Document").map_err(|_| Error::Truncated)?;
    let user = cfb.stream("Current User").ok();
    parse_stream(&doc, user.as_deref())
}

/// Open a password protected `.ppt` file from disk.
pub fn open_with_password<P: AsRef<std::path::Path>>(
    path: P,
    password: &str,
) -> Result<Presentation, Error> {
    let bytes = std::fs::read(path).map_err(|_| Error::Truncated)?;
    let cfb = office_cfb::Cfb::parse(bytes).map_err(|_| Error::Truncated)?;
    let doc = cfb.stream("PowerPoint Document").map_err(|_| Error::Truncated)?;
    let user = cfb.stream("Current User").ok();
    parse_stream_with_password(&doc, user.as_deref(), password)
}

// --- pictures -------------------------------------------------------------

/// Image formats that can appear inside a BLIP record.
pub mod write;

pub use office_blip::PictureKind;

/// One image from the `Pictures` stream.
#[derive(Debug, Clone)]
pub struct Picture {
    /// 1-based index: this is the value a slide's `pib` property refers to.
    pub index: u32,
    pub kind: PictureKind,
    /// The image bytes exactly as stored (PNG/JPEG/GIF/BMP/TIFF/EMF/WMF file).
    pub data: Vec<u8>,
}

/// Extract every image from the `Pictures` stream of a `.ppt` file.
///
/// The stream is normally an `OfficeArtBStoreContainer` holding `OfficeArtBSE`
/// records; a bare BLIP (a single picture written directly, which WPS does) is
/// accepted too. See [`office_blip`] for why the type is sniffed rather than
/// read from the record header.
pub fn pictures(stream: &[u8]) -> Vec<Picture> {
    let mut out = Vec::new();
    let push = |blip: office_blip::Blip<'_>, out: &mut Vec<Picture>| {
        out.push(Picture {
            index: out.len() as u32 + 1,
            kind: blip.kind,
            data: blip.data.to_vec(),
        });
    };

    for rec in records(stream, 0, stream.len()) {
        let payload = rec.payload(stream);
        match rec.kind {
            RT_OFFICEART_BSTORE_CONTAINER => {
                for bse in records(stream, rec.offset + 8, rec.offset + 8 + rec.len) {
                    if bse.kind != office_blip::RT_FBSE {
                        continue;
                    }
                    if let Some(blip) = office_blip::from_bse(bse.payload(stream)) {
                        push(blip, &mut out);
                    }
                }
            }
            k if office_blip::is_blip(k) => {
                if let Some(blip) = office_blip::from_blip(payload) {
                    push(blip, &mut out);
                }
            }
            _ => {}
        }
    }
    out
}

#[cfg(test)]
mod tests {
    use super::*;

    fn synth_record(kind: u16, instance: u16, payload: &[u8], out: &mut Vec<u8>) {
        out.extend_from_slice(&instance.to_le_bytes());
        out.extend_from_slice(&kind.to_le_bytes());
        out.extend_from_slice(&(payload.len() as u32).to_le_bytes());
        out.extend_from_slice(payload);
    }

    #[test]
    fn record_header_roundtrip() {
        let mut buf = Vec::new();
        synth_record(RT_TEXT_CHARS_ATOM, 0, &[b'h', 0, b'i', 0], &mut buf);
        let r = read_record(&buf, 0, buf.len()).unwrap();
        assert_eq!(r.kind, RT_TEXT_CHARS_ATOM);
        assert_eq!(r.len, 4);
        assert_eq!(decode_utf16le(r.payload(&buf)), "hi");
    }

    #[test]
    fn persist_directory_parses_both_count_conventions() {
        // WPS style: 0x00400001 = first id 1, 4 offsets (count field = 4).
        let mut p = Vec::new();
        p.extend_from_slice(&0x0040_0001u32.to_le_bytes());
        for off in [0u32, 1451, 44385, 46275] {
            p.extend_from_slice(&off.to_le_bytes());
        }
        let map = parse_persist_directory(&p);
        assert_eq!(map.len(), 4);
        assert_eq!(map[&1], 0);
        assert_eq!(map[&4], 46275);
    }

    #[test]
    fn finds_title_text_through_containers() {
        // Slide > PPDrawing > DgContainer > SpContainer > ClientTextbox > [hdr, chars]
        let mut chars = Vec::new();
        synth_record(RT_TEXT_CHARS_ATOM, 0, &[b'A', 0, b'B', 0], &mut chars);
        let mut hdr = Vec::new();
        synth_record(RT_TEXT_HEADER_ATOM, 0, &0u32.to_le_bytes(), &mut hdr);
        hdr.extend_from_slice(&chars);
        let mut textbox = Vec::new();
        synth_record(RT_OFFICEART_CLIENT_TEXTBOX, 0, &hdr, &mut textbox);
        let mut sp = Vec::new();
        synth_record(RT_OFFICEART_SP_CONTAINER, 0, &textbox, &mut sp);
        let mut dg = Vec::new();
        synth_record(RT_OFFICEART_DG_CONTAINER, 0, &sp, &mut dg);
        let mut drawing = Vec::new();
        synth_record(RT_PPDRAWING, 0, &dg, &mut drawing);
        let mut slide = Vec::new();
        synth_record(RT_SLIDE, 0, &drawing, &mut slide);

        let mut texts = Vec::new();
        let mut kind = TextKind::Other;
        let mut pics = Vec::new();
        walk(&slide, 0, slide.len(), &mut texts, &mut kind, &mut pics, &mut 0usize, 0);
        assert_eq!(texts.len(), 1);
        assert_eq!(texts[0].text, "AB");
        assert_eq!(texts[0].kind, TextKind::Title);
    }

    fn encrypted_deck() -> Option<(Vec<u8>, Option<Vec<u8>>)> {
        let p = concat!(env!("CARGO_MANIFEST_DIR"), "/../../samples/encrypted.ppt");
        if !std::path::Path::new(p).exists() {
            return None;
        }
        let bytes = std::fs::read(p).ok()?;
        let cfb = office_cfb::Cfb::parse(bytes).ok()?;
        let doc = cfb.stream("PowerPoint Document").ok()?;
        let user = cfb.stream("Current User").ok();
        Some((doc, user))
    }

    #[test]
    fn reports_an_encrypted_presentation_as_locked() {
        let Some((doc, user)) = encrypted_deck() else { return };
        let r = parse_stream(&doc, user.as_deref());
        assert!(matches!(r, Err(Error::Encrypted)), "got {r:?}");
        assert!(matches!(
            parse_stream_with_password(&doc, user.as_deref(), "wrong-password"),
            Err(Error::WrongPassword)
        ));
    }

    #[test]
    fn decrypts_the_sample_presentation() {
        let Some((doc, user)) = encrypted_deck() else { return };
        let deck = parse_stream_with_password(&doc, user.as_deref(), "secret123").expect("decrypt");
        assert!(!deck.slides.is_empty());
        let text = deck.plain_text();
        assert!(text.contains("ENCRYPTED-PPT-999"), "text was: {text}");
    }

    #[test]
    fn fopt_picture_finds_pib_among_other_properties() {
        // two properties: pid 0x0100 (fill), then pid 0x0104 (pib = 7)
        let mut p = Vec::new();
        p.extend_from_slice(&0x0100u16.to_le_bytes());
        p.extend_from_slice(&1u32.to_le_bytes());
        p.extend_from_slice(&PID_PIB.to_le_bytes());
        p.extend_from_slice(&7u32.to_le_bytes());
        assert_eq!(fopt_picture(2, &p), Some(7));
        // and with the fBid flag that real files set on `pib`
        let mut q = Vec::new();
        q.extend_from_slice(&0x4104u16.to_le_bytes());
        q.extend_from_slice(&3u32.to_le_bytes());
        assert_eq!(fopt_picture(1, &q), Some(3));
        assert_eq!(fopt_picture(1, &p), None);
    }

    /// Walk every `.ppt` in `samples/` and require that each
    /// `StyleTextPropAtom` is decoded to the last byte.
    ///
    /// A short read means one of the `PFMasks`/`CFMasks` field widths is wrong,
    /// which would silently shift every following property.
    #[test]
    fn sample_style_atoms_are_decoded_exactly() {
        let dir = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../samples");
        let Ok(entries) = std::fs::read_dir(&dir) else { return };
        let mut checked = 0usize;
        for entry in entries.flatten() {
            let path = entry.path();
            if path.extension().and_then(|e| e.to_str()) != Some("ppt") {
                continue;
            }
            let Ok(bytes) = std::fs::read(&path) else { continue };
            let Ok(cfb) = office_cfb::Cfb::parse(bytes) else { continue };
            // Skip protected decks: their atoms are still encrypted.
            if let Ok(user) = cfb.stream("Current User") {
                if let Some(at) = user.windows(4).position(|w| w == b"\x5f\xc0\x91\xe3") {
                    let _ = at;
                } else if user.len() >= 20 && u32_at(&user, 12) == Some(0xF3D1_C4DF) {
                    continue;
                }
            }
            let Ok(stream) = cfb.stream("PowerPoint Document") else { continue };
            checked += check_style_atoms(&stream, 0, stream.len(), 0, &mut 0usize, &path);
        }
        assert!(checked > 0, "no StyleTextPropAtom found in samples/", );
    }

    fn check_style_atoms(
        data: &[u8],
        start: usize,
        end: usize,
        depth: usize,
        last_cch: &mut usize,
        path: &std::path::Path,
    ) -> usize {
        if depth > MAX_DEPTH {
            return 0;
        }
        let mut found = 0usize;
        for rec in records(data, start, end) {
            match rec.kind {
                RT_TEXT_CHARS_ATOM => *last_cch = rec.len / 2 + 1,
                RT_TEXT_BYTES_ATOM => *last_cch = rec.len + 1,
                RT_STYLE_TEXT_PROP_ATOM => {
                    let payload = rec.payload(data);
                    let consumed = style_runs_len(payload, *last_cch);
                    assert_eq!(
                        consumed,
                        payload.len(),
                        "{}: StyleTextPropAtom at {} consumed {consumed} of {} bytes (cch {})",
                        path.display(),
                        rec.offset,
                        payload.len(),
                        *last_cch
                    );
                    found += 1;
                }
                _ => {}
            }
            if rec.is_container() || rec.kind == RT_OFFICEART_CLIENT_TEXTBOX {
                if rec.kind == RT_SLIDE_LIST_WITH_TEXT {
                    continue;
                }
                found += check_style_atoms(
                    data,
                    rec.offset + 8,
                    rec.offset + 8 + rec.len,
                    depth + 1,
                    last_cch,
                    path,
                );
            }
        }
        found
    }
}
