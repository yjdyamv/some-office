//! Writing PowerPoint 97-2003 presentations.
//!
//! The layout follows what a real file looks like (see `docs/ppt.md`): a
//! `Document` container that lists the slides, a `MainMaster`, one `Slide` per
//! page, then the `PersistDirectoryAtom` and `UserEditAtom` that say where each
//! persist object lives. The `Current User` stream points at the last edit.

/// A slide: a title line and zero or more body paragraphs.
#[derive(Clone, Debug, Default)]
pub struct SlideSpec {
    pub title: String,
    pub body: Vec<String>,
    /// Pictures placed on the slide, in order. Each one goes into the `Pictures`
    /// stream and gets a picture shape on the slide.
    pub pictures: Vec<PictureSpec>,
    /// Speaker notes, one line per paragraph (empty means no notes page).
    pub notes: String,
    /// The transition PowerPoint plays when this slide appears.
    pub transition: Option<Transition>,
}

/// How a slide comes onto the screen, and whether it advances by itself.
///
/// The numbers are the ones PowerPoint 2016 writes into its
/// `SlideShowSlideInfoAtom` (0x03F9) — measured, not taken from the enum in
/// MS-PPT, because the visible effect is what matters (`docs/ppt.md` §9.26).
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum TransitionKind {
    /// No transition (PowerPoint calls this "cut").
    Cut,
    /// Horizontal blinds.
    Blinds,
    /// A dissolve through random pixels.
    Dissolve,
    /// A cross-fade.
    Fade,
}

impl TransitionKind {
    /// The `effectType` byte of the atom.
    fn code(self) -> u8 {
        match self {
            TransitionKind::Cut => 0,
            TransitionKind::Blinds => 2,
            TransitionKind::Dissolve => 5,
            TransitionKind::Fade => 6,
        }
    }

    /// The `effectDirection` byte: blinds open downwards, everything else
    /// ignores the direction.
    fn direction(self) -> u8 {
        match self {
            TransitionKind::Blinds => 1,
            _ => 0,
        }
    }
}

/// A slide transition: the effect plus an optional auto-advance time.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct Transition {
    pub kind: TransitionKind,
    /// Advance after this many milliseconds; `None` waits for a click.
    pub advance_ms: Option<u32>,
}

impl Transition {
    /// A transition that waits for a click.
    pub fn new(kind: TransitionKind) -> Self {
        Transition { kind, advance_ms: None }
    }

    /// The same transition, advancing on its own after `ms` milliseconds.
    pub fn after(self, ms: u32) -> Self {
        Transition { advance_ms: Some(ms), ..self }
    }
}

/// One picture for a slide: the image bytes plus the size it is drawn at, in
/// points (1/72 inch).
#[derive(Clone, Debug)]
pub struct PictureSpec {
    pub data: Vec<u8>,
    pub width: i32,
    pub height: i32,
}

impl PictureSpec {
    pub fn new(data: impl Into<Vec<u8>>, width: i32, height: i32) -> Self {
        PictureSpec { data: data.into(), width, height }
    }
}

impl SlideSpec {
    pub fn new(title: impl Into<String>) -> Self {
        SlideSpec {
            title: title.into(),
            body: Vec::new(),
            pictures: Vec::new(),
            notes: String::new(),
            transition: None,
        }
    }

    /// Give the slide a transition (`docs/ppt.md` §9.26).
    pub fn with_transition(mut self, transition: Transition) -> Self {
        self.transition = Some(transition);
        self
    }

    /// Add speaker notes, one paragraph per line.
    pub fn with_notes(mut self, lines: &[&str]) -> Self {
        self.notes = lines.join("\r");
        self
    }

    pub fn with_body(mut self, lines: &[&str]) -> Self {
        self.body = lines.iter().map(|s| s.to_string()).collect();
        self
    }

    /// Put one picture on the slide, drawn at `width` x `height` points.
    pub fn with_picture(mut self, data: impl Into<Vec<u8>>, width: i32, height: i32) -> Self {
        self.pictures.push(PictureSpec::new(data, width, height));
        self
    }
}

// PowerPoint record types (MS-PPT 2.13).
const RT_DOCUMENT: u16 = 0x03E8;
const RT_END_DOCUMENT: u16 = 0x03EA;
const RT_SLIDE: u16 = 0x03EE;
const RT_SLIDE_ATOM: u16 = 0x03EF;
const RT_SLIDE_SHOW_SLIDE_INFO: u16 = 0x07F0;
/// `SlideShowSlideInfoAtom`: the transition of one slide.
const RT_SLIDE_SHOW_INFO_ATOM: u16 = 0x03F9;
const RT_SLIDE_PERSIST_ATOM: u16 = 0x03F3;
const RT_PPDRAWING: u16 = 0x040C;
const RT_TEXT_HEADER_ATOM: u16 = 0x0F9F;
const RT_TEXT_CHARS_ATOM: u16 = 0x0FA0;
const RT_TEXT_BYTES_ATOM: u16 = 0x0FA8;
const RT_STYLE_TEXT_PROP_ATOM: u16 = 0x0FA1;
const RT_TEXT_SPECIAL_INFO: u16 = 0x0FAA;
const RT_SLIDE_LIST_WITH_TEXT: u16 = 0x0FF0;
const RT_USER_EDIT_ATOM: u16 = 0x0FF5;
const RT_CURRENT_USER: u16 = 0x0FF6;
const RT_PERSIST_DIRECTORY_ATOM: u16 = 0x1772;
const RT_NOTES: u16 = 0x03F0;
const RT_NOTES_ATOM: u16 = 0x03F1;

// Escher record types.
const ESCHER_SPGR: u16 = 0xF009;
const ESCHER_DG_ATOM: u16 = 0xF008;
const ESCHER_DG_CONTAINER: u16 = 0xF002;
const ESCHER_SPGR_CONTAINER: u16 = 0xF003;
const ESCHER_SP_CONTAINER: u16 = 0xF004;
const ESCHER_SP: u16 = 0xF00A;
const ESCHER_FOPT: u16 = 0xF00B;
const ESCHER_CLIENT_TEXTBOX: u16 = 0xF00D;
const ESCHER_CLIENT_ANCHOR: u16 = 0xF010;
const ESCHER_CLIENT_DATA: u16 = 0xF011;

/// `TextHeaderAtom.textType`: 0 title, 1 body, 2 notes, 6 centre title, 5 centre body.
pub const TEXT_TYPE_TITLE: u32 = 6;
pub const TEXT_TYPE_BODY: u32 = 5;
pub const TEXT_TYPE_NOTES: u32 = 2;

/// The 32 byte `SlideShowSlideInfoAtom` every slide and notes page carries.
const SLIDE_SHOW_INFO: [u8; 32] = [
    0xFF, 0xFF, 0xFF, 0x00, 0x00, 0x00, 0x00, 0x00, 0xE8, 0xE8, 0xE8, 0x00, 0x0E, 0x28, 0x41,
    0x00, 0x15, 0x60, 0x82, 0x00, 0xE9, 0x71, 0x32, 0x00, 0x46, 0x78, 0x86, 0x00, 0x96, 0x60,
    0x7D, 0x00,
];

/// Builds one record: `[ver_inst][kind][len]` + payload. Exposed so that callers
/// assembling a whole stream (see `officedump`'s examples) can emit the records
/// this module does not generate yet.
pub fn ppt_public(ver: u16, instance: u16, kind: u16, payload: &[u8]) -> Vec<u8> {
    ppt(ver, instance, kind, payload)
}

fn ppt(ver: u16, instance: u16, kind: u16, payload: &[u8]) -> Vec<u8> {
    let mut out = Vec::with_capacity(8 + payload.len());
    out.extend_from_slice(&((instance << 4) | (ver & 0xF)).to_le_bytes());
    out.extend_from_slice(&kind.to_le_bytes());
    out.extend_from_slice(&(payload.len() as u32).to_le_bytes());
    out.extend_from_slice(payload);
    out
}

fn container(kind: u16, children: &[Vec<u8>]) -> Vec<u8> {
    let payload: Vec<u8> = children.iter().flatten().copied().collect();
    ppt(0xF, 0, kind, &payload)
}

fn escher(instance: u16, kind: u16, payload: &[u8]) -> Vec<u8> {
    ppt(0x0, instance, kind, payload)
}

fn u32b(v: u32) -> [u8; 4] {
    v.to_le_bytes()
}

/// Trimmed `PowerPoint Document` stream used by [`write_presentation`].
const TEMPLATE_STREAM: &[u8] = include_bytes!("../templates/minimal.bin");

/// Pack a finished `PowerPoint Document` stream into a compound file.
fn pack(stream: &[u8], edit_at: u32, pictures: &[u8]) -> Vec<u8> {
    let user = current_user(edit_at, "Rust");
    // PowerPoint pads the stream to 4096 bytes (measured on `samples/ms/pic.ppt`,
    // whose one 573-byte BLIP sits in a 4096-byte stream). Keeping it that size
    // also keeps it out of the mini stream, where the CFB writer would put it.
    let mut padded = Vec::new();
    if !pictures.is_empty() {
        padded = pictures.to_vec();
        padded.resize(pictures.len().max(4096), 0);
    }
    let mut streams = vec![
        office_cfb::OutputStream::new("PowerPoint Document", stream),
        office_cfb::OutputStream::new("Current User", &user),
    ];
    if !padded.is_empty() {
        streams.push(office_cfb::OutputStream::new("Pictures", &padded));
    }
    office_cfb::write(&streams)
}

/// Build a deck by cloning the shapes of a template's first slide.
///
/// The Escher shape tree of a slide carries bookkeeping (FOPT property ids,
/// `ClientData` blocks, the drawing-group shape ids) that PowerPoint validates
/// strictly but that has no documented minimal form. Copying a slide's shape
/// tree and swapping the text keeps all of that intact.
pub fn write_from_slide_template(
    template: &[u8],
    slides: &[SlideSpec],
) -> Result<Vec<u8>, crate::Error> {
    let cfb = office_cfb::Cfb::parse(template.to_vec())?;
    let stream = cfb.stream("PowerPoint Document")?.to_vec();
    let base = cfb.stream("Pictures").map(|d| d.to_vec()).unwrap_or_default();
    let (out, edit_at, pictures) = build_from_stream(&stream, slides, &base)?;
    Ok(pack(&out, edit_at, &pictures))
}

/// Build a `.ppt` from a list of slides.
///
/// Only the `Document` skeleton (environment, style, theme) and the two
/// master-side objects come from the template baked into this crate (a trimmed
/// deck written by Microsoft PowerPoint, trimmed by `examples/trimtemplate.rs`);
/// every slide is a fresh [`slide_record`] holding [`text_shape`] and
/// [`picture_shape`] records — no shape tree is cloned (docs/ppt.md §9.24).
///
/// The generated slide points its `masterIdRef` at the template's first layout,
/// which the deck carries as persist object 3.
pub fn write_presentation(slides: &[SlideSpec]) -> Result<Vec<u8>, crate::Error> {
    if slides.is_empty() {
        return Err(crate::Error::NoDocument);
    }
    let stream = TEMPLATE_STREAM;
    let top = crate::records(stream, 0, stream.len());
    let dir = persist_entries(stream, &top).ok_or(crate::Error::NoDocument)?;
    let blob = |id: u32| -> Option<Vec<u8>> {
        let at = dir.iter().find(|(i, _)| *i == id).map(|(_, o)| *o)?;
        let rec = crate::read_record(stream, at, stream.len())?;
        Some(stream[at..at + 8 + rec.len].to_vec())
    };
    let document_blob = blob(1).ok_or(crate::Error::NoDocument)?;
    // The template's master and layout carry the Office 2007 round-trip ZIPs;
    // PowerPoint rebuilds those on save (docs §9.22).
    let master = strip_round_trip(&blob(2).ok_or(crate::Error::NoDocument)?);
    let layout = strip_round_trip(&blob(3).ok_or(crate::Error::NoDocument)?);

    // The Document's atom, style and theme describe the deck as a whole.
    let mut document_atom = Vec::new();
    let mut style = Vec::new();
    let mut theme = Vec::new();
    for rec in crate::records(&document_blob, 8, document_blob.len()) {
        match rec.kind {
            0x03E9 => document_atom = rec.payload(&document_blob).to_vec(),
            0x03F2 => {
                for child in
                    crate::records(&document_blob, rec.offset + 8, rec.offset + 8 + rec.len)
                {
                    if child.kind == 0x0FA3 {
                        style = child.payload(&document_blob).to_vec();
                    }
                }
            }
            0x0428 => theme = rec.payload(&document_blob).to_vec(),
            _ => {}
        }
    }

    // Object 1 is the Document, 2 the main master, 3 the layout every slide
    // names, 4 the notes master; the slides follow from 5 on and the notes pages
    // after them. The template's notes master drawing sits in cluster 3.
    const FIRST_SLIDE_OBJECT: u32 = 5;
    const NOTES_MASTER_OBJECT: u32 = 4;
    const FIRST_CLUSTER: u16 = 4;
    const LAYOUT_SLIDE_ID: u32 = 0x8000_000C;

    // `DocumentAtom.notesMasterPersistIdRef` has to name the notes master.
    if document_atom.len() >= 28 {
        document_atom[24..28].copy_from_slice(&u32b(NOTES_MASTER_OBJECT));
    }

    let mut per_object = vec![0u32, 7, 7, 8];
    let mut blips: Vec<Vec<u8>> = Vec::new();
    let mut blip_refs: Vec<u32> = Vec::new();
    let mut objects: Vec<(u32, Vec<u8>)> = Vec::new();
    let mut slide_atoms: Vec<Vec<u8>> = Vec::new();
    let mut notes: Vec<(u32, Vec<u8>)> = Vec::new();
    let mut notes_atoms: Vec<Vec<u8>> = Vec::new();
    let mut cluster = FIRST_CLUSTER;
    for (index, spec) in slides.iter().enumerate() {
        let base = (cluster as u32) << 10;
        let mut next = base + 1;
        let mut shapes: Vec<Vec<u8>> = Vec::new();
        if !spec.title.is_empty() {
            shapes.push(text_shape(
                next,
                TEXT_TYPE_TITLE,
                &spec.title,
                [0x02F3, 0x1B10, 0x0893],
                44,
                true,
            ));
            next += 1;
        }
        // Body paragraphs share one placeholder, separated by paragraph marks,
        // exactly like the cloned path (`write_from_slide_template`).
        let body = spec.body.join("\r");
        if !body.is_empty() {
            shapes.push(text_shape(
                next,
                TEXT_TYPE_BODY,
                &body,
                [0x08C3, 0x1B10, 0x0C62],
                24,
                false,
            ));
            next += 1;
        }
        for picture in &spec.pictures {
            let blip = blip_record(&picture.data);
            let pib = match blips.iter().position(|b| *b == blip) {
                Some(index) => {
                    blip_refs[index] += 1;
                    index + 1
                }
                None => {
                    blips.push(blip);
                    blip_refs.push(1);
                    blips.len()
                }
            };
            shapes.push(picture_shape(next, pib as u32, picture.width, picture.height));
            next += 1;
        }

        let notes_id = FIRST_SLIDE_OBJECT + slides.len() as u32 + notes.len() as u32;
        let notes_ref = if spec.notes.is_empty() {
            0
        } else {
            0x100 + index as u32
        };
        let slide = slide_record_with_transition(
            cluster,
            LAYOUT_SLIDE_ID,
            notes_ref,
            7,
            &shapes,
            spec.transition,
        );
        // The cluster issued one id per shape, plus the group's own id.
        per_object.push(next - base);
        cluster += 1;
        let id = FIRST_SLIDE_OBJECT + index as u32;
        objects.push((id, slide));
        slide_atoms.push(ppt(
            0,
            0,
            0x03F3,
            &slide_persist_atom_full(id, 0, 0x100 + index as u32),
        ));

        if !spec.notes.is_empty() {
            let page = notes_record(cluster, 0x100 + index as u32, &spec.notes);
            // The group plus the one text shape.
            per_object.push(2);
            cluster += 1;
            notes.push((notes_id, page));
            notes_atoms.push(ppt(
                0,
                0,
                0x03F3,
                &slide_persist_atom_full(notes_id, 4, 0x100 + index as u32),
            ));
        }
    }

    let master_atoms = vec![
        ppt(0, 0, 0x03F3, &slide_persist_atom_full(2, 0, 0x8000_0000)),
        ppt(0, 0, 0x03F3, &slide_persist_atom_full(3, 0, LAYOUT_SLIDE_ID)),
    ];
    let store: Vec<(u32, u32)> = blips.iter().zip(&blip_refs).map(|(b, r)| (b.len() as u32, *r)).collect();
    let document = document_with(
        &per_object,
        &store,
        DeckLists { masters: &master_atoms, slides: &slide_atoms, notes: &notes_atoms },
        &style,
        &theme,
        &document_atom,
    );

    let notes_master = strip_round_trip(&blob(NOTES_MASTER_OBJECT).ok_or(crate::Error::NoDocument)?);
    let seed = FIRST_SLIDE_OBJECT + slides.len() as u32 + notes.len() as u32;
    let mut ordered: Vec<(u32, Vec<u8>)> =
        vec![(1, document), (2, master), (3, layout), (NOTES_MASTER_OBJECT, notes_master)];
    ordered.extend(objects);
    ordered.extend(notes);
    let mut out = Vec::new();
    let mut offsets = Vec::new();
    for (_, body) in &ordered {
        offsets.push(out.len() as u32);
        out.extend_from_slice(body);
    }
    let mut pd = Vec::new();
    pd.extend_from_slice(&(1u32 | ((ordered.len() as u32) << 20)).to_le_bytes());
    for off in &offsets {
        pd.extend_from_slice(&off.to_le_bytes());
    }
    let pd_at = out.len();
    out.extend_from_slice(&ppt(0, 0, RT_PERSIST_DIRECTORY_ATOM, &pd));

    let mut ue = Vec::new();
    ue.extend_from_slice(&0u32.to_le_bytes()); // lastSlideIdRef
    ue.extend_from_slice(&0x460Cu16.to_le_bytes());
    ue.extend_from_slice(&[0, 3]);
    ue.extend_from_slice(&0u32.to_le_bytes()); // offsetLastEdit
    ue.extend_from_slice(&(pd_at as u32).to_le_bytes());
    ue.extend_from_slice(&1u32.to_le_bytes()); // docPersistIdRef
    ue.extend_from_slice(&seed.to_le_bytes());
    ue.extend_from_slice(&1u16.to_le_bytes()); // lastView
    ue.extend_from_slice(&0u16.to_le_bytes());
    let ue_at = out.len();
    out.extend_from_slice(&ppt(0, 0, RT_USER_EDIT_ATOM, &ue));

    let pictures: Vec<u8> = blips.iter().flatten().copied().collect();
    Ok(pack(&out, ue_at as u32, &pictures))
}

/// Rebuild a presentation from a template stream, reusing every persist object.
///
/// Returns the new `PowerPoint Document` stream plus the offset of its
/// `UserEditAtom` (which `Current User` has to point at).
/// Records that only hold XML for Office 2007's round-trip store: the theme,
/// master/layout/shape drawing XML, and table styles. PowerPoint rebuilds them
/// on save and opens decks without them (docs §9.22), so the writer drops them
/// rather than carrying tens of kilobytes of the template's ZIP blobs.
const ROUND_TRIP_STORE: [u16; 5] = [0x040E, 0x0423, 0x041E, 0x0428, 0xC3A9];

/// Drops the round-trip XML records from one persist object and fixes the
/// container length. Atoms pass through untouched.
fn strip_round_trip(object: &[u8]) -> Vec<u8> {
    if object.len() < 8 {
        return object.to_vec();
    }
    let vi = u16::from_le_bytes([object[0], object[1]]);
    if vi & 0xF != 0xF {
        return object.to_vec();
    }
    let mut kids = Vec::new();
    for r in crate::records(object, 8, object.len()) {
        if ROUND_TRIP_STORE.contains(&r.kind) {
            continue;
        }
        kids.extend_from_slice(&object[r.offset..r.offset + 8 + r.len]);
    }
    if kids.len() == object.len() - 8 {
        return object.to_vec();
    }
    let mut out = object[0..8].to_vec();
    out[4..8].copy_from_slice(&(kids.len() as u32).to_le_bytes());
    out.extend_from_slice(&kids);
    out
}

fn build_from_stream(
    stream: &[u8],
    slides: &[SlideSpec],
    base_pictures: &[u8],
) -> Result<(Vec<u8>, u32, Vec<u8>), crate::Error> {
    let top = crate::records(stream, 0, stream.len());
    let dir = persist_entries(stream, &top).ok_or(crate::Error::NoDocument)?;

    // Objects in the template, in persist-directory order, with their ids.
    let mut objects: Vec<(u32, usize)> = Vec::new();
    for (index, (id, offset)) in dir.iter().enumerate() {
        let end = dir
            .get(index + 1)
            .map(|(_, next)| *next)
            .unwrap_or_else(|| {
                top.iter()
                    .find(|r| r.kind == RT_PERSIST_DIRECTORY_ATOM)
                    .map(|r| r.offset)
                    .unwrap_or(stream.len())
            });
        objects.push((*id, *offset));
        let _ = end;
    }
    let bounds = |i: usize| -> (usize, usize) {
        let start = objects[i].1;
        let end = if i + 1 < objects.len() {
            objects[i + 1].1
        } else {
            top.iter()
                .find(|r| r.kind == RT_PERSIST_DIRECTORY_ATOM)
                .map(|r| r.offset)
                .unwrap_or(stream.len())
        };
        (start, end.min(stream.len()))
    };

    // Which object is the Document, the MainMaster, and each Slide?
    let kind_at = |i: usize| -> u16 {
        let (s, _) = bounds(i);
        crate::read_record(stream, s, stream.len()).map(|r| r.kind).unwrap_or(0)
    };
    let document_i = (0..objects.len())
        .find(|i| kind_at(*i) == RT_DOCUMENT)
        .ok_or(crate::Error::NoDocument)?;
    let slide_is: Vec<usize> = (0..objects.len()).filter(|i| kind_at(*i) == RT_SLIDE).collect();
    if slide_is.is_empty() {
        return Err(crate::Error::NoDocument);
    }

    // The template's first slide supplies the shape tree we clone.
    let (shape_at, shape_end) = {
        let i = slide_is[0];
        let (s, e) = bounds(i);
        let rec = crate::read_record(stream, s, e).ok_or(crate::Error::NoDocument)?;
        (s + 8, s + 8 + rec.len)
    };
    // The drawing record inside the slide; the clone rebuilds only this.
    let drawing_range = {
        let mut range: Option<(usize, usize)> = None;
        let mut pos = shape_at;
        while pos + 8 <= shape_end {
            let Some(rec) = crate::read_record(stream, pos, shape_end) else { break };
            if rec.kind == RT_PPDRAWING {
                range = Some((pos, pos + 8 + rec.len));
            }
            pos += 8 + rec.len;
            if rec.len == 0 {
                break;
            }
        }
        range.unwrap_or((shape_at, shape_end))
    };
    let slide_atom = {
        let mut atom: Vec<u8> = Vec::new();
        let mut pos = shape_at;
        while pos + 8 <= shape_end {
            let Some(rec) = crate::read_record(stream, pos, shape_end) else { break };
            if rec.kind == RT_SLIDE_ATOM {
                atom = stream[pos..pos + 8 + rec.len].to_vec();
            }
            pos += 8 + rec.len;
            if rec.len == 0 {
                break;
            }
        }
        if atom.is_empty() {
            slide_atom(2, 0)
        } else {
            atom
        }
    };
    let trailing = {
        let mut out = Vec::new();
        let mut pos = shape_at;
        while pos + 8 <= shape_end {
            let Some(rec) = crate::read_record(stream, pos, shape_end) else { break };
            if rec.kind != RT_SLIDE_ATOM
                && rec.kind != RT_PPDRAWING
                && !ROUND_TRIP_STORE.contains(&rec.kind)
            {
                out.extend_from_slice(&stream[pos..pos + 8 + rec.len]);
            }
            pos += 8 + rec.len;
            if rec.len == 0 {
                break;
            }
        }
        out
    };

    // Rebuild every object; slides are replaced by clones with our text.
    let mut spid = 0x400u32;
    let mut pieces: Vec<(u32, Vec<u8>)> = Vec::new();
    let mut next_id = objects.iter().map(|(id, _)| *id).max().unwrap_or(1) + 1;
    // Pictures: `pib` is 1-based and shared by the whole deck, the stream is an
    // `OfficeArtBStoreContainer` of `OfficeArtFBSE` records.
    // Keep the template's images: the cloned slides still point at them.
    let mut blips: Vec<Vec<u8>> = split_blips(base_pictures);
    let mut refs: Vec<u32> = vec![1; blips.len()];
    let mut fresh: Vec<usize> = Vec::new();
    let mut shapes_added: u32 = 0;
    let mut slide_seen = 0usize;
    for (i, &(id, start)) in objects.iter().enumerate() {
        let (_, end) = bounds(i);
        if i == document_i {
            continue; // written first, below
        }
        if kind_at(i) == RT_SLIDE {
            // Extra template slides are dropped: the caller decides the count.
            if slide_seen >= slides.len() {
                continue;
            }
            let mut used = 0usize;
            let mut drawing = clone_with_text(stream, drawing_range.0, drawing_range.1, &slides[slide_seen], &mut used, &mut spid);
            add_pictures(&mut drawing, &slides[slide_seen], &mut spid, &mut blips, &mut refs, &mut fresh, &mut shapes_added);
            // The declared length must cover the trailing records too, so the
            // payload is assembled by hand instead of via `container`.
            let mut body = slide_atom.clone();
            body.extend_from_slice(&drawing);
            body.extend_from_slice(&trailing);
            pieces.push((id, ppt(0xF, 0, RT_SLIDE, &body)));
            slide_seen += 1;
        } else {
            pieces.push((id, strip_round_trip(&stream[start..end])));
        }
    }
    // More slides than the template has: append fresh objects.
    while slide_seen < slides.len() {
        let mut used = 0usize;
        let mut drawing = clone_with_text(stream, drawing_range.0, drawing_range.1, &slides[slide_seen], &mut used, &mut spid);
        add_pictures(&mut drawing, &slides[slide_seen], &mut spid, &mut blips, &mut refs, &mut fresh, &mut shapes_added);
        let mut body = slide_atom.clone();
        body.extend_from_slice(&drawing);
        body.extend_from_slice(&trailing);
        pieces.push((next_id, ppt(0xF, 0, RT_SLIDE, &body)));
        next_id += 1;
        slide_seen += 1;
    }

    // The Document keeps the template's structural records; only the slide
    // list is regenerated so it references the slide objects in order.
    let (doc_start, doc_end) = bounds(document_i);
    let document = crate::read_record(stream, doc_start, doc_end).ok_or(crate::Error::NoDocument)?;
    let mut boilerplate: Vec<Vec<u8>> = Vec::new();
    let mut masters_at = None;
    let mut slides_at = None;
    let mut pos = doc_start + 8;
    let doc_stop = doc_start + 8 + document.len;
    while pos + 8 <= doc_stop {
        let Some(rec) = crate::read_record(stream, pos, doc_stop) else { break };
        match rec.kind {
            RT_END_DOCUMENT => {}
            RT_SLIDE_LIST_WITH_TEXT if rec.instance == 1 => masters_at = Some(boilerplate.len()),
            RT_SLIDE_LIST_WITH_TEXT if rec.instance == 0 => slides_at = Some(boilerplate.len()),
            RT_SLIDE_LIST_WITH_TEXT => {}
            kind if ROUND_TRIP_STORE.contains(&kind) => {}
            _ => boilerplate.push(stream[pos..pos + 8 + rec.len].to_vec()),
        }
        pos += 8 + rec.len;
        if rec.len == 0 {
            break;
        }
    }
    // Master list: keep the template's own bytes (it references the master and
    // layout objects we copy verbatim).
    let master_list = {
        let mut found = Vec::new();
        let mut pos = doc_start + 8;
        while pos + 8 <= doc_stop {
            let Some(rec) = crate::read_record(stream, pos, doc_stop) else { break };
            if rec.kind == RT_SLIDE_LIST_WITH_TEXT && rec.instance == 1 {
                found = stream[pos..pos + 8 + rec.len].to_vec();
            }
            pos += 8 + rec.len;
            if rec.len == 0 {
                break;
            }
        }
        found
    };
    let slide_ids: Vec<u32> = pieces
        .iter()
        .filter(|(_, bytes)| crate::read_record(bytes, 0, bytes.len()).map(|r| r.kind) == Some(RT_SLIDE))
        .map(|(id, _)| *id)
        .collect();
    let mut slide_list = Vec::new();
    for (i, id) in slide_ids.iter().enumerate() {
        slide_list.extend(slide_persist_atom(*id, 0x100 + i as u32));
    }
    let at_masters = masters_at.unwrap_or(boilerplate.len()).min(boilerplate.len());
    boilerplate.insert(at_masters, master_list);
    let mut at_slides = slides_at.unwrap_or(boilerplate.len());
    if at_masters <= at_slides {
        at_slides += 1;
    }
    boilerplate.insert(
        at_slides.min(boilerplate.len()),
        ppt(0xF, 0, RT_SLIDE_LIST_WITH_TEXT, &slide_list),
    );
    boilerplate.push(ppt(0, 0, RT_END_DOCUMENT, &[]));

    let doc_id = objects[document_i].0;
    // The document's `PPDrawingGroup` carries the shape-id allocator
    // (`OfficeArtFDGGBlock`, 0xF006): its `spidMax`/`cspidCur` have to grow with
    // every shape we add, or PowerPoint calls the file damaged.
    let added_shapes = shapes_added;
    if added_shapes > 0 {
        let added: Vec<(u32, u32)> = fresh
            .iter()
            .map(|i| (blips[*i].len() as u32, refs[*i]))
            .collect();
        for rec in boilerplate.iter_mut() {
            if rec.len() >= 8 && u16::from_le_bytes([rec[2], rec[3]]) == 0x040B {
                patch_dgg(rec, added_shapes, &added);
            }
        }
    }
    let mut out = container(RT_DOCUMENT, &boilerplate);
    let mut written: Vec<(u32, u32)> = vec![(doc_id, 0)];
    for (id, bytes) in &pieces {
        written.push((*id, out.len() as u32));
        out.extend_from_slice(bytes);
    }
    let mut dir = Vec::new();
    dir.extend_from_slice(&u32b(written[0].0 | ((written.len() as u32) << 20)));
    for (_, offset) in &written {
        dir.extend_from_slice(&u32b(*offset));
    }
    let dir_at = out.len() as u32;
    out.extend(ppt(0, 0, RT_PERSIST_DIRECTORY_ATOM, &dir));
    let mut edit = Vec::new();
    edit.extend_from_slice(&u32b(0x100 + slide_ids.len().saturating_sub(1) as u32));
    edit.extend_from_slice(&0x17DFu16.to_le_bytes());
    edit.push(0);
    edit.push(3);
    edit.extend_from_slice(&u32b(0));
    edit.extend_from_slice(&u32b(dir_at));
    edit.extend_from_slice(&u32b(1));
    edit.extend_from_slice(&u32b(next_id));
    edit.extend_from_slice(&1u16.to_le_bytes());
    edit.extend_from_slice(&0u16.to_le_bytes());
    let edit_at = out.len() as u32;
    out.extend(ppt(0, 0, RT_USER_EDIT_ATOM, &edit));
    let picture_stream: Vec<u8> = blips.concat();
    Ok((out, edit_at, picture_stream))
}


/// One BLIP record. PowerPoint's `Pictures` stream is a plain sequence of these
/// (one per `pib`), with no `OfficeArtBStoreContainer` around them: that is what
/// WPS writes in `samples/image.ppt` and PowerPoint opens it happily.
fn blip_record(data: &[u8]) -> Vec<u8> {
    let kind = match office_blip::PictureKind::sniff(data) {
        office_blip::PictureKind::Png => 0xF01Eu16,
        office_blip::PictureKind::Jpeg => 0xF01Fu16,
        office_blip::PictureKind::Bmp => 0xF02Au16,
        _ => 0xF01Eu16,
    };
    let mut out = Vec::new();
    out.extend_from_slice(&0x6E00u16.to_le_bytes()); // ver 0, instance 0x6E0
    out.extend_from_slice(&kind.to_le_bytes());
    out.extend_from_slice(&((17 + data.len()) as u32).to_le_bytes());
    out.extend_from_slice(&BLIP_UID);
    out.push(0xFF);
    out.extend_from_slice(data);
    out
}

/// The fixed `rgbUid` Word uses for an embedded picture; readers only need it
/// to be stable, not unique.
const BLIP_UID: [u8; 16] = [
    0x2B, 0x0E, 0xFA, 0x72, 0xBC, 0x1B, 0x93, 0x6D, 0x44, 0x2E, 0xE6, 0xCF, 0x3B, 0x43, 0x35, 0xD2,
];

/// A picture shape, copied from a file PowerPoint itself wrote
/// (`samples/ms/pic.ppt`: `OfficeArtSpContainer` with `Sp`, `FOPT` and
/// `ClientAnchor`). Only the shape id and the `pib` are patched: PowerPoint
/// validates the rest of the property set, and neither a guessed minimal set nor
/// the WPS variant (which adds `0x8380`/`0x8381` string properties) survives it.
/// The `spid` of an `SpContainer`, read from its `OfficeArtFSP` child.
fn spid_of(shape: &[u8]) -> Option<u32> {
    let rec = crate::read_record(shape, 0, shape.len())?;
    let fsp = crate::records(shape, 8, 8 + rec.len)
        .into_iter()
        .find(|r| r.kind == ESCHER_SP)?;
    let payload = fsp.payload(shape);
    let bytes: [u8; 4] = payload.get(0..4)?.try_into().ok()?;
    Some(u32::from_le_bytes(bytes))
}

/// Builds one `Slide` container from scratch: a `SlideAtom` plus a `PPDrawing`
/// whose drawing container holds `shapes` inside a group.
///
/// `cluster` is the drawing cluster this slide occupies — the `OfficeArtFDG`
/// record's instance, one past the previous drawing object's — and `master` is
/// the `masterIdRef` of the `SlideAtom`: the *layout's* slide id (`0x8000_0000`
/// for the main master). `shapes` are `SpContainer` records such as
/// [`text_shape`] or [`picture_shape`].
///
/// The document's `FDGG` block (see [`drawing_group`]) has to be told how many
/// shape ids this cluster issued: `shapes.len() + 2` (docs/ppt.md §9.24).
pub fn slide_record(cluster: u16, master: u32, flags: u16, shapes: &[Vec<u8>]) -> Vec<u8> {
    slide_record_with_notes(cluster, master, 0, flags, shapes)
}

/// [`slide_record`] with the `notesIdRef` field set: the id the notes list in
/// the `Document` uses for this page's notes (docs/ppt.md §9.25).
pub fn slide_record_with_notes(
    cluster: u16,
    master: u32,
    notes: u32,
    flags: u16,
    shapes: &[Vec<u8>],
) -> Vec<u8> {
    slide_record_with_transition(cluster, master, notes, flags, shapes, None)
}

/// The `SlideShowSlideInfoAtom` payload: the transition PowerPoint plays.
///
/// Measured on PowerPoint 2016's own files: `slideTime` (ms) first, then
/// `soundIdRef`, then `effectDirection`, `effectType`, a constant 1 and the
/// flags byte where bit 2 means "advance on its own" (`docs/ppt.md` §9.26).
fn slide_show_info(transition: Option<Transition>) -> [u8; 16] {
    let mut out = [0u8; 16];
    let Some(transition) = transition else {
        out[10] = 1;
        return out;
    };
    out[0..4].copy_from_slice(&transition.advance_ms.unwrap_or(0).to_le_bytes());
    out[8] = transition.kind.direction();
    out[9] = transition.kind.code();
    out[10] = 1;
    if transition.advance_ms.is_some() {
        out[11] = 0x04;
    }
    out
}

/// Like [`slide_record_with_notes`], plus a slide transition.
pub fn slide_record_with_transition(
    cluster: u16,
    master: u32,
    notes: u32,
    flags: u16,
    shapes: &[Vec<u8>],
    transition: Option<Transition>,
) -> Vec<u8> {
    let atom = slide_atom_full(0x10, [0; 8], master, notes, flags);
    let mut body = ppt(2, 0, RT_SLIDE_ATOM, &atom);
    // PowerPoint writes its transition atom right after the slide atom; a
    // slide with no transition simply has none.
    if transition.is_some() {
        body.extend_from_slice(&ppt(0, 0, RT_SLIDE_SHOW_INFO_ATOM, &slide_show_info(transition)));
    }
    body.extend_from_slice(&container(RT_PPDRAWING, &[dg_container(cluster, shapes)]));
    // Every real slide carries the slide-show flags after its drawing.
    body.extend_from_slice(&ppt(0, 1, RT_SLIDE_SHOW_SLIDE_INFO, &SLIDE_SHOW_INFO));
    ppt(0xF, 0, RT_SLIDE, &body)
}

/// A notes page for one slide: its own drawing (a group plus one text shape
/// holding the notes) and, like a slide, the slide-show flags (docs §9.25).
///
/// `slide_id` is the `slideId` of the page the notes belong to (`0x100 + n`),
/// which the `SlideAtom.notesIdRef` field and the notes list in the `Document`
/// both have to agree with.
pub fn notes_record(cluster: u16, slide_id: u32, text: &str) -> Vec<u8> {
    let base = (cluster as u32) << 10;
    // The notes placeholder sits below the slide thumbnail, as in a real deck.
    let shape = text_shape(
        base + 1,
        TEXT_TYPE_NOTES,
        text,
        [0x0AD4, 0x01B0, 0x0F30],
        12,
        false,
    );
    let mut atom = Vec::new();
    atom.extend_from_slice(&u32b(slide_id));
    atom.extend_from_slice(&u32b(7));
    let mut body = ppt(1, 0, RT_NOTES_ATOM, &atom);
    body.extend_from_slice(&container(
        RT_PPDRAWING,
        &[dg_container(cluster, std::slice::from_ref(&shape))],
    ));
    body.extend_from_slice(&ppt(0, 1, RT_SLIDE_SHOW_SLIDE_INFO, &SLIDE_SHOW_INFO));
    ppt(0xF, 0, RT_NOTES, &body)
}

/// The `DgContainer` (F002) of one drawing: a group shape owning the cluster,
/// the `FDG` that counts them, and `shapes`.
fn dg_container(cluster: u16, shapes: &[Vec<u8>]) -> Vec<u8> {
    let base = (cluster as u32) << 10;
    let spid_max = shapes
        .iter()
        .filter_map(|s| spid_of(s))
        .max()
        .unwrap_or(base | shapes.len() as u32);

    // Every drawing starts with a group shape that owns the cluster.
    let mut group = Vec::new();
    group.extend_from_slice(&ppt(1, 0, ESCHER_SPGR, &[0u8; 16]));
    let mut group_fsp = Vec::new();
    group_fsp.extend_from_slice(&u32b(base));
    group_fsp.extend_from_slice(&u32b(0x05));
    group.extend_from_slice(&ppt(2, 0, ESCHER_SP, &group_fsp));

    let mut group_children = vec![container(ESCHER_SP_CONTAINER, &[group])];
    group_children.extend(shapes.iter().cloned());

    // `OfficeArtFDG` counts the `SpContainer`s in the drawing (the group is one)
    // and repeats the highest shape id; the record instance is the cluster.
    let mut fdg = Vec::new();
    fdg.extend_from_slice(&u32b(shapes.len() as u32 + 1));
    fdg.extend_from_slice(&u32b(spid_max));
    let mut dg = Vec::new();
    dg.extend_from_slice(&ppt(0, cluster, ESCHER_DG_ATOM, &fdg));
    dg.extend_from_slice(&container(ESCHER_SPGR_CONTAINER, &group_children));
    container(ESCHER_DG_CONTAINER, &[dg])
}

pub fn picture_shape(spid: u32, pib: u32, width: i32, height: i32) -> Vec<u8> {
    let mut shape = PICTURE_SHAPE_TEMPLATE.to_vec();
    // Layout: container header (8) + `Sp` header (8) + spid (4).
    if shape.len() >= 20 {
        shape[16..20].copy_from_slice(&spid.to_le_bytes());
    }
    // The `FOPT` holds `(pid, value)` pairs of 6 bytes; `0x4104` is the `pib`.
    // Skip the container header: the `Sp`, `FOPT` and `ClientAnchor` records live
    // inside it, so starting at 0 would jump straight past all of them.
    let mut pos = 8;
    while pos + 8 <= shape.len() {
        let kind = u16::from_le_bytes([shape[pos + 2], shape[pos + 3]]);
        let len = u32::from_le_bytes([
            shape[pos + 4],
            shape[pos + 5],
            shape[pos + 6],
            shape[pos + 7],
        ]) as usize;
        if kind == 0xF00B {
            let count = u16::from_le_bytes([shape[pos], shape[pos + 1]]) as usize >> 4;
            let body = pos + 8;
            for i in 0..count {
                let at = body + i * 6;
                if at + 6 > shape.len() {
                    break;
                }
                let pid = u16::from_le_bytes([shape[at], shape[at + 1]]);
                let value = match pid {
                    0x4104 => Some(pib),
                    // `cx` / `cy` are in EMU (1 pt = 12700).
                    0x0081 => Some(width as u32 * 12700),
                    0x0082 => Some(height as u32 * 12700),
                    _ => None,
                };
                if let Some(value) = value {
                    shape[at + 2..at + 6].copy_from_slice(&value.to_le_bytes());
                }
            }
        } else if kind == 0xF010 {
            // `ClientAnchor` for a .ppt picture is 8 bytes with **no flags**:
            // `top, left, right, bottom`, all in **1/8 pt**. Measured on
            // `samples/ms/pic3.ppt`, where PowerPoint placed the same image at
            // 40/100 (96x96) and 300/100 (96x48):
            //   `20 03 | 40 01 | 40 04 | 20 06` = 100, 40, 136, 196 pt
            //   `20 03 | 60 09 | 60 0c | a0 04` = 100, 300, 396, 148 pt
            let at = pos + 8;
            if shape.len() >= at + 8 {
                let top = i16::from_le_bytes([shape[at], shape[at + 1]]) as i64;
                let left = i16::from_le_bytes([shape[at + 2], shape[at + 3]]) as i64;
                let right = left + width as i64 * 8;
                let bottom = top + height as i64 * 8;
                shape[at + 4..at + 6].copy_from_slice(&(right as i16).to_le_bytes());
                shape[at + 6..at + 8].copy_from_slice(&(bottom as i16).to_le_bytes());
            }
        }
        pos += 8 + len;
    }
    shape
}

/// `OfficeArtSpContainer` of a picture shape, taken from `samples/ms/pic.ppt`.
const PICTURE_SHAPE_TEMPLATE: &[u8] = include_bytes!("../templates/picture-shape.bin");

/// Split a `Pictures` stream into its individual BLIP records.
fn split_blips(pictures: &[u8]) -> Vec<Vec<u8>> {
    let mut out = Vec::new();
    let mut pos = 0;
    while pos + 8 <= pictures.len() {
        let kind = u16::from_le_bytes([pictures[pos + 2], pictures[pos + 3]]);
        let len = u32::from_le_bytes([
            pictures[pos + 4],
            pictures[pos + 5],
            pictures[pos + 6],
            pictures[pos + 7],
        ]) as usize;
        let stop = (pos + 8 + len).min(pictures.len());
        if office_blip::is_blip(kind) {
            out.push(pictures[pos..stop].to_vec());
        }
        pos = stop;
    }
    out
}

/// Offset of the first record of `kind` at the top level of `data[start..]`.
fn find_record(data: &[u8], start: usize, kind: u16) -> Option<usize> {
    let mut pos = start;
    while pos + 8 <= data.len() {
        let k = u16::from_le_bytes([data[pos + 2], data[pos + 3]]);
        let len = u32::from_le_bytes([
            data[pos + 4],
            data[pos + 5],
            data[pos + 6],
            data[pos + 7],
        ]) as usize;
        if k == kind {
            return Some(pos);
        }
        pos += 8 + len;
    }
    None
}

/// The 36-byte `OfficeArtFBSE` PowerPoint writes into the document drawing group.
fn doc_fbse(image_len: u32, refs: u32) -> Vec<u8> {
    let mut payload = Vec::new();
    payload.extend_from_slice(&[0x06, 0x06]); // ver_inst
    payload.extend_from_slice(&BLIP_UID);
    payload.extend_from_slice(&[0xFF, 0x00]); // tag
    payload.extend_from_slice(&image_len.to_le_bytes());
    payload.extend_from_slice(&refs.to_le_bytes()); // cRef
    payload.extend_from_slice(&0u32.to_le_bytes()); // foDelay
    // unused (2), cbName (1), unused (1) — PowerPoint ends the 36-byte payload
    // with `00 00 ff ff` (measured on `samples/ms/pic.ppt`).
    payload.extend_from_slice(&[0x00, 0x00, 0xFF, 0xFF]);
    let mut rec = Vec::new();
    rec.extend_from_slice(&0x0062u16.to_le_bytes());
    rec.extend_from_slice(&0xF007u16.to_le_bytes());
    rec.extend_from_slice(&(payload.len() as u32).to_le_bytes());
    rec.extend_from_slice(&payload);
    rec
}

/// Register `added` new shapes in the document's drawing group.
///
/// `PPDrawingGroup` → `DggContainer` holds three things PowerPoint checks when a
/// shape is added: the `OfficeArtFDGGBlock` (0xF006) counters, the cluster table
/// that records how many ids each cluster issued, and an `OfficeArtBStoreContainer`
/// (0xF001) listing the images the document holds. A real file that gains a
/// picture grows all three (measured on `samples/ms/minimal.ppt` vs `pic.ppt`).
fn patch_dgg(rec: &mut Vec<u8>, added: u32, blips: &[(u32, u32)]) {
    let Some(dgg_at) = find_record(rec, 8, 0xF000) else {
        return;
    };
    let dgg_body = dgg_at + 8;
    if let Some(at) = find_record(rec, dgg_body, 0xF006) {
        let pay = at + 8;
        if pay + 16 <= rec.len() {
            // [0] next free shape id, [8] ids issued, [12] cluster count.
            for field in [0usize, 8] {
                let v = u32::from_le_bytes(rec[pay + field..pay + field + 4].try_into().unwrap());
                rec[pay + field..pay + field + 4].copy_from_slice(&(v + added).to_le_bytes());
            }
            let cidcl = u32::from_le_bytes(rec[pay + 12..pay + 16].try_into().unwrap()) as usize;
            if cidcl > 0 {
                let last = pay + 16 + (cidcl - 1) * 8 + 4;
                if last + 4 <= rec.len() {
                    let v = u32::from_le_bytes(rec[last..last + 4].try_into().unwrap());
                    rec[last..last + 4].copy_from_slice(&(v + added).to_le_bytes());
                }
            }
        }
    }
    if blips.is_empty() {
        return;
    }
    // One `OfficeArtFBSE` per image: `size` is that BLIP's length and `cRef` how
    // many shapes point at it (measured on `samples/ms/pic2.ppt`, where two
    // shapes share one BLIP and PowerPoint writes `cRef = 2`).
    let mut fbses = Vec::new();
    for (len, refs) in blips {
        fbses.push(doc_fbse(*len, *refs));
    }
    let body = fbses.concat();
    match find_record(rec, dgg_body, 0xF001) {
        // Append to the store the template already has.
        Some(at) => {
            let len = u32::from_le_bytes(rec[at + 4..at + 8].try_into().unwrap()) as usize;
            let end = at + 8 + len;
            let delta = body.len() as u32;
            rec.splice(end..end, body.iter().copied());
            rec[at + 4..at + 8].copy_from_slice(&(len as u32 + delta).to_le_bytes());
            for offset in [dgg_at, 0] {
                let v = u32::from_le_bytes(rec[offset + 4..offset + 8].try_into().unwrap());
                rec[offset + 4..offset + 8].copy_from_slice(&(v + delta).to_le_bytes());
            }
        }
        None => {
            let store = ppt(0xF, 1, 0xF001, &body); // instance 1, as PowerPoint writes it
            let at = find_record(rec, dgg_body, 0xF00B).unwrap_or(dgg_body + 8);
            let delta = store.len() as u32;
            rec.splice(at..at, store.iter().copied());
            for offset in [dgg_at, 0] {
                let v = u32::from_le_bytes(rec[offset + 4..offset + 8].try_into().unwrap());
                rec[offset + 4..offset + 8].copy_from_slice(&(v + delta).to_le_bytes());
            }
        }
    }
}

/// The shape id a new shape has to take: one past the drawing group's last id.
fn drawing_group_next(drawing: &[u8]) -> Option<u32> {
    fn walk(data: &[u8], start: usize, end: usize) -> Option<u32> {
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
            let stop = (pos + 8 + len).min(end);
            if kind == 0xF008 && len >= 8 {
                let last = u32::from_le_bytes([
                    data[pos + 12],
                    data[pos + 13],
                    data[pos + 14],
                    data[pos + 15],
                ]);
                return Some(last + 1);
            }
            if ver_inst & 0xF == 0xF {
                if let Some(found) = walk(data, pos + 8, stop) {
                    return Some(found);
                }
            }
            pos = stop;
        }
        None
    }
    walk(drawing, 0, drawing.len())
}

/// Keep the slide's drawing group in step with the shapes it holds: the
/// `OfficeArtFDG` (0xF008) carries the shape count and the last shape id.
fn patch_drawing_group(drawing: &mut [u8], spid: u32) {
    fn walk(data: &mut [u8], start: usize, end: usize, spid: u32) -> bool {
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
            let stop = (pos + 8 + len).min(end);
            if kind == 0xF008 && len >= 8 {
                let count = u32::from_le_bytes([
                    data[pos + 8],
                    data[pos + 9],
                    data[pos + 10],
                    data[pos + 11],
                ]);
                data[pos + 8..pos + 12].copy_from_slice(&(count + 1).to_le_bytes());
                data[pos + 12..pos + 16].copy_from_slice(&spid.to_le_bytes());
                return true;
            }
            if ver_inst & 0xF == 0xF && walk(data, pos + 8, stop, spid) {
                return true;
            }
            pos = stop;
        }
        false
    }
    walk(drawing, 0, drawing.len(), spid);
}

/// Append one picture shape per `PictureSpec` to the slide's shape group.
///
/// Images are deduplicated: `pib` is the 1-based index into the `Pictures`
/// stream, and `refs` counts how many shapes point at each BLIP (PowerPoint
/// writes `cRef = 2` when two shapes share one image, measured on
/// `samples/ms/pic2.ppt`).
fn add_pictures(
    drawing: &mut Vec<u8>,
    spec: &SlideSpec,
    spid: &mut u32,
    blips: &mut Vec<Vec<u8>>,
    refs: &mut Vec<u32>,
    fresh: &mut Vec<usize>,
    shapes: &mut u32,
) {
    for picture in &spec.pictures {
        let blip = blip_record(&picture.data);
        let pib = match blips.iter().position(|b| *b == blip) {
            Some(index) => {
                refs[index] += 1;
                index + 1
            }
            None => {
                blips.push(blip);
                refs.push(1);
                fresh.push(blips.len() - 1);
                blips.len()
            }
        };
        // The drawing group records the last shape id; a shape with an id below
        // the existing ones makes PowerPoint call the file damaged.
        let shape_id = drawing_group_next(drawing).unwrap_or(*spid + 1);
        *spid = shape_id;
        let shape = picture_shape(shape_id, pib as u32, picture.width, picture.height);
        if splice_into_spgr(drawing, &shape) {
            patch_drawing_group(drawing, shape_id);
            *shapes += 1;
        }
    }
}

/// Insert `shape` at the end of the slide's `OfficeArtSpgrContainer` and grow
/// every enclosing container's declared length by the same amount.
fn splice_into_spgr(drawing: &mut Vec<u8>, shape: &[u8]) -> bool {
    fn walk(
        data: &[u8],
        start: usize,
        end: usize,
        chain: &mut Vec<usize>,
        found: &mut Option<(usize, Vec<usize>)>,
    ) {
        let mut pos = start;
        while pos + 8 <= end {
            let ver_inst = u16::from_le_bytes([data[pos], data[pos + 1]]);
            let kind = u16::from_le_bytes([data[pos + 2], data[pos + 3]]);
            let len = u32::from_le_bytes([data[pos + 4], data[pos + 5], data[pos + 6], data[pos + 7]])
                as usize;
            let stop = (pos + 8 + len).min(end);
            if ver_inst & 0xF == 0xF {
                chain.push(pos);
                if kind == 0xF003 && found.is_none() {
                    *found = Some((stop, chain.clone()));
                }
                walk(data, pos + 8, stop, chain, found);
                chain.pop();
            }
            pos = stop;
        }
    }
    let mut chain = Vec::new();
    let mut found = None;
    walk(drawing, 0, drawing.len(), &mut chain, &mut found);
    let Some((at, chain)) = found else {
        return false;
    };
    let delta = shape.len();
    drawing.splice(at..at, shape.iter().copied());
    for offset in chain {
        if let Some(bytes) = drawing.get(offset + 4..offset + 8) {
            let len = u32::from_le_bytes([bytes[0], bytes[1], bytes[2], bytes[3]]) as usize + delta;
            drawing[offset + 4..offset + 8].copy_from_slice(&(len as u32).to_le_bytes());
        }
    }
    true
}

/// `(persistId, offset)` pairs of the last edit, sorted by offset.
fn persist_entries(stream: &[u8], top: &[crate::Record]) -> Option<Vec<(u32, usize)>> {
    let edit = top.iter().rev().find(|r| r.kind == RT_USER_EDIT_ATOM)?;
    let payload = edit.offset + 8;
    let dir_at = u32::from_le_bytes(stream.get(payload + 12..payload + 16)?.try_into().ok()?) as usize;
    let dir = crate::read_record(stream, dir_at, stream.len())?;
    if dir.kind != RT_PERSIST_DIRECTORY_ATOM {
        return None;
    }
    let body = dir.offset + 8;
    let end = body + dir.len;
    let mut out = Vec::new();
    let mut pos = body;
    while pos + 4 <= end {
        let word = u32::from_le_bytes(stream[pos..pos + 4].try_into().ok()?);
        let first = word & 0xF_FFFF;
        let mut count = ((word >> 20) & 0xFFF) as usize;
        if count == 0 {
            count = 1;
        }
        if 4 + count * 4 > end - pos {
            count = (end - pos - 4) / 4;
        }
        pos += 4;
        for i in 0..count {
            if pos + 4 > end {
                break;
            }
            let offset = u32::from_le_bytes(stream[pos..pos + 4].try_into().ok()?) as usize;
            out.push((first + i as u32, offset));
            pos += 4;
        }
    }
    out.sort_by_key(|(_, offset)| *offset);
    Some(out)
}

/// PowerPoint rejects a **zero-length** text atom (measured: a slide whose body
/// placeholder has no text made the whole file "damaged"); an empty paragraph is
/// written as a lone `\r`.
fn text_atom_body(text: String) -> String {
    if text.is_empty() {
        String::from('\r')
    } else {
        text
    }
}

/// Copy a record range, replacing the text inside every `ClientTextbox`.
fn clone_with_text(
    data: &[u8],
    start: usize,
    end: usize,
    spec: &SlideSpec,
    used: &mut usize,
    spid: &mut u32,
) -> Vec<u8> {
    if std::env::var("PPT_WRITE_PASSTHROUGH").is_ok() {
        return data[start..end].to_vec();
    }
    let mut out = Vec::new();
    let mut pos = start;
    let mut counts: Option<(usize, usize)> = None;
    while pos + 8 <= end {
        let Some(rec) = crate::read_record(data, pos, end) else { break };
        let body = pos + 8;
        let stop = (body + rec.len).min(end);
        let recurses = rec.kind == RT_PPDRAWING
            || rec.kind == ESCHER_CLIENT_TEXTBOX
            || rec.kind == ESCHER_DG_CONTAINER
            || rec.kind == ESCHER_SPGR_CONTAINER
            || rec.kind == ESCHER_SP_CONTAINER;
        if recurses {
            let rebuilt = clone_with_text(data, body, stop, spec, used, spid);
            out.extend(ppt(rec.ver as u16, rec.instance, rec.kind, &rebuilt));
        } else if rec.kind == ESCHER_SP && rec.len >= 8 {
            // Keep the template's shape ids: PowerPoint allocates them in
            // clusters (see `OfficeArtFDGGBlock`) and renumbering them from a
            // fresh counter puts them outside the allocated range.
            *spid = u32::from_le_bytes([data[body], data[body + 1], data[body + 2], data[body + 3]]);
            out.extend_from_slice(&data[pos..stop]);
        } else if rec.kind == RT_TEXT_BYTES_ATOM {
            // PowerPoint writes ASCII-only text as a byte atom. Keep that type
            // when our text is ASCII, otherwise switch to the UTF-16 atom.
            let text = text_atom_body(pick_text(spec, *used));
            *used += 1;
            counts = Some((stop - body, text.chars().count()));
            if text.is_ascii() {
                out.extend(ppt(rec.ver as u16, rec.instance, rec.kind, text.as_bytes()));
            } else {
                let utf16: Vec<u8> = text.chars().flat_map(|c| (c as u16).to_le_bytes()).collect();
                out.extend(ppt(rec.ver as u16, rec.instance, RT_TEXT_CHARS_ATOM, &utf16));
            }
        } else if rec.kind == RT_TEXT_CHARS_ATOM {
            let text = text_atom_body(pick_text(spec, *used));
            *used += 1;
            counts = Some((rec.len / 2, text.chars().count()));
            let utf16: Vec<u8> = text.chars().flat_map(|c| (c as u16).to_le_bytes()).collect();
            out.extend(ppt(rec.ver as u16, rec.instance, rec.kind, &utf16));
        } else if rec.kind == RT_TEXT_SPECIAL_INFO {
            // `TextSpecialInfoAtom` starts with the character count it covers;
            // PowerPoint cross-checks it against the text atom, and a stale
            // count makes the whole file "damaged". Producers disagree on
            // whether the paragraph mark is included (PowerPoint writes the atom
            // length, WPS writes length + 1 — `samples/hello.ppt` says 14 for a
            // 13-character title), so keep the template's own offset instead of
            // assuming one convention.
            let mut info = data[body..stop].to_vec();
            if let Some((old, new)) = counts {
                if info.len() >= 4 {
                    let before = u32::from_le_bytes(info[0..4].try_into().unwrap());
                    let after = (before as i64 + new as i64 - old as i64).max(0) as u32;
                    info[0..4].copy_from_slice(&after.to_le_bytes());
                }
            }
            out.extend(ppt(rec.ver as u16, rec.instance, rec.kind, &info));
        } else if rec.kind == RT_STYLE_TEXT_PROP_ATOM {
            // The atom holds one PF run and one CF run covering the text; both
            // `count` fields are the character count including the paragraph
            // mark. Re-emitting the two runs with the new count keeps the
            // template's formatting and cannot desync the two arrays.
            let cch = counts.map(|(_, new)| new as u32 + 1).unwrap_or(0);
            match restyle(&data[body..stop], cch) {
                Some(style) => out.extend(ppt(rec.ver as u16, rec.instance, rec.kind, &style)),
                None => out.extend_from_slice(&data[pos..stop]),
            }
        } else {
            out.extend_from_slice(&data[pos..stop]);
        }
        pos += 8 + rec.len;
        if rec.len == 0 {
            break;
        }
    }
    out
}

fn pick_text(spec: &SlideSpec, index: usize) -> String {
    // PowerPoint stores no trailing CR in the text atom: the paragraph break
    // is implied and the style atom's `cch` counts what it covers.
    if index == 0 {
        spec.title.clone()
    } else {
        spec.body.join("\r")
    }
}

/// One `Slide` record: atom, drawing with title/body placeholders, and the
/// per-slide extras PowerPoint expects next to the drawing.
fn slide_persist_atom(persist_id: u32, slide_id: u32) -> Vec<u8> {
    let mut payload = Vec::new();
    payload.extend_from_slice(&u32b(persist_id));
    payload.extend_from_slice(&u32b(4)); // cPersist: PowerPoint writes 4, WPS 0
    payload.extend_from_slice(&u32b(0)); // fIsEcho | fIsNotMaster | fIsDocPublished
    payload.extend_from_slice(&u32b(slide_id));
    payload.extend_from_slice(&u32b(0)); // reserved
    ppt(0, 0, RT_SLIDE_PERSIST_ATOM, &payload)
}

fn slide_atom(master: u32, notes: u32) -> Vec<u8> {
    let mut payload = Vec::new();
    payload.extend_from_slice(&u32b(16)); // geom: 10x7.5in slide
    payload.extend_from_slice(&[0u8; 8]); // rgPlaceholderTypes
    // Real files set the high bit of masterIdRef (0x8000_0000) and write
    // slideFlags = 7; PowerPoint keeps the value, so mirror it.
    payload.extend_from_slice(&u32b(0x8000_0000 | master)); // masterIdRef
    payload.extend_from_slice(&u32b(notes)); // notesIdRef
    payload.extend_from_slice(&7u16.to_le_bytes()); // slideFlags
    payload.extend_from_slice(&0u16.to_le_bytes()); // unused
    ppt(2, 0, RT_SLIDE_ATOM, &payload)
}

/// One text placeholder: `Sp`, `FOPT`, `ClientAnchor`, `ClientData`, and the
/// `ClientTextbox` that carries the text plus its formatting.
///
/// This is the from-scratch shape, with every field taken from a real file (see
/// `docs/ppt.md` §9.21). The library's own writer builds slides by cloning a
/// template's shape tree (`clone_with_text`) instead, because a template already
/// carries the master's placeholders and theme.
pub fn text_shape(
    spid: u32,
    text_type: u32,
    text: &str,
    anchor: [i16; 3],
    size_points: u16,
    centred: bool,
) -> Vec<u8> {
    let mut sp = Vec::new();
    sp.extend_from_slice(&u32b(spid));
    sp.extend_from_slice(&u32b(0x0A20)); // fHaveAnchor | fHaveSpt

    // OfficeArtFOPT: the fifteen properties PowerPoint writes for a slide text
    // shape, then twelve bytes of padding. The padding is load-bearing: a
    // shorter record makes PowerPoint reject the file (§9.21).
    let props: [(u16, u32); 15] = [
        (0x0004, 0),
        (0x007F, 0x01EF_0001), // fFitShapeToText etc.
        (0x0080, 0x108F_CCF0), // fill: theme colour
        (0x0081, 0x0001_6530), // line: width
        (0x0082, 0x0000_B298), // line: colour
        (0x0083, 0x0001_6530),
        (0x0084, 0x0000_B298),
        (0x0085, 0),
        (0x0087, 0),
        (0x0088, 0),
        (0x00BF, 0x0006_0000),
        (0x01FF, 0x0001_0000),
        (0x0301, 0x0000_0803),
        (0xC380, 0x0000_000C), // text direction
        (0x03BF, 0x0002_0000),
    ];
    let mut fopt = Vec::with_capacity(102);
    for (pid, value) in props {
        fopt.extend_from_slice(&pid.to_le_bytes());
        fopt.extend_from_slice(&value.to_le_bytes());
    }
    fopt.resize(102, 0);

    let mut anchor_bytes = Vec::new();
    anchor_bytes.extend_from_slice(&0x0240u16.to_le_bytes()); // flags
    for v in anchor {
        anchor_bytes.extend_from_slice(&(v as u16).to_le_bytes());
    }

    // OfficeArtClientData holds the client-data records real PowerPoint files
    // carry: a 0x0BC3 blob and a per-shape counter.
    let mut counter = Vec::new();
    counter.extend_from_slice(&0xFFFF_FFFFu32.to_le_bytes());
    counter.extend_from_slice(&0x0Fu32.to_le_bytes());
    let client_data = container(
        ESCHER_CLIENT_DATA,
        &[
            ppt(0, 0, 0x0BC3, &counter),
            ppt(0, 0, 0x041F, &u32b(1)),
        ],
    );

    // Both `cch` fields count the paragraph mark, and both atoms follow it: a
    // count one short makes PowerPoint declare the file damaged (§9.20).
    let cch = text.chars().count() as u32 + 1;
    let mut style = Vec::new();
    style.extend_from_slice(&u32b(cch)); // TextPFRun count
    style.extend_from_slice(&0u16.to_le_bytes()); // indentLevel
    style.extend_from_slice(&u32b(0x0000_0800)); // PFMasks: textAlignment
    style.extend_from_slice(&(if centred { 1u16 } else { 0 }).to_le_bytes());
    style.extend_from_slice(&u32b(cch)); // TextCFRun count
    style.extend_from_slice(&u32b(0x0002_0000)); // CFMasks: fontSize
    style.extend_from_slice(&size_points.to_le_bytes());

    // ASCII goes into the byte atom, exactly as PowerPoint writes it.
    let text_atom = if text.is_ascii() {
        ppt(0, 0, RT_TEXT_BYTES_ATOM, text.as_bytes())
    } else {
        let utf16: Vec<u8> = text
            .chars()
            .flat_map(|c| (c as u16).to_le_bytes())
            .collect();
        ppt(0, 0, RT_TEXT_CHARS_ATOM, &utf16)
    };
    let mut textbox: Vec<Vec<u8>> = Vec::new();
    textbox.push(ppt(0, 0, RT_TEXT_HEADER_ATOM, &u32b(text_type)));
    textbox.push(text_atom);
    textbox.push(ppt(0, 0, RT_STYLE_TEXT_PROP_ATOM, &style));
    // TextSpecialInfoAtom: the character count, then the run/paragraph layout
    // PowerPoint writes for a plain one-paragraph box.
    let mut special = Vec::new();
    special.extend_from_slice(&u32b(cch - 1));
    special.extend_from_slice(&[
        0x07, 0x00, 0x00, 0x00, 0x00, 0x00, 0x09, 0x04, 0x04, 0x08, 0x01, 0x00, 0x00, 0x00, 0x07,
        0x00, 0x00, 0x00, 0x00, 0x00, 0x04, 0x08, 0x09, 0x04,
    ]);
    textbox.push(ppt(0, 0, RT_TEXT_SPECIAL_INFO, &special));
    // TextMasterStyleAtom-shaped tab stops; harmless, and PowerPoint writes it.
    textbox.push(ppt(
        0,
        0,
        0x0FA6,
        &[0xF0, 0x00, 0x00, 0x00, 0xB0, 0x01, 0xD0, 0x02, 0xF0, 0x03, 0x10, 0x05],
    ));

    container(
        ESCHER_SP_CONTAINER,
        &[
            ppt(2, 1, ESCHER_SP, &sp),
            ppt(3, 15, ESCHER_FOPT, &fopt),
            escher(0, ESCHER_CLIENT_ANCHOR, &anchor_bytes),
            client_data,
            container(ESCHER_CLIENT_TEXTBOX, &textbox),
        ],
    )
}

pub fn current_user(offset_to_current_edit: u32, name: &str) -> Vec<u8> {
    // MS-PPT 2.3.2: size, headerToken, offsetToCurrentEdit, lenUserName,
    // docFileVersion, major, minor, unused, ANSI name, relVersion,
    // UTF-16 name (lenUserName * 2 bytes, no NUL, no length prefix).
    let ansi = name.as_bytes();
    let mut payload = Vec::new();
    payload.extend_from_slice(&0x14u32.to_le_bytes()); // size
    payload.extend_from_slice(&0xE391_C05Fu32.to_le_bytes()); // headerToken
    payload.extend_from_slice(&offset_to_current_edit.to_le_bytes());
    payload.extend_from_slice(&(ansi.len() as u16).to_le_bytes());
    payload.extend_from_slice(&0x03F4u16.to_le_bytes()); // docFileVersion
    payload.push(3); // majorVersion
    payload.push(0); // minorVersion
    payload.extend_from_slice(&0u16.to_le_bytes()); // unused
    payload.extend_from_slice(ansi);
    payload.extend_from_slice(&8u32.to_le_bytes()); // relVersion
    for ch in name.chars() {
        payload.extend_from_slice(&(ch as u16).to_le_bytes());
    }
    ppt(0, 0, RT_CURRENT_USER, &payload)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn writes_a_deck_the_reader_accepts() {
        let deck = vec![
            SlideSpec::new("WRITER-PPT-001").with_body(&["first body line", "second line"]),
            SlideSpec::new("Second slide"),
        ];
        let bytes = write_presentation(&deck).unwrap();
        let cfb = office_cfb::Cfb::parse(bytes).unwrap();
        let doc = cfb.stream("PowerPoint Document").unwrap().to_vec();
        let user = cfb.stream("Current User").unwrap().to_vec();
        let pres = crate::parse_stream(&doc, Some(&user)).unwrap();
        assert_eq!(pres.slides.len(), 2);
        assert_eq!(pres.slides[0].texts[0].text.trim_end(), "WRITER-PPT-001");
        assert_eq!(pres.slides[0].texts[1].text.trim_end(), "first body line\rsecond line");
        assert_eq!(pres.slides[1].texts[0].text.trim_end(), "Second slide");
        assert!(pres.ordered_by_persist_directory);
    }

    #[test]
    fn records_have_the_expected_headers() {
        let bytes = write_presentation(&[SlideSpec::new("T")]).unwrap();
        let cfb = office_cfb::Cfb::parse(bytes).unwrap();
        let doc = cfb.stream("PowerPoint Document").unwrap().to_vec();
        let first = crate::read_record(&doc, 0, doc.len()).unwrap();
        assert_eq!(first.kind, RT_DOCUMENT);
        assert_eq!(first.ver, 0xF, "the Document is a container");
        let user = cfb.stream("Current User").unwrap().to_vec();
        assert_eq!(&user[12..16], &0xE391_C05Fu32.to_le_bytes());
    }

    /// Every slide is a clone of the template's first shape tree, so a deck
    /// with more pages than the template has reuses those shape ids. PowerPoint
    /// accepts that (verified with 1/3/5/10 pages and with pictures, docs §9.19);
    /// this test pins the structure: one slide object per spec, each with its own
    /// text, distinct slide ids, and a slide list that matches.
    #[test]
    fn extra_slides_are_appended_with_their_own_text() {
        let Ok(tpl) = std::fs::read("../../samples/ms/minimal.ppt") else { return };
        let specs: Vec<SlideSpec> = (1..=5)
            .map(|i| {
                let mut s = SlideSpec::new(format!("Slide {i}"));
                s.body = vec![format!("body {i}")];
                s
            })
            .collect();
        let out = write_from_slide_template(&tpl, &specs).unwrap();
        let parsed = crate::parse_stream(
            &office_cfb::Cfb::parse(out.clone()).unwrap().stream("PowerPoint Document").unwrap().to_vec(),
            None,
        )
        .unwrap();
        assert_eq!(parsed.slides.len(), 5);
        for (i, slide) in parsed.slides.iter().enumerate() {
            assert_eq!(slide.texts[0].text.trim_end(), format!("Slide {}", i + 1));
            assert_eq!(slide.texts[1].text.trim_end(), format!("body {}", i + 1));
            // The clone keeps the template's shapes, so the count is stable.
            assert_eq!(slide.texts.len(), 2, "slide {} lost a placeholder", i + 1);
        }

        // Slide ids in the document's slide list are 0x100, 0x101, ... and each
        // `SlidePersistAtom` points at an object that really is a slide.
        let stream = office_cfb::Cfb::parse(out).unwrap().stream("PowerPoint Document").unwrap().to_vec();
        let doc = crate::records(&stream, 0, stream.len())
            .into_iter()
            .find(|r| r.kind == RT_DOCUMENT)
            .expect("document");
        let body = &stream[doc.offset + 8..doc.offset + 8 + doc.len];
        let mut ids = Vec::new();
        for list in crate::records(body, 0, body.len()) {
            // instance 0 is the slide list; instance 1 lists masters/layouts.
            if list.kind != RT_SLIDE_LIST_WITH_TEXT || list.instance != 0 {
                continue;
            }
            let start = list.offset + 8;
            for atom in crate::records(body, start, start + list.len) {
                if atom.kind == RT_SLIDE_PERSIST_ATOM {
                    let at = atom.offset + 8;
                    ids.push(u32::from_le_bytes(body[at + 12..at + 16].try_into().unwrap()));
                }
            }
        }
        assert_eq!(ids, vec![0x100, 0x101, 0x102, 0x103, 0x104]);
    }

    #[test]
    fn template_slide_clone_keeps_the_shape_tree() {
        // hello.ppt is WPS-written; its shape tree is known-good for our reader.
        let Ok(template) = std::fs::read("../../samples/hello.ppt") else {
            return;
        };
        let deck = vec![SlideSpec::new("CLONE-1").with_body(&["cloned body"])];
        let bytes = write_from_slide_template(&template, &deck).unwrap();
        let cfb = office_cfb::Cfb::parse(bytes).unwrap();
        let doc = cfb.stream("PowerPoint Document").unwrap().to_vec();
        let user = cfb.stream("Current User").ok().map(|d| d.to_vec());
        let pres = crate::parse_stream(&doc, user.as_deref()).unwrap();
        assert_eq!(pres.slides.len(), 1);
        assert_eq!(pres.slides[0].texts[0].text.trim_end(), "CLONE-1");
        assert_eq!(pres.slides[0].texts[1].text.trim_end(), "cloned body");
    }

    /// The `TextSpecialInfoAtom` count follows the template's own convention:
    /// WPS (`hello.ppt`) counts the paragraph mark, PowerPoint (`minimal.ppt`)
    /// does not. Writing the wrong one makes PowerPoint refuse the file.
    #[test]
    fn special_info_count_keeps_the_templates_convention() {
        fn first_special_info_count(bytes: &[u8]) -> Option<u32> {
            let cfb = office_cfb::Cfb::parse(bytes.to_vec()).ok()?;
            let doc = cfb.stream("PowerPoint Document").ok()?.to_vec();
            let slide = crate::records(&doc, 0, doc.len()).into_iter().find(|r| r.kind == RT_SLIDE)?;
            let start = slide.offset;
            let end = start + 8 + slide.len;
            for pos in start..end.saturating_sub(8) {
                if doc[pos + 2..pos + 4] == [0xAA, 0x0F] {
                    let len = u32::from_le_bytes(doc[pos + 4..pos + 8].try_into().ok()?) as usize;
                    if len >= 4 && pos + 12 <= end {
                        return Some(u32::from_le_bytes(doc[pos + 8..pos + 12].try_into().ok()?));
                    }
                }
            }
            None
        }
        let title = "Generated title"; // 15 characters
        for (template, want) in [("../../samples/hello.ppt", 16u32), ("../../samples/ms/minimal.ppt", 15)] {
            let Ok(tpl) = std::fs::read(template) else { continue };
            let out = write_from_slide_template(&tpl, &[SlideSpec::new(title)]).unwrap();
            assert_eq!(
                first_special_info_count(&out),
                Some(want),
                "{template} keeps its own paragraph-mark convention"
            );
        }
    }

    #[test]
    fn writes_a_picture_into_the_pictures_stream() {
        let Ok(png) = std::fs::read("../../samples/pixel.png") else {
            return;
        };
        let deck = vec![
            SlideSpec::new("PIC-1").with_picture(png.clone(), 96, 96),
            SlideSpec::new("PIC-2"),
        ];
        let bytes = write_presentation(&deck).unwrap();
        let cfb = office_cfb::Cfb::parse(bytes).unwrap();
        let pictures = cfb.stream("Pictures").unwrap().to_vec();
        let blips = crate::pictures(&pictures);
        assert_eq!(blips.len(), 1, "one picture in the store");
        assert_eq!(blips[0].kind, crate::PictureKind::Png);
        assert_eq!(blips[0].data, png, "the image survives byte for byte");
        let doc = cfb.stream("PowerPoint Document").unwrap().to_vec();
        let user = cfb.stream("Current User").unwrap().to_vec();
        let pres = crate::parse_stream(&doc, Some(&user)).unwrap();
        assert_eq!(pres.slides[0].pictures, vec![1]);
        assert!(pres.slides[1].pictures.is_empty());
        assert_eq!(pres.slides[0].texts[0].text.trim_end(), "PIC-1");
    }

    #[test]
    fn writes_a_presentation_with_a_picture() {
        // The picture has to be registered in three places PowerPoint checks: the
        // slide's `OfficeArtFDG`, the document's `OfficeArtFDGG` (counters, cluster
        // table, `BStoreContainer`) and the `Pictures` stream itself.
        let png = std::fs::read("../../samples/pixel.png").unwrap();
        let deck = vec![SlideSpec::new("PIC-TEST").with_picture(png.clone(), 96, 96)];
        let bytes = write_presentation(&deck).unwrap();
        let cfb = office_cfb::Cfb::parse(bytes).unwrap();
        let pictures = cfb.stream("Pictures").unwrap().to_vec();
        assert_eq!(pictures.len(), 4096, "padded like PowerPoint does");
        let found = crate::pictures(&pictures);
        assert_eq!(found.len(), 1);
        assert_eq!(found[0].data, png);
        let doc = cfb.stream("PowerPoint Document").unwrap().to_vec();
        let user = cfb.stream("Current User").unwrap().to_vec();
        let pres = crate::parse_stream(&doc, Some(&user)).unwrap();
        assert_eq!(pres.slides[0].pictures, vec![1], "the shape points at the image");
    }

    #[test]
    fn picture_size_comes_from_the_client_anchor() {
        // `ClientAnchor` is `top, left, right, bottom` in 1/8 pt with no flags.
        fn anchors(data: &[u8], start: usize, end: usize, out: &mut Vec<(i64, i64)>) {
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
                let stop = (pos + 8 + len).min(end);
                if kind == 0xF010 && len == 8 {
                    let at = pos + 8;
                    let top = i16::from_le_bytes([data[at], data[at + 1]]) as i64;
                    let left = i16::from_le_bytes([data[at + 2], data[at + 3]]) as i64;
                    let right = i16::from_le_bytes([data[at + 4], data[at + 5]]) as i64;
                    let bottom = i16::from_le_bytes([data[at + 6], data[at + 7]]) as i64;
                    out.push(((right - left) / 8, (bottom - top) / 8));
                }
                if ver_inst & 0xF == 0xF {
                    anchors(data, pos + 8, stop, out);
                }
                pos = stop;
            }
        }
        let png = std::fs::read("../../samples/pixel.png").unwrap();
        let deck = vec![SlideSpec::new("SIZES")
            .with_picture(png.clone(), 96, 96)
            .with_picture(png, 48, 24)];
        let bytes = write_presentation(&deck).unwrap();
        let cfb = office_cfb::Cfb::parse(bytes).unwrap();
        let doc = cfb.stream("PowerPoint Document").unwrap().to_vec();
        let mut sizes = Vec::new();
        anchors(&doc, 0, doc.len(), &mut sizes);
        assert!(sizes.contains(&(96, 96)), "96x96 anchor missing: {sizes:?}");
        assert!(sizes.contains(&(48, 24)), "48x24 anchor missing: {sizes:?}");
    }

    #[test]
    fn two_pictures_share_one_blip() {
        let png = std::fs::read("../../samples/pixel.png").unwrap();
        let deck = vec![SlideSpec::new("PIC-TWO")
            .with_picture(png.clone(), 96, 96)
            .with_picture(png, 64, 64)];
        let bytes = write_presentation(&deck).unwrap();
        let cfb = office_cfb::Cfb::parse(bytes).unwrap();
        let pictures = cfb.stream("Pictures").unwrap().to_vec();
        assert_eq!(crate::pictures(&pictures).len(), 1, "identical images are stored once");
        let doc = cfb.stream("PowerPoint Document").unwrap().to_vec();
        let user = cfb.stream("Current User").unwrap().to_vec();
        let pres = crate::parse_stream(&doc, Some(&user)).unwrap();
        assert_eq!(pres.slides[0].pictures, vec![1, 1], "both shapes use pib 1");
    }

    #[test]
    fn cloning_a_template_keeps_its_pictures() {
        // A deck written by PowerPoint that already holds one picture: the clone
        // must keep the shape, the `Pictures` stream and the `pib` wiring.
        let Ok(template) = std::fs::read("../../samples/ms/pic.ppt") else {
            return;
        };
        let deck = vec![SlideSpec::new("CLONE-PIC").with_body(&["picture kept"])];
        let bytes = write_from_slide_template(&template, &deck).unwrap();
        let cfb = office_cfb::Cfb::parse(bytes).unwrap();
        let pictures = cfb.stream("Pictures").unwrap().to_vec();
        assert_eq!(crate::pictures(&pictures).len(), 1, "the image survives");
        let doc = cfb.stream("PowerPoint Document").unwrap().to_vec();
        let user = cfb.stream("Current User").unwrap().to_vec();
        let pres = crate::parse_stream(&doc, Some(&user)).unwrap();
        assert_eq!(pres.slides.len(), 1);
        assert_eq!(pres.slides[0].pictures, vec![1], "the shape still points at it");
        assert_eq!(pres.slides[0].texts[0].text.trim_end(), "CLONE-PIC");
    }
}

/// Re-emit a single-PF, single-CF `StyleTextPropAtom` with a new character count.
fn restyle(style: &[u8], cch: u32) -> Option<Vec<u8>> {
    if style.len() < 10 || cch == 0 {
        return None;
    }
    let pf_masks = u32::from_le_bytes(style.get(6..10)?.try_into().ok()?);
    let (pf_end, _) = crate::pf_exception(style, 10, pf_masks)?;
    let cf_masks = u32::from_le_bytes(style.get(pf_end + 4..pf_end + 8)?.try_into().ok()?);
    let (cf_end, _) = crate::cf_exception(style, pf_end + 8, cf_masks)?;
    let mut out = Vec::with_capacity(cf_end);
    out.extend_from_slice(&cch.to_le_bytes());
    out.extend_from_slice(&style[4..pf_end]);
    out.extend_from_slice(&cch.to_le_bytes());
    out.extend_from_slice(&style[pf_end + 4..cf_end]);
    Some(out)
}

// --- the id-cluster table of `OfficeArtFDGGBlock` (docs/ppt.md §9.4.2) ------

/// Builds the payload of `OfficeArtFDGGBlock` (0xF006) from the number of ids
/// each persist object uses, in object order.
///
/// PowerPoint refuses a deck whose cluster table does not line up with the
/// documents it holds: `void.ms` (3 objects) carries `(1,7) (2,1)` and a blank
/// deck from the same program (14 objects) carries 13 clusters. Documenting the
/// rule as code is what lets the writer stop copying that table from a template.
pub fn fdgg_payload(per_object: &[u32]) -> Vec<u8> {
    let highest = per_object.len() as u32;
    let mut spid_max = 0u32;
    let mut clusters = Vec::new();
    for (i, used) in per_object.iter().enumerate() {
        if *used == 0 {
            continue;
        }
        spid_max = (i as u32) << 10 | (used - 1);
        clusters.push((i as u32, *used));
    }
    let issued: u32 = clusters.iter().map(|(_, n)| n).sum();
    let mut v = Vec::new();
    v.extend_from_slice(&(spid_max + 1).to_le_bytes());
    v.extend_from_slice(&(highest).to_le_bytes());
    v.extend_from_slice(&issued.to_le_bytes());
    v.extend_from_slice(&(clusters.len() as u32).to_le_bytes());
    for (id, used) in clusters {
        v.extend_from_slice(&id.to_le_bytes());
        v.extend_from_slice(&used.to_le_bytes());
    }
    v
}

#[cfg(test)]
mod fdgg_tests {
    /// `samples/void-ms.ppt`: Document + master (7 shapes) + slide (1 shape).
    #[test]
    fn three_objects_give_two_clusters() {
        let v = super::fdgg_payload(&[0, 7, 1]);
        // The reference file says `01 08 00 00 | 03 00 00 00 | 07 00 00 00 |
        // 02 00 00 00 | (1,7) (2,1)`; the third word ("issued ids") does not
        // follow from the cluster counts (7 vs 8 here, 75 vs 88 in minimal.ppt),
        // so it is left as the plain sum until it is understood.
        assert_eq!(&v[0..4], &0x0801u32.to_le_bytes());
        assert_eq!(&v[4..8], &3u32.to_le_bytes());
        assert_eq!(&v[12..16], &2u32.to_le_bytes());
        assert_eq!(&v[16..24], &[1, 0, 0, 0, 7, 0, 0, 0]);
        assert_eq!(&v[24..32], &[2, 0, 0, 0, 1, 0, 0, 0]);
    }

    /// The payload measured in `samples/void-ms.ppt`, so the rule is checked
    /// against the real file and not only against my own notes.
    #[test]
    fn matches_the_fdgg_inside_a_powerpoint_written_deck() {
        let Ok(bytes) = std::fs::read("../../samples/void-ms.ppt") else {
            return; // samples/ is not shipped with the crate
        };
        let cfb = office_cfb::Cfb::parse(bytes).expect("cfb");
        let stream = cfb.stream("PowerPoint Document").expect("stream");
        let at = stream
            .windows(4)
            .position(|w| w == [0x00, 0x00, 0x06, 0xf0])
            .expect("0xF006 record");
        let len = u32::from_le_bytes([
            stream[at + 4], stream[at + 5], stream[at + 6], stream[at + 7],
        ]) as usize;
        let reference = &stream[at + 8..at + 8 + len];
        let ours = super::fdgg_payload(&[0, 7, 1]);
        assert_eq!(ours.len(), reference.len());
        // The third word ("issued ids") is still undecoded: the file says 7
        // where the cluster counts add up to 8, so it is compared separately.
        assert_eq!(&ours[0..8], &reference[0..8]);
        assert_eq!(&ours[12..16], &reference[12..16]);
        assert_eq!(&ours[16..], &reference[16..]);
    }

    /// `samples/ms/minimal.ppt`: 1 master + 11 layouts + 1 slide over 14 objects;
    /// its last slide uses 4 ids, so `spidMax` is `0x3404`.
    #[test]
    fn fourteen_objects_put_the_slide_in_the_last_cluster() {
        let mut per = vec![0u32, 7];
        per.extend(std::iter::repeat(7).take(11));
        per.push(4);
        let v = super::fdgg_payload(&per);
        assert_eq!(v.len(), 16 + 13 * 8);
        assert_eq!(&v[0..4], &0x3404u32.to_le_bytes());
        assert_eq!(&v[4..8], &14u32.to_le_bytes());
        assert_eq!(&v[12..16], &13u32.to_le_bytes());
        assert_eq!(&v[16..24], &[1, 0, 0, 0, 7, 0, 0, 0]);
        assert_eq!(&v[v.len() - 8..], &[13, 0, 0, 0, 4, 0, 0, 0]);
    }
}

#[cfg(test)]
mod skeleton_tests {
    const CONTAINERS: [u16; 12] = [
        0x03E8, 0x03F8, 0x03EE, 0x040C, 0x040B, 0xF000, 0xF002, 0xF003, 0xF004, 0xF001, 0x03F2, 0x07D0,
    ];

    struct Seen {
        versions: Vec<(u16, u16)>,
        slide_atom_ver: Option<u16>,
        master_style_inst: Vec<u16>,
        layout_inst: Vec<u16>,
    }

    fn walk(d: &[u8], start: usize, end: usize, seen: &mut Seen) {
        let mut at = start;
        while at + 8 <= end {
            let vi = u16::from_le_bytes([d[at], d[at + 1]]);
            let kind = u16::from_le_bytes([d[at + 2], d[at + 3]]);
            let len = u32::from_le_bytes([d[at + 4], d[at + 5], d[at + 6], d[at + 7]]) as usize;
            if at + 8 + len > end {
                break;
            }
            let ver = vi & 0xF;
            if CONTAINERS.contains(&kind) {
                assert_eq!(ver, 0xF, "container 0x{kind:04X} has ver {ver}");
                walk(d, at + 8, at + 8 + len, seen);
            } else if kind == 0x03EF {
                seen.slide_atom_ver = Some(ver);
            } else {
                // Atoms carry their own fixed `ver` (DocumentAtom 1, SlideAtom 2,
                // FOPT 3 …), so only the bookkeeping records are pinned to 0.
                if matches!(kind, 0xF006 | 0x0FF5 | 0x1772 | 0x0428 | 0x0423 | 0x03EA) {
                    assert_eq!(ver, 0, "bookkeeping 0x{kind:04X} at {at} has ver {ver}");
                }
                seen.versions.push((kind, ver));
                if kind == 0x0FA3 {
                    seen.master_style_inst.push(vi >> 4);
                } else if kind == 0x041E {
                    seen.layout_inst.push(vi >> 4);
                }
            }
            at += 8 + len;
        }
    }

    /// The rules `docs/ppt.md` §9.5 records, checked against the real file.
    #[test]
    fn void_ms_follows_the_skeleton_rules() {
        let Ok(bytes) = std::fs::read("../../samples/void-ms.ppt") else {
            return;
        };
        let cfb = office_cfb::Cfb::parse(bytes).expect("cfb");
        let stream = cfb.stream("PowerPoint Document").expect("stream");
        let d: &[u8] = &stream;
        let mut seen = Seen {
            versions: Vec::new(),
            slide_atom_ver: None,
            master_style_inst: Vec::new(),
            layout_inst: Vec::new(),
        };
        walk(d, 0, d.len(), &mut seen);
        assert_eq!(seen.slide_atom_ver, Some(2), "SlideAtom is the one ver = 2 record");
        let ver_of = |kind: u16| seen.versions.iter().find(|(k, _)| *k == kind).map(|(_, v)| *v);
        assert_eq!(ver_of(0x03E9), Some(1), "DocumentAtom");
        assert_eq!(ver_of(0xF00B), Some(3), "FOPT");
        assert_eq!(ver_of(0x03EA), Some(0), "EndDocument");
        // `Environment` carries one `TextMasterStyleAtom` too (instance 4), so
        // the master's eight are the tail of the list.
        assert_eq!(seen.master_style_inst.len(), 9, "{:?}", seen.master_style_inst);
        assert_eq!(&seen.master_style_inst[1..], &[0, 1, 2, 4, 5, 6, 7, 8]);
        assert_eq!(seen.layout_inst.len(), 11);
        assert!(seen.layout_inst.windows(2).all(|w| w[0] > w[1]), "{:?}", seen.layout_inst);
    }
}

// --- atoms whose layout is confirmed against void-ms.ppt (docs/ppt.md §9.6) --

/// `SlideAtom` (0x03EF): 24 bytes. `ver` is 2 for this record, never 0xF.
pub fn slide_atom_full(geom: u32, placeholder_types: [u8; 8], master: u32, notes: u32, flags: u16) -> Vec<u8> {
    let mut v = Vec::with_capacity(24);
    v.extend_from_slice(&geom.to_le_bytes());
    v.extend_from_slice(&placeholder_types);
    v.extend_from_slice(&master.to_le_bytes());
    v.extend_from_slice(&notes.to_le_bytes());
    v.extend_from_slice(&flags.to_le_bytes());
    v.extend_from_slice(&0u16.to_le_bytes());
    v
}

/// `SlidePersistAtom` (0x03F3): 20 bytes. PowerPoint writes `cPersist = 4`.
pub fn slide_persist_atom_full(persist_id: u32, c_persist: u32, slide_id: u32) -> Vec<u8> {
    let mut v = Vec::with_capacity(20);
    v.extend_from_slice(&persist_id.to_le_bytes());
    v.extend_from_slice(&c_persist.to_le_bytes());
    v.extend_from_slice(&0u32.to_le_bytes()); // flags
    v.extend_from_slice(&slide_id.to_le_bytes());
    v.extend_from_slice(&0u32.to_le_bytes()); // reserved
    v
}

#[cfg(test)]
mod atom_tests {
    use super::*;

    fn records(d: &[u8]) -> Vec<(u16, Vec<u8>)> {
        let mut out = Vec::new();
        fn walk(d: &[u8], start: usize, end: usize, out: &mut Vec<(u16, Vec<u8>)>) {
            let mut at = start;
            while at + 8 <= end {
                let vi = u16::from_le_bytes([d[at], d[at + 1]]);
                let kind = u16::from_le_bytes([d[at + 2], d[at + 3]]);
                let len = u32::from_le_bytes([d[at + 4], d[at + 5], d[at + 6], d[at + 7]]) as usize;
                if at + 8 + len > end {
                    break;
                }
                out.push((kind, d[at + 8..at + 8 + len].to_vec()));
                if vi & 0xF == 0xF {
                    walk(d, at + 8, at + 8 + len, out);
                }
                at += 8 + len;
            }
        }
        walk(d, 0, d.len(), &mut out);
        out
    }

    #[test]
    fn slide_atoms_match_the_reference_deck() {
        let Ok(bytes) = std::fs::read("../../samples/void-ms.ppt") else {
            return;
        };
        let cfb = office_cfb::Cfb::parse(bytes).expect("cfb");
        let stream = cfb.stream("PowerPoint Document").expect("stream");
        let atoms: Vec<Vec<u8>> = records(&stream)
            .into_iter()
            .filter(|(k, _)| *k == 0x03EF)
            .map(|(_, p)| p)
            .collect();
        assert_eq!(atoms.len(), 2, "master + one slide");
        // master: geom 1, no placeholders, no master of its own, no flags
        assert_eq!(atoms[0], slide_atom_full(1, [1, 2, 7, 9, 8, 0, 0, 0], 0, 0, 2));
        // slide: geom 16, no placeholder types, master 0x80000000, flags 7
        assert_eq!(atoms[1], slide_atom_full(0x10, [0; 8], 0x8000_0000, 0, 7));
    }

    #[test]
    fn slide_persist_atoms_match_the_reference_deck() {
        let Ok(bytes) = std::fs::read("../../samples/void-ms.ppt") else {
            return;
        };
        let cfb = office_cfb::Cfb::parse(bytes).expect("cfb");
        let stream = cfb.stream("PowerPoint Document").expect("stream");
        let atoms: Vec<Vec<u8>> = records(&stream)
            .into_iter()
            .filter(|(k, _)| *k == 0x03F3)
            .map(|(_, p)| p)
            .collect();
        assert_eq!(atoms.len(), 2, "master list + slide list");
        assert_eq!(atoms[0], slide_persist_atom_full(2, 0, 0x8000_0000));
        assert_eq!(atoms[1], slide_persist_atom_full(3, 0, 0x0102));
    }
}

#[cfg(test)]
mod environment_tests {
    /// `Environment` (0x03F2) is a container of six records; the numbers are
    /// pinned here so `docs/ppt.md` §9.8 and any future builder cannot drift.
    #[test]
    fn environment_children_are_the_six_expected_records() {
        let Ok(bytes) = std::fs::read("../../samples/void-ms.ppt") else {
            return;
        };
        let cfb = office_cfb::Cfb::parse(bytes).expect("cfb");
        let stream = cfb.stream("PowerPoint Document").expect("stream");
        let d: &[u8] = &stream;
        // the top-level Document starts at 0; Environment is its second child
        let mut at = 8usize;
        let mut env = None;
        while at + 8 <= d.len() {
            let kind = u16::from_le_bytes([d[at + 2], d[at + 3]]);
            let len = u32::from_le_bytes([d[at + 4], d[at + 5], d[at + 6], d[at + 7]]) as usize;
            if kind == 0x03F2 {
                env = Some(&d[at + 8..at + 8 + len]);
                break;
            }
            at += 8 + len;
        }
        let env = env.expect("Environment");
        assert_eq!(env.len(), 430);
        let mut children = Vec::new();
        let mut p = 0usize;
        while p + 8 <= env.len() {
            let vi = u16::from_le_bytes([env[p], env[p + 1]]);
            let kind = u16::from_le_bytes([env[p + 2], env[p + 3]]);
            let len = u32::from_le_bytes([env[p + 4], env[p + 5], env[p + 6], env[p + 7]]) as usize;
            children.push((kind, vi >> 4, len));
            p += 8 + len;
        }
        assert_eq!(
            children,
            vec![
                (0x0FC8, 2, 12),   // colour scheme
                (0x07D5, 0, 228),  // font collection: 3 x FontEntity
                (0x0FA4, 0, 10),   // ruler
                (0x0FA5, 0, 12),
                (0x0FA9, 0, 10),
                (0x0FA3, 4, 110),  // the default text style
            ]
        );
    }
}

/// `Environment` (0x03F2) with the given fonts; `style` is the 110 byte
/// `TextMasterStyleAtom` payload (`docs/ppt.md` §9.8).
pub fn environment(fonts: &[(&str, [u8; 4])], style: &[u8]) -> Vec<u8> {
    const COLOUR_SCHEME: [u8; 12] = [0x30, 0x00, 0xd2, 0x0f, 0x04, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00];
    const RULER_0FA4: [u8; 10] = [0x80, 0x00, 0x60, 0x00, 0x00, 0x00, 0xff, 0xff, 0x00, 0x00];
    const RULER_0FA5: [u8; 12] = [0x00, 0x00, 0x00, 0x08, 0x2e, 0x00, 0x00, 0x00, 0x02, 0x00, 0x00, 0x00];
    const RULER_0FA9: [u8; 10] = [0x07, 0x00, 0x00, 0x00, 0x02, 0x00, 0x04, 0x08, 0x00, 0x00];

    let mut body = Vec::new();
    body.extend_from_slice(&ppt(0xF, 2, 0x0FC8, &COLOUR_SCHEME));
    let mut collection = Vec::new();
    for (index, (name, flags)) in fonts.iter().enumerate() {
        let mut f = Vec::with_capacity(68);
        let mut name_bytes = [0u8; 64];
        for (i, u) in name.encode_utf16().take(32).enumerate() {
            name_bytes[i * 2..i * 2 + 2].copy_from_slice(&u.to_le_bytes());
        }
        f.extend_from_slice(&name_bytes);
        f.extend_from_slice(flags);
        // the instance is the font index
        collection.extend_from_slice(&ppt(0, index as u16, 0x0FB7, &f));
    }
    body.extend_from_slice(&ppt(0xF, 0, 0x07D5, &collection));
    body.extend_from_slice(&ppt(0, 0, 0x0FA4, &RULER_0FA4));
    body.extend_from_slice(&ppt(0, 0, 0x0FA5, &RULER_0FA5));
    body.extend_from_slice(&ppt(0, 0, 0x0FA9, &RULER_0FA9));
    body.extend_from_slice(&ppt(0, 4, 0x0FA3, style));
    ppt(0xF, 0, 0x03F2, &body)
}

#[cfg(test)]
mod environment_build_tests {
    use super::*;

    fn environment_of(path: &str) -> Option<Vec<u8>> {
        let bytes = std::fs::read(path).ok()?;
        let cfb = office_cfb::Cfb::parse(bytes).ok()?;
        let stream = cfb.stream("PowerPoint Document").ok()?;
        let mut at = 8usize;
        while at + 8 <= stream.len() {
            let kind = u16::from_le_bytes([stream[at + 2], stream[at + 3]]);
            let len = u32::from_le_bytes([stream[at + 4], stream[at + 5], stream[at + 6], stream[at + 7]]) as usize;
            if kind == 0x03F2 {
                return Some(stream[at + 8..at + 8 + len].to_vec());
            }
            at += 8 + len;
        }
        None
    }

    /// Everything up to the style atom is generated from field values.
    #[test]

    fn fonts_and_rulers_are_reproduced_from_their_fields() {
        let Some(reference) = environment_of("../../samples/void-ms.ppt") else {
            return;
        };
        // the kind field is the *second* u16 of the header
        let style_at = reference
            .windows(2)
            .position(|w| w == [0xa3, 0x0f])
            .expect("0x0FA3")
            - 2;
        let style_len = u32::from_le_bytes([
            reference[style_at + 4],
            reference[style_at + 5],
            reference[style_at + 6],
            reference[style_at + 7],
        ]) as usize;
        let style = &reference[style_at + 8..style_at + 8 + style_len];
        let fonts: [(&str, [u8; 4]); 3] = [
            ("\u{7b49}\u{7ebf}", [0x86, 0x00, 0x04, 0x02]),
            ("Arial", [0x00, 0x00, 0x04, 0x00]),
            ("\u{7b49}\u{7ebf} Light", [0x86, 0x00, 0x04, 0x02]),
        ];
        let ours = environment(&fonts, style);
        // `environment()` returns the whole record; the reference is its payload.
        assert_eq!(ours.len(), reference.len() + 8);
        let got = &ours[8..];
        if got != reference.as_slice() {
            let at = got.iter().zip(&reference).position(|(a, b)| a != b).unwrap();
            panic!(
                "first difference at {at}: ours={:02x?} reference={:02x?}",
                &got[at.saturating_sub(8)..(at + 8).min(got.len())],
                &reference[at.saturating_sub(8)..(at + 8).min(reference.len())]
            );
        }
    }
}

/// `SlideListWithText` (0x0FF0): a container of `SlidePersistAtom`s. Instance
/// 0 = slides, 1 = masters, 2 = notes (`docs/ppt.md` §4).
pub fn slide_list_with_text(instance: u16, atoms: &[Vec<u8>]) -> Vec<u8> {
    let mut body = Vec::new();
    for atom in atoms {
        body.extend_from_slice(atom);
    }
    ppt(0xF, instance, 0x0FF0, &body)
}

#[cfg(test)]
mod slide_list_tests {
    use super::*;

    #[test]
    fn both_lists_in_the_reference_deck_are_reproduced() {
        let Ok(bytes) = std::fs::read("../../samples/void-ms.ppt") else {
            return;
        };
        let cfb = office_cfb::Cfb::parse(bytes).expect("cfb");
        let stream = cfb.stream("PowerPoint Document").expect("stream");
        let d: &[u8] = &stream;
        let mut lists = Vec::new();
        let mut at = 8usize;
        while at + 8 <= d.len() {
            let vi = u16::from_le_bytes([d[at], d[at + 1]]);
            let kind = u16::from_le_bytes([d[at + 2], d[at + 3]]);
            let len = u32::from_le_bytes([d[at + 4], d[at + 5], d[at + 6], d[at + 7]]) as usize;
            if at + 8 + len > d.len() {
                break;
            }
            if kind == 0x0FF0 {
                lists.push((vi >> 4, d[at + 8..at + 8 + len].to_vec()));
            }
            at += 8 + len;
        }
        assert_eq!(lists.len(), 2);
        let master_atom = ppt(0, 0, 0x03F3, &slide_persist_atom_full(2, 0, 0x8000_0000));
        let masters = slide_list_with_text(1, &[master_atom]);
        assert_eq!(lists[0].0, 1);
        assert_eq!(masters[8..], lists[0].1[..], "master list");
        let slide_atom = ppt(0, 0, 0x03F3, &slide_persist_atom_full(3, 0, 0x0102));
        let slides = slide_list_with_text(0, &[slide_atom]);
        assert_eq!(lists[1].0, 0);
        assert_eq!(slides[8..], lists[1].1[..], "slide list");
    }
}

/// `PPDrawingGroup` (0x040B) of a deck with no pictures: `DggContainer` holding
/// the id-cluster table (`docs/ppt.md` §9.4.2), the drawing-group `FOPT` and
/// `0xF11E` (`docs/ppt.md` §9.10).
pub fn drawing_group(per_object: &[u32]) -> Vec<u8> {
    drawing_group_with(per_object, &[])
}

/// [`drawing_group`] plus one `OfficeArtFBSE` per embedded image.
///
/// A deck with a `Pictures` stream has to list each BLIP in the document's
/// drawing group: `(blip length, how many shapes point at it)`. Without the
/// store PowerPoint refuses the file (docs/ppt.md §9.24).
pub fn drawing_group_with(per_object: &[u32], blips: &[(u32, u32)]) -> Vec<u8> {
    const FOPT: [u8; 36] = [0x81, 0x01, 0x04, 0x00, 0x00, 0x08, 0x83, 0x01, 0x00, 0x00, 0x00, 0x08, 0xbf, 0x01, 0x10, 0x00, 0x10, 0x00, 0xc0, 0x01, 0x01, 0x00, 0x00, 0x08, 0xff, 0x01, 0x08, 0x00, 0x08, 0x00, 0x01, 0x02, 0x02, 0x00, 0x00, 0x08];
    const F11E: [u8; 16] = [0x04, 0x00, 0x00, 0x08, 0x01, 0x00, 0x00, 0x08, 0x02, 0x00, 0x00, 0x08, 0xf7, 0x00, 0x00, 0x10];
    let mut dgg = Vec::new();
    dgg.extend_from_slice(&ppt(0, 0, 0xF006, &fdgg_payload(per_object)));
    if !blips.is_empty() {
        let body: Vec<u8> = blips
            .iter()
            .flat_map(|(len, refs)| doc_fbse(*len, *refs))
            .collect();
        dgg.extend_from_slice(&ppt(0xF, 1, 0xF001, &body));
    }
    dgg.extend_from_slice(&ppt(3, 6, 0xF00B, &FOPT));
    dgg.extend_from_slice(&ppt(0, 4, 0xF11E, &F11E));
    ppt(0xF, 0, 0x040B, &ppt(0xF, 0, 0xF000, &dgg))
}

#[cfg(test)]
mod drawing_group_tests {
    use super::*;

    #[test]
    fn the_reference_drawing_group_is_reproduced() {
        let Ok(bytes) = std::fs::read("../../samples/void-ms.ppt") else {
            return;
        };
        let cfb = office_cfb::Cfb::parse(bytes).expect("cfb");
        let stream = cfb.stream("PowerPoint Document").expect("stream");
        let d: &[u8] = &stream;
        let mut at = 8usize;
        let mut reference = None;
        while at + 8 <= d.len() {
            let kind = u16::from_le_bytes([d[at + 2], d[at + 3]]);
            let len = u32::from_le_bytes([d[at + 4], d[at + 5], d[at + 6], d[at + 7]]) as usize;
            if kind == 0x040B {
                reference = Some(d[at + 8..at + 8 + len].to_vec());
                break;
            }
            at += 8 + len;
        }
        let reference = reference.expect("PPDrawingGroup");
        let ours = drawing_group(&[0, 7, 1]);
        assert_eq!(ours.len(), reference.len() + 8);
        // `FDGG`'s third word ("issued ids") is still undecoded (docs/ppt.md
        // §9.4.2), so it is compared separately from the rest.
        let mut got = ours[8..].to_vec();
        let mut want = reference.clone();

        got[24..28].fill(0);
        want[24..28].fill(0);
        assert_eq!(got, want, "drawing group differs outside the FDGG counter");
    }
}

/// `SlideViewInfo` (0x07D0) plus the two `0x0FD9` atoms that follow it in the
/// `Document` record (`docs/ppt.md` §9.11). They describe the editing view, not
/// the slides, so their content is constant for a deck with no outline saved.
pub fn slide_view_info() -> Vec<u8> {
    const VIEW: [u8; 296] = [0x1f, 0x00, 0x14, 0x04, 0x1c, 0x00, 0x00, 0x00, 0x00, 0x00, 0x15, 0x04, 0x14, 0x00, 0x00, 0x00, 0x21, 0xc9, 0xf4, 0x08, 0x00, 0xca, 0x9a, 0x3b, 0xad, 0x07, 0x94, 0xc7, 0x00, 0xca, 0x9a, 0x3b, 0x01, 0x02, 0x00, 0x00, 0x1f, 0x00, 0x13, 0x04, 0x3c, 0x00, 0x00, 0x00, 0x00, 0x00, 0xfd, 0x03, 0x34, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x0f, 0x00, 0xfa, 0x03, 0x47, 0x00, 0x00, 0x00, 0x00, 0x00, 0xfe, 0x03, 0x03, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0xfd, 0x03, 0x34, 0x00, 0x00, 0x00, 0x38, 0x00, 0x00, 0x00, 0x64, 0x00, 0x00, 0x00, 0x38, 0x00, 0x00, 0x00, 0x64, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x2b, 0x02, 0x00, 0x00, 0x2d, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x0f, 0x00, 0x88, 0x13, 0x69, 0x00, 0x00, 0x00, 0x0f, 0x00, 0x8a, 0x13, 0x29, 0x00, 0x00, 0x00, 0x00, 0x00, 0xba, 0x0f, 0x10, 0x00, 0x00, 0x00, 0x5f, 0x00, 0x5f, 0x00, 0x5f, 0x00, 0x50, 0x00, 0x50, 0x00, 0x54, 0x00, 0x31, 0x00, 0x32, 0x00, 0x00, 0x00, 0x8b, 0x13, 0x09, 0x00, 0x00, 0x00, 0x00, 0x00, 0x25, 0x04, 0x01, 0x00, 0x00, 0x00, 0x01, 0x0f, 0x00, 0x8a, 0x13, 0x30, 0x00, 0x00, 0x00, 0x00, 0x00, 0xba, 0x0f, 0x10, 0x00, 0x00, 0x00, 0x5f, 0x00, 0x5f, 0x00, 0x5f, 0x00, 0x50, 0x00, 0x50, 0x00, 0x54, 0x00, 0x31, 0x00, 0x30, 0x00, 0x00, 0x00, 0x8b, 0x13, 0x10, 0x00, 0x00, 0x00, 0x00, 0x00, 0x0d, 0x04, 0x08, 0x00, 0x00, 0x00, 0x70, 0xb5, 0x00, 0x00, 0x70, 0xb5, 0x00, 0x00];
    ppt(0xF, 0, 0x07D0, &VIEW)
}

pub fn slide_view_atoms() -> Vec<u8> {
    const A: [u8; 12] = [0x00, 0x00, 0xda, 0x0f, 0x04, 0x00, 0x00, 0x00, 0x0d, 0x00, 0x02, 0x00];
    const B: [u8; 12] = [0x00, 0x00, 0xda, 0x0f, 0x04, 0x00, 0x00, 0x00, 0x0d, 0x00, 0x02, 0x00];
    let mut v = ppt(0xF, 3, 0x0FD9, &A);
    v.extend_from_slice(&ppt(0xF, 4, 0x0FD9, &B));
    v
}

#[cfg(test)]
mod view_tests {
    use super::*;

    #[test]
    fn the_reference_view_records_are_reproduced() {
        let Ok(bytes) = std::fs::read("../../samples/void-ms.ppt") else {
            return;
        };
        let cfb = office_cfb::Cfb::parse(bytes).expect("cfb");
        let stream = cfb.stream("PowerPoint Document").expect("stream");
        let d: &[u8] = &stream;
        let mut at = 8usize;
        let (mut view, mut atoms) = (None, Vec::new());
        while at + 8 <= d.len() {
            let vi = u16::from_le_bytes([d[at], d[at + 1]]);
            let kind = u16::from_le_bytes([d[at + 2], d[at + 3]]);
            let len = u32::from_le_bytes([d[at + 4], d[at + 5], d[at + 6], d[at + 7]]) as usize;
            if at + 8 + len > d.len() {
                break;
            }
            if kind == 0x07D0 {
                view = Some(d[at + 8..at + 8 + len].to_vec());
            } else if kind == 0x0FD9 {
                atoms.push((vi >> 4, d[at + 8..at + 8 + len].to_vec()));
            }
            at += 8 + len;
        }
        let view = view.expect("SlideViewInfo");
        assert_eq!(&slide_view_info()[8..], view.as_slice());
        assert_eq!(atoms.len(), 2);
        let ours = slide_view_atoms();
        assert_eq!(ours.len(), atoms[0].1.len() + atoms[1].1.len() + 16);
        assert_eq!(&ours[8..8 + atoms[0].1.len()], atoms[0].1.as_slice());
        assert_eq!(&ours[16 + atoms[0].1.len()..], atoms[1].1.as_slice());
    }
}

/// The `serverZoom` a fresh PowerPoint document carries: 1:2, i.e. 50%.
///
/// The two `u32`s are a numerator/denominator pair (`RatioStruct`), so 5:10 —
/// what PowerPoint 2016 writes — is the same ratio. Only the ratio matters.
pub const DEFAULT_ZOOM: (u32, u32) = (1, 2);

/// The `slideSizeType` values of the `DocumentAtom` (`SlideSizeEnum`, §9.6.3):
/// `PpSlideSizeType - 1`, i.e. `onScreen = 0 … custom = 6`.
pub const SLIDE_SIZE_ON_SCREEN: u16 = 0;
pub const SLIDE_SIZE_CUSTOM: u16 = 6;

/// Builds the 40 byte `DocumentAtom` payload (MS-PPT §2.4.2, `docs/ppt.md` §9.6.3).
///
/// ```
/// # use office_ppt::write::*;
/// let atom = document_atom((7680, 4320), (4320, 5760), DEFAULT_ZOOM, 1, SLIDE_SIZE_CUSTOM, 0, 0);
/// assert_eq!(atom.len(), 40);
/// ```
///
/// `notes_master` / `handout_master` are persist ids; `0` means the document has
/// no notes (or handout) master. The four trailing flags are `fSaveWithFonts`,
/// `fOmitTitlePlace`, `fRightToLeft`, `fShowComments` — PowerPoint writes
/// `0, 0, 0, 1`, which is what this function does.
pub fn document_atom(
    slide_size: (u32, u32),
    notes_size: (u32, u32),
    zoom: (u32, u32),
    first_slide_number: u16,
    slide_size_type: u16,
    notes_master: u32,
    handout_master: u32,
) -> Vec<u8> {
    let mut out = Vec::with_capacity(40);
    for v in [slide_size.0, slide_size.1, notes_size.0, notes_size.1, zoom.0, zoom.1] {
        out.extend_from_slice(&v.to_le_bytes());
    }
    out.extend_from_slice(&notes_master.to_le_bytes());
    out.extend_from_slice(&handout_master.to_le_bytes());
    out.extend_from_slice(&first_slide_number.to_le_bytes());
    out.extend_from_slice(&slide_size_type.to_le_bytes());
    // fSaveWithFonts, fOmitTitlePlace, fRightToLeft, fShowComments
    out.extend_from_slice(&[0, 0, 0, 1]);
    out
}

/// Assembles the `Document` record (`docs/ppt.md` §9.5): ten children, ending
/// with `EndDocument` **inside** the record. `style` is the 110 byte default text
/// style, `theme` the theme ZIP, `document_atom` the 40 byte `DocumentAtom`.
pub fn document(per_object: &[u32], style: &[u8], theme: &[u8], document_atom: &[u8]) -> Vec<u8> {
    let master = ppt(0, 0, 0x03F3, &slide_persist_atom_full(2, 0, 0x8000_0000));
    let mut atoms = Vec::new();
    for i in 0..per_object.len().saturating_sub(2) {
        atoms.push(ppt(0, 0, 0x03F3, &slide_persist_atom_full(3 + i as u32, 0, 0x102 + i as u32)));
    }
    document_with(
        per_object,
        &[],
        DeckLists { masters: &[master], slides: &atoms, notes: &[] },
        style,
        theme,
        document_atom,
    )
}

/// The three `SlideListWithText` atom lists a `Document` carries: the masters
/// (instance 1), the slides (instance 0) and the notes pages (instance 2).
pub struct DeckLists<'a> {
    pub masters: &'a [Vec<u8>],
    pub slides: &'a [Vec<u8>],
    pub notes: &'a [Vec<u8>],
}

/// [`document`] with explicit atom lists.
///
/// A slide's `masterIdRef` has to name a *layout*, not the main master, so a
/// deck that wants visible shapes lists both (docs/ppt.md §9.14). The atoms are
/// `SlidePersistAtom` records, as built by [`slide_persist_atom_full`]; the notes
/// ones carry `cPersist = 4` (docs §9.25).
pub fn document_with(
    per_object: &[u32],
    blips: &[(u32, u32)],
    lists: DeckLists,
    style: &[u8],
    theme: &[u8],
    document_atom: &[u8],
) -> Vec<u8> {
    const FONTS: [(&str, [u8; 4]); 3] = [
        ("\u{7b49}\u{7ebf}", [0x86, 0x00, 0x04, 0x02]),
        ("Arial", [0x00, 0x00, 0x04, 0x00]),
        ("\u{7b49}\u{7ebf} Light", [0x86, 0x00, 0x04, 0x02]),
    ];
    let mut body = Vec::new();
    body.extend_from_slice(&ppt(1, 0, 0x03E9, document_atom));
    body.extend_from_slice(&environment(&FONTS, style));
    body.extend_from_slice(&drawing_group_with(per_object, blips));
    body.extend_from_slice(&slide_list_with_text(1, lists.masters));
    body.extend_from_slice(&slide_view_info());
    body.extend_from_slice(&slide_view_atoms());
    body.extend_from_slice(&slide_list_with_text(0, lists.slides));
    if !lists.notes.is_empty() {
        body.extend_from_slice(&slide_list_with_text(2, lists.notes));
    }
    body.extend_from_slice(&ppt(0, 0, 0x0428, theme));
    body.extend_from_slice(&ppt(0, 0, 0x03EA, &[]));
    ppt(0xF, 0, 0x03E8, &body)
}

#[cfg(test)]
mod generated_tests {
    use super::*;

    fn open(bytes: &[u8]) -> crate::Presentation {
        let cfb = office_cfb::Cfb::parse(bytes.to_vec()).expect("cfb");
        let doc = cfb.stream("PowerPoint Document").expect("stream");
        let user = cfb.stream("Current User").ok();
        crate::parse_stream(&doc, user.as_deref()).expect("parse")
    }

    fn stream(bytes: &[u8]) -> Vec<u8> {
        office_cfb::Cfb::parse(bytes.to_vec())
            .expect("cfb")
            .stream("PowerPoint Document")
            .expect("stream")
    }

    /// Every slide object, as `(cluster, spidMax, shape count)` from its FDG.
    /// Every drawing of the objects of `kind`, as `(cluster, spidMax, count)`.
    /// Every `SlideShowSlideInfoAtom` payload in the stream, in file order.
    fn transition_atoms(stream: &[u8]) -> Vec<[u8; 16]> {
        let mut out = Vec::new();
        for rec in crate::records(stream, 0, stream.len()) {
            if rec.kind == RT_SLIDE {
                walk(stream, 8 + rec.offset, 8 + rec.offset + rec.len, &mut out);
            }
        }
        out
    }

    fn walk(stream: &[u8], start: usize, end: usize, out: &mut Vec<[u8; 16]>) {
        let mut at = start;
        while at + 8 <= end {
            let Some(rec) = crate::read_record(stream, at, end) else { break };
            if rec.kind == RT_SLIDE_SHOW_INFO_ATOM && rec.len == 16 {
                let mut atom = [0u8; 16];
                atom.copy_from_slice(&stream[at + 8..at + 8 + 16]);
                out.push(atom);
            }
            if rec.is_container() {
                walk(stream, at + 8, at + 8 + rec.len, out);
            }
            at += 8 + rec.len;
        }
    }

    fn drawings_of(stream: &[u8], kind: u16) -> Vec<(u16, u32, u32)> {
        let mut out = Vec::new();
        for slide in crate::records(stream, 0, stream.len())
            .into_iter()
            .filter(|r| r.kind == kind)
        {
            let body = slide.offset + 8;
            let end = body + slide.len;
            for drawing in crate::records(stream, body, end)
                .into_iter()
                .filter(|r| r.kind == RT_PPDRAWING)
            {
                let at = drawing.offset + 8;
                for rec in crate::records(stream, at, at + drawing.len) {
                    if rec.kind != ESCHER_DG_CONTAINER {
                        continue;
                    }
                    let inner = rec.offset + 8;
                    for dg in crate::records(stream, inner, inner + rec.len) {
                        if dg.kind != ESCHER_DG_ATOM {
                            continue;
                        }
                        let p = dg.payload(stream);
                        let count = u32::from_le_bytes(p[0..4].try_into().unwrap());
                        let spid_max = u32::from_le_bytes(p[4..8].try_into().unwrap());
                        out.push((dg.instance, spid_max, count));
                    }
                }
            }
        }
        out
    }

    /// The deck parses back with each slide's own text — the shape tree is ours,
    /// not the template's.
    #[test]
    fn generated_decks_parse_back() {
        let deck = vec![
            SlideSpec::new("Generated one").with_body(&["first paragraph", "second paragraph"]),
            SlideSpec::new("Generated two").with_body(&["only paragraph"]),
        ];
        let bytes = write_presentation(&deck).expect("write");
        let pres = open(&bytes);
        assert_eq!(pres.slides.len(), 2);
        let texts: Vec<Vec<String>> = pres
            .slides
            .iter()
            .map(|s| s.texts.iter().map(|t| t.text.trim_end().to_string()).collect())
            .collect();
        // The body placeholder holds the paragraphs separated by marks, exactly
        // like the cloned path.
        assert_eq!(
            texts,
            vec![
                vec![
                    "Generated one".to_string(),
                    "first paragraph\rsecond paragraph".to_string()
                ],
                vec!["Generated two".to_string(), "only paragraph".to_string()],
            ]
        );
    }

    /// Each slide owns one drawing cluster and one shape id per shape, starting
    /// at the cluster's base: nothing is inherited from the template's slide.
    #[test]
    fn generated_slides_have_their_own_clusters() {
        let deck = vec![
            SlideSpec::new("one").with_body(&["a", "b"]),
            SlideSpec::new("two"),
        ];
        let bytes = write_presentation(&deck).expect("write");
        let stream = stream(&bytes);
        let drawings = drawings_of(&stream, RT_SLIDE);
        assert_eq!(drawings.len(), 2, "one drawing per slide");
        // Clusters 1..3 are the master, the layout and the notes master.
        // First slide: title + one body placeholder = two shapes, plus the group.
        assert_eq!(drawings[0], (4, 0x1002, 3));
        // Second slide: title only.
        assert_eq!(drawings[1], (5, 0x1401, 2));
        // The template's slide shape ids must not appear anywhere.
        assert!(
            !drawings.iter().any(|(_, spid_max, _)| *spid_max & 0x3FF == 3 && *spid_max >> 10 == 13),
            "cloned the template's slide: {drawings:?}"
        );
    }

    /// Notes pages round-trip: the deck carries one per slide that has notes,
    /// each in its own cluster right after its slide.
    #[test]
    fn transitions_land_in_the_slide_show_atom() {
        let slides = [
            SlideSpec::new("fade").with_transition(Transition::new(TransitionKind::Fade)),
            SlideSpec::new("dissolve")
                .with_transition(Transition::new(TransitionKind::Dissolve).after(2500)),
            SlideSpec::new("plain"),
        ];
        let bytes = write_presentation(&slides).unwrap();
        let stream = stream(&bytes);
        let atoms = transition_atoms(&stream);
        assert_eq!(atoms.len(), 2, "only two slides asked for a transition");
        assert_eq!(atoms[0], slide_show_info(Some(Transition::new(TransitionKind::Fade))));
        let dissolve = slide_show_info(Some(Transition::new(TransitionKind::Dissolve).after(2500)));
        assert_eq!(atoms[1], dissolve);
        assert_eq!(&atoms[1][0..4], &2500u32.to_le_bytes(), "slideTime in ms");
        assert_eq!(atoms[1][9], 5, "dissolve");
        assert_eq!(atoms[1][11], 0x04, "advance on its own");
    }

    #[test]
    fn notes_pages_round_trip() {
        let deck = vec![
            SlideSpec::new("one")
                .with_body(&["a"])
                .with_notes(&["note one", "note two"]),
            SlideSpec::new("two").with_body(&["b"]),
            SlideSpec::new("three").with_body(&["c"]).with_notes(&["note three"]),
        ];
        let bytes = write_presentation(&deck).expect("write");
        let pres = open(&bytes);
        assert_eq!(pres.slides.len(), 3);
        let notes: Vec<String> = pres
            .slides
            .iter()
            .map(|s| s.notes.iter().map(|t| t.text.trim_end()).collect::<Vec<_>>().join("|"))
            .collect();
        assert_eq!(notes, vec!["note one\rnote two".to_string(), String::new(), "note three".to_string()]);

        let stream = stream(&bytes);
        // Two slides carry notes, so there are two notes objects; the notes
        // master is one too, in cluster 3.
        let pages: Vec<_> = drawings_of(&stream, RT_NOTES)
            .into_iter()
            .filter(|(cluster, _, _)| *cluster != 3)
            .collect();
        assert_eq!(pages.len(), 2, "one drawing per notes page: {pages:?}");
        // Slide 1 is cluster 4, its notes 5, slide 2 cluster 6 (no notes),
        // slide 3 cluster 7 and its notes 8.
        assert_eq!(pages[0], (5, 0x1401, 2));
        assert_eq!(pages[1], (8, 0x2001, 2));
    }

    /// The notes list of the `Document` and the slide's `notesIdRef` have to
    /// name the same `slideId`; a mismatch makes PowerPoint refuse the file.
    #[test]
    fn notes_atoms_share_the_slide_id() {
        let deck = vec![SlideSpec::new("one")
            .with_body(&["a"])
            .with_notes(&["note"])];
        let bytes = write_presentation(&deck).expect("write");
        let stream = stream(&bytes);
        let document = crate::records(&stream, 0, stream.len())
            .into_iter()
            .find(|r| r.kind == RT_DOCUMENT)
            .expect("document");
        let mut notes_atom = None;
        let mut slide_atom = None;
        for rec in crate::records(&stream, document.offset + 8, document.offset + 8 + document.len) {
            if rec.kind == RT_SLIDE_LIST_WITH_TEXT && rec.instance == 2 {
                let inner = rec.offset + 8;
                let atom = crate::records(&stream, inner, inner + rec.len)
                    .into_iter()
                    .find(|r| r.kind == RT_SLIDE_PERSIST_ATOM)
                    .expect("notes atom");
                notes_atom = Some(atom.payload(&stream).to_vec());
            }
        }
        for rec in crate::records(&stream, 0, stream.len())
            .into_iter()
            .filter(|r| r.kind == RT_SLIDE)
        {
            for child in crate::records(&stream, rec.offset + 8, rec.offset + 8 + rec.len) {
                if child.kind == RT_SLIDE_ATOM {
                    slide_atom = Some(child.payload(&stream).to_vec());
                }
            }
        }
        let notes_atom = notes_atom.expect("notes list atom");
        let slide_atom = slide_atom.expect("slide atom");
        // SlidePersistAtom: persistIdRef, cPersist, flags, slideId, reserved.
        assert_eq!(u32::from_le_bytes(notes_atom[4..8].try_into().unwrap()), 4, "cPersist");
        let listed = u32::from_le_bytes(notes_atom[12..16].try_into().unwrap());
        // SlideAtom: geom, placeholder types, masterIdRef, notesIdRef, flags.
        let referenced = u32::from_le_bytes(slide_atom[16..20].try_into().unwrap());
        assert_eq!(listed, 0x100);
        assert_eq!(referenced, listed, "notesIdRef vs the notes list");
    }

    /// A deck with a picture needs the BLIP listed in the document's drawing
    /// group (`OfficeArtBStoreContainer`); without it PowerPoint rejects the
    /// file even though every record parses (docs/ppt.md §9.24).
    #[test]
    fn pictures_get_a_store_entry() {
        let Ok(png) = std::fs::read("../../samples/pixel.png") else { return };
        let deck = vec![SlideSpec::new("picture").with_picture(png, 60, 60)];
        let bytes = write_presentation(&deck).expect("write");
        let pres = open(&bytes);
        assert_eq!(pres.slides.len(), 1);
        assert_eq!(pres.slides[0].pictures, vec![1], "one shape points at pib 1");

        let stream = stream(&bytes);
        let doc = crate::records(&stream, 0, stream.len())
            .into_iter()
            .find(|r| r.kind == RT_DOCUMENT)
            .expect("document");
        let body = doc.offset + 8;
        let dgg = crate::records(&stream, body, body + doc.len)
            .into_iter()
            .find(|r| r.kind == 0x040B)
            .expect("drawing group");
        let inner = dgg.offset + 8;
        let store = crate::records(&stream, inner, inner + dgg.len)
            .into_iter()
            .find(|r| r.kind == 0xF000)
            .and_then(|r| {
                let at = r.offset + 8;
                crate::records(&stream, at, at + r.len)
                    .into_iter()
                    .find(|c| c.kind == 0xF001)
            })
            .expect("BStoreContainer");
        assert_eq!(store.instance, 1);
        assert_eq!(store.len, 44, "one OfficeArtFBSE is 44 bytes");
    }
}

#[cfg(test)]
mod document_tests {
    use super::*;

    fn children(payload: &[u8]) -> Vec<(u16, Vec<u8>)> {
        let mut out = Vec::new();
        let mut at = 0usize;
        while at + 8 <= payload.len() {
            let kind = u16::from_le_bytes([payload[at + 2], payload[at + 3]]);
            let len = u32::from_le_bytes([payload[at + 4], payload[at + 5], payload[at + 6], payload[at + 7]]) as usize;
            if at + 8 + len > payload.len() {
                break;
            }
            out.push((kind, payload[at + 8..at + 8 + len].to_vec()));
            at += 8 + len;
        }
        out
    }

    /// `document_atom` must cover all 40 bytes: rebuilding each sample's atom
    /// from its own fields has to give the file's bytes back exactly.
    #[test]
    fn document_atom_rebuilds_every_sample() {
        let mut checked = 0usize;
        for dir in ["../../samples", "../../samples/ms"] {
            let Ok(entries) = std::fs::read_dir(dir) else { continue };
            for entry in entries.flatten() {
                let path = entry.path();
                if path.extension().and_then(|e| e.to_str()) != Some("ppt") {
                    continue;
                }
                let Ok(bytes) = std::fs::read(&path) else { continue };
                let Ok(cfb) = office_cfb::Cfb::parse(bytes) else { continue };
                let Ok(stream) = cfb.stream("PowerPoint Document") else { continue };
                let Some(doc) = crate::records(&stream, 0, stream.len())
                    .into_iter()
                    .find(|r| r.kind == RT_DOCUMENT)
                else {
                    continue;
                };
                let Some(atom) = crate::records(&stream, doc.offset + 8, doc.offset + 8 + doc.len)
                    .into_iter()
                    .find(|r| r.kind == 0x03E9)
                else {
                    continue;
                };
                let p = atom.payload(&stream);
                assert_eq!(p.len(), 40, "{}", path.display());
                let w = |i: usize| u32::from_le_bytes([p[i], p[i + 1], p[i + 2], p[i + 3]]);
                let u = |i: usize| u16::from_le_bytes([p[i], p[i + 1]]);
                let ours = document_atom(
                    (w(0), w(4)),
                    (w(8), w(12)),
                    (w(16), w(20)),
                    u(32),
                    u(34),
                    w(24),
                    w(28),
                );
                assert_eq!(ours, p, "{}", path.display());
                checked += 1;
            }
        }
        assert!(checked > 20, "only {checked} samples checked");
    }

    #[test]
    fn the_reference_document_record_is_reassembled() {
        let Ok(bytes) = std::fs::read("../../samples/void-ms.ppt") else {
            return;
        };
        let cfb = office_cfb::Cfb::parse(bytes).expect("cfb");
        let stream = cfb.stream("PowerPoint Document").expect("stream");
        let d: &[u8] = &stream;
        let len = u32::from_le_bytes([d[4], d[5], d[6], d[7]]) as usize;
        let reference = &d[8..8 + len];
        let kids = children(reference);
        let pick = |kind: u16| kids.iter().find(|(k, _)| *k == kind).map(|(_, p)| p.clone());
        let document_atom = pick(0x03E9).expect("DocumentAtom");
        let theme = pick(0x0428).expect("theme");
        let env = pick(0x03F2).expect("Environment");
        let style = children(&env).into_iter().find(|(k, _)| *k == 0x0FA3).map(|(_, p)| p).expect("style");

        let ours = document(&[0, 7, 1], &style, &theme, &document_atom);
        let ours_kids: Vec<(u16, usize)> = children(&ours[8..]).iter().map(|(k, p)| (*k, p.len())).collect();
        let ref_kids: Vec<(u16, usize)> = kids.iter().map(|(k, p)| (*k, p.len())).collect();
        assert_eq!(ours_kids, ref_kids, "children");
        assert_eq!(ours.len(), reference.len() + 8, "length");
        // `FDGG`'s third word ("issued ids") is producer-dependent — Microsoft
        // writes `sum(used) - clusters`, WPS writes `sum(used)`, and
        // `void-ms.ppt` matches neither (docs/ppt.md §9.4.3). PowerPoint accepts
        // either, so mask it before comparing the rest.
        let mut got = ours[8..].to_vec();
        let mut want = reference.to_vec();
        for buf in [&mut got, &mut want] {
            buf[518..522].fill(0);
        }
        let got = got.as_slice();
        let want = want.as_slice();
        if let Some(at) = got.iter().zip(want).position(|(a, b)| a != b) {
            panic!("first difference at {at}: ours={:02x?} reference={:02x?}",
                &got[at.saturating_sub(6)..(at + 10).min(got.len())],
                &want[at.saturating_sub(6)..(at + 10).min(want.len())]);
        }
    }
}

#[cfg(test)]
mod document_atom_tests {
    use super::*;

    /// Field-by-field expectation for each measured sample (`docs/ppt.md` §9.6.3).
    ///
    /// `fsn.ppt` and `void-ms.ppt` are both 4:3, but one carries `5:10` zoom and
    /// size type 1, the other `1:2` zoom and type 6 — which is why the old
    /// "zoom fields follow the aspect ratio" theory was wrong: `5:10` and `1:2`
    /// are the same ratio, and the size type is independent.
    #[test]
    fn document_atom_matches_the_measured_samples() {
        #[allow(clippy::type_complexity)]
        let cases: [(&str, (u32, u32), (u32, u32), (u32, u32), u16, u16, u32, u32); 8] = [
            // file, slide size, notes size, zoom, first, type, notes master, handout master
            ("../../samples/void-ms.ppt", (7680, 4320), (4320, 5760), (1, 2), 1, 6, 0, 0),
            ("../../samples/ms/fsn.ppt", (5760, 4320), (4320, 5760), (5, 10), 5, 1, 0, 0),
            ("../../samples/ms/fsn2.ppt", (5760, 4320), (4320, 5760), (5, 10), 1, 1, 0, 0),
            ("../../samples/ms/size1.ppt", (5760, 4320), (4320, 5760), (5, 10), 7, 0, 0, 0),
            ("../../samples/ms/size3.ppt", (6240, 4320), (4320, 5760), (5, 10), 7, 2, 0, 0),
            ("../../samples/ms/size4.ppt", (6480, 4320), (4320, 5760), (5, 10), 7, 3, 0, 0),
            ("../../samples/ms/size6.ppt", (4608, 576), (4320, 5760), (5, 10), 7, 5, 0, 0),
            ("../../samples/ms/atom-notes.ppt", (7680, 4320), (4320, 5760), (5, 10), 1, 6, 2, 0),
        ];
        let mut checked = 0;
        for (path, slide, notes, zoom, first, kind, notes_master, handout) in cases {
            let Ok(bytes) = std::fs::read(path) else { continue };
            let cfb = office_cfb::Cfb::parse(bytes).expect("cfb");
            let stream = cfb.stream("PowerPoint Document").expect("stream");
            let reference = sample_atom(&stream);
            let ours = document_atom(slide, notes, zoom, first, kind, notes_master, handout);
            assert_eq!(ours, reference, "{path}");
            checked += 1;
        }
        assert!(checked >= 7, "only {checked} samples available");
    }

    /// Pulls the `DocumentAtom` payload out of a `PowerPoint Document` stream.
    fn sample_atom(stream: &[u8]) -> Vec<u8> {
        let doc = crate::records(stream, 0, stream.len())
            .into_iter()
            .find(|r| r.kind == RT_DOCUMENT)
            .expect("Document container");
        crate::records(stream, doc.offset + 8, doc.offset + 8 + doc.len)
            .into_iter()
            .find(|r| r.kind == 0x03E9)
            .expect("DocumentAtom")
            .payload(stream)
            .to_vec()
    }

    /// The notes master reference is a persist id, so it only appears in a deck
    /// that has one; `atom-notes.ppt` keeps it at id 2 (kind `0x03F0`).
    #[test]
    fn notes_master_reference_is_a_persist_id() {
        let Ok(bytes) = std::fs::read("../../samples/ms/atom-notes.ppt") else { return };
        let cfb = office_cfb::Cfb::parse(bytes).expect("cfb");
        let stream = cfb.stream("PowerPoint Document").expect("stream");
        let atom = sample_atom(&stream);
        let notes_master = u32::from_le_bytes([atom[24], atom[25], atom[26], atom[27]]);
        assert_eq!(notes_master, 2);
        let dir = persist_entries(&stream, &crate::records(&stream, 0, stream.len()))
            .expect("persist directory");
        let offset = dir
            .iter()
            .find(|(id, _)| *id == notes_master)
            .map(|(_, offset)| *offset)
            .expect("notes master object");
        let rec = crate::read_record(&stream, offset, stream.len()).expect("record");
        assert_eq!(rec.kind, 0x03F0, "NotesContainer");
    }
}

#[cfg(test)]
mod text_shape_tests {
    use super::*;

    fn kids(data: &[u8]) -> Vec<(u16, u16, Vec<u8>)> {
        crate::records(data, 0, data.len())
            .into_iter()
            .map(|r| (r.kind, r.instance, r.payload(data).to_vec()))
            .collect()
    }

    fn find(kids: &[(u16, u16, Vec<u8>)], kind: u16) -> &(u16, u16, Vec<u8>) {
        kids.iter().find(|k| k.0 == kind).expect("child record")
    }

    /// A from-scratch shape is only accepted when every count agrees with every
    /// other one (`docs/ppt.md` §9.21); the four that matter are the `FOPT`
    /// length, the two style `cch` fields, and the special-info count.
    #[test]
    fn text_shape_counts_agree() {
        let text = "Generated shape";
        let cch = text.chars().count() as u32;
        let shape = text_shape(0xC04, TEXT_TYPE_BODY, text, [0x02F3, 0x1B10, 0x0893], 24, false);
        // `text_shape` returns the `SpContainer` record itself, so its children
        // are one level in.
        let sp = kids(&shape);
        let inner = kids(&find(&sp, ESCHER_SP_CONTAINER).2);

        let spid = u32::from_le_bytes(inner[0].2[0..4].try_into().unwrap());
        assert_eq!(spid, 0xC04);

        let fopt = find(&inner, ESCHER_FOPT);
        assert_eq!(fopt.1, 15, "one instance per named property");
        assert_eq!(fopt.2.len(), 102, "fifteen properties plus two slots of padding");

        let data = find(&inner, ESCHER_CLIENT_DATA);
        let cd = kids(&data.2);
        assert_eq!(find(&cd, 0x0BC3).2.len(), 8);
        assert_eq!(find(&cd, 0x041F).2, 1u32.to_le_bytes());

        let textbox = find(&inner, ESCHER_CLIENT_TEXTBOX);
        let tb = kids(&textbox.2);
        assert_eq!(find(&tb, RT_TEXT_HEADER_ATOM).2, TEXT_TYPE_BODY.to_le_bytes());
        // ASCII text goes into the byte atom, one byte per character.
        assert_eq!(find(&tb, RT_TEXT_BYTES_ATOM).2, text.as_bytes());
        // Paragraph run: cch, indent level, masks, alignment; character run:
        // cch, masks, font size. Both counts include the paragraph mark.
        let style = &find(&tb, RT_STYLE_TEXT_PROP_ATOM).2;
        assert_eq!(u32::from_le_bytes(style[0..4].try_into().unwrap()), cch + 1);
        assert_eq!(u32::from_le_bytes(style[12..16].try_into().unwrap()), cch + 1);
        let special = &find(&tb, RT_TEXT_SPECIAL_INFO).2;
        assert_eq!(u32::from_le_bytes(special[0..4].try_into().unwrap()), cch);
        assert_eq!(special.len(), 28);
        assert_eq!(find(&tb, 0x0FA6).2.len(), 12);
    }

    /// Non-ASCII text has to use the UTF-16 atom, which counts two bytes per
    /// character instead of one.
    #[test]
    fn text_shape_uses_the_wide_atom_for_non_ascii() {
        let shape = text_shape(0xC05, TEXT_TYPE_BODY, "中文", [0; 3], 18, true);
        let sp = kids(&shape);
        let inner = kids(&find(&sp, ESCHER_SP_CONTAINER).2);
        let tb = kids(&find(&inner, ESCHER_CLIENT_TEXTBOX).2);
        assert_eq!(find(&tb, RT_TEXT_CHARS_ATOM).2.len(), 4);
        assert!(tb.iter().all(|k| k.0 != RT_TEXT_BYTES_ATOM));
    }
}

#[cfg(test)]
mod round_trip_store_tests {
    use super::*;

    /// The writer drops Office 2007's round-trip XML (`docs/ppt.md` §9.22).
    /// PowerPoint opens those decks and rebuilds the blobs on save; keeping them
    /// would mean shipping ~30 KB of the template's ZIP bytes in every file.
    #[test]
    fn generated_decks_carry_no_round_trip_store() {
        let Ok(tpl) = std::fs::read("../../samples/ms/minimal.ppt") else { return };
        let specs = vec![SlideSpec::new("Stripped")];
        let out = write_from_slide_template(&tpl, &specs).unwrap();
        let cfb = office_cfb::Cfb::parse(out).unwrap();
        let stream = cfb.stream("PowerPoint Document").unwrap();
        let template = office_cfb::Cfb::parse(std::fs::read("../../samples/ms/minimal.ppt").unwrap())
            .unwrap()
            .stream("PowerPoint Document")
            .unwrap();
        assert!(
            stream.len() + 20_000 < template.len(),
            "template {} B, output {} B",
            template.len(),
            stream.len()
        );
        fn walk(data: &[u8], start: usize, end: usize) -> Vec<u16> {
            let mut out = Vec::new();
            for r in crate::records(data, start, end) {
                if ROUND_TRIP_STORE.contains(&r.kind) {
                    out.push(r.kind);
                }
                if r.is_container() {
                    out.extend(walk(data, r.offset + 8, r.offset + 8 + r.len));
                }
            }
            out
        }
        let found = walk(&stream, 0, stream.len());
        assert!(found.is_empty(), "round-trip records left behind: {found:04X?}");
        // The template still has them, so the check is not vacuous.
        assert!(!walk(&template, 0, template.len()).is_empty());
    }

    /// `strip_round_trip` leaves atoms alone and fixes the container length.
    #[test]
    fn strip_round_trip_drops_only_the_store() {
        let keep = ppt(0, 0, RT_SLIDE_ATOM, &[1, 2, 3, 4]);
        let store = ppt(0, 0, 0x040E, &[0u8; 16]);
        let object = container(RT_SLIDE, &[keep.clone(), store]);
        let stripped = strip_round_trip(&object);
        assert_eq!(stripped.len(), object.len() - 24);
        let kids = crate::records(&stripped, 8, stripped.len());
        assert_eq!(kids.len(), 1);
        assert_eq!(kids[0].kind, RT_SLIDE_ATOM);
        assert_eq!(
            u32::from_le_bytes(stripped[4..8].try_into().unwrap()) as usize,
            stripped.len() - 8
        );
        // An atom (not a container) passes through unchanged.
        assert_eq!(strip_round_trip(&keep), keep);
    }
}
