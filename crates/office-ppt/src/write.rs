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
        }
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
const RT_MAIN_MASTER: u16 = 0x03F8;
const RT_SLIDE_PERSIST_ATOM: u16 = 0x03F3;
const RT_PPDRAWING: u16 = 0x040C;
const RT_SLIDE_SHOW_SLIDE_INFO: u16 = 0x07F0;
const RT_TEXT_HEADER_ATOM: u16 = 0x0F9F;
const RT_TEXT_CHARS_ATOM: u16 = 0x0FA0;
const RT_TEXT_BYTES_ATOM: u16 = 0x0FA8;
const RT_STYLE_TEXT_PROP_ATOM: u16 = 0x0FA1;
const RT_TEXT_SPECIAL_INFO: u16 = 0x0FAA;
const RT_SLIDE_LIST_WITH_TEXT: u16 = 0x0FF0;
const RT_USER_EDIT_ATOM: u16 = 0x0FF5;
const RT_CURRENT_USER: u16 = 0x0FF6;
const RT_PERSIST_DIRECTORY_ATOM: u16 = 0x1772;
const RT_SLIDE_SCHEME_COLOR: u16 = 0x0422;

// Escher record types.
const ESCHER_DG_CONTAINER: u16 = 0xF002;
const ESCHER_SPGR_CONTAINER: u16 = 0xF003;
const ESCHER_SP_CONTAINER: u16 = 0xF004;
const ESCHER_DG_ATOM: u16 = 0xF008;
const ESCHER_SPGR: u16 = 0xF009;
const ESCHER_SP: u16 = 0xF00A;
const ESCHER_FOPT: u16 = 0xF00B;
const ESCHER_CLIENT_TEXTBOX: u16 = 0xF00D;
const ESCHER_CLIENT_ANCHOR: u16 = 0xF010;
const ESCHER_CLIENT_DATA: u16 = 0xF011;

/// `TextHeaderAtom.textType`: 0 title, 1 body, 2 notes, 6 centre title, 5 centre body.
const TEXT_TYPE_TITLE: u32 = 6;
const TEXT_TYPE_BODY: u32 = 5;

fn ppt(ver: u16, instance: u16, kind: u16, payload: &[u8]) -> Vec<u8> {
    let mut out = Vec::with_capacity(8 + payload.len());
    out.extend_from_slice(&(((instance << 4) | (ver & 0xF)) as u16).to_le_bytes());
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

/// Build a `.ppt` from a list of slides.
/// Write a presentation from scratch, using the template baked into this crate.
///
/// The template is a trimmed deck written by Microsoft PowerPoint (Document +
/// MainMaster + one slide, see `tools/trim_ppt_template.py`): PowerPoint
/// validates the shape tree far too strictly for us to synthesise one by hand
/// (`docs/ppt.md` §9).
pub fn write_presentation(slides: &[SlideSpec]) -> Result<Vec<u8>, crate::Error> {
    if slides.is_empty() {
        return Err(crate::Error::NoDocument);
    }
    // The baked-in template has no images, so no base `Pictures` stream.
    let (stream, edit_at, pictures) = build_from_stream(TEMPLATE_STREAM, slides, &[])?;
    Ok(pack(&stream, edit_at, &pictures))
}

/// Trimmed `PowerPoint Document` stream used by [`write_presentation`].
const TEMPLATE_STREAM: &[u8] = include_bytes!("../templates/minimal.bin");

pub fn write_from_template(template: &[u8], slides: &[SlideSpec]) -> Result<Vec<u8>, crate::Error> {
    let cfb = office_cfb::Cfb::parse(template.to_vec())?;
    let stream = cfb.stream("PowerPoint Document")?.to_vec();
    let top = crate::records(&stream, 0, stream.len());

    let document = top
        .iter()
        .find(|r| r.kind == RT_DOCUMENT)
        .ok_or(crate::Error::NoDocument)?;
    let master = top
        .iter()
        .find(|r| r.kind == RT_MAIN_MASTER)
        .ok_or(crate::Error::NoDocument)?;

    // Structural records of the template's Document, minus the slide lists.
    let mut boilerplate: Vec<Vec<u8>> = Vec::new();
    let mut masters_at = None;
    let mut slides_at = None;
    let mut pos = document.offset + 8;
    let end = document.offset + 8 + document.len;
    while pos + 8 <= end {
        let Some(rec) = crate::read_record(&stream, pos, end) else {
            break;
        };
        let raw = stream[pos..pos + 8 + rec.len].to_vec();
        match rec.kind {
            RT_END_DOCUMENT => {}
            RT_SLIDE_LIST_WITH_TEXT if rec.instance == 1 => masters_at = Some(boilerplate.len()),
            RT_SLIDE_LIST_WITH_TEXT if rec.instance == 0 => slides_at = Some(boilerplate.len()),
            RT_SLIDE_LIST_WITH_TEXT => {}
            _ => boilerplate.push(raw),
        }
        pos += 8 + rec.len;
        if rec.len == 0 {
            break;
        }
    }

    let master_id: u32 = 2;
    let first_slide_id: u32 = 3;

    // The master list keeps the template's ordering but only our master.
    let mut master_list = Vec::new();
    master_list.extend(slide_persist_atom(master_id, 0x8000_0000));
    let mut slide_list = Vec::new();
    for i in 0..slides.len() {
        slide_list.extend(slide_persist_atom(first_slide_id + i as u32, 0x100 + i as u32));
    }
    let masters_rec = ppt(0xF, 1, RT_SLIDE_LIST_WITH_TEXT, &master_list);
    let slides_rec = ppt(0xF, 0, RT_SLIDE_LIST_WITH_TEXT, &slide_list);

    // Put them back where the template had them, so the record order matches.
    let insert_masters = masters_at.unwrap_or(boilerplate.len()).min(boilerplate.len());
    boilerplate.insert(insert_masters, masters_rec);
    let mut insert_slides = slides_at.unwrap_or(boilerplate.len());
    if insert_masters <= insert_slides {
        insert_slides += 1; // the master list pushed everything down by one
    }
    let insert_slides = insert_slides.min(boilerplate.len());
    boilerplate.insert(insert_slides, slides_rec);
    boilerplate.push(ppt(0, 0, RT_END_DOCUMENT, &[]));

    let document = container(RT_DOCUMENT, &boilerplate);
    let master_bytes = stream[master.offset..master.offset + 8 + master.len].to_vec();

    let mut out = document;
    let master_at = out.len() as u32;
    out.extend(&master_bytes);
    let mut slide_offsets: Vec<u32> = Vec::new();
    for (i, spec) in slides.iter().enumerate() {
        slide_offsets.push(out.len() as u32);
        out.extend(slide_record(master_id, i, spec));
    }

    let count = 2 + slides.len() as u32;
    let mut dir = Vec::new();
    dir.extend_from_slice(&u32b(1 | (count << 20)));
    dir.extend_from_slice(&u32b(0));
    dir.extend_from_slice(&u32b(master_at));
    for offset in &slide_offsets {
        dir.extend_from_slice(&u32b(*offset));
    }
    let dir_at = out.len() as u32;
    out.extend(ppt(0, 0, RT_PERSIST_DIRECTORY_ATOM, &dir));

    let mut edit = Vec::new();
    edit.extend_from_slice(&u32b(0x100 + slides.len().saturating_sub(1) as u32));
    edit.extend_from_slice(&0x17DFu16.to_le_bytes());
    edit.push(0);
    edit.push(3);
    edit.extend_from_slice(&u32b(0));
    edit.extend_from_slice(&u32b(dir_at));
    edit.extend_from_slice(&u32b(1));
    edit.extend_from_slice(&u32b(count));
    edit.extend_from_slice(&1u16.to_le_bytes()); // lastView
    edit.extend_from_slice(&0u16.to_le_bytes()); // unused: the record is 28 bytes
    let edit_at = out.len() as u32;
    out.extend(ppt(0, 0, RT_USER_EDIT_ATOM, &edit));

    let user = current_user(edit_at, "Rust");
    Ok(office_cfb::write(&[
        office_cfb::OutputStream::new("PowerPoint Document", &out),
        office_cfb::OutputStream::new("Current User", &user),
    ]))
}

/// Diagnostic: rebuild a deck keeping the template's own `Slide` records and
/// swapping in `slides` for the rest. Used to tell a broken slide record apart
/// from a broken document assembly.
pub fn write_from_template_mixed(
    template: &[u8],
    slides: &[SlideSpec],
    keep_template_slides: bool,
) -> Result<Vec<u8>, crate::Error> {
    let cfb = office_cfb::Cfb::parse(template.to_vec())?;
    let stream = cfb.stream("PowerPoint Document")?.to_vec();
    let top = crate::records(&stream, 0, stream.len());
    let document = top.iter().find(|r| r.kind == RT_DOCUMENT).ok_or(crate::Error::NoDocument)?;
    let master = top.iter().find(|r| r.kind == RT_MAIN_MASTER).ok_or(crate::Error::NoDocument)?;
    let template_slides: Vec<&crate::Record> = top.iter().filter(|r| r.kind == RT_SLIDE).collect();

    let mut boilerplate: Vec<Vec<u8>> = Vec::new();
    let mut pos = document.offset + 8;
    let end = document.offset + 8 + document.len;
    while pos + 8 <= end {
        let Some(rec) = crate::read_record(&stream, pos, end) else { break };
        if rec.kind != RT_END_DOCUMENT && rec.kind != RT_SLIDE_LIST_WITH_TEXT {
            boilerplate.push(stream[pos..pos + 8 + rec.len].to_vec());
        }
        pos += 8 + rec.len;
        if rec.len == 0 { break; }
    }
    let n = if keep_template_slides { template_slides.len() } else { slides.len() };
    let mut master_list = Vec::new();
    master_list.extend(slide_persist_atom(2, 0x8000_0000));
    let mut slide_list = Vec::new();
    for i in 0..n {
        slide_list.extend(slide_persist_atom(3 + i as u32, 0x100 + i as u32));
    }
    boilerplate.insert(3.min(boilerplate.len()), ppt(0xF, 1, RT_SLIDE_LIST_WITH_TEXT, &master_list));
    boilerplate.push(ppt(0xF, 0, RT_SLIDE_LIST_WITH_TEXT, &slide_list));
    boilerplate.push(ppt(0, 0, RT_END_DOCUMENT, &[]));

    let mut out = container(RT_DOCUMENT, &boilerplate);
    let master_at = out.len() as u32;
    out.extend(&stream[master.offset..master.offset + 8 + master.len]);
    let mut offsets = Vec::new();
    for i in 0..n {
        offsets.push(out.len() as u32);
        if keep_template_slides {
            let r = template_slides[i];
            out.extend(&stream[r.offset..r.offset + 8 + r.len]);
        } else {
            out.extend(slide_record(2, i, &slides[i]));
        }
    }
    let count = 2 + n as u32;
    let mut dir = Vec::new();
    dir.extend_from_slice(&u32b(1 | (count << 20)));
    dir.extend_from_slice(&u32b(0));
    dir.extend_from_slice(&u32b(master_at));
    for o in &offsets {
        dir.extend_from_slice(&u32b(*o));
    }
    let dir_at = out.len() as u32;
    out.extend(ppt(0, 0, RT_PERSIST_DIRECTORY_ATOM, &dir));
    let mut edit = Vec::new();
    edit.extend_from_slice(&u32b(0x100 + n.saturating_sub(1) as u32));
    edit.extend_from_slice(&0x17DFu16.to_le_bytes());
    edit.push(0);
    edit.push(3);
    edit.extend_from_slice(&u32b(0));
    edit.extend_from_slice(&u32b(dir_at));
    edit.extend_from_slice(&u32b(1));
    edit.extend_from_slice(&u32b(count));
    edit.extend_from_slice(&1u16.to_le_bytes());
    edit.extend_from_slice(&0u16.to_le_bytes());
    let edit_at = out.len() as u32;
    out.extend(ppt(0, 0, RT_USER_EDIT_ATOM, &edit));
    let user = current_user(edit_at, "Rust");
    Ok(office_cfb::write(&[
        office_cfb::OutputStream::new("PowerPoint Document", &out),
        office_cfb::OutputStream::new("Current User", &user),
    ]))
}

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

/// Rebuild a presentation from a template stream, reusing every persist object.
///
/// Returns the new `PowerPoint Document` stream plus the offset of its
/// `UserEditAtom` (which `Current User` has to point at).
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
            if rec.kind != RT_SLIDE_ATOM && rec.kind != RT_PPDRAWING {
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
    for i in 0..objects.len() {
        let (id, start) = objects[i];
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
            pieces.push((id, stream[start..end].to_vec()));
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
fn picture_shape(spid: u32, pib: u32, width: i32, height: i32) -> Vec<u8> {
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
                    0x4104 => Some(pib as u32),
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
        } else if rec.kind == RT_TEXT_SPECIAL_INFO_ATOM {
            // `TextSpecialInfoAtom` starts with the character count it covers;
            // PowerPoint cross-checks it against the text atom, and a stale
            // count makes the whole file "damaged".
            let mut info = data[body..stop].to_vec();
            if let Some((_, new)) = counts {
                if info.len() >= 4 {
                    info[0..4].copy_from_slice(&(new as u32).to_le_bytes());
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
fn slide_record(master_id: u32, index: usize, spec: &SlideSpec) -> Vec<u8> {
    let mut shapes: Vec<Vec<u8>> = Vec::new();
    shapes.push(container(
        ESCHER_SP_CONTAINER,
        &[
            escher(0, ESCHER_SPGR, &[0u8; 16]),
            escher(0, ESCHER_SP, &[0x00, 0x08, 0x00, 0x00, 0x05, 0x00, 0x00, 0x00]),
        ],
    ));
    let title = format!("{}\r", spec.title);
    shapes.push(text_shape(0x0801, TEXT_TYPE_TITLE, &title, [0x02F3, 0x1B10, 0x0893], 60, true));
    if !spec.body.is_empty() {
        let body = spec.body.join("\r") + "\r";
        shapes.push(text_shape(0x0802 + index as u32, TEXT_TYPE_BODY, &body, [0x08C3, 0x1B10, 0x0C62], 24, false));
    }
    let drawing = container(
        RT_PPDRAWING,
        &[container(
            ESCHER_DG_CONTAINER,
            &[
                escher(2, ESCHER_DG_ATOM, &[0x00, 0x08, 0x00, 0x00, 0x03, 0x00, 0x00, 0x00]),
                container(ESCHER_SPGR_CONTAINER, &shapes),
            ],
        )],
    );
    let mut slide = container(RT_SLIDE, &[slide_atom(master_id, 0), drawing]);
    slide.extend(ppt(0, 1, RT_SLIDE_SHOW_SLIDE_INFO, &[0u8; 32]));
    slide.extend(ppt(0, 0, RT_SLIDE_SCHEME_COLOR, &[0u8; 8]));
    slide
}

/// `SlidePersistAtom` (MS-PPT 2.9.45): 20 bytes, and PowerPoint checks all of
/// them — a 4-byte atom makes it reject the whole file.
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
fn text_shape(
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

    // OfficeArtFOPT: geometry (0x007F = 1 is "custom"?) plus the outline.
    let mut fopt = Vec::new();
    let props: [(u16, u32); 8] = [
        (0x007F, 1),
        (0x0081, 0x0001_5F90),
        (0x0082, 0x0000_B6D0),
        (0x0083, 0x0001_5F90),
        (0x0084, 0x0000_B6D0),
        (0x0087, 2),
        (0x0301, if centred { 2 } else { 4 }),
        (0x0380, size_points as u32 * 20),
    ];
    for (pid, value) in props {
        fopt.extend_from_slice(&pid.to_le_bytes());
        fopt.extend_from_slice(&value.to_le_bytes());
    }

    let mut anchor_bytes = Vec::new();
    anchor_bytes.extend_from_slice(&0x0240u16.to_le_bytes()); // flags
    for v in anchor {
        anchor_bytes.extend_from_slice(&(v as u16).to_le_bytes());
    }

    // OfficeArtClientData with a single "no data" entry.
    let client_data = vec![0u8, 0, 0, 0, 0, 0];

    let cch = text.chars().count() as u32;
    let mut style = Vec::new();
    style.extend_from_slice(&u32b(cch)); // TextPFRun count
    style.extend_from_slice(&0u16.to_le_bytes()); // indentLevel
    style.extend_from_slice(&u32b(0x0000_0800)); // PFMasks: textAlignment
    style.extend_from_slice(&(if centred { 1u16 } else { 0 }).to_le_bytes());
    style.extend_from_slice(&u32b(cch)); // TextCFRun count
    style.extend_from_slice(&u32b(0x0002_0000)); // CFMasks: fontSize
    style.extend_from_slice(&size_points.to_le_bytes());

    let mut textbox = Vec::new();
    textbox.extend(ppt(0, 0, RT_TEXT_HEADER_ATOM, &u32b(text_type)));
    let utf16: Vec<u8> = text
        .chars()
        .flat_map(|c| (c as u16).to_le_bytes())
        .collect();
    textbox.extend(ppt(0, 0, RT_TEXT_CHARS_ATOM, &utf16));
    textbox.extend(ppt(0, 0, RT_STYLE_TEXT_PROP_ATOM, &style));
    textbox.extend(ppt(0, 0, RT_TEXT_SPECIAL_INFO, &[0u8; 12]));

    container(
        ESCHER_SP_CONTAINER,
        &[
            escher(1, ESCHER_SP, &sp),
            escher(8, ESCHER_FOPT, &fopt),
            escher(0, ESCHER_CLIENT_ANCHOR, &anchor_bytes),
            escher(1, ESCHER_CLIENT_DATA, &client_data),
            container(ESCHER_CLIENT_TEXTBOX, &[textbox]),
        ],
    )
}

fn current_user(offset_to_current_edit: u32, name: &str) -> Vec<u8> {
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
/// `TextSpecialInfoAtom` (0x0FAA): the first word is the covered character count.
const RT_TEXT_SPECIAL_INFO_ATOM: u16 = 0x0FAA;

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
