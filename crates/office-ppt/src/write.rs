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
}

impl SlideSpec {
    pub fn new(title: impl Into<String>) -> Self {
        SlideSpec {
            title: title.into(),
            body: Vec::new(),
        }
    }

    pub fn with_body(mut self, lines: &[&str]) -> Self {
        self.body = lines.iter().map(|s| s.to_string()).collect();
        self
    }
}

// PowerPoint record types (MS-PPT 2.13).
const RT_DOCUMENT: u16 = 0x03E8;
const RT_DOCUMENT_ATOM: u16 = 0x03E9;
const RT_END_DOCUMENT: u16 = 0x03EA;
const RT_SLIDE: u16 = 0x03EE;
const RT_SLIDE_ATOM: u16 = 0x03EF;
const RT_MAIN_MASTER: u16 = 0x03F8;
const RT_SLIDE_PERSIST_ATOM: u16 = 0x03F3;
const RT_PPDRAWING: u16 = 0x040C;
const RT_SLIDE_SHOW_SLIDE_INFO: u16 = 0x07F0;
const RT_TEXT_HEADER_ATOM: u16 = 0x0F9F;
const RT_TEXT_CHARS_ATOM: u16 = 0x0FA0;
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
pub fn write_presentation(slides: &[SlideSpec]) -> Vec<u8> {
    // Persist ids: 1 = Document, 2 = MainMaster, 3.. = slides.
    let master_id: u32 = 2;
    let first_slide_id: u32 = 3;

    // ---- Document -----------------------------------------------------------
    // DocumentAtom: slide size 10x7.5 inch in EMU (914400 per inch), 4:3.
    let mut doc_atom = Vec::new();
    doc_atom.extend_from_slice(&u32b(9144000)); // slideSizeX
    doc_atom.extend_from_slice(&u32b(6858000)); // slideSizeY
    doc_atom.extend_from_slice(&u32b(6858000)); // notesSizeX
    doc_atom.extend_from_slice(&u32b(9144000)); // notesSizeY
    doc_atom.extend_from_slice(&0u16.to_le_bytes()); // serverZoom
    doc_atom.extend_from_slice(&u32b(0)); // notesMasterPersistIdRef
    doc_atom.extend_from_slice(&u32b(0)); // handoutMasterPersistIdRef
    doc_atom.extend_from_slice(&1u16.to_le_bytes()); // firstSlideNum
    doc_atom.extend_from_slice(&0u16.to_le_bytes()); // slideSizeType (4:3)
    doc_atom.extend_from_slice(&0u16.to_le_bytes()); // fSaveWithFonts etc.
    doc_atom.resize(40, 0);

    let mut master_list = Vec::new();
    master_list.extend(slide_persist_atom(master_id, 0x8000_0000));

    let mut slide_list = Vec::new();
    for i in 0..slides.len() {
        slide_list.extend(slide_persist_atom(first_slide_id + i as u32, 0x100 + i as u32));
    }

    let document = container(
        RT_DOCUMENT,
        &[
            ppt(0, 0, RT_DOCUMENT_ATOM, &doc_atom),
            ppt(0xF, 1, RT_SLIDE_LIST_WITH_TEXT, &master_list), // instance 1: masters
            ppt(0xF, 0, RT_SLIDE_LIST_WITH_TEXT, &slide_list),  // instance 0: slides
            ppt(0, 0, RT_END_DOCUMENT, &[]),
        ],
    );

    // ---- MainMaster ---------------------------------------------------------
    let master = container(
        RT_MAIN_MASTER,
        &[
            slide_atom(0, 0),
            container(
                RT_PPDRAWING,
                &[container(
                    ESCHER_DG_CONTAINER,
                    &[escher(0, ESCHER_DG_ATOM, &[0u8; 8]), container(ESCHER_SPGR_CONTAINER, &[])],
                )],
            ),
        ],
    );

    // ---- Slides -------------------------------------------------------------
    let mut slide_records: Vec<Vec<u8>> = Vec::new();
    for (i, spec) in slides.iter().enumerate() {
        let mut shapes: Vec<Vec<u8>> = Vec::new();
        // A group shape wraps the placeholders, exactly like a real slide.
        shapes.push(container(
            ESCHER_SP_CONTAINER,
            &[
                escher(0, ESCHER_SPGR, &[0u8; 16]),
                escher(
                    0,
                    ESCHER_SP,
                    &[0x00, 0x08, 0x00, 0x00, 0x05, 0x00, 0x00, 0x00],
                ),
            ],
        ));
        let title = format!("{}\r", spec.title);
        shapes.push(text_shape(
            0x0801,
            TEXT_TYPE_TITLE,
            &title,
            [0x02F3, 0x1B10, 0x0893],
            60,
            true,
        ));
        if !spec.body.is_empty() {
            let body = spec.body.join("\r") + "\r";
            shapes.push(text_shape(
                0x0802 + i as u32,
                TEXT_TYPE_BODY,
                &body,
                [0x08C3, 0x1B10, 0x0C62],
                24,
                false,
            ));
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
        slide_records.push(slide);
    }

    // ---- Persist directory + user edit --------------------------------------
    let mut stream = document.clone();
    let master_at = stream.len() as u32;
    stream.extend(&master);
    let mut slide_offsets: Vec<u32> = Vec::new();
    for slide in &slide_records {
        slide_offsets.push(stream.len() as u32);
        stream.extend(slide);
    }

    // One PersistDirectoryAtom entry: persist id 1 with `count` offsets.
    let count = 2 + slides.len() as u32;
    let mut dir = Vec::new();
    dir.extend_from_slice(&u32b(1 | (count << 20)));
    dir.extend_from_slice(&u32b(0)); // Document
    dir.extend_from_slice(&u32b(master_at)); // MainMaster
    for offset in &slide_offsets {
        dir.extend_from_slice(&u32b(*offset));
    }
    let dir_at = stream.len() as u32;
    stream.extend(ppt(0, 0, RT_PERSIST_DIRECTORY_ATOM, &dir));

    // UserEditAtom: last edit points back at itself.
    let mut edit = Vec::new();
    edit.extend_from_slice(&u32b(0x100 + slides.len().saturating_sub(1) as u32)); // lastSlideIdRef
    edit.extend_from_slice(&0x17DFu16.to_le_bytes()); // version
    edit.push(0); // minor
    edit.push(3); // major
    edit.extend_from_slice(&u32b(0)); // offsetLastEdit
    edit.extend_from_slice(&u32b(dir_at)); // offsetPersistDirectory
    edit.extend_from_slice(&u32b(1)); // docPersistIdRef
    edit.extend_from_slice(&u32b(count + 1)); // persistIdSeed
    edit.extend_from_slice(&1u16.to_le_bytes()); // lastView
    edit.extend_from_slice(&0u16.to_le_bytes()); // unused (record is 28 bytes)
    let edit_at = stream.len() as u32;
    stream.extend(ppt(0, 0, RT_USER_EDIT_ATOM, &edit));

    // ---- Current User -------------------------------------------------------
    let user = current_user(edit_at, "Rust");

    office_cfb::write(&[
        office_cfb::OutputStream::new("PowerPoint Document", &stream),
        office_cfb::OutputStream::new("Current User", &user),
    ])
}

/// Build a `.ppt` using an existing deck as the boilerplate.
///
/// PowerPoint validates far more than the slide list: the document needs an
/// `Environment` (font collection + colour scheme), a `PPDrawingGroup` (the
/// drawing space) and a theme, and those are hundreds of bytes of values with
/// no documented defaults — a deck without them is rejected with "Office
/// detected a problem with the contents of this file" even though every record
/// is well formed. Rather than invent those numbers, this function keeps the
/// template's structural records (everything in `Document` except the two
/// slide lists) and its `MainMaster`, and generates the slides themselves.
///
/// `template` is any deck with a single master — a blank presentation works.
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
    let top = crate::records(&stream, 0, stream.len());
    let slide = *top
        .iter()
        .find(|r| r.kind == RT_SLIDE)
        .ok_or(crate::Error::NoDocument)?;
    let end = slide.offset + 8 + slide.len;

    let mut atom: Vec<u8> = Vec::new();
    let mut trailing: Vec<u8> = Vec::new();
    let mut pos = slide.offset + 8;
    while pos + 8 <= end {
        let Some(rec) = crate::read_record(&stream, pos, end) else { break };
        if rec.kind == RT_SLIDE_ATOM {
            atom = stream[pos..pos + 8 + rec.len].to_vec();
        } else if rec.kind != RT_PPDRAWING {
            trailing.extend_from_slice(&stream[pos..pos + 8 + rec.len]);
        }
        pos += 8 + rec.len;
        if rec.len == 0 {
            break;
        }
    }
    if atom.is_empty() {
        atom = slide_atom(2, 0);
    }

    let mut built: Vec<Vec<u8>> = Vec::new();
    let mut spid = 0x400u32;
    for spec in slides {
        let mut used = 0usize;
        let drawing = clone_with_text(&stream, slide.offset + 8, end, spec, &mut used, &mut spid);
        let mut rec = container(RT_SLIDE, &[atom.clone(), drawing]);
        rec.extend_from_slice(&trailing);
        built.push(rec);
    }

    let document = top
        .iter()
        .find(|r| r.kind == RT_DOCUMENT)
        .ok_or(crate::Error::NoDocument)?;
    let master = top
        .iter()
        .find(|r| r.kind == RT_MAIN_MASTER)
        .ok_or(crate::Error::NoDocument)?;
    let mut boilerplate: Vec<Vec<u8>> = Vec::new();
    let mut masters_at = None;
    let mut slides_at = None;
    let mut pos = document.offset + 8;
    let d_end = document.offset + 8 + document.len;
    while pos + 8 <= d_end {
        let Some(rec) = crate::read_record(&stream, pos, d_end) else { break };
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
    let mut master_list = Vec::new();
    master_list.extend(slide_persist_atom(2, 0x8000_0000));
    let mut slide_list = Vec::new();
    for i in 0..built.len() {
        slide_list.extend(slide_persist_atom(3 + i as u32, 0x100 + i as u32));
    }
    let at_masters = masters_at.unwrap_or(boilerplate.len()).min(boilerplate.len());
    boilerplate.insert(at_masters, ppt(0xF, 1, RT_SLIDE_LIST_WITH_TEXT, &master_list));
    let mut at_slides = slides_at.unwrap_or(boilerplate.len());
    if at_masters <= at_slides {
        at_slides += 1;
    }
    boilerplate.insert(
        at_slides.min(boilerplate.len()),
        ppt(0xF, 0, RT_SLIDE_LIST_WITH_TEXT, &slide_list),
    );
    boilerplate.push(ppt(0, 0, RT_END_DOCUMENT, &[]));

    let mut out = container(RT_DOCUMENT, &boilerplate);
    let master_at = out.len() as u32;
    out.extend_from_slice(&stream[master.offset..master.offset + 8 + master.len]);
    let mut offsets = Vec::new();
    for rec in &built {
        offsets.push(out.len() as u32);
        out.extend_from_slice(rec);
    }
    let count = 2 + built.len() as u32;
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
    edit.extend_from_slice(&u32b(0x100 + built.len().saturating_sub(1) as u32));
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

/// Copy a record range, replacing the text inside every `ClientTextbox`.
fn clone_with_text(
    data: &[u8],
    start: usize,
    end: usize,
    spec: &SlideSpec,
    used: &mut usize,
    spid: &mut u32,
) -> Vec<u8> {
    let mut out = Vec::new();
    let mut pos = start;
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
            // Shape ids must be unique across the whole document.
            *spid += 1;
            let mut sp = data[body..stop].to_vec();
            sp[0..4].copy_from_slice(&spid.to_le_bytes());
            out.extend(ppt(rec.ver as u16, rec.instance, rec.kind, &sp));
        } else if rec.kind == RT_TEXT_CHARS_ATOM {
            let text = pick_text(spec, *used);
            *used += 1;
            let utf16: Vec<u8> = text.chars().flat_map(|c| (c as u16).to_le_bytes()).collect();
            out.extend(ppt(rec.ver as u16, rec.instance, rec.kind, &utf16));
        } else if rec.kind == RT_STYLE_TEXT_PROP_ATOM {
            let text = pick_text(spec, used.saturating_sub(1));
            let cch = text.chars().count() as u32;
            let title = *used <= 1;
            let mut style = Vec::new();
            style.extend_from_slice(&u32b(cch));
            style.extend_from_slice(&0u16.to_le_bytes());
            style.extend_from_slice(&u32b(0x0000_0800));
            style.extend_from_slice(&(if title { 1u16 } else { 0 }).to_le_bytes());
            style.extend_from_slice(&u32b(cch));
            style.extend_from_slice(&u32b(0x0002_0000));
            style.extend_from_slice(&(if title { 60u16 } else { 24 }).to_le_bytes());
            out.extend(ppt(rec.ver as u16, rec.instance, rec.kind, &style));
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
    if index == 0 {
        format!("{}\r", spec.title)
    } else if spec.body.is_empty() {
        "\r".to_string()
    } else {
        format!("{}\r", spec.body.join("\r"))
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
    payload.extend_from_slice(&u32b(0)); // cPersist
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
        let bytes = write_presentation(&deck);
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
        let bytes = write_presentation(&[SlideSpec::new("T")]);
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
}
