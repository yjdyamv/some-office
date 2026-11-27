//! Dump a .ppt's record tree with persist ids and offsets.
use office_ppt::{read_record, records};

fn name(k: u16) -> &'static str {
    match k {
        0x03E8 => "Document",
        0x03E9 => "DocumentAtom",
        0x03EA => "EndDocument",
        0x03EE => "Slide",
        0x03EF => "SlideAtom",
        0x03F0 => "Notes",
        0x03F1 => "NotesAtom",
        0x03F2 => "Environment",
        0x03F3 => "SlidePersistAtom",
        0x03F8 => "MainMaster",
        0x03FA => "SlideViewInfo",
        0x03FF => "SlideViewInfoInstance",
        0x040C => "PPDrawingGroup",
        0x040D => "PPDrawing",
        0x0411 => "DrawingGroup",
        0x0428 => "Theme",
        0x0FA3 => "TextMasterStyleAtom",
        0x0FF5 => "UserEditAtom",
        0x1772 => "PersistDirectoryAtom",
        0x0FF0 => "CurrentUserAtom",
        0x0FBA => "SlideListWithText",
        0x0FBB => "SlidePersistAtom?",
        0x0F9F => "SlideListWithText2",
        0x0FA0 => "SlidePersistAtom2",
        0x0FA8 => "SlideListWithText?",
        0x0FA9 => "SlidePersistAtom?",
        0x0FB4 => "NotesListWithText",
        0x0FB5 => "NotesPersistAtom",
        0x0FC9 => "Handout",
        0x0FCA => "HandoutAtom",
        0x1388 => "OfficeArtDgg",
        0x0BDB => "OfficeArtDg",
        0x0BDC => "OfficeArtSpgr",
        0x0BDD => "OfficeArtSp",
        0xF009 => "FSPGR",
        0xF00A => "FSP",
        0xF00B => "FOPT",
        0xF010 => "ClientAnchor",
        0x0BE1 => "ClientTextbox",
        0x0BE7 => "ClientAnchor",
        0x0BE8 => "ClientData",
        0x0BDA => "OfficeArtBStore",
        0x0BD0 => "ClientRecord?",
        _ => "",
    }
}

fn dump(data: &[u8], start: usize, end: usize, depth: usize) {
    let mut off = start;
    while let Some(r) = read_record(data, off, end) {
        let nm = name(r.kind);
        println!(
            "{:indent$}@{:<7} {:04X} v{} i{} len={}{}",
            "",
            r.offset,
            r.kind,
            r.ver,
            r.instance,
            r.len,
            if nm.is_empty() { String::new() } else { format!("  {}", nm) },
            indent = depth * 2
        );
        if r.len <= 64 && std::env::var("HEX").is_ok() {
            let p = &data[r.offset + 8..r.offset + 8 + r.len];
            println!("{:indent$}  {}", "", p.iter().map(|b| format!("{:02x}", b)).collect::<Vec<_>>().join(" "), indent = depth * 2);
        }
        if r.is_container() && depth < 8 {
            dump(data, r.offset + 8, r.offset + 8 + r.len, depth + 1);
        }
        off += 8 + r.len;
    }
}

fn main() {
    let path = std::env::args().nth(1).expect("path");
    let only = std::env::args().nth(2);
    let data = std::fs::read(&path).expect("read");
    let cfb = office_cfb::Cfb::parse(data).expect("cfb");
    let stream = cfb.stream("PowerPoint Document").expect("stream").to_vec();
    let top = records(&stream, 0, stream.len());
    let edit = top.iter().rev().find(|r| r.kind == 0x0FF5).expect("edit");
    let p = edit.offset + 8;
    let dir_at = u32::from_le_bytes(stream[p + 12..p + 16].try_into().unwrap()) as usize;
    let dir = read_record(&stream, dir_at, stream.len()).expect("dir");
    let body = dir.offset + 8;
    let end = body + dir.len;
    let mut pos = body;
    let mut out = Vec::new();
    while pos + 4 <= end {
        let word = u32::from_le_bytes(stream[pos..pos + 4].try_into().unwrap());
        let first = word & 0xF_FFFF;
        let mut count = ((word >> 20) & 0xFFF) as usize;
        if count == 0 {
            count = 1;
        }
        pos += 4;
        for i in 0..count {
            let off = u32::from_le_bytes(stream[pos..pos + 4].try_into().unwrap()) as usize;
            pos += 4;
            out.push((first + i as u32, off));
        }
    }
    out.sort_by_key(|(_, o)| *o);
    for (id, off) in out {
        if let Some(only) = &only {
            if only != &id.to_string() {
                continue;
            }
        }
        let r = read_record(&stream, off, stream.len()).unwrap();
        println!("=== id {} @{} {:04X} {}", id, off, r.kind, name(r.kind));
        dump(&stream, off, off + 8 + r.len, 1);
    }
    if only.is_none() {
        println!("=== trailing top-level from {} to {}", dir_at, stream.len());
        dump(&stream, dir_at, stream.len(), 0);
    }
}
