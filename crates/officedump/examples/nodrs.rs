//! Drops the round-trip ZIP blobs (`drs/*.xml`, theme, table styles) from a
//! deck and rewrites the persist directory, to see whether PowerPoint needs
//! them at all (docs §9.22).
//!
//! `cargo run -q --example nodrs samples/scratch5-minclone.ppt out.ppt`

use office_ppt::{read_record, records, Record};

const RT_USER_EDIT_ATOM: u16 = 0x0FF5;
const RT_PERSIST_DIRECTORY_ATOM: u16 = 0x1772;

/// Records that only carry XML for the round-trip store; PowerPoint regenerates
/// them on save, so a deck written by this library does not need them.
const DRS: [u16; 5] = [0x040E, 0x0423, 0x041E, 0x0428, 0xC3A9];

fn persist_entries(stream: &[u8], top: &[Record]) -> Vec<(u32, usize)> {
    let edit = top.iter().rev().find(|r| r.kind == RT_USER_EDIT_ATOM).expect("UserEditAtom");
    let payload = edit.offset + 8;
    let dir_at =
        u32::from_le_bytes(stream[payload + 12..payload + 16].try_into().unwrap()) as usize;
    let dir = read_record(stream, dir_at, stream.len()).expect("persist directory");
    assert_eq!(dir.kind, RT_PERSIST_DIRECTORY_ATOM);
    let body = dir.offset + 8;
    let end = body + dir.len;
    let mut out = Vec::new();
    let mut pos = body;
    while pos + 4 <= end {
        let word = u32::from_le_bytes(stream[pos..pos + 4].try_into().unwrap());
        let first = word & 0xF_FFFF;
        let mut count = ((word >> 20) & 0xFFF) as usize;
        if count == 0 {
            count = 1;
        }
        pos += 4;
        for i in 0..count {
            if pos + 4 > end {
                break;
            }
            let offset = u32::from_le_bytes(stream[pos..pos + 4].try_into().unwrap()) as usize;
            out.push((first + i as u32, offset));
            pos += 4;
        }
    }
    out.sort_by_key(|(_, offset)| *offset);
    out
}

/// Rebuilds one persist object with its `drs` children removed.
fn strip(object: &[u8]) -> Vec<u8> {
    if object.len() < 8 {
        return object.to_vec();
    }
    let kind = u16::from_le_bytes([object[2], object[3]]);
    let vi = u16::from_le_bytes([object[0], object[1]]);
    if vi & 0xF != 0xF {
        return object.to_vec(); // an atom, not a container
    }
    let keep_all = std::env::var("KEEP_DRS").is_ok();
    let mut kids = Vec::new();
    for r in records(object, 8, object.len()) {
        if !keep_all && DRS.contains(&r.kind) {
            continue;
        }
        kids.extend_from_slice(&object[r.offset..r.offset + 8 + r.len]);
    }
    let mut out = object[0..8].to_vec();
    out[4..8].copy_from_slice(&(kids.len() as u32).to_le_bytes());
    out.extend_from_slice(&kids);
    assert_eq!(kind, u16::from_le_bytes([out[2], out[3]]));
    out
}

fn main() {
    let args: Vec<String> = std::env::args().skip(1).collect();
    let (src, dst) = (args[0].clone(), args[1].clone());
    let bytes = std::fs::read(&src).expect("read");
    let cfb = office_cfb::Cfb::parse(bytes).expect("cfb");
    let stream = cfb.stream("PowerPoint Document").expect("stream");
    let top = records(&stream, 0, stream.len());
    let entries = persist_entries(&stream, &top);
    let edit = top.iter().rev().find(|r| r.kind == RT_USER_EDIT_ATOM).expect("edit");

    // Re-lay the objects out in their old order, dropping the drs records.
    let mut out = Vec::new();
    let mut moved = Vec::new();
    for (id, offset) in &entries {
        let rec = read_record(&stream, *offset, stream.len()).expect("object");
        let object = &stream[rec.offset..rec.offset + 8 + rec.len];
        let stripped = strip(object);
        moved.push((*id, out.len()));
        out.extend_from_slice(&stripped);
    }

    // Fresh persist directory, in the form PowerPoint writes: one entry whose
    // word holds the first id and the count, then one offset per object.
    let mut ids: Vec<(u32, usize)> = moved.clone();
    ids.sort_by_key(|(id, _)| *id);
    assert!(
        ids.iter().enumerate().all(|(i, (id, _))| *id == ids[0].0 + i as u32),
        "persist ids must be contiguous for the one-entry form"
    );
    let mut dir = Vec::new();
    dir.extend_from_slice(&(ids[0].0 | ((ids.len() as u32) << 20)).to_le_bytes());
    for (_, at) in &ids {
        dir.extend_from_slice(&(*at as u32).to_le_bytes());
    }
    let dir_at = out.len();
    out.extend_from_slice(&[0x00, 0x00, 0x72, 0x17]);
    out.extend_from_slice(&(dir.len() as u32).to_le_bytes());
    out.extend_from_slice(&dir);

    let mut ue = stream[edit.offset + 8..edit.offset + 8 + edit.len].to_vec();
    ue[8..12].copy_from_slice(&0u32.to_le_bytes()); // offsetLastEdit
    ue[12..16].copy_from_slice(&(dir_at as u32).to_le_bytes()); // offsetPersistDirectory
    let ue_at = out.len();
    out.extend_from_slice(&[0x00, 0x00, 0xF5, 0x0F]);
    out.extend_from_slice(&(ue.len() as u32).to_le_bytes());
    out.extend_from_slice(&ue);

    // Current User: an `0x0FF6` record whose payload holds the offset of the
    // last edit (after the size and header token).
    let mut user = cfb.stream("Current User").expect("current user");
    assert_eq!(u16::from_le_bytes([user[2], user[3]]), 0x0FF6, "CurrentUserAtom");
    user[16..20].copy_from_slice(&(ue_at as u32).to_le_bytes());

    let mut streams = vec![
        office_cfb::OutputStream::new("PowerPoint Document", &out),
        office_cfb::OutputStream::new("Current User", &user),
    ];
    let pictures = cfb.stream("Pictures").ok();
    if let Some(pictures) = &pictures {
        streams.push(office_cfb::OutputStream::new("Pictures", pictures));
    }
    std::fs::write(&dst, office_cfb::write(&streams)).expect("write");
    println!(
        "{dst}: {} bytes (was {}), {} objects, {} bytes dropped",
        out.len(),
        stream.len(),
        entries.len(),
        stream.len().saturating_sub(out.len())
    );
}
