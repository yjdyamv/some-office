//! Rebuild the template embedded in `office-ppt` from `samples/ms/minimal.ppt`
//! and `samples/ms/notes2.ppt`.
//!
//! `write_presentation` only uses four of the template's persist objects — the
//! `Document`, the main master, the first layout and the notes master — plus the
//! trailing `UserEditAtom`/`PersistDirectoryAtom` pair. The other eleven layouts
//! and the sample slides are dead weight (170 KB of the 203 KB blob), so this
//! tool writes a stream that keeps just the four objects and re-numbers the
//! directory.
//!
//! The notes master comes from `samples/ms/notes2.ppt` (written by PowerPoint
//! 2016 with real speaker notes, so it carries a notes master); its drawing sits
//! in cluster 14 there and is moved to cluster 3 here.
//!
//! ```text
//! cargo run -q -p officedump --example trimtemplate
//! ```

fn u32b(value: u32) -> Vec<u8> {
    value.to_le_bytes().to_vec()
}

fn record(ver: u16, instance: u16, kind: u16, payload: &[u8]) -> Vec<u8> {
    let mut out = Vec::with_capacity(8 + payload.len());
    out.extend_from_slice(&(ver | (instance << 4)).to_le_bytes());
    out.extend_from_slice(&kind.to_le_bytes());
    out.extend_from_slice(&(payload.len() as u32).to_le_bytes());
    out.extend_from_slice(payload);
    out
}

/// `(persist id, start, end)` for every object, read from the directory.
fn persist_entries(stream: &[u8]) -> Vec<(u32, usize, usize)> {
    let mut entries = Vec::new();
    for rec in office_ppt::records(stream, 0, stream.len()) {
        if rec.kind != office_ppt::RT_USER_EDIT_ATOM {
            continue;
        }
        let payload = rec.payload(stream);
        let dir = u32::from_le_bytes(payload[12..16].try_into().unwrap()) as usize;
        let Some(dir_rec) = office_ppt::read_record(stream, dir, stream.len()) else { continue };
        let body = dir_rec.offset + 8;
        let mut pos = body;
        while pos + 4 <= body + dir_rec.len {
            let word = u32::from_le_bytes(stream[pos..pos + 4].try_into().unwrap());
            let first = word & 0x000F_FFFF;
            let count = word >> 20;
            pos += 4;
            for i in 0..count {
                let at = pos + (i as usize) * 4;
                let start = u32::from_le_bytes(stream[at..at + 4].try_into().unwrap()) as usize;
                entries.push((first + i, start, 0));
            }
            pos += count as usize * 4;
        }
    }
    entries.sort_by_key(|e| e.1);
    for entry in &mut entries {
        let start = entry.1;
        entry.2 = office_ppt::read_record(stream, start, stream.len())
            .map_or(stream.len(), |rec| start + 8 + rec.len);
    }
    entries
}

/// Moves every shape id in a drawing from one cluster to another: the `FSP`
/// records and the `FDG`'s instance and `spidMax`.
fn renumber(data: &mut [u8], start: usize, end: usize, old: u16, new: u16) {
    let mut at = start;
    while at + 8 <= end {
        let word = u16::from_le_bytes([data[at], data[at + 1]]);
        let kind = u16::from_le_bytes([data[at + 2], data[at + 3]]);
        let len = u32::from_le_bytes([data[at + 4], data[at + 5], data[at + 6], data[at + 7]]) as usize;
        let body = at + 8;
        let stop = body + len;
        if stop > end {
            break;
        }
        match kind {
            0xF00A if len >= 4 => {
                let spid = u32::from_le_bytes([data[body], data[body + 1], data[body + 2], data[body + 3]]);
                if spid >> 10 == old as u32 {
                    let moved = ((new as u32) << 10) | (spid & 0x03FF);
                    data[body..body + 4].copy_from_slice(&moved.to_le_bytes());
                }
            }
            0xF008 if len >= 8 && word >> 4 == old => {
                data[at..at + 2].copy_from_slice(&((word & 0xF) | (new << 4)).to_le_bytes());
                let max = u32::from_le_bytes([
                    data[body + 4],
                    data[body + 5],
                    data[body + 6],
                    data[body + 7],
                ]);
                let moved = ((new as u32) << 10) | (max & 0x03FF);
                data[body + 4..body + 8].copy_from_slice(&moved.to_le_bytes());
            }
            _ => {}
        }
        if kind == 0xF000 || kind == 0xF001 || kind == 0xF002 || kind == 0xF003 || kind == 0xF004
            || kind == 0xF00D || kind == 0xF006 || kind == 0xF007 || kind == 0x040C
            || kind == 0x03F0
        {
            renumber(data, body, stop, old, new);
        }
        at = stop;
    }
}

fn main() {
    let Ok(bytes) = std::fs::read("samples/ms/minimal.ppt") else {
        eprintln!("samples/ms/minimal.ppt is missing");
        return;
    };
    let cfb = office_cfb::Cfb::parse(bytes).expect("cfb");
    let stream = cfb.stream("PowerPoint Document").expect("stream");

    let entries = persist_entries(&stream);
    let keep: Vec<(u32, usize, usize)> = entries
        .iter()
        .filter(|(id, _, _)| *id <= 3)
        .copied()
        .collect();
    assert_eq!(keep.len(), 3, "expected Document, master and one layout");

    // The notes master, moved from the notes deck's cluster 14 to cluster 3 so
    // that the deck's clusters stay contiguous.
    let notes_bytes = std::fs::read("samples/ms/notes2.ppt").expect("samples/ms/notes2.ppt");
    let notes_cfb = office_cfb::Cfb::parse(notes_bytes).expect("cfb");
    let notes_stream = notes_cfb.stream("PowerPoint Document").expect("stream");
    let notes_dir = persist_entries(&notes_stream);
    let (_, start, end) = *notes_dir
        .iter()
        .find(|(id, _, _)| *id == 2)
        .expect("notes master");
    let mut notes_master = notes_stream[start..end].to_vec();
    let notes_len = notes_master.len();
    renumber(&mut notes_master, 0, notes_len, 14, 3);

    let mut out = Vec::new();
    let mut offsets = Vec::new();
    for (_, start, end) in &keep {
        offsets.push(out.len() as u32);
        out.extend_from_slice(&stream[*start..*end]);
    }
    offsets.push(out.len() as u32);
    out.extend_from_slice(&notes_master);

    // One `PersistDirectoryAtom` covering ids 1..=4, then the `UserEditAtom`.
    let mut payload = u32b(1 | (4 << 20));
    for offset in &offsets {
        payload.extend_from_slice(&u32b(*offset));
    }
    let directory = record(0, 0, office_ppt::RT_PERSIST_DIRECTORY_ATOM, &payload);
    let directory_at = out.len() as u32;
    out.extend_from_slice(&directory);

    let mut user_edit = Vec::new();
    user_edit.extend_from_slice(&u32b(0)); // lastSlideIdRef
    user_edit.extend_from_slice(&0x460Cu16.to_le_bytes());
    user_edit.extend_from_slice(&[0, 3]); // minor, major
    user_edit.extend_from_slice(&u32b(0)); // offsetLastEdit
    user_edit.extend_from_slice(&u32b(directory_at));
    user_edit.extend_from_slice(&u32b(1)); // docPersistIdRef
    user_edit.extend_from_slice(&u32b(5)); // persistIdSeed
    user_edit.extend_from_slice(&u32b(1)); // lastView
    user_edit.extend_from_slice(&u32b(0));
    out.extend_from_slice(&record(0, 0, office_ppt::RT_USER_EDIT_ATOM, &user_edit));

    let path = "crates/office-ppt/templates/minimal.bin";
    std::fs::write(path, &out).expect("write");
    println!("{path} {} bytes (was {})", out.len(), stream.len());
}
