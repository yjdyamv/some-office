//! Rebuild the template embedded in `office-ppt` from `samples/ms/minimal.ppt`.
//!
//! `write_presentation` only uses three of the template's persist objects — the
//! `Document`, the main master and the first layout — plus the trailing
//! `UserEditAtom`/`PersistDirectoryAtom` pair. The other eleven layouts and the
//! sample slide are dead weight (150 KB of the 203 KB blob), so this tool writes
//! a stream that keeps just the three objects and re-numbers the directory.
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

    let mut out = Vec::new();
    let mut offsets = Vec::new();
    for (_, start, end) in &keep {
        offsets.push(out.len() as u32);
        out.extend_from_slice(&stream[*start..*end]);
    }

    // One `PersistDirectoryAtom` covering ids 1..=3, then the `UserEditAtom`.
    let mut payload = u32b(1 | (3 << 20));
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
    user_edit.extend_from_slice(&u32b(4)); // persistIdSeed
    user_edit.extend_from_slice(&u32b(1)); // lastView
    user_edit.extend_from_slice(&u32b(0));
    out.extend_from_slice(&record(0, 0, office_ppt::RT_USER_EDIT_ATOM, &user_edit));

    let path = "crates/office-ppt/templates/minimal.bin";
    std::fs::write(path, &out).expect("write");
    println!("{path} {} bytes (was {})", out.len(), stream.len());
}
