//! Dump the persist directory of a .ppt: (persistId, offset, record kind, len).
use office_ppt::{read_record, records};

fn main() {
    let path = std::env::args().nth(1).expect("path");
    let data = std::fs::read(&path).expect("read");
    let cfb = office_cfb::Cfb::parse(data).expect("cfb");
    let stream = cfb.stream("PowerPoint Document").expect("stream").to_vec();
    let top = records(&stream, 0, stream.len());
    let edit = top
        .iter()
        .rev()
        .find(|r| r.kind == 0x0FF5)
        .expect("UserEditAtom");
    let p = edit.offset + 8;
    let dir_at = u32::from_le_bytes(stream[p + 12..p + 16].try_into().unwrap()) as usize;
    let dir = read_record(&stream, dir_at, stream.len()).expect("dir");
    println!("dir @{} kind={:04X} len={}", dir_at, dir.kind, dir.len);
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
        let r = read_record(&stream, off, stream.len());
        match r {
            Some(r) => println!(
                "id={:>2} off={:>6} kind={:04X} len={:>6} end={}",
                id,
                off,
                r.kind,
                r.len,
                off + 8 + r.len
            ),
            None => println!("id={:>2} off={:>6} <unparsable>", id, off),
        }
    }
}
