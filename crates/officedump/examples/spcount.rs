//! Count `OfficeArtSp` (0xF00A) records per persist object: `spcount FILE`.
//!
//! Used to work out the third word of `OfficeArtFDGGBlock` (`docs/ppt.md` §9.4.2).
use office_ppt as p;

fn fdgg(data: &[u8], from: usize, to: usize) -> Option<Vec<u8>> {
    let mut pos = from;
    while pos + 8 <= to {
        let rec = p::read_record(data, pos, to)?;
        let body = pos + 8;
        let stop = (body + rec.len).min(to);
        if rec.kind == 0xF006 {
            return Some(data[body..stop].to_vec());
        }
        if matches!(rec.kind, 0x03E9 | 0x03EE | 0x03F8 | 0xF000 | 0xF002 | 0xF003 | 0xF004 | 0x040B | 0x040C) {
            if let Some(v) = fdgg(data, body, stop) {
                return Some(v);
            }
        }
        pos = stop;
    }
    None
}

fn count(data: &[u8], from: usize, to: usize) -> usize {
    let mut n = 0;
    let mut pos = from;
    while pos + 8 <= to {
        let Some(rec) = p::read_record(data, pos, to) else { break };
        let body = pos + 8;
        let stop = (body + rec.len).min(to);
        if rec.kind == 0xF00A {
            n += 1;
        } else if rec.kind == 0xF009 {
            n += 1000; // groups counted separately: reported as n/1000
        } else if matches!(rec.kind, 0x03E9 | 0x03EE | 0x03F8 | 0xF000 | 0xF002 | 0xF003 | 0xF004 | 0xF005 | 0xF00D | 0x040B | 0x040C) {
            n += count(data, body, stop);
        }
        pos = stop;
    }
    n
}

fn main() {
    let a: Vec<String> = std::env::args().skip(1).collect();
    let cfb = office_cfb::Cfb::parse(std::fs::read(&a[0]).unwrap()).unwrap();
    let s = cfb.stream("PowerPoint Document").unwrap().to_vec();
    let top = p::records(&s, 0, s.len());
    let dir = top.iter().find(|r| r.kind == 0x1772).unwrap();
    let body = &s[dir.offset + 8..dir.offset + 8 + dir.len];
    let mut pos = 0;
    let mut total = 0;
    while pos + 4 <= body.len() {
        let w = u32::from_le_bytes(body[pos..pos + 4].try_into().unwrap());
        let (first, count_) = (w & 0xFFFFF, w >> 20);
        pos += 4;
        for i in 0..count_ {
            let off = u32::from_le_bytes(body[pos..pos + 4].try_into().unwrap()) as usize;
            pos += 4;
            let r = p::read_record(&s, off, s.len()).unwrap();
            let n = count(&s, off + 8, off + 8 + r.len);
            if let Some(v) = fdgg(&s, off + 8, off + 8 + r.len) {
                let w: Vec<u32> = v.chunks(4).map(|c| u32::from_le_bytes(c.try_into().unwrap())).collect();
                println!("  id {:<3} FDGG spidMax={} objects={} issued={} clusters={} {:?}", first + i, w[0], w[1], w[2], w[3], &w[4..]);
            }
            if n > 0 {
                println!("  id {:<3} kind {:04X} shapes {n}", first + i, r.kind);
            }
            total += n;
        }
    }
    println!("{} total FSP {} groups {}", a[0], total % 1000, total / 1000);
}
