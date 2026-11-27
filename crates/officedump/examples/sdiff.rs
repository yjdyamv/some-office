//! Byte-diff two persist objects (default id 3) of two files: `sdiff A B [id]`.
use office_ppt as p;

fn blob(path: &str, id: u32) -> Vec<u8> {
    let cfb = office_cfb::Cfb::parse(std::fs::read(path).unwrap()).unwrap();
    let s = cfb.stream("PowerPoint Document").unwrap().to_vec();
    let top = p::records(&s, 0, s.len());
    let dir = top.iter().find(|r| r.kind == 0x1772).unwrap();
    let body = &s[dir.offset + 8..dir.offset + 8 + dir.len];
    let mut pos = 0;
    while pos + 4 <= body.len() {
        let w = u32::from_le_bytes(body[pos..pos + 4].try_into().unwrap());
        let (first, count) = (w & 0xFFFFF, w >> 20);
        pos += 4;
        for i in 0..count {
            let off = u32::from_le_bytes(body[pos..pos + 4].try_into().unwrap()) as usize;
            pos += 4;
            if first + i == id {
                let r = p::read_record(&s, off, s.len()).unwrap();
                return s[off..off + 8 + r.len].to_vec();
            }
        }
    }
    panic!("no persist id {id}");
}

fn main() {
    let a: Vec<String> = std::env::args().skip(1).collect();
    let id: u32 = a.get(2).map(|s| s.parse().unwrap()).unwrap_or(3);
    let (x, y) = (blob(&a[0], id), blob(&a[1], id));
    println!("{} {} bytes / {} {} bytes", a[0], x.len(), a[1], y.len());
    let n = x.len().min(y.len());
    let mut shown = 0;
    let mut i = 0;
    while i < n && shown < 12 {
        if x[i] != y[i] {
            let from = i.saturating_sub(8);
            let to = (i + 16).min(n);
            println!("@{i}: A {:02x?}\n     B {:02x?}", &x[from..to], &y[from..to]);
            shown += 1;
            i = to;
        } else {
            i += 1;
        }
    }
}
