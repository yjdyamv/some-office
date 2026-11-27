//! Per text box: text length, paragraph count, `0x0FA1` and `0x0FAA` first
//! fields. `cargo run -q --example tbox FILE` — to settle what those counts mean.
fn decode_utf16le(b: &[u8]) -> String {
    let units: Vec<u16> = b.chunks_exact(2).map(|c| u16::from_le_bytes([c[0], c[1]])).collect();
    String::from_utf16_lossy(&units)
}

#[derive(Default)]
struct Box_ {
    text: Option<String>,
    style: Option<(u32, usize)>,
    special: Option<(u32, usize)>,
    kind: Option<u32>,
}

fn walk(data: &[u8], start: usize, end: usize, depth: usize, out: &mut Vec<Box_>) {
    let mut cur = Box_::default();
    for rec in office_ppt::records(data, start, end) {
        let p = rec.payload(data);
        match rec.kind {
            0x0F9F if p.len() >= 4 => cur.kind = Some(u32::from_le_bytes([p[0], p[1], p[2], p[3]])),
            0x0FA0 => cur.text = Some(decode_utf16le(p)),
            0x0FA8 => cur.text = Some(String::from_utf8_lossy(p).to_string()),
            0x0FA1 if p.len() >= 4 => cur.style = Some((u32::from_le_bytes([p[0], p[1], p[2], p[3]]), rec.len)),
            0x0FAA if p.len() >= 4 => cur.special = Some((u32::from_le_bytes([p[0], p[1], p[2], p[3]]), rec.len)),
            _ => {}
        }
        if depth < 8 && (rec.is_container() || rec.kind == 0xF00D) {
            walk(data, rec.offset + 8, rec.offset + 8 + rec.len, depth + 1, out);
        }
    }
    if cur.text.is_some() || cur.style.is_some() {
        out.push(cur);
    }
}

fn main() {
    let path = std::env::args().nth(1).expect("usage: tbox FILE");
    let cfb = office_cfb::Cfb::parse(std::fs::read(&path).unwrap()).unwrap();
    let data = cfb.stream("PowerPoint Document").unwrap().to_vec();
    let mut out = Vec::new();
    walk(&data, 0, data.len(), 0, &mut out);
    for b in &out {
        let text = b.text.clone().unwrap_or_default();
        let chars = text.chars().count();
        let paras = text.matches('\r').count() + 1;
        let shown: String = text.chars().take(24).collect::<String>().replace('\r', "|");
        println!(
            "type={:?} chars={:3} paras={} style={:?} special={:?} {:?}",
            b.kind, chars, paras, b.style, b.special, shown
        );
    }
}
