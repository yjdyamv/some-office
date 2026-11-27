//! Dump every text-related atom with its offset and decoded text, so count
//! mismatches can be traced to the box they belong to:
//! `cargo run -q --example tatoms FILE [--hex]`.
fn decode_utf16le(b: &[u8]) -> String {
    let units: Vec<u16> = b.chunks_exact(2).map(|c| u16::from_le_bytes([c[0], c[1]])).collect();
    String::from_utf16_lossy(&units)
}

fn walk(data: &[u8], start: usize, end: usize, depth: usize, hex: bool) {
    for rec in office_ppt::records(data, start, end) {
        match rec.kind {
            0x0F9F | 0x0FA0 | 0x0FA1 | 0x0FA8 | 0x0FAA => {
                let p = rec.payload(data);
                let head = if p.len() >= 4 { format!(" first={}", u32::from_le_bytes([p[0], p[1], p[2], p[3]])) } else { String::new() };
                let text = match rec.kind {
                    0x0FA0 => format!(" {:?}", decode_utf16le(p)),
                    0x0FA8 => format!(" {:?}", String::from_utf8_lossy(p)),
                    _ => String::new(),
                };
                let dump = if hex && (rec.kind == 0x0FAA || rec.kind == 0x0FA1) {
                    let n = p.len().min(48);
                    let s: Vec<String> = p[..n].iter().map(|b| format!("{b:02X}")).collect();
                    format!(" [{}]", s.join(" "))
                } else {
                    String::new()
                };
                println!("{:indent$}{:#06x} @{:#x} len={}{}{}{}", "", rec.kind, rec.offset, rec.len, head, text, dump, indent = depth * 2);
            }
            _ => {}
        }
        if depth < 8 && (rec.is_container() || rec.kind == 0xF00D) {
            walk(data, rec.offset + 8, rec.offset + 8 + rec.len, depth + 1, hex);
        }
    }
}

fn main() {
    let args: Vec<String> = std::env::args().skip(1).collect();
    let hex = args.iter().any(|a| a == "--hex");
    let path = &args[0];
    let cfb = office_cfb::Cfb::parse(std::fs::read(path).unwrap()).unwrap();
    let data = cfb.stream("PowerPoint Document").unwrap().to_vec();
    walk(&data, 0, data.len(), 0, hex);
}
