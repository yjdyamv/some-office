//! Print the 40-byte `DocumentAtom` payload as eight `u32`s plus the two `u16`
//! in the tail, side by side across files: `cargo run -q --example docatom FILE...`
fn main() {
    for path in std::env::args().skip(1) {
        let cfb = match office_cfb::Cfb::parse(std::fs::read(&path).unwrap()) {
            Ok(c) => c,
            Err(e) => {
                println!("{path}: {e}");
                continue;
            }
        };
        let data = cfb.stream("PowerPoint Document").unwrap().to_vec();
        let Some(doc) = office_ppt::records(&data, 0, data.len())
            .into_iter()
            .find(|r| r.kind == 0x03E8)
        else {
            println!("{path}: no Document container");
            continue;
        };
        let Some(rec) = office_ppt::records(&data, doc.offset + 8, doc.offset + 8 + doc.len)
            .into_iter()
            .find(|r| r.kind == 0x03E9)
        else {
            println!("{path}: no DocumentAtom");
            continue;
        };
        let p = rec.payload(&data);
        let w = |i: usize| u32::from_le_bytes([p[i], p[i + 1], p[i + 2], p[i + 3]]);
        let u = |i: usize| u16::from_le_bytes([p[i], p[i + 1]]);
        let name = std::path::Path::new(&path).file_name().map(|n| n.to_string_lossy().to_string()).unwrap_or_else(|| path.clone());
        print!("{name:24}");
        for i in (0..40).step_by(4) {
            print!(" {:>8}", w(i));
        }
        println!("   u16@32={} u16@34={} atom_len={}", u(32), u(34), rec.len);
        let hex: Vec<String> = p.iter().map(|b| format!("{b:02X}")).collect();
        println!("   {}", hex.join(" "));
    }
}
