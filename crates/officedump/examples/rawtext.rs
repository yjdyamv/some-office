//! Print the raw text stream of a .doc as code points, with no `ccp*` slicing,
//! so the exact characters at each CP can be compared between two files.

fn main() {
    let mut args = std::env::args().skip(1);
    let path = match args.next() {
        Some(p) => p,
        None => {
            eprintln!("usage: rawtext FILE.doc [FROM] [TO]");
            std::process::exit(2);
        }
    };
    let from: usize = args.next().map(|a| a.parse().unwrap()).unwrap_or(0);
    let to: usize = args.next().map(|a| a.parse().unwrap()).unwrap_or(from + 80);
    let bytes = std::fs::read(&path).expect("read");
    let cfb = office_cfb::Cfb::parse(bytes).expect("cfb");
    let wd = cfb.stream("WordDocument").expect("WordDocument");
    let u32_at = |d: &[u8], i: usize| u32::from_le_bytes([d[i], d[i + 1], d[i + 2], d[i + 3]]);
    let fc_min = u32_at(&wd, 0x18) as usize;
    let fc_mac = u32_at(&wd, 0x1C) as usize;
    let counts: Vec<u32> = (0..8).map(|i| u32_at(&wd, 0x4C + 4 * i)).collect();
    println!("ccp {counts:?} fcMin={fc_min} fcMac={fc_mac}");
    let mut at = fc_min + from * 2;
    for cp in from..to {
        if at + 2 > fc_mac {
            break;
        }
        let c = u16::from_le_bytes([wd[at], wd[at + 1]]);
        let shown = match c {
            0x0D => "\\r".to_string(),
            0x02 => "MARK".to_string(),
            0x03 | 0x04 => format!("\\u{{{c:04x}}}"),
            _ => char::from_u32(c as u32).unwrap_or('?').escape_debug().to_string(),
        };
        println!("  cp {cp:>4} = {shown}");
        at += 2;
    }
}
