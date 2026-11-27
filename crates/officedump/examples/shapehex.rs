//! Dumps one shape's `SpContainer` byte for byte, and the synthetic shape the
//! writer builds, so the two can be diffed.
//!
//! `cargo run -q --example shapehex samples/ms/minimal.ppt 0x3403`

use office_ppt::write;

fn rec(data: &[u8], at: usize) -> Option<(u16, u16, u16, usize)> {
    if at + 8 > data.len() {
        return None;
    }
    let w = u16::from_le_bytes([data[at], data[at + 1]]);
    let kind = u16::from_le_bytes([data[at + 2], data[at + 3]]);
    let len = u32::from_le_bytes([data[at + 4], data[at + 5], data[at + 6], data[at + 7]]) as usize;
    Some((w & 0xF, w >> 4, kind, len))
}

/// Walks `data`, looking for an `SpContainer` (0xF004) whose `Sp` (0xF00A)
/// carries `spid`; returns the container's byte range.
fn find_sp(data: &[u8], spid: u32) -> Option<(usize, usize)> {
    let mut at = 0usize;
    while let Some((_, _, kind, len)) = rec(data, at) {
        let end = at + 8 + len;
        if end > data.len() {
            return None;
        }
        if kind == 0xF004 {
            // Direct children: find the Sp and read its spid.
            let mut child = at + 8;
            while let Some((_, _, k, l)) = rec(data, child) {
                if k == 0xF00A && child + 12 <= data.len() {
                    let id = u32::from_le_bytes([data[child + 8], data[child + 9], data[child + 10], data[child + 11]]);
                    if id == spid {
                        return Some((at, end));
                    }
                }
                child += 8 + l;
            }
        }
        if let Some(found) = find_sp(&data[at + 8..end], spid) {
            return Some((at + 8 + found.0, at + 8 + found.1));
        }
        at = end;
    }
    None
}

fn dump(data: &[u8], indent: usize) {
    let mut at = 0usize;
    while let Some((ver, inst, kind, len)) = rec(data, at) {
        let name = match kind {
            0xF000 => "DggContainer",
            0xF002 => "DgContainer",
            0xF003 => "SpgrContainer",
            0xF004 => "SpContainer",
            0xF008 => "FDG",
            0xF009 => "FSPGR",
            0xF00A => "FSP",
            0xF00B => "FOPT",
            0xF00D => "ClientTextbox",
            0xF010 => "ClientAnchor",
            0xF011 => "ClientData",
            0x0F9F => "TextHeaderAtom",
            0x0FA0 => "TextCharsAtom",
            0x0FA1 => "StyleTextPropAtom",
            0x0FAA => "TextSpecialInfoAtom",
            0x0BC3 => "0x0BC3",
            0x041F => "0x041F",
            _ => "?",
        };
        println!(
            "{:indent$}@{} {name} v{ver} i{inst} len={len}",
            "",
            at,
            indent = indent
        );
        if kind == 0xF00A {
            let id = u32::from_le_bytes([data[at + 8], data[at + 9], data[at + 10], data[at + 11]]);
            println!("{:indent$}  spid=0x{id:X}", "", indent = indent);
        }
        if kind == 0xF00B {
            let mut p = at + 8;
            while p + 6 <= at + 8 + len {
                let pid = u16::from_le_bytes([data[p], data[p + 1]]);
                let v = u32::from_le_bytes([data[p + 2], data[p + 3], data[p + 4], data[p + 5]]);
                println!("{:indent$}  pid=0x{pid:04X} = 0x{v:08X} ({v})", "", indent = indent);
                p += 6;
            }
        }
        if kind == 0xF011 {
            let p = at + 8;
            println!("{:indent$}  hex {indent}", hex(&data[p..p + len]), indent = indent);
        }
        if kind == 0x0FA1 || kind == 0x0FAA || kind == 0x0F9F || kind == 0x0BC3 || kind == 0x041F {
            println!("{:indent$}  hex {indent}", hex(&data[at + 8..at + 8 + len]), indent = indent);
        }
        if kind == 0xF010 {
            println!("{:indent$}  hex {indent}", hex(&data[at + 8..at + 8 + len]), indent = indent);
        }
        if len > 0 && matches!(kind, 0xF000 | 0xF002 | 0xF003 | 0xF004 | 0xF00D) {
            dump(&data[at + 8..at + 8 + len], indent + 2);
        }
        at += 8 + len;
    }
}

fn hex(b: &[u8]) -> String {
    b.iter().map(|x| format!("{x:02x}")).collect::<Vec<_>>().join(" ")
}

/// Prints a one-line summary per shape: spid, FOPT instance/payload, and which
/// child records the shape carries.
fn list(data: &[u8], indent: usize) {
    let mut at = 0usize;
    while let Some((_, _, kind, len)) = rec(data, at) {
        let end = at + 8 + len;
        if end > data.len() {
            return;
        }
        if kind == 0xF004 {
            let mut child = at + 8;
            let mut spid = 0u32;
            let mut kinds = Vec::new();
            let mut fopt = String::new();
            while let Some((_, inst, k, l)) = rec(data, child) {
                if k == 0xF00A {
                    spid = u32::from_le_bytes([data[child + 8], data[child + 9], data[child + 10], data[child + 11]]);
                } else if k == 0xF00B {
                    fopt = format!("i{inst}/len{l}");
                } else if k == 0xF011 {
                    let p = &data[child + 8..child + 8 + l];
                    let a = if l >= 12 { u32::from_le_bytes([p[8], p[9], p[10], p[11]]) } else { 0 };
                    let b = if l >= 24 { u32::from_le_bytes([p[20], p[21], p[22], p[23]]) } else { 0 };
                    kinds.push(format!("CD(len{l},0x{a:X},0x{b:X})"));
                } else if k == 0xF00D {
                    let mut t = child + 8;
                    let mut inner = Vec::new();
                    while let Some((_, _, ik, il)) = rec(&data[..child + 8 + l], t) {
                        inner.push(format!("{ik:04X}/len{il}"));
                        t += 8 + il;
                    }
                    kinds.push(format!("TB[{}]", inner.join(",")));
                } else {
                    kinds.push(format!("{k:04X}"));
                }
                child += 8 + l;
            }
            println!("{:indent$}spid=0x{spid:04X} FOPT {fopt} | {}", "", kinds.join(" "), indent = indent);
        }
        if matches!(kind, 0xF000 | 0xF002 | 0xF003 | 0xF004 | 0xF00D | 0x040C | 0x040B) {
            list(&data[at + 8..end], indent + 1);
        }
        at = end;
    }
}

fn main() {
    let args: Vec<String> = std::env::args().skip(1).collect();
    if args.len() < 2 {
        println!("usage: shapehex FILE SPID|--list");
        return;
    }
    if args[1] == "--list" {
        let bytes = std::fs::read(&args[0]).expect("read");
        let cfb = office_cfb::Cfb::parse(bytes).expect("cfb");
        let stream = cfb.stream("PowerPoint Document").expect("stream");
        for r in office_ppt::records(&stream, 0, stream.len()) {
            if matches!(r.kind, 0x03F8 | 0x03EE) {
                println!("-- object @{} kind {:04X} len {}", r.offset, r.kind, r.len);
                list(r.payload(&stream), 1);
            }
        }
        return;
    }
    let spid = u32::from_str_radix(args[1].trim_start_matches("0x"), 16).unwrap();
    let bytes = std::fs::read(&args[0]).expect("read");
    let cfb = office_cfb::Cfb::parse(bytes).expect("cfb");
    let stream = cfb.stream("PowerPoint Document").expect("stream");
    for r in office_ppt::records(&stream, 0, stream.len()) {
        if let Some((a, b)) = find_sp(r.payload(&stream), spid) {
            println!("-- found in object @{} kind {:04X}", r.offset, r.kind);
            let (s, e) = (r.offset + 8 + a, r.offset + 8 + b);
            println!("== real shape @{s}..{e} ({} bytes)", e - s);
            dump(&stream[s..e], 0);
            println!("== hex {}", hex(&stream[s..e]));
            break;
        }
    }
    let (s, e) = (0usize, 0usize);
    let _ = (s, e);
    println!("== real shape @{s}..{e} ({} bytes)", e - s);
    dump(&stream[s..e], 0);
    println!("== hex {}", hex(&stream[s..e]));

    let ours = write::text_shape(spid, 6, "Generated shape", [0x02F3, 0x1B10, 0x0893], 24, false);
    println!("== text_shape ({} bytes)", ours.len());
    dump(&ours, 0);
    println!("== hex {}", hex(&ours));
}
