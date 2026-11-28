//! Print the section properties (`SEPX`) of a `.doc` file.

use std::env;
use std::fs;

fn u16_at(data: &[u8], at: usize) -> u16 {
    u16::from_le_bytes([data[at], data[at + 1]])
}

fn u32_at(data: &[u8], at: usize) -> u32 {
    u32::from_le_bytes([data[at], data[at + 1], data[at + 2], data[at + 3]])
}

/// Operand width in bytes, from the `spra` bits of a sprm's identifier.
fn operand_len(sprm: u16) -> usize {
    match (sprm >> 13) & 0x7 {
        0 | 1 => 1,
        2 | 4 | 5 => 2,
        3 => 4,
        6 => {
            // Variable length: the first byte holds the count.
            0
        }
        _ => 3,
    }
}

fn main() {
    let args: Vec<String> = env::args().skip(1).collect();
    if args.is_empty() {
        eprintln!("usage: sepxdump FILE.doc...");
        std::process::exit(2);
    }
    for path in args {
        let bytes = fs::read(&path).unwrap();
        let cfb = office_cfb::Cfb::parse(bytes).unwrap();
        let wd = cfb.stream("WordDocument").unwrap();
        let table_name = if wd[0x0B] & 0x02 != 0 { "1Table" } else { "0Table" };
        let table = cfb.stream(table_name).unwrap();
        let flags = u16_at(&wd, 0x0A);
        println!("== {path}  (table {table_name}, flags {flags:#06x})");
        // Pair 6 is fcPlcfsed/lcbPlcfsed.
        let fc = u32_at(&wd, 0x9A + 6 * 8) as usize;
        let lcb = u32_at(&wd, 0x9A + 6 * 8 + 4) as usize;
        if lcb < 16 {
            println!("  no section table");
            continue;
        }
        let sections = (lcb - 4) / 16;
        println!("  sections={sections} fcPlcfsed={fc} lcb={lcb}");
        for i in 0..sections {
            let at = fc + i * 16 + (sections + 1) * 4;
            let fc_sepx = u32_at(&table, at + 2) as usize;
            print!("  section {i}: fcSepx={fc_sepx}");
            if fc_sepx == 0xFFFF_FFFF || fc_sepx + 2 > wd.len() {
                println!(" (none)");
                continue;
            }
            // `fcSepx` is an offset into the WordDocument stream, not the table.
            let cb = u16_at(&wd, fc_sepx) as usize;
            println!(" cb={cb}");
            let body = &wd[fc_sepx + 2..(fc_sepx + 2 + cb).min(wd.len())];
            let mut pos = 0;
            while pos + 2 <= body.len() {
                let sprm = u16_at(body, pos);
                let fixed = operand_len(sprm);
                let (size, start) = if fixed == 0 {
                    let n = body.get(pos + 2).copied().unwrap_or(0) as usize;
                    (n, pos + 3)
                } else {
                    (fixed, pos + 2)
                };
                if start + size > body.len() {
                    println!("    {sprm:#06x}: truncated");
                    break;
                }
                let operand = &body[start..start + size];
                let shown: String = operand.iter().map(|b| format!("{b:02x}")).collect();
                let value = match size {
                    1 => operand[0] as i64,
                    2 => u16::from_le_bytes([operand[0], operand[1]]) as i16 as i64,
                    4 => u32::from_le_bytes([operand[0], operand[1], operand[2], operand[3]]) as i32 as i64,
                    _ => 0,
                };
                println!("    sprm {sprm:#06x} len={size} bytes={shown} value={value}");
                pos = start + size;
            }
        }
    }
}
