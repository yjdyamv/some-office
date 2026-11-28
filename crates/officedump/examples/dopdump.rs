//! Dump the document properties (`DOP`) of `.doc` files.

use std::env;

fn u32_at(data: &[u8], at: usize) -> u32 {
    u32::from_le_bytes([data[at], data[at + 1], data[at + 2], data[at + 3]])
}

fn main() {
    for path in env::args().skip(1) {
        let data = match std::fs::read(&path) {
            Ok(data) => data,
            Err(e) => {
                println!("{path}: {e}");
                continue;
            }
        };
        let cfb = match office_cfb::Cfb::parse(data) {
            Ok(cfb) => cfb,
            Err(e) => {
                println!("{path}: {e}");
                continue;
            }
        };
        let Ok(wd) = cfb.stream("WordDocument") else {
            println!("{path}: no WordDocument");
            continue;
        };
        let table_name = if wd[0x0B] & 0x02 == 0 { "0Table" } else { "1Table" };
        let Ok(table) = cfb.stream(table_name) else {
            println!("{path}: no {table_name}");
            continue;
        };
        let fc = u32_at(&wd, 0x9A + 31 * 8) as usize;
        let lcb = u32_at(&wd, 0x9A + 31 * 8 + 4) as usize;
        println!("== {path}  fcDop={fc} lcbDop={lcb}");
        if fc != 0 && fc + lcb <= table.len() {
            let dop = &table[fc..fc + lcb];
            for row in 0..dop.len().div_ceil(16) {
                let start = row * 16;
                let end = (start + 16).min(dop.len());
                let hex: Vec<String> = dop[start..end].iter().map(|b| format!("{b:02x}")).collect();
                println!("{start:04x}  {}", hex.join(" "));
            }
        }
    }
}
