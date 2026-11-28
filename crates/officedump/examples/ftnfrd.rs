//! Print `PlcffndRef` cut points and the `FRD` words that follow them.

use std::env;
use std::fs;

fn u16_at(d: &[u8], at: usize) -> u16 {
    u16::from_le_bytes([d[at], d[at + 1]])
}
fn u32_at(d: &[u8], at: usize) -> u32 {
    u32::from_le_bytes([d[at], d[at + 1], d[at + 2], d[at + 3]])
}

fn main() {
    let args: Vec<String> = env::args().skip(1).collect();
    for path in &args {
        let bytes = fs::read(path).unwrap();
        let cfb = office_cfb::Cfb::parse(bytes).unwrap();
        let wd = cfb.stream("WordDocument").unwrap();
        let table_name = if wd[0x0B] & 0x02 != 0 { "1Table" } else { "0Table" };
        let table = cfb.stream(table_name).unwrap();
        println!("== {path}");
        for (pair, name) in [(2u32, "PlcffndRef"), (46, "PlcfendRef")] {
            let fc = u32_at(&wd, 0x9A + pair as usize * 8) as usize;
            let lcb = u32_at(&wd, 0x9A + pair as usize * 8 + 4) as usize;
            if lcb == 0 {
                println!("  {name}: empty");
                continue;
            }
            // lcb = 4k + 2(k-1) -> k = (lcb + 2) / 6
            let k = (lcb + 2) / 6;
            let cps: Vec<u32> = (0..k).map(|i| u32_at(&table, fc + i * 4)).collect();
            let frds: Vec<u16> = (0..k - 1)
                .map(|i| u16_at(&table, fc + k * 4 + i * 2))
                .collect();
            println!("  {name}: lcb={lcb} cps={cps:?} frd={frds:?}");
            for (i, frd) in frds.iter().enumerate() {
                println!(
                    "    frd {i}: 0x{frd:04x} fAuto={} fSize={} reserved={}",
                    frd & 1,
                    (frd >> 1) & 1,
                    (frd >> 2)
                );
            }
        }
    }
}
