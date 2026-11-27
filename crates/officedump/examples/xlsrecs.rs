//! Dump the BIFF records of a `.xls` file: kind, length and payload hex.
//!
//! ```text
//! cargo run -q -p officedump --example xlsrecs -- FILE.xls [KIND...]
//! ```
//!
//! With no kinds, every record is printed; otherwise only those whose hex code
//! (e.g. `001C` for `NOTE`) matches one of the arguments. `MAX=0` prints whole
//! payloads instead of the first 64 bytes.

use std::env;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut args = env::args().skip(1);
    let path = args.next().ok_or("usage: xlsrecs FILE.xls [KIND...]")?;
    let filter: Vec<String> = args.map(|a| a.to_uppercase()).collect();
    let max: usize = env::var("MAX").ok().and_then(|v| v.parse().ok()).unwrap_or(64);

    let bytes = std::fs::read(&path)?;
    let cfb = office_cfb::Cfb::parse(bytes)?;
    let stream = cfb.stream("Workbook")?;

    let mut at = 0usize;
    while at + 4 <= stream.len() {
        let kind = u16::from_le_bytes([stream[at], stream[at + 1]]);
        let len = u16::from_le_bytes([stream[at + 2], stream[at + 3]]) as usize;
        let code = format!("{kind:04X}");
        let body = stream.get(at + 4..at + 4 + len).unwrap_or(&[]);
        if filter.is_empty() || filter.contains(&code) {
            let shown = &body[..body.len().min(max)];
            println!(
                "{at:06x} {code} len={len:<5} {}",
                shown.iter().map(|b| format!("{b:02x}")).collect::<Vec<_>>().join(" ")
            );
        }
        at += 4 + len;
    }
    Ok(())
}
