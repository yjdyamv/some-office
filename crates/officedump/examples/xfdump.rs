//! Dump the `XF` table of a workbook, one line per record.
//!
//! `cargo run -q -p officedump --example xfdump -- FILE.xls [row col ...]`
//!
//! Layout per MS-XLS §3.9.11 (`docs/xls.md` §19):
//!
//! ```text
//! 0 ifnt | 2 ifmt | 4 flags | 6 alc/fWrap/alcV | 7 trot | 8 indent
//! 9 unused | 10 dgLeft/dgRight/dgTop/dgBottom (4 bits each)
//! 12 icvLeft(7)/icvRight(7)/grbitDiag(2)
//! 14 icvTop(7)/icvBottom(7)/icvDiag(7)/dgDiag(4)/fls(6)
//! 18 icvFore(7)/icvBack(7)
//! ```

use std::path::PathBuf;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut args = std::env::args().skip(1);
    let path: PathBuf = args.next().ok_or("usage: xfdump FILE.xls")?.into();
    let bytes = std::fs::read(&path)?;
    let cfb = office_cfb::Cfb::parse(bytes)?;
    let stream = cfb.stream("Workbook").or_else(|_| cfb.stream("Book"))?;

    let mut at = 0;
    let mut xfs: Vec<Vec<u8>> = Vec::new();
    while at + 4 <= stream.len() {
        let kind = u16::from_le_bytes([stream[at], stream[at + 1]]);
        let len = u16::from_le_bytes([stream[at + 2], stream[at + 3]]) as usize;
        let p = &stream[at + 4..(at + 4 + len).min(stream.len())];
        if kind == 0x00E0 && p.len() >= 20 {
            xfs.push(p.to_vec());
        }
        if kind == 0x000A {
            break;
        }
        at += 4 + len;
    }

    for (i, p) in xfs.iter().enumerate() {
        let u16_at = |i: usize| u16::from_le_bytes([p[i], p[i + 1]]);
        let u32_at = |i: usize| {
            u32::from_le_bytes([p[i], p[i + 1], p[i + 2], p[i + 3]])
        };
        let dg = u16_at(10);
        let icv = u16_at(12);
        let tail = u32_at(14);
        let bit = |v: u32, from: u32, width: u32| (v >> from) & ((1 << width) - 1);
        println!(
            "xf {i:3} ifnt={:<2} ifmt={:<3} style={} alc={} dg={:x}/{:x}/{:x}/{:x} \
             icv={:x}/{:x}/{:x}/{:x} diag={} dgDiag={:x} fls={} icvFore={:x} icvBack={:x}",
            u16_at(0),
            u16_at(2),
            u16_at(4) >> 2 & 1,
            p[6] & 0x7,
            dg & 0xF,
            dg >> 4 & 0xF,
            dg >> 8 & 0xF,
            dg >> 12,
            icv & 0x7F,
            icv >> 7 & 0x7F,
            bit(tail, 0, 7),
            bit(tail, 7, 7),
            icv >> 14 & 0x3,
            bit(tail, 21, 4),
            bit(tail, 26, 6),
            u16_at(18) & 0x7F,
            u16_at(18) >> 7 & 0x7F,
        );
    }
    println!("{} XF records", xfs.len());

    // Cell -> XF, so the numbers above can be tied to a cell.
    for cell in args {
        let mut parts = cell.split(',');
        let (row, col): (u16, u16) = match (parts.next(), parts.next()) {
            (Some(r), Some(c)) => (r.parse()?, c.parse()?),
            _ => return Err(format!("bad cell {cell:?}, want ROW,COL").into()),
        };
        let mut at = 0;
        let mut found = None;
        while at + 4 <= stream.len() {
            let kind = u16::from_le_bytes([stream[at], stream[at + 1]]);
            let len = u16::from_le_bytes([stream[at + 2], stream[at + 3]]) as usize;
            let p = &stream[at + 4..(at + 4 + len).min(stream.len())];
            if kind == 0x0203 && p.len() >= 6 {
                let r = u16::from_le_bytes([p[0], p[1]]);
                let c = u16::from_le_bytes([p[2], p[3]]);
                if r == row && c == col {
                    found = Some(u16::from_le_bytes([p[4], p[5]]));
                }
            }
            if kind == 0x000A {
                break;
            }
            at += 4 + len;
        }
        println!("cell {row},{col} -> xf {found:?}");
    }
    Ok(())
}
