//! Dump the piece table and the formatting pages of a .doc, for comparing two
//! files that Word treats differently.

fn main() {
    let path = match std::env::args().nth(1) {
        Some(p) => p,
        None => {
            eprintln!("usage: fkpdump FILE.doc");
            std::process::exit(2);
        }
    };
    let bytes = std::fs::read(&path).expect("read");
    let cfb = office_cfb::Cfb::parse(bytes).expect("cfb");
    let wd = cfb.stream("WordDocument").expect("WordDocument");
    let table = cfb
        .stream(if wd[0x0B] & 0x02 != 0 { "1Table" } else { "0Table" })
        .expect("table stream");
    let u32_at = |d: &[u8], i: usize| u32::from_le_bytes([d[i], d[i + 1], d[i + 2], d[i + 3]]);
    let u16_at = |d: &[u8], i: usize| u16::from_le_bytes([d[i], d[i + 1]]);
    let counts: Vec<u32> = (0..8).map(|i| u32_at(&wd, 0x4C + 4 * i)).collect();
    println!("ccp {counts:?}");
    println!("flags={:#06x} fcMin={} fcMac={} cbMac={}", u16_at(&wd, 0x0A), u32_at(&wd, 0x18), u32_at(&wd, 0x1C), u32_at(&wd, 0x40));

    // Piece table: CLX is FIB pair 33.
    let clx_fc = u32_at(&wd, 0x9A + 33 * 8) as usize;
    let clx_lcb = u32_at(&wd, 0x9A + 33 * 8 + 4) as usize;
    let clx = &table[clx_fc..clx_fc + clx_lcb];
    let mut at = 0usize;
    while at < clx.len() {
        let kind = clx[at];
        let len = u32_at(clx, at + 1) as usize;
        if kind == 2 {
            let cps: Vec<u32> = (0..(len - 4) / 12).map(|i| u32_at(clx, at + 5 + i * 4)).collect();
            let pcds: Vec<(u32, u16)> = (0..(len - 4) / 12)
                .map(|i| (u32_at(clx, at + 5 + (len - 4) / 12 * 4 + i * 8) & 0x3FFF_FFFF, u16_at(clx, at + 5 + (len - 4) / 12 * 4 + i * 8 + 6)))
                .collect();
            println!("pieces cps={cps:?}");
            for (i, (fc, prm)) in pcds.iter().enumerate() {
                println!("  piece {i}: fc={fc} (compressed={}) prm={prm:#06x}", fc & 0x4000_0000 != 0);
            }
        } else {
            println!("clx record kind={kind} len={len}");
        }
        at += 5 + len;
    }

    println!("stream len={} cbMac={}", wd.len(), u32_at(&wd, 0x40));
    // Section table: FIB pair 6.
    let sed_fc = u32_at(&wd, 0x9A + 6 * 8) as usize;
    let sed_lcb = u32_at(&wd, 0x9A + 6 * 8 + 4) as usize;
    if sed_lcb >= 4 {
        let n = (sed_lcb - 4) / 16;
        let cps: Vec<u32> = (0..=n).map(|i| u32_at(&table, sed_fc + i * 4)).collect();
        println!("Plcfsed cps={cps:?}");
    }

    // Formatting pages: FIB pairs 12 (chpx) and 13 (papx) point at the page tables.
    for (label, pair) in [("chpx", 12usize), ("papx", 13)] {
        let plc_fc = u32_at(&wd, 0x9A + pair * 8) as usize;
        let plc_lcb = u32_at(&wd, 0x9A + pair * 8 + 4) as usize;
        if plc_lcb == 0 {
            println!("{label}: no pages");
            continue;
        }
        let plc = &table[plc_fc..plc_fc + plc_lcb];
        let n = (plc_lcb - 4) / 8;
        for i in 0..n {
            let pn = u32_at(plc, (n + 1) * 4 + i * 4) & 0x3FF_FFFF;
            let page = &wd[pn as usize * 512..pn as usize * 512 + 512];
            let crun = page[511] as usize;
            let fcs: Vec<u32> = (0..=crun).map(|j| u32_at(page, j * 4)).collect();
            println!("{label} page {pn} crun={crun} fcs={fcs:?}");
        }
    }
}
