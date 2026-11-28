//! Dump the header/footer story of a .doc: the `PlcfHdd` character positions
//! and the story text with control characters escaped.

fn main() {
    let path = match std::env::args().nth(1) {
        Some(p) => p,
        None => {
            eprintln!("usage: hdrdump FILE.doc");
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
    let counts: Vec<u32> = (0..8).map(|i| u32_at(&wd, 0x4C + 4 * i)).collect();
    println!(
        "ccpText={} ccpFtn={} ccpHdd={} ccpMcr={} ccpAtn={} ccpEdn={} ccpTxbx={} ccpHdrTxbx={}",
        counts[0], counts[1], counts[2], counts[3], counts[4], counts[5], counts[6], counts[7]
    );
    let fc = u32_at(&wd, 0xF2) as usize;
    let lcb = u32_at(&wd, 0xF6) as usize;
    println!("fcPlcfhdd={fc} lcbPlcfhdd={lcb}");
    let plc = &table[fc..fc + lcb];
    let cps: Vec<u32> = (0..lcb / 4).map(|i| u32_at(plc, i * 4)).collect();
    println!("PlcfHdd cps ({}): {cps:?}", cps.len());
    let sed_fc = u32_at(&wd, 0xCA) as usize;
    let sed_lcb = u32_at(&wd, 0xCE) as usize;
    println!("Plcfsed fc={sed_fc} lcb={sed_lcb} (sections={})", (sed_lcb - 4) / 16);

    // The two endnote PLCs sit at FIB pairs 46 and 47; pairs 4 and 5 are the
    // annotation PLCs, which look identical but hang the notes off the wrong
    // story.
    for (label, pair) in [("PlcfendRef", 46usize), ("PlcfendTxt", 47)] {
        let fc = u32_at(&wd, 0x9A + pair * 8) as usize;
        let lcb = u32_at(&wd, 0x9A + pair * 8 + 4) as usize;
        if lcb == 0 {
            println!("{label}: (empty)");
            continue;
        }
        let plc = &table[fc..fc + lcb];
        let cps: Vec<u32> = (0..lcb / 4).map(|i| u32_at(plc, i * 4)).collect();
        println!("{label} fc={fc} lcb={lcb} cps={cps:?}");
    }
    // The story text, with control characters escaped and their CP offsets.
    let doc = office_doc::open(&path).expect("open");
    let text = doc.story(office_doc::Story::Header);
    println!("header story ({} chars):", text.chars().count());
    let chars: Vec<char> = text.chars().collect();
    for (i, window) in cps.windows(2).enumerate() {
        let from = window[0] as usize;
        let to = window[1] as usize;
        let slice: String = chars.get(from..to).unwrap_or(&[]).iter().collect();
        println!("  slot {i:>2} {from:>3}..{to:<3}: {slice:?}");
    }
    let doc = office_doc::open(std::env::args().nth(1).unwrap()).unwrap();
    for (label, story) in [
        ("footnote", office_doc::Story::Footnote),
        ("endnote", office_doc::Story::Endnote),
    ] {
        let text = doc.story(story);
        println!("{label} story ({} chars): {text:?}", text.chars().count());
    }
    match doc.header_footer(0) {
        Some((header, footer)) => println!("header={header:?} footer={footer:?}"),
        None => println!("(no per-section header slots)"),
    }
}
