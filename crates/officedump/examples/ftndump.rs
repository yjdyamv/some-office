//! Dump the footnote-related parts of a .doc: the FibRgLw97 counts, the
//! `PlcffndRef`/`PlcffndTxt` pairs and the footnote reference characters in the
//! main text.
//!
//! `cargo run -q -p officedump --example ftndump -- samples/ms/footnote.doc`
use office_doc::Story;

fn pair(word: &[u8], index: usize) -> (u32, u32) {
    let o = 0x9A + index * 8;
    let rd = |at: usize| u32::from_le_bytes(word[at..at + 4].try_into().unwrap());
    (rd(o), rd(o + 4))
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let path = std::env::args().nth(1).ok_or("usage: ftndump FILE.doc")?;
    let bytes = std::fs::read(&path)?;
    let cfb = office_cfb::Cfb::parse(bytes)?;
    let word = cfb.stream("WordDocument")?;
    let table = cfb.stream(if word[0x0B] & 0x02 != 0 { "1Table" } else { "0Table" })?;

    let rd = |at: usize| u32::from_le_bytes(word[at..at + 4].try_into().unwrap());
    println!("ccpText={} ccpFtn={} ccpHdd={}", rd(0x4C), rd(0x50), rd(0x54));
    for (name, index) in [("PlcffndRef", 2usize), ("PlcffndTxt", 3)] {
        let (fc, lcb) = pair(&word, index);
        print!("{name} pair={index} fc={fc} lcb={lcb} ->");
        if lcb == 0 {
            println!(" (empty)");
            continue;
        }
        let plc = &table[fc as usize..(fc + lcb) as usize];
        let n = (lcb as usize - 4) / 6;
        let cps: Vec<u32> = (0..=n).map(|i| u32::from_le_bytes(plc[i * 4..i * 4 + 4].try_into().unwrap())).collect();
        let data: Vec<u16> = (0..n)
            .map(|i| u16::from_le_bytes(plc[(n + 1) * 4 + i * 2..(n + 1) * 4 + i * 2 + 2].try_into().unwrap()))
            .collect();
        println!(" cps={cps:?} data={data:?}");
    }
    let doc = office_doc::open(&path)?;
    let (start, len) = doc.story_range(Story::Main);
    let full = doc.full_text.clone();
    let main: Vec<char> = full.chars().skip(start as usize).take(len as usize).collect();
    println!("main text: {:?}", main.iter().collect::<String>());
    println!("main codes: {:?}", main.iter().map(|c| *c as u32).collect::<Vec<_>>());
    println!("hdd_slots={:?}", doc.hdd_slots);
    for story in [Story::Main, Story::Footnote, Story::Header] {
        let (a, b) = doc.story_range(story);
        let piece: Vec<u32> = full.chars().skip(a as usize).take((b - a) as usize).map(|c| c as u32).collect();
        println!("{story:?} [{a}..{b}] = {piece:?}");
    }
    Ok(())
}