//! Print the summary property sets of a .doc/.xls/.ppt file.
//!
//! `cargo run -q -p officedump --example meta -- FILE...`

use office_cfb::Cfb;

fn main() {
    for path in std::env::args().skip(1) {
        println!("{path}");
        let bytes = std::fs::read(&path).expect("read");
        let cfb = match Cfb::parse(bytes) {
            Ok(cfb) => cfb,
            Err(e) => {
                println!("  not a compound file: {e}");
                continue;
            }
        };
        match cfb.stream(office_oleps::SUMMARY_STREAM) {
            Ok(data) => match office_oleps::Summary::from_bytes(&data) {
                Some(s) => {
                    println!("  summary ({} bytes)", data.len());
                    println!("    title={:?} subject={:?}", s.title, s.subject);
                    println!("    author={:?} last_author={:?}", s.author, s.last_author);
                    println!("    keywords={:?} comments={:?}", s.keywords, s.comments);
                    println!("    template={:?} revision={:?}", s.template, s.revision);
                    println!("    application={:?}", s.application);
                    println!(
                        "    created=0x{:016x} saved=0x{:016x}",
                        s.created, s.saved
                    );
                    println!(
                        "    pages={} words={} chars={} security={}",
                        s.pages, s.words, s.characters, s.doc_security
                    );
                }
                None => println!("  summary: unrecognised property set"),
            },
            Err(_) => println!("  no \x05SummaryInformation"),
        }
        match cfb.stream(office_oleps::DOC_SUMMARY_STREAM) {
            Ok(data) => match office_oleps::DocSummary::from_bytes(&data) {
                Some(s) => println!(
                    "  doc summary: category={:?} manager={:?} company={:?} lines={} paras={} slides={} notes={}",
                    s.category, s.manager, s.company, s.lines, s.paragraphs, s.slides, s.notes
                ),
                None => println!("  doc summary: unrecognised property set"),
            },
            Err(_) => println!("  no \x05DocumentSummaryInformation"),
        }
    }
}
