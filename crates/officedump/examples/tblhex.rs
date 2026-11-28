//! Hexdump a range of a stream: `tblhex FILE STREAM START LEN`.

use std::env;
use std::fs;

fn main() {
    let args: Vec<String> = env::args().skip(1).collect();
    if args.len() != 4 {
        eprintln!("usage: tblhex FILE STREAM START LEN");
        std::process::exit(2);
    }
    let bytes = fs::read(&args[0]).unwrap();
    let cfb = office_cfb::Cfb::parse(bytes).unwrap();
    let stream = cfb.stream(&args[1]).unwrap();
    let start: usize = args[2].parse().unwrap();
    let len: usize = args[3].parse().unwrap();
    let mut at = start;
    while at < start + len && at < stream.len() {
        let end = (at + 16).min(start + len).min(stream.len());
        let row: Vec<String> = stream[at..end].iter().map(|b| format!("{b:02x}")).collect();
        println!("{at:08x} {}", row.join(" "));
        at = end;
    }
}
