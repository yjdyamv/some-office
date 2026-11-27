//! Hexdump a byte range of the `PowerPoint Document` stream.
//!
//! ```text
//! cargo run -q -p officedump --example hexrange samples/notes.ppt 61022 640
//! ```

fn main() {
    let args: Vec<String> = std::env::args().skip(1).collect();
    let path = args.first().expect("path");
    let start: usize = args.get(1).map(|v| v.parse().unwrap()).unwrap_or(0);
    let len: usize = args.get(2).map(|v| v.parse().unwrap()).unwrap_or(256);
    let bytes = std::fs::read(path).expect("read");
    let cfb = office_cfb::Cfb::parse(bytes).expect("cfb");
    let stream = cfb.stream("PowerPoint Document").expect("stream");
    let end = (start + len).min(stream.len());
    for (i, chunk) in stream[start..end].chunks(16).enumerate() {
        let hex: Vec<String> = chunk.iter().map(|b| format!("{b:02x}")).collect();
        let text: String = chunk
            .iter()
            .map(|b| if (0x20..0x7F).contains(b) { *b as char } else { '.' })
            .collect();
        println!("{:08x}  {:<47}  {}", start + i * 16, hex.join(" "), text);
    }
}
