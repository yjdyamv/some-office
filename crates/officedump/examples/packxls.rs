//! Wrap a raw `Workbook` stream in a CFB container: `packxls <stream.bin> <out.xls>`.
//!
//! Handy when debugging the writer — patch the stream with a script, repack it,
//! and hand it to Excel without rebuilding the crate.

fn main() {
    let mut args = std::env::args().skip(1);
    let (src, out) = match (args.next(), args.next()) {
        (Some(a), Some(b)) => (a, b),
        _ => {
            eprintln!("usage: packxls <stream.bin> <out.xls>");
            std::process::exit(2);
        }
    };
    let stream = std::fs::read(&src).unwrap();
    let bytes = office_cfb::write(&[office_cfb::OutputStream::new("Workbook", &stream)]);
    std::fs::write(&out, &bytes).unwrap();
    println!("{out} {} bytes (stream {})", bytes.len(), stream.len());
}
