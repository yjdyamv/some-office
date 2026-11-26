//! Turn a Word document into an XOR-obfuscated one.
//!
//! Usage: `obf <in> <out> <password>`
//!
//! Physical offsets come from `Cfb::physical_runs`: a stream is not necessarily
//! contiguous, and the directory can sit in the middle of the file.
fn main() {
    let a: Vec<String> = std::env::args().collect();
    let (input, output, pw) = (&a[1], &a[2], &a[3]);

    let bytes = std::fs::read(input).unwrap();
    let cfb = office_cfb::Cfb::parse(bytes.clone()).unwrap();
    let arr = office_crypto::xor::array(pw);
    let verifier = office_crypto::xor::verifier(pw);
    let mut file = bytes.clone();

    // 1. Obfuscate the streams. The keystream index is anchored at the end of a
    //    stream, so each is transformed over its logical bytes, then written
    //    back through its physical runs. The FIB prefix stays readable.
    let table = if cfb.stream("1Table").is_ok() { "1Table" } else { "0Table" };
    for (name, plain_prefix) in [("WordDocument", 0x44usize), (table, 0)] {
        let Ok(plain) = cfb.stream(name) else { continue };
        let mut body = office_crypto::xor::encrypt_run(&plain, &arr, 0, false);
        body[..plain_prefix].copy_from_slice(&plain[..plain_prefix]);
        let mut at = 0usize;
        for (off, len) in cfb.physical_runs(name).unwrap() {
            file[off..off + len].copy_from_slice(&body[at..at + len]);
            at += len;
        }
    }

    // 2. Patch the FIB *after* that: fEncrypted | fObfuscated, verifier in
    //    lKey. Both are plaintext.
    let wd_off = cfb.physical_runs("WordDocument").unwrap()[0].0;
    let flags = u16::from_le_bytes([file[wd_off + 10], file[wd_off + 11]]);
    file[wd_off + 10..wd_off + 12].copy_from_slice(&(flags | 0x8100).to_le_bytes());
    file[wd_off + 0x0E..wd_off + 0x12].copy_from_slice(&(verifier as u32).to_le_bytes());

    std::fs::write(output, &file).unwrap();
    println!("wrote {output} (verifier {verifier:#06X})");
}
