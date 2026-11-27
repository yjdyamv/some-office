fn main() {
    let png = std::fs::read("samples/pixel.png").unwrap();
    let paragraphs = vec![
        office_doc::write::Paragraph::text("before"),
        office_doc::write::Paragraph::styled(
            vec![office_doc::write::Run::picture(office_doc::write::InlinePicture::new(png, 240, 240))],
            office_doc::fmt::ParaProps::default(),
        ),
    ];
    let bytes = office_doc::write::write_document(&paragraphs);
    let doc = office_doc::parse(&bytes).unwrap();
    let d = &doc.data;
    let lcb = u32::from_le_bytes(d[0..4].try_into().unwrap()) as usize;
    let cb = u16::from_le_bytes(d[4..6].try_into().unwrap()) as usize;
    println!("lcb {lcb} cb {cb} len {}", d.len());
    let mut pos = cb.max(0x44);
    while pos + 8 <= lcb.min(d.len()) {
        let vi = u16::from_le_bytes([d[pos], d[pos + 1]]);
        let kind = u16::from_le_bytes([d[pos + 2], d[pos + 3]]);
        let len = u32::from_le_bytes(d[pos + 4..pos + 8].try_into().unwrap()) as usize;
        println!("pos {pos:#x} vi {vi:#x} kind {kind:#x} len {len}");
        pos += 8 + len;
    }
    println!("pictures {}", doc.pictures().len());
}
