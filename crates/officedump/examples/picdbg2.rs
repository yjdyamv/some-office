fn main() {
    let data = std::fs::read("crates/office-doc/templates/picture-data.bin").unwrap();
    println!("template len {}", data.len());
    println!("from_bse(template): {:?}", office_blip::from_bse(&data[0xD6 + 8..]).map(|b| (b.kind, b.data.len())));
    let png = std::fs::read("samples/pixel.png").unwrap();
    let paragraphs = vec![office_doc::write::Paragraph::styled(
        vec![office_doc::write::Run::picture(office_doc::write::InlinePicture::new(png, 240, 240))],
        office_doc::fmt::ParaProps::default(),
    )];
    let bytes = office_doc::write::write_document(&paragraphs);
    let doc = office_doc::parse(&bytes).unwrap();
    println!("written data len {}", doc.data.len());
    let fbse_at = 0xD6usize;
    let len = u32::from_le_bytes(doc.data[fbse_at + 4..fbse_at + 8].try_into().unwrap());
    println!("fbse len {}", len);
    println!("from_bse(written): {:?}", office_blip::from_bse(&doc.data[fbse_at + 8..]).map(|b| (b.kind, b.data.len())));
    let blip_at = fbse_at + 44;
    println!("blip hdr {:02x?}", &doc.data[blip_at..blip_at + 8]);
    println!("from_blip(written): {:?}", office_blip::from_blip(&doc.data[blip_at..]).map(|b| (b.kind, b.data.len())));
}
