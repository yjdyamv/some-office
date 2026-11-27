fn main() {
    let png = std::fs::read("samples/pixel.png").unwrap();
    let paragraphs = vec![
        office_doc::write::Paragraph::text("before"),
        office_doc::write::Paragraph::styled(
            vec![office_doc::write::Run::picture(office_doc::write::InlinePicture::new(png.clone(), 240, 240))],
            office_doc::fmt::ParaProps::default(),
        ),
        office_doc::write::Paragraph::text("after"),
    ];
    let bytes = office_doc::write::write_document(&paragraphs);
    let cfb = office_cfb::Cfb::parse(bytes.clone()).unwrap();
    println!("streams: {:?}", cfb.entries().iter().map(|e| (e.name.clone(), e.size)).collect::<Vec<_>>());
    let doc = office_doc::parse(&bytes).unwrap();
    println!("data len {} text {:?}", doc.data.len(), doc.text);
    for r in &doc.char_runs {
        println!("  run {:x}..{:x} pic {:?} spec {:?}", r.fc_start, r.fc_end, r.props.pic_location, r.props.spec);
    }
    println!("pictures {}", doc.pictures().len());
}
// extra probe appended
