//! Multi-slide decks through `write_from_slide_template`: how many pages does
//! PowerPoint accept, and do repeated shape ids hurt?
//!
//! `cargo run -q --example scratch9` → `samples/scratch9-*.ppt`.
use office_ppt::write::{self, PictureSpec, SlideSpec};

fn deck(n: usize) -> Vec<SlideSpec> {
    (1..=n)
        .map(|i| {
            let mut s = SlideSpec::new(format!("Slide {i}"));
            s.body = vec![format!("body line for slide {i}"), "second line".into()];
            s
        })
        .collect()
}

fn main() {
    let tpl = std::fs::read("samples/ms/minimal.ppt").unwrap();
    for n in [1usize, 3, 5, 10] {
        let out = write::write_from_slide_template(&tpl, &deck(n)).unwrap();
        let path = format!("samples/scratch9-{n}.ppt");
        std::fs::write(&path, &out).unwrap();
        println!("{path} {} bytes", out.len());
    }

    let png = std::fs::read("samples/pixel.png").unwrap();
    let mut with_pics: Vec<SlideSpec> = (1..=3)
        .map(|i| {
            let mut s = SlideSpec::new(format!("Picture {i}"));
            s.body = vec![format!("caption {i}")];
            s.pictures = vec![PictureSpec::new(png.clone(), 120, 120)];
            s
        })
        .collect();
    with_pics[1].pictures.push(PictureSpec::new(png.clone(), 60, 60));
    let out = write::write_from_slide_template(&tpl, &with_pics).unwrap();
    std::fs::write("samples/scratch9-pics.ppt", &out).unwrap();
    println!("samples/scratch9-pics.ppt {} bytes", out.len());

    // WPS template (no layout objects, cluster table absent): 4 pages.
    let hello = std::fs::read("samples/hello.ppt").unwrap();
    let out = write::write_from_slide_template(&hello, &deck(4)).unwrap();
    std::fs::write("samples/scratch9-hello4.ppt", &out).unwrap();
    println!("samples/scratch9-hello4.ppt {} bytes", out.len());

    // Read back what we just wrote: text per slide, so a mis-clone shows up here
    // before PowerPoint has to say it.
    let doc = office_ppt::open("samples/scratch9-5.ppt").unwrap();
    for (i, s) in doc.slides.iter().enumerate() {
        let texts: Vec<&str> = s.texts.iter().map(|t| t.text.trim_end()).collect();
        println!("  slide {}: {:?}", i + 1, texts);
    }
}
