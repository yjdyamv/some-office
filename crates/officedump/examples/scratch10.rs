//! Decks whose slides are generated from scratch — no cloned shape tree.
//!
//! ```text
//! cargo run -q --example scratch10
//! ```

use office_ppt::write::{self, SlideSpec};

fn main() {
    let text = vec![
        SlideSpec::new("Generated slide one")
            .with_body(&["the shape tree is built, not cloned", "second paragraph"]),
        SlideSpec::new("Generated slide two").with_body(&["two slides, four shapes"]),
    ];
    let file = write::write_presentation(&text).expect("generate");
    std::fs::write("samples/scratch10-text.ppt", &file).expect("write");
    println!("samples/scratch10-text.ppt {} bytes", file.len());

    let pixel = std::fs::read("samples/pixel.png").expect("pixel.png");
    let pics = vec![SlideSpec::new("Generated picture")
        .with_body(&["a generated picture shape"])
        .with_picture(pixel.clone(), 120, 120)];
    let file = write::write_presentation(&pics).expect("generate");
    std::fs::write("samples/scratch10-pic.ppt", &file).expect("write");
    println!("samples/scratch10-pic.ppt {} bytes", file.len());

    let only = vec![SlideSpec::new("").with_picture(pixel, 120, 120)];
    let file = write::write_presentation(&only).expect("generate");
    std::fs::write("samples/scratch10-piconly.ppt", &file).expect("write");
    println!("samples/scratch10-piconly.ppt {} bytes", file.len());
}
