//! `write_from_slide_template` on arbitrary template files (the embedded-template
//! path is covered by `scratch7`). Both share `build_from_stream`, which clones
//! the template's first slide and keeps every layout, so the output has the
//! shape tree PowerPoint accepts.
use office_ppt::write::{self, SlideSpec};

fn main() {
    let spec = SlideSpec { title: "Generated title".into(), body: vec!["body line one".into()], pictures: vec![], notes: String::new() };
    for path in ["samples/ms/minimal.ppt", "samples/hello.ppt", "samples/void-ms.ppt"] {
        let Ok(tpl) = std::fs::read(path) else { continue };
        let out = write::write_from_slide_template(&tpl, std::slice::from_ref(&spec)).unwrap();
        let stem = std::path::Path::new(path).file_stem().unwrap().to_string_lossy();
        let out_path = format!("samples/scratch6-{stem}.ppt");
        std::fs::write(&out_path, &out).unwrap();
        println!("{out_path} {} bytes", out.len());
    }
}
