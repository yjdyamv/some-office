use office_ppt::write::{self, SlideSpec};

fn main() {
    let spec = SlideSpec { title: "Generated title".into(), body: vec!["body line one".into()], pictures: vec![] };
    let out = write::write_presentation(&[spec]).unwrap();
    std::fs::write("samples/scratch7-pres.ppt", &out).unwrap();
    println!("samples/scratch7-pres.ppt {} bytes", out.len());
}
