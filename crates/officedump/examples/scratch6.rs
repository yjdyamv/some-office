use office_ppt::write::{self, SlideSpec};

fn main() {
    let tpl = std::fs::read("samples/ms/minimal.ppt").unwrap();
    let spec = SlideSpec { title: "Generated title".into(), body: vec!["body line one".into()], pictures: vec![] };
    for keep in [true, false] {
        let out = write::write_from_template_mixed(&tpl, &[spec.clone()], keep).unwrap();
        let path = format!("samples/scratch6-keep{keep}.ppt");
        std::fs::write(&path, &out).unwrap();
        println!("{path} {} bytes", out.len());
    }
}
