//! Print the page setup of `.doc` files, read through the library.

use std::env;

fn main() {
    let args: Vec<String> = env::args().skip(1).collect();
    if args.is_empty() {
        eprintln!("usage: pages FILE.doc...");
        std::process::exit(2);
    }
    for path in args {
        match office_doc::open(&path) {
            Ok(doc) => match doc.page {
                Some(page) => println!(
                    "{path}: {}x{} twips ({:.1}x{:.1} pt), margins L{} R{} T{} B{}",
                    page.width,
                    page.height,
                    page.width as f64 / 20.0,
                    page.height as f64 / 20.0,
                    page.left,
                    page.right,
                    page.top.map(|t| t.to_string()).unwrap_or_else(|| "-".into()),
                    page.bottom.map(|b| b.to_string()).unwrap_or_else(|| "-".into()),
                ),
                None => println!("{path}: no section properties"),
            },
            Err(e) => println!("{path}: {e}"),
        }
    }
}
