//! Print the print setup of `.xls` files.

use std::env;

fn main() {
    let args: Vec<String> = env::args().skip(1).collect();
    if args.is_empty() {
        eprintln!("usage: xlsprint FILE.xls...");
        std::process::exit(2);
    }
    for path in args {
        match office_xls::open(&path) {
            Ok(wb) => {
                for sheet in &wb.sheets {
                    if let Some(area) = sheet.print_area {
                        println!(
                            "{path} [{}]: print area rows {}..{}, columns {}..{}",
                            sheet.name, area.first.0, area.last.0, area.first.1, area.last.1
                        );
                    }
                    match &sheet.print {
                        Some(p) => println!(
                            "{path} [{}]: paper {}, {}, margins L{} R{} T{} B{} in, header {:?}, footer {:?}",
                            sheet.name,
                            p.paper_size,
                            if p.landscape { "landscape" } else { "portrait" },
                            p.left_margin,
                            p.right_margin,
                            p.top_margin,
                            p.bottom_margin,
                            p.header,
                            p.footer,
                        ),
                        None => println!("{path} [{}]: no print setup", sheet.name),
                    }
                }
            }
            Err(e) => println!("{path}: {e}"),
        }
    }
}
