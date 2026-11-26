//! `officedump` — a small command line front end for the three readers.
//!
//! ```text
//! officedump info <file>          # container + format summary
//! officedump tree <file>          # CFB directory tree
//! officedump text <file> [--json] # text (doc/ppt) or cell grid (xls)
//! officedump html <file>          # styled HTML (doc runs, xls grid, ppt slides)
//! officedump media <file> <dir>   # extract embedded images (ppt)
//! officedump stories <file>       # .doc: main text, footnotes, headers
//! officedump styles <file>        # .doc: the style sheet (istd -> name)
//!
//! Add `--password <pw>` to any command to open a protected document.
//! ```

use std::process::ExitCode;

use office_cfb::{Cfb, EntryKind};

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Format {
    Doc,
    Xls,
    Ppt,
    Unknown,
}

fn detect(cfb: &Cfb) -> Format {
    if cfb.stream("WordDocument").is_ok() {
        Format::Doc
    } else if cfb.stream("Workbook").is_ok() || cfb.stream("Book").is_ok() {
        Format::Xls
    } else if cfb.stream("PowerPoint Document").is_ok() {
        Format::Ppt
    } else {
        Format::Unknown
    }
}

fn main() -> ExitCode {
    let args: Vec<String> = std::env::args().skip(1).collect();
    if args.len() < 2 {
        eprintln!("{USAGE}");
        return ExitCode::from(2);
    }
    let cmd = args[0].as_str();
    let path = &args[1];
    let json = args[2..].iter().any(|a| a == "--json");
    let out_dir = args
        .get(2)
        .filter(|a| !a.starts_with("--"))
        .cloned()
        .filter(|_| cmd == "media");
    // `--password <pw>` (or `-p <pw>`) anywhere after the file name.
    let password = args
        .iter()
        .position(|a| a == "--password" || a == "-p")
        .and_then(|i| args.get(i + 1))
        .cloned();

    match run(cmd, path, json, out_dir.as_deref(), password.as_deref()) {
        Ok(()) => ExitCode::SUCCESS,
        Err(e) => {
            eprintln!("officedump: {e}");
            ExitCode::FAILURE
        }
    }
}

const USAGE: &str = "\
usage: officedump <command> <file> [--json] [--password <pw>]

commands:
  info   container version, sector size, streams, detected format
  tree   CFB directory tree
  text   document text (doc), cell grid (xls), slide text (ppt)
  html   styled HTML: doc character/paragraph runs, xls grid, ppt slides
  media  extract embedded images (ppt Pictures stream, doc Data stream)
  stories  .doc: split the text stream into main / footnote / header stories
  styles   .doc: print the style sheet (istd -> style name)
  raw    dump the record/byte layout header of the main stream

options:
  --json             machine readable output (text, tree)
  --password, -p     password for RC4/XOR protected files";

fn run(
    cmd: &str,
    path: &str,
    json: bool,
    out_dir: Option<&str>,
    password: Option<&str>,
) -> Result<(), Box<dyn std::error::Error>> {
    let bytes = std::fs::read(path)?;
    let cfb = Cfb::parse(bytes.clone())?;
    match cmd {
        "info" => info(&cfb, &bytes, path, password),
        "tree" => tree(&cfb, json),
        "text" => text(&cfb, &bytes, json, password),
        "html" => html(&cfb, &bytes, password),
        "media" => media(&cfb, &bytes, out_dir, password),
        "stories" => stories(&cfb, &bytes, password),
        "styles" => styles(&cfb, &bytes, password),
        "raw" => raw(&cfb),
        other => Err(format!("unknown command {other:?}\n\n{USAGE}").into()),
    }
}

/// Open a `.doc`, decrypting it when a password is given.
fn open_doc(
    bytes: &[u8],
    password: Option<&str>,
) -> Result<office_doc::Doc, Box<dyn std::error::Error>> {
    Ok(match password {
        Some(pw) => office_doc::parse_with_password(bytes, pw)?,
        None => office_doc::parse(bytes)?,
    })
}

/// Open a `.xls`, decrypting it when a password is given.
fn open_xls(
    bytes: &[u8],
    password: Option<&str>,
) -> Result<office_xls::Workbook, Box<dyn std::error::Error>> {
    Ok(match password {
        Some(pw) => office_xls::parse_with_password(bytes, pw)?,
        None => office_xls::parse(bytes)?,
    })
}

/// Open a `.ppt`, decrypting it when a password is given.
fn open_ppt(
    stream: &[u8],
    user: Option<&[u8]>,
    password: Option<&str>,
) -> Result<office_ppt::Presentation, Box<dyn std::error::Error>> {
    Ok(match password {
        Some(pw) => office_ppt::parse_stream_with_password(stream, user, pw)?,
        None => office_ppt::parse_stream(stream, user)?,
    })
}

/// `true` when the error means "this file is encrypted and no usable password
/// was supplied", i.e. the caller should ask for one instead of failing.
fn is_locked(e: &Box<dyn std::error::Error>) -> bool {
    e.to_string().contains("encrypted") || e.to_string().contains("password")
}

fn scheme_name(scheme: &Option<office_crypto::Scheme>) -> String {
    match scheme {
        Some(s) => s.name().to_string(),
        None => "none".to_string(),
    }
}

fn info(
    cfb: &Cfb,
    bytes: &[u8],
    path: &str,
    password: Option<&str>,
) -> Result<(), Box<dyn std::error::Error>> {
    let fmt = detect(cfb);
    println!("file        {path}");
    println!(
        "container   CFB v{}, sector {} B, mini sector {} B, cutoff {} B",
        cfb.version(),
        cfb.sector_size(),
        cfb.mini_sector_size(),
        cfb.mini_stream_cutoff()
    );
    println!("streams     {}", cfb.entries().iter().filter(|e| e.kind == EntryKind::Stream).count());
    println!("format      {fmt:?}");
    match fmt {
        Format::Doc => {
            let doc = match open_doc(bytes, password) {
                Ok(doc) => doc,
                Err(e) if is_locked(&e) => {
                    println!("encryption  protected (pass --password <pw>)");
                    return Ok(());
                }
                Err(e) => return Err(e),
            };
            println!("encryption  {}", scheme_name(&doc.encryption));
            println!("nFib        0x{:04X}", doc.fib.n_fib);
            println!("table       {}", doc.fib.table_stream);
            println!("pieces      {}", doc.pieces.len());
            println!("characters  {} (ccpText)", doc.fib.ccp_text);
            println!("text bytes  {} .. {}", doc.fib.fc_min, doc.fib.fc_mac);
            println!(
                "styles      {} of {} defined",
                doc.styles.iter().filter(|s| s.name.is_some()).count(),
                doc.styles.len()
            );
            println!("pictures    {}", doc.pictures().len());
        }
        Format::Xls => {
            let wb = match open_xls(bytes, password) {
                Ok(wb) => wb,
                Err(e) if is_locked(&e) => {
                    println!("encryption  protected (pass --password <pw>)");
                    return Ok(());
                }
                Err(e) => return Err(e),
            };
            println!("encryption  {}", scheme_name(&wb.encryption));
            println!("biff        0x{:04X}", wb.biff_version);
            println!("codepage    0x{:04X}", wb.codepage);
            println!("date system  {}", if wb.date_1904 { "1904" } else { "1900" });
            println!("strings      {} unique", wb.sst.len());
            for s in &wb.sheets {
                println!("sheet       {:<16} {:?} @ {} ({} cells)", s.name, s.kind, s.offset, s.cells.len());
            }
        }
        Format::Ppt => {
            let user = cfb.stream("Current User").ok();
            let stream = cfb.stream("PowerPoint Document")?;
            let deck = match open_ppt(&stream, user.as_deref(), password) {
                Ok(deck) => deck,
                Err(e) if is_locked(&e) => {
                    println!("encryption  protected (pass --password <pw>)");
                    return Ok(());
                }
                Err(e) => return Err(e),
            };
            println!("encryption  {}", scheme_name(&deck.encryption));
            let pics = cfb.stream("Pictures").map(|p| office_ppt::pictures(&p).len()).unwrap_or(0);
            println!("slides      {}", deck.slides.len());
            println!("pictures    {pics}");
            let noted = deck.slides.iter().filter(|s| !s.notes.is_empty()).count();
            println!("notes       {noted} of {} slides", deck.slides.len());
            if !deck.master_styles.is_empty() {
                let mut parts: Vec<String> = Vec::new();
                for (kind, levels) in [
                    (office_ppt::TextKind::Title, "title"),
                    (office_ppt::TextKind::Body, "body"),
                    (office_ppt::TextKind::Other, "other"),
                ] {
                    if let Some(l) = deck.master_styles.get(&kind) {
                        let size = l.first().and_then(|f| f.size).map(|s| format!("{s}pt"));
                        parts.push(match size {
                            Some(s) => format!("{levels_kind} {n} levels, {s}", levels_kind = levels, n = l.len(), s = s),
                            None => format!("{} {} levels", levels, l.len()),
                        });
                    }
                }
                println!("master      {}", parts.join("; "));
            }
            println!("slide order {}", if deck.ordered_by_persist_directory { "persist directory" } else { "stream order (fallback)" });
            for s in &deck.slides {
                println!("slide {:<6} offset {:<8} persist {:?} runs {} notes {}",
                    s.index + 1, s.offset, s.persist_id, s.texts.len(), s.notes.len());
            }
        }
        Format::Unknown => println!("(no WordDocument / Workbook / PowerPoint Document stream)"),
    }
    Ok(())
}

fn tree(cfb: &Cfb, json: bool) -> Result<(), Box<dyn std::error::Error>> {
    if json {
        let mut out = String::from("[");
        for (i, e) in cfb.entries().iter().enumerate() {
            if i > 0 {
                out.push(',');
            }
            out.push_str(&format!(
                "{{\"id\":{},\"name\":{},\"kind\":\"{:?}\",\"size\":{}}}",
                e.id,
                json_string(&e.name),
                e.kind,
                e.size
            ));
        }
        out.push(']');
        println!("{out}");
        return Ok(());
    }
    fn walk(cfb: &Cfb, id: u32, depth: usize) {
        let Some(e) = cfb.entries().iter().find(|e| e.id == id) else {
            return;
        };
        let pad = "  ".repeat(depth);
        match e.kind {
            EntryKind::Stream => println!("{pad}{} ({} B)", e.name, e.size),
            EntryKind::Root => {
                println!("{pad}/");
                for c in cfb.children(id) {
                    walk(cfb, c.id, depth + 1);
                }
            }
            EntryKind::Storage => {
                println!("{pad}{}/", e.name);
                for c in cfb.children(id) {
                    walk(cfb, c.id, depth + 1);
                }
            }
            EntryKind::Empty => {}
        }
    }
    walk(cfb, 0, 0);
    Ok(())
}

fn text(
    cfb: &Cfb,
    bytes: &[u8],
    json: bool,
    password: Option<&str>,
) -> Result<(), Box<dyn std::error::Error>> {
    match detect(cfb) {
        Format::Doc => {
            let doc = open_doc(bytes, password)?;
            if json {
                let paras: Vec<String> = doc.paragraphs().iter().map(|p| json_string(p)).collect();
                println!("{{\"format\":\"doc\",\"paragraphs\":[{}]}}", paras.join(","));
            } else {
                print!("{}", doc.plain_text());
            }
        }
        Format::Xls => {
            let wb = open_xls(bytes, password)?;
            for sheet in &wb.sheets {
                if json {
                    let cells: Vec<String> = sheet
                        .cells
                        .iter()
                        .map(|c| {
                            format!(
                                "{{\"row\":{},\"col\":{},\"value\":{}}}",
                                c.row,
                                c.col,
                                json_string(&wb.cell_text(c))
                            )
                        })
                        .collect();
                    println!(
                        "{{\"format\":\"xls\",\"sheet\":{},\"cells\":[{}]}}",
                        json_string(&sheet.name),
                        cells.join(",")
                    );
                } else {
                    println!("== {} ==", sheet.name);
                    for row in sheet.rows() {
                        // cell_text() applies the cell's number format, so dates
                        // and percentages read the way Excel shows them.
                        let cols: Vec<String> = row.iter().map(|c| wb.cell_text(c)).collect();
                        let label = row.first().map(|c| c.row + 1).unwrap_or(0);
                        println!("{label}	| {}", cols.join("	"));
                    }
                }
            }
        }
        Format::Ppt => {
            let user = cfb.stream("Current User").ok();
            let stream = cfb.stream("PowerPoint Document")?;
            let deck = open_ppt(&stream, user.as_deref(), password)?;
            if json {
                let slides: Vec<String> = deck
                    .slides
                    .iter()
                    .map(|s| {
                        let runs: Vec<String> = s
                            .texts
                            .iter()
                            .map(|t| {
                                format!(
                                    "{{\"kind\":\"{:?}\",\"text\":{}}}",
                                    t.kind,
                                    json_string(&t.text)
                                )
                            })
                            .collect();
                        format!("{{\"index\":{},\"runs\":[{}]}}", s.index, runs.join(","))
                    })
                    .collect();
                println!("{{\"format\":\"ppt\",\"slides\":[{}]}}", slides.join(","));
            } else {
                print!("{}", deck.plain_text());
            }
        }
        Format::Unknown => return Err("unrecognised Office binary format".into()),
    }
    Ok(())
}


/// Escape text for HTML.
fn html_escape(s: &str) -> String {
    let mut out = String::with_capacity(s.len());
    for ch in s.chars() {
        match ch {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '>' => out.push_str("&gt;"),
            '"' => out.push_str("&quot;"),
            c if (c as u32) < 0x20 => out.push(' '),
            c => out.push(c),
        }
    }
    out
}

/// `#RRGGBB` for a `sprmCCv` colour value (stored as 0x00BBGGRR).
fn css_colour(v: u32) -> String {
    format!("#{:02X}{:02X}{:02X}", v & 0xFF, (v >> 8) & 0xFF, (v >> 16) & 0xFF)
}

fn doc_html(doc: &office_doc::Doc) -> String {
    let mut out = String::from(
        "<!doctype html>\n<meta charset=\"utf-8\">\n<style>body{font-family:serif;max-width:46em;margin:2em auto;line-height:1.5}td{padding:.15em .6em;border:1px solid #bbb;vertical-align:top}</style>\n",
    );
    let paras = doc.paragraphs_rich();
    let mut i = 0usize;
    while i < paras.len() {
        // A run of `sprmPFInTable` paragraphs is a table: each one is a cell,
        // and the paragraph whose properties carry `sprmPFTtp` closes the row
        // (Word writes one extra, empty cell mark at the end of every row).
        if paras[i].props.in_table == Some(true) {
            out.push_str("<table>\n");
            // The row-end paragraph carries `TDefTable`: consecutive values are
            // cell edges in twips (20 per point), so the column widths are the
            // differences between them.
            if let Some(w) = paras[i..]
                .iter()
                .find_map(|p| (!p.props.col_widths.is_empty()).then_some(&p.props.col_widths))
            {
                out.push_str("<colgroup>");
                for pair in w.windows(2) {
                    let pt = (pair[1] - pair[0]) as f32 / 20.0;
                    out.push_str(&format!("<col style=\"width:{pt}pt\">"));
                }
                out.push_str("</colgroup>\n");
            }
            while i < paras.len() && paras[i].props.in_table == Some(true) {
                out.push_str("<tr>");
                loop {
                    // The paragraph whose properties carry `sprmPFTtp` is the
                    // row terminator, not a cell: it only holds the `TDefTable`
                    // that describes the columns.
                    if paras[i].props.ttp == Some(true) {
                        i += 1;
                        break;
                    }
                    out.push_str("<td>");
                    out.push_str(&para_inner(&paras[i]));
                    out.push_str("</td>");
                    i += 1;
                    if i >= paras.len() || paras[i].props.in_table != Some(true) {
                        break;
                    }
                }
                out.push_str("</tr>\n");
            }
            out.push_str("</table>\n");
            continue;
        }
        let style = para_style(&paras[i]);
        if style.is_empty() {
            out.push_str("<p>");
        } else {
            out.push_str(&format!("<p style=\"{style}\">"));
        }
        out.push_str(&para_inner(&paras[i]));
        out.push_str("</p>\n");
        i += 1;
    }
    out
}

/// The CSS a paragraph's own properties call for, or an empty string.
fn para_style(para: &office_doc::fmt::Paragraph) -> String {
    let mut style = String::new();
    match para.props.justify {
        Some(1) => style.push_str("text-align:center;"),
        Some(2) => style.push_str("text-align:right;"),
        Some(3) => style.push_str("text-align:justify;"),
        _ => {}
    }
    if let Some(d) = para.props.dxa_left {
        style.push_str(&format!("margin-left:{}pt;", d as f32 / 20.0));
    }
    if let Some(d) = para.props.dxa_left_first {
        style.push_str(&format!("text-indent:{}pt;", d as f32 / 20.0));
    }
    if let Some(d) = para.props.dya_before {
        style.push_str(&format!("margin-top:{}pt;", d as f32 / 20.0));
    }
    if let Some(d) = para.props.dya_after {
        style.push_str(&format!("margin-bottom:{}pt;", d as f32 / 20.0));
    }
    style
}

/// The inline HTML of one paragraph: its text with every character run
/// wrapped in the tags its properties call for.
fn para_inner(para: &office_doc::fmt::Paragraph) -> String {
    let mut out = String::new();
    let chars: Vec<char> = para.text.chars().collect();
    if para.runs.is_empty() {
        out.push_str(&html_escape(&para.text));
    }
    for run in &para.runs {
        let text: String = chars[run.start..run.end].iter().collect();
        let p = &run.props;
        let mut css = String::new();
        if let Some(c) = p.color {
            css.push_str(&format!("color:{};", css_colour(c)));
        }
        if let Some(h) = p.size_half_points {
            css.push_str(&format!("font-size:{}pt;", h as f32 / 2.0));
        }
        if p.caps == Some(true) {
            css.push_str("text-transform:uppercase;");
        }
        if p.small_caps == Some(true) {
            css.push_str("font-variant:small-caps;");
        }
        match p.iss {
            Some(1) => css.push_str("vertical-align:super;font-size:smaller;"),
            Some(2) => css.push_str("vertical-align:sub;font-size:smaller;"),
            _ => {}
        }
        let mut open = String::new();
        let mut close = String::new();
        if p.bold == Some(true) {
            open.push_str("<b>");
            close.insert_str(0, "</b>");
        }
        if p.italic == Some(true) {
            open.push_str("<i>");
            close.insert_str(0, "</i>");
        }
        if p.underline.is_some_and(|u| u != 0) {
            open.push_str("<u>");
            close.insert_str(0, "</u>");
        }
        if p.strike == Some(true) || p.double_strike == Some(true) {
            open.push_str("<s>");
            close.insert_str(0, "</s>");
        }
        out.push_str(&open);
        if css.is_empty() {
            out.push_str(&html_escape(&text));
        } else {
            out.push_str(&format!("<span style=\"{css}\">{}</span>", html_escape(&text)));
        }
        out.push_str(&close);
    }
    out
}

fn html(
    cfb: &Cfb,
    bytes: &[u8],
    password: Option<&str>,
) -> Result<(), Box<dyn std::error::Error>> {
    match detect(cfb) {
        Format::Doc => print!("{}", doc_html(&open_doc(bytes, password)?)),
        Format::Xls => {
            let wb = open_xls(bytes, password)?;
            let mut out = String::from("<!doctype html>\n<meta charset=\"utf-8\">\n");
            for sheet in &wb.sheets {
                out.push_str(&format!(
                    "<h2>{}</h2>\n<table border=1 cellspacing=0 cellpadding=3>\n",
                    html_escape(&sheet.name)
                ));
                for row in sheet.rows() {
                    out.push_str("<tr>");
                    let mut next = 0u16;
                    for cell in row {
                        while next < cell.col {
                            out.push_str("<td></td>");
                            next += 1;
                        }
                        out.push_str(&format!("<td>{}</td>", html_escape(&wb.cell_text(cell))));
                        next = cell.col + 1;
                    }
                    out.push_str("</tr>\n");
                }
                out.push_str("</table>\n");
            }
            print!("{out}");
        }
        Format::Ppt => {
            let user = cfb.stream("Current User").ok();
            let stream = cfb.stream("PowerPoint Document")?;
            let deck = open_ppt(&stream, user.as_deref(), password)?;
            let mut out = String::from("<!doctype html>\n<meta charset=\"utf-8\">\n");
            for slide in &deck.slides {
                out.push_str(&format!("<div class=\"slide\" id=\"slide{}\">\n", slide.index + 1));
                let mut i = 0;
                while i < slide.texts.len() {
                    let kind = slide.texts[i].kind;
                    let mut inner = String::new();
                    let mut align = None;
                    while i < slide.texts.len() && slide.texts[i].kind == kind {
                        let run = &slide.texts[i];
                        if align.is_none() {
                            align = run.format.align;
                        }
                        let css = run.format.css();
                        let body = html_escape(&run.text);
                        if css.is_empty() {
                            inner.push_str(&body);
                        } else {
                            inner.push_str(&format!("<span style=\"{css}\">{body}</span>"));
                        }
                        i += 1;
                    }
                    let tag = match kind {
                        office_ppt::TextKind::Title => "h2",
                        _ => "p",
                    };
                    let style = match align {
                        Some(1) => " style=\"text-align:center\"",
                        Some(2) => " style=\"text-align:right\"",
                        Some(3) => " style=\"text-align:justify\"",
                        _ => "",
                    };
                    out.push_str(&format!("<{tag}{style}>{inner}</{tag}>\n"));
                }
                out.push_str("</div>\n");
            }
            print!("{out}");
        }
        Format::Unknown => return Err("unrecognised Office binary format".into()),
    }
    Ok(())
}


/// `media`: write every image of the `Pictures` stream to `<dir>`.
fn media(
    cfb: &Cfb,
    bytes: &[u8],
    out_dir: Option<&str>,
    password: Option<&str>,
) -> Result<(), Box<dyn std::error::Error>> {
    let dir = out_dir.ok_or("media needs an output directory: officedump media <file> <dir>")?;

    // `.ppt` keeps its images in a `Pictures` stream, `.doc` keeps them in the
    // `Data` stream behind `sprmCPicLocation` runs.
    let images: Vec<(office_blip::PictureKind, Vec<u8>)> = match detect(cfb) {
        Format::Ppt => {
            let stream = cfb
                .stream("Pictures")
                .map_err(|_| "this presentation has no Pictures stream")?;
            office_ppt::pictures(&stream)
                .into_iter()
                .map(|p| (p.kind, p.data))
                .collect()
        }
        Format::Doc => open_doc(bytes, password)?
            .pictures()
            .into_iter()
            .map(|p| (p.kind, p.data))
            .collect(),
        _ => return Err("media supports .ppt (Pictures stream) and .doc (Data stream)".into()),
    };
    if images.is_empty() {
        return Err("no embedded images found".into());
    }
    std::fs::create_dir_all(dir)?;
    for (i, (kind, data)) in images.iter().enumerate() {
        let name = format!("{}/image{}.{}", dir, i + 1, kind.extension());
        std::fs::write(&name, data)?;
        println!("{name}  ({kind:?}, {} bytes)", data.len());
    }
    Ok(())
}


/// `stories`: Word keeps every story in one text stream; show them apart.
fn stories(
    cfb: &Cfb,
    bytes: &[u8],
    password: Option<&str>,
) -> Result<(), Box<dyn std::error::Error>> {
    if detect(cfb) != Format::Doc {
        return Err("only .doc has multiple text stories".into());
    }
    let doc = open_doc(bytes, password)?;
    for (kind, text) in doc.stories() {
        println!("--- {} ({} characters) ---", kind.name(), text.chars().count());
        println!("{}", text.replace('\r', "\n").trim_end());
    }
    Ok(())
}

/// List the style sheet: one line per `istd` that carries a name.
fn styles(
    cfb: &Cfb,
    bytes: &[u8],
    password: Option<&str>,
) -> Result<(), Box<dyn std::error::Error>> {
    if detect(cfb) != Format::Doc {
        return Err("only .doc carries a style sheet".into());
    }
    let doc = open_doc(bytes, password)?;
    println!("{:<5} {}", "istd", "name");
    for (istd, style) in doc.styles.iter().enumerate() {
        if let Some(name) = &style.name {
            let mut tags: Vec<String> = Vec::new();
            if let Some(j) = style.para.justify {
                tags.push(format!("justify={j}"));
            }
            if let Some(v) = style.para.dxa_left {
                tags.push(format!("indent={v}"));
            }
            if style.char.bold == Some(true) {
                tags.push("bold".into());
            }
            if let Some(v) = style.char.size_half_points {
                tags.push(format!("size={}pt", f32::from(v) / 2.0));
            }
            let suffix = if tags.is_empty() {
                String::new()
            } else {
                format!("  [{}]", tags.join(" "))
            };
            println!("{istd:<5} {name}{suffix}");
        }
    }
    Ok(())
}

/// Minimal JSON string escaping (no serde dependency anywhere in the workspace).
fn json_string(s: &str) -> String {
    let mut out = String::with_capacity(s.len() + 2);
    out.push('"');
    for ch in s.chars() {
        match ch {
            '"' => out.push_str("\\\""),
            '\\' => out.push_str("\\\\"),
            '\n' => out.push_str("\\n"),
            '\r' => out.push_str("\\r"),
            '\t' => out.push_str("\\t"),
            c if (c as u32) < 0x20 => out.push_str(&format!("\\u{:04x}", c as u32)),
            c => out.push(c),
        }
    }
    out.push('"');
    out
}

/// `raw`: show the first records of the main stream — handy while writing a
/// new parser by hand.
fn raw(cfb: &Cfb) -> Result<(), Box<dyn std::error::Error>> {
    let (name, data) = match detect(cfb) {
        Format::Doc => ("WordDocument", cfb.stream("WordDocument")?),
        Format::Xls => ("Workbook", cfb.stream("Workbook")?),
        Format::Ppt => ("PowerPoint Document", cfb.stream("PowerPoint Document")?),
        Format::Unknown => return Err("unrecognised Office binary format".into()),
    };
    println!("{name}: {} bytes", data.len());
    for (i, chunk) in data.chunks(16).take(8).enumerate() {
        let hex: Vec<String> = chunk.iter().map(|b| format!("{b:02X}")).collect();
        let ascii: String = chunk
            .iter()
            .map(|&b| if (0x20..0x7F).contains(&b) { b as char } else { '.' })
            .collect();
        println!("{:08X}  {:<47}  {}", i * 16, hex.join(" "), ascii);
    }
    Ok(())
}
