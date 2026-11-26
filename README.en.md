# some-office

Read **and write** the legacy Microsoft Office binary formats — `.doc` (Word 97-2003),
`.xls` (BIFF8 / BIFF5) and `.ppt` (PowerPoint 97-2003) — in pure Rust, with **no
third-party dependencies**.

The formats were reverse-engineered from the files themselves, from the binary that
implements them (Word, Excel, PowerPoint, `wwlib.dll`, `MSO.DLL`), and from the public
specifications (MS-CFB, MS-DOC, MS-XLS, MS-PPT, MS-OFFCRYPTO). Everything that the
specification gets wrong is listed in `docs/evidence.md` with the bytes that prove it.

The Chinese documentation is the primary one (`docs/README.md`); this file is the
English overview.

## Crates

| crate | what it does |
|---|---|
| `office-cfb` | Compound File Binary container: read **and write** |
| `office-enc` | code page 1252 / UTF-16LE helpers |
| `office-crypto` | MD5, SHA-1, RC4, XOR obfuscation (needed by the encrypted formats) |
| `office-blip` | Escher BLIP sniffing (PNG/JPEG/GIF/BMP/TIFF/EMF/WMF) |
| `office-doc` | `.doc`: FIB, pieces, CHPX/PAPX formatting, styles, tables, stories, pictures, encryption |
| `office-xls` | `.xls`: BIFF records, SST, cells, number formats, formula decompile **and compile**, charts, XOR/RC4 |
| `office-ppt` | `.ppt`: persist directory, slides, text formatting, master styles, notes, pictures, RC4 CryptoAPI |
| `officedump` | CLI: `info`, `text`, `tree`, `raw`, `html`, `media`, `styles`, `stories` |

## Quickstart

```sh
cargo test                      # 129 tests: unit + real files + randomised round-trips
cargo run --release -- info  samples/hello.doc
cargo run --release -- text  samples/hello.xls
cargo run --release -- html  samples/ms/formatted.doc > out.html
cargo run --release -- media samples/picture.doc out/     # extract embedded pictures
cargo run --release -- text  samples/encrypted.xls --password secret123
```

```rust
let bytes = std::fs::read("report.doc")?;
let doc = office_doc::parse(&bytes)?;
println!("{}", doc.plain_text());

for p in doc.paragraphs_rich() {
    println!("{:?} {:?}", p.props.justify, p.text);
}
```

Writing:

```rust
// .doc with formatting, tables, multiple pages
let doc = office_doc::write_table_document("before", &office_doc::Table::new(&[&["a", "b"]]), "after");

// .xls with formulas: the text is compiled into a real rgce token array
let cells = vec![office_xls::Cell {
    row: 0, col: 1, xf: 15,
    value: office_xls::CellValue::Formula {
        cached: office_xls::CachedValue::Number(30.0),
        rgce: Vec::new(),
        text: "A1+A2".into(),
    },
}];
let xls = office_xls::write_minimal("Sheet1", &cells);

// .ppt: clone the built-in blank deck's shape tree, replace the text
let deck = office_ppt::write::write_presentation(&[
    office_ppt::write::SlideSpec::new("Hello").with_body(&["from rust"]),
])?;
```

## What is verified, and how

`cargo test` proves the bytes are *self-consistent*. Only the real applications can
prove they are *valid*, so `tools/roundtrip.ps1` opens what the writers produce with
Word, Excel and PowerPoint and reads the content back:

```text
word   paragraphs=3 chars=57
word   formatted: [1] bold=-1 size=18 colour=255 align=1
word   table: paras=10 rows=2 cols=3   [1] H1 | H2 | H3
excel  sheet=Writer used=A1:A4
excel  formulas: [1] 10 | =A1+A2 = 30
ppt    slides=2   s1 WRITER-PPT-001 | first body line
```

Reading is validated against files written by both Microsoft Office and WPS Office
(they disagree on several details, which is where most of the interesting findings
came from) and against the Microsoft templates shipped with Office.

## Limitations

* Read-only for `.xls` charts (type/title/axes/series, no data references).
* No formula *evaluation* — formulas are decompiled to text and compiled back to tokens.
* Writers: `.doc` has no images or style-sheet names; `.xls` writes no charts and no
  per-cell number formats; `.ppt` writes no images, animations or notes, and needs a
  template (a 203 KB blank deck is compiled into the crate).
* `.ppt` master level 5 may be missing (some atoms do not decode exactly).

See `docs/README.md` § 已知限制 for the authoritative list, and `docs/evidence.md` for
the specification-vs-reality table.

## License

MIT OR Apache-2.0.
