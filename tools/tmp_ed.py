import re

# ---- 1) fmt.rs: add `spec` to CharProps, parse 0x0855 -------------------
p = "crates/office-doc/src/fmt.rs"
s = open(p, encoding="utf-8").read()
s = s.replace(
    """    /// `sprmCPicLocation`: offset of the picture inside the `Data` stream.
    pub pic_location: Option<u32>,""",
    """    /// `sprmCPicLocation`: offset of the picture inside the `Data` stream.
    pub pic_location: Option<u32>,
    /// `sprmCFSpec`: the character is special (a picture mark, a footnote
    /// reference, ...) and is not laid out as text.
    pub spec: Option<bool>,""",
    1,
)
s = s.replace(
    """            0x2A48 => props.iss = v.first().copied(),""",
    """            0x2A48 => props.iss = v.first().copied(),
            0x0855 => props.spec = v.first().map(|b| *b != 0),""",
    1,
)
open(p, "w", encoding="utf-8", newline="\n").write(s)

# ---- 2) write.rs -------------------------------------------------------
p = "crates/office-doc/src/write.rs"
s = open(p, encoding="utf-8").read()

s = s.replace(
    """use crate::fmt::{CharProps, ParaProps};""",
    """use crate::fmt::{CharProps, ParaProps};

/// A single picture, stored in the `Data` stream and referenced from a `CHPX`
/// through `sprmCPicLocation`.
///
/// The bytes are the complete image file (PNG, JPEG, ...) exactly as it should
/// be extracted again; the size is what Word uses to lay it out, in twips
/// (1/1440 inch -- a 16x16 pixel image at 96 dpi is 240x240).
#[derive(Clone, Debug)]
pub struct InlinePicture {
    pub data: Vec<u8>,
    pub width_twips: u16,
    pub height_twips: u16,
}

impl InlinePicture {
    pub fn new(data: impl Into<Vec<u8>>, width_twips: u16, height_twips: u16) -> Self {
        InlinePicture { data: data.into(), width_twips, height_twips }
    }
}

/// A `Data`-stream picture taken from a real Word file (one 16x16 PNG), used as
/// the template for the `PICF`/`OfficeArt` wrapper: Word validates that wrapper
/// far more strictly than it documents it, so the writer keeps its bytes and
/// patches only the sizes and the image itself.
const PICTURE_TEMPLATE: &[u8] = include_bytes!("../templates/picture-data.bin");

/// Offset of the `OfficeArtFBSE` record inside [`PICTURE_TEMPLATE`].
const TEMPLATE_FBSE: usize = 0xD6;""",
    1,
)

s = s.replace(
    """pub struct Run {
    pub text: String,
    pub props: CharProps,
}""",
    """pub struct Run {
    pub text: String,
    pub props: CharProps,
    /// An inline picture: the run's text becomes the picture mark (`\\x01`) and
    /// its `CHPX` gains `sprmCPicLocation` + `sprmCFSpec`.
    pub picture: Option<InlinePicture>,
}""",
    1,
)

s = s.replace(
    """    pub fn new(text: impl Into<String>) -> Self {
        Run {
            text: text.into(),
            props: CharProps::default(),
        }
    }""",
    """    pub fn new(text: impl Into<String>) -> Self {
        Run {
            text: text.into(),
            props: CharProps::default(),
            picture: None,
        }
    }

    /// An inline picture; its size is in twips (a 96-dpi pixel is 15 twips).
    pub fn picture(picture: InlinePicture) -> Self {
        Run {
            text: String::from('\\u{1}'),
            props: CharProps::default(),
            picture: Some(picture),
        }
    }""",
    1,
)

s = s.replace(
    """    pub fn styled(text: impl Into<String>, props: CharProps) -> Self {
        Run {
            text: text.into(),
            props,
        }
    }""",
    """    pub fn styled(text: impl Into<String>, props: CharProps) -> Self {
        Run {
            text: text.into(),
            props,
            picture: None,
        }
    }""",
    1,
)

s = s.replace(
    """    // ---- lay the text out and record where every run starts and ends --------
    let mut utf16: Vec<u8> = Vec::new();""",
    """    // ---- pictures: the `Data` stream is built first so each run can be given
    // the offset of its `PICF` while the text is laid out ---------------------
    let mut data_stream: Vec<u8> = Vec::new();
    let mut picture_offsets: Vec<u32> = Vec::new();
    for paragraph in paragraphs {
        for run in &paragraph.runs {
            if let Some(picture) = &run.picture {
                picture_offsets.push(data_stream.len() as u32);
                append_picture(&mut data_stream, picture);
            }
        }
    }

    // ---- lay the text out and record where every run starts and ends --------
    let mut utf16: Vec<u8> = Vec::new();
    let mut next_picture = 0usize;""",
    1,
)

s = s.replace(
    """            runs.push((start, TEXT_AT as u32 + utf16.len() as u32, run.props.clone()));""",
    """            let mut props = run.props.clone();
            if run.picture.is_some() {
                props.pic_location = Some(picture_offsets[next_picture]);
                props.spec = Some(true);
                next_picture += 1;
            }
            runs.push((start, TEXT_AT as u32 + utf16.len() as u32, props));""",
    1,
)

s = s.replace(
    """    office_cfb::write(&[
        office_cfb::OutputStream::new("WordDocument", &wd),
        office_cfb::OutputStream::new("1Table", &table),
    ])""",
    """    let mut streams = vec![
        office_cfb::OutputStream::new("WordDocument", &wd),
        office_cfb::OutputStream::new("1Table", &table),
    ];
    if !data_stream.is_empty() {
        streams.push(office_cfb::OutputStream::new("Data", &data_stream));
    }
    office_cfb::write(&streams)""",
    1,
)

s = s.replace(
    """/// Split items into formatting pages:""",
    """/// Append one picture to the `Data` stream: `PICF` + `OfficeArt` wrapper +
/// `OfficeArtFBSE` + BLIP.
///
/// The wrapper bytes come from [`PICTURE_TEMPLATE`] because Word checks more
/// than MS-DOC describes; only the sizes, the picture size in twips and the
/// image bytes are ours.
fn append_picture(out: &mut Vec<u8>, picture: &InlinePicture) {
    let start = out.len();
    let blip_len = 17 + picture.data.len(); // uid + tag + image
    let fbse_len = 36 + 8 + blip_len; // header fields + the nested BLIP record
    let total = TEMPLATE_FBSE + 8 + fbse_len;

    // `PICF` + `OfficeArt` wrapper, with the total length and the layout size
    // patched (`dxaGoal`/`dyaGoal` at 0x1C/0x1E).
    let mut picf = PICTURE_TEMPLATE[..TEMPLATE_FBSE].to_vec();
    picf[0..4].copy_from_slice(&(total as u32).to_le_bytes());
    picf[0x1C..0x1E].copy_from_slice(&picture.width_twips.to_le_bytes());
    picf[0x1E..0x20].copy_from_slice(&picture.height_twips.to_le_bytes());
    out.extend_from_slice(&picf);

    // `OfficeArtFBSE`: `len` covers its own fields plus the nested BLIP, and
    // `size` is the BLIP record length.
    let fbse = TEMPLATE_FBSE + 8;
    out.extend_from_slice(&0xF007u16.to_le_bytes());
    out.extend_from_slice(&(fbse_len as u32).to_le_bytes());
    out.extend_from_slice(&PICTURE_TEMPLATE[fbse..fbse + 18]); // ver_inst, uid, tag
    out.extend_from_slice(&((8 + blip_len) as u32).to_le_bytes()); // size
    out.extend_from_slice(&PICTURE_TEMPLATE[fbse + 22..fbse + 36]); // cRef ... unused

    // The BLIP: uid + tag byte + the image, typed by sniffing the magic bytes.
    let blip = fbse + 36;
    let kind = match office_blip::sniff(&picture.data) {
        office_blip::PictureKind::Png => 0xF01Eu16,
        office_blip::PictureKind::Jpeg => 0xF01Fu16,
        office_blip::PictureKind::Bmp => 0xF02Au16,
        _ => 0xF01Eu16,
    };
    out.extend_from_slice(&PICTURE_TEMPLATE[blip..blip + 2]); // ver_inst
    out.extend_from_slice(&kind.to_le_bytes());
    out.extend_from_slice(&(blip_len as u32).to_le_bytes());
    out.extend_from_slice(&PICTURE_TEMPLATE[blip + 8..blip + 25]); // uid + tag
    out.extend_from_slice(&picture.data);

    debug_assert_eq!(out.len() - start, total);
}

/// Split items into formatting pages:""",
    1,
)
open(p, "w", encoding="utf-8", newline="\n").write(s)
print("patched")
