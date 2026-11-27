//! OfficeArt BLIP extraction.
//!
//! Both `.ppt` and `.doc` store images as OfficeArt *BLIP* records
//! (`OfficeArtBlipPNG` 0xF01D, `…JPEG` 0xF01C/0xF02A, `…EMF` 0xF018, …), either
//! bare or wrapped in an `OfficeArtFBSE` (0xF007) inside an
//! `OfficeArtBStoreContainer` (0xF001). This crate holds the part that both
//! readers need, so the container walkers stay format specific.
//!
//! Two decisions are baked in, both taken from real files rather than the spec:
//!
//! * the image type is **sniffed from the magic bytes**, because the record
//!   type is not reliable — WPS Office writes a PNG into a record typed
//!   `OfficeArtBlipTIFF` (0xF01E);
//! * a BLIP payload starts with a 16-byte `rgbUid`, sometimes followed by a
//!   second one and a one-byte tag, so the image is located by scanning the
//!   first few dozen bytes for a known image header.

/// Image formats that can appear inside a BLIP record.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PictureKind {
    Png,
    Jpeg,
    Gif,
    Bmp,
    Tiff,
    Emf,
    Wmf,
    Unknown,
}

impl PictureKind {
    /// File extension for this kind.
    pub fn extension(&self) -> &'static str {
        match self {
            PictureKind::Png => "png",
            PictureKind::Jpeg => "jpg",
            PictureKind::Gif => "gif",
            PictureKind::Bmp => "bmp",
            PictureKind::Tiff => "tif",
            PictureKind::Emf => "emf",
            PictureKind::Wmf => "wmf",
            PictureKind::Unknown => "bin",
        }
    }

    /// Identify an image by its magic bytes.
    pub fn sniff(data: &[u8]) -> PictureKind {
        if data.starts_with(&[0x89, b'P', b'N', b'G']) {
            PictureKind::Png
        } else if data.starts_with(&[0xFF, 0xD8, 0xFF]) {
            PictureKind::Jpeg
        } else if data.starts_with(b"GIF8") {
            PictureKind::Gif
        } else if data.starts_with(b"BM") {
            PictureKind::Bmp
        } else if data.starts_with(&[0x49, 0x49, 0x2A, 0x00])
            || data.starts_with(&[0x4D, 0x4D, 0x00, 0x2A])
        {
            PictureKind::Tiff
        } else if data.len() > 44 && data[40..44] == [0x20, 0x45, 0x4D, 0x46] {
            PictureKind::Emf
        } else if data.starts_with(&[0xD7, 0xCD, 0xC6, 0x9A]) {
            PictureKind::Wmf
        } else {
            PictureKind::Unknown
        }
    }
}

/// `OfficeArtFBSE` — the record that wraps a BLIP in a store container.
pub const RT_FBSE: u16 = 0xF007;

/// `true` for the BLIP record range `0xF018..=0xF02A`.
pub fn is_blip(kind: u16) -> bool {
    (0xF018..=0xF02A).contains(&kind)
}

/// Offset of the first recognisable image header inside a BLIP payload.
pub fn find_image_start(payload: &[u8]) -> Option<usize> {
    let limit = payload.len().min(64);
    (0..limit).find(|at| PictureKind::sniff(&payload[*at..]) != PictureKind::Unknown)
}

/// One extracted image.
#[derive(Debug, Clone, Copy)]
pub struct Blip<'a> {
    pub kind: PictureKind,
    /// The image bytes exactly as stored (a complete PNG/JPEG/… file).
    pub data: &'a [u8],
}

impl<'a> Blip<'a> {
    pub fn len(&self) -> usize {
        self.data.len()
    }

    pub fn is_empty(&self) -> bool {
        self.data.is_empty()
    }
}

/// Image bytes from a BLIP record payload (the bytes *after* the 8-byte record
/// header).
pub fn from_blip(payload: &[u8]) -> Option<Blip<'_>> {
    let at = find_image_start(payload)?;
    let data = &payload[at..];
    Some(Blip {
        kind: PictureKind::sniff(data),
        data,
    })
}

/// Image bytes from an `OfficeArtFBSE` payload: `btWin32(1) btMacOS(1)
/// rgbUid(16) tag(2) size(4) cRef(4) foDelay(4) unused1(1) cbName(1)
/// unused2(1) unused3(1) name(cbName*2)` and then a BLIP record.
pub fn from_bse(payload: &[u8]) -> Option<Blip<'_>> {
    // The BLIP follows the 36-byte header (rgbUid, tag, size, cRef, foDelay,
    // unused, cbName). Word's own files are not consistent about the last two
    // bytes, so look for the BLIP record instead of trusting `cbName`.
    let mut at = 34usize;
    while at + 8 <= payload.len() {
        let kind = u16::from_le_bytes([payload[at + 2], payload[at + 3]]);
        let len = u32::from_le_bytes([
            payload[at + 4],
            payload[at + 5],
            payload[at + 6],
            payload[at + 7],
        ]) as usize;
        if is_blip(kind) {
            return from_blip(&payload[at..]);
        }
        if len == 0 || at + 8 + len > payload.len() {
            at += 1;
        } else {
            at += 8 + len;
        }
    }
    None
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn kinds_are_sniffed_from_magic_bytes() {
        assert_eq!(PictureKind::sniff(&[0x89, b'P', b'N', b'G', 0x0D]), PictureKind::Png);
        assert_eq!(PictureKind::sniff(&[0xFF, 0xD8, 0xFF, 0xE0]), PictureKind::Jpeg);
        assert_eq!(PictureKind::sniff(b"GIF89a"), PictureKind::Gif);
        assert_eq!(PictureKind::sniff(b"BM6\0\0\0"), PictureKind::Bmp);
        assert_eq!(PictureKind::sniff(&[0xD7, 0xCD, 0xC6, 0x9A]), PictureKind::Wmf);
        assert_eq!(PictureKind::sniff(b"nonsense"), PictureKind::Unknown);
    }

    #[test]
    fn the_uid_and_tag_are_skipped_before_the_image() {
        let mut payload = vec![0xABu8; 16]; // rgbUid
        payload.push(0xFF); // tag
        payload.extend_from_slice(&[0x89, b'P', b'N', b'G', 0x0D, 0x0A]);
        let blip = from_blip(&payload).unwrap();
        assert_eq!(blip.kind, PictureKind::Png);
        assert_eq!(blip.data, &[0x89, b'P', b'N', b'G', 0x0D, 0x0A]);
    }

    #[test]
    fn fbse_wrapper_is_unwrapped() {
        // FBSE header with cbName = 2 (a 4-byte UTF-16 name) then a BLIP record.
        let mut bse = vec![6u8, 6];
        bse.extend_from_slice(&[0u8; 16]); // rgbUid
        bse.extend_from_slice(&[0xFF, 0x00]); // tag
        bse.extend_from_slice(&100u32.to_le_bytes()); // size
        bse.extend_from_slice(&1u32.to_le_bytes()); // cRef
        bse.extend_from_slice(&0u32.to_le_bytes()); // foDelay
        bse.extend_from_slice(&[0, 2, 0, 0]); // unused1, cbName=2, unused2, unused3
        assert_eq!(bse.len(), 36);
        bse.extend_from_slice(&[b'a', 0, b'b', 0]); // name
        // BLIP record: header + payload (uid + tag + PNG magic)
        let mut body = vec![0u8; 16];
        body.push(0xFF);
        body.extend_from_slice(&[0x89, b'P', b'N', b'G', 0x0D, 0x0A, 0x1A, 0x0A]);
        bse.extend_from_slice(&0x006Eu16.to_le_bytes()); // ver/instance
        bse.extend_from_slice(&0xF01Eu16.to_le_bytes()); // type: BlipTIFF (lie)
        bse.extend_from_slice(&(body.len() as u32).to_le_bytes());
        bse.extend_from_slice(&body);

        let blip = from_bse(&bse).unwrap();
        assert_eq!(blip.kind, PictureKind::Png, "magic wins over the record type");
        assert_eq!(blip.data.len(), 8);
    }

    #[test]
    fn a_truncated_payload_is_rejected_instead_of_panicking() {
        assert!(from_bse(&[0u8; 10]).is_none());
        assert!(from_blip(&[0u8; 4]).is_none());
    }
}
