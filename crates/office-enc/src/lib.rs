//! # office-enc — the two encodings Office binary formats actually use
//!
//! * **UTF-16LE** — Word "uncompressed" pieces, `TextCharsAtom` in PowerPoint,
//!   BIFF8 strings with the high-byte flag set.
//! * **CP1252** — Word "compressed" pieces (one byte per character),
//!   `TextBytesAtom`, BIFF8 strings with the high-byte flag clear, BIFF5 text.
//!
//! CP1252 is Latin-1 except for 0x80..=0x9F, which holds punctuation such as
//! the curly quotes Word writes when "smart quotes" are on.

/// Decode CP1252. Invalid/unassigned bytes map to the matching C1 control
/// character (as Windows does), never to U+FFFD, so byte offsets stay 1:1.
pub fn decode_cp1252(bytes: &[u8]) -> String {
    const HIGH: [char; 32] = [
        '\u{20AC}', '\u{0081}', '\u{201A}', '\u{0192}', '\u{201E}', '\u{2026}', '\u{2020}', '\u{2021}',
        '\u{02C6}', '\u{2030}', '\u{0160}', '\u{2039}', '\u{0152}', '\u{008D}', '\u{017D}', '\u{008F}',
        '\u{0090}', '\u{2018}', '\u{2019}', '\u{201C}', '\u{201D}', '\u{2022}', '\u{2013}', '\u{2014}',
        '\u{02DC}', '\u{2122}', '\u{0161}', '\u{203A}', '\u{0153}', '\u{009D}', '\u{017E}', '\u{0178}',
    ];
    bytes
        .iter()
        .map(|&b| {
            if (0x80..=0x9F).contains(&b) {
                HIGH[(b - 0x80) as usize]
            } else {
                b as char
            }
        })
        .collect()
}

/// Decode UTF-16LE, replacing unpaired surrogates with U+FFFD.
pub fn decode_utf16le(bytes: &[u8]) -> String {
    let units: Vec<u16> = bytes
        .chunks_exact(2)
        .map(|c| u16::from_le_bytes([c[0], c[1]]))
        .collect();
    String::from_utf16_lossy(&units)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn cp1252_quotes_and_accent() {
        assert_eq!(decode_cp1252(&[0x93, b'h', b'i', 0x94]), "\u{201C}hi\u{201D}");
        assert_eq!(decode_cp1252(&[0xE9]), "é");
        assert_eq!(decode_cp1252(b"plain"), "plain");
    }

    #[test]
    fn utf16_cjk() {
        assert_eq!(decode_utf16le(&[0x2D, 0x4E, 0x87, 0x65]), "中文");
    }
}
