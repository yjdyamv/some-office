//! RC4 — the stream cipher behind both Office 97-2003 encryption variants.
//!
//! Office never uses a single long RC4 stream: it re-keys every *block*
//! (512 bytes for `.doc`, 1024 for `.xls`, per persist object for `.ppt`), so
//! the cipher is used as a keyed keystream generator that restarts on demand.

/// A resumable RC4 keystream state.
pub struct Rc4 {
    s: [u8; 256],
    i: u8,
    j: u8,
}

impl Rc4 {
    /// Key-schedule `key` (1..=256 bytes).
    pub fn new(key: &[u8]) -> Self {
        let mut s = [0u8; 256];
        for (i, b) in s.iter_mut().enumerate() {
            *b = i as u8;
        }
        let mut j = 0u8;
        for i in 0..256 {
            j = j
                .wrapping_add(s[i])
                .wrapping_add(key[i % key.len().max(1)]);
            s.swap(i, j as usize);
        }
        Self { s, i: 0, j: 0 }
    }

    /// XOR `data` with the next `data.len()` keystream bytes, in place.
    pub fn apply(&mut self, data: &mut [u8]) {
        for byte in data {
            self.i = self.i.wrapping_add(1);
            self.j = self.j.wrapping_add(self.s[self.i as usize]);
            self.s.swap(self.i as usize, self.j as usize);
            let k = self.s[(self.s[self.i as usize].wrapping_add(self.s[self.j as usize])) as usize];
            *byte ^= k;
        }
    }
}

/// One-shot RC4: XOR `data` with the keystream derived from `key`.
pub fn apply(key: &[u8], data: &[u8]) -> Vec<u8> {
    let mut out = data.to_vec();
    Rc4::new(key).apply(&mut out);
    out
}

#[cfg(test)]
mod tests {
    use super::{apply, Rc4};

    #[test]
    fn rfc6229_style_known_answer() {
        // Classic test vector: key "Key", plaintext "Plaintext".
        let ct = apply(b"Key", b"Plaintext");
        assert_eq!(
            ct.iter().map(|b| format!("{b:02X}")).collect::<String>(),
            "BBF316E8D940AF0AD3"
        );
        // And the inverse direction is the same operation.
        assert_eq!(apply(b"Key", &ct), b"Plaintext".to_vec());
    }

    #[test]
    fn resuming_matches_one_shot() {
        let key = [0x11u8; 16];
        let data: Vec<u8> = (0..1000u32).map(|i| (i % 251) as u8).collect();
        let one_shot = apply(&key, &data);

        let mut rc4 = Rc4::new(&key);
        let mut split = data.clone();
        rc4.apply(&mut split[..300]);
        rc4.apply(&mut split[300..]);
        assert_eq!(split, one_shot);
    }

    #[test]
    fn rekeying_changes_the_stream() {
        // This is why Office re-keys per block instead of using one stream.
        let data = [0u8; 8];
        let a = apply(b"key-0000", &data);
        let b = apply(b"key-0001", &data);
        assert_ne!(a, b);
    }
}
