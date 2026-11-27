//! XOR obfuscation (MS-OFFCRYPTO 2.3.7.1, "method 1") — the oldest Office
//! "password" scheme, still written by Office 2007+ when a `.doc`/`.xls` is
//! saved in 97-2003 format with a password and "weak encryption" is chosen.
//!
//! It is not encryption at all: a 16-byte array derived from the password is
//! XORed with the data and the result rotated right by 5 bits. The array index
//! is a function of the **absolute position in the stream**, so decrypting a
//! run of bytes requires knowing where that run starts.
//!
//! Verified against the reference implementation's output (see `tests`).

/// 15-byte filler used to pad the password to 16 bytes.
const PAD: [u8; 15] = [
    0xBB, 0xFF, 0xFF, 0xBA, 0xFF, 0xFF, 0xB9, 0x80, 0x00, 0xBE, 0x0F, 0x00, 0xBF, 0x0F, 0x00,
];

/// Per-password-length starting key (`initial_code[len - 1]`).
const INITIAL_CODE: [u16; 15] = [
    0xE1F0, 0x1D0F, 0xCC9C, 0x84C0, 0x110C, 0x0E10, 0xF1CE, 0x313E, 0x1872, 0xE139, 0xD40F, 0x84F9,
    0x280C, 0xA96A, 0x4EC3,
];

/// 15 x 7-bit mixing matrix.
const MATRIX: [u16; 105] = [
    0xAEFC, 0x4DD9, 0x9BB2, 0x2745, 0x4E8A, 0x9D14, 0x2A09, 0x7B61, 0xF6C2, 0xFDA5, 0xEB6B, 0xC6F7,
    0x9DCF, 0x2BBF, 0x4563, 0x8AC6, 0x05AD, 0x0B5A, 0x16B4, 0x2D68, 0x5AD0, 0x0375, 0x06EA, 0x0DD4,
    0x1BA8, 0x3750, 0x6EA0, 0xDD40, 0xD849, 0xA0B3, 0x5147, 0xA28E, 0x553D, 0xAA7A, 0x44D5, 0x6F45,
    0xDE8A, 0xAD35, 0x4A4B, 0x9496, 0x390D, 0x721A, 0xEB23, 0xC667, 0x9CEF, 0x29FF, 0x53FE, 0xA7FC,
    0x5FD9, 0x47D3, 0x8FA6, 0x0F6D, 0x1EDA, 0x3DB4, 0x7B68, 0xF6D0, 0xB861, 0x60E3, 0xC1C6, 0x93AD,
    0x377B, 0x6EF6, 0xDDEC, 0x45A0, 0x8B40, 0x06A1, 0x0D42, 0x1A84, 0x3508, 0x6A10, 0xAA51, 0x4483,
    0x8906, 0x022D, 0x045A, 0x08B4, 0x1168, 0x76B4, 0xED68, 0xCAF1, 0x85C3, 0x1BA7, 0x374E, 0x6E9C,
    0x3730, 0x6E60, 0xDCC0, 0xA9A1, 0x4363, 0x86C6, 0x1DAD, 0x3331, 0x6662, 0xCCC4, 0x89A9, 0x0373,
    0x06E6, 0x0DCC, 0x1021, 0x2042, 0x4084, 0x8108, 0x1231, 0x2462, 0x48C4,
];

fn ror8(v: u8, n: u32) -> u8 {
    v.rotate_right(n)
}

#[allow(dead_code)]
fn rol8(v: u8, n: u32) -> u8 {
    v.rotate_left(n)
}

/// Derive the 16-bit XOR key from the password (method 1).
pub fn key(password: &str) -> u16 {
    let bytes: Vec<u8> = password.chars().map(|c| c as u32 as u8).collect();
    if bytes.is_empty() || bytes.len() > 15 {
        return 0;
    }
    let mut xor_key = INITIAL_CODE[bytes.len() - 1];
    // The reference implementation walks the matrix downwards from index 104
    // and wraps around to the end when it runs off the front (only reachable
    // for passwords longer than 14 characters).
    let mut current: isize = 0x68;
    for ch in bytes.iter().rev() {
        let mut ch = *ch;
        for _ in 0..7 {
            if ch & 0x40 != 0 {
                xor_key ^= MATRIX[current.rem_euclid(MATRIX.len() as isize) as usize];
            }
            ch = ch.wrapping_shl(1);
            current -= 1;
        }
    }
    xor_key
}

/// Derive the 16-byte obfuscation array from the password (method 1).
pub fn array(password: &str) -> [u8; 16] {
    let bytes: Vec<u8> = password.chars().map(|c| c as u32 as u8).collect();
    let xor_key = key(password);
    let mut out = [0u8; 16];
    let hi = (xor_key >> 8) as u8;
    let lo = (xor_key & 0xFF) as u8;

    let mut index = bytes.len();
    if index % 2 == 1 {
        out[index] = ror8(PAD[0] ^ hi, 1);
        index -= 1;
        out[index] = ror8(bytes[bytes.len() - 1] ^ lo, 1);
    }
    while index > 0 {
        index -= 1;
        out[index] = ror8(bytes[index] ^ hi, 1);
        index -= 1;
        out[index] = ror8(bytes[index] ^ lo, 1);
    }

    let mut index: isize = 15;
    let mut pad_index: isize = 15 - bytes.len() as isize;
    // Python indexes PAD from the end when `pad_index` goes negative, so the
    // odd-length case reads PAD[14] once. Mirror that with rem_euclid.
    while pad_index > 0 {
        out[index as usize] = ror8(PAD[pad_index.rem_euclid(PAD.len() as isize) as usize] ^ hi, 1);
        index -= 1;
        pad_index -= 1;
        out[index as usize] = ror8(PAD[pad_index.rem_euclid(PAD.len() as isize) as usize] ^ lo, 1);
        index -= 1;
        pad_index -= 1;
    }
    out
}

/// The 16-bit verifier a file stores for `password` (MS-OFFCRYPTO 2.3.7.1).
///
/// This is the `verificationBytes` field of an `.xls` `FILEPASS` record, or the
/// FIB's `lKey` for `.doc`. It is *not* the obfuscation key: [`key`] derives
/// that separately, and readers should take the keystream from the password
/// rather than from the key stored in the file.
pub fn verifier(password: &str) -> u16 {
    let mut acc: u16 = 0;
    let mut chars: Vec<u8> = vec![password.chars().count() as u8];
    chars.extend(password.chars().map(|c| c as u32 as u8));
    chars.reverse();
    for byte in chars {
        let carry = if acc & 0x4000 == 0 { 0 } else { 1 };
        acc = (acc.wrapping_shl(1) & 0x7FFF) ^ carry;
        acc ^= byte as u16;
    }
    acc ^ 0xCE4B
}

/// Check a password against the 16-bit verifier stored in the file.
pub fn verify(password: &str, verification_bytes: u16) -> bool {
    verifier(password) == verification_bytes
}

/// Decrypt one run of obfuscated bytes.
///
/// `offset` is the absolute position of `data[0]` in the stream: the array
/// index is `(offset + data.len()) % 16`. For the 4-byte `lbPlyPos` field of a
/// `BoundSheet8` record the index is shifted by a further 4 bytes, which is
/// what `boundsheet_payload` requests.
pub fn decrypt_run(data: &[u8], arr: &[u8; 16], offset: usize, boundsheet_payload: bool) -> Vec<u8> {
    let mut index = (offset + data.len()) % 16;
    if boundsheet_payload {
        index = (index + 4) % 16;
    }
    let mut out = Vec::with_capacity(data.len());
    for byte in data {
        out.push(ror8(byte ^ arr[index], 5));
        index = (index + 1) % 16;
    }
    out
}

/// Apply the obfuscation (the inverse of [`decrypt_run`]).
///
/// Needed to build test fixtures and to write a document back; the transform is
/// an involution up to the rotation, so the inverse rotates the other way.
pub fn encrypt_run(data: &[u8], arr: &[u8; 16], offset: usize, boundsheet_payload: bool) -> Vec<u8> {
    let mut index = (offset + data.len()) % 16;
    if boundsheet_payload {
        index = (index + 4) % 16;
    }
    let mut out = Vec::with_capacity(data.len());
    for byte in data {
        out.push(byte.rotate_left(5) ^ arr[index]);
        index = (index + 1) % 16;
    }
    out
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn verifier_matches_reference_vector() {
        // "VelvetSweatshop" is Excel's default password for encrypted
        // templates; 0x9A0A is the verifier it produces.
        assert!(verify("VelvetSweatshop", 0x9A0A));
        assert!(!verify("VelvetSweatsho", 0x9A0A));
    }

    #[test]
    fn key_and_array_match_reference_implementation() {
        // Values produced by the reference implementation for "secret123".
        assert_eq!(key("secret123"), 0xDBF2);
        assert_eq!(
            array("secret123"),
            [
                0xC0, 0x5F, 0xC8, 0xD4, 0xCB, 0xD7, 0xE1, 0xF4, 0xE0, 0x30, 0x86, 0x12, 0x24, 0x12,
                0x86, 0x31
            ]
        );
    }

    #[test]
    fn decrypts_a_run_encrypted_by_the_reference_implementation() {
        // 64 bytes encrypted by the reference implementation at offset 0.
        const CIPHER: [u8; 64] = [
            0xA0, 0x1E, 0xEA, 0xD7, 0x28, 0x13, 0x44, 0x72, 0x87, 0x78, 0xAF, 0x18, 0xCE, 0xD9,
            0x2A, 0xBC, 0xAE, 0x10, 0xF8, 0xC5, 0x3A, 0x05, 0x52, 0x60, 0x95, 0x66, 0xB1, 0x0A,
            0xDC, 0xCB, 0x3C, 0xAA, 0xBC, 0x02, 0xF6, 0xCB, 0x34, 0x17, 0x40, 0x76, 0x83, 0x74,
            0xA3, 0x14, 0xC2, 0xD5, 0x2E, 0xB8, 0xAA, 0x14, 0xE4, 0xD9, 0x26, 0x19, 0x4E, 0x64,
            0x91, 0x62, 0xB5, 0x06, 0xD0, 0xC7, 0x30, 0xA6,
        ];
        let arr = array("secret123");
        let plain = decrypt_run(&CIPHER, &arr, 0, false);
        let expected: Vec<u8> = (0..64u32).map(|i| ((i * 7 + 3) & 0xFF) as u8).collect();
        assert_eq!(plain, expected);
    }

    #[test]
    fn encrypt_run_round_trips_through_decrypt_run() {
        let arr = array("secret123");
        let plain: Vec<u8> = (0..200u32).map(|i| (i * 7 + 3) as u8).collect();
        for offset in [0usize, 5, 16] {
            let enc = encrypt_run(&plain, &arr, offset, false);
            assert_ne!(enc, plain, "offset {offset}");
            assert_eq!(decrypt_run(&enc, &arr, offset, false), plain);
        }
    }

    #[test]
    fn rotation_is_an_involution_with_the_same_key() {
        let arr = array("pw");
        let plain = b"the quick brown fox";
        // Re-encrypt what we decrypted: rol5 then XOR, the inverse of decrypt.
        let mut idx = plain.len() % 16;
        let mut ct = Vec::new();
        for b in plain {
            ct.push(rol8(*b, 5) ^ arr[idx]);
            idx = (idx + 1) % 16;
        }
        assert_eq!(decrypt_run(&ct, &arr, 0, false), plain.to_vec());
    }
}
