//! # office-crypto — Office 97-2003 binary document encryption
//!
//! Word, Excel and PowerPoint store a password-protected document by encrypting
//! the *contents* of a few streams and leaving the container structure
//! readable. Three schemes exist, all identified by a version pair that lives at
//! a fixed place in the file:
//!
//! | version  | scheme          | key derivation                     |
//! |----------|-----------------|------------------------------------|
//! | 1.1      | legacy RC4      | `MD5` over a 336-byte salt buffer  |
//! | 2.2/3.2/4.2 | RC4 CryptoAPI | `SHA1(salt ‖ password)` + block no. |
//! | (no version) | XOR obfuscation | 15-entry mixing matrix          |
//!
//! ## Where the header lives
//!
//! * `.doc` — at the very start of the table stream (`0Table`/`1Table`), right
//!   after the FIB flags that say `fEncrypted`.
//! * `.xls` — in the `FILEPASS` record (`0x002F`), immediately after `BOF`.
//! * `.ppt` — inside a `CryptSession10Container` record (`0x2F14`) that the
//!   persist directory points at, reached through `UserEditAtom`.
//!
//! ## How the bytes are encrypted
//!
//! RC4 is never used as one long stream. The data is cut into blocks (512 bytes
//! for `.doc`, 1024 for `.xls`, one block per persist object for `.ppt`) and
//! **every block is re-keyed** with its block index. Decrypting a byte
//! therefore requires knowing both its block size and its block number, which
//! is why the format crates pass `block_size` and `first_block` here instead of
//! doing the cipher work themselves.
//!
//! All three schemes were verified against real files written by WPS Office and
//! against Microsoft's own templates; see `docs/encryption.md`.

pub mod md5;
pub mod rc4;
pub mod sha1;
pub mod xor;

use std::fmt;

/// `headerToken` value in the PPT `Current User` stream for a plain file.
pub const TOKEN_PLAINTEXT: u32 = 0xE391_C05F;
/// `headerToken` value for an encrypted file.
pub const TOKEN_ENCRYPTED: u32 = 0xF3D1_C4DF;

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Error {
    /// The buffer ended before the header did.
    Truncated(&'static str),
    /// `vMajor`/`vMinor` pair that this crate does not implement.
    UnsupportedVersion { major: u16, minor: u16 },
    /// The password did not reproduce the stored verifier hash.
    WrongPassword,
    /// Structurally valid header with an unsupported algorithm.
    Unsupported(&'static str),
}

impl fmt::Display for Error {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Error::Truncated(what) => write!(f, "encryption header truncated in {what}"),
            Error::UnsupportedVersion { major, minor } => write!(
                f,
                "unsupported encryption version {major}.{minor} (supported: 1.1 RC4, 2.2/3.2/4.2 RC4 CryptoAPI)"
            ),
            Error::WrongPassword => write!(f, "wrong password"),
            Error::Unsupported(what) => write!(f, "unsupported encryption: {what}"),
        }
    }
}

impl std::error::Error for Error {}

/// The encryption scheme of a document, together with everything needed to
/// check a password and to decrypt its streams.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Scheme {
    /// Version 1.1: MD5-based key schedule, 40-bit RC4 keys.
    LegacyRc4 {
        salt: [u8; 16],
        verifier: [u8; 16],
        verifier_hash: [u8; 16],
    },
    /// Version 2.2 / 3.2 / 4.2: SHA-1 based key schedule, RC4 or AES keys.
    Rc4CryptoApi {
        salt: [u8; 16],
        key_size: u32,
        verifier: [u8; 16],
        verifier_hash: Vec<u8>,
        verifier_hash_size: u32,
        alg_id: u32,
        alg_id_hash: u32,
    },
    /// XOR obfuscation: the 16-bit verifier is all the file stores.
    Xor { key: u16 },
}

impl Scheme {
    /// Human-readable name, used by the CLI and by error messages.
    pub fn name(&self) -> &'static str {
        match self {
            Scheme::LegacyRc4 { .. } => "RC4 (legacy, v1.1)",
            Scheme::Rc4CryptoApi { key_size, .. } => match key_size {
                40 => "RC4 CryptoAPI (40-bit)",
                56 => "RC4 CryptoAPI (56-bit)",
                128 => "RC4 CryptoAPI (128-bit)",
                _ => "RC4 CryptoAPI",
            },
            Scheme::Xor { .. } => "XOR obfuscation",
        }
    }

    /// Key length in bits.
    pub fn key_size(&self) -> u32 {
        match self {
            Scheme::LegacyRc4 { .. } => 40,
            Scheme::Rc4CryptoApi { key_size, .. } => *key_size,
            Scheme::Xor { .. } => 0,
        }
    }

    /// Check `password` against the verifier stored in the file.
    ///
    /// This is the only way to know a password is right before decrypting:
    /// the file stores `RC4(key0, verifier)` and `RC4(key0, hash(verifier))`.
    pub fn verify(&self, password: &str) -> bool {
        match self {
            Scheme::LegacyRc4 {
                salt,
                verifier,
                verifier_hash,
            } => {
                let key = legacy_key(password, salt, 0);
                let mut buf = verifier.to_vec();
                buf.extend_from_slice(verifier_hash);
                let plain = rc4::apply(&key, &buf);
                md5::digest(&plain[..16]) == plain[16..32]
            }
            Scheme::Rc4CryptoApi {
                salt,
                key_size,
                verifier,
                verifier_hash,
                ..
            } => {
                let key = cryptoapi_key(password, salt, *key_size, 0);
                let mut buf = verifier.to_vec();
                buf.extend_from_slice(verifier_hash);
                let plain = rc4::apply(&key, &buf);
                sha1::digest(&plain[..16])[..] == plain[16..]
            }
            Scheme::Xor { key } => xor::verify(password, *key),
        }
    }
}

/// A parsed encryption header plus the number of bytes it occupied.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Header {
    pub scheme: Scheme,
    /// Bytes consumed, so the caller can skip the header in the stream.
    pub consumed: usize,
}

/// Parse the versioned header that `.doc` and `.xls` store inline.
///
/// `data` must start at the `vMajor`/`vMinor` pair.
pub fn parse_header(data: &[u8]) -> Result<Header, Error> {
    if data.len() < 4 {
        return Err(Error::Truncated("EncryptionVersionInfo"));
    }
    let major = u16::from_le_bytes([data[0], data[1]]);
    let minor = u16::from_le_bytes([data[2], data[3]]);
    match (major, minor) {
        (1, 1) => {
            // v1.1 stores the verifier straight after the version: no
            // EncryptionHeader, no size fields — 48 bytes and then data.
            if data.len() < 4 + 48 {
                return Err(Error::Truncated("RC4 verifier"));
            }
            let mut salt = [0u8; 16];
            let mut verifier = [0u8; 16];
            let mut verifier_hash = [0u8; 16];
            salt.copy_from_slice(&data[4..20]);
            verifier.copy_from_slice(&data[20..36]);
            verifier_hash.copy_from_slice(&data[36..52]);
            Ok(Header {
                scheme: Scheme::LegacyRc4 {
                    salt,
                    verifier,
                    verifier_hash,
                },
                consumed: 52,
            })
        }
        (2..=4, 2) => {
            let (scheme, consumed) = parse_cryptoapi(data)?;
            Ok(Header { scheme, consumed })
        }
        _ => Err(Error::UnsupportedVersion { major, minor }),
    }
}

/// Parse a version-2/3/4 `EncryptionHeader` + `EncryptionVerifier` pair.
///
/// Layout (verified against `samples/encrypted.ppt`, a WPS-written file):
///
/// ```text
/// 0   vMajor u16, vMinor u16          // already consumed by the caller
/// 4   flags u32                       // fCryptoAPI | fAES
/// 8   headerSize u32                  // 126 in the sample
/// 12  EncryptionHeader {              // exactly headerSize bytes
///        flags u32, sizeExtra u32, algId u32, algIdHash u32,
///        keySize u32, providerType u32, reserved1 u32, reserved2 u32,
///        cspName: UTF-16LE, null terminated }
/// 12+headerSize  EncryptionVerifier {
///        saltSize u32, salt[16], encryptedVerifier[16],
///        verifierHashSize u32, encryptedVerifierHash[20 or 32] }
/// ```
fn parse_cryptoapi(data: &[u8]) -> Result<(Scheme, usize), Error> {
    if data.len() < 16 {
        return Err(Error::Truncated("EncryptionHeader"));
    }
    let rd = |o: usize| u32::from_le_bytes([data[o], data[o + 1], data[o + 2], data[o + 3]]);
    let header_size = rd(8) as usize;
    if header_size < 32 || data.len() < 12 + header_size {
        return Err(Error::Truncated("EncryptionHeader body"));
    }
    let body = &data[12..12 + header_size];
    let bd = |o: usize| u32::from_le_bytes([body[o], body[o + 1], body[o + 2], body[o + 3]]);
    let _body_flags = bd(0);
    let alg_id = bd(8);
    let alg_id_hash = bd(12);
    let mut key_size = bd(16);
    if key_size == 0 {
        // MS-OFFCRYPTO: a zero keySize means the default 40-bit key.
        key_size = 40;
    }

    let pos = 12 + header_size;
    if data.len() < pos + 40 {
        return Err(Error::Truncated("EncryptionVerifier"));
    }
    let salt_size = rd(pos);
    if salt_size != 16 {
        return Err(Error::Unsupported("salt size other than 16"));
    }
    let mut salt = [0u8; 16];
    salt.copy_from_slice(&data[pos + 4..pos + 20]);
    let mut verifier = [0u8; 16];
    verifier.copy_from_slice(&data[pos + 20..pos + 36]);
    let verifier_hash_size = rd(pos + 36);
    let hash_len = match alg_id {
        0x0000_6801 => 20,                                        // CALG_RC4 -> SHA-1
        0x0000_660E | 0x0000_660F | 0x0000_6610 | 0x0000_6611 => 32, // AES -> SHA-512/256
        _ => return Err(Error::Unsupported("algorithm id")),
    };
    if data.len() < pos + 40 + hash_len {
        return Err(Error::Truncated("EncryptedVerifierHash"));
    }
    let verifier_hash = data[pos + 40..pos + 40 + hash_len].to_vec();

    Ok((
        Scheme::Rc4CryptoApi {
            salt,
            key_size,
            verifier,
            verifier_hash,
            verifier_hash_size,
            alg_id,
            alg_id_hash,
        },
        pos + 40 + hash_len,
    ))
}

/// UTF-16LE encoding of a password, as every Office scheme expects it.
pub fn password_utf16le(password: &str) -> Vec<u8> {
    let mut out = Vec::with_capacity(password.len() * 2);
    for unit in password.encode_utf16() {
        out.extend_from_slice(&unit.to_le_bytes());
    }
    out
}

/// Legacy RC4 (v1.1) block key.
///
/// ```text
/// H0 = MD5(password_utf16le)
/// IB = (H0[0..5] ‖ salt) * 16          // 336 bytes
/// H1 = MD5(IB)
/// key = MD5(H1[0..5] ‖ block_le32)     // 16 bytes, but only 40 bits are used
/// ```
pub fn legacy_key(password: &str, salt: &[u8; 16], block: u32) -> [u8; 16] {
    let h0 = md5::digest(&password_utf16le(password));
    let mut ib = Vec::with_capacity(21 * 16);
    for _ in 0..16 {
        ib.extend_from_slice(&h0[..5]);
        ib.extend_from_slice(salt);
    }
    let h1 = md5::digest(&ib);
    let mut fin = Vec::with_capacity(9);
    fin.extend_from_slice(&h1[..5]);
    fin.extend_from_slice(&block.to_le_bytes());
    md5::digest(&fin)
}

/// RC4 CryptoAPI (v2.2/3.2/4.2) block key.
///
/// ```text
/// H0 = SHA1(salt ‖ password_utf16le)
/// key = SHA1(H0 ‖ block_le32)[0 .. key_size/8]     // 40-bit keys are padded
/// ```
pub fn cryptoapi_key(password: &str, salt: &[u8], key_size: u32, block: u32) -> Vec<u8> {
    let mut buf = salt.to_vec();
    buf.extend_from_slice(&password_utf16le(password));
    let h0 = sha1::digest(&buf);
    let mut buf2 = h0.to_vec();
    buf2.extend_from_slice(&block.to_le_bytes());
    let hf = sha1::digest(&buf2);
    if key_size == 40 {
        // 40-bit keys are stored right-padded with zeros to the full block size.
        let mut key = vec![0u8; 16];
        key[..5].copy_from_slice(&hf[..5]);
        key
    } else {
        let n = (key_size as usize / 8).min(20);
        hf[..n].to_vec()
    }
}

/// Decrypt `data` in re-keyed RC4 blocks, starting at block 0.
pub fn decrypt_legacy_blocks(
    data: &[u8],
    password: &str,
    salt: &[u8; 16],
    block_size: usize,
) -> Vec<u8> {
    decrypt_legacy_blocks_from(data, password, salt, block_size, 0)
}

/// Decrypt `data` in re-keyed RC4 blocks, starting at block `first_block`.
pub fn decrypt_legacy_blocks_from(
    data: &[u8],
    password: &str,
    salt: &[u8; 16],
    block_size: usize,
    first_block: u32,
) -> Vec<u8> {
    let block_size = block_size.max(1);
    let mut out = Vec::with_capacity(data.len());
    for (i, chunk) in data.chunks(block_size).enumerate() {
        let key = legacy_key(password, salt, first_block + i as u32);
        out.extend_from_slice(&rc4::apply(&key, chunk));
    }
    out
}

/// Decrypt `data` in re-keyed RC4 CryptoAPI blocks.
pub fn decrypt_cryptoapi_blocks(
    data: &[u8],
    password: &str,
    salt: &[u8],
    key_size: u32,
    block_size: usize,
    first_block: u32,
) -> Vec<u8> {
    let block_size = block_size.max(1);
    let mut out = Vec::with_capacity(data.len());
    for (i, chunk) in data.chunks(block_size).enumerate() {
        let key = cryptoapi_key(password, salt, key_size, first_block + i as u32);
        out.extend_from_slice(&rc4::apply(&key, chunk));
    }
    out
}

/// Decrypt a whole buffer with the scheme's default block layout.
///
/// `.doc` uses 512-byte blocks, `.xls` 1024-byte blocks. `.ppt` encrypts each
/// persist object separately and must call [`decrypt_cryptoapi_blocks`]
/// directly with the object's own block number.
pub fn decrypt(
    scheme: &Scheme,
    password: &str,
    data: &[u8],
    block_size: usize,
) -> Result<Vec<u8>, Error> {
    if !scheme.verify(password) {
        return Err(Error::WrongPassword);
    }
    Ok(match scheme {
        Scheme::LegacyRc4 { salt, .. } => decrypt_legacy_blocks(data, password, salt, block_size),
        Scheme::Rc4CryptoApi {
            salt, key_size, ..
        } => decrypt_cryptoapi_blocks(data, password, salt, *key_size, block_size, 0),
        Scheme::Xor { .. } => {
            return Err(Error::Unsupported(
                "XOR obfuscation needs per-record positions; use xor::decrypt_run",
            ))
        }
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn legacy_key_is_stable() {
        // Computed with the reference implementation for "secret123" and the
        // salt stored in samples/encrypted.doc.
        let salt: [u8; 16] = [
            0xb6, 0x6d, 0x0c, 0xf5, 0x06, 0xce, 0x13, 0xaf, 0xdc, 0xa7, 0xcc, 0x58, 0x75, 0x99,
            0x8d, 0xbf,
        ];
        let key = legacy_key("secret123", &salt, 0);
        assert_eq!(
            key.iter().map(|b| format!("{b:02x}")).collect::<String>(),
            "f4419525106dd874aca5e4f4c413d528"
        );
    }

    #[test]
    fn cryptoapi_key_is_stable() {
        let salt = [0x11u8; 16];
        // Deterministic; the exact bytes were cross-checked against the
        // reference implementation while developing the crate.
        let k40 = cryptoapi_key("pw", &salt, 40, 0);
        assert_eq!(k40.len(), 16);
        assert_eq!(&k40[5..], &[0u8; 11]);
        let k128 = cryptoapi_key("pw", &salt, 128, 0);
        assert_eq!(k128.len(), 16);
        assert_ne!(k40, k128);
    }

    #[test]
    fn header_round_trips_legacy_layout() {
        let mut data = vec![1, 0, 1, 0];
        data.extend_from_slice(&[0xAA; 48]);
        let h = parse_header(&data).unwrap();
        assert_eq!(h.consumed, 52);
        match h.scheme {
            Scheme::LegacyRc4 { salt, verifier, .. } => {
                assert_eq!(salt, [0xAA; 16]);
                assert_eq!(verifier, [0xAA; 16]);
            }
            other => panic!("unexpected scheme {other:?}"),
        }
    }

    #[test]
    fn rejects_unknown_versions() {
        let data = [9u8, 0, 9, 0, 0, 0, 0, 0];
        assert!(matches!(
            parse_header(&data),
            Err(Error::UnsupportedVersion { major: 9, minor: 9 })
        ));
    }

    #[test]
    fn password_encoding_is_utf16le() {
        assert_eq!(password_utf16le("A"), vec![0x41, 0x00]);
        assert_eq!(password_utf16le("Ω"), vec![0xA9, 0x03]);
    }
}
