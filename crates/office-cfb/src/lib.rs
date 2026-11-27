//! # office-cfb — Compound File Binary reader
//!
//! `.doc`, `.xls` and `.ppt` are not single files: they are **OLE2 compound
//! documents**, a tiny in-file FAT filesystem (MS-CFB). Every Office binary
//! format lives in a named *stream* inside this container:
//!
//! | format | stream name(s)                              |
//! |--------|---------------------------------------------|
//! | .doc   | `WordDocument`, `0Table`/`1Table`, `Data`   |
//! | .xls   | `Workbook` (BIFF8) or `Book` (BIFF5)        |
//! | .ppt   | `PowerPoint Document`, `Current User`       |
//!
//! This crate implements the container only; the format crates consume
//! [`Cfb::stream`].
//!
//! ## On-disk layout
//!
//! ```text
//! +--------------------+ 0
//! | header (512 bytes) |  signature D0 CF 11 E0 A1 B1 1A E1
//! +--------------------+ 512
//! | sector 0           |  FAT sector (usually)
//! | sector 1           |  directory sector (usually)
//! | ...                |
//! +--------------------+
//! ```
//!
//! A sector is 512 bytes (major version 3) or 4096 bytes (major version 4);
//! sector `n` starts at byte `(n + 1) * sector_size` — the header occupies the
//! first sector-sized slot. Small streams (< 4096 bytes, the *mini stream
//! cutoff*) are packed into the root entry's stream in 64-byte mini sectors.
//!
//! Verified against files written by Word/Excel/PowerPoint and WPS Office;
//! see `docs/cfb.md` for the byte-level walkthrough.

use std::collections::HashSet;
use std::fmt;
use std::fs;
use std::path::Path;

pub type Result<T> = std::result::Result<T, Error>;

/// CFB header magic: `D0 CF 11 E0 A1 B1 1A E1`.
pub const SIGNATURE: [u8; 8] = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];

const FREESECT: u32 = 0xFFFF_FFFF;
const ENDOFCHAIN: u32 = 0xFFFF_FFFE;
/// FAT sector marker; also a valid chain terminator.
#[allow(dead_code)]
const FATSECT: u32 = 0xFFFF_FFFD;
/// DIFAT sector marker; also a valid chain terminator.
#[allow(dead_code)]
const DIFSECT: u32 = 0xFFFF_FFFC;
/// "no stream" marker in directory sibling/child fields.
pub const NOSTREAM: u32 = 0xFFFF_FFFF;

#[derive(Debug)]
pub enum Error {
    Io(std::io::Error),
    BadSignature,
    Truncated { need: usize, have: usize },
    BadVersion(u16),
    BadSectorShift(u16),
    BadByteOrder(u16),
    BadDirectory(&'static str),
    CycleDetected(u32),
    NoSuchStream(String),
}

impl fmt::Display for Error {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Error::Io(e) => write!(f, "i/o error: {e}"),
            Error::BadSignature => write!(f, "not a compound file (bad D0CF11E0 signature)"),
            Error::Truncated { need, have } => {
                write!(f, "truncated file: need {need} bytes, have {have}")
            }
            Error::BadVersion(v) => write!(f, "unsupported CFB major version {v}"),
            Error::BadSectorShift(s) => write!(f, "unsupported sector shift {s}"),
            Error::BadByteOrder(b) => write!(f, "bad byte-order mark 0x{b:04X} (expected 0xFFFE)"),
            Error::BadDirectory(m) => write!(f, "bad directory: {m}"),
            Error::CycleDetected(s) => write!(f, "sector chain cycle at sector {s}"),
            Error::NoSuchStream(n) => write!(f, "stream {n:?} not found"),
        }
    }
}

impl std::error::Error for Error {}

impl From<std::io::Error> for Error {
    fn from(e: std::io::Error) -> Self {
        Error::Io(e)
    }
}

/// Object type of a directory entry (MS-CFB §2.6.1).
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum EntryKind {
    /// Unused directory slot.
    Empty,
    /// A storage (a "directory" that can contain streams).
    Storage,
    /// A stream (the actual data).
    Stream,
    /// The root storage; its stream is the mini-stream backing store.
    Root,
}

impl EntryKind {
    fn from_byte(b: u8) -> EntryKind {
        match b {
            1 => EntryKind::Storage,
            2 => EntryKind::Stream,
            5 => EntryKind::Root,
            _ => EntryKind::Empty,
        }
    }
}

/// One 128-byte directory entry.
#[derive(Clone, Debug)]
pub struct DirEntry {
    /// Index in the directory array (also the "stream id").
    pub id: u32,
    /// UTF-16 name with the trailing NUL stripped.
    pub name: String,
    pub kind: EntryKind,
    /// Starting sector (mini sector when `size < cutoff`).
    pub start_sector: u32,
    /// Stream length in bytes (0 for storages).
    pub size: u64,
    pub left: u32,
    pub right: u32,
    pub child: u32,
    pub clsid: [u8; 16],
}

/// A parsed compound file. All sectors are read eagerly, so the struct owns
/// the file bytes and every `stream()` call is a cheap chain walk.
pub mod write;

pub use write::{write, OutputStream};

pub struct Cfb {
    data: Vec<u8>,
    major_version: u16,
    sector_size: usize,
    mini_sector_size: usize,
    cutoff: u32,
    fat: Vec<u32>,
    mini_fat: Vec<u32>,
    entries: Vec<DirEntry>,
    mini_stream: Vec<u8>,
    /// Sector chain of the root entry's stream (the mini-stream backing store),
    /// kept so a mini sector can be mapped back to a byte range of the file.
    mini_stream_chain: Vec<u32>,
}

impl Cfb {
    /// Read and parse a file from disk.
    pub fn open<P: AsRef<Path>>(path: P) -> Result<Cfb> {
        Cfb::parse(fs::read(path)?)
    }

    /// Parse an in-memory compound file.
    pub fn parse(data: Vec<u8>) -> Result<Cfb> {
        if data.len() < 512 {
            return Err(Error::Truncated { need: 512, have: data.len() });
        }
        if data[..8] != SIGNATURE {
            return Err(Error::BadSignature);
        }
        let u16_at = |o: usize| u16::from_le_bytes([data[o], data[o + 1]]);
        let u32_at = |o: usize| {
            u32::from_le_bytes([data[o], data[o + 1], data[o + 2], data[o + 3]])
        };

        let major = u16_at(0x1A);
        if major != 3 && major != 4 {
            return Err(Error::BadVersion(major));
        }
        let byte_order = u16_at(0x1C);
        if byte_order != 0xFFFE {
            return Err(Error::BadByteOrder(byte_order));
        }
        let sector_shift = u16_at(0x1E);
        if !(9..=12).contains(&sector_shift) {
            return Err(Error::BadSectorShift(sector_shift));
        }
        let sector_size = 1usize << sector_shift;
        let mini_sector_size = 1usize << u16_at(0x20);

        let num_fat_sectors = u32_at(0x2C) as usize;
        let first_dir_sector = u32_at(0x30);
        let cutoff = u32_at(0x38);
        let first_mini_fat = u32_at(0x3C);
        let num_mini_fat = u32_at(0x40) as usize;
        let first_difat = u32_at(0x44);
        let num_difat = u32_at(0x48) as usize;

        // ---- DIFAT: 109 header slots + chained DIFAT sectors -------------
        let mut difat: Vec<u32> = (0..109).map(|i| u32_at(0x4C + i * 4)).collect();
        let mut sector = first_difat;
        let mut seen = HashSet::new();
        for _ in 0..num_difat {
            if sector == ENDOFCHAIN || sector == FREESECT {
                break;
            }
            if !seen.insert(sector) {
                return Err(Error::CycleDetected(sector));
            }
            let off = Self::sector_offset(sector, sector_size)?;
            let end = off + sector_size;
            if end > data.len() {
                return Err(Error::Truncated { need: end, have: data.len() });
            }
            // sector_size/4 - 1 entries + next DIFAT sector in the last slot
            for i in 0..(sector_size / 4 - 1) {
                difat.push(u32::from_le_bytes([
                    data[off + i * 4],
                    data[off + i * 4 + 1],
                    data[off + i * 4 + 2],
                    data[off + i * 4 + 3],
                ]));
            }
            sector = u32::from_le_bytes([
                data[end - 4],
                data[end - 3],
                data[end - 2],
                data[end - 1],
            ]);
        }
        // Header slot 0..num_fat_sectors are the real FAT sector numbers; the
        // rest are FREESECT padding. Some writers leave the count at 0, in
        // which case every non-free slot is used.
        let mut fat_sectors: Vec<u32> = difat
            .into_iter()
            .filter(|&s| s != FREESECT && s != ENDOFCHAIN)
            .collect();
        if num_fat_sectors > 0 && fat_sectors.len() > num_fat_sectors {
            fat_sectors.truncate(num_fat_sectors);
        }

        // ---- FAT ---------------------------------------------------------
        let mut fat: Vec<u32> = Vec::with_capacity(fat_sectors.len() * (sector_size / 4));
        for &fs in &fat_sectors {
            let off = Self::sector_offset(fs, sector_size)?;
            let end = off + sector_size;
            if end > data.len() {
                return Err(Error::Truncated { need: end, have: data.len() });
            }
            for i in 0..(sector_size / 4) {
                fat.push(u32::from_le_bytes([
                    data[off + i * 4],
                    data[off + i * 4 + 1],
                    data[off + i * 4 + 2],
                    data[off + i * 4 + 3],
                ]));
            }
        }

        // ---- directory ---------------------------------------------------
        let dir_data = Self::read_chain(&data, sector_size, &fat, first_dir_sector)?;
        if dir_data.len() < 128 {
            return Err(Error::BadDirectory("directory shorter than one entry"));
        }
        let mut entries = Vec::with_capacity(dir_data.len() / 128);
        for (i, raw) in dir_data.chunks_exact(128).enumerate() {
            let name_len = u16::from_le_bytes([raw[0x40], raw[0x41]]) as usize;
            let name = if (2..=64).contains(&name_len) {
                let units: Vec<u16> = raw[..name_len - 2]
                    .chunks_exact(2)
                    .map(|c| u16::from_le_bytes([c[0], c[1]]))
                    .collect();
                String::from_utf16_lossy(&units)
            } else {
                String::new()
            };
            let kind = EntryKind::from_byte(raw[0x42]);
            if kind == EntryKind::Empty {
                entries.push(DirEntry {
                    id: i as u32,
                    name,
                    kind,
                    start_sector: 0,
                    size: 0,
                    left: NOSTREAM,
                    right: NOSTREAM,
                    child: NOSTREAM,
                    clsid: [0; 16],
                });
                continue;
            }
            let mut clsid = [0u8; 16];
            clsid.copy_from_slice(&raw[0x50..0x60]);
            let size = u64::from_le_bytes(raw[0x78..0x80].try_into().unwrap());
            entries.push(DirEntry {
                id: i as u32,
                name,
                kind,
                start_sector: u32::from_le_bytes([raw[0x74], raw[0x75], raw[0x76], raw[0x77]]),
                // major version 3 stores only the low 32 bits; the upper half
                // is documented as "must be zero" and is ignored here.
                size: if major == 3 { size & 0xFFFF_FFFF } else { size },
                left: u32::from_le_bytes([raw[0x44], raw[0x45], raw[0x46], raw[0x47]]),
                right: u32::from_le_bytes([raw[0x48], raw[0x49], raw[0x4A], raw[0x4B]]),
                child: u32::from_le_bytes([raw[0x4C], raw[0x4D], raw[0x4E], raw[0x4F]]),
                clsid,
            });
        }

        // ---- mini FAT + mini stream --------------------------------------
        let mut mini_fat = Vec::new();
        if first_mini_fat != ENDOFCHAIN && first_mini_fat != FREESECT {
            let raw = Self::read_chain(&data, sector_size, &fat, first_mini_fat)?;
            let want = num_mini_fat * (sector_size / 4);
            let take = want.min(raw.len() / 4);
            mini_fat.reserve(take);
            for i in 0..take {
                mini_fat.push(u32::from_le_bytes([
                    raw[i * 4],
                    raw[i * 4 + 1],
                    raw[i * 4 + 2],
                    raw[i * 4 + 3],
                ]));
            }
        }
        let root = entries
            .iter()
            .find(|e| e.kind == EntryKind::Root)
            .ok_or(Error::BadDirectory("no root entry"))?;
        let mini_stream_chain = if root.start_sector == ENDOFCHAIN {
            Vec::new()
        } else {
            Self::sector_chain(&fat, root.start_sector)
        };
        let mini_stream = if root.start_sector == ENDOFCHAIN {
            Vec::new()
        } else {
            Self::read_chain(&data, sector_size, &fat, root.start_sector)?
        };

        Ok(Cfb {
            data,
            major_version: major,
            sector_size,
            mini_sector_size,
            cutoff,
            fat,
            mini_fat,
            entries,
            mini_stream,
            mini_stream_chain,
        })
    }

    /// Follow a FAT chain without reading the bytes.
    fn sector_chain(fat: &[u32], start: u32) -> Vec<u32> {
        let mut out = Vec::new();
        let mut sector = start;
        let mut seen = HashSet::new();
        while sector != ENDOFCHAIN && sector != FREESECT && (sector as usize) < fat.len() {
            if !seen.insert(sector) {
                break;
            }
            out.push(sector);
            sector = fat[sector as usize];
        }
        out
    }

    /// Where a stream's bytes physically live in the file, as
    /// `(file_offset, length)` runs in stream order.
    ///
    /// Writes need this: a stream is not necessarily contiguous, and the
    /// directory (or the FAT) may sit in the middle of the file, so "assume the
    /// stream starts at some offset" is wrong — `hello.doc` keeps its second
    /// directory sector at 0xC00 right where its `WordDocument` stream lives.
    pub fn physical_runs(&self, name: &str) -> Result<Vec<(usize, usize)>> {
        let e = self
            .entries
            .iter()
            .find(|e| e.name == name && e.kind == EntryKind::Stream)
            .ok_or_else(|| Error::NoSuchStream(name.to_string()))?;
        let mut runs: Vec<(usize, usize)> = Vec::new();
        let mut left = e.size as usize;
        if left == 0 {
            return Ok(runs);
        }
        if e.size < self.cutoff as u64 {
            let per_sector = self.sector_size / self.mini_sector_size;
            let mut sector = e.start_sector;
            let mut seen = HashSet::new();
            while sector != ENDOFCHAIN && sector != FREESECT && (sector as usize) < self.mini_fat.len()
            {
                if !seen.insert(sector) {
                    return Err(Error::CycleDetected(sector));
                }
                let holder = self
                    .mini_stream_chain
                    .get(sector as usize / per_sector)
                    .copied()
                    .ok_or(Error::Truncated { need: sector as usize, have: self.mini_stream_chain.len() })?;
                let off = Self::sector_offset(holder, self.sector_size)?
                    + (sector as usize % per_sector) * self.mini_sector_size;
                let len = left.min(self.mini_sector_size);
                runs.push((off, len));
                left -= len;
                if left == 0 {
                    break;
                }
                sector = self.mini_fat[sector as usize];
            }
        } else {
            for sector in Self::sector_chain(&self.fat, e.start_sector) {
                let off = Self::sector_offset(sector, self.sector_size)?;
                let len = left.min(self.sector_size);
                runs.push((off, len));
                left -= len;
                if left == 0 {
                    break;
                }
            }
        }
        Ok(runs)
    }

    /// The whole file, for callers that patch streams in place.
    pub fn as_bytes(&self) -> &[u8] {
        &self.data
    }

    fn sector_offset(sector: u32, sector_size: usize) -> Result<usize> {
        (sector as usize)
            .checked_add(1)
            .and_then(|s| s.checked_mul(sector_size))
            .ok_or(Error::BadDirectory("sector offset overflow"))
    }

    fn read_chain(data: &[u8], sector_size: usize, fat: &[u32], start: u32) -> Result<Vec<u8>> {
        let mut out = Vec::new();
        let mut sector = start;
        let mut seen = HashSet::new();
        while sector != ENDOFCHAIN && sector != FREESECT {
            if sector as usize >= fat.len() {
                break; // dangling chain: stop at the end of the FAT
            }
            if !seen.insert(sector) {
                return Err(Error::CycleDetected(sector));
            }
            let off = Self::sector_offset(sector, sector_size)?;
            let end = off + sector_size;
            if end > data.len() {
                return Err(Error::Truncated { need: end, have: data.len() });
            }
            out.extend_from_slice(&data[off..end]);
            sector = fat[sector as usize];
        }
        Ok(out)
    }

    pub fn sector_size(&self) -> usize {
        self.sector_size
    }

    /// CFB major version: 3 (512-byte sectors) or 4 (4096-byte sectors).
    pub fn version(&self) -> u16 {
        self.major_version
    }

    pub fn mini_sector_size(&self) -> usize {
        self.mini_sector_size
    }

    /// Size threshold below which streams live in the mini stream.
    pub fn mini_stream_cutoff(&self) -> u32 {
        self.cutoff
    }

    /// All directory entries, in directory order.
    pub fn entries(&self) -> &[DirEntry] {
        &self.entries
    }

    /// Look up a directory entry by exact name (case-insensitive, as Office
    /// itself matches stream names that way).
    pub fn entry(&self, name: &str) -> Option<&DirEntry> {
        self.entries
            .iter()
            .find(|e| e.kind != EntryKind::Empty && e.name.eq_ignore_ascii_case(name))
    }

    /// Read a stream by name.
    pub fn stream(&self, name: &str) -> Result<Vec<u8>> {
        let e = self
            .entry(name)
            .ok_or_else(|| Error::NoSuchStream(name.to_string()))?;
        self.stream_of(e)
    }

    /// Read a stream by directory index.
    pub fn stream_by_id(&self, id: u32) -> Result<Vec<u8>> {
        let e = self
            .entries
            .get(id as usize)
            .ok_or_else(|| Error::NoSuchStream(format!("dir entry #{id}")))?;
        self.stream_of(e)
    }

    fn stream_of(&self, e: &DirEntry) -> Result<Vec<u8>> {
        if e.kind != EntryKind::Stream {
            return Err(Error::NoSuchStream(e.name.clone()));
        }
        if e.size < self.cutoff as u64 {
            // mini stream: 64-byte sectors inside the root entry's stream
            let mut out = Vec::with_capacity(e.size as usize);
            let mut sector = e.start_sector;
            let mut seen = HashSet::new();
            while sector != ENDOFCHAIN && sector != FREESECT {
                if sector as usize >= self.mini_fat.len() {
                    break;
                }
                if !seen.insert(sector) {
                    return Err(Error::CycleDetected(sector));
                }
                let off = sector as usize * self.mini_sector_size;
                let end = off + self.mini_sector_size;
                if end > self.mini_stream.len() {
                    return Err(Error::Truncated { need: end, have: self.mini_stream.len() });
                }
                out.extend_from_slice(&self.mini_stream[off..end]);
                sector = self.mini_fat[sector as usize];
            }
            out.truncate(e.size as usize);
            Ok(out)
        } else {
            let mut out = Self::read_chain(&self.data, self.sector_size, &self.fat, e.start_sector)?;
            out.truncate(e.size as usize);
            Ok(out)
        }
    }

    /// Direct children of a storage (walks the red-black sibling tree).
    pub fn children(&self, storage_id: u32) -> Vec<&DirEntry> {
        let mut out = Vec::new();
        let mut stack = Vec::new();
        if let Some(s) = self.entries.get(storage_id as usize) {
            if s.child != NOSTREAM {
                stack.push(s.child);
            }
        }
        let mut seen = HashSet::new();
        while let Some(id) = stack.pop() {
            if id == NOSTREAM || !seen.insert(id) {
                continue;
            }
            if let Some(e) = self.entries.get(id as usize) {
                out.push(e);
                if e.left != NOSTREAM {
                    stack.push(e.left);
                }
                if e.right != NOSTREAM {
                    stack.push(e.right);
                }
            }
        }
        out.sort_by_key(|e| e.id);
        out
    }

    /// `Root Entry/Storage/Stream` style paths for every non-empty entry.
    pub fn paths(&self) -> Vec<(String, u32)> {
        let mut out = Vec::new();
        let root = match self.entries.iter().find(|e| e.kind == EntryKind::Root) {
            Some(r) => r.id,
            None => return out,
        };
        fn walk(c: &Cfb, id: u32, prefix: &str, out: &mut Vec<(String, u32)>) {
            for e in c.children(id) {
                let path = if prefix.is_empty() {
                    e.name.clone()
                } else {
                    format!("{prefix}/{}", e.name)
                };
                out.push((path.clone(), e.id));
                if e.kind == EntryKind::Storage {
                    walk(c, e.id, &path, out);
                }
            }
        }
        walk(self, root, "", &mut out);
        out
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    /// A hand-built compound file: header + FAT + directory + mini FAT +
    /// mini stream, holding one 5-byte stream (`Tiny`) in the mini stream.
    fn tiny_cfb() -> Vec<u8> {
        let ss = 512usize;
        let mut d = vec![0u8; ss * 5];
        d[..8].copy_from_slice(&SIGNATURE);
        d[0x18..0x1A].copy_from_slice(&0x3Eu16.to_le_bytes()); // minor
        d[0x1A..0x1C].copy_from_slice(&3u16.to_le_bytes()); // major
        d[0x1C..0x1E].copy_from_slice(&0xFFFEu16.to_le_bytes());
        d[0x1E..0x20].copy_from_slice(&9u16.to_le_bytes()); // 512
        d[0x20..0x22].copy_from_slice(&6u16.to_le_bytes()); // 64
        d[0x2C..0x30].copy_from_slice(&1u32.to_le_bytes()); // 1 FAT sector
        d[0x30..0x34].copy_from_slice(&1u32.to_le_bytes()); // dir at sector 1
        d[0x38..0x3C].copy_from_slice(&4096u32.to_le_bytes());
        d[0x3C..0x40].copy_from_slice(&2u32.to_le_bytes()); // mini FAT at sector 2
        d[0x40..0x44].copy_from_slice(&1u32.to_le_bytes()); // 1 mini FAT sector
        d[0x44..0x48].copy_from_slice(&ENDOFCHAIN.to_le_bytes());
        d[0x4C..0x50].copy_from_slice(&0u32.to_le_bytes()); // FAT sector 0

        // FAT: sector0 = FATSECT, sectors 1..3 = single-sector chains
        let fat = (0..128u32)
            .map(|i| match i {
                0 => FATSECT,
                1..=3 => ENDOFCHAIN,
                _ => FREESECT,
            })
            .collect::<Vec<_>>();
        for (i, v) in fat.iter().enumerate() {
            d[ss + i * 4..ss + i * 4 + 4].copy_from_slice(&v.to_le_bytes());
        }

        // directory sector at 2*512: root + one stream
        let dir = 2 * ss;
        let write_entry = |d: &mut [u8], off: usize, name: &str, kind: u8, start: u32, size: u64, child: u32| {
            let units: Vec<u16> = name.encode_utf16().collect();
            for (i, u) in units.iter().enumerate() {
                d[off + i * 2..off + i * 2 + 2].copy_from_slice(&u.to_le_bytes());
            }
            d[off + 0x40..off + 0x42].copy_from_slice(&((units.len() * 2 + 2) as u16).to_le_bytes());
            d[off + 0x42] = kind;
            d[off + 0x44..off + 0x48].copy_from_slice(&NOSTREAM.to_le_bytes());
            d[off + 0x48..off + 0x4C].copy_from_slice(&NOSTREAM.to_le_bytes());
            d[off + 0x4C..off + 0x50].copy_from_slice(&child.to_le_bytes());
            d[off + 0x74..off + 0x78].copy_from_slice(&start.to_le_bytes());
            d[off + 0x78..off + 0x80].copy_from_slice(&size.to_le_bytes());
        };
        // root: mini stream at sector 3, 64 bytes long
        write_entry(&mut d, dir, "Root Entry", 5, 3, 64, 1);
        write_entry(&mut d, dir + 128, "Tiny", 2, 0, 5, NOSTREAM);
        // mini FAT at sector 2: mini sector 0 is the only one, end of chain
        d[3 * ss..3 * ss + 4].copy_from_slice(&ENDOFCHAIN.to_le_bytes());
        // mini stream at sector 3: "hello" in the first 64-byte mini sector
        d[4 * ss..4 * ss + 5].copy_from_slice(b"hello");
        d
    }

    #[test]
    fn parses_header_and_entries() {
        let d = tiny_cfb();
        let c = Cfb::parse(d).unwrap();
        assert_eq!(c.sector_size(), 512);
        assert_eq!(
            c.entries().iter().filter(|e| e.kind != EntryKind::Empty).count(),
            2
        );
        assert_eq!(c.entry("Tiny").unwrap().size, 5);
        assert_eq!(c.stream("Tiny").unwrap(), b"hello");
        assert!(matches!(c.stream("nope"), Err(Error::NoSuchStream(_))));
    }

    #[test]
    fn walks_paths() {
        let c = Cfb::parse(tiny_cfb()).unwrap();
        let paths: Vec<String> = c.paths().into_iter().map(|(p, _)| p).collect();
        assert_eq!(paths, vec!["Tiny".to_string()]);
    }

    #[test]
    fn rejects_bad_signature() {
        let mut d = tiny_cfb();
        d[0] = 0;
        assert!(matches!(Cfb::parse(d), Err(Error::BadSignature)));
    }
}

