//! Writing version-3 compound files.
//!
//! Only what a reader needs is produced: 512-byte sectors, one FAT, one mini
//! FAT, a single directory chain, and no DIFAT extension (so at most 109 FAT
//! sectors, roughly 7 MB of file). Streams smaller than the 4096-byte cutoff go
//! into the mini stream, as the format requires.
//!
//! The point of this module is not to be a general CFB writer but to prove the
//! reader's model: a file produced here must parse back to the same streams, and
//! the same file must also open in Word/Excel (`tools/roundtrip.ps1`).

const SECTOR: usize = 512;
const MINI: usize = 64;
const CUTOFF: usize = 4096;
const ENDOFCHAIN: u32 = 0xFFFF_FFFE;
const FREESECT: u32 = 0xFFFF_FFFF;
const FATSECT: u32 = 0xFFFF_FFFD;
const ENTRY: usize = 128;

/// One stream to write into a new compound file.
pub struct OutputStream<'a> {
    pub name: &'a str,
    pub data: &'a [u8],
}

impl<'a> OutputStream<'a> {
    pub fn new(name: &'a str, data: &'a [u8]) -> Self {
        OutputStream { name, data }
    }
}

/// Append `value` to `out` as little-endian bytes.
fn le32(out: &mut Vec<u8>, value: u32) {
    out.extend_from_slice(&value.to_le_bytes());
}

/// Build a compound file holding `streams` as direct children of the root.
pub fn write(streams: &[OutputStream<'_>]) -> Vec<u8> {
    let per_sector = SECTOR / 4;

    // ---- pack the mini stream -------------------------------------------------
    // Small streams live in 64-byte mini sectors inside the root entry's stream;
    // their chain lives in the mini FAT.
    let mut mini_bytes: Vec<u8> = Vec::new();
    let mut mini_fat: Vec<u32> = Vec::new();
    // (stream index, first mini sector)
    let mut mini_layout: Vec<(usize, u32)> = Vec::new();
    for (i, s) in streams.iter().enumerate() {
        if s.data.len() >= CUTOFF {
            continue;
        }
        let start = (mini_bytes.len() / MINI) as u32;
        mini_layout.push((i, start));
        let count = s.data.len().div_ceil(MINI).max(1);
        for k in 0..count {
            mini_fat.push(if k + 1 == count {
                ENDOFCHAIN
            } else {
                start + k as u32 + 1
            });
        }
        mini_bytes.extend_from_slice(s.data);
        // Pad this stream's last mini sector. Never `resize(count * MINI, 0)`:
        // that truncates everything written by earlier streams.
        while mini_bytes.len() % MINI != 0 {
            mini_bytes.push(0);
        }
    }
    if mini_bytes.is_empty() {
        // A file with no small stream still needs a readable mini stream: keep
        // one empty mini sector so the root entry is not dangling.
        mini_fat.push(ENDOFCHAIN);
        mini_bytes.resize(MINI, 0);
    }

    // ---- sector allocation ---------------------------------------------------
    // Order: big streams, mini stream, mini FAT, directory, FAT.
    let mut big_layout: Vec<(usize, u32, usize)> = Vec::new(); // (stream, start, count)
    let mut next_sector = 0u32;
    for (i, s) in streams.iter().enumerate() {
        if s.data.len() < CUTOFF {
            continue;
        }
        let count = s.data.len().div_ceil(SECTOR);
        big_layout.push((i, next_sector, count));
        next_sector += count as u32;
    }
    // The mini stream is padded to 64-byte mini sectors, so its sector count
    // must round *up*: a 256-byte mini stream still occupies one 512-byte sector.
    let mini_sectors = mini_bytes.len().div_ceil(SECTOR);
    let mini_start = next_sector;
    next_sector += mini_sectors as u32;
    let minifat_count = mini_fat.len().div_ceil(per_sector);
    let minifat_start = next_sector;
    next_sector += minifat_count as u32;

    let dir_entries = streams.len() + 2; // root + streams + one empty slot
    let dir_count = dir_entries.div_ceil(SECTOR / ENTRY);
    let dir_start = next_sector;
    next_sector += dir_count as u32;

    // The FAT must also describe itself, so grow until there is room.
    let mut fat_count = next_sector.div_ceil(per_sector as u32).max(1) as usize;
    while (next_sector as usize + fat_count) > fat_count * per_sector {
        fat_count += 1;
    }
    let fat_start = next_sector;

    // ---- FAT -----------------------------------------------------------------
    let mut fat = vec![FREESECT; fat_count * per_sector];
    let chain = |fat: &mut Vec<u32>, start: u32, count: usize| {
        for k in 0..count {
            fat[start as usize + k] = if k + 1 == count {
                ENDOFCHAIN
            } else {
                start + k as u32 + 1
            };
        }
    };
    for (_, start, count) in &big_layout {
        chain(&mut fat, *start, *count);
    }
    chain(&mut fat, mini_start, mini_sectors);
    chain(&mut fat, minifat_start, minifat_count);
    chain(&mut fat, dir_start, dir_count);
    for k in 0..fat_count {
        fat[fat_start as usize + k] = FATSECT;
    }

    // ---- directory -----------------------------------------------------------
    let mut dir: Vec<u8> = Vec::with_capacity(dir_count * SECTOR);
    let push_entry = |dir: &mut Vec<u8>,
                      name: &str,
                      kind: u8,
                      start: u32,
                      size: u64,
                      child: u32,
                      right: u32| {
        let mut e = [0u8; ENTRY];
        let utf16: Vec<u8> = name.encode_utf16().flat_map(|c| c.to_le_bytes()).collect();
        // The stored length includes the terminating NUL.
        let name_len = if utf16.is_empty() { 0 } else { utf16.len() + 2 };
        e[..utf16.len()].copy_from_slice(&utf16);
        e[0x40..0x42].copy_from_slice(&(name_len as u16).to_le_bytes());
        e[0x42] = kind;
        e[0x43] = 1; // black: a flat sibling chain is a valid red-black tree
        e[0x44..0x48].copy_from_slice(&FREESECT.to_le_bytes());
        e[0x48..0x4C].copy_from_slice(&right.to_le_bytes());
        e[0x4C..0x50].copy_from_slice(&child.to_le_bytes());
        e[0x74..0x78].copy_from_slice(&start.to_le_bytes());
        e[0x78..0x80].copy_from_slice(&size.to_le_bytes());
        dir.extend_from_slice(&e);
    };
    // Siblings form a red-black tree ordered by (name length, uppercase name);
    // a flat chain along `right` is ordered by the same key and needs no
    // rebalancing, which keeps the writer honest without a tree rotation pass.
    let mut siblings: Vec<(&str, u32, u64)> = Vec::with_capacity(streams.len());
    for (i, s) in streams.iter().enumerate() {
        match big_layout.iter().find(|(idx, _, _)| *idx == i) {
            Some((_, start, _)) => siblings.push((s.name, *start, s.data.len() as u64)),
            None => {
                let (_, start) = mini_layout.iter().find(|(idx, _)| *idx == i).copied().unwrap();
                siblings.push((s.name, start, s.data.len() as u64));
            }
        }
    }
    siblings.sort_by(|a, b| {
        let ka = (a.0.encode_utf16().count(), a.0.to_uppercase());
        let kb = (b.0.encode_utf16().count(), b.0.to_uppercase());
        ka.cmp(&kb)
    });

    let root_child = if siblings.is_empty() { FREESECT } else { 1 };
    push_entry(
        &mut dir,
        "Root Entry",
        5,
        mini_start,
        mini_bytes.len() as u64,
        root_child,
        FREESECT,
    );
    for (i, (name, start, size)) in siblings.iter().enumerate() {
        let right = if i + 1 < siblings.len() { (i + 2) as u32 } else { FREESECT };
        push_entry(&mut dir, name, 2, *start, *size, FREESECT, right);
    }
    // Unused slots must still be well formed: type 0 with every tree pointer
    // FREESECT. Zero-filled slots make strict readers walk into entry 0.
    while dir.len() < dir_count * SECTOR {
        let mut e = [0u8; ENTRY];
        e[0x44..0x48].copy_from_slice(&FREESECT.to_le_bytes());
        e[0x48..0x4C].copy_from_slice(&FREESECT.to_le_bytes());
        e[0x4C..0x50].copy_from_slice(&FREESECT.to_le_bytes());
        e[0x74..0x78].copy_from_slice(&FREESECT.to_le_bytes());
        dir.extend_from_slice(&e);
    }

    // ---- assemble ------------------------------------------------------------
    let total = (fat_start as usize + fat_count + 1) * SECTOR;
    let mut out = vec![0u8; total];
    let put = |out: &mut [u8], sector: u32, bytes: &[u8]| {
        let at = (sector as usize + 1) * SECTOR;
        out[at..at + bytes.len()].copy_from_slice(bytes);
    };
    for (i, start, _) in &big_layout {
        put(&mut out, *start, streams[*i].data);
    }
    put(&mut out, mini_start, &mini_bytes);
    let mut mf: Vec<u8> = Vec::with_capacity(minifat_count * SECTOR);
    for v in &mini_fat {
        le32(&mut mf, *v);
    }
    mf.resize(minifat_count * SECTOR, 0xFF);
    put(&mut out, minifat_start, &mf);
    put(&mut out, dir_start, &dir);
    let mut fat_bytes: Vec<u8> = Vec::with_capacity(fat_count * SECTOR);
    for v in &fat {
        le32(&mut fat_bytes, *v);
    }
    put(&mut out, fat_start, &fat_bytes);

    // ---- header --------------------------------------------------------------
    let h = &mut out[..SECTOR];
    h[0..8].copy_from_slice(&[0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1]);
    h[0x18..0x1A].copy_from_slice(&0x003Eu16.to_le_bytes()); // minor version
    h[0x1A..0x1C].copy_from_slice(&3u16.to_le_bytes()); // major version (512-byte sectors)
    h[0x1C..0x1E].copy_from_slice(&0xFFFEu16.to_le_bytes()); // little endian
    h[0x1E..0x20].copy_from_slice(&9u16.to_le_bytes()); // sector shift
    h[0x20..0x22].copy_from_slice(&6u16.to_le_bytes()); // mini sector shift
    h[0x2C..0x30].copy_from_slice(&(fat_count as u32).to_le_bytes());
    h[0x30..0x34].copy_from_slice(&dir_start.to_le_bytes());
    h[0x38..0x3C].copy_from_slice(&(CUTOFF as u32).to_le_bytes());
    h[0x3C..0x40].copy_from_slice(&minifat_start.to_le_bytes());
    h[0x40..0x44].copy_from_slice(&(minifat_count as u32).to_le_bytes());
    h[0x44..0x48].copy_from_slice(&ENDOFCHAIN.to_le_bytes()); // no DIFAT extension
    for i in 0..109 {
        let v = if i < fat_count {
            fat_start + i as u32
        } else {
            FREESECT
        };
        h[0x4C + i * 4..0x50 + i * 4].copy_from_slice(&v.to_le_bytes());
    }
    out
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::{Cfb, EntryKind};

    #[test]
    fn round_trips_through_the_reader() {
        let big: Vec<u8> = (0..5000u32).map(|i| (i % 251) as u8).collect();
        let small: Vec<u8> = (0..100u32).map(|i| (i * 3 % 256) as u8).collect();
        let tiny = b"abc".to_vec();
        let bytes = write(&[
            OutputStream::new("Big", &big),
            OutputStream::new("Small", &small),
            OutputStream::new("Tiny", &tiny),
        ]);

        let cfb = Cfb::parse(bytes).unwrap();
        assert_eq!(cfb.version(), 3);
        let names: Vec<&str> = cfb
            .entries()
            .iter()
            .filter(|e| e.kind == EntryKind::Stream)
            .map(|e| e.name.as_str())
            .collect();
        // Directory order is the CFB sort key: name length first, then the
        // uppercase name — so "Big"(3) < "Tiny"(4) < "Small"(5).
        assert_eq!(names, vec!["Big", "Tiny", "Small"]);
        assert_eq!(cfb.stream("Big").unwrap(), big);
        assert_eq!(cfb.stream("Small").unwrap(), small);
        assert_eq!(cfb.stream("Tiny").unwrap(), tiny);
    }

    #[test]
    fn writes_a_file_with_no_streams() {
        let bytes = write(&[]);
        let cfb = Cfb::parse(bytes).unwrap();
        assert_eq!(
            cfb.entries()
                .iter()
                .filter(|e| e.kind != EntryKind::Empty)
                .count(),
            1 // the root
        );
        assert!(cfb.stream("Nope").is_err());
    }

    #[test]
    fn physical_runs_point_at_the_written_bytes() {
        let payload: Vec<u8> = (0..200u32).map(|i| (i % 97) as u8).collect();
        let bytes = write(&[OutputStream::new("S", &payload)]);
        let cfb = Cfb::parse(bytes.clone()).unwrap();
        let runs = cfb.physical_runs("S").unwrap();
        let mut rebuilt = Vec::new();
        for (off, len) in runs {
            rebuilt.extend_from_slice(&bytes[off..off + len]);
        }
        assert_eq!(rebuilt, payload);
    }
}
