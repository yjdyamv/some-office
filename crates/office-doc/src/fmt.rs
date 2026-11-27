//! Character and paragraph formatting: the FKP pages and their property
//! modifiers (`sprm`s).
//!
//! Word does not store formatting next to the text. It stores *runs* of
//! characters in 512-byte "formatted disk pages" (FKP), each entry a list of
//! `sprm` opcodes to apply to the character properties accumulated so far.
//! A run with no `sprm`s inherits everything from the previous run, so a
//! reader has to walk the page in order and keep merging.
//!
//! Everything here is best effort: unknown `sprm`s are skipped by their
//! declared operand length (which is encoded in the opcode itself), so an
//! unknown property never desynchronises the walk.

/// Character properties (`CHPX`). Every field is `Option` because a `sprm`
/// only says "set this", never "this is the whole truth".
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct CharProps {
    pub bold: Option<bool>,
    pub italic: Option<bool>,
    /// `sprmCKul`: 0 none, 1 single, 2 word, 3 double, 4 dotted, 5 hidden,
    /// 6 thick, 7 dash, 8 dot-dash, 9 dot-dot-dash, 10 wave.
    pub underline: Option<u8>,
    pub strike: Option<bool>,
    pub double_strike: Option<bool>,
    pub size_half_points: Option<u16>,
    /// `sprmCCv`: a 0x00BBGGRR colour value.
    pub color: Option<u32>,
    /// `sprmCIco`: an index into Word's fixed 16-colour palette.
    pub color_index: Option<u8>,
    pub font_id: Option<u16>,
    /// `sprmCIss`: 0 normal, 1 superscript, 2 subscript.
    pub iss: Option<u8>,
    pub caps: Option<bool>,
    pub small_caps: Option<bool>,
    pub highlight: Option<u8>,
    pub style_id: Option<u16>,
    /// `sprmCPicLocation`: offset of the picture inside the `Data` stream.
    pub pic_location: Option<u32>,
    /// `sprmCFSpec`: the character is special (a picture mark, a footnote
    /// reference, ...) and is not laid out as text.
    pub spec: Option<bool>,
}

impl CharProps {
    /// Overlay `other` on top of `self`: a `sprm` set later in the same run
    /// wins, and a property the run never touches stays `None`.
    pub fn merge(&mut self, other: &CharProps) {
        macro_rules! take {
            ($($f:ident),*) => { $( if other.$f.is_some() { self.$f = other.$f; } )* };
        }
        take!(
            bold,
            italic,
            underline,
            strike,
            double_strike,
            size_half_points,
            color,
            color_index,
            font_id,
            iss,
            caps,
            small_caps,
            highlight,
            style_id
        );
    }
}

/// Paragraph properties (`PAPX`).
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct ParaProps {
    /// Index into the style sheet (`sprmPIstd` / the PAPX header).
    pub style_id: Option<u16>,
    /// `sprmPJc`: 0 left, 1 centre, 2 right, 3 both, 4 distribute.
    pub justify: Option<u8>,
    /// Twips.
    pub dxa_right: Option<i16>,
    pub dxa_left: Option<i16>,
    pub dxa_left_first: Option<i16>,
    pub dya_before: Option<u16>,
    pub dya_after: Option<u16>,
    /// `sprmPFInTable`: the paragraph sits inside a table cell.
    pub in_table: Option<bool>,
    /// `sprmPFTtp`: this paragraph mark ends a table **row**.
    pub ttp: Option<bool>,
    /// `sprmTDefTable` (`0xD608`): the row's cell boundaries in twips, one
    /// more entry than there are cells.
    pub col_widths: Vec<i16>,
    /// `sprmPIlvl`: list level.
    pub ilvl: Option<u8>,
    /// `sprmPIlfo`: list format override index.
    pub ilfo: Option<u16>,
    /// `sprmPOutLvl`: outline level.
    pub outline_level: Option<u8>,
}

impl ParaProps {
    pub fn merge(&mut self, other: &ParaProps) {
        macro_rules! take {
            ($($f:ident),*) => { $( if other.$f.is_some() { self.$f = other.$f; } )* };
        }
        if !other.col_widths.is_empty() {
            self.col_widths = other.col_widths.clone();
        }
        take!(
            style_id,
            justify,
            dxa_right,
            dxa_left,
            dxa_left_first,
            dya_before,
            dya_after,
            in_table,
            ttp,
            ilvl,
            ilfo,
            outline_level
        );
    }
}

/// One character run: a CP range plus the properties in force over it.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct CharRun {
    /// Byte offsets into the WordDocument stream, as stored in the FKP.
    pub fc_start: u32,
    pub fc_end: u32,
    pub props: CharProps,
}

/// One paragraph run (a PAPX entry).
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct ParaRun {
    pub fc_start: u32,
    pub fc_end: u32,
    pub props: ParaProps,
}

fn u16_at(b: &[u8], o: usize) -> Option<u16> {
    Some(u16::from_le_bytes([*b.get(o)?, *b.get(o + 1)?]))
}

fn u32_at(b: &[u8], o: usize) -> Option<u32> {
    Some(u32::from_le_bytes([
        *b.get(o)?,
        *b.get(o + 1)?,
        *b.get(o + 2)?,
        *b.get(o + 3)?,
    ]))
}

/// Operand size for a `sprm`, from the `spra` bits (13..=15).
///
/// | spra | operand                                  |
/// |------|------------------------------------------|
/// | 0,1  | 1 byte                                   |
/// | 2,4,5| 2 bytes                                  |
/// | 3    | 4 bytes                                  |
/// | 6    | 1 length byte, then that many bytes      |
/// | 7    | 3 bytes                                  |
fn operand_len(sprm: u16, data: &[u8], pos: usize) -> Option<usize> {
    match (sprm >> 13) & 0x7 {
        0 | 1 => Some(1),
        2 | 4 | 5 => Some(2),
        3 => Some(4),
        6 => Some(*data.get(pos)? as usize),
        7 => Some(3),
        _ => None,
    }
}

/// Walk a `grpprl` and hand every `(sprm, operand)` pair to `f`.
fn walk_sprms(grpprl: &[u8], mut f: impl FnMut(u16, &[u8])) {
    let mut i = 0usize;
    while i + 2 <= grpprl.len() {
        let sprm = u16::from_le_bytes([grpprl[i], grpprl[i + 1]]);
        i += 2;
        let n = match operand_len(sprm, grpprl, i) {
            Some(n) => n,
            None => break,
        };
        let start = if (sprm >> 13) & 0x7 == 6 { i + 1 } else { i };
        // Word sometimes declares a length longer than the bytes that follow
        // (the row-end `TDefTable` of samples/table.doc says 70 and ships 57).
        // Clamp instead of giving up: dropping the whole sprm loses the table.
        let end = (start + n).min(grpprl.len());
        if start >= grpprl.len() {
            break;
        }
        f(sprm, &grpprl[start..end]);
        i = end;
    }
}

fn bit(v: &[u8]) -> Option<bool> {
    // `ToggleOperand` (MS-DOC 2.6.1): 0 = off, 1 = on, 0x80 = leave unchanged,
    // 0x81 = toggle. A CHPX/PAPX entry carries the whole state of its run, so
    // "toggle" against the default (off) means on — Microsoft Word writes 0x81
    // for *every* bold run, and reading it as "off" loses all of them.
    match v.first() {
        None => None,
        Some(0x80) => None,
        Some(0) => Some(false),
        Some(_) => Some(true),
    }
}

/// Parse a `TDefTable` operand: `itcMac` (one byte), `itcMac + 1` cell
/// boundaries (`dxaCenter`, twips from the left margin), then `itcMac` 20-byte
/// `TC` structures.
///
/// Word writes one undocumented byte before `itcMac`, and it may declare more
/// bytes than it ships (the row-end `TDefTable` of `samples/table.doc` claims
/// 70 and provides 57), so the cell count is found by trying both offsets and
/// requiring the boundaries to be increasing:
/// `samples/table.doc` → `itcMac = 3`, `[-108, 2732, 5573, 8414]`.
fn tdef_table(v: &[u8]) -> Vec<i16> {
    for lead in 0..2usize {
        let Some(&itc) = v.get(lead) else { continue };
        let itc = itc as usize;
        if !(1..=63).contains(&itc) {
            continue;
        }
        let base = lead + 1;
        if base + (itc + 1) * 2 > v.len() {
            continue;
        }
        let bounds: Vec<i16> = (0..=itc)
            .map(|k| i16::from_le_bytes([v[base + k * 2], v[base + k * 2 + 1]]))
            .collect();
        if bounds.windows(2).all(|w| w[0] <= w[1]) {
            return bounds;
        }
    }
    Vec::new()
}

/// Apply character `sprm`s to `props`.
pub fn apply_char_sprms(props: &mut CharProps, grpprl: &[u8]) {
    walk_sprms(grpprl, |sprm, v| {
        let n2 = |v: &[u8]| u16::from_le_bytes([v[0], v[1]]);
        let n4 = |v: &[u8]| u32::from_le_bytes([v[0], v[1], v[2], v[3]]);
        match sprm {
            0x0835 => props.bold = bit(v),
            0x0836 => props.italic = bit(v),
            0x0837 => props.strike = bit(v),
            0x083B => props.caps = bit(v),
            0x083A => props.small_caps = bit(v),
            0x2A3E => props.underline = v.first().copied(),
            0x2A42 => props.color_index = v.first().copied(),
            0x6870 => props.color = u32_at(v, 0),
            0x4A43 => props.size_half_points = Some(n2(v)),
            0x4A4F => props.font_id = Some(n2(v)),
            0x2A48 => props.iss = v.first().copied(),
            0x0855 => props.spec = v.first().map(|b| *b != 0),
            0x2A0C => props.highlight = v.first().copied(),
            0x6A03 => props.pic_location = Some(n4(v)),
            0x2A53 => props.double_strike = bit(v),
            0x4A30 => props.style_id = Some(n2(v)),
            _ => {}
        }
    });
}

/// Apply paragraph `sprm`s to `props`.
pub fn apply_para_sprms(props: &mut ParaProps, grpprl: &[u8]) {
    walk_sprms(grpprl, |sprm, v| {
        let n2 = |v: &[u8]| u16::from_le_bytes([v[0], v[1]]);
        let s2 = |v: &[u8]| i16::from_le_bytes([v[0], v[1]]);
        match sprm {
            0x2403 => props.justify = v.first().copied(),
            0x840E => props.dxa_right = Some(s2(v)),
            0x840F => props.dxa_left = Some(s2(v)),
            0x8411 => props.dxa_left_first = Some(s2(v)),
            0xA413 => props.dya_before = Some(n2(v)),
            0xA414 => props.dya_after = Some(n2(v)),
            0x2416 => props.in_table = bit(v),
            0x2417 => props.ttp = bit(v),
            0xD608 => props.col_widths = tdef_table(v),
            0x260A => props.ilvl = v.first().copied(),
            0x460B => props.ilfo = Some(n2(v)),
            0x2640 => props.outline_level = v.first().copied(),
            0x4600 => props.style_id = Some(n2(v)),
            _ => {}
        }
    });
}

/// Parse a `PlcfBteChpx` / `PlcfBtePapx` (a PLC of FCs followed by page
/// numbers) into the WordDocument page numbers it points at.
pub fn bte_pages(plc: &[u8]) -> Vec<u32> {
    if plc.len() < 4 {
        return Vec::new();
    }
    let n = (plc.len() - 4) / 8;
    (0..n)
        .filter_map(|k| u32_at(plc, (n + 1) * 4 + k * 4))
        .collect()
}

/// Decode one 512-byte CHPX FKP page into character runs.
///
/// Layout: `crun+1` FCs, `crun` 1-byte offsets (×2 to get the byte position),
/// then `cb`-prefixed `grpprl`s. A run whose offset is 0 has no properties of
/// its own and inherits everything.
pub fn chpx_page(page: &[u8]) -> Vec<CharRun> {
    let mut out = Vec::new();
    if page.len() < 512 {
        return out;
    }
    let crun = page[511] as usize;
    if crun == 0 || (crun + 1) * 4 + crun > 511 {
        return out;
    }
    for i in 0..crun {
        let (Some(fc_start), Some(fc_end)) = (u32_at(page, i * 4), u32_at(page, (i + 1) * 4))
        else {
            break;
        };
        // Each CHPX entry is self-contained: Word's model is "style properties,
        // then this run's sprms". Runs do *not* accumulate, which is why a
        // non-bold paragraph following a bold one carries no "bold off" sprm.
        let mut props = CharProps::default();
        let off = page[(crun + 1) * 4 + i] as usize * 2;
        if off != 0 {
            if let Some(cb) = page.get(off) {
                let cb = *cb as usize;
                if let Some(grpprl) = page.get(off + 1..off + 1 + cb) {
                    apply_char_sprms(&mut props, grpprl);
                }
            }
        }
        out.push(CharRun {
            fc_start,
            fc_end,
            props,
        });
    }
    out
}

/// Decode one 512-byte PAPX FKP page into paragraph runs.
///
/// Same shape as CHPX, except each entry starts with `cb` (0 means "the real
/// length is the next byte ×2") and the payload begins with the `istd`.
pub fn papx_page(page: &[u8]) -> Vec<ParaRun> {
    let mut out = Vec::new();
    if page.len() < 512 {
        return out;
    }
    let crun = page[511] as usize;
    // Each run has a 13-byte `BX` (one `bOffset` byte plus a 12-byte `PHE`),
    // not a one-byte offset like a CHPX page.
    if crun == 0 || (crun + 1) * 4 + crun * 13 > 511 {
        return out;
    }
    for i in 0..crun {
        let (Some(fc_start), Some(fc_end)) = (u32_at(page, i * 4), u32_at(page, (i + 1) * 4))
        else {
            break;
        };
        // Reading `+ i` instead of `+ i * 13` here silently gives every run
        // the *first* run's properties — invisible in a one-paragraph test.
        let bo = page[(crun + 1) * 4 + i * 13];
        // A zero `bOffset` means "no PAPX for this run": the run keeps the
        // previous run's properties. Word relies on this to compress runs of
        // identically formatted paragraphs, so treating it as "default
        // properties" would silently drop the style of every paragraph after
        // the first. A run that *has* a PAPX starts from scratch — the entries
        // do not accumulate, exactly like CHPX.
        let mut props = if bo == 0 {
            out.last()
                .map(|r: &ParaRun| r.props.clone())
                .unwrap_or_default()
        } else {
            ParaProps::default()
        };
        let off = bo as usize * 2;
        if off != 0 {
            if let Some(cb) = page.get(off).copied() {
                let (payload_at, cb) = if cb == 0 {
                    (off + 2, page.get(off + 1).copied().unwrap_or(0) as usize * 2)
                } else {
                    (off + 1, cb as usize)
                };
                if let Some(body) = page.get(payload_at..payload_at + cb) {
                    if let Some(istd) = u16_at(body, 0) {
                        props.style_id = Some(istd);
                    }
                    apply_para_sprms(&mut props, body.get(2..).unwrap_or(&[]));
                }
            }
        }
        out.push(ParaRun {
            fc_start,
            fc_end,
            props,
        });
    }
    out
}

/// One entry of the style sheet: a name plus the properties the style itself
/// carries (`STD.grupe`, the `UPX` blocks).
///
/// The `UPX` array holds one block for the paragraph properties and, for
/// paragraph styles, a second for the character properties. Rather than trust
/// that order, every block is fed to both walkers — each only reacts to the
/// `sprm` ids it knows, so a misclassification cannot corrupt the result.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct StyleDef {
    /// The style's name; `None` when the file did not write a definition.
    pub name: Option<String>,
    /// Paragraph properties carried by the style (`sprmP*`).
    pub para: ParaProps,
    /// Character properties carried by the style (`sprmC*`).
    pub char: CharProps,
}

/// Does `grpprl` parse as a complete run of `sprm` + operand pairs?
fn sprms_cover(grpprl: &[u8]) -> bool {
    if grpprl.is_empty() {
        return false;
    }
    let mut i = 0usize;
    while i + 2 <= grpprl.len() {
        let sprm = u16::from_le_bytes([grpprl[i], grpprl[i + 1]]);
        i += 2;
        let Some(n) = operand_len(sprm, grpprl, i) else {
            return false;
        };
        let start = if (sprm >> 13) & 0x7 == 6 { i + 1 } else { i };
        let end = start + n;
        if end > grpprl.len() {
            return false;
        }
        i = end;
    }
    i == grpprl.len()
}

/// Read the `grupe` array of a `STD`: a sequence of `cbUPX u16` + that many
/// bytes of `grpprl`, running to the end of the record.
///
/// Word writes a `cupx` count in front of the array, and its `grpprl`s parse
/// exactly. WPS omits the count, pads the first length by two bytes, and leaves
/// one stray byte before the next block — so the scan also accepts a block whose
/// first two bytes are not part of the `sprm` run, and resynchronises by single
/// bytes when a length field is impossible. A block is only accepted when its
/// `sprm`s cover it exactly, which is what keeps the tolerance from inventing
/// properties.
pub fn style_upx(body: &[u8]) -> Vec<&[u8]> {
    let mut out = Vec::new();
    let mut pos = 0usize;
    let mut skipped = 0usize;
    while pos + 2 <= body.len() {
        let cb = u16::from_le_bytes([body[pos], body[pos + 1]]) as usize;
        if cb == 0 {
            pos += 2;
            continue;
        }
        if pos + 2 + cb <= body.len() {
            let whole = &body[pos + 2..pos + 2 + cb];
            if sprms_cover(whole) {
                out.push(whole);
                pos += 2 + cb;
                continue;
            }
            // WPS counts two padding bytes in `cbUPX`.
            if cb > 2 {
                let trimmed = &body[pos + 4..pos + 2 + cb];
                if sprms_cover(trimmed) {
                    out.push(trimmed);
                    pos += 2 + cb;
                    continue;
                }
            }
        }
        // Impossible length: slide one byte and try again, but only for a
        // handful of bytes — a long run of garbage means this is not `grupe`.
        if skipped >= 8 {
            break;
        }
        skipped += 1;
        pos += 1;
    }
    out
}

/// A paragraph as the document author wrote it: its text plus the character
/// runs that cover it.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct Paragraph {
    pub text: String,
    pub props: ParaProps,
    pub runs: Vec<CharSpan>,
    /// The paragraph style's name, resolved from the style sheet.
    pub style_name: Option<String>,
}

/// The formatting of a slice of a [`Paragraph`]'s text (character offsets).
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct CharSpan {
    pub start: usize,
    pub end: usize,
    pub props: CharProps,
}

impl CharSpan {
    /// The characters this span covers, taken from the paragraph's text.
    pub fn text<'a>(&self, paragraph: &'a str) -> &'a str {
        let byte = |want: usize| {
            paragraph
                .char_indices()
                .nth(want)
                .map(|(i, _)| i)
                .unwrap_or(paragraph.len())
        };
        &paragraph[byte(self.start)..byte(self.end)]
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn sprm_operand_lengths_are_derived_from_the_opcode() {
        // spra 0 -> 1 byte, spra 2 -> 2 bytes, spra 3 -> 4 bytes, spra 6 -> length byte
        assert_eq!(operand_len(0x0835, &[], 0), Some(1)); // sprmCFBold
        assert_eq!(operand_len(0x4A43, &[], 0), Some(2)); // sprmCHps
        assert_eq!(operand_len(0x6870, &[], 0), Some(4)); // sprmCCv
        assert_eq!(operand_len(0xC615, &[3, 1, 2, 3], 0), Some(3)); // length-prefixed
    }

    #[test]
    fn char_sprms_set_bold_size_and_colour() {
        let mut p = CharProps::default();
        // sprmCFBold=1, sprmCHps=36, sprmCCv=0x000000FF
        apply_char_sprms(&mut p, &[0x35, 0x08, 1, 0x43, 0x4A, 0x24, 0x00, 0x70, 0x68, 0xFF, 0, 0, 0]);
        assert_eq!(p.bold, Some(true));
        assert_eq!(p.size_half_points, Some(36));
        assert_eq!(p.color, Some(0xFF));
    }

    #[test]
    fn unknown_sprms_do_not_desync_the_walk() {
        let mut p = CharProps::default();
        // an invented 4-byte sprm, then sprmCFItalic
        apply_char_sprms(&mut p, &[0x00, 0x60, 1, 2, 3, 4, 0x36, 0x08, 1]);
        assert_eq!(p.italic, Some(true));
    }

    #[test]
    fn paragraph_sprms_read_justification_and_indent() {
        let mut p = ParaProps::default();
        // sprmPJc=1 (centre), sprmPDxaLeft=720, sprmPDxaLeft1=360
        apply_para_sprms(&mut p, &[0x03, 0x24, 1, 0x0F, 0x84, 0xD0, 0x02, 0x11, 0x84, 0x68, 0x01]);
        assert_eq!(p.justify, Some(1));
        assert_eq!(p.dxa_left, Some(720));
        assert_eq!(p.dxa_left_first, Some(360));
    }
}
