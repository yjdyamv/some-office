//! # office-oleps — OLE property sets
//!
//! Word, Excel and PowerPoint all store their document metadata in two
//! property-set streams inside the compound file:
//!
//! * `\x05SummaryInformation` — title, subject, author, keywords, revision,
//!   application, timestamps (MS-OLEPS, "summary information" property set).
//! * `\x05DocumentSummaryInformation` — company, manager, page/word/character
//!   counts and the like.
//!
//! A property set is a small binary format: a 28-byte header naming the
//! sections, then each section as a table of `(id, offset)` pairs pointing at
//! typed values (a `VT_` variant tag followed by the data).
//!
//! We write every string as `VT_LPWSTR` with codepage 1200 (UTF-16), which is
//! what WPS's `samples/hello.doc` does and what Word itself does for
//! non-Latin-only metadata. That keeps the writer free of codepage guessing:
//! no string ever needs a code page other than UTF-16.

use std::collections::BTreeMap;

/// `VT_I2`, a 16-bit signed integer padded to 4 bytes.
const VT_I2: u32 = 0x0002;
/// `VT_I4`, a 32-bit signed integer.
const VT_I4: u32 = 0x0003;
/// `VT_LPSTR`, a counted CP1252 string.
const VT_LPSTR: u32 = 0x001E;
/// `VT_LPWSTR`, a counted UTF-16 string.
const VT_LPWSTR: u32 = 0x001F;
/// `VT_FILETIME`, 100-nanosecond intervals since 1601-01-01 UTC.
const VT_FILETIME: u32 = 0x0040;

/// The codepage that says "every string here is UTF-16".
pub const CODEPAGE_UTF16: i16 = 1200;

/// `FMTID_SummaryInformation`: the first section of `\x05SummaryInformation`.
pub const FMTID_SUMMARY: [u8; 16] = [
    0xE0, 0x85, 0x9F, 0xF2, 0xF9, 0x4F, 0x68, 0x10, 0xAB, 0x91, 0x08, 0x00, 0x2B, 0x27, 0xB3, 0xD9,
];

/// `FMTID_DocSummaryInformation`: the first section of
/// `\x05DocumentSummaryInformation`.
pub const FMTID_DOC_SUMMARY: [u8; 16] = [
    0x02, 0xD5, 0xCD, 0xD5, 0x9C, 0x2E, 0x1B, 0x10, 0x93, 0x97, 0x08, 0x00, 0x2B, 0x2C, 0xF9, 0xAE,
];

/// Stream name of the summary information property set.
pub const SUMMARY_STREAM: &str = "\u{5}SummaryInformation";
/// Stream name of the document summary information property set.
pub const DOC_SUMMARY_STREAM: &str = "\u{5}DocumentSummaryInformation";

/// One typed property value.
#[derive(Clone, Debug, PartialEq)]
pub enum Value {
    /// A 16-bit integer (`VT_I2`).
    I2(i16),
    /// A 32-bit integer (`VT_I4`).
    I4(i32),
    /// A string, stored as UTF-16 (`VT_LPWSTR`).
    Text(String),
    /// A timestamp in 100-ns intervals since 1601-01-01 (`VT_FILETIME`).
    Time(u64),
}

impl Value {
    fn encode(&self, out: &mut Vec<u8>) {
        match self {
            Value::I2(v) => {
                out.extend_from_slice(&VT_I2.to_le_bytes());
                out.extend_from_slice(&i32::from(*v).to_le_bytes());
            }
            Value::I4(v) => {
                out.extend_from_slice(&VT_I4.to_le_bytes());
                out.extend_from_slice(&v.to_le_bytes());
            }
            Value::Text(s) => {
                out.extend_from_slice(&VT_LPWSTR.to_le_bytes());
                let units: Vec<u16> = s.encode_utf16().collect();
                out.extend_from_slice(&((units.len() + 1) as u32).to_le_bytes());
                for unit in units {
                    out.extend_from_slice(&unit.to_le_bytes());
                }
                out.extend_from_slice(&0u16.to_le_bytes());
            }
            Value::Time(t) => {
                out.extend_from_slice(&VT_FILETIME.to_le_bytes());
                out.extend_from_slice(&(*t as u32).to_le_bytes());
                out.extend_from_slice(&((*t >> 32) as u32).to_le_bytes());
            }
        }
    }
}

/// One section of a property set: a `FMTID` and its properties by id.
#[derive(Clone, Debug, Default, PartialEq)]
pub struct Section {
    pub fmtid: [u8; 16],
    pub properties: BTreeMap<u32, Value>,
}

impl Section {
    /// An empty section carrying `fmtid` and the UTF-16 codepage property.
    pub fn new(fmtid: [u8; 16]) -> Self {
        let mut properties = BTreeMap::new();
        properties.insert(1, Value::I2(CODEPAGE_UTF16));
        Section { fmtid, properties }
    }

    /// Set one property, replacing any previous value.
    pub fn set(&mut self, id: u32, value: Value) -> &mut Self {
        self.properties.insert(id, value);
        self
    }

    /// Set a string property. An empty string is stored as such rather than
    /// dropped: Word treats a missing title and an empty one the same way, but
    /// keeping it makes the round trip lossless.
    pub fn text(&mut self, id: u32, text: impl Into<String>) -> &mut Self {
        self.set(id, Value::Text(text.into()))
    }

    fn encode(&self, out: &mut Vec<u8>) {
        // Section header: cbSection, cProperties, then the (id, offset) table.
        // Offsets are relative to the section start, so the values cannot be
        // written until the table's size is known.
        let count = self.properties.len() as u32;
        let table_len = 8 + count * 8;
        let mut values = Vec::new();
        let mut offsets: Vec<(u32, u32)> = Vec::with_capacity(count as usize);
        for (id, value) in &self.properties {
            offsets.push((*id, table_len + values.len() as u32));
            value.encode(&mut values);
        }
        let section_len = table_len + values.len() as u32;
        out.extend_from_slice(&section_len.to_le_bytes());
        out.extend_from_slice(&count.to_le_bytes());
        for (id, at) in offsets {
            out.extend_from_slice(&id.to_le_bytes());
            out.extend_from_slice(&at.to_le_bytes());
        }
        out.extend_from_slice(&values);
    }
}

/// A whole property-set stream: a header and its sections.
#[derive(Clone, Debug, Default, PartialEq)]
pub struct PropertySet {
    pub sections: Vec<Section>,
}

impl PropertySet {
    /// Serialise the stream. The system identifier is `0x0002000A`, the value
    /// Word and Excel write (MS-OLEPS names `0x00020006`, but every Office file
    /// we measured carries `0x0002000A`).
    pub fn to_bytes(&self) -> Vec<u8> {
        let mut out = Vec::new();
        out.extend_from_slice(&0xFFFEu16.to_le_bytes()); // byte order
        out.extend_from_slice(&0u16.to_le_bytes()); // version
        out.extend_from_slice(&0x0002_000Au32.to_le_bytes()); // system identifier
        out.extend_from_slice(&[0u8; 16]); // CLSID
        out.extend_from_slice(&(self.sections.len() as u32).to_le_bytes());
        // Section list: fmtid + offset, offsets relative to the stream start.
        let mut at = 28 + self.sections.len() * 20;
        let mut bodies = Vec::new();
        for section in &self.sections {
            out.extend_from_slice(&section.fmtid);
            out.extend_from_slice(&(at as u32).to_le_bytes());
            let mut body = Vec::new();
            section.encode(&mut body);
            at += body.len();
            bodies.extend_from_slice(&body);
        }
        out.extend_from_slice(&bodies);
        out
    }

    /// Parse a property-set stream.
    pub fn parse(data: &[u8]) -> Option<PropertySet> {
        let read_u16 = |at: usize| -> Option<u16> {
            Some(u16::from_le_bytes([*data.get(at)?, *data.get(at + 1)?]))
        };
        let read_u32 = |at: usize| -> Option<u32> {
            Some(u32::from_le_bytes([
                *data.get(at)?,
                *data.get(at + 1)?,
                *data.get(at + 2)?,
                *data.get(at + 3)?,
            ]))
        };
        if read_u16(0)? != 0xFFFE {
            return None;
        }
        let count = read_u32(24)? as usize;
        let mut sections = Vec::new();
        for i in 0..count {
            let entry = 28 + i * 20;
            let mut fmtid = [0u8; 16];
            fmtid.copy_from_slice(data.get(entry..entry + 16)?);
            let start = read_u32(entry + 16)? as usize;
            let pairs = read_u32(start + 4)? as usize;
            let mut properties = BTreeMap::new();
            for k in 0..pairs {
                let pair = start + 8 + k * 8;
                let id = read_u32(pair)?;
                let at = start + read_u32(pair + 4)? as usize;
                if let Some(value) = read_value(data, at) {
                    properties.insert(id, value);
                }
            }
            sections.push(Section { fmtid, properties });
        }
        Some(PropertySet { sections })
    }

    /// The section whose `fmtid` matches, if any.
    pub fn section(&self, fmtid: &[u8; 16]) -> Option<&Section> {
        self.sections.iter().find(|s| &s.fmtid == fmtid)
    }
}

fn read_value(data: &[u8], at: usize) -> Option<Value> {
    let tag = u32::from_le_bytes([
        *data.get(at)?,
        *data.get(at + 1)?,
        *data.get(at + 2)?,
        *data.get(at + 3)?,
    ]);
    let body = at + 4;
    match tag {
        VT_I2 => Some(Value::I2(i16::from_le_bytes([
            *data.get(body)?,
            *data.get(body + 1)?,
        ]))),
        VT_I4 => Some(Value::I4(i32::from_le_bytes([
            *data.get(body)?,
            *data.get(body + 1)?,
            *data.get(body + 2)?,
            *data.get(body + 3)?,
        ]))),
        VT_LPSTR => {
            let count = u32::from_le_bytes([
                *data.get(body)?,
                *data.get(body + 1)?,
                *data.get(body + 2)?,
                *data.get(body + 3)?,
            ]) as usize;
            let raw = data.get(body + 4..body + 4 + count)?;
            // The count includes the terminating NUL, and Office sometimes pads
            // with more of them: trim every trailing zero byte.
            let end = raw.iter().rposition(|b| *b != 0).map_or(0, |at| at + 1);
            Some(Value::Text(office_enc::decode_cp1252(&raw[..end])))
        }
        VT_LPWSTR => {
            let count = u32::from_le_bytes([
                *data.get(body)?,
                *data.get(body + 1)?,
                *data.get(body + 2)?,
                *data.get(body + 3)?,
            ]) as usize;
            let raw = data.get(body + 4..body + 4 + count * 2)?;
            // Same as the narrow case: the count covers the NUL terminator(s).
            let mut end = raw.len();
            while end >= 2 && raw[end - 1] == 0 && raw[end - 2] == 0 {
                end -= 2;
            }
            Some(Value::Text(office_enc::decode_utf16le(&raw[..end])))
        }
        VT_FILETIME => {
            let low = u32::from_le_bytes([
                *data.get(body)?,
                *data.get(body + 1)?,
                *data.get(body + 2)?,
                *data.get(body + 3)?,
            ]);
            let high = u32::from_le_bytes([
                *data.get(body + 4)?,
                *data.get(body + 5)?,
                *data.get(body + 6)?,
                *data.get(body + 7)?,
            ]);
            Some(Value::Time(u64::from(low) | (u64::from(high) << 32)))
        }
        _ => None,
    }
}

/// The metadata Word, Excel and PowerPoint share: what the File > Properties
/// dialog shows on its Summary tab.
#[derive(Clone, Debug, Default, PartialEq)]
pub struct Summary {
    pub title: String,
    pub subject: String,
    pub author: String,
    pub keywords: String,
    pub comments: String,
    pub last_author: String,
    pub revision: String,
    pub template: String,
    pub application: String,
    /// Creation time in 100-ns intervals since 1601, or 0 for "unset".
    pub created: u64,
    /// Last-saved time, same units.
    pub saved: u64,
    /// Pages, words and characters, as Word counts them.
    pub pages: i32,
    pub words: i32,
    pub characters: i32,
    pub doc_security: i32,
}

impl Summary {
    /// A summary with every field empty but the codepage.
    pub fn new() -> Self {
        Summary::default()
    }

    /// Build the property set, in the property-id order MS-OLEPS assigns.
    pub fn to_property_set(&self) -> PropertySet {
        let mut section = Section::new(FMTID_SUMMARY);
        section.text(2, &self.title);
        section.text(3, &self.subject);
        section.text(4, &self.author);
        section.text(5, &self.keywords);
        section.text(6, &self.comments);
        if !self.template.is_empty() {
            section.text(7, &self.template);
        }
        section.text(8, &self.last_author);
        if !self.revision.is_empty() {
            section.text(9, &self.revision);
        }
        if self.created != 0 {
            section.set(12, Value::Time(self.created));
        }
        if self.saved != 0 {
            section.set(13, Value::Time(self.saved));
        }
        if self.pages != 0 {
            section.set(14, Value::I4(self.pages));
        }
        if self.words != 0 {
            section.set(15, Value::I4(self.words));
        }
        if self.characters != 0 {
            section.set(16, Value::I4(self.characters));
        }
        section.text(18, &self.application);
        if self.doc_security != 0 {
            section.set(19, Value::I4(self.doc_security));
        }
        PropertySet { sections: vec![section] }
    }

    /// Serialised `\x05SummaryInformation` stream.
    pub fn to_bytes(&self) -> Vec<u8> {
        self.to_property_set().to_bytes()
    }

    /// Read a summary back out of a `\x05SummaryInformation` stream.
    pub fn from_bytes(data: &[u8]) -> Option<Self> {
        let set = PropertySet::parse(data)?;
        let section = set.section(&FMTID_SUMMARY)?;
        let text = |id: u32| -> String {
            match section.properties.get(&id) {
                Some(Value::Text(s)) => s.clone(),
                _ => String::new(),
            }
        };
        let int = |id: u32| -> i32 {
            match section.properties.get(&id) {
                Some(Value::I4(v)) => *v,
                Some(Value::I2(v)) => i32::from(*v),
                _ => 0,
            }
        };
        let time = |id: u32| -> u64 {
            match section.properties.get(&id) {
                Some(Value::Time(t)) => *t,
                _ => 0,
            }
        };
        Some(Summary {
            title: text(2),
            subject: text(3),
            author: text(4),
            keywords: text(5),
            comments: text(6),
            template: text(7),
            last_author: text(8),
            revision: text(9),
            created: time(12),
            saved: time(13),
            pages: int(14),
            words: int(15),
            characters: int(16),
            application: text(18),
            doc_security: int(19),
        })
    }
}

/// The metadata only the second stream carries: company and manager, plus the
/// counts each application fills differently.
#[derive(Clone, Debug, Default, PartialEq)]
pub struct DocSummary {
    pub category: String,
    pub manager: String,
    pub company: String,
    /// Excel: lines, paragraphs. Word: paragraphs. PowerPoint: slides, notes.
    pub lines: i32,
    pub paragraphs: i32,
    pub slides: i32,
    pub notes: i32,
}

impl DocSummary {
    pub fn to_property_set(&self) -> PropertySet {
        let mut section = Section::new(FMTID_DOC_SUMMARY);
        if !self.category.is_empty() {
            section.text(2, &self.category);
        }
        if self.lines != 0 {
            section.set(5, Value::I4(self.lines));
        }
        if self.paragraphs != 0 {
            section.set(6, Value::I4(self.paragraphs));
        }
        if self.slides != 0 {
            section.set(7, Value::I4(self.slides));
        }
        if self.notes != 0 {
            section.set(8, Value::I4(self.notes));
        }
        if !self.manager.is_empty() {
            section.text(14, &self.manager);
        }
        if !self.company.is_empty() {
            section.text(15, &self.company);
        }
        PropertySet { sections: vec![section] }
    }

    pub fn to_bytes(&self) -> Vec<u8> {
        self.to_property_set().to_bytes()
    }

    pub fn from_bytes(data: &[u8]) -> Option<Self> {
        let set = PropertySet::parse(data)?;
        let section = set.section(&FMTID_DOC_SUMMARY)?;
        let text = |id: u32| -> String {
            match section.properties.get(&id) {
                Some(Value::Text(s)) => s.clone(),
                _ => String::new(),
            }
        };
        let int = |id: u32| -> i32 {
            match section.properties.get(&id) {
                Some(Value::I4(v)) => *v,
                Some(Value::I2(v)) => i32::from(*v),
                _ => 0,
            }
        };
        Some(DocSummary {
            category: text(2),
            lines: int(5),
            paragraphs: int(6),
            slides: int(7),
            notes: int(8),
            manager: text(14),
            company: text(15),
        })
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn a_summary_round_trips() {
        let summary = Summary {
            title: "Quarterly report".into(),
            author: "Jamie Yates".into(),
            keywords: "budget, q3".into(),
            application: "some-office".into(),
            pages: 4,
            words: 1200,
            ..Summary::new()
        };
        let bytes = summary.to_bytes();
        assert_eq!(Summary::from_bytes(&bytes).as_ref(), Some(&summary));
    }

    #[test]
    fn non_ascii_text_survives() {
        let summary = Summary { title: "季度报告".into(), ..Summary::new() };
        let bytes = summary.to_bytes();
        let back = Summary::from_bytes(&bytes).unwrap();
        assert_eq!(back.title, "季度报告");
        // Stored as UTF-16 code units (4 characters + a NUL terminator), not
        // as UTF-8: the stream must contain the raw wide characters.
        let wide: Vec<u8> = "季度报告"
            .encode_utf16()
            .flat_map(|u| u.to_le_bytes())
            .collect();
        assert!(bytes.windows(wide.len()).any(|w| w == wide), "no UTF-16 title");
    }

    #[test]
    fn the_header_matches_what_word_writes() {
        let bytes = Summary::new().to_bytes();
        assert_eq!(&bytes[0..2], &[0xFE, 0xFF]);
        assert_eq!(&bytes[2..4], &[0, 0]);
        assert_eq!(&bytes[4..8], &[0x0A, 0x00, 0x02, 0x00]);
        assert_eq!(&bytes[24..28], &1u32.to_le_bytes());
        assert_eq!(&bytes[28..44], &FMTID_SUMMARY);
    }

    #[test]
    fn a_doc_summary_carries_company_and_manager() {
        let doc = DocSummary {
            company: "Example Ltd".into(),
            manager: "Alex Lee".into(),
            paragraphs: 7,
            ..DocSummary::default()
        };
        let bytes = doc.to_bytes();
        let back = DocSummary::from_bytes(&bytes).unwrap();
        assert_eq!(back.company, "Example Ltd");
        assert_eq!(back.manager, "Alex Lee");
        assert_eq!(back.paragraphs, 7);
    }

    #[test]
    fn a_reader_takes_the_cp1252_strings_office_writes() {
        // Word writes `VT_LPSTR` with a non-1200 codepage; the reader must
        // still decode it rather than return an empty string.
        let mut section = Section::new(FMTID_SUMMARY);
        section.set(1, Value::I2(1252));
        let mut out = Vec::new();
        out.extend_from_slice(&0xFFFEu16.to_le_bytes());
        out.extend_from_slice(&0u16.to_le_bytes());
        out.extend_from_slice(&0x0002_000Au32.to_le_bytes());
        out.extend_from_slice(&[0u8; 16]);
        out.extend_from_slice(&1u32.to_le_bytes());
        out.extend_from_slice(&FMTID_SUMMARY);
        out.extend_from_slice(&48u32.to_le_bytes());
        // Section: two properties, title as an old-style CP1252 string.
        let mut values = Vec::new();
        let title = b"caf\xE9\0";
        values.extend_from_slice(&VT_LPSTR.to_le_bytes());
        values.extend_from_slice(&(title.len() as u32).to_le_bytes());
        values.extend_from_slice(title);
        // Two properties: the table is 8 + 2x8 = 24 bytes, the codepage value
        // sits at 24 and the title at 32.
        out.extend_from_slice(&(24u32 + 8 + values.len() as u32).to_le_bytes());
        out.extend_from_slice(&2u32.to_le_bytes());
        out.extend_from_slice(&1u32.to_le_bytes());
        out.extend_from_slice(&24u32.to_le_bytes());
        out.extend_from_slice(&2u32.to_le_bytes());
        out.extend_from_slice(&32u32.to_le_bytes());
        // Codepage value, then the title.
        out.extend_from_slice(&VT_I2.to_le_bytes());
        out.extend_from_slice(&1252i32.to_le_bytes());
        out.extend_from_slice(&values);
        let back = Summary::from_bytes(&out).unwrap();
        assert_eq!(back.title, "café");
    }
}