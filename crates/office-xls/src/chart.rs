//! Chart sheet substream: a `BOF` record with `dt = 0x0020`.
//!
//! A chart carries almost no typed "type" field. MS-XLS 2.2.3.4 (Chart Group)
//! says the *chart group type* is implied by **which record is present**, so the
//! reader's job is "哪条记录出现了" rather than "读哪个枚举字段".
//!
//! | record | id | meaning |
//! |---|---|---|
//! | `Line` | `0x1018` | line chart |
//! | `Pie` | `0x1019` | pie chart (`pcDonut != 0` ⇒ doughnut) |
//! | `Bar` | `0x1017` | bar/column (`fTranspose != 0` ⇒ horizontal bar) |
//! | `Scatter` | `0x103A` | XY scatter |
//! | `Area` | `0x101A` | area chart |
//! | `SeriesText` | `0x100D` | one text: chart title / axis title / series name |
//! | title frame | `0x104F` | precedes the chart title's `SeriesText` |
//! | text descriptor | `0x1051` | precedes every `SeriesText`; longer than 8 bytes ⇒ series |

use crate::{Record, REC_EOF};
use office_enc::{decode_cp1252, decode_utf16le};

pub const REC_UNITS: u16 = 0x1001;
pub const REC_CHART: u16 = 0x1002;
pub const REC_SERIES: u16 = 0x1003;
pub const REC_SERIES_TEXT: u16 = 0x100D;
pub const REC_BAR: u16 = 0x1017;
pub const REC_LINE: u16 = 0x1018;
pub const REC_PIE: u16 = 0x1019;
pub const REC_AREA: u16 = 0x101A;
pub const REC_BEGIN: u16 = 0x1033;
pub const REC_END: u16 = 0x1034;
pub const REC_SCATTER: u16 = 0x103A;
pub const REC_TITLE_FRAME: u16 = 0x104F;
pub const REC_TEXT_DESC: u16 = 0x1051;
/// `BOF.dt` of a chart substream.
pub const BOF_DT_CHART: u16 = 0x0020;

/// Chart group type, implied by the records present in the substream.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ChartKind {
    Area,
    Bar,
    Column,
    Line,
    Pie,
    Doughnut,
    Scatter,
    Other,
}

impl ChartKind {
    pub fn name(self) -> &'static str {
        match self {
            ChartKind::Area => "area",
            ChartKind::Bar => "bar",
            ChartKind::Column => "column",
            ChartKind::Line => "line",
            ChartKind::Pie => "pie",
            ChartKind::Doughnut => "doughnut",
            ChartKind::Scatter => "scatter",
            ChartKind::Other => "other",
        }
    }
}

/// One chart, parsed from its substream.
#[derive(Debug, Clone, Default)]
pub struct Chart {
    pub kind: Option<ChartKind>,
    /// The chart title.
    pub title: Option<String>,
    /// Axis titles, in the order they appear (category axis first).
    pub axis_titles: Vec<String>,
    /// Series names.
    pub series: Vec<String>,
    /// Offset of this chart's `BOF` record in the workbook stream.
    pub offset: u32,
}

impl Chart {
    pub fn kind_name(&self) -> &'static str {
        self.kind.map(ChartKind::name).unwrap_or("other")
    }

    /// One-line summary, e.g. `line "CHART-TITLE-XLS" series=[Sales]`.
    pub fn summary(&self) -> String {
        let mut s = self.kind_name().to_string();
        if let Some(t) = &self.title {
            s.push_str(&format!(" {t:?}"));
        }
        if !self.series.is_empty() {
            s.push_str(&format!(" series={:?}", self.series));
        }
        if !self.axis_titles.is_empty() {
            s.push_str(&format!(" axes={:?}", self.axis_titles));
        }
        s
    }
}

/// `SeriesText` payload: `u16 link`, `cch u8`, `grbit u8`, chars.
fn text(payload: &[u8]) -> Option<String> {
    if payload.len() < 4 {
        return None;
    }
    let cch = payload[2] as usize;
    let high = payload[3] & 1 != 0;
    let body = payload.get(4..)?;
    if high {
        if body.len() < cch * 2 {
            return None;
        }
        Some(decode_utf16le(&body[..cch * 2]))
    } else {
        if body.len() < cch {
            return None;
        }
        Some(decode_cp1252(&body[..cch]))
    }
}

/// Parse the chart whose `BOF` is `recs[start]`.
///
/// Two discriminators are measured rather than assumed:
///
/// * **nesting depth** (via `Begin`/`End`) separates the chart title from the
///   axis titles: the chart title's `SeriesText` sits at depth ≤ 2, axis titles
///   sit deeper (inside the axis sub-block). The title frame record `0x104F`
///   appears before *every* text object, including empty ones, so its presence
///   alone says nothing.
/// * the `0x1051` descriptor's length separates series names from titles: a
///   series descriptor is longer than 8 bytes.
///
/// The kind is decided at the end by priority, because a scatter chart also
/// carries a `Bar` record (it inherits the bar-like grouping): `Scatter` >
/// `Line` > `Pie` > `Area` > `Bar`/`Column`.
pub fn parse(recs: &[Record<'_>], start: usize) -> Chart {
    let mut chart = Chart {
        offset: recs.get(start).map(|r| r.offset as u32).unwrap_or(0),
        ..Chart::default()
    };
    let mut depth = 0usize;
    let mut desc_len = 0usize;
    let mut bar_transpose = None;
    let (mut line, mut pie, mut donut, mut scatter, mut area) = (false, false, false, false, false);
    for rec in &recs[start..] {
        match rec.kind {
            REC_EOF => break,
            REC_BEGIN => depth += 1,
            REC_END => depth = depth.saturating_sub(1),
            REC_TEXT_DESC => desc_len = rec.data.len(),
            REC_SERIES_TEXT => {
                let Some(s) = text(rec.data) else { continue };
                if desc_len > 8 {
                    chart.series.push(s);
                } else if depth <= 2 {
                    chart.title.get_or_insert(s);
                } else {
                    chart.axis_titles.push(s);
                }
            }
            REC_BAR => bar_transpose = Some(rec.data.get(4).copied().unwrap_or(0) != 0),
            REC_LINE => line = true,
            REC_PIE => {
                pie = true;
                donut = rec.data.get(2).copied().unwrap_or(0) != 0;
            }
            REC_SCATTER => scatter = true,
            REC_AREA => area = true,
            _ => {}
        }
    }
    chart.kind = Some(if scatter {
        ChartKind::Scatter
    } else if line {
        ChartKind::Line
    } else if pie {
        if donut {
            ChartKind::Doughnut
        } else {
            ChartKind::Pie
        }
    } else if area {
        ChartKind::Area
    } else if let Some(transpose) = bar_transpose {
        if transpose {
            ChartKind::Bar
        } else {
            ChartKind::Column
        }
    } else {
        ChartKind::Other
    });
    chart
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn series_text_decodes_both_widths() {
        // "Sales" as UTF-16 (grbit bit0 set).
        let mut p = vec![0, 0, 5, 1];
        for ch in "Sales".encode_utf16() {
            p.extend_from_slice(&ch.to_le_bytes());
        }
        assert_eq!(text(&p).as_deref(), Some("Sales"));
        // 8-bit form.
        assert_eq!(text(&[0, 0, 3, 0, b'a', b'b', b'c']).as_deref(), Some("abc"));
        // Truncated.
        assert!(text(&[0, 0, 9, 0, b'a']).is_none());
    }

    #[test]
    fn bar_transpose_picks_bar_or_column() {
        let rec = |data: &'static [u8]| Record {
            offset: 0,
            kind: REC_BAR,
            data,
        };
        // fTranspose at offset 4: 1 = horizontal bar, 0 = column.
        let mut bar: Vec<Record> = vec![rec(&[0, 0, 0x96, 0, 1, 0])];
        bar.push(Record {
            offset: 0,
            kind: REC_EOF,
            data: &[],
        });
        assert_eq!(parse(&bar, 0).kind, Some(ChartKind::Bar));
        let mut col: Vec<Record> = vec![rec(&[0, 0, 0x96, 0, 0, 0])];
        col.push(Record {
            offset: 0,
            kind: REC_EOF,
            data: &[],
        });
        assert_eq!(parse(&col, 0).kind, Some(ChartKind::Column));
    }
}
