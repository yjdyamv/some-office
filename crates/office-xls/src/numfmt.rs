//! Excel number formats: the `FORMAT`/`XF` records that turn a bare `f64` into
//! what the grid shows (`2024-03-05`, `25.00%`, `¥1,234.57`).
//!
//! This is a *display* formatter, not a full XLSX-style format engine. It covers
//! the four sections (`pos;neg;zero;text`), dates and times, percentages,
//! thousands separators, decimal places and literal prefixes/suffixes such as
//! `$` or `¥`. Anything it does not understand degrades to the plain number
//! rather than to garbage.

/// Built-in `ifmt` values (MS-XLS 2.4.126 `IFmt`). Locale-dependent ones are the
/// English/`en-US` spellings; a workbook that cares about them ships its own
/// `FORMAT` record, which always wins.
pub fn builtin(ifmt: u16) -> Option<&'static str> {
    Some(match ifmt {
        0 => "General",
        1 => "0",
        2 => "0.00",
        3 => "#,##0",
        4 => "#,##0.00",
        5 => "$#,##0;($#,##0)",
        6 => "$#,##0;[Red]($#,##0)",
        7 => "$#,##0.00;($#,##0.00)",
        8 => "$#,##0.00;[Red]($#,##0.00)",
        9 => "0%",
        10 => "0.00%",
        11 => "0.00E+00",
        12 => "# ?/?",
        13 => "# ??/??",
        14 => "m/d/yyyy",
        15 => "d-mmm-yy",
        16 => "d-mmm",
        17 => "mmm-yy",
        18 => "h:mm AM/PM",
        19 => "h:mm:ss AM/PM",
        20 => "h:mm",
        21 => "h:mm:ss",
        22 => "m/d/yyyy h:mm",
        37 => "#,##0 ;(#,##0)",
        38 => "#,##0 ;[Red](#,##0)",
        39 => "#,##0.00;(#,##0.00)",
        40 => "#,##0.00;[Red](#,##0.00)",
        45 => "mm:ss",
        46 => "[h]:mm:ss",
        47 => "mm:ss.0",
        48 => "##0.0E+0",
        49 => "@",
        _ => return None,
    })
}

/// `true` if the code should be applied to the date part of a serial number.
pub fn is_date_code(code: &str) -> bool {
    let section = first_section(code);
    let mut chars = section.chars().peekable();
    while let Some(c) = chars.next() {
        match c {
            '"' => {
                for c in chars.by_ref() {
                    if c == '"' {
                        break;
                    }
                }
            }
            '\\' | '_' | '*' => {
                chars.next();
            }
            '[' => {
                // [Red], [$¥-804], [>=100] — but [h]/[m]/[s] are elapsed time.
                let mut inner = String::new();
                for c in chars.by_ref() {
                    if c == ']' {
                        break;
                    }
                    inner.push(c);
                }
                let lower = inner.to_ascii_lowercase();
                if matches!(lower.as_str(), "h" | "hh" | "m" | "mm" | "s" | "ss") {
                    return true;
                }
            }
            'y' | 'Y' | 'd' | 'D' | 'h' | 'H' | 's' | 'S' => return true,
            'm' | 'M' => return true,
            _ => {}
        }
    }
    false
}

/// The first `;`-separated section, ignoring separators inside quotes/brackets.
pub fn first_section(code: &str) -> &str {
    split_sections(code).into_iter().next().unwrap_or(code)
}

/// Split on `;` while respecting `"..."`, `[...]` and `\x` escapes.
pub fn split_sections(code: &str) -> Vec<&str> {
    let bytes = code.as_bytes();
    let mut out = Vec::new();
    let mut start = 0usize;
    let mut i = 0usize;
    while i < bytes.len() {
        match bytes[i] {
            b'"' => {
                i += 1;
                while i < bytes.len() && bytes[i] != b'"' {
                    i += 1;
                }
            }
            b'[' => {
                while i < bytes.len() && bytes[i] != b']' {
                    i += 1;
                }
            }
            b'\\' => i += 1,
            b';' => {
                out.push(&code[start..i]);
                start = i + 1;
            }
            _ => {}
        }
        i += 1;
    }
    out.push(&code[start..]);
    out
}

/// Render `value` the way the given format code would.
///
/// `date_1904` selects the Mac epoch; Excel's phantom 1900-02-29 is honoured for
/// serial 60 because files in the wild contain it.
pub fn format_value(value: f64, code: &str, date_1904: bool) -> String {
    let sections = split_sections(code);
    let code = code.trim();
    if code.is_empty() || code.eq_ignore_ascii_case("general") {
        return general(value);
    }
    // A single section applies to every sign; with two or more, pick by sign.
    let section = if sections.len() >= 2 && value < 0.0 {
        sections[1]
    } else if sections.len() >= 3 && value == 0.0 {
        sections[2]
    } else {
        sections[0]
    };
    let section = section.trim();
    if section.is_empty() {
        return general(value);
    }
    if is_date_code(section) {
        return format_date(value, section, date_1904);
    }
    if section.contains('@') {
        // Text section: nothing sensible for a number, show the number.
        return general(value);
    }
    format_number(value, section)
}

fn general(value: f64) -> String {
    if value.fract() == 0.0 && value.abs() < 1e15 {
        format!("{}", value as i64)
    } else {
        let s = format!("{value}");
        s
    }
}

/// Count `0`/`#` placeholders after the decimal point.
fn decimals(code: &str) -> usize {
    match code.split_once('.') {
        Some((_, frac)) => frac
            .chars()
            .take_while(|c| *c == '0' || *c == '#' || *c == '?')
            .count(),
        None => 0,
    }
}

fn format_number(value: f64, code: &str) -> String {
    // `(#,##0.00)` — the section supplies its own parentheses; strip them here
    // and re-add once at the end so they are not doubled.
    let mut code = code;
    let mut parens = false;
    if code.starts_with('(') && code.ends_with(')') {
        code = &code[1..code.len() - 1];
        parens = true;
    }
    let percent = code.contains('%');
    let v = if percent { value * 100.0 } else { value };
    let v = if v < 0.0 { -v } else { v };
    let d = decimals(code);
    let grouping = code.contains(',') && !code.contains(",,");
    let mut body = format!("{v:.d$}");
    if grouping {
        body = group_thousands(&body);
    }
    // Literal text: quoted strings, `\x` escapes, `$`, `¥`, `€`, `£`, and any
    // other run of non-placeholder characters.
    let (prefix, suffix) = literals(code);
    let mut out = String::new();
    out.push_str(&prefix);
    out.push_str(&body);
    if percent {
        out.push('%');
    }
    out.push_str(&suffix);
    if parens || (value < 0.0 && code.contains('(')) {
        out = format!("({out})");
    } else if value < 0.0 && !prefix.contains('-') {
        out.insert(0, '-');
    }
    out
}

/// Split a number code into literal text before and after the digit placeholders.
fn literals(code: &str) -> (String, String) {
    let mut prefix = String::new();
    let mut suffix = String::new();
    let mut in_digits = false;
    let mut chars = code.chars().peekable();
    while let Some(c) = chars.next() {
        match c {
            '"' => {
                let mut s = String::new();
                for c in chars.by_ref() {
                    if c == '"' {
                        break;
                    }
                    s.push(c);
                }
                if in_digits {
                    suffix.push_str(&s);
                } else {
                    prefix.push_str(&s);
                }
            }
            '\\' => {
                if let Some(c) = chars.next() {
                    if in_digits {
                        suffix.push(c);
                    } else {
                        prefix.push(c);
                    }
                }
            }
            '_' => {
                chars.next(); // width placeholder: render as a space
                if in_digits {
                    suffix.push(' ');
                } else {
                    prefix.push(' ');
                }
            }
            '*' => {
                chars.next(); // fill character: ignore
            }
            '[' => {
                // colour / condition / locale prefix — drop, except [$¥-804]
                let mut inner = String::new();
                for c in chars.by_ref() {
                    if c == ']' {
                        break;
                    }
                    inner.push(c);
                }
                if let Some(rest) = inner.strip_prefix('$') {
                    let sym: String = rest.chars().take_while(|c| *c != '-').collect();
                    prefix.push_str(&sym);
                }
            }
            '0' | '#' | '?' | '.' | ',' | '%' | 'E' | 'e' | '+' => {
                in_digits = true;
            }
            other => {
                if in_digits {
                    suffix.push(other);
                } else {
                    prefix.push(other);
                }
            }
        }
    }
    (prefix, suffix)
}

fn group_thousands(s: &str) -> String {
    let (int_part, frac) = match s.split_once('.') {
        Some((a, b)) => (a, Some(b)),
        None => (s, None),
    };
    let mut out = String::new();
    let digits: Vec<char> = int_part.chars().collect();
    for (i, c) in digits.iter().enumerate() {
        if i > 0 && (digits.len() - i) % 3 == 0 {
            out.push(',');
        }
        out.push(*c);
    }
    match frac {
        Some(f) => format!("{out}.{f}"),
        None => out,
    }
}

/// Days since 1970-01-01 → (year, month, day). Howard Hinnant's civil calendar.
fn civil_from_days(z: i64) -> (i32, u32, u32) {
    let z = z + 719_468;
    let era = if z >= 0 { z } else { z - 146_096 } / 146_097;
    let doe = (z - era * 146_097) as u64;
    let yoe = (doe - doe / 1460 + doe / 36_524 - doe / 146_096) / 365;
    let y = yoe as i64 + era * 400;
    let doy = doe - (365 * yoe + yoe / 4 - yoe / 100);
    let mp = (5 * doy + 2) / 153;
    let d = (doy - (153 * mp + 2) / 5 + 1) as u32;
    let m = if mp < 10 { mp + 3 } else { mp - 9 } as u32;
    ((if m <= 2 { y + 1 } else { y }) as i32, m, d)
}

/// Excel serial → (year, month, day, hour, minute, second).
pub fn serial_to_parts(serial: f64, date_1904: bool) -> (i32, u32, u32, u32, u32, u32) {
    let days = serial.floor();
    let frac = serial - days;
    // Round to the nearest second; 23:59:59.7 must not become 24:00:00.
    let secs = (frac * 86_400.0).round() as i64;
    let (secs, extra_day) = if secs >= 86_400 { (secs - 86_400, 1) } else { (secs, 0) };
    let days = days as i64 + extra_day;

    if !date_1904 && days == 60 {
        // Excel's phantom leap day: serial 60 is 1900-02-29, a date that never
        // existed. Files contain it (Lotus 1-2-3 heritage), so reproduce it.
        let h = (secs / 3600) as u32;
        let mi = ((secs % 3600) / 60) as u32;
        let s = (secs % 60) as u32;
        return (1900, 2, 29, h, mi, s);
    }
    // Day numbers are counted from 1970-01-01, so these epochs are negative.
    let epoch = if date_1904 {
        // 1904-01-01 is serial 0.
        -24_107
    } else if days >= 61 {
        // 1899-12-30, i.e. the day before the phantom 1900-02-29.
        -25_569
    } else {
        // 1899-12-31: serial 1 is 1900-01-01.
        -25_568
    };
    let (y, m, d) = civil_from_days(epoch + days);
    let h = (secs / 3600) as u32;
    let mi = ((secs % 3600) / 60) as u32;
    let s = (secs % 60) as u32;
    (y, m, d, h, mi, s)
}

const MONTHS: [&str; 12] = [
    "January", "February", "March", "April", "May", "June", "July", "August", "September",
    "October", "November", "December",
];
const DAYS: [&str; 7] = [
    "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday",
];

fn format_date(serial: f64, code: &str, date_1904: bool) -> String {
    let (y, mo, d, h24, mi, sec) = serial_to_parts(serial, date_1904);
    let days_total = serial.floor() as i64;
    // 1899-12-30 was a Saturday, so serial 0 is a Saturday too (1904: 1904-01-01
    // was a Friday, serial 0).
    let weekday = if date_1904 {
        ((days_total % 7 + 7) % 7 + 5) % 7
    } else {
        ((days_total % 7 + 7) % 7 + 6) % 7
    } as usize;

    let mut out = String::new();
    let chars: Vec<char> = code.chars().collect();
    let mut i = 0usize;
    let mut last_was_hour = false;
    let lower = code.to_ascii_lowercase();
    let has_ampm = lower.contains("am/pm") || lower.contains("a/p");
    let (h12, ampm) = if has_ampm {
        let am = h24 < 12;
        let h = match h24 % 12 {
            0 => 12,
            v => v,
        };
        (h, if am { "AM" } else { "PM" })
    } else {
        (h24, "")
    };

    while i < chars.len() {
        let c = chars[i];
        let run = |i: usize, ch: char| -> usize {
            let mut n = 0;
            while i + n < chars.len() && chars[i + n].eq_ignore_ascii_case(&ch) {
                n += 1;
            }
            n
        };
        match c.to_ascii_lowercase() {
            '"' => {
                i += 1;
                while i < chars.len() && chars[i] != '"' {
                    out.push(chars[i]);
                    i += 1;
                }
                i += 1;
                continue;
            }
            '\\' => {
                if i + 1 < chars.len() {
                    out.push(chars[i + 1]);
                }
                i += 2;
                continue;
            }
            '_' | '*' => {
                i += 2;
                continue;
            }
            '[' => {
                let mut inner = String::new();
                let mut j = i + 1;
                while j < chars.len() && chars[j] != ']' {
                    inner.push(chars[j]);
                    j += 1;
                }
                match inner.to_ascii_lowercase().as_str() {
                    "h" | "hh" => out.push_str(&format!("{}", serial * 24.0)),
                    "m" | "mm" => out.push_str(&format!("{}", serial * 1440.0)),
                    "s" | "ss" => out.push_str(&format!("{}", serial * 86_400.0)),
                    _ => {}
                }
                i = j + 1;
                continue;
            }
            'y' => {
                let n = run(i, 'y');
                if n >= 3 {
                    out.push_str(&format!("{y:04}"));
                } else {
                    out.push_str(&format!("{:02}", y % 100));
                }
                i += n;
                continue;
            }
            'd' => {
                let n = run(i, 'd');
                if n >= 4 {
                    out.push_str(DAYS[weekday]);
                } else if n == 3 {
                    out.push_str(&DAYS[weekday][..3]);
                } else if n == 2 {
                    out.push_str(&format!("{d:02}"));
                } else {
                    out.push_str(&format!("{d}"));
                }
                i += n;
                continue;
            }
            'h' => {
                let n = run(i, 'h');
                let v = if has_ampm { h12 } else { h24 };
                if n >= 2 {
                    out.push_str(&format!("{v:02}"));
                } else {
                    out.push_str(&format!("{v}"));
                }
                last_was_hour = true;
                i += n;
                continue;
            }
            's' => {
                let n = run(i, 's');
                if n >= 2 {
                    out.push_str(&format!("{sec:02}"));
                } else {
                    out.push_str(&format!("{sec}"));
                }
                i += n;
                continue;
            }
            'm' => {
                let n = run(i, 'm');
                // `m` is minutes when it follows an hour token or precedes `s`.
                let next_is_sec = chars.get(i + n).map(|c| c.eq_ignore_ascii_case(&'s')) == Some(true);
                let is_minute = last_was_hour || next_is_sec;
                if is_minute {
                    if n >= 2 {
                        out.push_str(&format!("{mi:02}"));
                    } else {
                        out.push_str(&format!("{mi}"));
                    }
                } else if n >= 4 {
                    out.push_str(MONTHS[(mo as usize - 1).min(11)]);
                } else if n == 3 {
                    out.push_str(&MONTHS[(mo as usize - 1).min(11)][..3]);
                } else if n == 2 {
                    out.push_str(&format!("{mo:02}"));
                } else {
                    out.push_str(&format!("{mo}"));
                }
                last_was_hour = false;
                i += n;
                continue;
            }
            'a' => {
                // AM/PM and A/P
                let rest: String = chars[i..].iter().take(5).collect::<String>().to_ascii_uppercase();
                if rest.starts_with("AM/PM") {
                    out.push_str(ampm);
                    i += 5;
                    continue;
                }
                if rest.starts_with("A/P") {
                    out.push(if ampm == "AM" { 'A' } else { 'P' });
                    i += 3;
                    continue;
                }
                out.push(c);
                i += 1;
                continue;
            }
            _ => {
                out.push(c);
                last_was_hour = false;
                i += 1;
                continue;
            }
        }
    }
    out
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn builtin_formats_exist() {
        assert_eq!(builtin(0), Some("General"));
        assert_eq!(builtin(49), Some("@"));
        assert_eq!(builtin(200), None);
    }

    #[test]
    fn date_detection() {
        assert!(is_date_code("yyyy-mm-dd"));
        assert!(is_date_code("m/d/yyyy"));
        assert!(is_date_code("[h]:mm:ss"));
        assert!(!is_date_code("#,##0.00"));
        assert!(!is_date_code("0.00%"));
        // `m` inside a quoted literal is not a month.
        assert!(!is_date_code("#,##0\" m\""));
    }

    #[test]
    fn serial_1900_and_1904() {
        assert_eq!(serial_to_parts(45356.0, false), (2024, 3, 5, 0, 0, 0));
        assert_eq!(serial_to_parts(1.0, false), (1900, 1, 1, 0, 0, 0));
        // Excel's phantom leap day, reproduced rather than silently shifted.
        assert_eq!(serial_to_parts(60.0, false), (1900, 2, 29, 0, 0, 0));
        assert_eq!(serial_to_parts(61.0, false), (1900, 3, 1, 0, 0, 0));
        assert_eq!(serial_to_parts(0.0, true), (1904, 1, 1, 0, 0, 0));
        // 12:00 exactly.
        assert_eq!(serial_to_parts(45356.5, false).3, 12);
    }

    #[test]
    fn formats_dates_percent_and_thousands() {
        assert_eq!(format_value(45356.0, "yyyy-mm-dd", false), "2024-03-05");
        assert_eq!(format_value(45356.0, "m/d/yyyy", false), "3/5/2024");
        assert_eq!(format_value(0.25, "0.00%", false), "25.00%");
        assert_eq!(format_value(1234.5678, "0.00", false), "1234.57");
        assert_eq!(format_value(1234.5678, "#,##0.00", false), "1,234.57");
        assert_eq!(format_value(1234567.0, "#,##0", false), "1,234,567");
        assert_eq!(format_value(1234.5, "\"$\"#,##0.00", false), "$1,234.50");
        assert_eq!(format_value(12.0, "General", false), "12");
    }

    #[test]
    fn negative_sections() {
        assert_eq!(format_value(-1234.5, "#,##0.00;(#,##0.00)", false), "(1,234.50)");
        assert_eq!(format_value(-42.0, "0", false), "-42");
    }
}
