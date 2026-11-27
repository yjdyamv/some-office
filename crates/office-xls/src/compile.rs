//! Compile a formula *text* into a BIFF8 `rgce` token array.
//!
//! This is the inverse of [`crate::formula::decompile`]: `compile("A1+2*B2")`
//! produces the byte array Excel stores in a `FORMULA` record, so the writer
//! can emit real formulas instead of a frozen value.
//!
//! Supported: numbers, strings, booleans, cell references (`A1`, `$A$1`),
//! ranges (`A1:B2`), unary `-`/`+`, `%`, `^`, `*`, `/`, `+`, `-`, `&`,
//! comparisons (`=`, `<>`, `<`, `>`, `<=`, `>=`), parentheses and function
//! calls. Three-dimensional references (`Sheet2!A1`) need a `SUPBOOK` +
//! `EXTERNSHEET` pair, which the writer does not emit yet — they are rejected
//! with a clear error instead of being written as something Excel would refuse.

use std::fmt;

use crate::formula::func_name;

/// Why a formula could not be compiled.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Error {
    Unexpected(char),
    Unclosed,
    BadNumber(String),
    BadRef(String),
    UnknownFunction(String),
    TooManyArgs,
    ThreeDimensional,
    Trailing,
}

impl fmt::Display for Error {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Error::Unexpected(c) => write!(f, "unexpected character {c:?}"),
            Error::Unclosed => write!(f, "unbalanced parentheses or quote"),
            Error::BadNumber(s) => write!(f, "bad number {s:?}"),
            Error::BadRef(s) => write!(f, "bad reference {s:?}"),
            Error::UnknownFunction(s) => write!(f, "unknown function {s:?}"),
            Error::TooManyArgs => write!(f, "too many arguments"),
            Error::ThreeDimensional => write!(f, "sheet-qualified references are not supported yet"),
            Error::Trailing => write!(f, "trailing characters after the formula"),
        }
    }
}

impl std::error::Error for Error {}

/// `A1` → `0`, `Z1` → `25`, `AA1` → `26`, `XFD1` → `16383`.
fn column_of(letters: &str) -> Option<u16> {
    let mut col: u32 = 0;
    for ch in letters.chars() {
        let d = ch.to_ascii_uppercase() as u32;
        if !(b'A' as u32..=b'Z' as u32).contains(&d) {
            return None;
        }
        col = col * 26 + (d - b'A' as u32 + 1);
    }
    if col == 0 || col > 16384 {
        return None;
    }
    Some((col - 1) as u16)
}

/// Reverse lookup in the function table the decompiler uses.
pub fn func_id(name: &str) -> Option<u16> {
    (0..=368u16).find(|id| func_name(*id) == Some(name))
}

#[derive(Debug, Clone, PartialEq)]
enum Tok {
    Num(f64),
    Str(String),
    Bool(bool),
    /// `(row, col, flags)` — flags are the `fRwRel`/`fColRel` bits.
    Ref(u16, u16, u16),
    Area(u16, u16, u16, u16),
    Func(String),
    LParen,
    RParen,
    Comma,
    /// Binary operator, carrying its `ptg` id.
    Bin(u8),
    Percent,
}

const P_COL_REL: u16 = 0x4000;
const P_ROW_REL: u16 = 0x8000;

fn is_ref_char(c: char) -> bool {
    c.is_ascii_alphabetic() || c.is_ascii_digit() || c == '$'
}

fn tokenize(src: &str) -> Result<Vec<Tok>, Error> {
    let chars: Vec<char> = src.chars().collect();
    let mut out = Vec::new();
    let mut i = 0;
    while i < chars.len() {
        let c = chars[i];
        if c.is_whitespace() {
            i += 1;
            continue;
        }
        if c == '"' {
            i += 1;
            let mut s = String::new();
            loop {
                let Some(&d) = chars.get(i) else {
                    return Err(Error::Unclosed);
                };
                if d == '"' {
                    if chars.get(i + 1) == Some(&'"') {
                        s.push('"');
                        i += 2;
                        continue;
                    }
                    i += 1;
                    break;
                }
                s.push(d);
                i += 1;
            }
            out.push(Tok::Str(s));
            continue;
        }
        if c.is_ascii_digit() || (c == '.' && chars.get(i + 1).is_some_and(|d| d.is_ascii_digit())) {
            let start = i;
            while i < chars.len() && (chars[i].is_ascii_digit() || chars[i] == '.') {
                i += 1;
            }
            if i < chars.len() && (chars[i] == 'e' || chars[i] == 'E') {
                i += 1;
                if i < chars.len() && (chars[i] == '+' || chars[i] == '-') {
                    i += 1;
                }
                while i < chars.len() && chars[i].is_ascii_digit() {
                    i += 1;
                }
            }
            let text: String = chars[start..i].iter().collect();
            let n: f64 = text.parse().map_err(|_| Error::BadNumber(text.clone()))?;
            out.push(Tok::Num(n));
            continue;
        }
        if c == '$' || c.is_ascii_alphabetic() {
            let start = i;
            while i < chars.len() && is_ref_char(chars[i]) {
                i += 1;
            }
            let word: String = chars[start..i].iter().collect();
            // A function call: letters (and digits) followed by '('.
            if i < chars.len() && chars[i] == '(' && !word.contains('$') && word.chars().all(|d| d.is_ascii_alphabetic()) {
                out.push(Tok::Func(word.to_ascii_uppercase()));
                continue;
            }
            if i < chars.len() && chars[i] == '!' {
                return Err(Error::ThreeDimensional);
            }
            match parse_ref(&word) {
                Some((r1, c1, f1)) => {
                    if i < chars.len() && chars[i] == ':' {
                        i += 1;
                        let start2 = i;
                        while i < chars.len() && is_ref_char(chars[i]) {
                            i += 1;
                        }
                        let word2: String = chars[start2..i].iter().collect();
                        let Some((r2, c2, f2)) = parse_ref(&word2) else {
                            return Err(Error::BadRef(word2));
                        };
                        out.push(Tok::Area(r1, r2, c1 | f1, c2 | f2));
                    } else {
                        out.push(Tok::Ref(r1, c1, f1));
                    }
                }
                None => {
                    let upper = word.to_ascii_uppercase();
                    match upper.as_str() {
                        "TRUE" => out.push(Tok::Bool(true)),
                        "FALSE" => out.push(Tok::Bool(false)),
                        _ => return Err(Error::BadRef(word)),
                    }
                }
            }
            continue;
        }
        match c {
            '(' => out.push(Tok::LParen),
            ')' => out.push(Tok::RParen),
            ',' => out.push(Tok::Comma),
            '%' => out.push(Tok::Percent),
            '+' => out.push(Tok::Bin(0x03)),
            '-' => out.push(Tok::Bin(0x04)),
            '*' => out.push(Tok::Bin(0x05)),
            '/' => out.push(Tok::Bin(0x06)),
            '^' => out.push(Tok::Bin(0x07)),
            '&' => out.push(Tok::Bin(0x08)),
            '=' => out.push(Tok::Bin(0x0B)),
            '<' => {
                if chars.get(i + 1) == Some(&'>') {
                    i += 1;
                    out.push(Tok::Bin(0x0E));
                } else if chars.get(i + 1) == Some(&'=') {
                    i += 1;
                    out.push(Tok::Bin(0x0A));
                } else {
                    out.push(Tok::Bin(0x09));
                }
            }
            '>' => {
                if chars.get(i + 1) == Some(&'=') {
                    i += 1;
                    out.push(Tok::Bin(0x0C));
                } else {
                    out.push(Tok::Bin(0x0D));
                }
            }
            other => return Err(Error::Unexpected(other)),
        }
        i += 1;
    }
    Ok(out)
}

/// `$A$1` → `(0, 0, 0)`; `A1` → `(0, 0, 0xC000)`.
fn parse_ref(word: &str) -> Option<(u16, u16, u16)> {
    let mut chars = word.chars().peekable();
    let col_abs = chars.peek() == Some(&'$');
    if col_abs {
        chars.next();
    }
    let mut letters = String::new();
    while let Some(&c) = chars.peek() {
        if c.is_ascii_alphabetic() {
            letters.push(c);
            chars.next();
        } else {
            break;
        }
    }
    if letters.is_empty() || letters.len() > 3 {
        return None;
    }
    let row_abs = chars.peek() == Some(&'$');
    if row_abs {
        chars.next();
    }
    let mut digits = String::new();
    while let Some(&c) = chars.peek() {
        if c.is_ascii_digit() {
            digits.push(c);
            chars.next();
        } else {
            return None;
        }
    }
    if digits.is_empty() || chars.peek().is_some() {
        return None;
    }
    let row: u32 = digits.parse().ok()?;
    if row == 0 || row > 65536 {
        return None;
    }
    let col = column_of(&letters)?;
    if col > 255 {
        return None;
    }
    let mut flags = 0;
    if !col_abs {
        flags |= P_COL_REL;
    }
    if !row_abs {
        flags |= P_ROW_REL;
    }
    Some(((row - 1) as u16, col, flags))
}

struct Parser {
    toks: Vec<Tok>,
    pos: usize,
    out: Vec<u8>,
}

impl Parser {
    fn peek(&self) -> Option<&Tok> {
        self.toks.get(self.pos)
    }

    fn bump(&mut self) -> Option<Tok> {
        let t = self.toks.get(self.pos).cloned();
        if t.is_some() {
            self.pos += 1;
        }
        t
    }

    fn expr(&mut self, min_prec: u8) -> Result<(), Error> {
        self.unary()?;
        while let Some(Tok::Bin(ptg)) = self.peek().cloned() {
            let prec = precedence(ptg);
            if prec < min_prec {
                break;
            }
            self.bump();
            self.expr(prec + 1)?;
            self.out.push(ptg);
        }
        Ok(())
    }

    fn unary(&mut self) -> Result<(), Error> {
        match self.peek().cloned() {
            Some(Tok::Bin(0x04)) => {
                self.bump();
                self.unary()?;
                self.out.push(0x13); // ptgUminus
                Ok(())
            }
            Some(Tok::Bin(0x03)) => {
                self.bump();
                self.unary()
            }
            _ => self.primary(),
        }
    }

    fn primary(&mut self) -> Result<(), Error> {
        let Some(tok) = self.bump() else {
            return Err(Error::Unclosed);
        };
        match tok {
            Tok::Num(n) => {
                if (0.0..=65535.0).contains(&n) && n.fract() == 0.0 {
                    self.out.push(0x1E); // ptgInt
                    self.out.extend_from_slice(&(n as u16).to_le_bytes());
                } else {
                    self.out.push(0x1F); // ptgNum
                    self.out.extend_from_slice(&n.to_le_bytes());
                }
            }
            Tok::Str(s) => {
                self.out.push(0x17); // ptgStr
                let ascii = s.is_ascii();
                self.out.push(s.chars().count().min(255) as u8);
                self.out.push(if ascii { 0 } else { 1 });
                if ascii {
                    self.out.extend_from_slice(s.as_bytes());
                } else {
                    for ch in s.chars() {
                        self.out.extend_from_slice(&(ch as u16).to_le_bytes());
                    }
                }
            }
            Tok::Bool(b) => {
                self.out.push(0x1D); // ptgBool
                self.out.push(u8::from(b));
            }
            Tok::Ref(row, col, flags) => {
                self.out.push(0x44); // ptgRef
                self.out.extend_from_slice(&row.to_le_bytes());
                self.out.extend_from_slice(&(col | flags).to_le_bytes());
            }
            Tok::Area(r1, r2, c1, c2) => {
                self.out.push(0x25); // ptgArea
                self.out.extend_from_slice(&r1.to_le_bytes());
                self.out.extend_from_slice(&r2.to_le_bytes());
                self.out.extend_from_slice(&c1.to_le_bytes());
                self.out.extend_from_slice(&c2.to_le_bytes());
            }
            Tok::Func(name) => {
                let Some(id) = func_id(&name) else {
                    return Err(Error::UnknownFunction(name));
                };
                if self.bump() != Some(Tok::LParen) {
                    return Err(Error::Unclosed);
                }
                let mut args = 0u8;
                if self.peek() != Some(&Tok::RParen) {
                    loop {
                        self.expr(0)?;
                        args = args.checked_add(1).ok_or(Error::TooManyArgs)?;
                        match self.bump() {
                            Some(Tok::Comma) => continue,
                            Some(Tok::RParen) => break,
                            _ => return Err(Error::Unclosed),
                        }
                    }
                } else {
                    self.bump();
                }
                self.out.push(0x42); // ptgFuncVar
                self.out.push(args);
                self.out.extend_from_slice(&id.to_le_bytes());
            }
            Tok::LParen => {
                self.expr(0)?;
                if self.bump() != Some(Tok::RParen) {
                    return Err(Error::Unclosed);
                }
            }
            Tok::Percent => return Err(Error::Unexpected('%')),
            _ => return Err(Error::Unexpected('?')),
        }
        while self.peek() == Some(&Tok::Percent) {
            self.bump();
            self.out.push(0x14); // ptgPercent
        }
        Ok(())
    }
}

fn precedence(ptg: u8) -> u8 {
    match ptg {
        0x07 => 7,                       // ^
        0x05 | 0x06 => 6,                // * /
        0x03 | 0x04 => 5,                // + -
        0x08 => 4,                       // &
        _ => 3,                          // comparisons
    }
}

/// Compile `formula` (without the leading `=`) into a BIFF8 token array.
pub fn compile(formula: &str) -> Result<Vec<u8>, Error> {
    let src = formula.trim().strip_prefix('=').unwrap_or(formula.trim());
    let toks = tokenize(src)?;
    if toks.is_empty() {
        return Err(Error::Unclosed);
    }
    let mut p = Parser { toks, pos: 0, out: Vec::new() };
    p.expr(0)?;
    if p.pos != p.toks.len() {
        return Err(Error::Trailing);
    }
    Ok(p.out)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::formula::{decompile, Ctx};

    fn round(formula: &str) -> String {
        let rgce = compile(formula).unwrap();
        let ctx = Ctx { row: 0, col: 0, sheet: "", names: &[], extern_sheets: &[], biff8: true };
        decompile(&rgce, &ctx)
    }

    #[test]
    fn compiles_numbers_refs_and_operators() {
        assert_eq!(round("A1+A2"), "A1+A2");
        assert_eq!(round("A1*2"), "A1*2");
        assert_eq!(round("(A1+A2)*2"), "(A1+A2)*2");
        assert_eq!(round("$A$1+B$2"), "$A$1+B$2");
        assert_eq!(round("-A1"), "-A1");
        assert_eq!(round("A1&\"x\""), "A1&\"x\"");
        assert_eq!(round("A1<>B1"), "A1<>B1");
        assert_eq!(round("2^3"), "2^3");
    }

    #[test]
    fn compiles_ranges_and_functions() {
        assert_eq!(round("SUM(A1:A2)"), "SUM(A1:A2)");
        assert_eq!(round("SUM(A1:A2)*2"), "SUM(A1:A2)*2");
        assert_eq!(round("IF(A1>5,\"big\",\"small\")"), "IF(A1>5,\"big\",\"small\")");
        assert_eq!(round("ROUND(A1,1)"), "ROUND(A1,1)");
        assert_eq!(round("AVERAGE(A1:B2)"), "AVERAGE(A1:B2)");
    }

    #[test]
    fn matches_the_bytes_excel_wrote() {
        // A1+A2 as WPS wrote it in samples/formulas.xls.
        assert_eq!(compile("A1+A2").unwrap(), vec![0x44, 0x00, 0x00, 0x00, 0xC0, 0x44, 0x01, 0x00, 0x00, 0xC0, 0x03]);
        // SUM(A1:A2) as Excel wrote it in samples/formulas.xls (R3C2, cce 13).
        assert_eq!(
            compile("SUM(A1:A2)").unwrap(),
            vec![0x25, 0x00, 0x00, 0x01, 0x00, 0x00, 0xC0, 0x00, 0xC0, 0x42, 0x01, 0x04, 0x00]
        );
    }

    #[test]
    fn rejects_what_it_cannot_encode() {
        assert_eq!(compile("Sheet2!A1"), Err(Error::ThreeDimensional));
        assert_eq!(compile("FOO(A1)"), Err(Error::UnknownFunction("FOO".into())));
        assert_eq!(compile("A1+"), Err(Error::Unclosed));
        assert_eq!(compile("A1)"), Err(Error::Trailing));
        assert!(matches!(compile("1 @ 2"), Err(Error::Unexpected('@'))));
    }
}
