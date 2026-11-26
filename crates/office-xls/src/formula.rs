//! Formula decompiler: BIFF `rgce` (reverse-polish bytecode) → infix text.
//!
//! The tokens are the ones MS-XLS calls `Ptg` (2.5.198). Two encoding details
//! are worth spelling out because they trip up naive readers:
//!
//! * Bytes below `0x20` are plain tokens (`0x01` is `ptgExp`, `0x03` is
//!   `ptgAdd`, …). From `0x20` upwards the low five bits are the token and the
//!   two bits above them are the *class* (`0x20` reference, `0x40` value,
//!   `0x60` array), so `0x24` and `0x44` are both `ptgRef` — one returns a
//!   reference, the other its value.
//! * A BIFF8 `ptgRef` stores the **absolute** row and column plus two "is this
//!   relative" flags. The flags exist so a copied formula can be shifted; they
//!   are not offsets. Verified against `samples/formulas.xls`, where `=A1+A2`
//!   in B1 is `44 0000 00c0 44 0100 00c0 03`: rows 0 and 1, both flagged
//!   relative, columns 0 and 0 — absolute storage.

use office_enc::{decode_cp1252, decode_utf16le};

/// Everything the decompiler needs that is not in the token stream itself.
pub struct Ctx<'a> {
    /// Row/column of the cell that owns the formula (for BIFF5 relative refs).
    pub row: u16,
    pub col: u16,
    /// Name of the sheet the formula lives on, for same-sheet `3d` refs.
    pub sheet: &'a str,
    /// Workbook-level defined names, 1-based (`ptgName` indexes this).
    pub names: &'a [String],
    /// `EXTERNSHEET` table: one entry per `ixti`, already resolved to text.
    pub extern_sheets: &'a [String],
    /// BIFF8 stores strings as `cch u8, grbit u8`; BIFF5 as `cch u8`.
    pub biff8: bool,
}

/// Column index → `A`, `Z`, `AA`, …
pub fn col_name(col: u16) -> String {
    let mut n = col as u32;
    let mut out = Vec::new();
    loop {
        out.push((b'A' + (n % 26) as u8) as char);
        if n < 26 {
            break;
        }
        n = n / 26 - 1;
    }
    out.iter().rev().collect()
}

fn cell_name(row: u16, col: u16) -> String {
    format!("{}{}", col_name(col), row as u32 + 1)
}

fn abs_cell_name(row: u16, _col: u16, grbit: u16) -> String {
    let mut s = String::new();
    if grbit & 0x8000 == 0 {
        s.push('$');
    }
    s.push_str(&col_name(grbit & 0xFF));
    if grbit & 0x4000 == 0 {
        s.push('$');
    }
    s.push_str(&(row as u32 + 1).to_string());
    s
}

/// A stack entry: the rendered text plus its operator precedence, so that
/// `(A1+A2)*2` keeps its parentheses when it comes back out.
type Item = (String, u8);

const P_PRIMARY: u8 = 9;
const P_UNARY: u8 = 8;
const P_POW: u8 = 7;
const P_MUL: u8 = 6;
const P_ADD: u8 = 5;
const P_CONCAT: u8 = 4;
const P_CMP: u8 = 3;

fn wrap(item: &Item, min: u8) -> String {
    if item.1 < min {
        format!("({})", item.0)
    } else {
        item.0.clone()
    }
}

fn quote(s: &str) -> String {
    format!("\"{}\"", s.replace('"', "\"\""))
}

/// Error code → the `#DIV/0!` style text Excel shows.
fn error_text(code: u8) -> &'static str {
    match code {
        0x00 => "#NULL!",
        0x07 => "#DIV/0!",
        0x0F => "#VALUE!",
        0x17 => "#REF!",
        0x1D => "#NAME?",
        0x24 => "#NUM!",
        0x2A => "#N/A",
        _ => "#ERR!",
    }
}

/// Is this cell a follower of a shared formula (`ptgExp` to another cell)?
pub fn shared_master(rgce: &[u8]) -> Option<(u16, u16)> {
    if rgce.len() == 5 && rgce[0] == 0x01 {
        let row = u16::from_le_bytes([rgce[1], rgce[2]]);
        let col = u16::from_le_bytes([rgce[3], rgce[4]]);
        return Some((row, col));
    }
    None
}

/// Shift every relative reference in `rgce` by `(dr, dc)`.
///
/// Shared formulas store the master's bytecode once; followers differ only by
/// this offset. Because BIFF8 keeps absolute coordinates, shifting is a plain
/// add on the flagged parts.
pub fn shift_relative(rgce: &[u8], dr: i32, dc: i32) -> Vec<u8> {
    let mut out = rgce.to_vec();
    let mut i = 0usize;
    while i < out.len() {
        let b = out[i];
        let (len, refs): (usize, &[(usize, usize)]) = if b < 0x20 {
            match b {
                0x01 => (5, &[]),        // ptgExp
                0x02 => (6, &[]),        // ptgTbl
                0x17 => {
                    // ptgStr: length depends on the string itself
                    if i + 2 > out.len() {
                        break;
                    }
                    let cch = out[i + 1] as usize;
                    let wide = i + 2 < out.len() && out[i + 2] & 0x01 != 0;
                    (3 + cch * if wide { 2 } else { 1 }, &[])
                }
                0x19 => (4, &[]),        // ptgAttr
                0x1C | 0x1D => (2, &[]), // ptgErr / ptgBool
                0x1E => (3, &[]),        // ptgInt
                0x1F => (9, &[]),        // ptgNum
                _ => (1, &[]),           // operators
            }
        } else {
            match b & 0x1F {
                0x04 => (5, &[(1usize, 5usize)]), // ptgRef
                0x05 => (9, &[(1, 5), (5, 9)]),   // ptgArea
                0x0C => (5, &[(1, 5)]),           // ptgRefN
                0x0D => (9, &[(1, 5), (5, 9)]),   // ptgAreaN
                0x1A => (7, &[(3, 7)]),           // ptgRef3d
                0x1B => (11, &[(3, 7), (7, 11)]), // ptgArea3d
                0x00 => (8, &[]),                 // ptgArray
                0x01 => (3, &[]),                 // ptgFunc
                0x02 => (4, &[]),                 // ptgFuncVar
                0x03 => (5, &[]),                 // ptgName
                0x06 | 0x07 | 0x08 | 0x09 => {
                    let cce = match read_u16(&out, i + 3) {
                        Some(v) => v as usize,
                        None => break,
                    };
                    (5 + cce, &[]) // ptgMem*
                }
                0x0A => (5, &[]),  // ptgRefErr
                0x0B => (9, &[]),  // ptgAreaErr
                0x19 => (7, &[]),  // ptgNameX
                0x1C => (7, &[]),  // ptgRefErr3d
                0x1D => (11, &[]), // ptgAreaErr3d
                _ => (1, &[]),
            }
        };
        // `refs` lists (start, end) windows of the token: [start..start+2] is the
        // row, [end-2..end] is the grbit word holding column + relative flags.
        for (start, end) in refs {
            if i + end > out.len() {
                continue;
            }
            let row = u16::from_le_bytes([out[i + start], out[i + start + 1]]);
            let grbit = u16::from_le_bytes([out[i + end - 2], out[i + end - 1]]);
            if grbit & 0x8000 != 0 {
                let r = (row as i32 + dr).clamp(0, 65535) as u16;
                out[i + start] = (r & 0xFF) as u8;
                out[i + start + 1] = (r >> 8) as u8;
            }
            if grbit & 0x4000 != 0 {
                let c = ((grbit & 0xFF) as i32 + dc).clamp(0, 255) as u16;
                let g = (grbit & 0xFF00) | c;
                out[i + end - 2] = (g & 0xFF) as u8;
                out[i + end - 1] = (g >> 8) as u8;
            }
        }
        i += len.max(1);
    }
    out
}

/// `iftab` → function name. Unknown ids come back as `None` and render as
/// `FUNC(id)` so a wrong guess can never masquerade as a real function.
pub fn func_name(iftab: u16) -> Option<&'static str> {
    Some(match iftab {
        0 => "COUNT",
        1 => "IF",
        2 => "ISNA",
        3 => "ISERROR",
        4 => "SUM",
        5 => "AVERAGE",
        6 => "MIN",
        7 => "MAX",
        8 => "ROW",
        9 => "COLUMN",
        10 => "NA",
        11 => "NPV",
        12 => "STDEV",
        13 => "DOLLAR",
        14 => "FIXED",
        15 => "SIN",
        16 => "COS",
        17 => "TAN",
        18 => "ATAN",
        19 => "PI",
        20 => "SQRT",
        21 => "EXP",
        22 => "LN",
        23 => "LOG10",
        24 => "ABS",
        25 => "INT",
        26 => "SIGN",
        27 => "ROUND",
        28 => "LOOKUP",
        29 => "INDEX",
        30 => "REPT",
        31 => "MID",
        32 => "LEN",
        33 => "VALUE",
        34 => "TRUE",
        35 => "FALSE",
        36 => "AND",
        37 => "OR",
        38 => "NOT",
        39 => "MOD",
        40 => "DCOUNT",
        41 => "DSUM",
        42 => "DAVERAGE",
        43 => "DMIN",
        44 => "DMAX",
        45 => "DSTDEV",
        46 => "VAR",
        47 => "DVAR",
        48 => "TEXT",
        49 => "LINEST",
        50 => "TREND",
        51 => "LOGEST",
        52 => "GROWTH",
        56 => "PV",
        57 => "FV",
        58 => "NPER",
        59 => "PMT",
        60 => "RATE",
        61 => "MIRR",
        62 => "IRR",
        63 => "RAND",
        64 => "MATCH",
        65 => "DATE",
        66 => "TIME",
        67 => "DAY",
        68 => "MONTH",
        69 => "YEAR",
        70 => "WEEKDAY",
        71 => "HOUR",
        72 => "MINUTE",
        73 => "SECOND",
        74 => "NOW",
        75 => "AREAS",
        76 => "ROWS",
        77 => "COLUMNS",
        78 => "OFFSET",
        82 => "SEARCH",
        83 => "TRANSPOSE",
        86 => "TYPE",
        97 => "ATAN2",
        98 => "ASIN",
        99 => "ACOS",
        100 => "CHOOSE",
        101 => "HLOOKUP",
        102 => "VLOOKUP",
        105 => "ISREF",
        109 => "LOG",
        111 => "CHAR",
        112 => "LOWER",
        113 => "UPPER",
        114 => "PROPER",
        115 => "LEFT",
        116 => "RIGHT",
        117 => "EXACT",
        118 => "TRIM",
        119 => "REPLACE",
        120 => "SUBSTITUTE",
        121 => "CODE",
        124 => "FIND",
        125 => "CELL",
        126 => "ISERR",
        127 => "ISTEXT",
        128 => "ISNUMBER",
        129 => "ISBLANK",
        130 => "T",
        131 => "N",
        140 => "DATEVALUE",
        141 => "TIMEVALUE",
        142 => "SLN",
        143 => "SYD",
        144 => "DDB",
        148 => "INDIRECT",
        162 => "CLEAN",
        163 => "MDETERM",
        164 => "MINVERSE",
        165 => "MMULT",
        167 => "IPMT",
        168 => "PPMT",
        169 => "COUNTA",
        183 => "PRODUCT",
        184 => "FACT",
        189 => "DPRODUCT",
        190 => "ISNONTEXT",
        193 => "STDEVP",
        194 => "VARP",
        195 => "DSTDEVP",
        196 => "DVARP",
        197 => "TRUNC",
        198 => "ISLOGICAL",
        199 => "DCOUNTA",
        204 => "USDOLLAR",
        205 => "FINDB",
        206 => "SEARCHB",
        207 => "REPLACEB",
        208 => "LEFTB",
        209 => "RIGHTB",
        210 => "MIDB",
        211 => "LENB",
        212 => "ROUNDUP",
        213 => "ROUNDDOWN",
        214 => "ASC",
        215 => "DBCS",
        216 => "RANK",
        219 => "ADDRESS",
        220 => "DAYS360",
        221 => "TODAY",
        222 => "VDB",
        227 => "MEDIAN",
        228 => "SUMPRODUCT",
        229 => "SINH",
        230 => "COSH",
        231 => "TANH",
        232 => "ASINH",
        233 => "ACOSH",
        234 => "ATANH",
        235 => "DGET",
        244 => "INFO",
        247 => "DB",
        252 => "FREQUENCY",
        261 => "ERROR.TYPE",
        269 => "AVEDEV",
        270 => "BETADIST",
        271 => "GAMMALN",
        272 => "BETAINV",
        273 => "BINOMDIST",
        274 => "CHIDIST",
        275 => "CHIINV",
        276 => "COMBIN",
        277 => "CONFIDENCE",
        278 => "CRITBINOM",
        279 => "EVEN",
        280 => "EXPONDIST",
        281 => "FDIST",
        282 => "FINV",
        283 => "FISHER",
        284 => "FISHERINV",
        285 => "FLOOR",
        286 => "GAMMADIST",
        287 => "GAMMAINV",
        288 => "CEILING",
        289 => "HYPGEOMDIST",
        290 => "LOGNORMDIST",
        291 => "LOGINV",
        292 => "NEGBINOMDIST",
        293 => "NORMDIST",
        294 => "NORMSDIST",
        295 => "NORMINV",
        296 => "NORMSINV",
        297 => "STANDARDIZE",
        298 => "ODD",
        299 => "PERMUT",
        300 => "POISSON",
        301 => "TDIST",
        302 => "WEIBULL",
        303 => "SUMXMY2",
        304 => "SUMX2MY2",
        305 => "SUMX2PY2",
        306 => "CHITEST",
        307 => "CORREL",
        308 => "COVAR",
        309 => "FORECAST",
        310 => "FTEST",
        311 => "INTERCEPT",
        312 => "PEARSON",
        313 => "RSQ",
        314 => "STEYX",
        315 => "SLOPE",
        316 => "TTEST",
        317 => "PROB",
        318 => "DEVSQ",
        319 => "GEOMEAN",
        320 => "HARMEAN",
        321 => "SUMSQ",
        322 => "KURT",
        323 => "SKEW",
        324 => "ZTEST",
        325 => "LARGE",
        326 => "SMALL",
        327 => "QUARTILE",
        328 => "PERCENTILE",
        329 => "PERCENTRANK",
        330 => "MODE",
        331 => "TRIMMEAN",
        332 => "TINV",
        336 => "CONCATENATE",
        337 => "POWER",
        342 => "RADIANS",
        343 => "DEGREES",
        344 => "SUBTOTAL",
        345 => "SUMIF",
        346 => "COUNTIF",
        347 => "COUNTBLANK",
        350 => "ISPMT",
        351 => "DATEDIF",
        352 => "DATESTRING",
        353 => "NUMBERSTRING",
        354 => "ROMAN",
        358 => "GETPIVOTDATA",
        359 => "HYPERLINK",
        360 => "PHONETIC",
        361 => "AVERAGEA",
        362 => "MAXA",
        363 => "MINA",
        364 => "STDEVPA",
        365 => "VARPA",
        366 => "STDEVA",
        367 => "VARA",
        368 => "BAHTTEXT",
        _ => return None,
    })
}

/// Fixed arity for the functions that appear as `ptgFunc` (not `ptgFuncVar`).
/// `ptgFuncVar` carries its own argument count, so it needs nothing here.
///
/// Only the fixed-arity half of the `iftab` table is listed; everything else
/// falls back to one argument, which is what the overwhelming majority of
/// `ptgFunc` tokens carry. A wrong guess can only produce a wrongly grouped
/// call, never a crash.
fn fixed_arity(iftab: u16) -> usize {
    match iftab {
        // no arguments
        10 | 19 | 34 | 35 | 63 | 74 | 221 => 0,
        // two arguments
        27 | 39 | 97 | 117 | 212 | 213 | 337 => 2,
        // three arguments
        142 => 3,
        _ => 1,
    }
}

fn read_u16(d: &[u8], i: usize) -> Option<u16> {
    Some(u16::from_le_bytes([*d.get(i)?, *d.get(i + 1)?]))
}

/// One reference operand as `(row, col, relative-flags, byte length)`.
///
/// The two encodings differ in where the column and the "relative" flags live.
/// BIFF8 has a 16-bit column word whose top two bits are `fColRel`/`fRwRel`.
/// BIFF5 keeps the column in a single byte and puts the very same two flags in
/// the top of the **row** word — measured on a Microsoft-written Excel 95 file,
/// where `=A2` is `44 01 c0 00` (row word `0xC001` = flags `0xC000` | row 1).
fn read_ref(d: &[u8], i: usize, biff8: bool) -> Option<(u16, u16, u16, usize)> {
    if biff8 {
        let row = read_u16(d, i + 1)?;
        let col = read_u16(d, i + 3)?;
        Some((row, col & 0xFF, col & 0xC000, 5))
    } else {
        let word = read_u16(d, i + 1)?;
        let col = *d.get(i + 3)? as u16;
        Some((word & 0x3FFF, col, word & 0xC000, 4))
    }
}

fn read_u32(d: &[u8], i: usize) -> Option<u32> {
    Some(u32::from_le_bytes([
        *d.get(i)?,
        *d.get(i + 1)?,
        *d.get(i + 2)?,
        *d.get(i + 3)?,
    ]))
}

/// Decompile `rgce` into infix text. Never panics on truncated input: it stops
/// and returns what it has, prefixed with `…` when it had to bail out.
pub fn decompile(rgce: &[u8], ctx: &Ctx) -> String {
    let mut stack: Vec<Item> = Vec::new();
    let mut i = 0usize;
    let mut truncated = false;

    macro_rules! push {
        ($s:expr, $p:expr) => {
            stack.push(($s, $p))
        };
    }
    macro_rules! pop {
        () => {
            match stack.pop() {
                Some(v) => v,
                None => {
                    truncated = true;
                    ("#VALUE!".to_string(), P_PRIMARY)
                }
            }
        };
    }

    while i < rgce.len() {
        let b = rgce[i];
        // Plain tokens live in 0x00..=0x1F; class-carrying ones are moved to
        // 0x80..=0x9F so that `ptgExp` (0x01) and `ptgFunc` (0x21) cannot
        // collide in the same match.
        let id = if b < 0x20 { b } else { 0x80 | (b & 0x1F) };
        match id {
            // ---- constants ------------------------------------------------
            0x1E => {
                match read_u16(rgce, i + 1) {
                    Some(v) => push!(v.to_string(), P_PRIMARY),
                    None => {
                        truncated = true;
                        break;
                    }
                }
                i += 3;
            }
            0x1F => {
                if i + 9 > rgce.len() {
                    truncated = true;
                    break;
                }
                let v = f64::from_le_bytes(rgce[i + 1..i + 9].try_into().unwrap());
                push!(crate::numfmt::format_value(v, "General", false), P_PRIMARY);
                i += 9;
            }
            0x1D => {
                match rgce.get(i + 1) {
                    Some(v) => push!(
                        if *v != 0 { "TRUE".to_string() } else { "FALSE".to_string() },
                        P_PRIMARY
                    ),
                    None => {
                        truncated = true;
                        break;
                    }
                }
                i += 2;
            }
            0x1C => {
                match rgce.get(i + 1) {
                    Some(v) => push!(error_text(*v).to_string(), P_PRIMARY),
                    None => {
                        truncated = true;
                        break;
                    }
                }
                i += 2;
            }
            0x17 => {
                let cch = match rgce.get(i + 1) {
                    Some(v) => *v as usize,
                    None => {
                        truncated = true;
                        break;
                    }
                };
                let (grbit, start) = if ctx.biff8 {
                    match rgce.get(i + 2) {
                        Some(g) => (*g, i + 3),
                        None => {
                            truncated = true;
                            break;
                        }
                    }
                } else {
                    (0u8, i + 2)
                };
                let wide = ctx.biff8 && grbit & 0x01 != 0;
                let nbytes = cch * if wide { 2 } else { 1 };
                let body = match rgce.get(start..start + nbytes) {
                    Some(s) => s,
                    None => {
                        truncated = true;
                        break;
                    }
                };
                let s = if wide {
                    decode_utf16le(body)
                } else {
                    decode_cp1252(body)
                };
                push!(quote(&s), P_PRIMARY);
                i = start + nbytes;
            }
            0x16 => {
                // ptgMissArg: an omitted argument, e.g. the middle of IF(a,,b)
                push!(String::new(), P_PRIMARY);
                i += 1;
            }
            // ---- references -----------------------------------------------
            0x84 | 0x8C => {
                // ptgRef / ptgRefN
                let (rw, grbit, size) = match read_ref(rgce, i, ctx.biff8) {
                    Some((row, col, flags, size)) => (row, col | flags, size),
                    None => {
                        truncated = true;
                        break;
                    }
                };
                let (row, col) = if id == 0x8C {
                    (
                        (ctx.row as i32 + rw as i16 as i32).clamp(0, 65535) as u16,
                        (ctx.col as i32 + (grbit & 0xFF) as i32).clamp(0, 255) as u16,
                    )
                } else {
                    (rw, grbit & 0xFF)
                };
                let text = if id == 0x8C {
                    cell_name(row, col)
                } else {
                    abs_cell_name(row, col, grbit)
                };
                push!(text, P_PRIMARY);
                i += size;
            }
            0x85 | 0x8D => {
                // ptgArea / ptgAreaN
                let parsed = if ctx.biff8 {
                    let (r1, r2, g1, g2) = match (
                        read_u16(rgce, i + 1),
                        read_u16(rgce, i + 3),
                        read_u16(rgce, i + 5),
                        read_u16(rgce, i + 7),
                    ) {
                        (Some(a), Some(b), Some(c), Some(d)) => (a, b, c, d),
                        _ => (0, 0, 0, 0),
                    };
                    Some(((r1, g1 & 0xFF, g1 & 0xC000), (r2, g2 & 0xFF, g2 & 0xC000), 9))
                } else {
                    // BIFF5: two row words (flags in the high bits) then two
                    // one-byte columns.
                    match (
                        read_u16(rgce, i + 1),
                        read_u16(rgce, i + 3),
                        rgce.get(i + 5).copied(),
                        rgce.get(i + 6).copied(),
                    ) {
                        (Some(a), Some(b), Some(c), Some(d)) => Some((
                            (a & 0x3FFF, c as u16, a & 0xC000),
                            (b & 0x3FFF, d as u16, b & 0xC000),
                            7,
                        )),
                        _ => None,
                    }
                };
                let (a, b, size) = match parsed {
                    Some(v) => v,
                    None => {
                        truncated = true;
                        break;
                    }
                };
                let text = if id == 0x8D {
                    format!(
                        "{}:{}",
                        cell_name(
                            (ctx.row as i32 + a.0 as i16 as i32).clamp(0, 65535) as u16,
                            (ctx.col as i32 + a.1 as i32).clamp(0, 255) as u16
                        ),
                        cell_name(
                            (ctx.row as i32 + b.0 as i16 as i32).clamp(0, 65535) as u16,
                            (ctx.col as i32 + b.1 as i32).clamp(0, 255) as u16
                        )
                    )
                } else {
                    format!(
                        "{}:{}",
                        abs_cell_name(a.0, a.1, a.2),
                        abs_cell_name(b.0, b.1, b.2)
                    )
                };
                push!(text, P_PRIMARY);
                i += size;
            }
            0x9A => {
                // ptgRef3d: ixti, rw, grbit
                let (ixti, rw, grbit) = match (
                    read_u16(rgce, i + 1),
                    read_u16(rgce, i + 3),
                    read_u16(rgce, i + 5),
                ) {
                    (Some(a), Some(b), Some(c)) => (a, b, c),
                    _ => {
                        truncated = true;
                        break;
                    }
                };
                let sheet = ctx
                    .extern_sheets
                    .get(ixti as usize)
                    .map(|s| s.as_str())
                    .unwrap_or("");
                let prefix = if sheet.is_empty() {
                    String::new()
                } else if sheet.contains(' ') || sheet.contains('\'') {
                    format!("'{}'!", sheet.replace('\'', "''"))
                } else {
                    format!("{sheet}!")
                };
                push!(
                    format!("{prefix}{}", abs_cell_name(rw, grbit & 0xFF, grbit)),
                    P_PRIMARY
                );
                i += 7;
            }
            0x9B => {
                // ptgArea3d: ixti, rwFirst, rwLast, grbitFirst, grbitLast
                let (ixti, r1, r2, g1, g2) = match (
                    read_u16(rgce, i + 1),
                    read_u16(rgce, i + 3),
                    read_u16(rgce, i + 5),
                    read_u16(rgce, i + 7),
                    read_u16(rgce, i + 9),
                ) {
                    (Some(a), Some(b), Some(c), Some(d), Some(e)) => (a, b, c, d, e),
                    _ => {
                        truncated = true;
                        break;
                    }
                };
                let sheet = ctx
                    .extern_sheets
                    .get(ixti as usize)
                    .map(|s| s.as_str())
                    .unwrap_or("");
                let prefix = if sheet.is_empty() {
                    String::new()
                } else {
                    format!("{sheet}!")
                };
                push!(
                    format!(
                        "{prefix}{}:{}",
                        abs_cell_name(r1, g1 & 0xFF, g1),
                        abs_cell_name(r2, g2 & 0xFF, g2)
                    ),
                    P_PRIMARY
                );
                i += 11;
            }
            0x8A | 0x9C => {
                // ptgRefErr / ptgRefErr3d
                push!("#REF!".to_string(), P_PRIMARY);
                i += if id == 0x9C { 7 } else { 5 };
            }
            0x8B | 0x9D => {
                // ptgAreaErr / ptgAreaErr3d
                push!("#REF!".to_string(), P_PRIMARY);
                i += if id == 0x9D { 11 } else { 9 };
            }
            // ---- names ----------------------------------------------------
            0x83 => {
                // ptgName: u32 index into the NAME table (1-based)
                let idx = match read_u32(rgce, i + 1) {
                    Some(v) => v,
                    None => {
                        truncated = true;
                        break;
                    }
                };
                let name = ctx
                    .names
                    .get(idx.saturating_sub(1) as usize)
                    .cloned()
                    .unwrap_or_else(|| format!("#NAME{idx}"));
                push!(name, P_PRIMARY);
                i += 5;
            }
            0x99 => {
                // ptgNameX: ixti, u32 index
                let idx = match read_u32(rgce, i + 3) {
                    Some(v) => v,
                    None => {
                        truncated = true;
                        break;
                    }
                };
                let name = ctx
                    .names
                    .get(idx.saturating_sub(1) as usize)
                    .cloned()
                    .unwrap_or_else(|| format!("#NAME{idx}"));
                push!(name, P_PRIMARY);
                i += 7;
            }
            // ---- operators ------------------------------------------------
            0x03 | 0x04 | 0x05 | 0x06 | 0x07 | 0x08 | 0x09 | 0x0A | 0x0B | 0x0C | 0x0D | 0x0E
            | 0x0F | 0x10 | 0x11 => {
                // The plain-operator ids (0x03..0x11) live *below* 0x20, so they
                // only reach here when b < 0x20 — see the guard above.
                let (op, prec, right_assoc) = match id {
                    0x03 => ("+", P_ADD, false),
                    0x04 => ("-", P_ADD, false),
                    0x05 => ("*", P_MUL, false),
                    0x06 => ("/", P_MUL, false),
                    0x07 => ("^", P_POW, true),
                    0x08 => ("&", P_CONCAT, false),
                    0x09 => ("<", P_CMP, false),
                    0x0A => ("<=", P_CMP, false),
                    0x0B => ("=", P_CMP, false),
                    0x0C => (">=", P_CMP, false),
                    0x0D => (">", P_CMP, false),
                    0x0E => ("<>", P_CMP, false),
                    0x0F => (" ", P_MUL, false), // intersection
                    0x10 => (",", P_ADD, false), // union
                    0x11 => (":", P_POW, false), // range
                    _ => unreachable!(),
                };
                let right = pop!();
                let left = pop!();
                let l = if right_assoc {
                    wrap(&left, prec + 1)
                } else {
                    wrap(&left, prec)
                };
                let r = if right_assoc {
                    wrap(&right, prec)
                } else {
                    wrap(&right, prec + 1)
                };
                push!(format!("{l}{op}{r}"), prec);
                i += 1;
            }
            0x12 => {
                let v = pop!();
                push!(format!("+{}", wrap(&v, P_UNARY)), P_UNARY);
                i += 1;
            }
            0x13 => {
                let v = pop!();
                push!(format!("-{}", wrap(&v, P_UNARY)), P_UNARY);
                i += 1;
            }
            0x14 => {
                let v = pop!();
                push!(format!("{}%", wrap(&v, P_UNARY)), P_UNARY);
                i += 1;
            }
            0x15 => {
                let v = pop!();
                push!(format!("({})", v.0), P_PRIMARY);
                i += 1;
            }
            // ---- functions ------------------------------------------------
            0x81 => {
                // ptgFunc: fixed arity
                let iftab = match read_u16(rgce, i + 1) {
                    Some(v) => v,
                    None => {
                        truncated = true;
                        break;
                    }
                };
                let n = fixed_arity(iftab);
                let mut args = Vec::with_capacity(n);
                for _ in 0..n {
                    args.push(pop!().0);
                }
                args.reverse();
                let name = func_name(iftab)
                    .map(|s| s.to_string())
                    .unwrap_or_else(|| format!("FUNC({iftab})"));
                push!(format!("{name}({})", args.join(",")), P_PRIMARY);
                i += 3;
            }
            0x82 => {
                // ptgFuncVar: cparams, iftab
                let (nparams, iftab) = match (rgce.get(i + 1).copied(), read_u16(rgce, i + 2)) {
                    (Some(a), Some(b)) => (a, b),
                    _ => {
                        truncated = true;
                        break;
                    }
                };
                let mut args = Vec::with_capacity(nparams as usize);
                for _ in 0..nparams {
                    args.push(pop!().0);
                }
                args.reverse();
                let name = func_name(iftab)
                    .map(|s| s.to_string())
                    .unwrap_or_else(|| format!("FUNC({iftab})"));
                push!(format!("{name}({})", args.join(",")), P_PRIMARY);
                i += 4;
            }
            // ---- bookkeeping tokens ---------------------------------------
            0x19 => {
                // ptgAttr is four bytes in every file we have seen: the flag byte,
                // then the 16-bit operand. (The spec's 2-byte `grbit` would make
                // `19 02 0d 00` mean a 0x0d02 flag word with a truncated operand,
                // which cannot produce the correct `IF` jump distance of 13.)
                let grbit = match rgce.get(i + 1) {
                    Some(g) => *g,
                    None => {
                        truncated = true;
                        break;
                    }
                };
                let data = read_u16(rgce, i + 2).unwrap_or(0);
                if grbit & 0x04 != 0 {
                    // tAttrChoose: followed by (cOffset+1) u16 jump offsets
                    i += 4 + (data as usize + 1) * 2;
                } else if grbit & 0x10 != 0 {
                    // tAttrSum: the reference on the stack is wrapped in SUM.
                    let v = pop!();
                    push!(format!("SUM({})", v.0), P_PRIMARY);
                    i += 4;
                } else {
                    i += 4;
                }
            }
            0x86 | 0x87 | 0x88 | 0x89 => {
                // ptgMemArea / MemErr / MemNoMem / MemFunc: reserved u16, cce u16
                let cce = match read_u16(rgce, i + 3) {
                    Some(v) => v as usize,
                    None => {
                        truncated = true;
                        break;
                    }
                };
                i += 5 + cce;
            }
            0x80 => {
                // ptgArray: 7 bytes of reserved data, then the array constants
                push!("{…}".to_string(), P_PRIMARY);
                i += 8;
            }
            0x02 => {
                // ptgTbl (only reachable as the plain byte 0x02)
                push!("{TABLE}".to_string(), P_PRIMARY);
                i += 6;
            }
            0x01 => {
                // ptgExp — a shared-formula follower we could not resolve
                push!("{…}".to_string(), P_PRIMARY);
                i += 5;
            }
            _ => {
                // Unknown token: stop rather than mis-align the whole stream.
                truncated = true;
                break;
            }
        }
    }

    let text = stack
        .pop()
        .map(|v| v.0)
        .unwrap_or_else(|| String::new());
    if truncated && !text.is_empty() {
        format!("{text}…")
    } else {
        text
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn ctx() -> Ctx<'static> {
        Ctx {
            row: 0,
            col: 1,
            sheet: "F",
            names: &[],
            extern_sheets: &[],
            biff8: true,
        }
    }

    #[test]
    fn column_names() {
        assert_eq!(col_name(0), "A");
        assert_eq!(col_name(25), "Z");
        assert_eq!(col_name(26), "AA");
        assert_eq!(col_name(701), "ZZ");
        assert_eq!(col_name(702), "AAA");
    }

    /// Bytes lifted from samples/formulas.xls, one per formula, with the text
    /// Excel/WPS wrote into the cell.
    #[test]
    fn decompiles_the_sample_formulas() {
        let cases: &[(&str, &str)] = &[
            ("44000000c044010000c003", "A1+A2"),
            ("44000000c044010000c005", "A1*A2"),
            ("250000010000c000c042010400", "SUM(A1:A2)"),
            ("44000000c044010000c006", "A1/A2"),
            ("44000000c01e020007", "A1^2"),
            ("44000000c013", "-A1"),
            ("44000000c0170101780008", "A1&\"x\""),
            ("44000000c01e05000d", "A1>5"),
            ("44020000c01e0100411b00", "ROUND(A3,1)"),
            ("250000010000c000c0420104001e020005", "SUM(A1:A2)*2"),
            ("44000000c044010000c004411800", "ABS(A1-A2)"),
            ("44000000c044010000c00e", "A1<>A2"),
            ("24000000c024010000c042020400", "SUM(A1,A2)"),
        ];
        for (rgce, want) in cases {
            let bytes: Vec<u8> = (0..rgce.len() / 2)
                .map(|k| u8::from_str_radix(&rgce[k * 2..k * 2 + 2], 16).unwrap())
                .collect();
            assert_eq!(decompile(&bytes, &ctx()), *want, "rgce {rgce}");
        }
    }

    #[test]
    fn if_keeps_its_branches() {
        let bytes: Vec<u8> = (0.."44000000c01e05000d19020d001703016200690067001908140017050173006d0061006c006c001908030042030100".len() / 2)
            .map(|k| {
                u8::from_str_radix(
                    &"44000000c01e05000d19020d001703016200690067001908140017050173006d0061006c006c001908030042030100"
                        [k * 2..k * 2 + 2],
                    16,
                )
                .unwrap()
            })
            .collect();
        assert_eq!(decompile(&bytes, &ctx()), "IF(A1>5,\"big\",\"small\")");
    }

    #[test]
    fn parentheses_survive_the_round_trip() {
        // (A1+A2)*2
        let bytes: Vec<u8> = (0.."44000000c044010000c0031e020005".len() / 2)
            .map(|k| {
                u8::from_str_radix(&"44000000c044010000c0031e020005"[k * 2..k * 2 + 2], 16).unwrap()
            })
            .collect();
        assert_eq!(decompile(&bytes, &ctx()), "(A1+A2)*2");
    }

    #[test]
    fn shared_formula_shift_moves_relative_refs() {
        // =A1+A2 shifted down one row and right one column is =B2+B3
        let master: Vec<u8> = (0.."44000000c044010000c003".len() / 2)
            .map(|k| {
                u8::from_str_radix(&"44000000c044010000c003"[k * 2..k * 2 + 2], 16).unwrap()
            })
            .collect();
        let shifted = shift_relative(&master, 1, 1);
        let mut c = ctx();
        c.row = 1;
        c.col = 2;
        assert_eq!(decompile(&shifted, &c), "B2+B3");
    }

    #[test]
    fn shared_master_detection() {
        assert_eq!(shared_master(&[0x01, 0x05, 0x00, 0x02, 0x00]), Some((5, 2)));
        assert_eq!(shared_master(&[0x44, 0, 0, 0, 0xc0]), None);
    }
}
