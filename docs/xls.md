# .xls：Excel 97-2003 二进制工作表（BIFF8）

`.xls` 在 CFB 里只有一个有用的流：`Workbook`（Excel 5/95 叫 `Book`）。流的内容是一条
**记录流**（BIFF8），没有索引，只能从头顺序扫。本仓库实现：`crates/office-xls/src/lib.rs`。

## 1. 记录格式

每条记录 4 字节头 + 负载：

```
u16  record id     （小端，如 0x0809 = BOF）
u16  payload len
u8[] payload
```

记录可以跨 CONTINUE（0x003C）续接——字符串超过 8224 字节时就会被切开，SST 尤其常见。

## 2. 两条子流：globals + sheets

```
Workbook 流
  [BOF dt=0x0005 globals]
      CODEPAGE / DATEMODE / FONT / XF / STYLE / BOUNDSHEET / SST / EOF
  [BOF dt=0x0010 worksheet]  <- 偏移由 BOUNDSHEET 给出
      ROW / LABELSST / RK / NUMBER / FORMULA / ... / EOF
  [BOF dt=0x0010 worksheet]
      ...
```

`samples/hello.xls` 的实测布局（`Workbook` 流共 16137 字节）：

| 偏移 | 记录 | 说明 |
|---|---|---|
| 0x0000 | BOF len=16 | version 0x0600 (BIFF8), dt=0x0005 globals |
| 0x0014 | INTERFACEHDR | |
| 0x0024 | WRITEACCESS len=112 | 用户名，定长 |
| 0x0098 | CODEPAGE = 0x04B0 | 1200 = UTF-16，Excel 97 中文版常见 |
| 0x00F0 | DATEMODE = 0 | 0 = 1900 日期系统，1 = 1904 |
| 0x0108 | FONT ×23 len=20 | |
| 0x06C0 | XF ×66 len=20 | 单元格格式索引表 |
| 0x1A7A | STYLE + ... | |
| 0x39F7 | BOUNDSHEET len=20 | "Sheet2" @ 0x3AB6 |
| 0x3A0F | BOUNDSHEET len=16 | "Data" @ 0x3C72 |
| 0x3A53 | SST len=77 | 3 个字符串 |
| 0x3AA4 | EXTSST | SST 加速表（可忽略） |
| 0x3AB2 | EOF | globals 结束 |
| 0x3AB6 | BOF dt=0x0010 | 工作表 "Sheet2" |
| 0x3C72 | BOF dt=0x0010 | 工作表 "Data" |
| 0x3F05 | EOF | 文件结束 |

**BOUNDSHEET 的顺序就是标签页顺序**，与工作表在流里的物理顺序无关（本例标签页
"Sheet2" 在前、"Data" 在后，物理顺序相同，但 SOLVSAMP.XLS 里就不一定了）。

## 3. BOF

```
00 06        version（0x0600 = BIFF8，0x0500 = BIFF5/7）
05 00        dt：0x0005 globals，0x0010 worksheet，0x0020 chart，0x0040 macro
...          其余字段可跳过
```

判断"这是不是 BIFF8"就看第一个 BOF 的 version。BIFF5 的字符串是 8-bit + 代码页，
没有 SST（用 LABEL 0x0204 直接内联），本仓库只完整支持 BIFF8。

## 4. BOUNDSHEET（0x0085）

```
u32  lbPlyPos     该工作表 BOF 在流中的绝对偏移
u8   grbit        低 2 位：0 可见 / 1 隐藏 / 2 很深隐藏
u8   dt           0 工作表 / 1 Excel4 宏表 / 2 图表 / 6 VBA 模块
u8   cch          名字字符数
u8   grbitName    bit0 = 1 → 名字是 UTF-16，否则 8-bit
u8[cch]          名字
```

`hello.xls` 的第一条：`b6 3a 00 00 | 00 | 00 | 06 | 01 | 53 00 68 00 65 00 65 00 74 00 32 00`
→ 偏移 0x3AB6、可见、worksheet、6 字符、UTF-16、"Sheet2"。

## 5. SST（0x00FC）+ CONTINUE

**这是全格式最容易写错的地方：两个计数器各占 4 字节，不是 2 字节。**

```
u32 cstTotal    字符串总数（含重复）
u32 cstUnique   唯一字符串数
然后 cstUnique 个 XLUnicodeRichExtendedString：
  u16 cch
  u8  grbit    bit0 fHighByte(1=UTF-16) / bit2 fExtSt / bit3 fRichSt
  [u16 cRun]   仅 fRichSt
  [u32 cbExtRst] 仅 fExtSt
  cch 个字符（UTF-16LE 或 8-bit）
  [4*cRun]     格式索引
  [cbExtRst]   扩展数据
```

`hello.xls` 的 SST（0x3A53，77 字节）：

```
03 00 00 00   cstTotal = 3
03 00 00 00   cstUnique = 3
0c 00 01      cch=12, grbit=1 (UTF-16)
73 00 65 00 63 00 6f 00 6e 00 64 00 2d 00 73 00 68 00 65 00 65 00 74 00   "second-sheet"
0d 00 01      cch=13  "HELLO-XLS-001"
05 00 01      cch=5   "中文单元格"
```

注意：WPS 把纯 ASCII 字符串也写成 UTF-16（grbit=1）；Microsoft Excel 通常对纯 ASCII
用 8-bit（grbit=0）以省空间。两种都要处理。

读成 `u16` 时 `cstUnique` 会变成 0，于是整个字符串池为空、所有文本单元格显示空白——
而且不报错。本仓库第一版就踩了这个坑。

**CONTINUE 的陷阱**：字符串被切开时，新段开头会**重复一个 grbit 字节**表示该段用
8-bit 还是 UTF-16。所以不能简单地把所有 CONTINUE 负载拼接起来再解析，必须按段读取、
每段独立处理字符宽度。

## 6. 单元格记录

| id | 名称 | 负载 |
|---|---|---|
| 0x00FD | LABELSST | row u16, col u16, xf u16, isst u32 |
| 0x0204 | LABEL | row, col, xf, 字符串（BIFF8 里 XLUnicodeString） |
| 0x027E | RK | row, col, xf, rk u32 |
| 0x0203 | NUMBER | row, col, xf, f64 |
| 0x0205 | BOOLERR | row, col, xf, value u8, isError u8 |
| 0x0201 | BLANK | row, col, xf |
| 0x00BD | MULRK | row, colFirst, (xf u16, rk u32)×n, colLast |
| 0x00BE | MULBLANK | row, colFirst, xf×n, colLast |
| 0x0006 | FORMULA | row, col, xf, result 8B, grbit u16, chn u32, cce u16, rgce |
| 0x0207 | STRING | 字符串结果（紧跟在文本类 FORMULA 之后） |
| 0x0208 | ROW | 行属性（高度、隐藏等） |

行号列号都从 0 开始。`hello.xls` 的 "Data" 表（0x3C72 起）：

```
0x3E04 LABELSST  A1  isst=1  -> "HELLO-XLS-001"
0x3E12 MULRK     B1..B2     -> 1, 2
0x3E28 LABELSST  A2  isst=2  -> "中文单元格"
0x3E36 RK        A3          -> 1234.5
0x3E44 RK        C1          -> 45356 (日期 2024-03-05)
0x3E52 FORMULA   B3          -> 缓存值 3.0, cce=13
0x3E79 RK        A4          -> 42
0x3E87 RK        A5          -> -7
```

## 7. RK 解码

RK 用 32 位塞一个 double 或一个小整数：

```rust
fn decode_rk(rk: u32) -> f64 {
    let mut v = if rk & 0x02 != 0 {
        // 整数：高 30 位是有符号整数
        ((rk as i32) >> 2) as f64
    } else {
        // 浮点：这 32 位是 double 的高 32 位，低 32 位为 0
        f64::from_bits((rk as u64 & 0xFFFF_FFFC) << 32)
    };
    if rk & 0x01 != 0 { v /= 100.0; }
    v
}
```

实测：`0x40000000` → 2.0；`0x40934A00` → 1234.5；`0xC01C0000` → -7.0；
整数形式 `(0x00000005 << 2) | 2 = 0x16` → 1。

## 8. FORMULA（0x0006）：缓存值 + RPN

Excel 不保存公式文本，只保存**上次计算结果**和 RPN token 流（`rgce`）。
记录载荷的偏移（注意 `cce` 在 20，`rgce` 在 22——我第一次按 18/20 读，结果每个公式
都是空的）：

```text
0  row      u16
2  col      u16
4  xf       u16
6  result   8 B    数字是 f64；字符串/布尔/错误是特殊标记
14 grbit    u16    bit0 fAlwaysCalc, bit1 fCalcOnLoad, bit2 fShrFmla
16 chn      u32
20 cce      u16    rgce 字节数
22 rgce     cce B
```

结果若是字符串，缓存区前两字节是 `0xFFFF`，真值在下一条 `STRING`（0x0207）里。

### 8.1 token 流怎么读

每个 token 首字节是「类 + 操作码」：**bit5-6 是类**，低 5 位才是操作码。

| 类 | 值 | 含义 |
|---|---|---|
| 0x00 | `0x01..0x1F` | 基础 token（Exp / Add / Str / Attr / Func …） |
| 0x20 | `0x21..0x3F` | 引用类（Ref / Area / NameX / Ref3d） |
| 0x40 | `0x41..0x5F` | 值类（Func / FuncVar / Number / Str） |
| 0x60 | `0x61..0x7F` | 数组类 |

所以同一个「函数调用」可能是 `0x21`、`0x41`、`0x61`，取决于上下文。实现里把
`id = 0x80 | (b & 0x1F)` 做归一化，`0x01`（ptgExp）和 `0x21`（ptgFunc）就不会在
同一个 `match` 里撞车——这是本仓库 `formula.rs` 的核心技巧。

实现过的 token：

| token | 名称 | 载荷 | 说明 |
|---|---|---|---|
| `0x01` | ptgExp | row u16, col u16 | 共享公式的「从属」格子 |
| `0x03..0x08` | Add/Sub/Mul/Div/Pow/Concat | — | 二元运算，需要优先级表 |
| `0x0D/0x0E/0x0F..0x12` | GT/NE/LE/GE/LT/EQ | — | 比较 |
| `0x13` | Uminus | — | 一元负号 |
| `0x17` | ptgStr | cch u8, grbit u8, chars | grbit bit0 = UTF-16 |
| `0x19` | ptgAttr | grbit u8, data u16 | **4 字节**（见坑 4） |
| `0x1C/0x1D` | ptgErr / ptgBool | u8 | |
| `0x1E` | ptgInt | u16 | |
| `0x1F` | ptgNum | f64 | |
| `0x24/0x44/0x64` | ptgRef | row u16, col u16(+flags) | **存绝对行列**，见下 |
| `0x25` | ptgArea | 2×4 B | 同上 |
| `0x39/0x59/0x79` | ptgNameX | ixti u16, nameidx u32 | 跨表名字 |
| `0x3A` | ptgRef3d | ixti u16, row, col | `Sheet!A1` |
| `0x3B` | ptgArea3d | ixti u16, 2×4 B | `Sheet!A1:B2` |
| `0x41` | ptgFunc | iftab u16 | 固定参数个数 |
| `0x42` | ptgFuncVar | cparams u8, iftab u16 | 变参（SUM/IF/…） |

**ptgRef 存的是绝对行列 + 相对标志，不是相对偏移。** `=A1+A2` 在 B1 里的 rgce 是：

```text
44 00 00 00 c0    ptgRef row=0 col=0 grbit=0xC000（行相对|列相对）
44 01 00 00 c0    ptgRef row=1 col=0
03                Add
```

grbit 高位：`0x8000` 行绝对、`0x4000` 列绝对（`$A$1` 两个都置位）。共享公式（从属格子
的 `ptgExp`）才需要 `shift_relative()` 按偏移量改写行列。

**ptgAttr 只有 4 字节。** 规范写 `grbit u16 + data u16`，但真实文件里
`IF(A1>5,"big","small")` 的 attr 是 `19 02 0d 00`——只有按 `grbit u8 = 0x02`
（tAttrIf）、`data u16 = 13`（跳转长度）读，两个分支才能对上；按 5 字节读会立刻跑飞。
`grbit` 的位：`0x01` volatile、`0x02` if、`0x04` choose、`0x08` goto、`0x10` sum、
`0x20` assign。`0x10`（tAttrSum）表示「把栈顶区域求和」，也就是 `SUM(D17:F17)`
被压缩成 `ptgArea + tAttrSum`。

实测（`samples/formulas.xls`，18 个公式全部对上 WPS 写入的原文）：

```text
=A1+A2              =IF(A1>5,"big","small")     =SUM(G!A1,3)
=SUM(A1:A2)         =ROUND(A3,1)                =A1&"x"
=SUM(A1:A2)*2       =ABS(A1-A2)                 =A1<>A2
```

优先级：`^` 7 > 一元负号 8 > `*` `/` 6 > `+` `-` 5 > `&` 4 > 比较 3。生成文本时按
优先级决定是否加括号，否则 `(A1+A2)*2` 会变成 `A1+A2*2`。

## 9. 数字格式与日期

`FORMAT`（0x041E）存自定义格式串，`XF`（0x00E0）把格式串挂到单元格上：

```text
XF 载荷: 0 font u16 | 2 ifmt u16 | 4 之后是 alignment/border/pattern…
```

`ifmt < 164` 是内置格式（`164` 起才是自定义），但**WPS 连标准编号都写 FORMAT 记录**，
所以查表顺序必须是「先自定义表，再内置表」。

内置表里值得记的几个：0 `General`、1 `0`、2 `0.00`、9 `0%`、10 `0.00%`、
14 `m/d/yyyy`（区域相关）、22 `m/d/yyyy h:mm`、49 `@`（文本）。

日期序列号换算（`hello.xls` C1 = 45356 → 2024-03-05）：

```text
1900 系统: serial 1 = 1900-01-01，serial 60 = 1900-02-29（Excel 的幽灵闰日）
           serial >= 61 时 天数 = serial - 25569（1970-01-01 基准）
           serial <  61 时 天数 = serial - 25568
1904 系统: 天数 = serial - 24107
```

`%` 会把数值 ×100，`,#` 是千分位，负号段带括号时输出 `(1,234.50)`。本仓库实现了
`y/yy/yyyy`、`d/dd/ddd/dddd`、`m/mm/mmm/mmmm`（分钟与月份的歧义靠「前面是 h 或
后面是 s」判定）、`h/hh`、`s/ss`、`AM/PM`、`[h]`（累计小时），以及 `_`/`*` 占位跳过。


## 10. Rust API

```rust
use office_xls::{parse, CellValue};

let wb = parse(&std::fs::read("samples/hello.xls")?)?;
println!("biff={:#06X} sheets={}", wb.biff_version, wb.sheets.len());
for sheet in &wb.sheets {
    for row in sheet.rows() {
        for cell in row {
            match &cell.value {
                CellValue::Text(t) => print!("{t}\t"),
                CellValue::Number(n) => print!("{n}\t"),
                other => print!("{other}\t"),
            }
        }
        println!();
    }
}
```

公式文本在 `CellValue::Formula { text }` 里；`wb.formula_text(cell)` 直接取：

```rust
if let Some(f) = wb.formula_text(cell) {
    println!("= {f}");       // 例: SUM(A1:A2)*2
}
```

CLI：

```console
$ cargo run -q -p officedump -- text samples/hello.xls
$ cargo run -q -p officedump -- text samples/formulas.xls
$ cargo run -q -p officedump -- html samples/formulas.xls > /tmp/x.html
```

## 11. 加密文档

`FILEPASS`（`0x002F`）出现即表示加密，负载是 `wEncryptionType u16` + 加密头，
所以旧式 RC4 的 `lcb = 54`。整个 `Workbook` 流按 **1024 字节块**加密，但
**每条记录的 4 字节头保持明文**，BOF/FilePass/UsrExcl/FileLock/InterfaceHdr/RRDInfo/RRDHead
的负载和 `BoundSheet8` 的前 4 字节（`lbPlyPos`）也保持明文 —— 否则记录流无法遍历。

```rust
let wb = office_xls::parse_with_password(&bytes, "secret123")?;
```

细节见 [`encryption.md`](encryption.md)。

## 12. BIFF5（Excel 5.0 / 95）

`BOF` 的 `vers` 是 `0x0500` 而不是 `0x0600`，读取器据此切到「无 SST」的分支：

| 差异 | BIFF8（`0x0600`） | BIFF5（`0x0500`） |
|---|---|---|
| 字符串表 | `SST`（0x00FC）+ `LABEL` 存索引 | 没有 SST |
| `LABEL` 载荷 | `row, col, xf, XLUnicodeString(grbit+cch+chars)` | `row, col, xf, cch u16, 代码页字节` |
| `BOUNDSHEET` 名字 | `cch u8, grbit u8, chars` | `cch u8, chars`（无 grbit） |
| 公式 token | 同左 | 少部分 token 宽度不同（本仓库只解析公共子集） |
| `ptgRef` / `ptgArea` 引用编码 | `row u16` + `col u16`（相对标志在 col 高两位） | `row u16`（相对标志在**行字**高两位）+ `col u8` |

`LABEL` 的字符串按 `CODEPAGE` 记录解码（1252 / 936 …），`café` 这种 8-bit 字节
在 CP1252 下就是 `café`。

### 12.1 引用编码：BIFF5 把列压缩成一个字节

这是唯一一处**必须按版本分叉**的 token 宽度，实测自微软自己写的 Excel 95 文件
（`samples/ms/excel95.xls`、`biff5mix.xls`，见 `tools/ms_samples.ps1`）：

| 公式 | BIFF8 编码 | BIFF5 编码 |
|---|---|---|
| `=A2` | `44 01 00 c0 00`（5 B） | `44 01 c0 00`（4 B） |
| `=B1+C1` | `44 00 00 c0 01 44 00 00 c0 02 03` | `44 00 c0 01 44 00 c0 02 03` |
| `=SUM(A1:A3)` | `25 00 00 02 00 00 00 00 19 10 00 00` | `25 00 c0 02 c0 00 00 19 10 00 00` |

BIFF5 的 `row` 字里 `0x4000 = fColRel`、`0x8000 = fRwRel`（与 BIFF8 `col` 字的高两位
含义相同），剩下 14 位是行号（正好 16384 行）；列号只占 1 字节（256 列，BIFF5 的上限）。
所以 `=A2` 的 `01 c0` 是「行 1 + 两个相对标志」，不是行 `0xC001`。

```rust
// crates/office-xls/src/formula.rs
fn read_ref(d: &[u8], i: usize, biff8: bool) -> Option<(u16, u16, u16, usize)> {
    if biff8 {
        let row = read_u16(d, i + 1)?;
        let col = read_u16(d, i + 3)?;
        Some((row, col & 0xFF, col & 0xC000, 5))
    } else {
        let word = read_u16(d, i + 1)?;      // 行 + 相对标志
        let col = *d.get(i + 3)? as u16;     // 列只有一个字节
        Some((word & 0x3FFF, col, word & 0xC000, 4))
    }
}
```

按 BIFF8 去读 BIFF5 的 `44 01 c0 00` 会得到 `$A$49154` —— 不报错，只是全错，
所以 `decompiles_real_biff5_token_arrays` 这个回归测试把两种编码的结果都钉死了。

```rust
// BIFF5 分支的回归测试不依赖任何外部文件：直接喂一条手工拼出来的记录流。
let wb = office_xls::parse_stream(&biff5_stream)?;
assert_eq!(wb.biff_version, 0x0500);
```

`crates/office-xls` 的 `biff5_tests` 模块用 `rec(kind, body)` / `bof(vers, dt)` 两个小助手
拼出 globals + 一张工作表的完整流（`BOUNDSHEET` 的偏移依赖自身长度，先留 0 再回填），
覆盖了「无 SST 的 8-bit 字符串 + NUMBER + 代码页解码」。

真正的 BIFF5 文件由微软 Excel 写出（`samples/ms/excel95.xls`，`SaveAs` 的第二参数 `39`
= `xlExcel9795`）——WPS 会忽略这个参数、照样写 BIFF8，所以只有微软的实现能给出
BIFF5 的真值。

## 13. 图表：Chart Sheet Substream

图表**没有「类型字段」**。MS-XLS 2.2.3.4（Chart Group）规定：类型由**出现了哪条记录**决定，
读取器要做的是「看记录」而不是「读枚举」——这和 `.doc`/`.ppt` 里到处是显式字段的习惯正好相反。

| 类型 | 判据（记录） | 说明 |
|---|---|---|
| 折线 `line` | `Line`（`0x1018`，2 字节） | 出现即折线 |
| 饼图 `pie` | `Pie`（`0x1019`） | 偏移 2 的 `pcDonut` = 0 |
| 圆环 `doughnut` | 同上 | `pcDonut` ≠ 0 |
| 条形 `bar` | `Bar`（`0x1017`） | 偏移 4 的 `fTranspose` = 1（水平） |
| 柱形 `column` | 同上 | `fTranspose` = 0（垂直） |
| 散点 `scatter` | `Scatter`（`0x103A`） | |
| 面积 `area` | `Area`（`0x101A`） | |

**实测证据**（`samples/ms/charts.xls`，用 COM 把 `ChartType` 分别设成
`4`/`51`/`5`/`57`/`-4100`）：

```text
C_Scatter 1017 len=6 00 00 96 00 00 00   + 103A len=14 …
C_Bar     1017 len=6 00 00 96 00 01 00        ← fTranspose = 1
C_Column  1017 len=6 00 00 96 00 00 00        ← fTranspose = 0
C_Pie     1019 len=6 00 00 00 00 00 00        ← pcDonut = 0
C_Line    1018 len=2 00 00
```

注意散点图**同时**带 `Bar` 记录，所以判定要有优先级：
`scatter > line > pie > area > bar/column`。按出现顺序「先到先得」会把散点图误判成柱形图。

### 13.1 文字：标题 / 轴标题 / 系列名

三者都是 `SeriesText`（`0x100D`），payload = `u16 link`、`cch u8`、`grbit u8`、字符
（`grbit` 的 bit0 = UTF-16）。区分靠两个**实测**出来的信号：

| 信号 | 含义 |
|---|---|
| `0x1051` 描述符长度 > 8 | 系列名（长的描述符带系列范围/类型） |
| 出现位置在 `Begin`/`End` 的深度 ≤ 2 | 图表标题 |
| 深度 ≥ 3 | 轴标题（在轴的子块里） |

`0x104F`（文本框：`u16`、`u16`、四个 `i32` 的矩形）在**每一个**文字对象前都会出现，
包括空文本的那些，所以「有 `0x104F` 就是标题」是错的——第一版就是这么把轴标题「月份」
当成图表标题的。深度才是可靠信号：

```text
104F @00454C depth=2 rect=(317, 520, 3574, 3042)   ← 空标题框
104F @004FF8 depth=3 rect=(-16, 0, 52, 39)  → 100D '月份'    ← 轴标题
104F @0054AD depth=3 rect=(-16, -4, 42, 56) → 100D 'Sales'   ← 轴标题
104F @005A02 depth=2 rect=(-6, -36, 239, 61) → 100D 'CHART-TITLE-XLS'
```

### 13.2 图表放在哪里

图表子流是 `BOF` 且 `dt = 0x0020`。两种存在方式：

* **图表工作表**：`BOUNDSHEET` 的 `dt = 2`，偏移直接指向该子流（`sheets` 里会出现
  `kind == Chart`、`cells` 为空的条目）。
* **嵌入图表**：`BOUNDSHEET` 里没有它，子流挂在所属工作表记录流的末尾，
  只能扫描全部记录找 `BOF dt = 0x0020`。`samples/ms/chart.xls` 就是这种。

```rust
for chart in &wb.charts {
    println!("{} {}", chart.kind_name(), chart.summary());
    // line "CHART-TITLE-XLS" series=["Sales"] axes=["月份", "Sales"]
}
```

```console
$ officedump info samples/ms/chart.xls | tail -2
charts      1
chart       @15792    line "CHART-TITLE-XLS" series=["Sales"] axes=["月份", "Sales"]
```

## 14. 已知坑

1. **SST 计数器 4 字节**（§5），写错不报错，只是全空。
2. CONTINUE 分段时每段开头有 grbit 字节，不能盲目拼接。
3. 工作表在流里的物理顺序 ≠ 标签页顺序，以 BOUNDSHEET 为准。
4. `MULRK` 的 n 不是显式字段，靠 `len` 反推：`n = (len - 6) / 6`（行 2 + 列首 2 + 列尾 2）。
5. BIFF5 的 `LABEL` 用代码页解码，没有 SST；判断 version 后再走不同分支。
6. 加密工作簿（FILEPASS 0x002F）之后的记录全被 XOR/RC4 加密，遇到就直接放弃。
7. 图表/宏表也是 BOUNDSHEET 指向的 BOF，但记录集完全不同（`dt=0x0020` 图表、
   `dt=0x0040` 宏表）。本仓库**照样枚举**它们（`SheetKind::Chart` / `MacroSheet`），
   只是不解单元格：图表子流里没有 `LABELSST`/`RK`，`cells` 为空。
8. `FORMULA` 的 `cce` 在载荷偏移 20、`rgce` 在 22（§8），按 18/20 读会让每个公式变空。
9. `ptgRef` 存绝对行列（§8.1），当相对偏移读会得到整表错位的引用。
10. `ptgAttr` 是 4 字节（§8.1 坑 4），按规范 5 字节读会立刻跑飞。
11. 共享公式（`grbit` bit2）的从属格子只有 `ptgExp`，得先找主格子的 `rgce` 再按偏移改写。
12. 内置格式编号也可能被 `FORMAT` 记录覆盖（WPS 就写了 ifmt 5/6/7/8），先查自定义表。
12. 图表的「类型」不是字段：由**出现了哪条记录**决定，且散点图同时带 `Bar` 记录，
    判定必须按 `scatter > line > pie > area > bar/column` 的优先级，否则散点图会变成柱形图（§13）。
13. 图表标题不能用「前面有没有 `0x104F` 文本框」判定——每个文字对象（含空文本）都有它；
    要用 `Begin`/`End` 的嵌套深度（≤2 是图表标题，≥3 是轴标题），系列名则用 `0x1051` 的长度区分（§13.1）。

## 15. 写：最小 .xls

`office_xls::write_minimal(sheet_name, &[Cell]) -> Vec<u8>` 直接复用 crate 自己的
`Cell`/`CellValue` 类型，因此写出来的文件必须能被自己的读取器原样读回 —— 这是
对 BIFF8 记录布局最直接的一次交叉验证。

### 15.0 公式：把文本编回 `rgce`

`CellValue::Formula { cached, rgce, text }` 里 `text` 非空时，写入器调用
`office_xls::compile::compile(text)` 把它编译成 token 数组；`text` 为空则直接沿用
`rgce`（例如从另一个文件读出来的公式）。编译不出 token 时退化成「只写缓存值」。

```rust
let cells = vec![
    Cell { row: 0, col: 0, xf: 15, value: CellValue::Number(10.0) },
    Cell { row: 0, col: 1, xf: 15, value: CellValue::Formula {
        cached: CachedValue::Number(30.0), rgce: Vec::new(), text: "A1+A2".into(),
    }},
];
let bytes = office_xls::write_minimal("Formulas", &cells);
```

`FORMULA`(0x0006) 的载荷：`row, col, xf, result(f64), grbit u16, chn u32, cce u16, rgce`。
写入器把 `grbit` 置成 `0x0001`（`fAlwaysCalc`），因此 Excel 打开时会重算，
非数值的缓存值（布尔/文本/错误）也就不会被当成结果。

编码表（与读取器共用 `formula.rs` 的约定）：

| 语法 | token | 字节 |
|---|---|---|
| `A1` | `ptgRef` 0x44 | `row u16` + `(col\|flags) u16`，flags：bit14 列相对、bit15 行相对 |
| `A1:B2` | `ptgArea` 0x25 | `r1 r2 (c1\|f1) (c2\|f2)` |
| `1` / `1.5` | `ptgInt` 0x1E / `ptgNum` 0x1F | u16 / f64 |
| `"x"` | `ptgStr` 0x17 | `cch u8, grbit u8, chars` |
| `SUM(..)` | `ptgFuncVar` 0x42 | `cparams u8, iftab u16` |
| `ROUND(..)` | `ptgFunc` 0x41 | `iftab u16` |
| `+ - * / ^ &` | 0x03..0x08 | — |
| `-x` | `ptgUminus` 0x13 | — |

验证：`samples/written-formulas.xls` 被 Excel 打开后 `A1+A2` 显示为公式、算出 30，
`IF(A1>5,"big","small")` 算出 `big`，`ROUND(A3,1)` 算出 1234.6 —— 说明 token 数组
本身是合法的，而不只是缓存值对。

### 15.1 记录清单（少一条 Excel 就报"文件格式无效"）

| 子流 | 记录（按顺序） |
|---|---|
| globals | `BOF`(dt 0x0005) → `CODEPAGE` → `DATEMODE` → `FONT` → `XF`×16 → `STYLE` → `WINDOW1` → `SST` → `BOUNDSHEET` → `EOF` |
| sheet | `BOF`(dt 0x0010) → `DIMENSIONS` → 单元格 → `WINDOW2` → `EOF` |

- **`BOF` 载荷是 16 字节**：`vers, dt, rupBuild, rupYear, bfh u32, verLowestBiff u32`。
  我一开始只写了 14 字节，Excel 直接拒绝。真实值（`hello.xls`）：
  `00 06 05 00 bb 0d cd 07 c1 80 01 00 06 06 00 00`，即 build 0x0DBB、year 0x07CD、
  bfh 0x000180C1、lowest 0x00000606。
- `FONT` 不能省：`XF` 的 `ifnt` 必须指向一条存在的字体记录（Arial 10pt：
  `dyHeight 200`、`icv 0x7FFF`、`bls 400`）。
- `XF` 至少 16 条，且第 0 条的 `flags` 要带 `fStyle`（0x0004）；单元格统一用 `xf = 15`。
- `WINDOW2` 的 `flags` 用 `0x06B6`（显示网格线、显示行列表头等），后面补 16 字节 0。

### 15.2 一个字节序的坑

`BOOLERR` 的载荷是 `row, col, xf, bBoolErr, fError` —— **先值后标志**。我按
"先 fError" 写，读侧（`crates/office-xls/src/lib.rs` 里 `data[6]` 是值、`data[7]` 是
错误标志）读出来就是 `Bool(false)`，测试 `matches!(cells[2].value, CellValue::Bool(true))`
当场失败。这类"两侧约定必须一致"的地方，正是自写自读测试的价值。

### 15.3 验证

```
$ pwsh tools/roundtrip.ps1
excel  sheet=Writer used=A1:A4
  [1] WRITER-XLS-001
  [2] 1234.5
  [3] 42
  [4] TRUE
```

`SST` 里字符串按出现顺序去重（`isst` 索引），`RK` 对整数值用
`((n as i32) << 2) | 0x2` 的紧凑形式，其余用 f64 高 32 位。

## 15.1 逐格数字格式（`FORMAT` + 每个格式一条 `XF`）

`write_with_formats(sheet_name, cells, formats)` 让每个单元格用 `xf` 选一个格式槽：

```rust
let cells = [
    Cell { row: 0, col: 0, xf: 0, value: CellValue::Number(45356.0) },
    Cell { row: 1, col: 0, xf: 1, value: CellValue::Number(0.25) },
];
let bytes = office_xls::write_with_formats("Fmt", &cells, &["yyyy-mm-dd", "0.00%"]);
```

写入器做的事：

1. 每个格式串写一条 `FORMAT`(0x041E)：`ifmt u16`（从 **164** 起，0..=163 留给内置格式）
   + `XLUnicodeString`（`cch u16`、`grbit u8`（bit0 = UTF-16）、字符）。
2. 每个格式再写一条 `XF`，`ifmt` 指向上面那条。
3. 单元格的 XF 索引 = `CELL_XF + 1 + xf`；`xf` 超出 `formats` 长度时退回默认 XF 15（General）。

顺序与标志位是硬要求（详见 §15.2）：**`FORMAT` 必须在 `XF` 之前**，
**单元格 XF 的 `flags` 必须是 `0xFFF1`（fStyle 清掉）**。

真 Excel 读回（`tools/check_xls.ps1`）：`45356` → `2024-03-05`、`0.25` → `25.00%`、
`1234.5` → `1,234.50`、`-42` → `-42.00`。

## 15.2 真实 Excel 曾经打不开：两个真 bug + 一份模板

第一版的最小 `.xls` 被真实 Excel 拒绝（「Office 检测到文件数据有一个问题」），而
**WPS 能打开** —— WPS 的 `Application.Name`/`Version` 会把自己报成
`Microsoft Excel` / `16.0`，只看名字会把 WPS 当成 Excel。判断办法是打印
`$app.Path`：真 Excel 是 `C:\Program Files\Microsoft Office\Root\Office16`，
WPS 是 `...\Kingsoft\WPS Office\...`。`tools/check_xls.ps1` 现在一上来就打印它。

二分办法：把「真 Excel 写的 globals」与「我们写的 sheet 子流」拼起来，重新打包成
CFB 再交给 Excel。

| 组合 | 结果 |
|---|---|
| 真 Excel 的 globals + **我们的 sheet 子流** | ✅ 打开，值正确 |
| **我们的 globals** + 真 Excel 的 sheet 子流 | ❌ 拒绝 |

→ **sheet 子流（`BOF`/`DIMENSIONS`/`LABELSST`/`NUMBER`/`BOOLERR`/`WINDOW2`/`EOF`）完全正确，问题在 globals。**

逐个把我们的记录替换进 Excel 的 globals（每次只换一种），得到三条确定的结论：

1. **单元格 `XF` 的 `flags` 不能置 `fStyle`（bit 2）**。单元格 XF 若声称自己是「样式 XF」，
   Excel 就当它是样式定义、**忽略它的 `ifmt`**，所有数字都显示成「常规」。
   单元格 XF 用 `0xFFF1`（各属性组都适用、只清 fStyle），样式 XF 用 `0xFFF5`。
2. **`FORMAT` 必须排在引用它的 `XF` 之前**。Excel 边读边解析 `ifmt`，XF 指到后面才出现的
   `FORMAT` 时退回「常规」。
3. `STYLE` 的 `flags` 里 `fBuiltin` 是 **bit 0**（写 `0x01`），不是 `0x80`。

即使三处都改对，**手写的 globals 仍被 Excel 拒绝**；而把我们的记录逐条换进 Excel 的
globals 时每一条单独都可接受 —— 说明缺的是**某个我们完全没写的记录**，而
`INTERFACEHDR`/`MMS`/`INTERFACEEND`、`DSF`、`TABID`、`PRECISION`、`BOOKBOOL`、
`USESELFS`、`BUILTINFMTCNT`、`HIDEOBJ` 逐个补进去都没解决。

于是采用与 `.ppt` 相同的办法：**内嵌一份真 Excel 写出的 globals**
（`crates/office-xls/templates/globals.bin`，14698 字节，由
`tools/extract_xls_globals.py` 从 `samples/ms/one.xls` 提取；去掉了 `SST`/`EXTSST`/
`BOUNDSHEET`，因为这三样依赖具体内容，由写入器自己发）。写入器读模板、丢掉模板里的
`FONT`/`FORMAT`/`XF`/`STYLE`/`SST`/`EXTSST`/`BOUNDSHEET`，再写自己的那几样。

验证（`tools/check_xls.ps1`，真 Excel）：

```
written.xls          sheet: Writer    A1 WRITER-XLS-001   A2 1234.5   A3 42
written-formats.xls  sheet: Formats   A1 yyyy-mm-dd  A2 25.00%  A3 1,234.50  A4 -42.00
written-formulas.xls sheet: Formulas  B1 =A1+A2 → 30   B3 =ROUND(A3,1) → 1234.6
                                     B4 =IF(A1>5,"big","small") → big   B5 =SUM(A1:A2)/2 → 15
```

（`written-formats.xls` 的 A1 显示 `########` 只是列宽不够显示日期，格式本身是对的。）

**教训**：把「读得对」当成「写得对」是最容易犯的错；验证写入器必须用真实程序，
而且先确认 COM 连到的到底是哪个程序。


### 15.3 globals 不再是内嵌二进制

`templates/globals.bin` 已经**不再被写入器读取**。`crates/office-xls/src/globals.rs` 里
286 条记录以 Rust 源码形式列出（`const RECORDS: &[(u16, &[u8])]`），`write_globals()`
逐条发出；`.bin` 只留作**测试夹具**，`globals_reproduce_that_workbook_byte_for_byte`
拿它做逐字节比对，防止这份源码悄悄漂移。

删除内嵌模板的过程记录（也是这次踩的三次坑）：

1. 先按记录类型逐组删除、让真 Excel 裁决 —— **一条都删不掉**（见 §15.2 的二分表）。
2. 于是改成「生成」而不是「裁剪」：`tools/gen_globals_rs.py` 从参照文件生成源码。
3. 期间连续三次字段切分失败，根因是**用了一个没验证同步性的遍历器当标尺**：
   简易遍历器在 `globals` 里失步，于是我拿错位的字节去对齐 `STYLE`/`StyleExt` 的字段。
   验证同步只需一行：从 `BOF` 走到 `EOF` 时 `at == 文件长度`（实测 286 条 / 14698 B / 同步）。
   教训与 §14 第 1 条同一个：**先验标尺，再量东西**。

### 15.4 `XF` 载荷的字段解码（把「28 种尾巴」变成 6 个字段）

`XF` 的 20 字节载荷逐字节解码（62 条实测统计）：

| 偏移 | 字段 | 实测取值 |
|---|---|---|
| 0..2 | `ifnt` 字体 | 0/1/2/5/7/8/9/12..14/15..17/20/21 |
| 2..4 | `ifmt` 数字格式 | 0（少数 9） |
| 4..6 | `flags` | `0xFFF5`（样式 XF，`fStyle` 置位）/ `0x0001`（单元格 XF） |
| 6 | `alc` 对齐 | 恒为 `0x10` |
| 7 | `irot`+`indent` | 恒为 `0` |
| 8..12 | `brd` 边框 | 11 种：`00000000` `00f40000` `00b40000` `00d40020/50/60/61` `00f80000` `00941111` `00946666` `009c1111` |
| 12..14 | `fill` 填充 | `0000`，两条例外（`970b`、`bf1f` 主题填充色） |
| 14..18 | 保留 | 恒为 0 |
| 18..20 | 主题/扩展索引 | 18 种（`c020` `9b20` `b720` …），与 `XFEXT` 一一对应 |

即：**62 条 XF 只由 `(ifnt, ifmt, flags, brd, fill, tail)` 六个字段决定**，不需要"抄 28 份尾巴"——
生成器按字段拼即可，`brd`/`fill` 的语义按 MS-XLS 2.4.353 是
`{left:{style,color}, right, top, bottom}` 与 `{icvFore, icvBack}`。

## 16. 表格外观：合并单元格、列宽、行高、冻结窗格

`office_xls::write::write_with_layout(sheet_name, cells, formats, &Layout)` 在写单元格之外
再写四条记录，`Layout` 用链式构造：

```rust
let layout = Layout::default()
    .merge(0, 0, 0, 2)      // A1:C1
    .column(0, 20.0)        // A 列 20 个字符宽
    .row(0, 40.0)           // 第 1 行 40 磅高
    .freeze(1, 1);          // 冻结首行 + 首列
```

记录顺序与 Excel 一致：`BOF`、`DEFCOLWIDTH`、`COLINFO`*、`DIMENSIONS`、`ROW`*、单元格、
`WINDOW2`、`PANE`、`MERGEDCELLS`、`EOF`。逐条实测（`samples/ms/layout.xls`，Excel 2016
用 COM 写：合并 `A1:C1`、列宽 20/5/30、行高 40/22.5、在 `B2` 冻结窗格）：

| 记录 | 载荷 | 实测 |
|---|---|---|
| `DEFCOLWIDTH` 0x0055 | `8` | 默认列宽 8 个字符 |
| `COLINFO` 0x007D | `colFirst, colLast, coldx, ixfe, grbit, reserved` | `00 00 00 00 a4 14 0f 00 02 00 00 00` = A 列、`coldx` 5284、`ixfe` 15、`grbit` 0x0002（用户设定） |
| `ROW` 0x0208 | `rw, colMic, colMac, miyRw, irwMac, reserved, grbit, ixfe` | `00 00 00 00 03 00 20 03 …40 01 0f 00` = 第 1 行、`miyRw` 800 twips（40 磅）、`grbit` 0x0140 |
| `WINDOW2` 0x023E | `grbit, rwTop, colLeft, icvHdr, wScaleSLV, wScaleNormal, reserved` | 不冻结 `b6 06`；冻结 `be 07`（多 `0x0108` = `fFrozen`+`fFrozenNoSplit`），尾部 `reserved` 也从 0 变 2 |
| `PANE` 0x0041 | `xSplit, ySplit, rwTop, colLeft, pnnAct` | `01 00 01 00 01 00 01 00 00 00` = 冻 1 列、1 行 |
| `MERGEDCELLS` 0x00E5 | `cref` + 每段 8 字节 | `01 00 | 00 00 00 00 00 00 02 00` = `A1:C1` |

**列宽的单位是坑**：`coldx` 是 1/256 个字符。Excel 自己写的是 `width * 256 + 164`，但那个 164
是「5 像素」按**工作簿默认字体**的最大数字宽度折算出来的：我们的 `FONT` 记录固定是 Arial 10，
Excel 读回来按 `coldx / 256 - 0.71875` 算，所以生成器用 **184**。实测（`examples/xlslayout.rs`
写、`tools/check_layout.ps1` 用真 Excel 读回）：请求 10/20/50 读回 10.00/20.00/50.00；请求 12.5
读回 12.54（Excel 把宽度量化到整像素，奇数宽度会有几百分之一的偏差）。

行高没这个问题：`miyRw = 磅 * 20`，40 和 22.5 都精确读回。

冻结窗格要两件事：`WINDOW2.grbit` 的 `0x0008`（`fFrozen`）与 `0x0100`（`fFrozenNoSplit`），
加一条 `PANE`（`xSplit` = 冻结列数、`ySplit` = 冻结行数，冻结时这两个字段是**个数**而不是
1/20 磅）。只写 `PANE` 不置位 `WINDOW2` 时 Excel 会当成「拆分窗口」而不是冻结。

读侧同步支持（`Sheet::furniture`，`Furniture { merged, columns, rows, frozen }`），所以往返测试
不需要真 Excel：`write::layout_tests::layout_round_trips_through_the_reader` 与
`lib::tests::reads_sheet_furniture_from_the_sample`（后者读的就是 `samples/ms/layout.xls`，
断言 `coldx` 5284/1444/7844、`miyRw` 800/450、`frozen = Some((1, 1))`）。

真 Excel 复核：`tools/check_layout.ps1 samples/written-layout.xls` 输出
`A1 merge: A1:C1`、`width A/B/C: 20/5/30`、`height 1/3: 40/22.5`、
`freeze panes: True splitRow=1 splitCol=1`。

## 17. 多工作表

`write::write_workbook(&[SheetData])` 把几张表写进同一个 `Workbook` 流：

```rust
use office_xls::write::{Layout, SheetData, write_workbook};

let a = [Cell { row: 0, col: 0, xf: 15, value: CellValue::Text("first".into()) }];
let b = [Cell { row: 0, col: 0, xf: 0, value: CellValue::Number(42.0) }];
let bytes = write_workbook(&[
    SheetData::new("One", &a).with_formats(&["yyyy-mm-dd"]),
    SheetData::new("Two", &b).with_layout(Layout::default().column(0, 20.0)),
]);
```

`write_minimal` / `write_with_formats` / `write_with_layout` 都是它的单表包装，产物逐字节不变
（`sheets_tests::a_single_sheet_workbook_matches_write_minimal` 守着这一点）。

三张表共享、但各自独立的四样东西：

| 东西 | 作用域 | 说明 |
|---|---|---|
| `BOUNDSHEET` | globals 里一条 | `lbPlyPos` 指向该表的子流起点，要等子流拼完才回填 |
| 子流 | 每表一份 | `BOF … EOF`，列宽/行高/合并/冻结都在自己那份里 |
| `SST` | 全工作簿 | 字符串索引是全局的，重复文字只存一次 |
| `XF` / `FORMAT` | 全工作簿 | 格式表全局，但每格用**本表** `formats` 的下标，见下 |

**`xf` 的语义**：`Cell::xf` 是**本表 `formats` 的下标**，不是全局下标 —— `write_workbook` 先把
各表的格式串去重成全局表，再为每张表建一张 `xf → 全局下标` 的映射（`xf_for`）。所以两张表
各自写 `with_formats(&["0.00%"])` 时，它们的 `xf = 0` 都指向自己那条格式；`xf` 超出本表
`formats` 长度时落到默认 `XF`（General）。

真 Excel 复核：`tools/check_sheets.ps1 samples/written-sheets.xls` 输出

```
sheets: 3
  First: A1=WRITER-XLS-SHEETS text=WRITER-XLS-SHEETS fmt=General width=8.54
  Second: A1=second sheet text=second sheet fmt=General width=8.54
  Third: A1=3 text=3 fmt=General width=24
```

（`samples/written-sheets.xls` 的第一张表还带 `yyyy-mm-dd` 与 `0.00%` 两格，
`tools/check_xls.ps1` 读回 `A2 = 2024-03-05`、`A3 = 25.00%`。）

`tools/check_sheets.ps1` 与 `check_xls.ps1` 一样，先确认 `$app.Path` 是
`C:\Program Files\Microsoft Office\Root\Office16`（WPS 会自称 Microsoft Excel），
并且要把路径转成绝对路径再 `Workbooks.Open`（相对路径会报「找不到文件」）。

## 18. 单元格字体（`FONT` / `XF.ifnt` / `PALETTE`）

### 18.1 记录

`FONT`（`0x0031`，18 字节固定头 + 名字）：

| 偏移 | 字段 | 说明 |
|---|---|---|
| 0 | `dyHeight` | 字号 × 20（14 磅 = 280） |
| 2 | `grbit` | bit0 粗体、bit1 斜体、bit2 下划线、bit3 删除线 |
| 4 | `icv` | 颜色索引，`0x7FFF` = 自动（跟随系统文字色） |
| 6 | `bls` | 字重，粗体写 700、常规写 400 |
| 8 | `sss` | 上/下标，我们写 0 |
| 10 | `uls` | 下划线样式，我们写 0 |
| 12 | `bFamily` / `bCharSet` | 2 / 0 |
| 14 | `cch` | 名字长度（字符数） |
| 15 | `fHighByte` | 1 = UTF-16LE 名字 |
| 16 | `rgch` | 名字 |

`XF` 的 `ifnt`（偏移 0）指向这张表。`icv` 小于 8 是内建色，8 起是 `PALETTE`
（`0x0092`，`ccv` + 每项 4 字节，字节序 `0x00BBGGRR`）里的下标 —— 写侧把用到的
颜色按首次出现顺序排进去，`icv = 8 + 序号`。

### 18.2 Excel 会从索引 4 起重排字体表（这是本节的全部难点）

写侧最早按规范把自定义字体接在 5 张基础字体之后，`ifnt` 直接写表内下标。真
Excel 打开后**索引 ≥ 4 的字体全部错位一位**：14 磅的标题显示成 10 磅、斜体显示成
粗体，只有前四个字体是对的。

二分（每份文件只改一个变量，都用 `tools/check_styles.ps1` 读回）：

| 字体表 | `ifnt` | Excel 实际显示 |
|---|---|---|
| 4 个（10/11/12/13 磅） | 0..3 | 10 / 11 / 12 / 13 ✓ |
| 5 个（10..14 磅） | 0..4 | 10 / 11 / 12 / 13 / **10** |
| 7 个（10..16 磅） | 0..6 | 10 / 11 / 12 / 13 / **10** / 14 / 15 |
| 5 个，第 5 个是「斜体 10 磅」 | 0..4 | 索引 4 仍是**正体 10 磅**，索引 5 才是斜体 |

最后一行说明 Excel 不是「替换」第 5 个字体，而是**在索引 4 插进一个默认字体**，
把后面的整体挤后一位。用 Excel 自己另存这份文件，读回来有 26 个字体：我们的
0..3、**插入的 Arial 10（第 5 个）**、我们原来的 4..5，然后是 13 个内建样式字体
（等线 11）—— 也就是说 Excel 打开时就是这么理解的，不是显示层的问题。

排掉的原因（都不影响错位）：5 张基础字体写成 5 个完全相同的副本；删掉 globals 里
的 `0x0892`（StyleExt ×16）与 `0x087C`；把 `0x0892` 引用的 XF 号改成存在的；
把 16 个基础 `XF` 全写成 `fStyle = 1`；写 64 个基础 `XF` 让 `0x0892` 引用的
30/34/38 真的存在。Excel 自己写的 `samples/ms/layout.xls` 里前 5 个字体是 5 个
**完全相同**的「等线 11」，且没有 `0x0293`（`STYLE`）记录 —— 所以「前 5 个必须是
内建样式字体」这条也不成立。

### 18.3 对策：每张自定义字体写两遍

规范读者按 `ifnt` 直取第 `ifnt` 个 `FONT`；Excel 把 ≥ 4 的下标减一。让同一个脸
占**两条相邻记录** `p`、`p+1`，`ifnt` 写 `p + 1`，两种解释就都落在同一张脸上：

```
FONT 0..3   基础脸（默认、粗体、斜体、粗斜体）      ifnt = 0..3
FONT 4, 5   自定义脸 A（同一张脸写两遍）            ifnt = 5
FONT 6, 7   自定义脸 B                              ifnt = 7
```

代价是字体表里有重复记录（Excel 的字体下拉框会多出几条同名项），换来的是「无论
对端是否重排都对」。基础脸因此只剩 4 张（原来的下划线脸改成普通自定义脸）。

### 18.4 API

```rust
use office_xls::{Font, SheetData, XfSpec, write::write_workbook};

let bold = XfSpec::plain().font(Font { bold: true, ..Font::default_face() });
let red  = XfSpec::plain().font(Font { colour: Some((0xC0, 0x00, 0x00)), ..Font::default_face() });
let pct  = XfSpec::format("0.00%").font(Font { bold: true, ..Font::default_face() });
let bytes = write_workbook(&[SheetData::new("S", &cells).with_xfs(&[bold, red, pct])]);
```

`XfSpec` 把「数字格式 + 字体」绑成一条 `XF`，`Cell::xf` 是 `xfs` 的下标（`with_formats`
的老路径仍然可用）。读侧：`Workbook::fonts`、`Workbook::palette`、
`Workbook::font(cell.xf) -> Option<&Font>`。

真 Excel 复核 `tools/check_styles.ps1 samples/written-styles.xls`：

```
app path: C:\Program Files\Microsoft Office\Root\Office16
  A1 text=WRITER-XLS-STYLES bold=True italic=False size=14 colour=#000000 name=Arial
  A2 text=italic 14pt       bold=False italic=True  size=14 colour=#000000 name=Arial
  A3 text=red              bold=False italic=False size=10 colour=#c00000 name=Arial
  A4 text=25.00%           bold=True italic=False size=10 colour=#000000 name=Arial fmt=0.00%
```

## 19. 边框与对齐（`XF` 的 `dg*` / `icv*` / `alc`）

### 19.1 `XF` 的 20 个字节

MS-XLS §3.9.11 的 `Xf` 结构，逐字节（本节所有偏移都对着它核对过）：

| 偏移 | 长度 | 字段 |
|---|---|---|
| 0 | 2 | `ifnt`，`FONT` 下标（§18） |
| 2 | 2 | `ifmt`，数字格式（§15.1） |
| 4 | 2 | `fLocked`/`fHidden`/`fStyle`/`f123Prefix`/`ixfParent(12)` |
| 6 | 1 | `alc`(3) / `fWrap`(1) / `alcV`(3) / `fJustLast`(1) |
| 7 | 1 | `trot`（文字旋转） |
| 8 | 1 | `cIndent`(4) / `fShrinkToFit`(1) / `reserved`(1) / `iReadOrder`(2) |
| 9 | 1 | 未命名（Excel 的样式 XF 写 `0xF4`/`0xB4`/`0xD4`，单元格 XF 写 `0x20`；我们沿用 `0xF4`，Excel 接受） |
| 10 | 2 | `dgLeft`(4) / `dgRight`(4) / `dgTop`(4) / `dgBottom`(4) |
| 12 | 2 | `icvLeft`(7) / `icvRight`(7) / `grbitDiag`(2) |
| 14 | 4 | `icvTop`(7) / `icvBottom`(7) / `icvDiag`(7) / `dgDiag`(4) / `reserved`(1) / `fls`(6) |
| 18 | 2 | `icvFore`(7) / `icvBack`(7) |

要点：

- **四个边的线型在同一个 2 字节里**，顺序是左、右、上、下，每个 4 bit。
- **边框颜色的「自动」是 `0x40`**，不是字体的 `0x7FFF` —— 两者是不同字段的不同约定，
  混用会让 Excel 把边框画成黑色以外的怪色。
- 对角线：`grbitDiag` 的 bit0 = 左上到右下、bit1 = 左下到右上（两个都给就是交叉），
  线型在 `dgDiag`，颜色在 `icvDiag`。
- `alc`：0 常规、1 左、2 居中、3 右、4 填充、5 两端对齐、6 跨列居中、7 分散对齐。
  `alcV`：0 顶、1 居中、2 底、3 两端、4 分散。

线型码（`dg*`，MS-XLS `BorderStyle`）：

| 码 | 样式 | 码 | 样式 |
|---|---|---|---|
| 0 | 无 | 4 | 点线 |
| 1 | 细线 | 5 | 粗线 |
| 2 | 中线 | 6 | 双线 |
| 3 | 虚线 | 7 | 极细线（hair） |

注意 `Hair` 是 **7**（不是 8），而 Excel 的 `Border.Weight` 枚举是另一套
（`xlHairline = 1`、`xlThin = 2`、`xlThick = 4`、`xlDouble = -4119`），两边不要混。

### 19.2 怎么确认的

`samples/ms/borders.xls` 是 Excel 2016 自己写的：A1 四边细线、B2 四边粗线、C3 只有
下边框双线、A5:B5 合并加虚线框、D1 两条对角线。`cargo run -q -p officedump --example
xfdump -- samples/ms/borders.xls` 把每一条 `XF` 解出来，四个单边样本的差异只落在
`dg` 的一个 nibble 上，于是左/右/上/下的顺序、`icv` 的位置、`grbitDiag` 的两位都能
一一对上：

```
xf  16 dg=1/1/1/1   icv=40/40/40/40 diag=0     # A1 细线框
xf  17 dg=5/5/5/5   icv=40/40/40/40 diag=0     # B2 粗线框
xf  18 dg=0/0/0/6   icv=40/40/40/40 diag=0     # C3 下边双线
xf  22 dg=0/0/0/0   icv=40/40/40/40 diag=3     # D1 交叉对角线
```

（`icv=40` 就是「自动」；真彩色时是 `PALETTE` 的下标，从 8 开始。）

### 19.3 API

```rust
use office_xls::write::{write_workbook, SheetData, XfSpec, Layout};
use office_xls::{Align, Border, BorderStyle, BoxBorder, VAlign};

let thin_box = BoxBorder::thin_box();                       // 四边细线
let red = BoxBorder::all(Border::thin().colour((0xFF, 0, 0)));
let underline = BoxBorder { bottom: Border::new(BorderStyle::Double), ..BoxBorder::default() };
let diagonal = BoxBorder {
    diagonal: Border::new(BorderStyle::Thin),
    diagonal_down: true,
    ..BoxBorder::default()
};

let xfs = [
    XfSpec::plain().border(thin_box),
    XfSpec::plain().border(red).align(Align::Centre).valign(VAlign::Centre).wrap(true),
    XfSpec::plain().border(underline).indent(2),
    XfSpec::plain().border(diagonal),
];
let bytes = write_workbook(&[SheetData::new("S", &cells).with_xfs(&xfs)]);
```

读侧：`Workbook::xfs[xf]` 现在带 `align`、`valign`、`wrap`、`indent`、`border`
（`BoxBorder`，每条边是 `Border { style, colour }`，`colour: None` 表示自动）。
`Align`/`VAlign`/`BorderStyle` 都实现了 `code()`/`from_code()`，未认识的码会退回
「常规/底端对齐/细线」，不会让边框凭空消失。

真 Excel 复核 `tools/check_borders.ps1 samples/written-borders.xls`（每个样式隔一行
放，避免相邻单元格共用一条边导致读数串味）：

```
app path: C:\Program Files\Microsoft Office\Root\Office16
  A1 boxed       h=1 v=-4107 wrap=False indent=0 [7=style1/weight2 10=… 8=… 9=…]
  A3 thick       h=1 v=-4107 wrap=False indent=0 [7=style1/weight4 10=… 8=… 9=…]
  A5 underline   h=1 v=-4107 wrap=False indent=0 [9=style-4119/weight4]
  A7 centred and wrapped text h=-4108 v=-4108 wrap=True indent=0 []
  A9 indented    h=1 v=-4107 wrap=False indent=2 []
  A11 red edge   h=1 v=-4107 wrap=False indent=0 [7..9 colour16711680]
  C1 diagonal    h=1 v=-4107 wrap=False indent=0 [5=style1/weight2]
```

**相邻单元格共用一条边**：A1 的下边框就是 A2 的上边框，Excel 读 `Borders.Item(8)`
时会报出更粗的那条。所以核对边框一定要把样本行隔开放，否则会以为自己的写入串了行。

### 19.4 对齐字段的实测

`samples/ms/align.xls`（Excel 2016 自己写，A1/A2/A3 = 左/中/右，A5/A6/A7 = 顶/中/底，
A9 = 自动换行 + 缩进 3，A11 = 旋转 45°）里，只有这几条单元格 `XF` 的第 6～8 字节不同：

| 单元格 | 字节 6 | 字节 7 | 字节 8 | 含义 |
|---|---|---|---|---|
| A1 | `0x11` | 0 | 0 | `alc=1` 左 |
| A2 | `0x12` | 0 | 0 | `alc=2` 居中 |
| A3 | `0x13` | 0 | 0 | `alc=3` 右 |
| A5 | `0x00` | 0 | 0 | `alcV=0` 顶端 |
| A6 | `0x10` | 0 | 0 | `alcV=1` 垂直居中 |
| A7 | `0x20` | 0 | 0 | `alcV=2` 底端（Excel 的默认） |
| A9 | `0x19` | 0 | `0x03` | `alcV=1` + `fWrap` + 缩进 3 |
| A11 | `0x10` | `0x2D` | 0 | 旋转 45°（`trot`，度） |

于是「字节 6 = `alc | fWrap<<3 | alcV<<4`」这一条被四个水平值、三个垂直值和一个
换行位同时钉住。注意 Excel 的默认垂直对齐是**底端**（`0x20`），写侧 `VAlign::default()`
也是 `Bottom`，两边一致；`trot` 我们只读不写（`XfSpec` 没有旋转字段）。

## 20. 单元格填充（`fls` / `icvFore` / `icvBack`）

填充是三个字段拼出来的（都在 `XF` 里，位置见 §19.1）：

- `fls`：6 bit，`XF` 字节 14～17 那个 u32 的**最高 6 位**（bit 26..31），图案类型；
- `icvFore`：字节 18 的低 7 位，图案的颜色（纯色填充时就是填充色本身）；
- `icvBack`：字节 18 的 bit 7..13，图案背后的颜色。

图案码（MS-XLS `FillPattern`）：

| 码 | 图案 | Excel 常量 | 码 | 图案 | Excel 常量 |
|---|---|---|---|---|---|
| 0 | 无 | `xlNone` (-4142) | 9 | 深色网格 | `xlGrid` (9) |
| 1 | 纯色 | `xlSolid` (1) | 10 | 深色斜网格 | `xlTrellis` (10) |
| 2 | 50% 灰 | `xlGray50` (-4125) | 11 | 浅色横条 | `xlLightHorizontal` |
| 3 | 75% 灰 | `xlGray75` (-4124) | 12 | 浅色竖条 | `xlLightVertical` |
| 4 | 25% 灰 | `xlGray25` (-4126) | 13 | 浅色斜条 | `xlLightDown` |
| 5 | 深色横条 | `xlHorizontal` (2) | 14 | 浅色反斜条 | `xlLightUp` |
| 6 | 深色竖条 | `xlVertical` (3) | 15 | 浅色网格 | `xlLightGrid` |
| 7 | 深色斜条 | `xlDown` (4) | 16 | 浅色斜网格 | `xlLightTrellis` |
| 8 | 深色反斜条 | `xlUp` (5) | 17/18 | 12.5% / 6.25% 灰 | `xlGray125` / `xlGray0625` |

**「自动」色有两套值**：字体用 `0x7FFF`，边框和填充用 `0x40`（`icvBack` 还有一个
`0x41`，Excel 自己写默认单元格时就用它）。写侧统一写 `0x40`，读侧 `0x40`/`0x41`
都当「没有颜色」。

### 20.1 内建调色板

Excel 写文件时**通常不带 `PALETTE` 记录**（`samples/ms/fills.xls`、`borders.xls`
都没有），这时 `icv >= 8` 指的是 BIFF 的内建调色板，而不是「查不到」。所以读侧
内置了那张 56 色的表（索引 8..63：黑、白、红、绿、蓝、黄、品红、青，然后 8 个
深色，再往后是 40 个 Office 97 配色），只有工作簿真的带 `PALETTE` 时才以文件里的
为准。漏掉这张表的直接后果是：所有 Excel 自己写的彩色单元格都读成「没有颜色」。

### 20.2 API

```rust
use office_xls::write::{write_workbook, Fill, Pattern, SheetData, XfSpec};

let xfs = [
    XfSpec::plain().fill(Fill::solid((0xFF, 0x00, 0x00))),      // 纯红
    XfSpec::plain().fill(Fill {                                  // 50% 灰
        pattern: Pattern::Gray50,
        fore: Some((0xC0, 0xC0, 0xC0)),
        back: Some((0xFF, 0xFF, 0xFF)),
    }),
    XfSpec::plain().fill(Fill {                                  // 深色网格 + 藏青
        pattern: Pattern::DarkGrid,
        fore: Some((0x00, 0x00, 0x80)),
        back: None,                                             // 自动
    }),
];
let bytes = write_workbook(&[SheetData::new("Fills", &cells).with_xfs(&xfs)]);
```

读侧：`Xf::fill` 是 `Fill { pattern, fore, back }`，颜色是 `Option<(u8,u8,u8)>`，
`None` 就是「自动」。`Pattern` 有 `code()`/`from_code()`，未知码读成 `Pattern::None`
（宁可少画，也不要画错）。

真 Excel 复核 `tools/check_fills.ps1 samples/written-fills.xls`：

```
  A1 solid red      pattern=1     fore=16711680   patternColor=0
  A3 grey 50%       pattern=-4125 fore=16777215   patternColor=12632256
  A5 dark grid      pattern=9     fore=0          patternColor=128
  A7 light trellis  pattern=16    fore=32768      patternColor=16776960
  A9 no fill        pattern=-4142 fore=16777215   patternColor=0
```

`Interior.Color` 是**底色**、`Interior.PatternColor` 是**图案色**：纯色填充时 Excel
把填充色放到底色里（`A1 fore=16711680` 就是红色），图案填充时图案色才走
`PatternColor`（`A5` 的藏青是 `128 = 0x000080`）。读回时看这两个属性就能分辨
`icvFore`/`icvBack` 有没有写反。

## 21. 批注与超链接

批注和超链接都不在单元格记录里，而是同一段工作表子流里的独立记录：批注是
`NOTE`(0x001C)，超链接是 `HLINK`(0x01B8)。样本 `samples/ms/notes-hlink.xls`
由 Excel 2016 自己写出来（`tools/tmp_notes_xls.ps1`）：A1 批注「reviewed by ops」
（作者 ops）、A4 批注两行（作者 qa）、B2 链接到 `https://example.com/page`（显示
「click me」）、B5 是 `mailto:`。

### 21.1 批注：`NOTE` 与 `OBJ`/`TXO`

`NOTE` 记录（`docs/xls.md` 之外，见 MS-XLS §2.4.151）：

```
rw(2) col(2) grbit(2) idObj(2) cchAuthor(2) flags(1) author(XLUnicodeString)
```

批注的**正文**不在 `NOTE` 里，而在一条 `OBJ`(0x005D) 指向的 `TXO`(0x01B6) 里：
`OBJ` 的 `ftCmo` 结构里 `ot == 0x19`（批注），`id` 就是 `NOTE.idObj`。记录顺序是

```
MSODRAWING, OBJ, MSODRAWING, TXO, CONTINUE, CONTINUE     // 一个批注框
...
NOTE, NOTE                                              // 全部批注框之后
HLINK, HLINK
```

所以读侧要两遍：先扫 `OBJ`（只认 `ot == 0x19`）记住 `id`，再扫 `TXO` 取正文，
最后把 `NOTE` 按 `idObj` 配对。`TXO` 的正文在它后面的 `CONTINUE`(0x003C) 里，
而 `CONTINUE` 的**第一个字节是压缩标志**（0x00 = cp1252，0x01 = UTF-16LE），不是
`XLUnicodeString` 的那个 flags 字节——照搬字符串解析会整体错一位。

```rust
for cell_comment in &sheet.comments {
    println!("{}:{} {} — {}", cell_comment.row, cell_comment.col,
             cell_comment.author, cell_comment.text);
}
```

### 21.2 超链接：`HLINK` 的字节布局

```
ref8         rwFirst(2) rwLast(2) colFirst(2) colLast(2)   // 0 基
CLSID        CLSID_Hyperlink = 79EAC9D0-BAF9-11CE-8C82-00AA004BA90B
u32          2
u32          0x17
u32          cchText                       // 含结尾 NUL
text         UTF-16LE + NUL
CLSID        CLSID_StdURLMoniker = 79EAC9E0-BAF9-11CE-8C82-00AA004BA90B
u32          byte size of the address + 26
address      UTF-16LE + NUL
tail         26 bytes, identical in every sample
```

两个坑：

- `ref8` 是**0 基**行号/列号，和 BIFF 里几乎所有其它行列字段都相反；
- URL 前面那个 u32 不是字符数，而是**字节长度**（地址 2 倍 + 26 字节尾巴）。
  照字符数读会多要一倍，`read_utf16z` 直接返回 `None`，看起来像「记录不存在」。

写侧：

```rust
let links = [
    Link { row: 1, col: 1, url: "https://example.com/page", text: "example.com" },
    Link { row: 3, col: 1, url: "mailto:someone@example.com", text: "mail us" },
];
let bytes = write_workbook(&[SheetData::new("Links", &cells).with_links(&links)]);
```

`text` 为空时用 `url` 兜底（Excel 显示的就是地址本身）。`HLINK` 写在单元格记录
之后、`EOF` 之前，位置不重要，Excel 按 `ref8` 挂到单元格上。

真 Excel 复核 `tools/check_links.ps1 samples/written-links.xls`：

```
app path: C:\Program Files\Microsoft Office\Root\Office16
$B$2: text=example.com address=https://example.com/page type=0
$B$4: text=mail us     address=mailto:someone@example.com type=0
```

### 21.3 脚本要「删」注册表键，不是「改」它

`tools/check_*.ps1` 里原来把
`HKCU\Software\Classes\CLSID\{00024500-…}\LocalServer32` 改写成 `EXCEL.EXE`。
在这台机器上**不生效**：WPS 把该键写在 HKCU，改写成 Excel 之后 COM 仍然起 WPS
（`$app.Path` 会如实报出 `...\Kingsoft\WPSOFF~1\...\office6`）。正确做法是**删掉**
这个键，让 HKLM 的机器注册生效，脚本结束时再按原值重建：

```powershell
$old = (Get-ItemProperty -Path $key)."(default)"
Remove-Item -Path $key -Force -ErrorAction SilentlyContinue
...
finally { New-Item -Path $key -Force | Out-Null
          Set-ItemProperty -Path $key -Name "(default)" -Value $old }
```

凡是打印 `app path` 的脚本都该看这一行——WPS 自称 `Application.Name = "Microsoft
Excel"`、`Version = 16.0`，只有 `Path` 骗不了人。
