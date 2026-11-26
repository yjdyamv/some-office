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

## 13. 已知坑

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
