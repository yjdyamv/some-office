# .doc：Word 97-2003 二进制文档

`WordDocument` 流是"目录"，`0Table` / `1Table` 流是"附录"。两者配合才能拿到正文：

```
WordDocument 流
  0x0000  FIB（File Information Block，约 2 KB）
  0x0800  正文（piece 指向这里）
  ...     CHPX / PAPX / 各种 FKPs
0Table (或 1Table)
  CLX → PlcPcd（piece table）→ 每个 piece 的 fc
  SttbfFfn（字体表）、PlcfSed（节）、Dop（文档属性）...
```

本仓库实现：`crates/office-doc/src/lib.rs`。

## 1. FIB 布局

`hello.doc` 的关键字段（`WordDocument` 前 0x1B0 字节）：

| 偏移 | 大小 | 字段 | hello.doc | 说明 |
|---|---|---|---|---|
| 0x00 | 2 | wIdent | 0xA5EC | 魔数，Word 文档固定 |
| 0x02 | 2 | nFib | 0x00C1 | 0xC1 = Word 97；< 0xC1 是 Word 6/95 |
| 0x06 | 2 | lid | 0x0804 | 语言 ID（简体中文） |
| 0x0A | 2 | flags | 0x5014 | 见下 |
| 0x0C | 2 | nFibBack | 0x00BF | 能读这个文件的最低版本 |
| 0x18 | 4 | fcMin | 0x0800 | 正文起点 |
| 0x1C | 4 | fcMac | 0x08A4 | 正文终点 |
| 0x20 | 2 | csw | 0x000E | 后面有 14 个 u16 |
| 0x22 | 28 | fibRgW97 | | 14 × u16，含 lidFE 等 |
| 0x3E | 2 | cslw | 0x0016 | 后面有 22 个 u32 |
| 0x40 | 4 | cbMac | 3630 | **等于 WordDocument 流长度**（自检用） |
| 0x44 | 8 | reserved1/2 | 0 | |
| 0x4C | 4 | ccpText | 82 | 主文档字符数 |
| 0x50 | 4 | ccpFtn | 0 | 脚注字符数 |
| 0x54 | 4 | ccpHdd | 0 | 页眉页脚 |
| 0x5C | 4 | ccpAtn | 0 | 批注 |
| 0x60 | 4 | ccpEdn | 0 | 尾注 |
| 0x64 | 4 | ccpTxbx | 0 | 文本框 |
| 0x68 | 4 | ccpHdrTxbx | 0 | 页眉文本框 |
| 0x98 | 2 | cbRgFcLcb | 164 / 183 | **不可信**，见 §2 |
| 0x9A | 8×N | fibRgFcLcb | | (fc, lcb) 对，索引 33 = fcClx |
| 0x1A2 | 4 | fcClx | 0x52A | CLX 在 table 流中的偏移 |
| 0x1A6 | 4 | lcbClx | 21 | CLX 字节数 |

`flags` 位：

| 位 | 名字 | 含义 |
|---|---|---|
| 2 | fComplex | 有 piece table（几乎总是 1） |
| 8 | fEncrypted | 文档加密（本仓库不处理） |
| 9 | fWhichTblStm | **1 → 用 `1Table`，0 → 用 `0Table`** |
| 12 | fExtChar | 使用扩展字符集 |
| 14 | fFarEast | 远东版式 |

hello.doc 是 `0x5014`：bit2=1、bit4=1、bit12=1、bit14=1，bit9=0 → `0Table`。
Microsoft 的 `PROTTPLN.DOC` 是 `0xF012`：bit9=1 → `1Table`。两个分支都要实现。

## 2. 坑：cbRgFcLcb 不可信

规范说 nFib=0xC1 时 `cbRgFcLcb = 0x005D`（93 对）。实测：

| 文件 | 写入者 | 0x98 处的值 |
|---|---|---|
| samples/hello.doc | WPS | 164 (0xA4) |
| PROTTPLN.DOC | Microsoft | 183 (0xB7) |

都不是 93，也不是 8 的倍数。结论：**不要用它算 fibRgFcLcb 的长度**，直接按 nFib
决定读多少对，或者按固定索引取字段。fcClx 永远是第 33 对（0 起算）：

```rust
const FC_CLX_PAIR: usize = 33;
let off = 0x9A + FC_CLX_PAIR * 8;   // 0x1A2
```

## 3. CLX：piece table

CLX 在 table 流的 `[fcClx, fcClx+lcbClx)`。开头可能有若干 `Prc`（clxt=0x01，
后跟 2 字节 cbGrpprl），真正的数据是 `Pcdt`（clxt=0x02）：

```
hello.doc 的 CLX（0Table 0x52A，21 字节）:
02           clxt = 0x02 (Pcdt)
10 00 00 00  lcb = 16   （PlcPcd 的字节数）
00 00 00 00  CP[0] = 0
52 00 00 00  CP[1] = 0x52 = 82      <- ccpText
00 00        PCD.flags
00 08 00 00  PCD.fc = 0x800
00 00        PCD.prm
```

`PlcPcd` = (n+1) 个 `u32` 的 CP 数组 + n 个 8 字节 PCD，`n = (lcb - 4) / 12`。
第 i 个 piece 覆盖字符区间 `[CP[i], CP[i+1])`。

### PCD.fc 的两个位

```
fc = 0x00000800   bit30 = 0 → UTF-16LE
fc = 0x40000000   bit30 = 1 → 8-bit(CP1252)，真实偏移 = (fc & 0x3FFFFFFF) / 2
```

**8-bit 情况下偏移要除以 2**，因为 Word 用"字节偏移 × 2"来统一表示两种编码的
位置（一个 8-bit 字符占半个"字符槽"）。漏掉这一步，读出来的全是乱码，而且不会报错。

## 4. 解码正文

```rust
for p in &doc.pieces {
    // p.cp_start..p.cp_end 是逻辑字符区间
    // p.fc 已经换算成 WordDocument 流里的字节偏移
    // p.compressed 决定 8-bit / UTF-16
}
```

- 8-bit：`decode_cp1252(&wd[p.fc .. p.fc + (cp_end-cp_start)])`
- UTF-16：`decode_utf16le(&wd[p.fc .. p.fc + 2*(cp_end-cp_start)])`

hello.doc 只有 1 个 piece：CP 0..82，fc=0x800，UTF-16LE，82 字符 = 164 字节
（0x800..0x8A4，正好是 fcMac）。

## 5. 控制字符

正文里夹着控制字符，直接打印会看到方块：

| 字节 | 含义 | 建议输出 |
|---|---|---|
| 0x0D | 段落结束 | `\n` |
| 0x07 | 单元格/行结束 | `\t` |
| 0x0B | 软回车（Shift+Enter） | `\n` |
| 0x0C | 分页/分节 | 丢弃或 `\n` |
| 0x13/0x14/0x15 | 域开始/分隔/结束 | 丢弃 |
| 0x01 | 图片 | 丢弃 |
| 0x02 | 脚注引用 | 丢弃 |
| 0x08 | 图形 | 丢弃 |

本仓库 `Doc::plain_text()` / `paragraphs()` 就是这么做的。

## 6. Word 6/95（nFib < 0xC1）

老文档没有 CLX（`fComplex` 通常为 0），正文就是 `WordDocument[fcMin..fcMac]` 的
原始 8-bit 字节，直接 CP1252 解码。本仓库用 `fib.n_fib < 0xC1` 走这条分支。

## 7. 其他流

| 流 | 内容 |
|---|---|
| `\x05SummaryInformation` | 属性集：标题、作者、页数、创建时间 |
| `Data` | 老版本用，现代 Word 常为空（hello.doc 的 size = 0） |
| `WpsCustomData` | WPS 私有，可忽略 |
| `1Table`/`0Table` | 见上 |

## 8. 字符与段落格式：CHPX / PAPX / sprm

正文的字符位置有了，格式还得另一套结构。Word 用**格式页**（formatting page，FKP）
把「字节偏移区间 → 属性」存成 512 字节一页，再用 PLC 把页号串起来：

| 索引（`fibRgFcLcb97`） | 结构 | 页内容 |
|---|---|---|
| 12 | `PlcfBteChpx` | CHPX 页：字符属性 |
| 13 | `PlcfBtePapx` | PAPX 页：段落属性 |

**注意**：这两个 PLC 的 `fc/lcb` 在 **WordDocument 流**的 FIB 里（`0x9A + 12*8`、`0x9A + 13*8`），
不在 table 流里。PLC 本身很小：`n = (lcb - 4) / 8`，`n+1` 个 `fc` 后跟 `n` 个页号
（`hello.doc` 两个 PLC 都是 `lcb = 12` → 1 页）。

格式页（以 PAPX 为例，`crun` 在 `page[511]`）：

```text
0x000  rgfc[(crun+1)]     u32 × (crun+1)   每段的起止字节偏移
       rgbx[crun]         13 B each        bOffset(u8) + PHE(12 B)
       page[bOffset*2..]  PAPX             cb(u8) → istd(u16) → grpprl
```

`cb == 0` 表示「长度放不下 1 字节」，真实长度是下一个字节 ×2。CHPX 页更简单：
`crun` 之后是 `crun+1` 个 `fc`、`crun` 个 `u8` 偏移（同样 ×2），每个 CHPX 是
`cb(u8) + grpprl`。

`grpprl` 是 sprm 序列。sprm 是 16 位：

```text
sprm = sgc(3) | spra(3) | sgc/spra 之外的 10 位操作码
spra = (sprm >> 13) & 7   → 操作数长度
  0,1 → 1 字节      2,4,5 → 2 字节      3 → 4 字节
  6   → 首字节是长度，后面是数据        7 → 3 字节
```

**关键结论（样本推翻了我的第一版实现）：CHPX / PAPX 条目之间不累加。**
每个条目描述的是「该区间完整的属性增量」，套用时要从一个空的属性集开始。
`formatted.doc` 的第 2 段是斜体+下划线、没有粗体，而它前面一段是粗体——如果按
「继承上一个条目的属性」实现，后面所有段落都会变粗。这个错误我在第一版里真的写出来了，
靠对照样本才发现。

`formatted.doc` 第 1 段的 CHPX 原样字节（第 5 页，`fc 0x800`）：

```text
42 2a 06                    sprmCIco   = 6        红色（16 色调色板索引）
70 68 ff 00 00 00           sprmCCv    = 0x0000FF RGB，sprmCCv 是 0x00BBGGRR
43 4a 24 00                 sprmCHps   = 0x24     36 半点 = 18 pt
35 08 01                    sprmCFBold = 1
6f 28 01 / 6e 48 04 08 / 74 48 04 08   WPS 私有 sprm，按 spra 长度跳过
```

常用 sprm（本仓库实现了这些）：

| sprm | 含义 | 类型 |
|---|---|---|
| `0x0835` | 粗体 | bool |
| `0x0836` | 斜体 | bool |
| `0x0837` / `0x2A53` | 删除线 / 双删除线 | bool |
| `0x2A3E` | 下划线类型（0 无 … 1 单线 … 10 波浪） | u8 |
| `0x4A43` | 字号（半点） | u16 |
| `0x6870` | 文字颜色 `0x00BBGGRR` | u32 |
| `0x2A42` | 调色板颜色索引 | u8 |
| `0x4A4F` | 字体序号（→ `SttbfFfn`） | u16 |
| `0x2A48` | 上下标（0 正常 1 上 2 下） | u8 |
| `0x2403` | 对齐（0 左 1 中 2 右 3 两端 4 分散） | u8 |
| `0x840E/0x840F/0x8411` | 右/左/首行缩进（twips） | i16 |
| `0xA413/0xA414` | 段前/段后间距（twips） | u16 |
| `0x2417` | 表格行结束标记 | bool |
| `0x260A` / `0x460B` | 列表级别 / 列表覆盖 | u8 / u16 |

**坑：WPS 把段落格式写进了样式表，PAPX 里一个 sprm 都没有。**
`formatted.doc` 的 6 个 PAPX 条目全部只有 `istd`、没有 `grpprl`（居中、缩进都在 STSH 里）。
段落格式的完整取值应当是「样式属性 + PAPX sprms」，所以只读 PAPX 会漏掉格式；
本仓库读得到 `istd`，再由 §8.1 的样式表拿到样式名。

**坑：`bOffset == 0` 的 PAPX 条目不是「无格式」，而是「同上一条」。**
`formatted.doc` 的 PAPX 页有 6 个 run，只有第 0 个带数据。Word 靠这个把连续同样格式的
段落压缩成一条；如果把它当成默认属性，第 2 段之后的样式名、缩进、对齐会全部丢失。
同理，PAPX 页里没有被任何 run 覆盖的段落标记，也继承上一段。

### 8.1 样式表：STSH / STD / Xstz

样式表位置是 `fibRgFcLcb97` 的**第 1 对**（`0x9A + 8`），在表流里：

| 偏移 | 字段 | 说明 |
|---|---|---|
| 0 | `cbStshi` u16 | **STSHI 结构自身的长度**（实测 1118 / 1062） |
| 2 | `cstd` u16 | 样式条目数（实测 18 / 15 / 19） |
| 4 | `cbSTDBaseInFile` u16 | 每个 `STD` 的固定头长度（实测 18） |
| 6 | `flags` u16 | bit0 `fStdStylenamesWritten` |
| 8 | `stiMaxWhenSaved` u16 | 保存时的内置样式上限（267 / 260） |
| 10 | `istdMaxFixedWhenSaved` u16 | 15 |
| 12 | `nVerBuiltInNamesWhenSaved` u16 | 7 / 4 |
| 14 | `rgftcStandardChpStsh[3]` u16×3 | 标准字符样式字体 |

`STD` 数组从 `2 + cbStshi` 开始，每条是 `cbStd u16` + `cbStd` 字节：

```text
STD = STDBase(cbSTDBaseInFile 字节) | Xstz(cch u16 + UTF-16 字符 + NUL u16) | grupe(UPX 数组)
```

所以样式名就在 `STD + cbSTDBaseInFile` 处。**`cbStd == 0` 表示该 `istd` 只有编号、
定义没写进文件**（Word 会省略仍处于默认值的内置样式），这一条没有名字。

```rust
let doc = office_doc::parse(&bytes)?;
assert_eq!(doc.style_name(0), Some("正文"));            // WPS 样例
assert_eq!(doc.style_name(0), Some("Normal"));          // PROTTPLN.DOC
for para in doc.paragraphs_rich() {
    println!("{:?} -> {:?}", para.text, para.style_name);
}
```

#### `grupe`：样式自己带的格式

名字后面（`Xstz` 的 NUL 之后）就是 `grupe`：一串 `UPX`，每个是 `cbUPX u16` + 这么多字节的
`grpprl`——和 `CHPX`/`PAPX` 里的格式**完全同一套 sprm 编码**，所以直接复用同一个解析器。
Word 会在数组前写一个 `cupx` 计数；WPS 不写，而且第一个 `cbUPX` 多算 2 个填充字节、
块之间还夹 1 个多余字节。

因此 `fmt::style_upx()` 用「**只接受能被 sprm 恰好走完的块**」这条判据做容错扫描：

```rust
for upx in fmt::style_upx(&body[rest..limit]) {
    fmt::apply_para_sprms(&mut style.para, upx);   // 段属性
    fmt::apply_char_sprms(&mut style.char, upx);   // 字符属性
}
```

两个 walker 各自只认自己的 sprm id，所以**不需要判断这块是段落还是字符属性**——喂给两边即可。

样式属性位于「直接格式」**之下**：`paragraphs_rich()` 先取样式属性，再把 `PAPX`/`CHPX`
的 sprm 合并上去（`CharProps::merge` / `ParaProps::merge`，后者覆盖前者）。

实测结果：

```text
$ officedump styles samples/formatted.doc
istd  name
0     正文  [justify=3 size=10.5pt]
10    默认段落字体
11    普通表格

$ officedump styles ".../PROTTPLN.DOC"
0     Normal  [size=12pt]
11    Table Normal
17    Balloon Text  [size=8pt]
```

`cargo run -p officedump -- styles samples/formatted.doc`：

```text
istd  name
0     正文
10    默认段落字体
11    普通表格
```

**规范对照**：MS-DOC 2.9.272 的 `STSHI` 第一个字段就是 `cstd`，但两个真实文件
（微软的 `PROTTPLN.DOC` 与 WPS 生成的样例）都在最前面多写了一个 `cbStshi`。
只有按「`2 + cbStshi` 起、`cbSTDBaseInFile` 偏移取名」走，`STD` 数组才**恰好**走完
`lcb` 个字节（1118+…=1642、1062+…=1298），否则第一步就跑到别人的数据里。
样式级格式（`STD` 里的 `UPX` sprm）本仓库尚未解析，只取了名字。

## 9. 文本故事（stories）：脚注、页眉、文本框

Word 把所有「故事」**拼成一条文本流**，靠 FIB 里的字符数切开：

```text
[ 主文档 ccpText ][ 脚注 ccpFtn ][ 页眉页脚 ccpHdd ][ 宏 ccpMcr ]
[ 批注 ccpAtn ][ 尾注 ccpEdn ][ 文本框 ccpTxbx ][ 页眉文本框 ccpHdrTxbx ]
```

这 8 个计数在 `fibRgLw97` 里，从 `0x4C` 开始每 4 字节一个（`0x4C` = ccpText，
`0x50` = ccpFtn，`0x54` = ccpHdd，`0x58` = ccpMcr，`0x5C` = ccpAtn，`0x60` = ccpEdn，
`0x64` = ccpTxbx，`0x68` = ccpHdrTxbx）。piece table 覆盖整条流，所以只要把
piece 文本按顺序拼起来，再按这些计数切片即可——不需要额外结构。

`samples/stories.doc`（页眉 `HEADER-MARK-11`、页脚 `FOOTER-MARK-22`、脚注
`FOOTNOTE-MARK-33`）实测：

```text
ccp_counts = [57, 19, 33, 0, 0, 0, 0]
0..57    "Body paragraph one.\rBody paragraph two.\rFOOTNOTE-MARK-33\r"
57..76   "FOOTNOTE-MARK-33 \r\r"
76..109  "HEADER-MARK-11\r\rFOOTER-MARK-22\r\r\r\r"
```

两个观察：

1. **页眉和页脚是同一个故事**（`ccpHdd`）。要分开得看 `Plcfhdd`（索引 11）：它的 CP
   序列把这条流切成 6 段，顺序是「首页页眉、偶数页页眉、奇数页页眉、首页页脚、偶数页
   页脚、奇数页页脚」，每段以 `\r` 结束。
2. **WPS 把脚注正文又写了一份进主文档区间**（字符 40..57），所以 `story(Main)`
   里会看到脚注文字。按 FIB 计数切分是规范做法（Word 自己就这么读），这里照做并记录
   这个现象。

## 10. 图片：Data 流 + PICF + OfficeArt

`.doc` 的图片**不在 `WordDocument` 里**，而在独立的 `Data` 流。链路是：

```text
CHPX ──sprmCPicLocation(0x6A03, 4 字节)──▶ Data 流的字节偏移
Data[偏移] = PICF 头（68 字节）
           ├ lcb        u32  @0   = 整个图片的总长度（含头部）
           ├ cbHeader   u16  @4   = 0x44（68）
           ├ mfHeader   u16  @6
           ├ bm_rcWinMF 8 B  @8   （WMF 边界矩形）
           ├ dxaGoal/dyaGoal/mx/my （显示尺寸）
           └ 裁剪、brcl …
PICF 头之后是 OfficeArt 记录：
   OfficeArtSpContainer (0xF004) ──▶ Sp(0xF00A) / FOPT(0xF00B) / ClientAnchor
   OfficeArtFBSE        (0xF007) ──▶ BLIP 记录 ──▶ 图片字节
```

实测 `samples/picture.doc`（WPS 插入一张 16×16 PNG）：

```text
Data 流 831 B
0x000  PICF  lcb=831  cbHeader=0x44
0x044  SpContainer 0xF004 len=138
0x0D6  FBSE        0xF007 len=609
         payload: btWin32 btMacOS rgbUid(16) tag size cRef foDelay … cbName
         0x102  BLIP 0xF01E len=565   ← 记录类型说 TIFF，其实是 PNG
         0x11A  rgbUid(16) + tag(0xFF) + "89 50 4E 47 0D 0A 1A 0A"
```

两点与 `.ppt` 完全一致，所以共用 `office-blip` crate：

1. **图片类型只看 magic bytes**，`0xF01E` 这类记录类型会说谎（这里明明是 PNG）；
2. **BLIP 负载前面有 16 字节 `rgbUid`**（有时再加 1 字节 tag），所以要扫描前几十字节找图片头。

```rust
let doc = office_doc::parse(&bytes)?;
for pic in doc.pictures() {
    println!("{:?} {} bytes", pic.kind, pic.data.len());   // Png 548 bytes
}
```

```console
$ cargo run -q -p officedump -- media samples/picture.doc out
out/image1.png  (Png, 548 bytes)
```

导出的 PNG 与插入的源文件**逐字节相同**。

## 10.1 表格

表格没有单独的流：它完全由**段落标记的属性**和**单元格标记 0x07** 表达。

- 每个单元格以 `0x07` 结尾，`0x07` 之前的文字就是单元格内容；
- 每行再多一个 `0x07`（这一条是"空单元格"），它的段落属性带 `sprmPFTtp`（`0x2417`）
  和 `sprmTDefTable`（`0xD608`，列宽）；
- 表格内的普通单元格段落带 `sprmPFInTable`（`0x2416`）。

所以读表格只要把 `0x07` 也当成段落边界，再看这两个 sprm：

```text
samples/table.doc 的文本流：
  "Before table.
"                     ← 不在表里
  "H1" "H2" "H3" ""   ← 第一行：3 个单元格 + 行结束
  "a"  "b"  "c"  ""   ← 第二行
  "After table.
"
```

```rust
let doc = office_doc::parse(&bytes)?;
for para in doc.paragraphs_rich() {
    if para.props.in_table == Some(true) {
        let row_end = para.props.ttp == Some(true);
        println!("{:?} row_end={row_end}", para.text);
    }
}
```

`officedump html` 会把连续的 `in_table` 段落渲染成 `<table>`，行结束标记不开 `<td>`：

```html
<p style="text-align:justify;">Before table.</p>
<table>
<tr><td>H1</td><td>H2</td><td>H3</td></tr>
<tr><td>a</td><td>b</td><td>c</td></tr>
</table>
<p style="text-align:justify;">After table.</p>
```

列宽在行结束标记的 `sprmTDefTable`（`0xD608`）里，实测布局是：

```text
TDefTable = <未文档化的 1 字节> itcMac(u8) dxaCenter[itcMac+1](i16, twips) TC[itcMac](20 字节)
```

`samples/table.doc` 的实测操作数是 `00 03 94 ff ac 0a c5 15 de 20 …` → `itcMac = 3`，
边界 `[-108, 2732, 5573, 8414]`（相邻差 2841 twips ≈ 142pt，三列等宽）。

两个坑：

1. `itcMac` 是 **1 字节**，按 `u16` 读会得到 768，边界全部错位；
2. 声明长度可能**大于实际字节数**（这一条声明 70 字节，实际只写了 57 字节），
   解析器必须夹取而不是整条丢弃，否则整张表都读不出来。

`sprm` 走查里的夹取就一行：

```rust
let end = (start + n).min(grpprl.len());
```

WPS 还会额外写 `0xD634`（本实现不解释，只跳过），所以 HTML 里除了 `<colgroup>` 的宽度
之外不依赖它：

```html
<table>
<colgroup><col style="width:142pt"><col style="width:142.05pt"><col style="width:142.05pt"></colgroup>
<tr><td>H1</td><td>H2</td><td>H3</td></tr>
<tr><td>a</td><td>b</td><td>c</td></tr>
</table>
```

## 11. 加密文档

`flags` 的 bit 8（`0x0100`，fEncrypted）置位时，`WordDocument` 与表流都按 **512 字节块**
用 RC4 加密，但 `WordDocument[0..0x44)` 的 FIB 头保持明文（否则连 `nFib` 都读不到）。
加密头（52 字节旧式 / CryptoAPI 变长）放在**表流的开头**，表流的选择仍然看 bit 9。

```rust
let doc = office_doc::parse_with_password(&bytes, "secret123")?;
```

细节（密钥派生、`(h0[:5]+salt)*16`、连续 RC4 校验）见 [`encryption.md`](encryption.md)。

## 12. 已知坑

1. `fWhichTblStm` 选错流 → CLX 位置读到垃圾，表现为 piece 数量离谱。
2. `cbRgFcLcb` 不可信（§2）。
3. 8-bit piece 的 `fc` 要 `/2`（§3）。
4. `fcMin/fcMac` 只对老文档/无 piece 的情况有效，**有 piece table 时以 piece 为准**。
5. `ccpText` 只是主文档字符数，脚注/页眉/文本框在 ccpFtn/ccpHdd/ccpTxbx 里，CP 区间
   在 `ccpText` 之后继续延伸；只取 `[0, ccpText)` 会漏掉它们。
6. 加密文档（flags bit8）的 FIB 之外全部被加密，`fcClx` 指向垃圾——先判断再解析。
7. CHPX/PAPX 条目之间**不累加**（§8），这是最容易写错的地方。
8. 段落格式可能只存在于样式表里（WPS 就是这样），PAPX 空不代表没有格式（§8）；
   `bOffset == 0` 的 PAPX 条目表示「同上一条」，不是「无格式」。
9. 只取 `[0, ccpText)` 会丢掉脚注/页眉/文本框，它们在同一条文本流后面（§9）。
10. **PAPX 页的 `BX` 是 13 字节**（1 字节 `bOffset` + 12 字节 `PHE`），CHPX 页才是 1 字节。
    按 1 字节步长读 `bOffset` 会让每个 run 都拿到**第一个 run** 的属性——单段落文档
    测不出来，多段落文档里所有段落的格式会整体"漂移"到错误的段落上（本仓库踩过）。
11. 表格不是独立结构，而是 `0x07` 单元格标记 + `sprmPFInTable` / `sprmPFTtp`（§10.1）。
12. **压缩片段的 `fc` 是否已经折半，两个写入器不一致**：WPS 写 `0x800`（已经是字节偏移），
    微软写 `0x40001000`（`bit30` 只表示 `fCompressed`，偏移还得 `/2`）。
    拿 `fc & 0x3FFFFFFF` 直接和 CHPX/PAPX 页里的字节偏移比较，微软写的文件会
    **一条 run 都命中不了**——格式全丢、表格全不认，而且不报错（本仓库踩过）。
    正确做法：`let raw = fc & 0x3FFFFFFF; let piece_fc = if compressed { raw / 2 } else { raw };`
13. **`ToggleOperand` 的 `0x81` 是「取反」，不是「关」**（MS-DOC 2.6.1）：
    `0`=关、`1`=开、`0x80`=保持原状、`0x81`=相对默认值取反。微软 Word 对**每一个**加粗 run
    都写 `sprmCFBold = 0x81`，按「非零且不等于 0x81 才算开」读，微软写的文件会全部丢粗体/斜体。

## 13. Rust API

```rust
use office_doc::parse;

// 入参是整份 .doc 文件（CFB 容器）的字节，不是 WordDocument 流
let doc = parse(&std::fs::read("samples/hello.doc")?)?;
println!("nFib={:#06X} pieces={}", doc.fib.n_fib, doc.pieces.len());
println!("{}", doc.plain_text());
for p in doc.paragraphs() { println!("{p}"); }
for piece in &doc.pieces {
    println!("cp {}..{} fc {:#X} compressed={}",
        piece.cp_start, piece.cp_end, piece.fc, piece.compressed);
}

// 带格式的段落：每段有段落属性 + 若干字符 span
for p in doc.paragraphs_rich() {
    println!("{:?} justify={:?} 缩进={:?}/{:?}",
        p.text, p.props.justify, p.props.dxa_left, p.props.dxa_left_first);
    for span in &p.runs {
        println!("  {:?} bold={:?} italic={:?} size={:?} color={:?}",
            span.text(&p.text), span.props.bold, span.props.italic,
            span.props.size_half_points, span.props.color);
    }
}

// 各故事分开取：主文档 / 脚注 / 页眉页脚 …
for (story, text) in doc.stories() {
    println!("--- {} ({} 字符) ---
{}", story.name(), text.chars().count(), text);
}
```

`parse(bytes)` 内部先走 `office_cfb`，再按 `fWhichTblStm` 自己挑 `0Table` / `1Table`，
调用方不用关心。

CLI：

```console
$ cargo run -q -p officedump -- text samples/hello.doc
$ cargo run -q -p officedump -- info samples/hello.doc --json
$ cargo run -q -p officedump -- html samples/formatted.doc > /tmp/fmt.html   # 带格式
$ cargo run -q -p officedump -- stories samples/stories.doc                  # 分故事
```
