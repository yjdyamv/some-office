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

1. **页眉和页脚是同一个故事**（`ccpHdd`）。要分开得看 `Plcfhdd`（索引 11 = FIB
   `0xF2`）：它的 CP 序列把这条流切成槽位，每槽以 `\r` 结束。槽位不是六段而是
   「六个文档级分隔槽 + 每节六槽 + 结尾槽」，节内顺序是「偶数页头、primary 头、
   偶数页脚、primary 脚、首页头、首页脚」——详见 §14.7（用 Word 逐槽造样本实测）。
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

## 14. 写：自己生成一个 Word 能打开的 .doc

`office_doc::write_minimal(text) -> Vec<u8>` 从零生成一个 Word 97 文档，只依赖
`office-cfb` 的写入器。目标不是"能写多少"，而是**把读模型反过来验证一遍**：读和写
两侧对同一批字节达成一致，才说明字段语义没理解错。

### 14.1 Word 到底强制要求什么（切除实验）

拿 `samples/hello.doc` 逐对把 FIB 的 `fc/lcb` 清零，再用自己的 CFB 写入器重打包，
交给 Word 打开：

| 清掉的对 | 索引 | Word 的反应 |
|---|---|---|
| `PlcfBteChpx` + `PlcfBtePapx` | 12, 13 | **挂死**（COM 调用 90 秒无返回） |
| 只清 `PlcfBteChpx` | 12 | 挂死 |
| 只清 `PlcfBtePapx` | 13 | 挂死 |
| `StshfOrig` + `Stshf` | 0, 1 | 打开失败 |
| `PlcfSed` | 6 | 打开失败 |

结论：格式页（FKP）不是"可选优化"，**没有它 Word 连文件都不打开**；样式表与节表
则是硬性字段。反过来，`Data`、`SummaryInformation`、`CompObj` 全都可以不要。

### 14.2 写出来的结构

| 位置 | 内容 |
|---|---|
| `WordDocument[0x000..0x380]` | FIB：`wIdent 0xA5EC`、`nFib 0x00C1`、`flags 0x1204`（fComplex｜fWhichTblStm｜fExtChar）、`cbRgFcLcb 0x005D` |
| `WordDocument[0x800..]` | 文本，UTF-16LE，`
` 作段落标记 |
| `WordDocument` 下一页 | CHPX FKP：`crun 1`，`rgfc = [fcMin, fcMac]`，`rgb[0] = 0`（无字符属性） |
| `WordDocument` 再下一页 | PAPX FKP：`crun = 段落数`，每段一条 `rgfc`，`BX` 步长 13 字节 |
| `1Table` | CLX（单片段）→ STSH（只有 `Normal`）→ PlcfSed（一节）→ PlcfBteChpx/Papx |

`flags` 里的 **fComplex（bit 2）必须置位**：只要存在 CLX，Word 就按"复杂文档"读，
否则它会忽略 CLX、把 `fcMin..fcMac` 当成一整段 8 位文本。

### 14.3 三个踩过的坑

1. **`rgfc` 是 UTF-16 码元偏移，不是 UTF-8 字节偏移。** 用
   `text.char_indices()` 的字节下标算 `fc` 时，一个汉字占 3 字节而流里只占 2，
   于是段落边界越算越大（实测出现 `rgfc = 2198 > fcMac = 2146`），Word 只读出第一段。
   正确写法是按 `ch.len_utf16()` 累计：
   ```rust
   let mut units = 0u32;
   for ch in text.chars() {
       units += ch.len_utf16() as u32;
       if ch == '
' { marks.push(TEXT_AT as u32 + units * 2); }
   }
   ```
2. **`cbMac` 要包含格式页。** 它等于 `WordDocument` 整个流的长度，不是文本结尾。
3. **PAPX 的 `BX` 是 13 字节**（`bOffset` + 12 字节 `PHE`），一个 512 字节页最多放
   29 条；超过就退回"整篇一条"，让 Word 自己按段落标记切分。

### 14.4 验证

```bash
cargo run --example mkfiles      # 生成 samples/written.doc / written.xls
pwsh tools/roundtrip.ps1         # 交给真的 Word / Excel 打开
```

```
word   paragraphs=3 chars=57
  [1] WRITER-DOC-001
  [2] Second line from the writer.
  [3] END-WRITER-9
excel  sheet=Writer used=A1:A4
```

`cargo test` 里另有 `writes_the_formatting_pages_word_insists_on`：把生成的文件用
自己的读取器走一遍，断言 `para_runs.len() == 3`、段落边界是
`[0x808, 0x810, 0x814]`（UTF-16 码元），这样"文件里有格式页"和"文件声称有格式页"
就能区分开。

### 14.5 写表格

表格不是一种"对象"，就是一批**以 `U+0007` 结尾的普通段落**：每个单元格一个段落，
行末再来一个空段落，它带 `sprmPFTtp` 和 `sprmTDefTable`（列宽）。读取侧
（§10.1）就是这么认的，写入侧因此不需要新机制，只要让段落标记变成 `0x07`。

```rust
let table = office_doc::Table::new(&[&["H1", "H2", "H3"], &["a", "b", "c"]]);
let bytes = office_doc::write_table_document("Before table.", &table, "After table.");
```

生成的文本流与 Word 自己写的一模一样：

```text
Before table.\r
H1\x07H2\x07H3\x07\x07      <- 三个单元格 + 行末标记
a\x07b\x07c\x07\x07
After table.\r
```

每个表格段落的 PAPX 必须带三条 sprm，少一条 Word 就拒绝打开：

| sprm | 含义 | 值 |
|---|---|---|
| `0x2416` | `sprmPFInTable` | 1 |
| `0x6649` | `sprmPItap`（表格深度） | 1 |
| `0x2417` | `sprmPFTtp`（只在行末段落） | 1 |

行末段落另外带 `0xD608`（`sprmTDefTable`）。它是 spra 6 的可变长 sprm，**操作数前面
要有一个长度字节**（`08 d6 | 46 | …`），操作数本身是：

```text
00            <- 未文档化的前导字节，实测恒为 0
03            <- itcMac（列数，1 字节）
94 ff ac 0a c5 15 de 20   <- itcMac+1 个单元格边界（twips）
<itcMac × 20 字节 TC>     <- 全 0 = 无边框、无合并
```

三个坑：

1. **`sprmPItap` 不能省**。只写 `fInTable` + `fTtp` 时我们的读取器能读回表格，
   但 Word 报「值不在预期的范围内」并拒绝打开 —— 这是必须靠真实程序才能发现的。
2. **`0xD608` 的长度字节**。用通用的 `push_sprm` 直接塞操作数会让读取器把首字节
   （那个前导 0）当成长度，列宽变成空数组。
3. **列宽是"边界"不是"宽度"**：`itcMac + 1` 个从左边距起算的坐标，
   因此 3 列 4 个值。

验证（`tools/roundtrip.ps1`）：

```text
word   table: paras=10 rows=2 cols=3
  [1] H1 | H2 | H3
  [2] a | b | c
```

限制：不写边框（`TC` 全 0）、不写合并单元格、不写表格样式。

### 14.6 写图片：`Data` 流 + `sprmCPicLocation`

`.doc` 的图片不放在 `WordDocument` 里，而是集中在 **`Data` 流**中；正文里只有一个 `` 字符，
它的 CHPX 用两个 sprm 指向图片：

| sprm | 名字 | 操作数 |
|---|---|---|
| `0x6A03` | `sprmCPicLocation` | `u32`，图片在 `Data` 流里的偏移 |
| `0x0855` | `sprmCFSpec` | `u8` = 1（这个字符是"特殊字符"） |

`Data` 流里一条图片的结构（`samples/picture.doc` 实测，共 831 字节）：

| 偏移 | 内容 | 说明 |
|---|---|---|
| `0x000` | `PICF`（68 字节） | `lcb`@0 = 831（含自己）、`cbHeader`@4 = 0x44、`mm`@6 = 100、`dxaGoal`@0x1C / `dyaGoal`@0x1E = 显示尺寸（twips） |
| `0x044` | `OfficeArtSpContainer`（`0xF004`，146 字节） | 形状属性：`Sp`、`FOPT`、`ClientAnchor` |
| `0x0D6` | `OfficeArtFBSE`（`0xF007`，617 字节） | 36 字节字段（rgbUid、tag、size、cRef…）+ 内嵌的 BLIP |
| `0x102` | BLIP（565 字节） | `ver_inst 0x6E00`、**`recType 0xF01E`（TIFF 编号，装的却是 PNG）**、`rgbUid` 16 字节、`0xFF`、图片本体 |

`crates/office-doc/src/write.rs` 的 `Run::picture(...)` 就是这么写的：模板的 831 字节
（`crates/office-doc/templates/picture-data.bin`）里只替换 BLIP 段，偏移按 `214 + 44 + 25 + 图片长度` 重算，
`PICF.dxaGoal/dyaGoal` 填显示尺寸。Word 实测：

```text
word   written-picture.doc: paras=3 words=15 chars=51 shapes=0 inline=1
  inline shape type=3 w=24.0 h=24.0
```

`type=3` 是 `wdInlineShapePicture`，24 pt = 480 twips = 我们写进 `dyaGoal` 的值。

三个坑：

1. **Escher 记录头是 `[ver_inst][kind][len]`**，不是 `[kind][ver_inst][len]` —— 写反了整体错位 2 字节。
2. **`OfficeArtFBSE` 后面还有 2 个多余字节**（`62 00`），所以内嵌 BLIP 的偏移是 `0xD6` 而不是 `0xD8`；
   读取器因此改成"从 `payload[34]` 起扫描 BLIP 魔数"，不再依赖 `cbName` 推算。
3. **`recType` 会说谎**：Word 把 PNG 写在 `0xF01E`（TIFF）里，判断图片格式只能看魔数。

### 14.7 写页眉和页脚

页眉页脚不是一个对象，而是文本流末尾的一段故事（`ccpHdd`），再靠 `PlcfHdd`
的 CP 序列切成槽位。

```rust
let header = vec![Paragraph::text("Draft header")];
let footer = vec![Paragraph::text("Draft footer")];
let bytes = office_doc::write::write_document_with(
    &body,
    &office_doc::write::HeaderFooter { header: &header, footer: &footer },
);
```

#### 槽位顺序（实测，与规范措辞不同）

拿 Word 2016 单独设每一个 `Headers.Item(i)` / `Footers.Item(i)` 造六份样本
（`samples/ms/hdd-{hdr,ftr}-{primary,firstpage,evenpages}.doc`）加一份双节样本
（`hdd-twosections.doc`），用 `examples/hdrdump.rs` 读出 CP 与逐槽文本：

```text
hdd-hdr-primary.doc   cps = [0,3,6,6,9,12,12,14,23,25,34,36,38,38]
  slot 1  "MARK-hdr-primary"      <- primary 页眉在 slot 1，不是 slot 0
  slot 3  "MARK-ftr-primary"      <- primary 页脚在 slot 3
twosections.doc       cps(20) = [0,3,6,6,9,12,12,14,23,25,34,36,38,38,38,38,38,38,38,41]
  slot 7  "MARK-h2"
  slot 9  "MARK-f2"
```

顺序是：

| CP 段 | 内容 |
|---|---|
| 0..5 | 六个**文档级分隔槽**：`U+0003` + 两个 `\r`、`U+0004` + 两个 `\r`、空段，重复一遍 |
| 6..11 | 每节六槽：偶数页头、**primary 头**、偶数页脚、**primary 脚**、首页头、首页脚 |
| 末段 | 一个结尾槽（三个 `\r`） |

所以单节文档要写 **13 个 CP**（6 个分隔 + 6 个节槽 + 结尾）。只写 7 个（一节六槽 +
结尾）会被 Word 判定为文件损坏。

#### 五条硬性要求

1. **CP 是页眉故事内的相对偏移**，写绝对偏移会让 Word 报「Word 无法读取文档中的
   页眉/页脚」。
2. **每个槽至少要有一个段落符**。空槽必须写成 `U+0003`+`\r\r`、`U+0004`+`\r\r`
   或 `\r\r`，零长度槽同样被拒。
3. **`ccpHdd` 恰好比页眉故事的文本短 2，最后一个 CP 落在故事末尾**。Word 自己
   写 `ccpHdd=49`、末 CP 51（故事 51 字符）。把 `ccpHdd` 写成故事长度（51）会让
   Word 打开文档时**没有页眉**，而正文完全正常 —— 只有 `tools/check_header.ps1`
   能看出来。
4. **piece table 必须按故事切分**。一条 piece 跨过故事边界会让正文全变成乱码，
   CLX 因此要写两条 piece（正文一条、页眉故事一条），PCD 的 `fc` 各自加偏移。
5. **没有页眉时不要写这个故事**。Word 接受 `ccpHdd=0`，此时 `PlcfHdd` 也应该不写
   —— 我们的写入器只在任一槽非空时才生成它，普通文档的字节因此完全不变。
   但**有脚注或尾注时必须写**：`ccpHdd=0` 时 Word 会报「无法读取页眉/页脚」。

#### 有尾注时页眉会消失（已知取舍）

尾注故事必须从 `ccp*` 之和开始，而那个位置正是 `ccpHdd`；页眉却要求故事比
`ccpHdd` 多两个字符。两者不可兼得：

- 让两个故事重叠（尾注从 `ccpHdd` 起、页眉故事仍多两个字符）→ Word 判定文件损坏；
- 让尾注从物理末尾起（不重叠）→ `ccp*` 之和对不上，Word 同样判定损坏；
- 把 `ccpHdd` 提到故事长度（不重叠且账目自洽）→ 文件有效、尾注可读，**页眉不显示**。

写入器选第三种：有尾注时页眉故事写成 `ccpHdd` 长度、末 CP 等于 `ccpHdd`。
`write::tests::an_endnote_story_starts_where_the_counts_put_it` 把这个形状钉住，
`docs/doc.md` 这里记下取舍。真 Word 复核（`tools/check_header.ps1` + `check_endnote.ps1`）：

```text
--- hdronly   section 1 header: [Draft header\r]        <- 没有尾注，页眉正常
--- hdrftn   section 1 header: [Draft header\r]        <- 有脚注也正常
--- hdredn   section 1 header: []   endnote 1: [endnote text]   <- 有尾注，页眉不显示
```

读取侧对应 API：

```rust
let doc = office_doc::open("samples/ms/hdd-hdr-primary.doc")?;
assert_eq!(doc.hdd_slots.len(), 14);
assert_eq!(doc.header_footer(0).unwrap().0, "MARK-hdr-primary");
```

`Doc::hdd_story_slots()` 返回每个槽的原始文本（含分隔槽），`Doc::header_footer(section)`
取出该节的 primary 页眉与页脚，`Doc::hdd_slots` 保留 `PlcfHdd` 的原始 CP。

真 Word 复核（`tools/check_header.ps1`）：

```text
app: C:\Program Files\Microsoft Office\Root\Office16
sections=1 pages=1
  body: HEADER-DOC-123 body text.
  body: Second body paragraph.
  section 1 header: [Draft header\r]
  section 1 footer: [Draft footer\r]
```

### 14.8 写脚注

脚注是文本流里的第二段故事（`ccpFtn`，紧跟正文），靠两张 PLC 与正文挂上钩：
`PlcffndRef` 说每个标记落在正文的哪个 CP，`PlcffndTxt` 说每条注释的文本落在脚注
故事里的哪个位置。

```rust
use office_doc::write::{write_document_with_footnotes, Documents, Paragraph, FOOTNOTE_MARK};
let body = vec![
    Paragraph::text(format!("Body one carries a note.{FOOTNOTE_MARK}")),
    Paragraph::text("No note here."),
];
let notes = vec![vec![Paragraph::text("the note")]];
let bytes = write_document_with_footnotes(&body, &Documents::default().notes(&notes));
```

调用方负责把 `FOOTNOTE_MARK`（`U+0002`）放进正文段落，一条注释一个；写入器扫正文
找这些字符，CP 本来就对，不需要额外传位置。`Documents` 同时带页眉、页脚与脚注，
三者一起写才保证字符计数与 PLC 一致：

```rust
write_document_with_footnotes(&body, &Documents::new(&header, &footer).notes(&notes));
```

#### 两张 PLC 的形状（实测）

用 Word 2016 的 `Footnotes.Add` 造两份样本（`samples/ms/footnote2.doc` 一条注释、
`footnote.doc` 两条），再读 Word 自己写的 CP：

```text
           ccpText ccpFtn ccpHdd  PlcfRef CP        FRD     PlcfTxt CP
footnote2   37       12     13    [17, 63]          [1]     [0, 11, 14]
footnote   101       51     13    [31, 75, 166]     [1, 1]  [0, 13, 50, 53]
```

1. **`PlcffndRef` 的 CP 是标记在正文里的位置**，最后再加一个收尾 CP = 所有故事
   长度之和 + 1（`ccpText + ccpFtn + ccpHdd + 1`，上面是 63 与 166）。每个 CP 后面
   跟一个 2 字节 `FRD`，所以 `lcb = 4k + 2(k-1)`：`lcb=10` 是 `[17, 63]` 加一个
   `FRD`。**不能把整段 `lcb` 当 u32 数组读** —— 那样每个 `FRD` 会变成一个假 CP
   （`[17, 63]` 读成 `[17, 65537]`）。
2. **`FRD` 的 bit 0（`fAuto`）为 1 表示自动编号，0 表示自定义标记**。自动编号的
   注释在正文里放 `0x02` 占位、`FRD` 写 1；自定义标记（如 `*`、`†`）则把那个字面
   字符同时写进正文与脚注故事开头，`FRD` 写 0（见 §14.8.1）。
3. **`PlcffndTxt` 是裸 CP 数组**，相对脚注故事起点：一条注释一个起点，然后一对收尾
   CP（`lcb = 4(n+2)`）。最后那个 CP 落在 `ccpFtn + 2`，比故事本身还长两个字符。
4. **`ccpFtn` 是整段故事的长度**。每条注释是「标记 + 文本 + 段落符」，故事末尾还有
   一个段落符；`ccpFtn` 比最后一个起点多 1。少写一个就是「每条脚注少显示最后一个
   字符」（实测 `second note, longer than the firs`）。
5. **`Plcfsed` 的 CP[1] 必须等于 `PlcffndRef` 的收尾 CP**。有子故事时它是
   `ccpText + ccpFtn + ccpHdd + 1`，没有子故事时才是 `ccpText`。两者差 1 时 Word
   的报错会指向页眉（「Word 无法读取文档中的页眉/页脚，请将文档另存为 RTF 格式…」），
   与真正出问题的 PLC 毫无关系 —— 这条浪费了最久。

#### 三条连带要求

1. **故事顺序 = 正文 → 脚注 → 页眉**，所以 CLX 要按故事切成多条 piece（正文一条、
   脚注一条、页眉一条）。跨故事的一条 piece 会让 Word 读到乱码。
2. **有脚注时必须有页眉故事**。Word 的 `footnote2.doc` 自带 13 个字符的纯分隔符
   故事；写成 `ccpHdd = 0` 会被拒收并甩出页眉错误。写入器因此在有脚注时也生成
   分隔槽（`PlcfHdd` 的 CP 相对页眉故事起点 = `ccpText + ccpFtn`）。
3. **没有脚注时两对 FIB 偏移保持 0/0**，普通文档的字节因此完全不变。

#### 读侧 API

```rust
let doc = office_doc::open("samples/ms/footnote.doc")?;
assert_eq!(doc.footnote_slots, vec![31, 75, 166]);       // 正文里的标记位置
assert_eq!(doc.footnote_text_slots, vec![0, 13, 50, 53]); // 脚注故事内的切点
assert_eq!(doc.fib.ccp_counts[1], 51);
assert!(doc.story(office_doc::Story::Footnote).contains("second note"));
```

#### 真 Word 复核（`tools/check_footnote.ps1`）

```text
app path: C:\Program Files\Microsoft Office\Root\Office16
paras=4 footnotes=2 endnotes=0
  para: [WRITER-DOC-001 carries a note.]  codes=...,2,13
  footnote 1: [note one]
  footnote 2: [note two, with more words]
```

Word 会在文末补一个空段落（`paras = 段落数 + 1`），它自己的样本也一样。同一份文件
再叠页眉页脚（`Documents::new(...).notes(...)`）也能被接受，正文、脚注、页眉三者
同时正确。

#### 14.8.1 自定义脚注标记

默认注释是自动编号的：正文里放 `FOOTNOTE_MARK`（`0x02`），`FRD.fAuto = 1`，
Word 排版时算出数字。要换成自定义标记（`*`、`†`、`§`…），用
`Documents::note_marks(&[Option<char>])` 给每条注释一个标记，`None` 表示那条仍自动
编号：

```rust
use office_doc::write::{write_document_with_footnotes, Documents, Paragraph, FOOTNOTE_MARK};
let body = vec![
    Paragraph::text(format!("Starred.{FOOTNOTE_MARK}")),
    Paragraph::text(format!("Numbered.{FOOTNOTE_MARK}")),
];
let notes = vec![vec![Paragraph::text("star note")], vec![Paragraph::text("auto note")]];
let marks = [Some('*'), None];
let documents = Documents::default().notes(&notes).note_marks(&marks);
let bytes = write_document_with_footnotes(&body, &documents);
```

写入器做两件事：把第 i 条自定义标记的字符覆盖到正文里第 i 个 `FOOTNOTE_MARK` 的
位置，并在脚注故事里该条的开头写同一个字符（替换 `0x02`）；`FRD.fAuto` 对应写 0。
自动编号的那条不动。`samples/ms/footnote-custom.doc`（Word 2016 自己写、一条 `*`
一条自动）实测：`FRD = [0, 1]`；正文里是字面 `*`；脚注故事是
`* star note\r\u{2} auto note\r\r`。读侧 `Doc::footnote_marks()` 按 `FRD.fAuto`
回读：自定义的返回那个字符，自动的返回 `None`（`[Some('*'), None]`）。

单测 `write::tests::custom_footnote_marks_ride_in_the_text` 与读侧
`tests::reads_custom_footnote_marks_from_the_sample`；`tools/check_footnote.ps1`
读回 `footnote 1: ref=*`、`footnote 3: ref=†`。

### 14.9 写尾注

尾注和脚注是同一个机制的两次实例：一张 PLC 列出正文里每个标记的 CP
（`PlcffndRef` / `PlcfendRef`），另一张把注释故事切成一条条文本
（`PlcffndTxt` / `PlcfendTxt`）。区别只在**FIB 里占哪一对偏移**，以及故事排在
文本流的哪一段。

```rust
use office_doc::write::{
    write_document_with_footnotes, Documents, Paragraph, ENDNOTE_MARK, FOOTNOTE_MARK,
};
let body = vec![
    Paragraph::text(format!("Body one carries a footnote.{FOOTNOTE_MARK}")),
    Paragraph::text(format!("Body two carries an endnote.{ENDNOTE_MARK}")),
];
let notes = vec![vec![Paragraph::text("footnote text")]];
let endnotes = vec![vec![Paragraph::text("endnote text")]];
let documents = Documents::new(&[], &[]).notes(&notes).endnotes(&endnotes);
let bytes = write_document_with_footnotes(&body, &documents);
```

#### 五条硬性要求

1. **尾注的 PLC 在 FIB 的第 46、47 对**（标记 `fc = 0x9A + 46*8 = 0x20A`、
   `lcb = 0x20E`；文本 `0x212` / `0x216`）。第 4、5 对是 `PlcfandRef` /
   `PlcfandTxt`（批注），结构一模一样 —— 写成那一对，Word 不报错，尾注整个消失。
2. **脚注标记和尾注标记在文本流里都是 `U+0002`**，没有任何一个字节区分它们，
   唯一的区分就是「哪个 PLC 列了这个 CP」。所以调用方用 `U+0005` 标注尾注，
   写入器落笔前改写成 `U+0002`：

   ```rust
   let mark = FOOTNOTE_MARK as u16;
   utf16[character * 2..character * 2 + 2].copy_from_slice(&mark.to_le_bytes());
   ```

3. **故事顺序 = 正文 → 脚注 → 页眉 → 尾注**。`ccp*` 是累加的，读者按
   `ccpText + ccpFtn + ccpHdd + ccpEdn` 定位，所以尾注必须排在页眉之后，CLX 也要
   为它单独切一条 piece（跨故事的 piece 会让正文变乱码）。
4. **`Plcfsed` 的 CP[1] 要算进 `ccpEdn`**：`ccpText + ccpFtn + ccpHdd + ccpEdn + 1`。
   少算这一段，Word 又会甩出那条与真正病因无关的「无法读取文档中的页眉/页脚，
   请将文档另存为 RTF 格式…」。
5. **`ccpEdn` 是整段尾注故事的长度**，`PlcfendTxt` 的末 CP 落在故事末尾 + 2（与
   脚注完全同形）。

#### 实测（`examples/hdrdump.rs` 现在也打这两张 PLC）

```text
samples/ms/endnote2.doc         ccp=[37,  0, 13, 0, 0, 15, 0, 0]
  PlcfendRef cps=[17, 66]     PlcfendTxt cps=[0, 14, 17]
samples/ms/both3.doc            ccp=[59, 17, 13, 0, 0, 16, 0, 0]
  PlcfendRef cps=[37, 106]    PlcfendTxt cps=[0, 15, 18]
samples/written-endnotes.doc    ccp=[91, 16, 25, 0, 0, 15, 0, 0]
  PlcfendRef cps=[62, 148]    PlcfendTxt cps=[0, 14, 17]
```

`both3.doc` 是 Word 自己写的「脚注 + 尾注」样本（脚注标记在 CP 18、尾注标记在
CP 37），三张表的形状与我们的输出一致；`sed_end = 91 + 16 + 25 + 15 + 1 = 148`
对得上 `PlcfendRef` 的收尾 CP。注意 `written-endnotes.doc` 的 `ccpHdd` 是 25 ——
有尾注时页眉故事不写那多出来的两个字符，原因见 §14.7「有尾注时页眉会消失」。

#### 一次「Word 一直转圈」的排查

第一版尾注写完后，Word 打开 `samples/written-endnotes.doc` 时界面一直转圈、
没有任何报错。病因不在尾注结构里，而在**段落格式页的 run 覆盖**：

1. 注释的引用标记是单独 push 进文本流的，于是它前面的段落停在标记之前，标记本身
   不属于任何段落 —— `para_ends` 里留下 1 个字符的洞（`gap 2154..2156`、
   `gap 2212..2214`）。
2. 格式页的 `rgfc` 数组只写**第一段的起点**和**每段的终点**，所以页内的洞看不
   出来；只有当分页边界正好落在洞上，后一页的第一个 `rgfc` 才比前一页的最后一
   个多 2 个字节。Word 读到不接续的 `rgfc` 就无限循环。
3. 修法：写完所有故事后把 run 拉平，把标记归给它后面那个段落。

   ```rust
   /// Make the paragraph runs cover `fc_min..fc_mac` without holes.
   fn close_gaps(para_ends: &mut [(u32, u32, ParaProps)], fc_min: u32, fc_mac: u32) {
       if para_ends.is_empty() {
           return;
       }
       para_ends[0].0 = fc_min;
       for i in 1..para_ends.len() {
           if para_ends[i].0 != para_ends[i - 1].1 {
               para_ends[i].0 = para_ends[i - 1].1;
           }
       }
       let last = para_ends.len() - 1;
       if para_ends[last].1 < fc_mac {
           para_ends[last].1 = fc_mac;
       }
   }
   ```

顺带把格式页的布局改成 Word 的样子：**数据紧跟在数组后面**，不再固定在
`0x100`。原来是「数组 `0x00..0x100`、数据 `0x100..0x1FF`」，一页只能装 14 段，
分页边界因此常常落在任意位置；现在装得下就装（`array_len(n) + data <= 511`），
「3 段正文 + 1 条脚注 + 1 条尾注 + 12 个页眉分隔槽」这种组合正好挤进一页。
两条一起改，`samples/written-endnotes.doc` 才被 Word 正常打开
（`paras=3 footnotes=1 endnotes=1`）。

> **验证陷阱**：Word 把打开失败的文件记在
> `HKCU\Software\Microsoft\Office\16.0\Word\Resiliency\DisabledItems` 里（按
> **路径**记）。之后再打开同一个路径会直接卡死，哪怕文件已经修好 —— 这次排查
> 里两份 md5 完全相同的文件（`3ecf13a060f267f61386f00c994ac731`），一个每次都
> 打开、另一个每次转圈，`cp` 到新路径后又立刻打开。所有 `tools/check_*.ps1`
> 现在都会先删掉这个键。

#### 读侧 API

```rust
let doc = office_doc::open("samples/ms/both3.doc")?;
assert_eq!(doc.footnote_slots, vec![18, 106]);      // 正文里的标记位置
assert_eq!(doc.endnote_slots, vec![37, 106]);       // 同上，尾注那张 PLC
assert_eq!(doc.endnote_text_slots, vec![0, 15, 18]); // 尾注故事内的切点
assert_eq!(doc.fib.ccp_counts[5], 16);              // ccpEdn
assert!(doc.story(office_doc::Story::Endnote).contains("endnote text"));
```

#### 真 Word 复核（`tools/check_endnote.ps1`）

```text
app path: C:\Program Files\Microsoft Office\Root\Office16
paras=3 footnotes=1 endnotes=1
  para: [WRITER-DOC-003 has a footnote.]
  para: [WRITER-DOC-004 has an endnote.]
  footnote 1: [footnote text] ref=[]
  endnote 1: [endnote text] ref=[]
```

`ref=[]` 是 COM 的怪癖（Word 自己写的 `samples/ms/both3.doc` 也是 `ref=[]`），
不是写入器的问题。

### 14.10 写页面设置

纸张与页边距不是段落属性，而是**节属性**：它们住在 `SEPX` 里，由 `Plcfsed`
的 `fcSepx` 指向。`SEPX` 本身是一个 `cb`（2 字节）+ 一串 sprm。

```rust
let page = office_doc::PageSetup::letter()
    .landscape()                 // 交换宽高，Word 靠形状判断横向
    .margins(1440, 720, 1080, 900); // 左 右 上 下，单位 twip（1/20 pt）
let documents = office_doc::write::Documents::default().page_setup(page);
let bytes = office_doc::write::write_document_with_footnotes(&body, &documents);
```

#### 五个 sprm（`examples/sepxdump.rs` 实测）

拿 Word 2016 造四份样本（`tools/tmp_page.ps1` → `samples/ms/page-{a4,letter,landscape,margins}.doc`），
用 `sepxdump` 读 `fcSepx` 指向的字节：

| sprm | 含义 | A4 样本 | margins 样本 |
|---|---|---|---|
| `0xB01F` | `sprmSXaPage` 纸宽 | 11906 | 11906 |
| `0xB020` | `sprmSYaPage` 纸高 | 16838 | 16838 |
| `0xB021` | `sprmSDxaLeft` 左边距 | 1418 | 2400 |
| `0xB022` | `sprmSDxaRight` 右边距 | 1418 | 800 |
| `0x9023` | `sprmSDyaTop` 上边距 | 1418 | 2000 |
| `0x9024` | `sprmSDyaBottom` 下边距 | 1418 | 1200 |

（70.9 pt = 1418 twip，90 pt = 1800，120 pt = 2400 —— 全部是 `pt × 20`。）
Word 自己还会写 `sprmSFpc`（0xB017）、`sprmSBkc`、`sprmSDxaGutter` 和一串打印机
默认值；这些都不影响 `PageSetup` 报出来的数字，写入器不写。

#### 两个坑

1. **`fcSepx` 指向 `WordDocument` 流，不是表流**。我第一次按表流去读，`cb` 读出
   26126，把 `SEPX` 解析成了垃圾。`Plcfsed` 的 CP 数组在表流里，`SEPX` 本体在
   正文流末尾。
2. **`SED` 是 12 字节，`fcSepx` 在第 2..6 字节**：`fn`(2) + `fcSepx`(4) +
   `fnMpr`(2) + `fcMpr`(4)。按「`fcSepx` 在最前面」去读会拿到 `fn` 的 2 字节，
   于是整个偏移错位。

#### 读侧

```rust
let doc = office_doc::open("samples/ms/page-landscape.doc")?;
let page = doc.page.unwrap();
assert_eq!((page.width, page.height), (15840, 12240));
assert_eq!((page.left, page.right), (1440, 1440));
assert_eq!(page.top, Some(1440));
```

`Doc::page` 取第一节的 `SEPX`；没有节属性时为 `None`。`examples/pages.rs` 把
它打印成 pt：

```text
samples/ms/page-margins.doc: 11906x16838 twips (595.3x841.9 pt), margins L2400 R800 T2000 B1200
samples/written-page.doc:    12240x15840 twips (612.0x792.0 pt), margins L1440 R1440 T1440 B1440
```

#### 真 Word 复核（`tools/check_page.ps1`）

```text
app: C:\Program Files\Microsoft Office\Root\Office16
page: 612 x 792 pt
margins: L72 R72 T72 B72 pt
orientation: 0 (0 portrait, 1 landscape)
```

写入器默认仍是 A4 纸 + 左右 1800 twip（1.25"）、上下交给 Word 的默认值（1440），
所以不设置页面设置时文件的字节与以前完全一致 —— 单测
`write::tests::page_setup_lands_in_the_section_properties` 钉住这一点。

### 14.11 首页与偶数页的不同页眉

`PlcfHdd` 的每节六个槽位里，除了「每页都用」的那一对，还有两对：第 5、6 个槽
是首页页眉/页脚，第 1、3 个是偶数页页眉/页脚。**但只把文字写进槽里 Word 不
显示** —— 两对各有自己的开关，而且开关不在同一个地方。

```rust
let documents = office_doc::write::Documents::new(&header, &footer)
    .first_page(&first_header, &[])   // 首页
    .even_pages(&even_header, &[]);   // 偶数页
let bytes = office_doc::write::write_document_and(&body, &documents);
```

#### 14.11.1 首页开关是节属性

`sprmSFTitlePage = 0x300A`，一个字节的布尔值，放在 `SEPX` 里。用
`examples/sepxdump.rs` 对比 Word 2016 造的两份样本
（`tools/tmp_diffhdr.ps1` → `samples/ms/hdr-{firstpage,evenodd}.doc`）：勾了
「首页不同」的那份多出一条 `sprm 0x300a len=1 value=1`，另一份没有。

#### 14.11.2 奇偶页开关在 `DOP` 里

勾了「奇偶页不同」的样本，`SEPX` 里**一条多余的 sprm 都没有** —— 这个开关是
文档级的 `fFacingPages`，`DOP` 第一个字节的 bit 0。用 `examples/dopdump.rs`
读 `fcDop/lcbDop`（FIB 第 31 对，`0x9A + 31 * 8`）就能看到：

```text
hdr-firstpage.doc  DOP[0] = 0x20
hdr-evenodd.doc    DOP[0] = 0x21   ← fFacingPages
```

所以写入器只在真的需要偶数页页眉时才写 `DOP`（84 字节的 `DopBase`，除 bit 0
之外全零），其余文档的第 31 对保持 0/0 —— 这样没有首页/偶数页页眉的文件字节
不会因为写入器学会了这个功能而变化。

#### 14.11.3 读回来

```rust
let doc = office_doc::open("samples/written-variants.doc")?;
assert_eq!(doc.header_footer(0), Some(("Every-page header".into(), "Every-page footer".into())));
assert_eq!(doc.first_page_header_footer(0), Some(("First-page header".into(), String::new())));
assert_eq!(doc.even_page_header_footer(0), Some(("Even-page header".into(), String::new())));
```

三个方法都是读 `PlcfHdd` 的槽位再裁掉段落符，所以「开关关着但槽里有字」的文件
也会把字读出来 —— 槽位是数据，开关才是显示条件。

#### 14.11.4 验证

`tools/check_variants.ps1`（Word COM，先删注册表键拿真 Word）读
`samples/written-variants.doc`：

```text
app: C:\Program Files\Microsoft Office\Root\Office16
first page header: -1
odd and even:      -1
primary header: [Every-page header]
first header:   [First-page header]
even header:    [Even-page header]
```

单测 `write::tests::first_and_even_headers_land_in_their_slots` 同时盯三件事：
三个槽位的文字、`WordDocument` 流里出现 `sprmSFTitlePage` 的三个字节
（`0A 30 01`）、以及 `DOP` 的 `fFacingPages` 置位。

### 14.12 写域（页码 / 总页数）

域在文本流里就是三个控制字符加两段普通文字：

```text
0x13  " PAGE "  0x14  "1"  0x15      域开始、域代码、分隔符、结果、域结束
```

两句都要：

1. 字符原样放进文本流（`\u{13}` / `\u{14}` / `\u{15}`）；
2. **这三个字符所在的那一跳 CHPX 必须带 `sprmCFSpec`（0x0855）** ——
   没有它 Word 把整段当成普通文字显示（`Fields.Count == 0`，页眉直接印出
   「PAGE 1」字样）。第一版只做了第 1 条，Word 就是这样处理的。

写入器的两个细节：`lay_out` 会把一个段落的 run 整条复制进 CHPX，所以落笔前
`split_field_characters` 先把每个控制字符切成独立的 run 再打 `spec = true`；
控制字符的判定放在 `push_story`（页眉页脚的空槽分隔符）与这个切分里各一份。

```rust
let page = office_doc::write::page_field("1");
let total = office_doc::write::numpages_field("3");
let header = [Paragraph::text(&format!("Page {page} of {total}"))];
let doc = write_document_with(&body, &HeaderFooter { header: &header, footer: &[] });
```

结果里的 `1` / `3` 只是缓存值，Word 打开后按 `Fields.Update` 或打印时重算
（实测：`samples/written-fields.doc` 存的是 `3`，Word 更新后页眉变成
`Page 1 of 1`）。

真 Word 2016 复核（`tools/check_fields.ps1`）：

```text
app: C:\Program Files\Microsoft Office\Root\Office16
header text: [Page 1 of 1]
fields: 2
  field 1: code=[ PAGE 1] result=[1]
  field 2: code=[ NUMPAGES 3] result=[1]
```

单测 `write::tests::field_characters_are_marked_special`（生成的文档里能搜到
`sprmCFSpec` 的字节 `55 08 01`）。

### 14.13 写多节文档

`write_document_sections(&[SectionSpec, ..])` 把文档切成若干节，每节有自己
的纸张与方向。`SectionSpec::new(paragraphs).page(PageSetup)`；页眉、页脚、
脚注、尾注仍然属于整个文档（`write_document_sections_all` 接受 header 槽与两条
注故事），不按节分。

Word 的存储方式（用 `samples/ms/sections2.doc` 实测：A4 纵向一节 + A4 横向一节，
由 Word 2016 自己写出）：

```text
ccpText = 40, 正文 = "Portrait section.\r\u{0c}Landscape section.\r\r\r"
Plcfsed CPs = [0, 19, 40]     节边界落在节标记 0x0C 之后
SEPX #0: 0xb01f 11906  0xb020 16838  0xb021 1800  0xb022 1800  (纵向 A4)
SEPX #1: 0x301d 2  0xb01f 16838  0xb020 11906  0xb021 1440  0xb022 1440
```

四条要点：

1. **节标记 `0x0C`（`\u{0c}`）是正文的一部分**，位于每节最后一个段落符之后、
   下一节第一个字符之前；它算进前一个节，所以 `Plcfsed` 的边界 CP 是「`0x0C`
   之后的下一个位置」。
2. `PlcfSed` 的 CP 数组有 `节数 + 1` 项；最后一项等于 `sed_end`（有脚注/页眉/
   尾注故事时是 `ccp*` 之和 + 1，否则是 `ccpText`），与 `PlcffndRef`、`PlcfHdd`
   的收尾 CP 必须一致。
3. **`sprmSBOrientation`（0x301D）是一字节，值 `2` = 横向**。只把宽高互换
   （`0xb01f` 写 16838、`0xb020` 写 11906）Word 的页面确实是横的，但
   `PageSetup.Orientation` 仍读回 0（纵向）—— 两者都要写。
4. 每个 `SEPX` 在 `WordDocument` 流里紧跟格式页，各自有独立偏移；
   `PlcfSed` 的每条 12 字节 SED 里，`fcSepx` 在第 2..6 字节。一节若没改过
   页面设置，Word 会写 `fcSepx = 0xFFFF0000` 表示「继承上一节」，读侧
   `Doc::sections` 会把上一节的设置重复一遍。

读侧：`Doc::sections: Vec<PageSetup>`（按文档顺序，继承的节重复上一节）；
`Doc::page` 仍等于第一节的设置。验证脚本 `tools/check_sections.ps1`：

```text
sections: 2
  1: 595.3x841.9 pt orient=0 text=[Portrait section.\n<sect>]
  2: 841.9x595.3 pt orient=1 text=[Landscape section.\n]
```

单测 `write::tests::sections_get_their_own_page_setups`。

### 14.14 写超链接

Word 的超链接不是一条独立记录，而是一个**域**：指令 `HYPERLINK "url"` 夹在
`0x13` 与 `0x14` 之间，显示的文本夹在 `0x14` 与 `0x15` 之间（见 §14.12 的域机制）。
所以写超链接不需要新的 PLC 或存储，只要把域字符标成 `sprmCFSpec`，Word 打开时自己
把域变成可点的链接：

```rust
use office_doc::write::{hyperlink, write_document, Paragraph};
let body = vec![
    Paragraph::text("Before the link."),
    Paragraph::text(format!("Read {} for details.", hyperlink("https://example.com/guide", "the guide"))),
];
let bytes = write_document(&body);
```

真 Word 复核（`tools/tmp_hlinkdoc.ps1` 用 COM 打开并打印 `Hyperlinks`）：

```text
hyperlinks=1
  text=[] addr=[https://example.com/guide] sub=[] type=0
```

Word 自己写的样本（`samples/ms/hyperlink.doc`，用 `Hyperlinks.Add` 造）里指令是
`HYPERLINK "https://example.com/guide" \o "the site"` —— 多一个 `\o` 开关带
屏幕提示；我们只写最简单的 `HYPERLINK "url"`，Word 同样认。显示文本 Word 会在末尾
补一个空格再接 `0x15`，读侧 `Doc::hyperlinks()` 会 `trim_end()` 掉。

读侧：`Doc::hyperlinks() -> Vec<(String, String)>`（`(url, 显示文本)`），按域字符切
正文再解析 `HYPERLINK` 指令。单测 `write::tests::a_hyperlink_is_a_field_with_a_url`、
`tests::reads_hyperlinks_as_fields`、`tests::reads_a_word_hyperlink_field_from_the_sample`。
注意与 `.xls` 的区别：那边一个链接是一条 `HLINK` 记录（xls.md §21），这里是域。

### 14.15 写列表（项目符号与编号）

列表分两层存：**定义**（每个图标长什么样）和**归属**（哪个段落属于哪个列表的哪一级）。
归属放在段落的 PAPX 里，只两条 sprm：`sprmPIlfo`（0x460B，一基的列表号，0 表示不在
列表里）与 `sprmPIlvl`（0x260A，层级）。定义放在两个 FIB 对里：

| FIB 对 | 名字 | 内容 |
|---|---|---|
| 73 | `fcPlfLst`/`lcbPlfLst` | `PlfLst`：`cLst`（u16）+ `cLst` 个 28 字节 `LSTF`（`lsid`、`tplc`、`rgistdPara[9]`、位域字节 `fSimpleList` bit0、`grfhic`）；**`LVL` 数组紧跟在 `fc+lcb` 处**（`lcb` 只算 `PlfLst` 本体，不含 `LVL`） |
| 74 | `fcPlfLfo`/`lcbPlfLfo` | `PlfLfo`：`lfoMac`（u32）+ `lfoMac` 个 16 字节 `LFO`（`lsid`、`unused1/2`、`clfolvl`、`ibstFltAutoNum`、`grfhic`、`unused3`）+ `lfoMac` 个 `LFOData`（`cp`，无覆盖时 `0xFFFFFFFF`） |

`LVL` = `LVLF`（28 字节）+ `grpprlPapx` + `grpprlChpx` + `Xst`（`cch` u16 + UTF-16）。
`LVLF` 的关键字段：`iStartAt`（i32）、`nfc`（u8，**0x17 = 项目符号，0x00 = 阿拉伯数字**）、
`rgbxchNums[9]`（模板里哪些字符是占位符的**一基**下标，零结尾）、`cbGrpprlChpx`、
`cbGrpprlPapx`。模板 `Xst` 的占位符在文件里是不可打印码点（`U+0000`/`U+0001`…），
读侧换成 `1` 显示。

两个坑，都是从 Word 2016 自写的 `samples/ms/lists.doc` 逐字节比出来的：

1. **项目符号的 `rgbxchNums` 必须是全 0**。项目符号那一级 `nfc = 0x17`、`Xst` 是**单个
   字面字符**（Symbol 字体的圆点是 `U+F06C`），不是占位符。若照数字模板把
   `rgbxchNums[0] = 1`，Word 会把那个字符当占位符换成序号，圆点变「1」。
2. **`lcbPlfLst` 不含 `LVL` 数组**。Word 写的是 `cLst + 2×28 = 58`，`LVL` 从
   `fc + 58` 才起。读侧若按 `lcb` 去找 `LVL` 就会落空。写侧照此填 `lcb`。

写入 API：

```rust
use office_doc::write::{write_document_and, Documents, ListSpec, Paragraph};
use office_doc::fmt::ParaProps;
let lists = [ListSpec::bullets('•'), ListSpec::numbers("1.")];
let item = |text: &str, ilfo: u16| Paragraph {
    props: ParaProps { ilfo: Some(ilfo), ilvl: Some(0), ..ParaProps::default() },
    ..Paragraph::text(text)
};
let body = vec![Paragraph::text("Bullets:"), item("first", 1), item("second", 1),
                Paragraph::text("Numbers:"), item("one", 2), item("two", 2)];
let bytes = write_document_and(&body, &Documents::new(&[], &[]).lists(&lists));
```

`ListSpec::bullets(c)` 造单级项目符号列表，`ListSpec::numbers("1.")` 造单级编号列表
（模板里一个 `1` 是占位，`"1)"`、`"(1)"` 都行），`ListSpec::levels(vec![...])` 造多级。
`ilfo` 是一基的列表号（对应 `&lists` 的顺序），`ilvl` 从 0 起。没有列表时两个 FIB 对
保持 `0/0`，文件字节不变。

读侧：`Doc::lists -> Vec<ListDef>`（每个 `ListDef { lsid, levels: Vec<ListFormat> }`）、
`Doc::list_overrides -> Vec<i32>`（一基的 `ilfo` → `lsid`），`paragraphs_rich()` 里每段
的 `fmt::Paragraph::list: Option<ListFormat>` 已解析好（`kind` 是 `Bullet`/`Number`/`None`，
`level`、`text` 模板、`start` 起始号）。真 Word 复核（`tools/check_lists.ps1`）：

```text
paras=6 listtemplates=2
  para 2: type=2 level=1 listString=[]   text=[WRITER-DOC-101 first bullet]
  para 5: type=3 level=1 listString=[1.] text=[WRITER-DOC-103 first number]
  para 6: type=3 level=1 listString=[2.] text=[WRITER-DOC-104 second number]
```

单测 `write::tests::lists_carry_their_markers_and_levels`、
`write::tests::a_document_without_lists_keeps_the_pairs_at_zero`。另外修了一个读侧
崩溃：列表段落的 PAPX 里 `sprmPIlfo` 只带 **1 个字节**操作数（Word 截断），旧的
`apply_para_sprms` 按 2 字节读会 `index out of bounds`（`fmt.rs`），现在按「拿得到几个
字节算几个」处理。

### 14.16 写批注与读回来

Word 的批注（comment）不是独立对象，而是**锚在正文里的一个引用字符 + 一段批注故事**。
批注的引用字符是 `0x05`：正文里它标出批注位置，批注故事里每条也以 `0x05` 开头。
脚注/尾注用的是 `0x02`，所以批注和脚注在正文里长得不一样，Word 靠这个区分。

写入 API：

```rust
use office_doc::write::{write_document_and, CommentSpec, Documents, Paragraph, COMMENT_MARK};
let body = vec![
    Paragraph::text(format!("First para with a note.{COMMENT_MARK}")),
    Paragraph::text(format!("Second para also noted.{COMMENT_MARK}")),
];
let comments = [
    CommentSpec { author: "Jamie Yates", initials: "jy", text: "please cite this" },
    CommentSpec { author: "Alex Lee", initials: "al", text: "second remark here" },
];
let bytes = write_document_and(&body, &Documents::new(&[], &[]).comments(&comments));
```

`COMMENT_MARK`（`0x06`）是**源文本里的占位**，排版前会被改写成 `0x05`；正文里有几个
`COMMENT_MARK`，就得给几个 `CommentSpec`（顺序对应）。批注正文里用 `\r` 分段。

涉及四样东西，缺一不可：

| FIB 对 | 名字 | 内容 |
|---|---|---|
| 4 | `fcPlcfandRef`/`lcbPlcfandRef` | 正文里每个批注锚点的 CP 数组 + 收尾 CP（= `sed_end`）+ 每条批注 30 字节的 `ATRDPre10` |
| 5 | `fcPlcfandTxt`/`lcbPlcfandTxt` | 批注故事里每条批注的起始 CP + 收尾 CP |
| 36 | `fcGrpXstAtnOwners`/`lcbGrpXstAtnOwners` | 作者名表：一列 `XST`（`cch` u16 + UTF-16），每个作者一条 |
| — | `ccpAtn`（FIB 偏移 0x5C） | 整个批注故事的长度 |

`ATRDPre10`（30 字节，MS-DOC §2.9.1）：`xstUsrInitl` 是 20 字节的 `LPXCharBuffer9`
——**前 2 字节是 `cch`**，随后最多 9 个 UTF-16 首字母；接着 2 字节 `ibst`（作者名在
`GrpXstAtnOwners` 里的 0 基索引）、2 字节 `bitsNotUsed`（0）、2 字节 `grfNotused`（0）、
4 字节 `lTagBkmk`（零长度批注范围写 `-1`）。

四个硬性要求，每一个都单独让 Word 拒收过：

1. **批注锚点必须带 `sprmCFSpec`**。`PlcfandRef` 里记录的每个 CP 都必须是「特殊字符」
   （`sprmCFSpec = 1` 的 run）。写入器把 `0x05`/`0x02` 都当作特殊字符切出独立 run
   （`split_special_characters`），否则 Word 读不到锚点。
2. **`GrpXstAtnOwners` 必须存在**。MS-DOC 原文：文档只要有批注，`fcGrpXstAtnOwners`
   就必须指向一张合法的 `XST` 数组，且每条 `ATRDPre10.ibst` 是它的合法索引。缺了它
   Word 报「文件内容有问题」。
3. **故事顺序是 正文 → 脚注 → 页眉 → 批注 → 尾注**。读侧是按 `ccpText`/`ccpFtn`/
   `ccpHdd`/`ccpMcr`/`ccpAtn`/`ccpEdn` 累加来定位故事的，批注故事必须排在页眉之后、
   尾注之前，`ccpAtn` 才算得对。
4. **CLX 的 piece table 要为批注故事单切一条 piece**。`build_from_stream` 里每多一个
   故事就多一条 piece（正文、脚注、页眉、批注、尾注各一条），否则 Word 用错误的
   fc 去读文本，正文全变问号。

`sed_end` 也要把 `ccpAtn` 算进去：有任何附属故事时 `sed_end = ccpText + ccpFtn +
ccpHdd + ccpAtn + ccpEdn + 1`。

读侧：`Doc::comments -> Vec<DocComment>`，每个 `DocComment { anchor: u32, initials:
String, text: String }`，`anchor` 是正文里 `0x05` 所在的 CP，`initials` 来自 `ATRDPre10`，
`text` 是该条批注故事的文本（去掉开头的 `0x05` 与结尾的段落符）。

真 Word 2016 复核（`tools/check_comments.ps1 samples/written-comments.doc`）：

```text
app path: C:\Program Files\Microsoft Office\Root\Office16
paras=2 comments=2
  para: [WRITER-DOC-201 first para with a note.]
  para: [WRITER-DOC-202 second para also noted.]
  comment 1: author=Jamie Yates initials=jy scope=[] text=[please cite this]
  comment 2: author=Alex Lee initials=al scope=[] text=[second remark here]
```

`scope=[]` 是 COM 的怪癖：我们的批注锚点是零长度范围，`Comment.Scope.Text` 取回来是
空串（Word 自写的同形文件也一样）。批注文字与作者都正确读回。

单测 `write::tests::comments_land_in_the_annotation_story`、
`lib::tests::reads_comments_from_the_sample`（读 `samples/ms/comments.doc` 得到两条批注、
initials 为 `jy`、正文 `0x05` 位置与两个 anchor 一致）。

#### 14.16.1 页眉故事的收尾（批注/尾注共用同一规则）

批注与尾注都排在页眉故事之后，所以「页眉故事写几个收尾段落符」由「后面是否还有故事」
决定，而不是由故事种类决定。Word 自己的两份样本给出了对照：

| 文件 | 页眉故事长度 | `ccpHdd` | 最后一个 `PlcfHdd` CP | 页眉之后 |
|---|---|---|---|---|
| `samples/ms/header.doc` | 51 | 49 | 51 | 无（页眉是最后的故事） |
| `samples/ms/header-note.doc` | 49 | 49 | 51 | 尾注 |

两次都是 `ccpHdd = 49`、末 CP = 51，只有物理收尾段落符的个数不同：页眉是最后一个故事
时写 3 个 `\r`，后面还有故事时写 1 个（因为下一个故事必须正好从 `ccp*` 之和、也就是
`ccpHdd` 开始）。写入器据此判断 `story_follows = !comments.is_empty() ||
!endnotes.is_empty()`。

### 14.17 文档属性（SummaryInformation）

Word 的每个文件里都有两个属性集流，`文件 > 属性` 对话框显示的就是它们：

```text
\x05SummaryInformation          标题/主题/作者/关键字/备注/模板/修订号/应用程序/时间/页数/字数
\x05DocumentSummaryInformation  类别/经理/公司/行数/段落数/幻灯片数
```

我们的写入器原来从不生成它们（`samples/written*.doc` 的复合文件只有
`Root Entry` / `1Table` / `WordDocument`），读侧也没有解析。现在由
`office-oleps` 这个 crate 统一处理，`.doc` / `.xls` / `.ppt` 共用。

#### 14.17.1 字节布局（MS-OLEPS）

流头 28 字节：

| 偏移 | 长度 | 内容 |
|---|---|---|
| 0 | 2 | 字节序，`0xFFFE` |
| 2 | 2 | 版本，0 |
| 4 | 4 | 系统标识，Word 与 Excel 都写 `0x0002000A` |
| 8 | 16 | CLSID，全 `0` |
| 24 | 4 | 节数 |

随后每个节一项 20 字节：**16 字节 `FMTID` 在前，4 字节节偏移在后**（偏移相对流起点）。
（按「偏移在前」读会把 `FMTID` 的后 4 字节当成偏移，得到一个 40 亿的地址——实测踩过。）

节体：`cbSection u32` + `cProperties u32` + `cProperties × (pid u32, offset u32)`，
offset 相对**节起点**，值 = `VT_` 标签（4 字节）+ 数据。属性按 `pid` 升序，
第一个永远是 `pid 1 = codepage`。

用到的变体：

| 标签 | 名称 | 数据 |
|---|---|---|
| `0x0002` | `VT_I2` | 2 字节，补到 4 |
| `0x0003` | `VT_I4` | 4 字节 |
| `0x001E` | `VT_LPSTR` | `cch u32`（含结尾 NUL）+ 单字节字符 |
| `0x001F` | `VT_LPWSTR` | `cch u32`（含结尾 NUL）+ UTF-16 字符 |
| `0x0040` | `VT_FILETIME` | 8 字节，1601-01-01 起的 100 纳秒数 |

`FMTID_SummaryInformation` = `E0859FF2-F94F-6810-AB91-08002B27B3D9`，
`FMTID_DocSummaryInformation` = `02D5CDD5-9C2E-1B10-9397-08002B2CF9AE`。

#### 14.17.2 两个坑

**① 结尾 NUL 不止一个。** `VT_LPSTR` 的 `cch` 已经含结尾 NUL，Word 还会多补几个：
`samples/ms/header.doc` 的空标题读出来是 `"\0\0\0"`，作者是 `"jd yuan"`、模板是
`"Normal\0"`。按 `cch` 读、只砍一个 NUL 会得到带 `\0` 的字符串（终端上看着像乱码）。
正确做法是把**所有**尾部零字节都修剪掉（窄字符按字节、宽字符按码元）。

**② 编码有两种。** Word 写 `VT_LPSTR` + codepage 936（或 1252），WPS 写
`VT_LPWSTR` + codepage 1200。两者都合法，读侧都要认；写侧统一用 `VT_LPWSTR` +
codepage 1200，这样中文与任何非拉丁文字都不需要码页表。

#### 14.17.3 API

```rust
use office_doc::write::{write_document_and, Documents, Paragraph};

let summary = office_oleps::Summary {
    title: "Quarterly report".into(),
    author: "Jamie Yates".into(),
    template: "Normal".into(),
    revision: "2".into(),
    application: "some-office".into(),
    pages: 1,
    words: 5,
    ..office_oleps::Summary::new()
};
let doc_summary = office_oleps::DocSummary {
    company: "Example Ltd".into(),
    manager: "Robin Fox".into(),
    paragraphs: 1,
    ..office_oleps::DocSummary::default()
};
let documents = Documents::new(&[], &[])
    .summary(summary)
    .doc_summary(doc_summary);
let bytes = write_document_and(&[Paragraph::text("Has metadata.")], &documents);
```

读回来：

```rust
let doc = office_doc::open("samples/written-metadata.doc")?;
let summary = doc.summary.expect("summary stream");
assert_eq!(summary.title, "Quarterly report");
assert_eq!(doc.doc_summary.unwrap().company, "Example Ltd");
```

`Summary` 的字段与属性 id 一一对应（2 标题、3 主题、4 作者、5 关键字、6 备注、
7 模板、8 最后作者、9 修订号、12/13 时间、14 页数、15 字数、16 字符数、18 应用程序、
19 安全性）；空字符串照样写进去，但**没有 `Summary` 就完全不写这个流**，所以
`samples/written.doc` 的字节一个都不变。

#### 14.17.4 验证

两个流是纯附加的：文本、FIB、表流一律不动，因此判断「Word 是否读懂了」最直接的办法
是**让 Word 自己另存一份**，再看元数据还在不在（Word 会把自己不认识的属性丢掉、
并覆盖它自己维护的字段）：

```text
$ powershell -File tools/resave_doc.ps1      # 打开 written-metadata.doc，另存为 tmp-resaved.doc
app path: C:\Program Files\Microsoft Office\Root\Office16
opened paras=1

$ cargo run -q -p officedump --example meta -- samples/written-metadata.doc
    title="Quarterly report" subject="Budget"
    author="Jamie Yates" last_author="Alex Lee"
    keywords="budget, q3" comments="Written for the board."
    template="Normal" revision="2"
    application="some-office"
    pages=1 words=5 chars=26 security=0
  doc summary: manager="Robin Fox" company="Example Ltd" paragraphs=1 lines=1

$ cargo run -q -p officedump --example meta -- samples/tmp-resaved.doc
    title="Quarterly report" subject="Budget"
    author="Jamie Yates" last_author="jd yuan"
    keywords="budget, q3" comments="Written for the board."
    template="Normal" revision="2"
    application="Microsoft Office Word"
    pages=1 words=4 chars=25 security=4
  doc summary: manager="Robin Fox" company="Example Ltd" paragraphs=1 lines=1
```

我们写的标题、主题、作者、关键字、备注、模板、修订号、公司、经理全部被 Word 原样保留；
被改动的只有 Word 自己维护的字段（最后作者、应用程序、时间、字数、安全性）。

`tools/check_metadata.ps1` 想直接读 `Document.BuiltInDocumentProperties`，但这台机器上
该 COM 对象取 `.Count` 就抛「未将对象引用设置到对象的实例」（Word 自己写的文件也一样），
所以验证走「另存 + 读回」这条路。单测是
`write::tests::metadata_lands_in_its_own_streams`、
`write::tests::a_document_without_metadata_has_no_property_streams`、
`lib::tests::reads_the_metadata_word_writes`（后者同时覆盖 Word 的 `VT_LPSTR` 与
WPS 的 `VT_LPWSTR` 两种写法）。

### 14.18 写书签与读回来

书签是「有名字的字符区间」。它和脚注、批注都不一样：**书签不往正文里放任何字符**，
所以一个带书签的文件和同样的文件相比，只有表流里的三张表不同，文本、`ccpText`、
`Plcfsed` 全部原样。

Word 把书签拆成三份存放，三个 FIB 对连在一起：

| FIB 对 | 名字 | 内容 |
|---|---|---|
| 21 | `fcSttbfBkmk` | 书签名表（`STTB`，`fExtend = 0xFFFF` 即名字是 UTF-16） |
| 22 | `fcPlcfBkf` | 每个书签的**起点** CP，按 `ibkl` 索引 |
| 23 | `fcPlcfBkl` | 每个书签的**终点** CP，同样的 `ibkl` 索引 |

`samples/ms/bookmarks.doc`（Word 2016 自造）实测：

```text
SttbfBkmk  fc=7534 lcb=38
  ff ff   fExtend = 0xFFFF（名字是 UTF-16）
  03 00   cData   = 3
  00 00   cbExtra = 0
  05 00 61 00 6c 00 70 00 68 00 61 00   "alpha"
  04 00 62 00 65 00 74 00 61 00         "beta"
  04 00 74 00 61 00 69 00 6c 00         "tail"

PlcfBkf   fc=7572 lcb=28   CPs [6, 19, 31, 51] + 3 × (ibkl u16, bits u16)
PlcfBkl   fc=7600 lcb=16   CPs [12, 23, 36, 51]（没有数据）
```

几个要点：

1. **两个 PLC 的收尾 CP 都是 `ccpText + 2`**（上例 `49 + 2 = 51`）。它比正文末
   还多一个字符，读者不用它，但 Word 自己就是这么写的，照抄最稳。
2. `PlcfBkl` 的 `lcb` 正好是 `4 × (书签数 + 1)`：它只有 CP，没有条目数据。
   按 `(lcb - 4) / 4` 去数条目会得到两倍的个数（实测越界 panic）。
3. `BKF` 的 4 个字节是 `ibkl u16`（书签序号）+ `bits u16`（`fCol` 表示列书签，
   我们写 0）。它是**独立编号**，与 `SttbfBkmk` 的顺序一致。
4. `SttbfBkmk` 的 `fExtend` 是 `0xFFFF`（不是 `fExtend + 字符宽度` 那套），
   所以每项 = `cch u16` + UTF-16 字符。

写：

```rust
use office_doc::write::{write_document_and, BookmarkSpec, Documents, Paragraph};

let body = vec![
    Paragraph::text("Alpha target line."),
    Paragraph::text("Beta target line."),
    Paragraph::text("Tail line."),
];
// (段落下标, 段内字符下标)：区间左闭右开
let bookmarks = [
    BookmarkSpec { name: "alpha", start: (0, 6), end: (0, 12) },
    BookmarkSpec { name: "beta",  start: (1, 0), end: (1, 4) },
    BookmarkSpec { name: "tail",  start: (2, 0), end: (2, 5) },
];
let documents = Documents::new(&[], &[]).bookmarks(&bookmarks);
let bytes = write_document_and(&body, &documents);
```

位置用「段落下标 + 段内字符下标」而不是 CP，是因为调用方手里只有段落；
写入器在排版之后从每个段落记下的 `fc` 反推出 CP（`(fc - TEXT_AT) / 2 + 列`）。
`start == end` 会写出一个零长度书签（Word 里显示为一根竖线）。

读回来：

```rust
let doc = office_doc::open("samples/written-bookmarks.doc")?;
for b in &doc.bookmarks {
    let text: String = doc.text.chars().skip(b.start as usize)
        .take((b.end - b.start) as usize).collect();
    println!("{} = {:?}", b.name, text);
}
// alpha = "target"
// beta  = "Beta"
// tail  = "Tail "
```

真 Word 2016 复核（`tools/check_bookmarks.ps1`，与其它脚本一样先删掉用户的 COM 覆盖键
才拿到微软 Word，并打印 `$app.Path`）：

```text
$ powershell -File tools/check_bookmarks.ps1 samples/written-bookmarks.doc
app path: C:\Program Files\Microsoft Office\Root\Office16
paras=3 bookmarks=3
  alpha = [target] start=6 end=12
  beta = [Beta] start=19 end=23
  tail = [Tail ] start=37 end=42
```

名字与区间都被 Word 认下。没有书签时三个 FIB 对保持 `0/0`，
`samples/written.doc` 的字节一个都不变。单测是
`write::tests::bookmarks_land_in_the_three_tables`、
`write::tests::a_document_without_bookmarks_keeps_the_pairs_at_zero`、
`lib::tests::reads_bookmarks_from_the_sample`。
