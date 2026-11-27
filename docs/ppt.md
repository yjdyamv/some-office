# .ppt：PowerPoint 97-2003 二进制演示文稿

`.ppt` 在 CFB 里是两个流：`PowerPoint Document`（记录树）和 `Current User`（一个 20 字节
的指针，告诉你"最后一次编辑"在哪）。本仓库实现：`crates/office-ppt/src/lib.rs`。

## 1. 记录头

每条记录 8 字节：

```
u16 ver_inst   低 4 位 = ver，高 12 位 = instance
u16 type
u32 len        负载长度（不含 8 字节头）
```

- `ver == 0xF` → 这是一个**容器**，负载是子记录序列；
- `ver == 0` / `1` / `2` → 原子（atom），负载是裸数据；
- `instance` 对原子常常没意义，对某些记录是语义（如 `SlideListWithText` 的 0/1/2）。

## 2. 顶层布局

`samples/hello.ppt` 的 `PowerPoint Document` 流（46999 字节）顶层只有 6 条记录：

| 偏移 | 类型 | len | 说明 |
|---|---|---|---|
| 0x0000 | 0x03E8 Document | 1443 | 文档级：原子 + 幻灯片清单 |
| 0x05AB | 0x03F8 MainMaster | 42926 | 母版，几乎全是 Escher 图形 |
| 0xAD61 | 0x03EE Slide | 1882 | 第 1 页 |
| 0xB4C3 | 0x03EE Slide | 652 | 第 2 页 |
| 0xB757 | 0x1772 PersistDirectoryAtom | 20 | persistId → 偏移 |
| 0xB773 | 0x0FF5 UserEditAtom | 28 | 最后一次编辑 |

```
Document
├── DocumentAtom        0x03E9  @0x0008 len=40
├── Environment         0x03F2  @0x0038 len=726
├── PPDrawingGroup      0x040B  @0x0316 len=116
├── SlideListWithText   0x0FF0  @0x0392 inst=1 len=28   <- 母版清单
├── ...                 0x07D0 / 0x0FD9 ...
├── SlideListWithText   0x0FF0  @0x0519 inst=0 len=56   <- 幻灯片清单
├── ...                 0x1388  @0x0559 len=66
└── EOF                 0x03EA  @0x05A3
```

`Slide` 容器：

```
Slide
├── SlideAtom           0x03EF  @+8  len=24
├── PPDrawing           0x040C  @+40 len=1786   <- 这一页的所有形状
├── ...                 0x07F0 / 0x0422
```

## 3. SlideListWithText：页序的唯一来源

`instance` 决定这份清单是什么：

| instance | 内容 |
|---|---|
| 0 | 幻灯片（SlidePersistAtom 0x03F3） |
| 1 | 母版 |
| 2 | 备注页 |

hello.ppt：`inst=1` 的清单里有 1 个 SlidePersistAtom（1 个母版），`inst=0` 的有 2 个
（2 页）。**每个 SlidePersistAtom 的 `persistIdRef`（负载前 4 字节）才是页序**，
不是 `Slide` 记录在流里的物理顺序。

## 4. PersistDirectoryAtom：persistId → 文件偏移

PowerPoint 用 persistId 而不是绝对偏移来引用对象，因为文件是**增量保存**的：
每次保存只在尾部追加新记录，旧记录还在。目录记录（0x1772）把 persistId 映射到偏移：

```
hello.ppt 的 PersistDirectoryAtom 负载（20 字节）:
01 00 40 00   u32 0x00400001
              低 20 位 = persistId = 1
              高 12 位 = 条目数 - 1 = 4  → 4 个条目
00 00 00 00   persistId 1 → 0x00000000  (Document)
ab 05 00 00   persistId 2 → 0x000005AB  (MainMaster)
61 ad 00 00   persistId 3 → 0x0000AD61  (Slide 1)
c3 b4 00 00   persistId 4 → 0x0000B4C3  (Slide 2)
```

一个 PersistDirectoryAtom 可以包含多个 `(persistId, count, offsets...)` 组，直到负载用完。

**坑**：规范说高 12 位是"条目数减 1"，**实测是条目数本身**：

| 文件 | 高 12 位 | 负载长度 | 实际条目 |
|---|---|---|---|
| samples/hello.ppt | 4 | 20 | 4 |
| samples/notes.ppt | 7 | 32 | 7 |
| samples/encrypted.ppt | 4 | 20 | 4 |

所以 `n = (w >> 20) & 0xFFF`，再夹取到 `(len - 4) / 4` 以防记录被截断（增量保存会留下
半条目录）。早期版本按"减 1 再加 1"实现，靠夹取才碰巧正确。

## 5. Current User 流 + UserEditAtom

`Current User` 流（hello.ppt 共 47 字节，PROTTPLN.PPT 38 字节）：

```
00 00 f6 0f    rh: recVer=0, recInstance=0, recType=0x0FF6
1d 00 00 00    recLen = 29
14 00 00 00    size = 20              <- 固定部分长度
5f c0 91 e3    headerToken = 0xE391C05F
73 b7 00 00    offsetToCurrentEdit = 0xB773   <- UserEditAtom 在文档流里的偏移
05 00          lenUserName = 5
f4 03          docFileVersion = 0x03F4
03             majorVersion = 3
00             minorVersion = 0
00 00          unused
"admin"        ansiUserName
08 00 00 00    relVersion / 对齐填充
"admin"(UTF-16) unicodeUserName
```

关键两个字段（MS-PPT 2.3.2 / 文件结构表实测一致）：

- **偏移 12**：`headerToken = 0xE391C05F`，固定魔数，可做完整性校验；
- **偏移 16**：`offsetToCurrentEdit`，即 UserEditAtom 在 `PowerPoint Document` 里的偏移
  （hello.ppt = 0xB773 = 46963，PROTTPLN.PPT = 0x0001061F = 67103）。

注意 RecordHeader 在这里是 8 字节（含 recLen），所以 `size` 字段从偏移 8 开始，
不能把前 8 字节当成普通的 4 字节头再往下读。

`UserEditAtom`（0x0FF5）负载：

| 偏移 | 大小 | 字段 | hello.ppt |
|---|---|---|---|
| 0 | 4 | lastSlideIdRef | 0x101 |
| 4 | 2 | version | 0x17DF |
| 6 | 1 | minorVersion | 0 |
| 7 | 1 | majorVersion | 3 |
| 8 | 4 | offsetLastEdit | 0（没有更早的编辑） |
| 12 | 4 | **offsetPersistDirectory** | 0xB757 |
| 16 | 4 | docPersistIdRef | 1 |
| 20 | 4 | persistIdSeed | 4 |
| 24 | 4 | lastView | 1 |

正确的解析顺序（本仓库的实现）：

```
1. 读 Current User 流，拿 offsetToCurrentEdit
2. 在 PowerPoint Document 里读该偏移的 0x0FF5 记录（读不到就找最后一条 0x0FF5）
3. 取 offsetPersistDirectory → 解析 PersistDirectoryAtom
4. 在 Document 容器里找 inst=0 的 SlideListWithText，取 SlidePersistAtom 序列
5. 每个 persistIdRef 查目录 → 偏移 → 读 Slide 容器 → 遍历 PPDrawing
```

**为什么不能直接从头扫**：增量保存后，流里可能存在多份同名记录，从头扫会读到旧版本
（比如删掉的页还在）。只有走 Current User → UserEditAtom → PersistDirectory 才能
拿到当前状态。

## 6. 文字在哪

幻灯片正文在 `PPDrawing`（0x040C）里的 Escher 容器中：

```
PPDrawing
└── 0xF002 DgContainer        （画布）
    └── 0xF003 SpgrContainer （形状组）
        └── 0xF004 SpContainer（一个形状）
            └── 0xF00D OfficeArtClientTextbox
                ├── 0x0F9F TextHeaderAtom  len=4  （文字类型）
                ├── 0x0FA0 TextCharsAtom   （UTF-16LE 正文）
                └── 0x0FA8 TextBytesAtom   （8-bit 正文，纯 ASCII 时用）
```

hello.ppt 第 1 页的实际记录：

```
0x0B0C5  TextHeaderAtom  len=4   textType = 6  (Tx_TYPE_CENTERTITLE)
0x0B0D1  TextCharsAtom   len=26  "HELLO-PPT-001"
0x0B40B  TextHeaderAtom  len=4   textType = 5  (Tx_TYPE_CENTERBODY)
0x0B417  TextCharsAtom   len=38  "Slide one body text"
```

### TextTypeEnum（MS-PPT 2.13.33）

| 值 | 名字 | 含义 |
|---|---|---|
| 0 | Tx_TYPE_TITLE | 标题占位符 |
| 1 | Tx_TYPE_BODY | 正文占位符 |
| 2 | Tx_TYPE_NOTES | 备注 |
| 3 | （未使用） | |
| 4 | Tx_TYPE_OTHER | 其它（形状里的文字） |
| 5 | Tx_TYPE_CENTERBODY | 居中正文（标题页的副标题） |
| 6 | Tx_TYPE_CENTERTITLE | 居中标题（标题页的主标题） |
| 7 | Tx_TYPE_HALFBODY | 双栏正文 |
| 8 | Tx_TYPE_QUARTERBODY | 四栏正文 |

注意 WPS 的标题页用的是 6/5（居中版式），不是 0/1。**只认 0/1 会把所有标题判成
"其它"**——这是本仓库第一版的 bug，实测后改成 `0|6 → Title`、`1|5|7|8 → Body`。

正文里的段落分隔是 `\r`（0x0D），与 Word 一致；软回车 0x0B。

## 6.1 备注（notes）：另一个 persist 对象

备注**不在**幻灯片里。每页幻灯片有一个独立的 `Notes` 持久对象（记录类型 `0x03F0`），
它在 `SlideListWithText` 的 **instance = 2** 列表里按幻灯片顺序排列：

```text
samples/notes.ppt 的 Document 子记录：
  SlideListWithText inst=1  len=28   1 × 0x03F3 (母版)
  SlideListWithText inst=2  len=56   2 × 0x03F3 (备注！与幻灯片同序)
  SlideListWithText inst=0  len=56   2 × 0x03F3 (幻灯片)

persist 目录（7 个对象）：
  1 → 0x00000     Document
  2 → 0x005B7     MainMaster
  3 → 0x0EBB8     Slide 1
  4 → 0x0EE5E     Notes 1     ← 记录类型 0x03F0
  5 → 0x0F0B4     Slide 2
  6 → 0x0F35A     Notes 2
  7 → 0x0AD6D     NotesMaster
```

备注里的文字和幻灯片正文一样，藏在 `Notes` 的 `PPDrawing` → `0xF004 SpContainer` →
`0xF00D OfficeArtClientTextbox` → `0x0FA0 TextCharsAtom` 里，所以直接复用同一个 walker：

```rust
let deck = office_ppt::parse_stream(&stream, Some(&current_user))?;
for slide in &deck.slides {
    println!("{}", slide.notes_plain_text());   // 演讲者备注
}
```

如果 `instance = 2` 的列表缺失（被裁剪过的文件），实现会退化成按 persist 偏移排序后
依次配对。

## 7. 母版为什么那么大

hello.ppt 的 `MainMaster` 有 42926 字节，占整个流的 91%。母版里存了 12 个占位符形状
的完整 Escher 图形（填充、线条、阴影、文本样式）。只取文字的话可以完全跳过它。

## 8. 图片：Pictures 流 + BLIP + pib

图片不在 `PowerPoint Document` 里，单独放在 **`Pictures` 流**；幻灯片里的形状只存一个
整数索引 `pib`（picture index），指过去取。

`Pictures` 流的结构（本仓库 `pictures()` 两种都吃）：

```text
OfficeArtBStoreContainer (0xF001)
└─ OfficeArtBSE (0xF007)   ← 一张图一条
     OfficeArtFBSE: btWin32(1) btMacOS(1) rgbUid(16) tag(2) size(4)
                    cRef(4) foDelay(4) unused1(1) cbName(1) unused2(1)
                    unused3(1) name(cbName*2)      → 共 36 + cbName*2 字节
     然后紧跟着 BLIP 记录（0xF018..0xF02A），BLIP 载荷 = rgbUid(16) + 图片字节
```

WPS 会跳过容器，直接把一条 BLIP 记录写在流开头——所以解析时两种都要试。

**坑：BLIP 的记录类型不可信。** `samples/image.ppt` 里那张 PNG，记录类型写的是
`0xF01E`（规范里是 `OfficeArtBlipTIFF`），但 `instance >> 8` 是 6（PNG），载荷也
确实是 PNG。所以本仓库**用 magic bytes 判断格式**（PNG `89 50 4E 47`、JPEG
`FF D8 FF`、GIF `GIF8`、BMP `BM`、TIFF `49 49 2A 00`、EMF 在第 40 字节处有
`" EMF"`、WMF `D7 CD C6 9A`），记录类型只当参考。

`pib` 在哪：形状的 `OfficeArtFOPT`（`0xF00B`）属性表里，`pid = 0x0104`。

```text
OfficeArtFOPT 载荷 = instance 个 6 字节条目
  条目 = pid u16 | value u32
  pid 低 14 位才是属性号：bit14 = fBid（值是 BLIP 序号），bit15 = fComplex
  真实文件里 pib 写作 0x4104（带 fBid），不是 0x0104 —— 必须 & 0x3FFF
```

`image.ppt` 幻灯片里图片形状的 FOPT 原样字节（偏移 44449）：

```text
7f 00 80 00 80 00      pid 0x007F  = 0x00800080   （fHaveAnchor 等标志位）
04 41 01 00 00 00      pid 0x4104  = 1            ← pib = 1，指向第 1 张图
05 81 0c 00 00 00      pid 0x8105  = 12           （fComplex 的 pibName）
```

Escher 容器层级（`image.ppt` 幻灯片实测）：

```text
PPDrawing 0x040C
└─ 0xF002 DgContainer
   └─ 0xF003 SpgrContainer
      └─ 0xF004 SpContainer          ← 一个形状
         ├─ 0xF00A Sp               几何（type / flags / id）
         ├─ 0xF00B FOPT             属性表（pib、填充、线条）
         ├─ 0xF010 ClientAnchor     在页面上的位置
         ├─ 0xF011 ClientData       形状类型（图片 = 0x000B）
         └─ 0xF00D ClientTextbox    文字（如果有）
```

`office_ppt::pictures()` 提取出的 PNG 与插入的源文件**逐字节相同**（548 B），
`Slide::pictures` 给出该页用到的 pib 列表。

## 8.1 文字格式：StyleTextPropAtom

文字格式不在文本记录里，而在紧随 `TextCharsAtom` 的 `StyleTextPropAtom`
（`0x0FA1`）里。它的 payload 是**两段首尾相接的数组**，没有任何长度字段
（MS-PPT 2.9.35）：

```text
rgTextPFRun:   每段 count u32 + indentLevel u16 + TextPFException
rgTextCFRun:   每段 count u32 + TextCFException
TextPFException = PFMasks u32 + 被掩码选中的字段（按固定顺序）
TextCFException = CFMasks u32 + 被掩码选中的字段（按固定顺序）
```

`count` 是该段覆盖的字符数。**麻烦在于两段数组之间没有分隔标记**：PF 段在哪里
结束、CF 段从哪里开始，只能靠「CF 段必须正好走到 payload 末尾」反推。本仓库的
做法是从小到大枚举 PF 段的每个可能边界，逐个试解 CF 段，取第一个能精确走到末尾
的（`parse_style`）。

`samples/你好.ppt`（Microsoft PowerPoint 2016 写的）里标题那一段，26 字节全部
解完：

```text
03 00 00 00               # PF: count 3（"你好" 2 字 + 段落回车）
00 00                     # indentLevel 0
00 08 00 00               # PFMasks 0x00000800 = bit11 align
01 00                     # align = 1（居中）
03 00 00 00               # CF: count 3
00 00 06 00               # CFMasks 0x00060000 = bit17 size | bit18 color
3c 00                     # size 60（点，不是半磅）
ff 00 00 fe               # color: r=FF g=00 b=00，索引 0xFE 表示"用这组 RGB"
```

同一文件副标题那一段只有 20 字节：PF 段 `06 00 00 00 00 00 00 00 00 00`（count 6、
无掩码），CF 段 `06 00 00 00 | 00 00 02 00 | 18 00`（24 磅）。

用到的掩码位（`PFMasks` 见 2.9.29、`CFMasks` 见 2.9.6，位序从 0 起）：

| CF 位 | 含义 | 宽度 |
| --- | --- | --- |
| 0 / 1 / 2 | bold / italic / underline（三个位共用 2 字节 `fontStyle`） | 2 |
| 16 | typeface（`fontRef`，主字体表下标） | 2 |
| 17 | size（点） | 2 |
| 18 | color（`ColorIndexStruct`：r, g, b, 索引；索引 `0xFE` = 用 RGB） | 4 |
| 19 | position（上/下标） | 2 |
| 21 / 22 / 23 | 旧东亚字体 / 西文 / 符号字体引用 | 各 2 |

| PF 位 | 含义 | 宽度 |
| --- | --- | --- |
| 0..=3 | 项目符号标志（共用 2 字节 `bulletFlags`） | 2 |
| 4 / 5 / 6 / 7 | bullet 字体 / 颜色 / 字号 / 字符 | 2 / 4 / 2 / 2 |
| 11 | align（0 左 1 中 2 右 3 两端） | 2 |
| 12 / 13 / 14 | 行距 / 段前 / 段后 | 各 2 |
| 8 / 10 / 15 | 左缩进 / 首行缩进 / 默认制表位 | 各 2 |
| 20 | tabStops（count u16 + count×4） | 变长 |
| 16 / 17..19 / 21 | fontAlign / wrapFlags / textDirection | 各 2 |

三个坑：

1. **数组之间没有长度字段**，只能靠"CF 段精确走到末尾"确定边界（见上）。
2. **空文本框也有 `StyleTextPropAtom`**。若实现成"格式化上一个文本记录"，空正文框
   的格式会盖到前一个形状上——实测 `samples/pptfmt.ppt` 第 2 页标题被第 2 页空
   正文的 24 磅覆盖。正确做法是记住"紧邻的那个文本记录"的序号，用过就失效。
3. **PF 段的 count 有时不含段末回车**：`samples/hello.ppt` 一个 2 字符的
   `cch`，PF 段只写 count 1。所以不能像规范要求的那样断言"各段 count 之和 ==
   字符数"，只能要求不超过。

早期笔记里"count 字段在 WPS 样本间不一致（1 / 2 / 0 / 6）"其实是把 **PF 段的
count** 当成了 CF 段的——两段数组是分开的，各自都自洽。

代码：

```rust
use office_ppt::{style_runs, TextFormat};

// payload 是 0x0FA1 的 payload，cch 是前一个文本记录里的字符数
let runs: Vec<(usize, usize, TextFormat)> = style_runs(payload, cch);
for (start, end, fmt) in runs {
    println!("{start}..{end} {:?}", fmt);   // fmt.size / fmt.bold / fmt.colour ...
}
```

`officedump html` 会把它们渲染成内联样式：

```html
<h2 style="text-align:center"><span style="font-size:60pt;color:#FF0000;">你好</span></h2>
<p><span style="font-size:24pt;">这是副标题</span></p>
```

## 8.2 占位符继承：主母版的 `TextMasterStyleAtom`

幻灯片里的文本 run 只写"和默认不一样"的属性。所以"副标题 24 磅"这种信息在
`StyleTextPropAtom` 里常常根本没有——它在**主母版**的 `TextMasterStyleAtom`
（`0x0FA3`）里，按占位符类型和缩进级别查表继承。

### 记录形状

主母版的 `PPDrawing` 里有 8 个 `0x0FA3`，`recInstance` 就是"样式类型"：

| instance | 含义 | 级别数 |
|---|---|---|
| 0 / 4 | 标题 | 1 |
| 1 / 5 | 正文 | 5 |
| 2、6、7、8 | 其他形状 | 5 或 1 |

实例号 0/1 是母版自己的占位符，4/5 是"九级样式"那一套；两者都在时取小的那个。

payload 是 `cLevels u16`，然后**每个级别**依次一个 `TextPFException`
（掩码 + 字段）和一个 `TextCFException`——和 `StyleTextPropAtom` 的 run 不同，
这里没有 `count`，因为一个级别就是整级。

```text
05 00                                   # cLevels = 5
00 00 00 00                             # 级别 0 的 PF 掩码 = 0（无段落属性）
00 00 02 00                             # 级别 0 的 CF 掩码 = 0x00020000（只有字号）
18 00                                   # 24 磅
01 00 00 00 ...                         # 级别 1 …
```

### 实测

```text
$ officedump info samples/你好.ppt | grep master
master      title 1 levels, 44pt; body 5 levels, 28pt; other 5 levels, 12pt

$ officedump info samples/hello.ppt | grep master
master      title 1 levels, 36pt; body 5 levels, 18pt; other 5 levels, 12pt
```

`你好.ppt`（微软 PowerPoint 2016）的 44 磅标题、正文 28/24/20/18 磅阶梯，
正是 PowerPoint 的默认值；`hello.ppt`（WPS）则是它自己的 36 磅主题。

### 合并规则

**run 自己写的属性永远赢**，母版只补 `None`：

```rust
let deck = office_ppt::parse_stream(&doc, user.as_deref())?;
let body = deck.master_styles.get(&office_ppt::TextKind::Body).unwrap();
let size = body[0].size;                  // 正文一级的默认字号
for run in &deck.slides[0].texts {
    println!("{:?} {:?} {:?}", run.kind, run.text, run.format.size);
}
```

`inherit_master_styles` 就是这个合并：`merged = master[level]; merged.absorb(&own)`。

### 坑

1. **不要用 `TextHeaderAtom` 判断类型**：母版里这些 atom 出现在占位符形状之前，
   前面根本没有 `TextHeaderAtom`。用 `recInstance`（§8.2 表格）。
2. **级别 4 有时解不出来**：部分 atom 的尾部字段宽度和规范对不上（`0x0FA3`
   的 62 字节变体末尾多 2 字节），所以最后一级可能缺属性。缺就缺，不要瞎猜——
   继承是"补默认值"，宁可少补也不能补错。
3. 备注页用的是 **NotesMaster**，不是主母版，本仓库不处理它的样式继承。

## 9. Rust API

```rust
use office_ppt::{parse_stream, TextKind};

let cfb = office_cfb::Cfb::open("samples/hello.ppt")?;
let data = cfb.stream("PowerPoint Document")?;
let user = cfb.stream("Current User").ok();
let pres = parse_stream(&data, user.as_deref())?;

println!("slides={} ordered_by_persist_directory={}", pres.slides.len(), pres.ordered_by_persist_directory);
for s in &pres.slides {
    println!("--- slide {} (persist {}) @ {:#X}", s.index + 1, s.persist_id, s.offset);
    for t in &s.texts {
        let tag = match t.kind { TextKind::Title => "T", TextKind::Body => "B",
                                 TextKind::Notes => "N", TextKind::Other => "?" };
        println!("[{tag}] {}", t.text);
    }
}
```

图片：

```rust
let pics = office_ppt::pictures(&cfb.stream("Pictures")?);
for p in &pics {
    println!("#{} {:?} {} 字节 → 写成 image{}.{}",
        p.index, p.kind, p.data.len(), p.index, p.kind.extension());
}
// 幻灯片上用到的图：pres.slides[0].pictures == [1]（pib 列表）
```

CLI：

```console
$ cargo run -q -p officedump -- text samples/hello.ppt
$ cargo run -q -p officedump -- tree samples/hello.ppt        # 打印记录树
$ cargo run -q -p officedump -- media samples/image.ppt ./out  # 导出图片
$ cargo run -q -p officedump -- html samples/hello.ppt > /tmp/p.html
```

## 9.1 写：生成一个 .ppt

写入器在 `crates/office-ppt/src/write.rs`，两个入口：

```rust
use office_ppt::write::{write_presentation, write_from_slide_template, SlideSpec};

let deck = vec![
    SlideSpec::new("第一页标题").with_body(&["正文第一行", "正文第二行"]),
    SlideSpec::new("Second slide").with_body(&["中文幻灯片内容"]),
];

// 自包含：模板已经编译进 crate，不需要任何外部文件
std::fs::write("out.ppt", write_presentation(&deck)?)?;

// 想用别的版式时，换成自己提供的模板（同样必须是微软 PowerPoint 写出的）
let tpl = std::fs::read("samples/ms/minimal.ppt")?;
std::fs::write("out2.ppt", write_from_slide_template(&tpl, &deck)?)?;
```

`write_presentation` 用 `include_bytes!("../templates/minimal.bin")` 内嵌了一份
**裁剪后的 `PowerPoint Document` 流**（203 KB，由 `tools/extract_ppt_template.py`
从微软 PowerPoint 写出的空白稿里提取），克隆其中首张幻灯片的形状树再替换文字。
这样调用方不需要装 Office、也不需要准备模板文件。

### 必须逐字节对上的记录

| 记录 | 要求 | 踩错的表现 |
|---|---|---|
| `SlidePersistAtom` | **20 字节**：`persistIdRef, cPersist(0), flags(0), slideId, reserved(0)` | 只写 4 字节，整份文件被拒 |
| 容器记录 | `ver_inst` 的 ver **必须是 0xF**（`0x1F` = SlideListWithText inst 1） | 写成 ver 0 时 PowerPoint 认不出子记录 |
| `SlideAtom` | **ver = 2**，`masterIdRef = 0x8000_0000 \| id`，`slideFlags = 7` | 版本号错时整份文件被拒 |
| `UserEditAtom` | **28 字节**：末尾 `lastView u16` + `unused u16`，`persistIdSeed` = 条目数 | 少 2 字节 → recLen 0x1A |
| `PersistDirectoryAtom` | 首字 = `1 \| (count << 20)`，count = 2 + 幻灯片数 | 少算时后面的对象全找不到 |
| `Current User` | `size(0x14) headerToken offsetToCurrentEdit lenUserName docFileVersion(0x03F4) major(3) minor(0) unused ansiName relVersion(8) unicodeName` | 多写一个长度前缀 → PowerPoint 判定文件损坏 |

`Current User` 的 `offsetToCurrentEdit` 必须指向 `UserEditAtom` 的偏移，`headerToken` 是 `0xE391C05F`。

### 验证到哪一步了

- `cargo run --example mkfiles` 生成 `samples/written.doc`、`written-formatted.doc`、`written-long.doc`、`written.xls`、`written.ppt`。
- **Word 与 Excel 都接受我们写的 `.doc`/`.xls`**（`tools/roundtrip.ps1` 对拍：段数、字数、粗体/字号/颜色/对齐、单元格值全部一致）。
- `.ppt`：**从零写已经打通**（`write_presentation` 用内嵌模板，见 §9.4），模板必须由**微软 PowerPoint 自己写出**（`samples/ms/minimal.ppt`，`Presentations.Add(1)` 生成的空白稿）。
  - 用 `hello.ppt`（WPS 写的）当模板时，克隆出来的文件 PowerPoint 拒收；
  - 用 `minimal.ppt` 当模板、并且**不改文字**时，输出与模板的 `PowerPoint Document` 流除 `UserEditAtom` 外**完全相同**，PowerPoint 正常打开（`slides=2`，形状文字正确）—— 说明 Document 组装、对象布局、persist 目录、`Current User` 全部正确；
  - **替换文字** → 一开始被拒收（「Office 检测到文件有问题」）。定位方法见 §9.3，根因是 `TextSpecialInfoAtom` 的首字 `cch` 没跟着改；修好后 **PowerPoint 正常打开，两页文字（含中文）全对**。

### 9.2 克隆模板时踩到的三件事

| 现象 | 原因 | 处理 |
|---|---|---|
| Slide 记录声明长度比实际短，文件被拒 | `container()` 按子记录算长度，之后追加的字节不算进去 | 手工拼 payload：`slide_atom + drawing + trailing` 再 `ppt(0xF, 0, RT_SLIDE, &body)` |
| SlideAtom 与尾随记录被复制两遍 | 克隆范围取了整个 Slide 体 | 只克隆 `PPDrawing` 记录的范围（`drawing_range`） |
| 文字一变长就拒收 | `StyleTextPropAtom` 里 PF run 与 CF run 的 `count` 都是「字符数 + 1」（含段落标记），且文字原子本身**不带结尾 `\r`** | 用 `pf_exception`/`cf_exception` 精确定位两段，重新生成原子（`restyle`），两处 `count` 都写 `cch + 1` |
| 换文字后仍拒收（最难的一个） | `TextSpecialInfoAtom`（0x0FAA）的**首个 u32 是它覆盖的字符数**，PowerPoint 会拿它和文字原子对账；克隆模板时这个值还停在模板文字的旧长度 | 改文字时一并把首字改成新长度（`info[0..4] = cch`），见 §9.3 |

`TextBytesAtom`（0x0FA8）只用于纯 ASCII；含非 ASCII 时必须换成 `TextCharsAtom`（0x0FA0，UTF-16LE）。微软写 ASCII 用 0x0FA8，WPS 写 0x0FA0。

### 9.3 怎么找到那条「隐含校验」

规范里没有一条说 PowerPoint 会因为「文字长度与某个附带字段不一致」拒收文件，报错也永远是同一句「Office 检测到文件有问题」。定位办法是**让 PowerPoint 自己写一份同样文字的文件，然后 diff**：

1. 用 COM 生成 `samples/ms/same.ppt`：空白稿 + 标题 `WRITER-PPT-001` + 正文 `first body line` / `second line`（与写入器的目标完全一致）；
2. 用 `olefile` 取出两个文件的 `PowerPoint Document`，找到最后一个 `Slide` 记录，递归打印子树里每个记录的 `ver/instance/kind/len` 与文字相关原子的十六进制；
3. 逐条对比 —— 形状树、`FOPT`、`ClientData`、`TextHeaderAtom`、`TextBytesAtom`、`StyleTextPropAtom` 全都一致，只有一处不同：

```
PowerPoint: 0x0faa len=28  0e000000 07000000 0000 0904 0408 0100 ...
我们的:     0x0faa len=28  0b000000 07000000 0000 0904 0408 0100 ...
                          ^^^^^^^^ 0x0E = 14 = 标题字符数；我们是模板里 "MIN-PPT-001" 的 11
```

`TextSpecialInfoAtom` 的首字就是「这段文字有多少字符」，正文那条同理（PowerPoint 27、我们 12）。把它一起改掉之后 PowerPoint 立刻接受。

> 教训：这类「结构全对但被拒收」的问题**不要猜规范，去造一份对照文件**。同一份内容让目标程序自己写一遍，diff 出来的那几字节就是答案 —— 比读十页 MS-PPT 快得多。

### 9.4 内嵌模板：为什么不能只留「一个母版 + 一个版式」

为了让 `write_presentation` 自包含，我试过把模板裁剪成 `Document + 主母版 + 一个版式 + 一张幻灯片`（203 KB → 19 KB），
并把 persist 目录、`UserEditAtom.persistIdSeed`、两个 `SlideListWithText` 的原子（`cPersist`、`slideId`）全部改写成自洽的值
（幻灯片 `SlideAtom.masterIdRef = 0x8000000C` 正好指向保留下来的那个版式）。**PowerPoint 仍然拒收**，逐级实验：

| 变体 | 内容 | 结果 |
|---|---|---|
| A 对照 | 原样 14 个对象 | 通过 |
| B | 14 个对象，但母版列表只列 1 个原子 | 拒收 |
| C | 只留 `Document + 母版 + 版式 + 幻灯片`（4 个对象），列表自洽 | 拒收 |
| D | 同 C，但母版原子的 `cPersist` 改成 0（与微软写法一致）、`slideId` 逐一对齐 | 拒收 |

结论：PowerPoint 2016 对「空白稿的形状/版式集合」有比 persist 目录更严的隐含要求（它甚至拒绝一个**完全自洽的子集**），
没有可推导的最小形式。所以最终**保留整条 203 KB 的流**内嵌进 crate —— 体积换来的是「零外部依赖」，
而这正是 `write_presentation` 存在的意义。想再瘦身只能靠压缩（例如内嵌 deflate + 自带解压），代价是给零依赖的 crate 引入解码器。

> 这条也解释了 §9.3 的方法论为什么重要：遇到「结构全对却被拒收」，规范读不出来，**只能靠对照实验**；
> 而对照实验的结论有时是「这条路不通，换一条」。

### 9.4.1 试过「完全自己写」：容器全对，PowerPoint 仍拒收

既然模板是 203 KB 的二进制，我按用户要求试了一次**完全不依赖模板**的写入：`Document`、`MainMaster`、
每张 `Slide`、`PersistDirectoryAtom`、`UserEditAtom` 全部由代码生成，只有几块「不含文字」的数值表
（`DocumentAtom`、`Environment`、`DggContainer`、两个占位符的 `FOPT`/`ClientData`）以常量形式写在源码里。
生成的文件 3–19 KB，结构上与微软写出的 `samples/ms/minimal.ppt` 逐记录对得上（同一批 record id、同样的顺序、
同样 20 字节的 `SlidePersistAtom`、`SlideAtom` ver=2、28 字节的 `UserEditAtom`）。

逐级实验（每一级都用 `tools/check_ppt.ps1` 让真 PowerPoint 打开）：

| # | 这一级做了什么 | 结果 |
|---|---|---|
| 1 | 自己生成 Document + 母版 + 幻灯片（占位符用自己拼的 `FOPT`/`ClientData`） | 拒收 |
| 2 | 占位符改用微软的 `FOPT`/`ClientData` 常量 | 拒收 |
| 3 | 补上 `SlideViewInfo`(0x07D0)、两个 `0x0FD9`、主题 `0x0428`/`0x0423` | 拒收 |
| 4 | 母版里补齐 8 条 `TextMasterStyleAtom` 与 `0x07F0`(instance 6)、把母版的 `PPDrawing` 直接换成微软那份 | 拒收 |
| 5 | 连幻灯片记录也换成微软写的那 835 字节（只把 `masterIdRef` 指回我们的母版） | 拒收 |

第 4、5 级很有信息量：**连 PowerPoint 自己写的母版绘图和幻灯片记录都用上了，文件仍然被判「已损坏」**，
说明差异不在形状树，而在「多个持久对象之间的相互约束」——最可能是 `DggContainer` 里 `FDGG`(0xF006)
记录的 **id 簇分配**：微软的空白稿有 14 个对象共享同一份簇表（`已发 id = 0x4B`），而我们的文件只有 4 个对象，
两者不可能同时满足；`§9.4` 的裁剪实验（把 14 个对象砍成 4 个、全部改写成自洽值）也失败了，两条路收敛到同一个结论。

**结论**：`.ppt` 的形状/版式/簇分配是一张「全有或全无」的一致性网，没有可推导的最小形式。

### 9.4.2 用户提供的 `void.ppt` 破解了那个「隐含约束」：`FDGG` 的 id 簇分配

用户在 PowerPoint 里另存了一份**只有一张空幻灯片**的稿子（`samples/void.pptx`，我再用 COM 另存为
`samples/void-ms.ppt`，85504 B）。它和 `samples/ms/minimal.ppt` 的关键差别是**持久对象数量不同**：
`void-ms.ppt` 只有 3 个（Document / 主母版 / 幻灯片），`minimal.ppt` 有 14 个（Document / 主母版 / 11 个版式 / 幻灯片）。
两份对照，`DggContainer` 里那个 `FDGG`(0xF006) 的规律就出来了：

```
FDGG 载荷 = spidMax u32 | 持久对象数 u32 | 已发 id 数 u32 | 簇数 u32 | 簇数 × (簇号 u32, 已用 u32)

void-ms.ppt : 01 08 00 00 | 03 00 00 00 | 07 00 00 00 | 02 00 00 00 | (1,7) (2,1)
minimal.ppt : 04 34 00 00 | 0e 00 00 00 | 4b 00 00 00 | 0d 00 00 00 | (1,7) (2,7) … (12,7) (13,4)
```

而每个对象自己的 `OfficeArtFDG`(0xF008) 正好对应一个簇（簇号 `i` 覆盖 `i<<10 .. (i<<10)+0x3FF`）：

| 文件 | 对象 | `FDG` 载荷 | 簇 | 形状数 |
|---|---|---|---|---|
| void | 主母版 | `06 00 00 00 06 04 00 00` | 1（spid `0x400..0x406`） | 7 |
| void | 幻灯片 | `01 00 00 00 00 08 00 00` | 2（spid `0x800`） | 1 |
| minimal | 主母版 | `06 00 00 00 06 04 00 00` | 1（spid `0x400..0x406`） | 7 |
| minimal | 幻灯片 | `03 00 00 00 03 34 00 00` | 13（spid `0x3400..0x3403`） | 4 |

**规则**：每个持久对象独占一个 id 簇，簇号按对象顺序递增（Document 不画图，不占簇），簇内 id 从
`簇号<<10` 起连续分配；`FDGG.spidMax` = 已分配的最大 spid + 1，`FDGG.簇数` = 对象数 − 1，
每个簇的 `已用` 就是那个对象的形状数。

这正好解释了 §9.4.1 的五次失败：我们的文件只有 4 个对象，却原封不动沿用了 `minimal.ppt` 的
「13 个簇 / 已发 0x4B」簇表，PowerPoint 一核对「对象数 × 簇表 × 每个 `FDG` 的 spid」就对不上。
这条规则已经落成代码：`crates/office-ppt/src/write.rs::fdgg_payload(&[u32])`
按「每个对象的形状数」生成簇表，两个单元测试分别用它重放 `void-ms.ppt`（3 对象 → `(1,7) (2,1)`）
与 `minimal.ppt`（14 对象 → 13 簇、`spidMax 0x3404`）。

验证方式是**直接读真实文件**：`matches_the_fdgg_inside_a_powerpoint_written_deck` 从
`samples/void-ms.ppt` 的 `PowerPoint Document` 流里按 `[ver_inst][kind][len]` 找到
`0xF006`（实测 `ver_inst = 0x0000`，`kind = 0xF006`，`len = 0x20`），拿它的载荷和
`fdgg_payload(&[0, 7, 1])` 逐字节比 —— 只跳过下面那个未解字段。

`已发 id 数` 这个字段当时没解出来：`void` 是 7（簇内 8 个 id）、`minimal` 是 75（88 个 id），
既不是总和也不是总和减簇数，所以代码里先按总和写，并在测试注释里标明"未解出、不依赖它"。

#### 9.4.3 `已发 id 数` 是「生产者相关」的：微软写 `Σused − 簇数`，WPS 写 `Σused`

`crates/officedump/examples/spcount.rs`（`cargo run -q --example spcount FILE`）把每个持久对象的
drawing 走一遍，数 `FSP`(0xF00A) 与 `FSPGR`(0xF009)，再把 Document 里 `0x040B → F000 → 0xF006`
的载荷整段打出来。12 份样本并列（`Σ` = 簇表 `used` 之和，`FSP`/`组` = 实测记录数）：

| 文件 | 写出者 | `Σ` | 簇数 | `已发 id` | `Σ−簇数` | `FSP` | 组 |
|---|---|---|---|---|---|---|---|
| `ms/minimal.ppt` | PowerPoint | 88 | 13 | 75 | **75** | 87 | 13 |
| `ms/same.ppt` | PowerPoint | 88 | 13 | 75 | **75** | 87 | 13 |
| `ms/fsn.ppt` | PowerPoint | 84 | 12 | 72 | **72** | 84 | 12 |
| `ms/pic.ppt` | PowerPoint | 89 | 13 | 76 | **76** | 88 | 13 |
| `ms/pic2.ppt` | PowerPoint | 90 | 13 | 77 | **77** | 89 | 13 |
| `ms/pic3.ppt` | PowerPoint | 90 | 13 | 77 | **77** | 89 | 13 |
| `ms/formatted.ppt` | PowerPoint | 106 | 16 | 90 | **90** | 103 | 16 |
| `void-ms.ppt` | PowerPoint | 8 | 2 | 7 | 6 ✗ | 8 | 2 |
| `hello.ppt` | WPS | 13 | 3 | 13 | 10 ✗ | 13 | 3 |
| `pptfmt.ppt` | WPS | 13 | 3 | 13 | 10 ✗ | 13 | 3 |
| `image.ppt` | WPS | 11 | 2 | 11 | 9 ✗ | 11 | 2 |
| `notes.ppt` | WPS | 27 | 6 | 27 | 21 ✗ | 27 | 6 |

微软写出的 7 份里有 6 份严格等于 `Σused − 簇数`；WPS 写出的 4 份一律等于 `Σused`；
`void-ms.ppt` 两者都不符（`Σ − 1`），可能因为它的簇 2 只发了 1 个 id（簇没填满）。

**这个字段不影响打开**：`write::fdgg_payload` 一直按 `Σused` 写，而由它组装的
`samples/scratch4-minimal.ppt`、`scratch5-*`、`scratch6-*` 都被真 PowerPoint 接受并正常显示形状。
所以保持现状，只把「未解」改成「生产者相关、PowerPoint 不校验」；测试里继续抹零后比，
注释指向本节。

于是结论要改：**不是不能从零写，而是之前把「对象数」这个维度当成了常量**。
> 教训：只要样例的代表性不够（同一份文件的两个变体对象数一样），就会把「随参数变化的量」误判成常量；
> 要破局就需要一份**参数不同**的对照文件 —— 这次靠用户提供。
因此 `write_presentation` 仍内嵌 203 KB 模板（换来零外部依赖），`write_from_scratch` 这条路撤掉；
真要瘦身只有两个方向：① 内嵌 deflate 压缩模板 + 自带解压（几十行 inflate）；② 逆向 `FDGG` 的簇分配算法
（需要微软写出的、对象数不同的对标文件，例如「空白稿 + 删掉所有版式」）。

### 9.5 图片：`Pictures` 流 + 图片形状（写到一半）

`.ppt` 的图片和 `.doc` 不同：图片**不放在幻灯片里**，而是集中放在一个独立的 `Pictures` 流中，
幻灯片上的图片形状只保存一个 `pib`（图片编号，从 1 开始）指向它。

实测（`samples/image.ppt`，WPS 写出、PowerPoint 能打开）：

```
Pictures 流（573 字节，就是一条 BLIP，没有 BStoreContainer 包着）
00 6e        ver 0, instance 0x6E0
1e f0        recType 0xF01E          ← 又是「PNG 写在 TIFF 编号里」
35 02 00 00  565 = 16 + 1 + 548
2b 0e fa 72 …  rgbUid（16 字节，同一张图在所有文件里都一样）
ff           tag
89 50 4e 47 …  PNG 本体
```

图片形状是 `OfficeArtSpContainer`（`0xF004`）里的三条记录：`Sp`（`0xF00A`，形状 id）、
`FOPT`（`0xF00B`，8 个属性，其中 `0x4104` = `pib`）、`ClientAnchor`（`0xF010`）。
关键属性 `0x007F = 0x00800080`（旋转 0、翻转 0、宽高各 1.0 的定点数），不是文本形状里的 `1`。

`crates/office-ppt/src/write.rs` 已经实现：`SlideSpec::with_picture(data, width, height)`
会把图片写进 `Pictures` 流、并把一个图片形状拼进幻灯片的 `OfficeArtSpgrContainer`（同时放大每一层外层容器的长度字段、
更新 `OfficeArtFDG` 的「最后一个形状 id」）。写出来的文件**我们的读取器能完整读回**
（`office_ppt::pictures()` 拿到的 PNG 与输入逐字节相同，`slide.pictures == [1]`，见单元测试
`writes_a_picture_into_the_pictures_stream`）。

**两条路径都已打通**（PowerPoint 2016 实测）：

| 路径 | 做法 | PowerPoint |
|---|---|---|
| A | `write_from_slide_template` 克隆一个本来就带图片的模板，只换文字 | 通过 |
| B | `write_presentation` 往一张没有图片的幻灯片里拼进图片形状 | 通过 |

B 路径一度被拒收，逐条排除后一共需要 **六处记账**，缺一处 PowerPoint 就报「文件有问题」：

| # | 要改的地方 | 细节 | 症状 |
|---|---|---|---|
| 1 | 形状 id 的位置 | 打在 `Sp` 的 payload 第一个 u32（记录头之后 16 字节），不是记录头 | `Sp` 的 `kind` 变 0，形状消失 |
| 2 | 形状 id 的取值 | 取幻灯片 `OfficeArtFDG` 的「最后一个 id + 1」（`0x3400 \| 簇号<<10`），不能用自增计数器 | id 落在 PowerPoint 未分配的簇外 |
| 3 | 幻灯片 `OfficeArtFDG`(0xF008) | `spidMax`（形状数）与 `cspidCur`（最后 id）各 +n | — |
| 4 | 文档 `OfficeArtFDGG`(0xF006) | `[0] spidMax`、`[8] 已发 id 数` 各 +n，**最后一个簇的 count +n** | — |
| 5 | 文档 `OfficeArtBStoreContainer`(0xF001) | 每张图一条 36 字节 `OfficeArtFBSE`：`size` = 该 BLIP 长度、`cRef` = 引用它的形状数、结尾 `00 00 ff ff` | 图片不显示 |
| 6 | `Pictures` 流 | 裸 BLIP 串联，**补零到 4096 字节**（PowerPoint 就是这么写的，也避免落进 mini 流） | 整份被拒 |

另外两个「不是记账、但同样致命」的坑：

- **空的文字原子会让整份文件损坏**。图片独占一张幻灯片时正文占位符没有文字，我们原来写出 `TextBytesAtom` 长度 0；PowerPoint 直接判「文件有问题」。空的段落要写成单个 `
`。
- **`cRef` 是「引用计数」不是「数量」**：同一张图放在两个形状里时 PowerPoint 只存一个 BLIP、`cRef = 2`（见 `samples/ms/pic2.ppt`）。`pib` 是这个 BLIP 在流里的 1 基下标。

实测结果（`tools/roundtrip.ps1` 的 ppt 段，PowerPoint 2016）：

```
ppt    slides=2
s1     WRITER-PPT-001 | first body line␍second line
s2     WRITER-PPT-002 | 第二页正文
pic    slide 1: type=14 720x188 | type=14 720x130.4 | type=13 96x96
```

`type=13` 就是图片形状（`msoPicture`）。

**`ClientAnchor` 已经全部解出**：PPT 里图片形状的 `OfficeArtClientAnchor`(0xF010) 只有 **8 字节，而且没有 flags**，
就是 `top, left, right, bottom` 四个 i16，单位 **1/8 pt**。

证据是 `samples/ms/pic3.ppt`（PowerPoint 自己把同一张图放在 40/100 处 96×96、300/100 处 96×48）：

| 图片 | 锚点原始字节 | 解码 | 核对 |
|---|---|---|---|
| 96×96 @ (40,100) | `20 03 40 01 40 04 20 06` | 100, 40, 136, 196 | 宽 = 136−40 = 96 ✓ 高 = 196−100 = 96 ✓ |
| 96×48 @ (300,100) | `20 03 60 09 60 0c a0 04` | 100, 300, 396, 148 | 宽 = 396−300 = 96 ✓ 高 = 148−100 = 48 ✓ |

所以写入器只要把 `right = left + 宽×8`、`bottom = top + 高×8` 写回去，尺寸就完全可控：

```
$ cargo run -q --release --example mkfiles   # 生成 samples/written.ppt 等
```

PowerPoint 实测（一张幻灯片放三张不同尺寸的图）：

```
OK  p5.ppt slides=1
  slide 1: type=14 720x188 | type=14 720x130.4 | type=13 96x96 | type=13 48x48 | type=13 120x60
```

> 早先踩过的坑：`picture_shape()` 的游标从 0 开始走记录，第一步读到容器头（len 98）就直接跳到末尾，
> `Sp`/`FOPT`/`ClientAnchor` 一个都没遍历到 —— 补丁「没生效」其实是被跳过了。从 8 开始才对。

## 10. 加密文档

加密的 .ppt **不是一条长 RC4 流**：每个持久对象（`Document`、`MainMaster`、每张 `Slide`）
各自一条流，**块号 = 持久 ID**，块大小按密钥位数取整（`keySize * (total/keySize + 1)`）。
加密头在 `CryptSession10Container`（`0x2F14`）里，偏移来自
`UserEditAtom.payload[28..32]` 的 `encryptSessionPersistIdRef`。

```rust
let deck = office_ppt::parse_stream_with_password(&stream, user.as_deref(), "secret123")?;
```

细节见 [`encryption.md`](encryption.md)。

## 11. 已知坑

1. **不要顺序扫 Slide 记录**，走 Current User → UserEditAtom → PersistDirectory（§5）。
2. PersistDirectoryAtom 的 count 有两种写法（§4），要容错。
3. `SlideListWithText` 的 instance 必须区分 0/1/2，否则母版会被当成幻灯片。
4. 文字可能在任何 Escher 容器深处，递归深度要有上限（本仓库 32 层），否则遇到
   损坏文件会栈溢出。
5. 记录 `len` 可能超出流末尾（截断文件），读取前先 clamp，否则 panic。
6. `TextBytesAtom` 是 8-bit，用 CP1252 解码；`TextCharsAtom` 是 UTF-16LE。判断错了
   会得到"每个字符后面多一个 \0"或乱码。
7. 加密的 .ppt（`EncryptedSummary` / CryptoAPI）里记录是密文，扫出来全是垃圾——
   先看 DocumentAtom 或 SummaryInformation 的加密标记。
8. 图片的 BLIP 记录类型会写错（§8），信 magic bytes 而不是信 `recType`。
9. `pib` 属性号带 `fBid` 位（`0x4104`），不 `& 0x3FFF` 就永远找不到图（§8）。
10. `Pictures` 流可能是裸 BLIP 而不是 `BStoreContainer`（WPS 就这样），两种都要处理。
11. `StyleTextPropAtom` 的两段数组没有长度字段，只能靠"CF 段正好走到末尾"定边界（§8.1）。
12. 空文本框也带 `StyleTextPropAtom`，别把它套到上一个形状的文本上（§8.1 坑 2）。
13. 格式化会把一个文本原子切成多个 run，输出纯文本时要按原子合并，否则每段格式
    都会变成一行（`Slide::paragraphs` 就是干这个的）。

### 9.5 从零写：`void-ms.ppt` 的完整记录骨架

`write_from_scratch` 要照着写的就是这张表（`samples/void-ms.ppt`，微软 PowerPoint 2016 写出的
最小演示文稿：1 张空幻灯片 + 11 个版式，共 3 个持久对象）。深度用缩进表示，`ver` 是记录头低 4 位、
`inst` 是高 12 位：

```
Document(0x03E8) ver=15 inst=0 len=3028
  DocumentAtom(0x03E9) ver=1 inst=0 len=40
  Environment(0x03F2) ver=15 inst=0 len=430
  PPDrawingGroup(0x040B) ver=15 inst=0 len=116
    DggContainer(0xF000) ver=15 inst=0 len=108
      FDGG(0xF006) ver=0 inst=0 len=32
      FOPT(0xF00B) ver=3 inst=6 len=36
      rec(0xF11E) ver=0 inst=4 len=16
  SlideListWithText(0x0FF0) ver=15 inst=1 len=28
  SlideViewInfo(0x07D0) ver=15 inst=0 len=296
  rec(0x0FD9) ver=15 inst=3 len=12
  rec(0x0FD9) ver=15 inst=4 len=12
  SlideListWithText(0x0FF0) ver=15 inst=0 len=28
  theme(zip)(0x0428) ver=0 inst=0 len=1986
  EndDocument(0x03EA) ver=0 inst=0 len=0
MainMaster(0x03F8) ver=15 inst=0 len=33383
  SlideAtom(0x03EF) ver=2 inst=0 len=24
  SlideShowSlideInfoAtom(0x07F0) ver=0 inst=6 len=32
  TextMasterStyle(0x0FA3) ver=0 inst=0 len=62
  TextMasterStyle(0x0FA3) ver=0 inst=1 len=118
  TextMasterStyle(0x0FA3) ver=0 inst=2 len=110
  TextMasterStyle(0x0FA3) ver=0 inst=4 len=110
  TextMasterStyle(0x0FA3) ver=0 inst=5 len=82
  TextMasterStyle(0x0FA3) ver=0 inst=6 len=12
  TextMasterStyle(0x0FA3) ver=0 inst=7 len=62
  TextMasterStyle(0x0FA3) ver=0 inst=8 len=62
  theme(zip)(0x0423) ver=0 inst=0 len=1635
  PPDrawing(0x040C) ver=15 inst=0 len=9134
    DgContainer(0xF002) ver=15 inst=0 len=9126
      DgAtom(0xF008) ver=0 inst=1 len=8
      SpgrContainer(0xF003) ver=15 inst=0 len=7705
      SpContainer(0xF004) ver=15 inst=0 len=1389
  SlideShowSlideInfoAtom(0x07F0) ver=0 inst=1 len=32
  rec(0x040E) ver=0 inst=0 len=3448
  rec(0x040F) ver=0 inst=0 len=314
  layout(zip)(0x041E) ver=0 inst=11 len=1548
  layout(zip)(0x041E) ver=0 inst=10 len=1484
  layout(zip)(0x041E) ver=0 inst=9 len=2221
  layout(zip)(0x041E) ver=0 inst=8 len=1850
  layout(zip)(0x041E) ver=0 inst=7 len=1154
  layout(zip)(0x041E) ver=0 inst=6 len=1332
  layout(zip)(0x041E) ver=0 inst=5 len=1914
  layout(zip)(0x041E) ver=0 inst=4 len=1560
  layout(zip)(0x041E) ver=0 inst=3 len=1726
  layout(zip)(0x041E) ver=0 inst=2 len=1453
  layout(zip)(0x041E) ver=0 inst=1 len=1654
  rec(0x041C) ver=0 inst=0 len=4
  rec(0x0FBA) ver=0 inst=2 len=22
Slide(0x03EE) ver=15 inst=0 len=176
  SlideAtom(0x03EF) ver=2 inst=0 len=24
  PPDrawing(0x040C) ver=15 inst=0 len=80
    DgContainer(0xF002) ver=15 inst=0 len=72
      DgAtom(0xF008) ver=0 inst=2 len=8
      SpgrContainer(0xF003) ver=15 inst=0 len=48
  SlideShowSlideInfoAtom(0x07F0) ver=0 inst=1 len=32
  rec(0x0422) ver=0 inst=0 len=8
PersistDirectoryAtom(0x1772) ver=0 inst=0 len=16
UserEditAtom(0x0FF5) ver=0 inst=0 len=28
```

三个可以直接照抄的结论：

1. **只有真正的容器是 `ver = 0xF`**（`Document`/`MainMaster`/`Slide`/`PPDrawing`/`PPDrawingGroup`/`DgContainer`/`SpgrContainer`/`SpContainer`/`DggContainer`/`Environment`/`SlideViewInfo`/`SlideListWithText`）。簿记类记录（`FDGG`、`UserEditAtom`、`PersistDirectoryAtom`、`theme`、`EndDocument`）是 `ver = 0`；**而 atom 各有自己的固定 `ver`**：`DocumentAtom` = 1、`SlideAtom` = 2、`FOPT` = 3、`SlidePersistAtom` = 0 —— 写错同样会被拒收。
2. **母版的 8 条 `TextMasterStyleAtom` 的 `inst` 是 0/1/2/4/5/6/7/8**（没有 3），顺序固定；这是占位符继承的类型表（§8.2）。注意 `Environment` 里还带**一条** `inst = 4` 的同名记录（默认样式），扫描时要跳过它。
3. **版式 ZIP 的顺序是 11…1 递减**（`0x041E` 的 `inst` 从 11 降到 1），即 Excel/PowerPoint 把「最后一个版式」写在最前面。写反了 PowerPoint 也能打开（它按 `inst` 索引），但逐字节对拍会失败。

要复现这条骨架，需要的**不透明载荷**只有三处：`theme` 的 1986/1635 字节 ZIP、`0x040E` 的 3448 字节 ZIP、
以及 11 个版式 ZIP（合计约 19 KB）。其余全部可由字段生成（§9.4.2 的簇表 + §8.2 的母版样式 + 占位符的
`FOPT`/`ClientData`/`ClientTextbox`）。**这三处后来被证明可以整个丢掉**（§9.22）：它们是 Office 2007
的回写缓存，PowerPoint 打开时会自己重建，写入器现在不再携带。

### 9.6 `DocumentAtom`（40 字节）实测

`samples/void-ms.ppt` 里那条（`ver = 1`、`len = 40`）：

```
00 1e 00 00 | e0 10 00 00 | e0 10 00 00 | 80 16 00 00 | 01 00 00 00 | 02 00 00 00 |
00 00 00 00 | 00 00 00 00 | 01 00 06 00 | 00 00 00 01
```

当时只对上了前 16 字节（两个尺寸），后面的字段全是猜的；**完整字段表见 §9.6.4**，
按 MS-PPT §2.4.2 逐字段核对后的正确读法：

| 偏移 | 字段 | 本例值 | 含义 |
|---|---|---|---|
| 0 | `slideSize.x` | `0x1E00` = 7680 | 8 英寸 × 960 |
| 4 | `slideSize.y` | `0x10E0` = 4320 | 4.5 英寸 |
| 8 | `notesSize.x` | `0x10E0` = 4320 | 备注页宽 |
| 12 | `notesSize.y` | `0x1680` = 5760 | 备注页高 |
| 16 | `serverZoom.numerator` | 1 | 比例 1:2 |
| 20 | `serverZoom.denominator` | 2 | |
| 24 | `notesMasterPersistIdRef` | 0 | 本文件没有备注母版 |
| 28 | `handoutMasterPersistIdRef` | 0 | 没有讲义母版 |
| 32 | `firstSlideNumber` | 1 | |
| 34 | `slideSizeType` | 6 | `Custom`（0 基枚举） |
| 36–39 | 四个 `bool1` | `0 0 0 1` | `fSaveWithFonts`…`fShowComments` |

### 9.7 `SlideAtom` / `SlidePersistAtom` 已可生成（与真实文件对拍通过）

| 记录 | 长度 | 布局（实测） |
|---|---|---|
| `SlideAtom` 0x03EF | 24 B，`ver = 2` | `geom u32` + `rgPlaceholderTypes [u8;8]` + `masterIdRef u32` + `notesIdRef u32` + `slideFlags u16` + `unused u16` |
| `SlidePersistAtom` 0x03F3 | 20 B，`ver = 0` | `persistIdRef u32` + `cPersist u32` + `flags u32` + `slideId u32` + `reserved u32` |

实测值（`samples/void-ms.ppt`）：

- 主母版 `SlideAtom` = `geom 1`、占位符类型 `[1,2,7,9,8,0,0,0]`、`masterIdRef 0`、`notesIdRef 0`、`slideFlags 2`。
- 幻灯片 `SlideAtom` = `geom 0x10`（16 = 全屏）、占位符类型全 0、`masterIdRef 0x80000000`、`notesIdRef 0`、`slideFlags 7`。
- 母版列表原子 `(persist 2, cPersist 0, slideId 0x80000000)`；幻灯片列表原子 `(persist 3, cPersist 0, slideId 0x0102)`。

代码：`write.rs::slide_atom_full()` / `slide_persist_atom_full()`，两个测试
（`slide_atoms_match_the_reference_deck`、`slide_persist_atoms_match_the_reference_deck`）
直接读 `void-ms.ppt` 逐字节比对。**这两个 atom 现在已经可以完全由字段生成**，
剩下的是 theme 的 ZIP 负载（`DocumentAtom` 已全字段可生成，§9.6.5；这处 ZIP 其实可以不写，§9.22）。

### 9.8 `Environment`（430 字节）的六个子记录

`Environment` 是个容器，实测子记录（`samples/void-ms.ppt`，测试 `environment_children_are_the_six_expected_records` 固化）：

| 记录 | `inst` | 长度 | 内容 |
|---|---|---|---|
| `0x0FC8` ColorScheme | 2 | 12 | 配色方案引用（`30 00 d2 0f 04 00 00 00 00 00 00 00`，指主题里的配色） |
| `0x07D5` FontCollection | 0 | 228 | **3 条** `0x0FB7` FontEntity（每条 68 B：名字 UTF-16 补满 64 B + charset/family/pitch/pad） |
| `0x0FA4` TextRuler | 0 | 10 | `80 00 60 00 00 00 ff ff 00 00` |
| `0x0FA5` | 0 | 12 | 默认制表位 |
| `0x0FA9` | 0 | 10 | |
| `0x0FA3` TextMasterStyle | **4** | 110 | 文档级默认样式（母版那 8 条之外的额外一条，§9.5 的坑） |

合计 20 + 236 + 18 + 20 + 18 + 118 = **430** ✓。

这六条里，`0x0FA3` 已经能由 §8.2 的 `master_style_levels()` 解析；三条字体实体是「名字 + 4 个标志字节」的纯数据；
剩下 `0x0FC8`/`0x0FA4`/`0x0FA5`/`0x0FA9` 是 10–12 字节的常量。也就是说 `Environment` **可以完全生成**，
不需要任何不透明二进制 —— 只要把三条字体的名字与四个标志列成表。

**已验证**：`write.rs::environment(fonts, style)` 用字段拼出的 430 字节载荷与
`void-ms.ppt` **逐字节相同**（测试 `fonts_and_rulers_are_reproduced_from_their_fields`）。

过程中挖出三条只有对拍才看得见的细节：

1. `0x0FB7` FontEntity **不是容器**（`ver = 0`），而且它的 `instance` 是**字体下标**（0/1/2）——
   我一开始按容器写成 `ver = 0xF`，diff 到第 28 字节才暴露。
2. 三条标尺 `0x0FA4`/`0x0FA5`/`0x0FA9` 同样是 `ver = 0`。
3. 差 8 字节那次是我把「整条记录（含 8 字节头）」和「载荷」当成了同一个东西比 ——
   和 §15.3 里 `.xls` 那次口径混淆是同一类错误。教训：**比对前先确认两边是同一个口径**
   （记录 vs 载荷），否则差出来的 8 字节会被误当成格式问题去追。

### 9.9 `SlideListWithText` 也已可生成

`slide_list_with_text(instance, atoms)`：容器记录（`ver = 0xF`），`instance` 0 = 幻灯片、1 = 母版、2 = 备注，
载荷是**完整的** `SlidePersistAtom` 记录（`[ver_inst][kind 0x03F3][len]` + 20 字节），不是裸的 20 字节。

实测 `void-ms.ppt`：母版列表 28 B（1 个原子 `(2, 0, slideId 0x80000000)`）、幻灯片列表 28 B
（1 个原子 `(3, 0, slideId 0x0102)`），两条都被 `both_lists_in_the_reference_deck_are_reproduced` 逐字节复现。

**踩坑**：我一开始把 `slide_persist_atom_full()` 的 20 字节载荷直接塞进列表，结果比参考短 8 字节 ——
列表里的原子是**带记录头的完整记录**。这是本节第三次「记录 vs 载荷」口径错误，写下来提醒自己。

### 9.10 `PPDrawingGroup`（116 字节）可生成

```rust
drawing_group(per_object: &[u32]) -> Vec<u8>   // 116 字节载荷
```

结构（实测 `void-ms.ppt`）：

```
PPDrawingGroup 0x040B  ver 0xF  len 116
└─ DggContainer 0xF000 ver 0xF  len 108
   ├─ FDGG   0xF006 ver 0    len 32   ← fdgg_payload(per_object)，§9.4.2
   ├─ FOPT   0xF00B ver 3 inst 6 len 36   ← 文档级绘图属性（常量）
   └─ 0xF11E ver 0 inst 4 len 16        ← 常量
```

`the_reference_drawing_group_is_reproduced` 与真实文件逐字节比较，**除 `FDGG` 第三个字
（"已发 id 数"，生产者相关、不影响打开，见 §9.4.3）之外全部相等**；测试里把这个字显式抹零后再比，并在注释里说明原因。

至此 `Document` 的七个子记录里有四个（`Environment`、`PPDrawingGroup`、两条 `SlideListWithText`）
已经是**逐字节对拍通过**的生成器；当时剩下 `DocumentAtom` 尾部 16 字节（**已解，§9.6.4/§9.6.5**）、
`SlideViewInfo`(296 B)、两个 `0x0FD9`(12 B) 与 theme 的 1986 字节 ZIP。

### 9.11 `SlideViewInfo`（296 B）与两个 `0x0FD9`（12 B）

```rust
slide_view_info()  -> Vec<u8>   // 0x07D0 容器，ver 0xF，载荷 296 字节
slide_view_atoms() -> Vec<u8>   // 0x0FD9 inst 3 + 0x0FD9 inst 4，各 12 字节
```

这三条描述的是**编辑视图**（当前选中页、缩放、网格等），与幻灯片内容无关，对没有保存大纲的
演示文稿是常量。测试 `the_reference_view_records_are_reproduced` 与 `void-ms.ppt` 逐字节相等。

`0x0FD9` 的载荷（12 B）= `0x0FDA` 记录（4 B 载荷 `0d 00 02 00`）套在外层里；两个实例只差一个字节，
按 `inst` 3/4 分别是「当前页」与「当前页的备注」。

至此 `Document` 的七个子记录里**五个已是逐字节对拍通过的生成器**：
`Environment`、`PPDrawingGroup`、`SlideListWithText` ×2、`SlideViewInfo` + `0x0FD9` ×2。
只剩 theme 的 1986 字节 ZIP —— `DocumentAtom` 的 40 字节已全部可生成（§9.6.5），而这处 ZIP 可以不写（§9.22）。

### 9.12 `Document` 整条已能组装（3028 字节）

```rust
document(per_object: &[u32], style: &[u8], theme: &[u8], document_atom: &[u8]) -> Vec<u8>
```

十个子记录按实测顺序拼装，`EndDocument`（`0x03EA`）**在 `Document` 内部**：

| # | 记录 | 长度 | 来源 |
|---|---|---|---|
| 1 | `DocumentAtom` 0x03E9 | 40 | `write::document_atom()`（§9.6.5） |
| 2 | `Environment` 0x03F2 | 430 | `environment()` 生成 |
| 3 | `PPDrawingGroup` 0x040B | 116 | `drawing_group()` 生成 |
| 4 | `SlideListWithText` 0x0FF0 inst 1 | 28 | `slide_list_with_text()` |
| 5 | `SlideViewInfo` 0x07D0 | 296 | `slide_view_info()` |
| 6-7 | `0x0FD9` inst 3 / inst 4 | 12 / 12 | `slide_view_atoms()` |
| 8 | `SlideListWithText` 0x0FF0 inst 0 | 28 | `slide_list_with_text()` |
| 9 | theme 0x0428 | 1986 | 传入（ZIP；可省，§9.22） |
| 10 | `EndDocument` 0x03EA | 0 | 生成 |

`the_reference_document_record_is_reassembled` 与 `void-ms.ppt` 的 3028 字节比较，
**除 `FDGG` 第三个字外逐字节相同**（测试里抹零后比，注释写明原因）。

组装过程中又踩了两个坑，都是「参数没传对」而不是格式错：

1. `per_object` 少算了幻灯片自己的形状数（传 `[0,7]` 而不是 `[0,7,1]`），`FDGG` 因此少一簇、
   `PPDrawingGroup` 短 8 字节 —— 而单测 `drawing_group` 是好的，只有组装测试能抓到。
2. 早先我用「在字节流里搜 `28 04`」从参考里取 theme，会匹配到载荷里的巧合字节；
   改成**按记录顺序遍历**（`children()`）后一次就取对了。

#### 9.6.1 第二份样本：`firstSlideNum` 与 16:9

用 COM 另存了一份「起始页码 5、16:9」的空白稿 `samples/ms/fsn.ppt`（`PageSetup.FirstSlideNumber = 5`、
`PageSetup.SlideSize = 2`，脚本 `tools/tmp_fsn.ps1`），与 `void-ms.ppt` 的 `DocumentAtom` 对比：

```
void-ms: 00 1e 00 00 | e0 10 00 00 | e0 10 00 00 | 80 16 00 00 | 01 00 00 00 | 02 00 00 00 | 00…00 | 01 00 06 00 | 00 00 00 01
fsn    : 80 16 00 00 | e0 10 00 00 | e0 10 00 00 | 80 16 00 00 | 05 00 00 00 | 0a 00 00 00 | 00…00 | 05 00 01 00 | 00 00 00 01
```

已确认的字段（前 24 字节全部解开）：

| 偏移 | 字段 | void | fsn |
|---|---|---|---|
| 0 | `slideSizeX` | 7680 | **5760**（16:9 变窄） |
| 4 | `slideSizeY` | 4320 | 4320 |
| 8 | `notesSizeX` | 4320 | 4320 |
| 12 | `notesSizeY` | 5760 | 5760 |
| 16 | 随 `FirstSlideNumber` 变 | 1 | **5** |
| 20 | 同上的两倍 | 2 | **10** |

当时把偏移 16/20 当成 `firstSlideNum` 与它的两倍 —— **错**：那两格是 `serverZoom`，页码在偏移 32。
`fsn.ppt` 与 `fsn2.ppt` 的页码不同（5 与 1）而 16/20 不变，正好说明这一点；最终结论见 §9.6.3/§9.6.4。

#### 9.6.2 第三份样本推翻了「偏移 16 = 页码」的猜测

再存一份「16:9 但不设页码」的 `samples/ms/fsn2.ppt`（脚本 `tools/tmp_fsn2.ps1`），三份并列：

| 文件 | 尺寸 | 偏移 16 | 偏移 20 | 尾 12 字节 |
|---|---|---|---|---|
| `void-ms.ppt` | 4:3 7680×4320 | 1 | 2 | `00 00 00 00 01 00 06 00 00 00 00 01` |
| `fsn.ppt` | 16:9 5760×4320，页码 5 | 5 | 10 | `00 00 00 00 05 00 01 00 00 00 00 01` |
| `fsn2.ppt` | 16:9 5760×4320，页码 1 | 5 | 10 | `00 00 00 00 01 00 01 00 00 00 00 01` |

**当时的结论**（后两条已被 §9.6.3/§9.6.4 推翻，见下）：

* 偏移 16/20 跟着尺寸走 —— **错**，见 §9.6.4；
* **页码**在尾部偏移 32 的 `u16` 上：`fsn.ppt` = 5，另两份 = 1 —— 正确，5 份样本一致；
* 尾部偏移 34 的 `u16` 是「尺寸相关的标志位」—— 方向对，具体是 `slideSizeType` 的 0 基编号。

#### 9.6.3 五份新样本：偏移 34 = `slideSizeType - 1`

`tools/tmp_sizes.ps1` 用 COM 把 `PageSetup.SlideSize` 依次设成 `ppSlideSizeType` 的 1/3/4/6
（`FirstSlideNumber = 7`）另存为 `samples/ms/size{1,3,4,6}.ppt`；`tools/tmp_multi.ps1` 存一份
两页的 `samples/ms/multi3.ppt`。注意 `Presentations.Add()` 建出来的是**零页**演示文稿
（`Slides.Count = 0`，所以 `size*.ppt` 里没有 `0x03EE` 对象），加页要用 `Slides.Add(1, layout)`。

| 文件 | `SlideSize` | `slideSizeX/Y` | 偏移 16 | 偏移 20 | 偏移 32 | 偏移 34 |
|---|---|---|---|---|---|---|
| `void-ms.ppt` | 自定义 8×4.5 in | 7680 / 4320 | 1 | 2 | 1 | **6** |
| `ms/fsn.ppt` | 2 | 5760 / 4320 | 5 | 10 | 5 | **1** |
| `ms/fsn2.ppt` | 2 | 5760 / 4320 | 5 | 10 | 1 | **1** |
| `ms/size1.ppt` | 1 | 5760 / 4320 | 5 | 10 | 7 | **0** |
| `ms/size3.ppt` | 3 | 6240 / 4320 | 5 | 10 | 7 | **2** |
| `ms/size4.ppt` | 4 | 6480 / 4320 | 5 | 10 | 7 | **3** |
| `ms/size6.ppt` | 6 | 4608 / 576 | 5 | 10 | 7 | **5** |
| `ms/multi3.ppt` | 1（两页） | 5760 / 4320 | 5 | 10 | 1 | **0** |

* **偏移 34 就是 `slideSizeType`，0 基**：`OnScreen=0, Letter=1, A4=2, 35mm=3, Overhead=4,
  Banner=5, Custom=6`，即 `PpSlideSizeType - 1`。六份样本一一对应，没有例外（`void-ms.ppt` 是
  自定义尺寸 → 6）。
* **偏移 32 是 `firstSlideNum`**（u16），设成 7 的三份都写 `07 00`，确认。
* **偏移 16/20 是 `serverZoom` 分子/分母**，不是尺寸：四种不同尺寸（含 576×72 的横幅）全是
  `5 / 10`，两页的也是 `5 / 10`；`void-ms.ppt` 是 `1 / 2`。`5:10` 与 `1:2` 是**同一个比例**
  （50%），`void-ms.ppt` 由 `void.pptx` 转换而来，只是写了未约分/已约分两种形式。§9.6.2 里
  「偏移 16 跟着尺寸走」的观感，是因为当时的样本恰好一个 4:3 一个 16:9，比例与尺寸共变。

#### 9.6.4 偏移 16 的语义：MS-PPT 原文 + 备注母版实验

[MS-PPT §2.4.2 `DocumentAtom`](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ppt/121f2728-3497-4a0a-829e-6f416fee2ee6)
给出了完整字段表，40 字节一个不多一个不少：

```
0  slideSize.x  u32          4  slideSize.y  u32
8  notesSize.x  u32         12  notesSize.y  u32
16 serverZoom.numerator u32 20  serverZoom.denominator u32   (RatioStruct, MUST be > 0)
24 notesMasterPersistIdRef   u32  (0 = 没有备注母版)
28 handoutMasterPersistIdRef u32  (0 = 没有讲义母版)
32 firstSlideNumber u16     34  slideSizeType u16  (SlideSizeEnum)
36 fSaveWithFonts u8  37 fOmitTitlePlace u8  38 fRightToLeft u8  39 fShowComments u8
```

三条实测确认（`tools/tmp_docatom.ps1` → `samples/ms/atom-*.ppt`，`cargo run -q --example docatom`）：

| 变体 | 做法 | 偏移 16/20 | 偏移 24 |
|---|---|---|---|
| `atom-base.ppt` | `Presentations.Add(1)`（零页，11 版式） | 5 / 10 | 0 |
| `atom-zoom.ppt` | `Windows.Item(1).View.Zoom = 66` | 5 / 10 | 0 |
| `atom-notes.ppt` | 访问 `$p.NotesMaster` 造出备注母版 | 5 / 10 | **2** |
| `atom-fewlayouts.ppt` | 删到只剩 2 个版式 | 5 / 10 | 0 |
| `atom-notessize.ppt` | 改 `NotesOrientation` | 5 / 10 | 0 |

* 窗口缩放、版式个数、备注页尺寸都不动这两个字段 —— 只有「比例」在变，所以是 `serverZoom`。
* 偏移 24 在造出备注母版后变成 **2**：`pdump` 显示 `atom-notes.ppt` 的 id 2 是
  `0x03F0 NotesContainer`（`off=202531 len=19957`），正是备注母版对象 ⇒ 偏移 24 是
  `notesMasterPersistIdRef`（持久对象 id，不是 slideId）。`notes.ppt`（WPS 写）该字段是 7。
* 尾部 `00 00 00 01` 就是 `fSaveWithFonts=0, fOmitTitlePlace=0, fRightToLeft=0, fShowComments=1`
  四个 bool，按书写顺序照抄字节即可（这个项目里把标志位当整数读已经错过一次：`.xls` 的 `BOF`）。
* `slideSizeType` 就是 spec 的 `SlideSizeEnum`，0 基 —— 即 `PpSlideSizeType - 1`，两者不矛盾。

教训：**一份样本只能定「常量 vs 变量」，两份能定「哪个变量」，但两个变量共变时必须靠实验拆开**
—— 八份样本里 `5/10` 与 `type=1` 恰好同时出现在 4:3 上，看起来像「尺寸类型」，实际一个是比例、
一个是枚举。真正的拆解来自「只改一个东西」的 COM 实验和规范原文。

#### 9.6.5 已生成器化

```rust
document_atom(
    slide_size: (u32, u32), notes_size: (u32, u32), zoom: (u32, u32),
    first_slide_number: u16, slide_size_type: u16,
    notes_master: u32, handout_master: u32,
) -> Vec<u8>                                        // 40 字节
```

旧版本是 `document_atom(width, height, first_slide)`，内部按 `width == 5760` 猜尺寸类型 ——
那是错的（`fsn.ppt` 是 4:3 却写 `5:10` + `type=1`）。现在七个字段全由调用方给定，
`DEFAULT_ZOOM = (1, 2)` 与 `SLIDE_SIZE_ON_SCREEN = 0` / `SLIDE_SIZE_CUSTOM = 6` 是常量。

三个测试守着它：

* `document_atom_matches_the_measured_samples` —— 8 份样本的**显式参数表**逐字节比对
  （`void-ms`、`fsn`、`fsn2`、`size1/3/4/6`、`atom-notes`），缺样本时跳过；
* `document_atom_rebuilds_every_sample` —— 遍历 `samples/` 与 `samples/ms/`，用每份文件自己的
  字段重建再逐字节比，证明 40 字节**没有剩余未覆盖的字节**；
* `notes_master_reference_is_a_persist_id` —— `atom-notes.ppt` 的偏移 24 == 2，且该持久对象
  的 kind 是 `0x03F0`（`NotesContainer`）。

### 9.13 从零组装整份 `.ppt`：**PowerPoint 接受了**

`crates/officedump/examples/scratch3.rs` 用 `write::document()` + 生成的 `PersistDirectoryAtom`
/`UserEditAtom`/`Current User` 拼出 `samples/scratch3.ppt`（39424 B），其中 `MainMaster` 与 `Slide`
的形状树暂时以 `void-ms.ppt` 的 blob 携带（把变量控制在「文档级记账」这一层）。

```
Document 3036 B（载荷 3028，全部由代码生成）
MainMaster @3036 · Slide @36427 · PersistDirectoryAtom @36611 · UserEditAtom @36635
```

`tools/check_ppt.ps1` 结果：**`OK scratch3.ppt slides=1`** —— 真 PowerPoint 打开成功。

这是 §9.4.1 里五次被拒收之后第一次通过，关键差别就是 `FDGG` 簇表按「对象数」生成
（`per_object = [0, 7, 1]` → 两簇），而不是沿用 14 对象文件的 13 簇。

剩下的差异是幻灯片里没有形状（`slide 1:` 后面为空）：blob 里的 `SlideAtom.masterIdRef`
指向 `0x8000000C`（`void-ms.ppt` 的某个版式），而我们只保留了主母版。下一步把 `masterIdRef`
指回主母版（`0x80000002`）即可。

**回退记录**：把携带过来的 `SlideAtom.masterIdRef` 从 `0x8000000C` 改成 `0x80000002`（指回我们保留的
主母版）后，PowerPoint **立刻又报「文件有问题」** —— 说明 `masterIdRef` 不能指向主母版，幻灯片必须指向
一个版式（`0x03F8` 的某个实例），而我们只保留了主母版。这也解释了为什么 `slide 1:` 没有形状：
形状挂在版式上，版式缺席时 PowerPoint 静默丢弃。

所以「保留 1 个主母版 + 1 个版式」这条最小组合是对的（§9.4.1 的 C 变体就是这么做的，当时失败另有原因），
下一步应当保留主母版 + 它引用的那个版式，而不是把幻灯片指回主母版。

> **修正（§9.14）**：上面「形状挂在版式上」的说法不准确。版式只提供占位符模板；幻灯片可见的形状
> 在它自己的 `PPDrawing` 里。带进版式后 `slide 1:` 仍为空，是因为 `void-ms.ppt` 的 slide 本身没有形状。

### 9.14 形状在幻灯片自己的 drawing 里，不在版式上

`crates/officedump/examples/scratch4.rs` 的三个变体（真 PowerPoint 对拍）：

| 变体 | master | layout | slide | `tools/check_ppt.ps1` |
| --- | --- | --- | --- | --- |
| `sparse` | void-ms | minimal id3（persist 12，簇 11） | void-ms id3 | 拒收 |
| `renum` | void-ms | minimal id3（persist 4，簇 3，spid 0xC00..） | void-ms id3 | OK，`slide 1:` 空 |
| `minimal` | minimal id2 | minimal id3（persist 3，簇 2） | minimal id14（簇 13→3） | OK，`720x188 \| 720x130.4` |

`minimal.ppt` 的两个可见形状来自 **slide 自身**：`F00A` spid `0x3402`（title）与 `0x3403`（body），
`ClientAnchor` 分别为 `c3 02 c0 03 40 1a a3 08` / `dd 08 c0 03 40 1a f0 0c`。版式 `id3` 的 6 个形状
只是模板；`SlideAtom.rgPlaceholderTypes` 全 0 时 PowerPoint 不会实例化版式形状。所以「把版式带进来」
本身不会让空的 slide 出现形状。

### 9.15 加形状：克隆真实形状可以，`text_shape` 的合成树仍被拒收

`crates/officedump/examples/scratch5.rs` 在同一套「生成的 Document + 携带的 master/layout/slide」上做增量：

* `minclone`：复制 minimal slide 的最后一个 `SpContainer`（真实 body 形状），spid `0xC03→0xC04`，
  追加到 `0xF003`，`F008` 改 count `3→4` / spidMax `0xC03→0xC04`，FDGG 簇 3 的 used `4→5`。
  **PowerPoint 接受，`slide 1:` 出现 3 个形状**（原有 2 个 + 克隆 1 个）。
* `mintext`：改为追加 `write::text_shape(0xC04, 6, "Generated shape", ...)`，**仍被拒收** ——
  问题在 `text_shape` 的合成字节，不在追加逻辑。

与真实形状对比，`text_shape` 已补但仍缺：

* `F00A` 的 `ver` 必须是 2（`escher()` 原来写 0），`F00B` 的 `ver` 必须是 3；
* `F011` 必须是 `ver=0xF` 的容器，子记录 `0x0BC3`（8 B）+ `0x041F`（4 B，占位符序号）；
* 真实 `FOPT` 是 `instance=15`（15 条属性），`text_shape` 只有 8 条；仅补前两项仍被拒收，
  说明缺少的那组 `FOPT` 属性也是 PowerPoint 校验的一部分。

当时的结论是「生成形状树要从已验证的真实形状出发，只改 spid/文字/位置」，纯手写 `FOPT`
还差属性集。**这条已被 §9.21 推翻**：缺的不是属性集，而是 `FOPT` 末尾那 12 字节填充。

### 9.16 追加形状的三个记账点（缺一即拒收）

1. 外层 `0x03EE`（Slide）记录长度必须随内部 drawing 增长（`slide[4..8] = len - 8`）；
   只改 `0x040C`/`0xF002`/`0xF003` 会让顶层 `records()` 错位，PowerPoint 报文件损坏。
2. `0xF008`（FDG）payload 是 `shapeCount u32 | spidMax u32`（count 在前，不是 spidMax）；
   record instance = 簇号，`spidMax` 必须覆盖新 spid（inclusive）。
3. Document 的 `FDGG` 中该簇 `used = spidMax - (cluster << 10) + 1`；persist 目录里的 id 不能重复。
   `scratch5` 最初把 layout 和 slide 都写成 persist id 3，导致拒收；改成 layout=3 / slide=4 后
   `carry` 与 `minclone` 立刻通过。

另外：`void-ms` master + minimal drawing 但**不带版式**（`carryonly`，`masterIdRef=0x80000000`）
会卡在 PowerPoint 的修复对话框；带上版式（`carry`，`masterIdRef=0x8000000C`）则正常。
即形状在 slide 上，但 slide 仍必须指向一个有效版式。

### 9.17 `write_from_template` 曾走一条没人验证的路径（已合并）

`write::write_from_template` 与 `write::write_from_template_mixed` 原来用 `slide_record` + `text_shape`
从零拼幻灯片，只保留主母版、`masterIdRef = 0x80000002`。真 PowerPoint 对这两个函数生成的
`samples/scratch6-keep{true,false}.ppt` **全部拒收**，而同一个库里 `write_presentation`
（内嵌模板）与 `write_from_slide_template` 生成的 20 万字节文件却能打开 —— 因为它们走的是
`build_from_stream`：保留模板的全部版式对象，把**模板第一张幻灯片的形状树整体克隆**再换文字。

现在只保留 `write_from_slide_template`（同一份 `build_from_stream`），
`write_from_template`、`write_from_template_mixed`、`slide_record` 与其专用的三个常量已删除；`text_shape` 保留但注明
「单独使用会被 PowerPoint 拒收」，只供实验用。`crates/officedump/examples/scratch6.rs` 用三份
不同来源的模板（PowerPoint 的 `minimal.ppt`、WPS 的 `hello.ppt`、`void-ms.ppt`）复现：

| 模板 | 输出 | 真 PowerPoint |
| --- | --- | --- |
| `samples/ms/minimal.ppt`（PowerPoint） | 207872 B | OK，2 个形状 |
| `samples/hello.ppt`（WPS，2 页） | 49152 B | OK，2 个形状 |
| `samples/void-ms.ppt`（模板本身无形状） | 39424 B | OK，`slide 1:` 空 |

### 9.18 `TextSpecialInfoAtom` 的字符数：两家写法差 1

`hello.ppt` 一开始仍然打不开（`Presentations.Open` 抛 `HRESULT E_FAIL`），而把形状树原样透传
（`PPT_WRITE_PASSTHROUGH=1`）就能打开 —— 说明问题在换文字，不在文档组装。逐字节对比
（`crates/officedump/examples/sdiff.rs`，用法 `cargo run --example sdiff A.ppt B.ppt 3`）定位到
`0x0FAA`：`hello.ppt` 的标题是 13 个字符，`0x0FA0` 长度 26 B，但 `0x0FAA` 第一个 u32 写的是 **14**；
`0x0FA1` 的 cch 也是 14。也就是说 WPS 在 `0x0FAA` 里也把段落符算进去了，而 PowerPoint 写的
`minimal.ppt` 里该值等于文本原子长度。

修法是不假设约定，按模板自己的偏移平移：

```rust
let before = u32::from_le_bytes(info[0..4].try_into().unwrap());
let after = (before as i64 + new as i64 - old as i64).max(0) as u32;
info[0..4].copy_from_slice(&after.to_le_bytes());
```

改完后 `hello.ppt` 模板生成的文件被接受（1 页、2 页两种都对），`minimal.ppt` 路径不受影响
（它本来就 delta = 0）。

**`0x0FAA` 的首字段到底是不是 cch（用 `tbox` 工具逐框量）**：

| 文件 | 文本框 | 字符数 | 段落 | `0x0FA1` cch | `0x0FAA` 首字段 | `0x0FAA` 长度 |
|---|---|---|---|---|---|---|
| `ms/minimal.ppt` | 母版标题提示 | 12 | 1 | — | 13 | 14 |
| `ms/minimal.ppt` | 母版正文提示 | 24 | 5 | — | 25 | 14 |
| `ms/minimal.ppt` | 占位符 `*` | 1 | 1 | 2 | 2 | 14 |
| `hello.ppt` | 标题 `HELLO-PPT-001` | 13 | 1 | 14 | 14 | 12 |
| `hello.ppt` | 正文 `Slide one body text` | 19 | 1 | 20 | 20 | 12 |
| `hello.ppt` | 母版正文提示 | 28 | 5 | — | **12** | **140** |

单段落框里两家都写 `chars + 1`（含段落符），只有 WPS 那个多段落母版提示写 `12`、载荷长到 140 B。
**那个「语义未解」已解开**（§9.23）：`0x0FAA` 是一串 `TextSIRun`，`12` 只是第一个 run 的字符数，
140 B 是 10 个 run。写入器仍按模板偏移平移，读侧改成解析 run 后校验总和（§9.20、§9.23）。

### 9.20 `officedump anomalies`：读侧交叉校验文本计数

`office_ppt::text_anomalies(&stream) -> Vec<TextAnomaly>`（一趟记录树遍历，不解码文本），
每项记 `chars`、`text_atom` 偏移、`style_cch`、`special_cch`，`describe()` 出人话。判定：

* `0x0FA1` 的 `cch` 必须**等于** `chars + 1`（段落符计入）—— 全部样本一致，且
  `style_atoms_are_fully_decoded` 就是靠它精确解完 run 数组；
* `0x0FAA` 解成 `TextSIRun` 数组后，各 run 的 `count` 之和必须落在 `chars ..= chars + 1`
  （§9.23：PowerPoint 把段落符单独算一个 run，我们的写入器写 `chars` + `1`）；记录尾部有
  未初始化字节（WPS 常见）导致解不完时，不算异常。

命令行：`officedump anomalies FILE`（`main.rs` 里 `fn anomalies(cfb)`，`USAGE` 已加一行）。
45 份样本扫描结果：**只有四个被 PowerPoint 拒收的 `scratch5-*` 实验文件被点名**，其余全部
`no text anomalies` —— 其中 `scratch5-mintext.ppt` 报的正是当初的病因：

```
text atom @0xd348: 15 chars, style cch 15 (want 16), special-info cch 0 (at most 16)
```

（`text_shape` 合成的 `0x0FA1` 少算了段落符。）单测两条：
`text_anomalies_flag_stale_counts`（合成 `0xF00D` 容器，3 字符 + cch 3 → 命中，cch 4 → 干净）
和 `samples_have_no_unexpected_text_anomalies`（遍历 `samples/`、`samples/ms/`，除四个已知
拒收文件外必须零告警，且那四个必须各报 1 条）。

### 9.19 `write_from_slide_template` 的多页与图片：1/3/5/10 页都过

`build_from_stream` 把**模板第一张幻灯片的形状树**克隆给每一页，所以多页稿里各页的 `spid`
完全相同（`minimal.ppt` 模板下都是簇 2 的 `0x3400..0x3403`）。这点以前只验过 2 页，
`crates/officedump/examples/scratch9.rs` 把它铺开测：

| 输出 | 模板 | 页数 | 真 PowerPoint |
|---|---|---|---|
| `scratch9-1.ppt` 207872 B | `ms/minimal.ppt` | 1 | OK |
| `scratch9-3.ppt` 209408 B | 同上 | 3 | OK，3 页各 2 形状 |
| `scratch9-5.ppt` 211456 B | 同上 | 5 | OK |
| `scratch9-10.ppt` 216064 B | 同上 | 10 | OK |
| `scratch9-pics.ppt` 214528 B | 同上 | 3 | OK，图片 `type=13`，第 2 页两张（120×120、60×60） |
| `scratch9-hello4.ppt` 54784 B | `hello.ppt`（WPS，无版式） | 4 | OK |

用 COM 逐页读文本（`tools/tmp_text.ps1`）确认**每页文字是自己的**，不是克隆串页：

```
== samples/scratch9-3.ppt slides=3
  slide 1: Slide 1 / body line for slide 1
  slide 2: Slide 2 / body line for slide 2
  slide 3: Slide 3 / body line for slide 3
== samples/scratch9-pics.ppt slides=3
  slide 1: Picture 1 / caption 1
  slide 2: Picture 2 / caption 2
```

**结论：跨页复用同一批 `spid` 不需要重编号**（PowerPoint 按「每页独立」处理），
页数超过模板页数时新对象按 `next_id` 顺序追加，`SlideListWithText` 的 `slideId`
写成 `0x100 + i`。单测 `extra_slides_are_appended_with_their_own_text` 锁住这条：
5 页 → 5 个 slide 对象、每页两段自己的文字、slide 列表的 `slideId` 恰为 `0x100..0x104`。

### 9.21 从零拼的文本形状也能被接受：`FOPT` 必须恰好 102 字节

§9.15 里 `text_shape` 的合成形状一直被拒收，当时归因于「缺属性」。逐级替换才找到真凶：
用 `scratch5` 的 `morph`（`fn morph(real, spid, text, stage)`）拿**已被接受的真实形状**当底，
每级只换一块合成件：

| 变体 | 换了什么 | 真 PowerPoint |
|---|---|---|
| `m0` | 什么都不换，只改 spid（对照） | OK，3 形状 |
| `m1` | 换文字（`0FA8` 载荷 + `0FA1` 两个 `cch` + `0FAA` 首字段） | OK |
| `m2` | 换 `FOPT`：15 条属性、`len=90`（去掉尾部垃圾） | **ERR 文件损坏** |
| `m5` | `m2` + 12 个零字节，`len=102` | OK |
| `m7` | `m2` + 12 字节真垃圾，`len=102` | OK |
| `m6` / `m8` / `m9` | `len=96` / `100` / `108` | 均 ERR `HRESULT E_FAIL` |

**规则：幻灯片文本形状的 `FOPT` 必须恰好 102 字节**（15 条命名属性 90 B + 12 B 填充；
填充内容随意 —— 真垃圾和全零都行），少一字节或多一字节都被拒。真实文件里那 12 字节是
未初始化内存（两家形状都有 `07 68 98 98 20 00`，尾两字节是形状序号 `0x31`/`0x32`），
照抄没有意义，写零即可。

补齐后 `text_shape` 的其余四件事（都实测过）：

* `0FA1` 的 PF/CF 两个 `cch` 都写 `chars + 1`（§9.20 的 `anomalies` 早就报过 `want 16`）；
* ASCII 文字用 `0FA8` 字节原子（与 `clone_with_text` 一致），非 ASCII 才用 `0FA0`；
* `0FAA` 是 28 字节：`cch u32` + 24 字节真实布局（只写 12 个零字节仍被拒）；
* `F011` 里 `0x0BC3 = {0xFFFFFFFF, 0x0F}`、`0x041F = 1`；再补一条 12 字节的 `0FA6`。

结果：`samples/scratch5-text15.ppt` 与改用库里 `text_shape` 的 `samples/scratch5-mintext.ppt`
都被接受，`slide 1:` 是 `type=14 720x188 | type=14 720x130.4 | type=14 771.6x202.4`，
COM 读文本得到 `MIN-PPT-001 / minimal body / Generated shape` —— 第三个形状就是从零拼的。
单测 `write::text_shape_tests::{text_shape_counts_agree, text_shape_uses_the_wide_atom_for_non_ascii}`
锁住 102 字节、15 条属性、两个 `cch`、`0FAA` 的 28 字节与非 ASCII 的宽原子。

不过库的默认路径仍是克隆模板形状树（`clone_with_text`）：模板自带母版占位符与主题，
从零拼只能得到一个白底文本框，样式要自己一项项补。

### 9.22 theme / 版式 / 母版 ZIP 可以不写：PowerPoint 会自己重建

一直以来的「未解」是这些 ZIP：`0x0428`（1986 B，`tableStyles.xml`）、`0x0423`（1635 B，`drs/slideMasters/slideMaster1.xml`）、
`0x040E`（3448 B，`theme/theme/theme1.xml` + `themeManager.xml`）、每个版式一条 `0x041E`（1914 B，
`drs/slideLayouts/slideLayout1.xml`）、每个形状 XML 一条 `0xC3A9`。它们是 Office 2007 之后写下的
「回写缓存」（`drs/` = drawing record store），97-2003 的渲染并不读它们。

验证方式：`crates/officedump/examples/nodrs.rs` 把 persist 目录里的对象重新铺一遍，丢掉这些记录并修正
容器长度、重写 persist 目录与 `UserEditAtom`、改 `Current User` 里的 `offsetToCurrentEdit`：

```
cargo run -q --example nodrs samples/scratch5-minclone.ppt samples/scratch5-nodrs.ppt
# 54426 → 22581 字节（丢掉 31845 字节），真 PowerPoint：OK，3 形状
KEEP_DRS=1 cargo run -q --example nodrs ...   # 对照组：只重排不删，同样 OK
```

对照组很关键：第一次跑 `nodrs` 时被拒收，以为是 ZIP 必需，其实是重排本身有两个 bug ——
`UserEditAtom` 的 `offsetLastEdit` 在**偏移 8**（不是 4，偏移 4 是 `version`/`minor`/`major`），
`Current User` 的 `offsetToCurrentEdit` 在**偏移 16**（前面有 8 字节记录头和 size、headerToken）。
先让「不删只重排」通过，才能确认「删掉」是被接受的。

于是 `write.rs` 里加了 `ROUND_TRIP_STORE: [u16; 5] = [0x040E, 0x0423, 0x041E, 0x0428, 0xC3A9]`
与 `strip_round_trip()`，在 `build_from_stream` 的三处生效：拷贝过来的母版/版式对象、幻灯片除
`SlideAtom`/`PPDrawing` 之外的尾部记录、`Document` 的样板记录。效果（真 PowerPoint 全部 OK）：

| 输出 | 之前 | 之后 |
|---|---|---|
| `scratch9-1.ppt` | 207872 B | 125440 B |
| `scratch9-hello4.ppt` | 54784 B | 26112 B |
| `scratch6-void-ms.ppt` | 39424 B | 14336 B |

代价：这些 XML 是「新版 Office 存回 pptx 时的往返保真」用的，丢掉后 PowerPoint 会按二进制记录重新
生成一份默认主题 —— 视觉不变（颜色/字体仍在母版的 `ColorScheme`/`TextMasterStyle` 里），但如果模板带
自定义主题，回写 pptx 时拿不到原来的 XML。对 97-2003 写入器来说这是值得的：文件小 40%，也不再依赖
模板里那 30 KB 的 ZIP 字节。

单测：`write::round_trip_store_tests::{generated_decks_carry_no_round_trip_store, strip_round_trip_drops_only_the_store}`
（前者遍历生成文件的所有容器断言没有这五种记录，并断言模板里确实有，避免空检查）。

### 9.23 `TextSpecialInfoAtom` 是一串 `TextSIRun`：多段落语义

MS-PPT §2.9.54/§2.9.55：`0x0FAA` 的载荷**没有总长度字段**，就是一串 `TextSIRun` 紧排到记录末尾。

```
TextSIRun       = count (4 B, 必须 ≥ 1) + si (TextSIException, 变长)
TextSIException = mask (4 B 位域) + 按位出现的可选字段
    bit 0 spell     → spellInfo (2 B)
    bit 1 lang      → lid       (2 B, TxLCID：0x0804 简中 / 0x0409 en-US)
    bit 2 altLang   → altLid    (2 B)
    bit 5 fPp10ext  → pp10runid(4 bit) + reserved3(27 bit) + grammarError(1 bit) = 4 B
    bit 6 fBidi     → bidi      (2 B)
    bit 9 smartTag  → smartTags (变长，本库不解)
    其余位保留，必须为 0
```

实测（`cargo run -q --example tbox FILE` 现在会打印 run 计数）：

| 文件 | 文本框 | 字符 | 载荷 | run 的 `count` | mask | lid / altLid |
|---|---|---|---|---|---|---|
| `hello.ppt`（WPS） | 标题 | 13 | 12 B | `14` | 6 | 0409 / 0409 |
| `ms/minimal.ppt` | 母版标题提示 | 12 | 14 B | `13` | 7 | 0804 / 0409 |
| `ms/minimal.ppt` | 幻灯片标题 | 11 | 28 B | `11+1` | 7, 7 | 0804 / 0409 |
| `hello.ppt`（WPS） | 母版正文提示（5 段） | 28 | 140 B | `12+1+3+1+3+1+3+1+3+1` | 7 × 10 | 0804 / 0409 |

所以 §9.18 里那个「WPS 写 12」就是**第一段（含段落符）的字符数**，140 B 是 10 个 run —— 一段一个 run。
PowerPoint 反过来把同一段文字拆成 `11+1`（正文 11 字 + 段落符各一个 run），我们的写入器写
`chars` + `1` 两个 run（与 `minimal.ppt` 的幻灯片标题同形）。三种划分都被 PowerPoint 接受：run 只是
语言/拼写信息的分组，唯一约束是 `Σcount` 落在 `chars ..= chars + 1`。

另一个坑：WPS 有些 `0x0FAA` 在最后一个 run 之后留着未初始化字节（例如 `01 00 00 00 07 00 00 00
00 00 04 81 00 00 00 82 00 00 00 83 ...`，lid 会解出 `0x8104` 这种不存在的语言码）。所以读侧不能把
「解不完」当成损坏：`office_ppt::special_info_runs(payload) -> Option<Vec<SpecialInfoRun>>`
（`SpecialInfoRun { count, mask, lid, alt_lid }`）解到长度不够就返回 `None`，`text_anomalies` 只在
解成功时校验总和（`TextAnomaly::special_total`）。

单测：`special_info_tests::{measured_payloads_decode_into_runs, run_counts_cover_the_text,
junk_after_the_last_run_is_reported_as_none}`。
## 9.24 `write_presentation`：从零生成幻灯片，不再克隆形状树

`write::write_presentation(&[SlideSpec])` 的每一页都是**新拼的** `Slide`：`SlideAtom` + 一个
`PPDrawing`，里面是 `OfficeArtDgContainer`（`F008` 记账 + `F003` 组容器），组容器第一个成员是
组形状（`F009` 16 个零 + `F00A` 的 `spid = cluster << 10`、flags `0x05`），后面才是
`text_shape` / `picture_shape` 生成的形状。模板只贡献 Document 的骨架（`Environment`、样式、
theme）、主母版和一个版式。

拼一页需要四件事，缺任何一条 PowerPoint 都报「文件有问题」：

1. **`SlideAtom`**：`geom = 0x10`、占位符全 0、`masterIdRef` 指向**版式**（`0x8000_000C`）、
   `flags = 7`，并且要自己包一层 `0x03EF`（`slide_atom_full` 返回的是 24 字节载荷）。
2. **`F008`（`OfficeArtFDG`）**：载荷 = `SpContainer` 个数（含组形状）| 最大 spid，记录的
   `instance` 是簇号。
3. **`0x07F0`（`SlideShowSlideInfoAtom`）**：真幻灯片在 `PPDrawing` 之后都跟着这条 32 字节记录
   （`void-ms.ppt` 之后还有 `0x0422`，`minimal.ppt` 是 `0x1388` + `0x041D`）。少了它，即使
   `SlideAtom` 和绘图完全正确，`Presentations.Open` 也直接抛「Office 检测到文件有问题」——
   这是整条路上最后一个坑。
4. **FDGG 记账**（文档级 `0x040B → F000 → F006`）：第 `i` 个簇的 `used` 必须等于
   `spidMax - (i << 10) + 1`，也就是「形状数 + 1」（组形状自己占一个 id）。

有图片时还多一条：**`Pictures` 流里的每个 BLIP 必须在文档级绘图组里有对应的
`OfficeArtFBSE`**（`0x040B → F000 → F001`，`instance = 1`，每条 FBSE 44 字节，`cRef` 是引用它的
形状数）。只写 `Pictures` 流而忘了 BStore，即使流内容和图片形状都合法，PowerPoint 同样拒收——
`write_presentation` 的生成路径一开始就栽在这里（对照组 `scratch10-streamonly.ppt` 证明了这一点）。

验证（真 PowerPoint 2016，`tools/check_ppt.ps1`）：

| 文件 | 内容 | 体积 | 结果 |
|---|---|---|---|
| `scratch10-text.ppt` | 2 页，标题 + 3 段正文 | 28160 B | OK，5 个 `type=14` 形状，文字逐页正确 |
| `scratch10-pic.ppt` | 1 页，标题 + 正文 + 120×120 图片 | 31744 B | OK，`type=14 | type=14 | type=13 120x120` |
| `scratch10-piconly.ppt` | 1 页，只有图片 | 30720 B | OK，`type=13 120x120` |

文字用 `tools/tmp_text.ps1` 读回：`slide 1: Generated slide one / the shape tree is built, not
cloned / second paragraph`、`slide 2: Generated slide two / two slides, four shapes`。

公开 API（`crates/office-ppt/src/write.rs`）：

| 函数 | 用途 |
|---|---|
| `slide_record(cluster, master, flags, &shapes)` | 拼一页 `Slide`（含 `F008` 与 `0x07F0`） |
| `drawing_group_with(per_object, &[(blip_len, cRef)])` | 文档级绘图组，带 `BStoreContainer` |
| `document_with(per_object, blips, master_atoms, slide_atoms, style, theme, atom)` | Document，可列多个母版/版式 |
| `write_presentation(&[SlideSpec])` | 整份演示文稿（内嵌模板 + 生成的形状树） |
| `write_from_slide_template(&[u8], &[SlideSpec])` | 用调用方的模板；仍克隆模板第一页的形状树 |

`write_presentation` 还会对模板的母版/版式跑 `strip_round_trip`（§9.22），所以 2 页文字稿只有
28160 B，而克隆路径的 `samples/written.ppt` 是 126464 B。正文段落和克隆路径一样放进同一个占位符、
用段落符分隔，所以两条路径读回来的文字完全一致。

`write_from_slide_template` 保留克隆行为：外来模板的版式数量、形状树都可能和我们的生成器不一致，
克隆是最保守的做法（§9.19）。内嵌模板则不再需要「模板里那张幻灯片」——只用到 Document 骨架、
主母版和一个版式。

单测：`write::generated_tests::{generated_decks_parse_back,
generated_slides_have_their_own_clusters, pictures_get_a_store_entry}`；
`write::tests::{writes_a_deck_the_reader_accepts, writes_a_presentation_with_a_picture,
picture_size_comes_from_the_client_anchor}` 继续覆盖生成路径，
`extra_slides_are_appended_with_their_own_text` 覆盖克隆路径。
