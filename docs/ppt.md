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
