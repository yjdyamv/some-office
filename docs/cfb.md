# CFB：Word / Excel / PowerPoint 共同的容器格式

`.doc`、`.xls`、`.ppt`（97-2003 二进制格式）在外层都是同一个东西：
**Microsoft Compound File Binary**，也就是常说的 OLE2 复合文档。所以逆向这三种格式
的第一步都是同一个：先把 CFB 拆开，拿到里面的"流"（stream）。

- 规范：MS-CFB *Compound File Binary File Format*
- 本仓库实现：`crates/office-cfb/src/lib.rs`
- 验证样本：`samples/hello.doc` / `.xls` / `.ppt`（WPS 生成）、
  `PROTTPLN.DOC`、`SOLVSAMP.XLS`、`PROTTPLN.PPT`（Microsoft 生成）

```
+---------------------------------------------------+
| Header 512 B (magic D0 CF 11 E0 A1 B1 1A E1)      |
|   sector size, FAT sector list (DIFAT)            |
+---------------------------------------------------+
| Sector 0 | Sector 1 | Sector 2 | ...               |  <- 每个 512 B（v3）或 4096 B（v4）
+---------------------------------------------------+
   其中某些扇区被标记为: FAT / 目录 / MiniFAT / 数据
```

## 1. 头部（512 字节）

`samples/hello.doc` 的真实字节：

```
0000  d0 cf 11 e0 a1 b1 1a e1 00 00 00 00 00 00 00 00  ................
0010  00 00 00 00 00 00 00 00 3e 00 03 00 fe ff 09 00  ........>.......
0020  06 00 00 00 00 00 00 00 00 00 00 00 01 00 00 00  ................
0030  01 00 00 00 00 00 00 00 00 10 00 00 02 00 00 00  ................
0040  01 00 00 00 fe ff ff ff 00 00 00 00 00 00 00 00  ................
0050  ff ff ff ff ff ff ff ff ...                       ................
```

| 偏移 | 大小 | 字段 | hello.doc | 说明 |
|---|---|---|---|---|
| 0x00 | 8 | signature | `D0CF11E0A1B11AE1` | 固定魔数 |
| 0x18 | 2 | minorVersion | 0x003E | 无关紧要 |
| 0x1A | 2 | majorVersion | 3 | 3 = 512B 扇区，4 = 4096B |
| 0x1C | 2 | byteOrder | 0xFFFE | 固定 |
| 0x1E | 2 | sectorShift | 9 | `sector = 1 << shift` |
| 0x20 | 2 | miniSectorShift | 6 | `mini sector = 64 B` |
| 0x2C | 4 | numFatSectors | 1 | FAT 占几个扇区 |
| 0x30 | 4 | firstDirSector | 1 | 目录首扇区 |
| 0x38 | 4 | miniStreamCutoff | 0x1000 | 4096：小于它的流进 mini stream |
| 0x3C | 4 | firstMiniFatSector | 2 | |
| 0x40 | 4 | numMiniFatSectors | 1 | |
| 0x44 | 4 | firstDifatSector | 0xFFFFFFFE | 无扩展 DIFAT |
| 0x48 | 4 | numDifatSectors | 0 | |
| 0x4C | 436 | difat[109] | `01 00 00 00`, `FE FF FF FF`, ... | FAT 扇区号列表 |

## 2. 扇区寻址（最容易写错的地方）

**扇区号是"逻辑"编号，第 0 个扇区从文件的 512 字节处开始**，因为头部自己占了第一个
"扇区位置"：

```rust
fn sector_offset(sector: u32, sector_size: usize) -> usize {
    (sector as usize + 1) * sector_size
}
```

写成 `sector * sector_size` 会让所有数据整体偏移 512 字节。本仓库第一版就踩了这个坑：
读出来的目录名全是乱码，`WordDocument` 流的首字节不是 `EC A5`。

## 3. FAT：扇区链表

FAT 本身是一个 `u32` 数组，`fat[s]` 给出扇区 `s` 的后继扇区号，特殊值：

| 值 | 含义 |
|---|---|
| `0xFFFFFFFD` | FATSECT：该扇区被 FAT 自己占用 |
| `0xFFFFFFFE` | ENDOFCHAIN：链尾 |
| `0xFFFFFFFF` | FREESECT：空闲 |
| `0xFFFFFFFC` | DIFSECT：该扇区属于 DIFAT |

`hello.doc` 的 FAT 前 6 项：`FD FF FF FF | 05 | FE FF FF FF | 04 | 06 | FE FF FF FF`
——扇区 0 是 FAT 自身，扇区 1 之后是目录（`1 -> 5`），扇区 3 之后是 mini stream
（`3 -> 4 -> 6 -> END`）。

**扇区号 ≥ 0xFFFFFFFA 一律当链尾**，不要试图去查表（数组越界）。

### DIFAT：FAT 扇区表本身

头部只有 109 个槽位，FAT 更大时，DIFAT 以扇区链表形式续接：每个 DIFAT 扇区末尾 4 字节
指向下一个 DIFAT 扇区，前面 `sector_size/4 - 1` 个槽位是 FAT 扇区号。

## 4. Mini stream：小流的二级分配

小于 `miniStreamCutoff`（通常 4096）的流**不占普通扇区**，而是全部塞进 Root Entry 的
那条流里（"mini stream"），再按 64 字节的 mini sector 切分，用 miniFAT 串起来。

- 目录项里的 `startSector` 此时是 **mini sector 号**，不是普通扇区号；
- 读取时：`mini_stream[mini_sector * 64 .. +64]`，后继查 miniFAT；
- Root Entry 的 `startSector` 才是 mini stream 在普通 FAT 里的起点。

`hello.doc`：`WordDocument` 3630 B < 4096，所以它在 mini stream 里（mini sector 16 开始），
而 `0Table` 2448 B 同样在 mini stream 里（mini sector 73）。Root Entry 的 mini stream
本身 7168 B，占普通扇区 3 → 4 → 6。

## 5. 目录：128 字节条目 + 红黑树

| 偏移 | 大小 | 字段 |
|---|---|---|
| 0x00 | 64 | 名字（UTF-16LE，含结尾 NUL） |
| 0x40 | 2 | 名字字节数（含 NUL） |
| 0x42 | 1 | 类型：0 空 / 1 storage / 2 stream / 5 root |
| 0x43 | 1 | 颜色（红黑树） |
| 0x44 | 4 | left sibling id |
| 0x48 | 4 | right sibling id |
| 0x4C | 4 | child id（storage 的第一个孩子） |
| 0x50 | 16 | CLSID |
| 0x74 | 4 | startSector |
| 0x78 | 8 | size（v3 只用低 4 字节） |

`hello.doc` 的目录（id / 类型 / 名字 / 起始 / 大小）：

```
[0] Root Entry                      type=5 start=3   size=7168
[1] \x05SummaryInformation           type=2 start=0   size=332
[2] \x05DocumentSummaryInformation   type=2 start=6   size=464
[3] WordDocument                    type=2 start=16  size=3630
[4] 0Table                          type=2 start=73  size=2448
[5] Data                            type=2 start=ENDOFCHAIN size=0
[6] WpsCustomData                   type=2 start=14  size=68
```

条目 id 就是数组下标，所以 `startSector` 之外还要靠 id 走红黑树：storage 的 `child`
是树根，然后 `left`/`right` 中序遍历得到该 storage 下的所有条目。本仓库的
`Cfb::children()` / `Cfb::paths()` 就是中序遍历。

## 6. 三种文档用到的流

| 格式 | 关键流 | 说明 |
|---|---|---|
| .doc | `WordDocument` + `0Table`/`1Table` | FIB 在 WordDocument 开头，piece table 在 table 流 |
| .xls | `Workbook`（老版本叫 `Book`） | 整个 BIFF 记录流 |
| .ppt | `PowerPoint Document` + `Current User` | 记录树 + 当前编辑指针 |

另外都有 `\x05SummaryInformation`（属性集，标题/作者/页数等），本仓库暂不解析。

## 7. 代码映射

```rust
use office_cfb::{Cfb, EntryKind};

let cfb = Cfb::open("samples/hello.doc")?;
println!("{} streams", cfb.entries().iter().filter(|e| e.kind == EntryKind::Stream).count());
let wd = cfb.stream("WordDocument")?;      // Vec<u8>，已按 FAT/miniFAT 拼好
let table = cfb.stream("0Table")?;
```

- `Cfb::parse(Vec<u8>)` 一次性读入所有扇区，之后 `stream()` 只是走链 + 拷贝；
- 错误类型 `office_cfb::Error` 区分 `BadSignature` / `BadVersion` / `Truncated` /
  `CycleDetected`（防 FAT 环）/ `NoSuchStream`。

## 8. 已知坑清单

1. **扇区偏移要 +1**（见 §2）。
2. **mini stream 判定用 `<` cutoff**，等于 4096 的流仍在普通扇区里。
3. **空流**的 `startSector` 是 `ENDOFCHAIN`（`hello.doc` 的 `Data`），别去读。
4. v4（4096 字节扇区）文档的目录项仍是 128 字节，但扇区里能放 32 个。
5. FAT 数组可能比"实际扇区数"长（尾部是 FREESECT），越界前先判 `s < fat.len()`。
6. 文件名带前导 `\x05` 的是"属性集"流（SummaryInformation），不是普通文本。
