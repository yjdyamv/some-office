# some-office：Office 97-2003 二进制格式的逆向笔记 + Rust 实现

微软没有公开 `.doc` / `.xls` / `.ppt` 的完整二进制格式（只有 [MS-CFB]、[MS-DOC]、[MS-XLS]、
[MS-PPT] 这些"事后公开"的规范，且规范与真实文件经常对不上）。这个仓库做三件事：

1. **逆向**：用本地 Office 安装（WINWORD.EXE / EXCEL.EXE / POWERPNT.EXE / wwlib.dll）
   配合 Ghidra（通过 ghlab MCP）验证格式细节，记录在 `docs/`；
2. **实现**：用 Rust 从零写读取器（不依赖任何第三方 crate，`cargo build` 无网络也能过）；
3. **验证**：用 WPS 生成的样本 + Microsoft 官方模板双向对拍；再把自己写出的文件
   交给真的 Word / Excel 打开（`cargo run --example mkfiles` + `tools/roundtrip.ps1`）。

```
docs/            逆向文档（每个格式一篇，含真实字节、偏移、踩坑记录）
  evidence.md    二进制证据：MSO.DLL / wwlib.dll 反汇编，规范与实测的差异
  cfb.md         Compound File Binary：三种格式共同的容器
  doc.md         .doc：FIB + piece table + CHPX/PAPX 格式 + 文本故事
  xls.md         .xls：BIFF8 记录流 + SST + RK + 公式反编译 + 数字格式
  ppt.md         .ppt：记录树 + persist directory + Escher + 图片
  encryption.md  加密：RC4（旧式 / CryptoAPI）+ XOR 混淆，三种格式的加密边界
crates/
  office-cfb/    容器层：FAT / miniFAT / 目录 / 流
  office-enc/    编码层：CP1252 / UTF-16LE
  office-crypto/ 加密层：MD5 / SHA1 / RC4 / XOR 混淆（零依赖）
  office-blip/   OfficeArt BLIP 提取（.ppt 与 .doc 共用）
  office-doc/    .doc 读取器 + 写入器（write.rs：字符/段落格式、格式页、多页）
  office-xls/    .xls 读取器 + 最小写入器（write.rs）
  office-ppt/    .ppt 读取器 + 写入器（write.rs：克隆内嵌模板的形状树并替换文字）
  officedump/    命令行工具（info / text / tree / raw / html / media / stories / styles / --password）
samples/         自造样本（WPS 通过 COM 生成，仍是合法的 .doc/.xls/.ppt）
samples/ms/      微软 Office 自己写的参考样本（见 tools/ms_samples.ps1）
tools/           cfb_list.py（Python 原型）、make_samples*.ps1（WPS 样本生成）
                 ms_samples.ps1（临时改写 HKCU 的 CLSID，让 COM 落到微软 Office）
                 scan_consts.py / disasm.py（静态逆向：常量扫描、反汇编）
                 roundtrip.ps1（把我们自己写的文件交给真 Word/Excel/PowerPoint 打开）
                 extract_ppt_template.py（从微软空白稿提取内嵌模板）
```

## 快速开始

```console
$ cargo test                        # 136 个测试（单元 + 真实文件集成 + 随机往返）
$ cargo run -q -p officedump -- info samples/hello.doc
$ cargo run -q -p officedump -- text samples/hello.doc
$ cargo run -q -p officedump -- text samples/hello.xls
$ cargo run -q -p officedump -- text samples/hello.ppt
$ cargo run -q -p officedump -- tree samples/hello.ppt      # 记录树
$ cargo run -q -p officedump -- raw  samples/hello.xls      # 主流前 128 字节

# 新增能力
$ cargo run -q -p officedump -- html samples/formatted.doc  # 字符/段落格式 → HTML
$ cargo run -q -p officedump -- stories samples/stories.doc # 主文档/脚注/页眉分开
$ cargo run -q -p officedump -- text samples/formulas.xls   # 公式反编译 + 数字格式
$ cargo run -q -p officedump -- media samples/image.ppt out # 导出图片（.ppt）
$ cargo run -q -p officedump -- media samples/picture.doc out # 导出图片（.doc）
$ cargo run -q -p officedump -- styles samples/formatted.doc # 样式表：istd → 名字 + 样式级格式
$ cargo run -q -p officedump -- text samples/notes.ppt       # 幻灯片 + 演讲者备注
$ cargo run -q -p officedump -- text samples/longstr.xls     # SST CONTINUE（9000 字符单元格）
$ cargo run -q -p officedump -- info samples/ms/chart.xls    # 图表：类型/标题/轴标题/系列

# 自己写文件，再交给真 Office 打开
$ cargo run -q --example mkfiles    # 从零生成 samples/written.doc / written.xls
$ pwsh tools/roundtrip.ps1          # Word 读出 3 段、Excel 读出 A1:A4

# 加密文件（密码都是 secret123）
$ cargo run -q -p officedump -- info samples/encrypted.doc             # → protected
$ cargo run -q -p officedump -- text samples/encrypted.doc -p secret123
$ cargo run -q -p officedump -- text samples/encrypted.xls -p secret123
$ cargo run -q -p officedump -- text samples/encrypted.ppt -p secret123
```

作为库使用（三个解析器的入参都是**整份文件的字节**，内部自己拆 CFB）：

```rust
use office_cfb::Cfb;

fn read_any(path: &str) -> Result<String, Box<dyn std::error::Error>> {
    let bytes = std::fs::read(path)?;
    let cfb = Cfb::parse(bytes.clone())?;
    if cfb.stream("WordDocument").is_ok() {
        Ok(office_doc::parse(&bytes)?.plain_text())
    } else if cfb.stream("Workbook").is_ok() {
        Ok(office_xls::parse(&bytes)?.plain_text())
    } else {
        let doc = cfb.stream("PowerPoint Document")?;
        let user = cfb.stream("Current User").ok();
        Ok(office_ppt::parse_stream(&doc, user.as_deref())?.plain_text())
    }
}
```

## 三种格式的关键差异（一张表）

| | .doc | .xls | .ppt |
|---|---|---|---|
| 主数据流 | `WordDocument` | `Workbook` | `PowerPoint Document` |
| 辅助流 | `0Table` / `1Table` | — | `Current User` |
| 结构 | 固定 FIB + piece table | 扁平记录流 | 记录树 + persist 目录 |
| 文本定位 | CLX → PlcPcd → fc | SST 字符串池 + isst | Escher 容器里的 TextCharsAtom |
| 文本编码 | UTF-16 或 CP1252（按 piece） | UTF-16 或 8-bit（按 grbit） | UTF-16（TextCharsAtom）或 8-bit |
| 最坑的地方 | fc 的 bit30 与 `/2` | SST 两个计数器是 **4 字节** | 必须走 Current User 才能拿到当前页 |
| 增量保存 | 无（整份重写） | 无 | **有**，所以有 persist 目录 |
| 格式 | CHPX/PAPX 格式页 + sprm | XF + FORMAT 格式串 | Escher 属性表 |
| 公式 | 域代码（未解析） | RPN token 流（已反编译） | 无 |
| 图表 | 无此概念 | Chart Sheet Substream（已解析类型/标题） | 无此概念 |
| 图片 | `Data` 流（sprmCPicLocation） | 未提取 | `Pictures` 流 + pib |

## 逆向方法

### 1. 先写 Python 原型，再做 Rust

`tools/cfb_list.py` 是最早的原型：把 CFB 拆开、打印目录、dump 任意流的 hex。
所有"真实字节"都来自它，文档里的 hexdump 可以复现：

```console
$ uv run --no-project python tools/cfb_list.py samples/hello.doc
```

### 2. 样本：两种来源，互相印证

- **WPS 生成**（`tools/make_samples.ps1`）：本机 `Word.Application` 等 ProgID 被 WPS
  劫持（CLSID `000209FF-...` → `wps.exe`），所以 COM 生成的其实是 WPS 写的文件。
  好处是它们足够"脏"（非标准 cbRgFcLcb、ASCII 也用 UTF-16、标题页用居中版式），
  正好用来发现规范之外的写法。
- **Microsoft 官方文件**（不需要生成）：
  `C:\Program Files\Microsoft Office\root\Office16\2052\PROTTPLN.DOC`、`PROTTPLV.DOC`、
  `PROTTPLN.XLS`、`PROTTPLN.PPT`、`...\SAMPLES\SOLVSAMP.XLS`（6 个工作表、159 个字符串、
  含真实公式缓存值）。同一个解析器必须两种都能读，否则就是"只对 WPS 有效"。

### 3. Ghidra：规范说不清的地方去问二进制

规范与实测冲突时（例如 persist 目录的 count 语义、SST 计数器宽度），用 Ghidra 反编译
确认常量与分发逻辑，细节见 [`docs/evidence.md`](evidence.md)。

关键结论：`POWERPNT.EXE` 只是壳（.text 只有 7 KB，整个文件里搜不到 `0xE391C05F`），
真正的 PPT 引擎在 `MSO.DLL`；Word 的引擎是 `wwlib.dll`；Excel 的引擎在 `EXCEL.EXE` +
`excelcnv.exe`。所以逆向要找对模块，否则会在壳里空转。

`tools/scan_consts.py` 和 `tools/disasm.py`（capstone）是不依赖 Ghidra 的兜底手段：
前者按 32 位常量/字符串扫 PE 并给出 VA，后者按 VA 反汇编，用来读 `CurrentUserAtom`
校验函数、记录类型分发表这类小片段。

## 已打通的能力

| 能力 | 状态 |
|---|---|
| 文本（三种格式） | ✅ 含中日韩 UTF-16、CP1252、8-bit piece |
| 公式反编译 | ✅ .xls：18 个公式与 WPS 原文逐字一致，SOLVSAMP.XLS 无截断 |
| 数字格式 / 日期 | ✅ 内置 40+ 格式、自定义 FORMAT、1900/1904 序列号、幽灵闰日 |
| `.xls` 图表 | ✅ 类型（折线/饼/圆环/条形/柱形/散点/面积，由**出现了哪条记录**判定）、图表标题、轴标题、系列名；嵌入图表与图表工作表都能找到 |
| `.xls` BIFF5 | ✅ 真 BIFF5 `0x0500`（微软写的 `samples/ms/excel95.xls`），`ptgRef` 的行字带相对标志、列只有 1 字节 |
| 字符格式 | ✅ .doc CHPX：粗/斜/下划线/删除线/字号/颜色/上下标 |
| 段落格式 | ✅ .doc PAPX：对齐/缩进/间距/列表级别（若写进 PAPX） |
| 样式表 | ✅ .doc STSH：`istd` → 样式名（`正文`/`Normal`/`Table Grid`…），微软与 WPS 文件都通过 |
| 样式级格式 | ✅ .doc `STD.grupe` 的 `UPX` sprm（`正文` 的对齐/字号、`Balloon Text` 的 8pt），与直接格式正确分层 |
| 表格 | ✅ .doc：`0x07` 单元格标记 + `sprmPFInTable`/`sprmPFTtp` + `TDefTable` 列宽，`html` 输出 `<table>`/`<colgroup>` |
| 文本故事 | ✅ .doc：主文档 / 脚注 / 页眉页脚 / 批注 / 尾注 / 文本框 |
| 演讲者备注 | ✅ .ppt：`Notes` 持久对象（`SlideListWithText` instance 2 定序） |
| .ppt 文字格式 | ✅ `StyleTextPropAtom`：粗/斜/下划线/字号/颜色/对齐，`html` 输出内联样式；微软样本与 WPS 样本都对拍通过 |
| .ppt 占位符继承 | ✅ 主母版 `TextMasterStyleAtom`（`recInstance` 定类型、按缩进级别查表），run 未写的属性从母版补（`你好.ppt` 标题 44pt、正文 28/24/20/18 阶梯） |
| 图片 | ✅ 读：.ppt `Pictures` 流、.doc `Data` 流（sprmCPicLocation）→ PNG/JPEG/GIF/BMP/TIFF/EMF/WMF 原样导出；写：.doc 内联图片、.ppt 图片形状（Word/PowerPoint 已验证） |
| HTML 输出 | ✅ 三种格式都带样式（doc 富文本、xls 表格、ppt 分页） |
| 页序 / 工作表序 | ✅ .ppt 走 persist 目录，.xls 走 BOUNDSHEET |
| 写文件 | ✅ `office_cfb::write`（容器，目录树按规范排序）、`office_doc::write_document`（.doc：粗/斜/下划线/字号/颜色/对齐/缩进 + 格式页 + 多页，Word 已验证）、`office_xls::write_minimal`（.xls，Excel 已验证）、`office_ppt::write_presentation`（.ppt：**自包含**，模板编译进 crate，支持文字与图片，PowerPoint 已验证） |
| 加密 | ✅ 旧式 RC4 v1.1 / RC4 CryptoAPI 128 位（三种格式）、XOR 混淆（.xls 逐记录解密；.doc 的 `fObfuscated` 走 `lKey` 校验 + 整体去混淆） |

## 测试策略

三层，缺一不可：

| 层 | 位置 | 证明什么 |
|---|---|---|
| 单元 | 各 crate 的 `mod tests` | 单个解析器的字节级行为（RK 编码、sprm 长度、BLIP magic…） |
| 真实文件集成 | `crates/officedump/tests/formats.rs` | 微软与 WPS 写出的真实文件能读对（`samples/`、`samples/ms/`） |
| 随机往返 | `crates/officedump/tests/roundtrip_props.rs` | 写入器的输出能被自己的读取器还原：CFB 流集合、`.xls` 单元格、公式 `compile`↔`decompile`、`.doc` 格式/表格、`.ppt` 多页文字 |

随机测试用自带的 xorshift64* 生成数据（无第三方依赖），失败信息里带 seed，
因此可复现。最后一道关是**真实程序**：`tools/roundtrip.ps1` 让 Word、Excel、
PowerPoint 打开写入器产出的文件并回读内容 —— 只有它们能证明字节"合法"而不只是
"自洽"。

证据链见 `docs/evidence.md`（规范与实测的对照表、Ghidra 反编译片段、复现命令）。

## 已知限制

| 限制 | 说明 |
|---|---|
| `.ppt` 写入器内嵌了一份 203 KB 模板 | 自己合成 Escher 形状树会被 PowerPoint 拒收（它连「完全自洽的子集」都拒），所以 `write_presentation` 用 `crates/office-ppt/templates/minimal.bin`（微软空白稿的 `PowerPoint Document` 流）克隆形状树再换文字；模板不能裁剪，见 ppt.md §9.4 |
| 写入器不支持的细节 | `.doc` 不写样式表名（表格、图片可以写：`TDefTable` / `sprmCPicLocation`+`Data` 流，都经 Word 验证）；`.xls` 不写图表/逐格数字格式（公式可以写：`compile` 把文本编成 `rgce`，Excel 会重算）；`.ppt` 不写动画/备注；图片尺寸可控（`ClientAnchor` = `top,left,right,bottom`，1/8 pt，见 ppt.md §9.5） |
| XOR 只测了合成流 | ~~已关闭~~：`samples/ms/xor.xls` 是微软 Excel 95 带修改密码保存的真文件（4 字节 `FILEPASS`），往返测试仍保留作无文件回归 |
| `.doc` 的 XOR 混淆没有微软样本 | 密钥派生由真 `.xls` 钉死、边界由自造的 `samples/obfuscated.doc` 验证（§5.2）；Word 2016 已不再写这种保护 |
| 图表只读「是什么」 | .xls 图表能给出类型/标题/轴标题/系列名，但不读数据引用、图例、坐标轴刻度、数据标签 |
| 无公式求值 | 只还原公式文本，不计算（缓存值来自文件） |
| .ppt 备注页不继承母版样式 | 备注页的样式在 NotesMaster 里，只做主母版的继承 |
| .ppt 母版第 5 级可能缺属性 | 部分 `0x0FA3` 尾部字段宽度与规范不符，最后一级解不出时留空而不猜 |
| BIFF5 只测了合成流 | ~~已关闭~~：`samples/ms/excel95.xls` 是真正的 BIFF5（`0x0500`），并据此发现 BIFF5 的 `ptgRef` 编码与 BIFF8 不同（§12.1） |

## 参考规范

- [MS-CFB] Compound File Binary File Format
- [MS-DOC] Word (.doc) Binary File Format
- [MS-XLS] Excel Binary File Format (.xls)
- [MS-PPT] PowerPoint (.ppt) Binary File Format
- [MS-ODRAW] Office Drawing (Escher) Binary File Format

规范是"事后文档"，与真实文件有出入的地方在 `docs/` 里都标了「实测」与「坑」。
