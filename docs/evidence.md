# 逆向证据：规范说不清的地方，去二进制里找

`docs/cfb.md` / `doc.md` / `xls.md` / `ppt.md` 里凡标注「实测」的结论，来源有三类：

1. **文件字节**：用 `tools/cfb_list.py` 把真实文件的字节 dump 出来（最可靠，可复现）；
2. **静态扫描**：`tools/scan_consts.py` 在 Office 二进制里找常量/字符串；
3. **反汇编**：`tools/disasm.py`（capstone）按虚拟地址反汇编；Ghidra（`ghlab`）用来做
   跨函数的调用关系与反编译。

本文记录第 2、3 类证据，以及它们如何修正了规范/第三方实现里的说法。

## 0. 先搞清楚代码在哪

一个容易踩的坑：**`POWERPNT.EXE` 不是 PowerPoint 的引擎**。

```console
$ uv run --no-project python tools/scan_consts.py "POWERPNT.EXE" 0xE391C05F
C:/Program Files/Microsoft Office/root/Office16/POWERPNT.EXE: 1878624 bytes, 8 sections
  .text    va 0x00001001 raw 0x000400 size 0x001c00   <- 只有 7 KB 代码
  .rdata   va 0x00003001 raw 0x002000 size 0x001400
  ...
  .rsrc    va 0x00009001 raw 0x003e00 size 0x1c1c00   <- 1.8 MB 资源
constant 0xe391c05f (bytes 5f c0 91 e3)
  total 0
```

1.8 MB 的 exe 里 96% 是资源，常量一个都找不到——它只是个启动壳。真正的引擎在
**`root\vfs\ProgramFilesCommonX64\Microsoft Shared\OFFICE16\MSO.DLL`**（38.7 MB，
`.text` 段 0x19E6600 = 27 MB）。同目录下还有 `MSORES.DLL`（81 MB，资源）。

| 应用 | 启动壳 | 真正的格式代码 |
|---|---|---|
| PowerPoint | `root\Office16\POWERPNT.EXE`（1.8 MB） | `vfs\...\OFFICE16\MSO.DLL`（38 MB） |
| Word | `root\Office16\WINWORD.EXE`（1.6 MB） | `root\Office16\wwlib.dll`（54 MB） |
| Excel | `root\Office16\EXCEL.EXE`（69 MB，未拆分） | 同左；另有 `excelcnv.exe`（49 MB，格式转换器） |

**教训**：先看 `.text` 段大小，再决定把哪个文件丢给 Ghidra。把 1.8 MB 的壳丢进去
分析，只会得到一堆 `WinMain` 和资源。

## 1. CurrentUserAtom 的校验（MSO.DLL @ 0x3ADC5D）

`docs/ppt.md` §5 说「headerToken 在偏移 12，offsetToCurrentEdit 在偏移 16」，
而不少第三方实现（按 MS-PPT 2.3.2 的字面理解）用 +8/+12。用 MSO.DLL 定案：

```console
$ uv run --with capstone python tools/disasm.py MSO.DLL 0x3adc80 120
```

```asm
0x003adc80  lea     r8d, [rbx + 0x14]              ; 请求读 20 字节
0x003adc84  lea     rdx, [rbp + 7]                 ; 缓冲区 = rbp+7
0x003adc8b  mov     rax, qword ptr [rax + 0x18]    ; vtable+0x18 = IStream::Read
0x003adc8f  call    qword ptr [rip + 0x16ad844]
0x003adc95  test    eax, eax
0x003adc9d  cmp     dword ptr [rbp - 9], 0x14      ; 实际读到的字节数 == 20
0x003adca7  mov     eax, 0xff6
0x003adcac  cmp     word ptr [rbp + 9], ax        ; 缓冲区 +0x02 = recType == 0x0FF6
0x003adcb6  cmp     dword ptr [rbp + 0x13], 0xf3d1c4df   ; 缓冲区 +0x0C = headerToken（旧）
0x003adcbf  mov     al, 1
0x003adcc3  cmp     dword ptr [rbp + 0x13], 0xe391c05f   ; 缓冲区 +0x0C = headerToken（新）
0x003adcd3  mov     eax, dword ptr [rbp + 0x17]           ; 缓冲区 +0x10 = offsetToCurrentEdit
0x003adce4  mov     rdx, qword ptr [rbp - 1]
0x003adcee  call    qword ptr [rip + 0x16ad7e5]    ; vtable+0x28 = IStream::Seek
0x003adcff  mov     r8d, 0x28                      ; 再读 40 字节
0x003add05  lea     rdx, [rbp + 0x1f]              ; 缓冲区 = rbp+0x1F
0x003add20  mov     eax, 0xff5
0x003add25  cmp     word ptr [rbp + 0x21], ax      ; 缓冲区 +0x02 = recType == 0x0FF5
0x003add2b  cmp     dword ptr [rbp + 0x23], 0x20   ; recLen == 32（新版 UserEditAtom）
0x003add31  cmp     dword ptr [rbp + 0x43], ebx    ; payload+0x1C != 0
0x003add36  mov     bl, 1
0x003add5e  ret
```

缓冲区基址是 `rbp+7`，所以：

| 汇编里的偏移 | 相对缓冲区 | 含义 |
|---|---|---|
| `[rbp+9]` | +0x02 | recType（0x0FF6） |
| `[rbp+0x13]` | +0x0C | **headerToken** |
| `[rbp+0x17]` | +0x10 | **offsetToCurrentEdit** |
| `[rbp+0x21]` | +0x02 | recType（0x0FF5） |
| `[rbp+0x23]` | +0x04 | recLen |
| `[rbp+0x43]` | +0x24 | payload+0x1C（新版 UserEditAtom 的附加字段） |

结论（三条，全部进 `docs/ppt.md`）：

1. `headerToken` 在**偏移 12**、`offsetToCurrentEdit` 在**偏移 16** —— 与本仓库
   `crates/office-ppt/src/lib.rs` 的实现一致；按 +8/+12 读会读到 `size` 字段
   （0x14）和 token，然后把 0xE391C05F 当成偏移去 seek，必然失败。
2. **合法 token 有两个**：`0xE391C05F`（PPT 97+）与 `0xF3D1C4DF`（更早的版本）。
   只认一个会让老文件被判为"不是 PPT"。
3. 校验要求「先读 20 字节的 CurrentUserAtom 头，再按 offsetToCurrentEdit 读 40 字节
   的 UserEditAtom，并检查 recType == 0x0FF5」——这正是 `docs/ppt.md` §5 的解析顺序。

`recLen == 0x20` 那条分支：只有新版（32 字节）UserEditAtom 才继续，且 payload+0x1C
非零才返回 true。这看起来是「文档是否加密/需要密码」的判定（新版 UserEditAtom 在
`lastView` 之后多了 `encryptSessionPersistIdRef`），不是我们需要的路径，但说明
**UserEditAtom 有 28 字节和 32 字节两种长度**，解析时要按 recLen 分支。

## 2. 记录类型常量表（MSO.DLL @ 0x3EA5C5）

一个 `switch` 把内部枚举（0x2710 起）映射成 PPT 记录类型：

```asm
0x003ea5ca  add     edx, 0xffffd8f0          ; edx -= 0x2710
0x003ea5d4  cmp     edx, 0x22                ; 35 个 case
0x003ea5e8  mov     edx, [r8 + rax*4 + 0x3ea724]   ; 跳转表 @ 0x3EA725
0x003ea5f5  mov     eax, 0x412
0x003ea5fe  mov     byte ptr [rbp + 0x1a], 1
...
0x003ea61e  mov     eax, 0x3e8               ; Document
0x003ea62c  mov     eax, 0x3ed
0x003ea633  mov     eax, 0x3ee               ; Slide
0x003ea63a  mov     eax, 0x3ef               ; SlideAtom
0x003ea64f  mov     eax, 0x3f2               ; Environment
0x003ea66b  mov     eax, 0x3f8               ; MainMaster
0x003ea6dd  mov     eax, 0x40b               ; PPDrawingGroup
0x003ea6e7  mov     eax, 0x40c               ; PPDrawing
0x003ea6f1  mov     eax, 0x414
0x003ea719  mov     eax, 0x418
```

这条 switch 出现的常量集合：
`0x3E8 0x3EC 0x3ED 0x3EE 0x3EF 0x3F0 0x3F1 0x3F2 0x3F4 0x3F5 0x3F6 0x3F8 0x3FB 0x3FC
0x3FE 0x3FF 0x401 0x402 0x403 0x404 0x405 0x406 0x407 0x408 0x409 0x40A 0x40B 0x40C
0x412 0x414 0x415 0x416 0x417 0x418`
—— 与 `docs/ppt.md` §2 用到的 `0x3E8 Document`、`0x3EE Slide`、`0x3EF SlideAtom`、
`0x3F8 MainMaster`、`0x40B PPDrawingGroup`、`0x40C PPDrawing` 完全对上。

`0x3F3 SlidePersistAtom`、`0x3EA EOF`、`0x0FF0 SlideListWithText`、`0x1772
PersistDirectoryAtom` 不在这个 switch 里（它们由别的分发器处理），所以又单独扫了常量：

```console
$ uv run --no-project python tools/scan_consts.py MSO.DLL 0x1772
constant 0x00001772 (bytes 72 17 00 00)
  @0x1d6eb5c .rdata+0x387f5c (va 0x1d6ff5d)  ctx 34 17 00 00 52 17 00 00 72 17 00 00 b4 17 00 00 b6 17 00 00
  total 3
```

`0x1D6FF5D` 处是一张**升序的 u32 记录类型表**：`0x1734 0x1752 0x1772 0x17B4 0x17B6`。
0x1772 落在里面，说明它确实是 PPT 解析器认识的记录类型，而不是我们猜的偏移。

## 3. Word 的流名（wwlib.dll @ 0x236CC40，UTF-16）

`docs/doc.md` 说表流叫 `0Table` / `1Table`。wwlib.dll 里这两个名字是**紧挨着**的
UTF-16 字符串：

```console
$ uv run --no-project python - <<'EOF'
data = open(r"...\wwlib.dll", "rb").read()
i = data.find("1Table".encode("utf-16-le"))
print(hex(i), data[i-24:i+64])
EOF
0x236cc40 b'M\x00a\x00c\x00r\x00o\x00s\x00...\x001\x00T\x00a\x00b\x00l\x00e\x00\x00\x000\x00T\x00a\x00b\x00l\x00e\x00\x00\x00FinalDocSize...'
```

即 `"Macros"` `"1Table"` `"0Table"` `"WordDocument"`（`0x236CD08`）在同一张字符串表里，
和 `"FinalDocSize"`、`"Intermed..."` 等内部流名并列。这解释了为什么 `fWhichTblStm`
只有 1 个 bit：两个名字在实现里就是相邻的两个常量。

**扫描坑**：ASCII 版本的 `"1Table"` 在 wwlib.dll 里命中 **0** 次；ASCII 版 `"0Table"`
的 5 次命中全是 `UseWord2010TableStyleRules`、`dopidUseWord2010TableStyleRules` 之类
的**子串误报**。只扫 ASCII 会得出"Word 不用 1Table"的错误结论。

## 4. Excel：BIFF8 的命名（EXCEL.EXE @ 0x33B3C68）

`EXCEL.EXE` 的 .rdata 里有一组 UTF-16 剪贴板格式名：

```
"Value"  "BIFF8"  "RTF"  "SYLK"  "Bitmap"
```

`excelcnv.exe` 里同样有 `"BIFF8"`（@0x24B5448）。这是"BIFF8 就是 Excel 97 二进制格式"
的直接旁证。Excel 的记录 id（0x0809 BOF、0x00FC SST、0x00FD LABELSST…）没有像 PPT 那样
集中在少数几个常量里，本轮没有做逐条反编译——`docs/xls.md` 的结论全部来自真实文件
字节（WPS 样本 + Microsoft 的 SOLVSAMP.XLS），可复现性更强。

## 5. Word：流名选择与 fcClx/lcbClx（wwlib.dll）

### 一个函数返回四种流名：wwlib.dll `0x180403DDC`

```asm
0x180403DDC  sub  rsp, 0x28
0x180403DE0  cmp  ecx, 6            ; ecx = 流编号
0x180403DE3  jg   0x180403E60
0x180403DED  test ecx, ecx
0x180403DEF  je   0x180403E4E      ; 0 → "WordDocument"
0x180403DF1  sub  r8d, 1
0x180403DF5  je   0x180403E45      ; 1 → "Data"
0x180403DFB  je   0x180403E27      ; 2 → 0Table / 1Table
0x180403E01  je   0x180403E27      ; 3 → 0Table / 1Table
...
0x180403E27  cmp  ecx, 3
0x180403E2A  lea  rax, [rip + 0x1F6A41F]   ; → 0x18236E250  "0Table"
0x180403E31  lea  rcx, [rip + 0x1F6A408]   ; → 0x18236E240  "1Table"
0x180403E38  sete r9b                       ; r9b = (which == 3)
0x180403E3C  cmp  r9d, edx                  ; edx = fWhichTblStm（调用方传入）
0x180403E3F  cmovne rax, rcx
0x180403E43  jmp  0x180403EB8
0x180403E45  lea  rax, [rip + 0x1F68A24]   ; → 0x18236C870  "Data"
0x180403E4E  lea  rax, [rip + 0x1F6A4B3]   ; → 0x18236E308  "WordDocument"
```

结论（可直接对上文档）：

- 四个流名在同一个 helper 里：`WordDocument` / `Data` / `0Table` / `1Table`；
- **表流的选择是一个布尔入参**（`edx`），不是这个函数自己去读 FIB；
- `which == 2` 与 `which == 3` 是同一分支，靠 `sete`（which==3）与 `edx` 异或出结果——
  也就是「0Table / 1Table」这一对，由 `fWhichTblStm` 决定。

### fcClx / lcbClx 就是 FIB 的 `+0x1A2` / `+0x1A6`：wwlib.dll `0x180416132`

```asm
0x18041612B  movzx r14d, word ptr [r15 + 6]
0x180416130  mov   edx, ebx
0x180416132  mov   r8d, dword ptr [r13 + 0x1A2]    ; fcClx
0x180416139  mov   ecx, esi
0x18041613B  shr   r14d, 5
0x18041613F  mov   r12, rax
0x180416142  and   r14d, 1
0x180416146  mov   dword ptr [rbp - 0x50], r14d
0x18041614A  call  0x18046A70C
0x18041614F  mov   eax, dword ptr [r13 + 0x1A6]    ; lcbClx
```

`r13` 是 FIB 基址，`+0x1A2` 取 fcClx、紧接着 `+0x1A6` 取 lcbClx——**两个字段相邻**，
正好是 `fibRgFcLcb97` 的第 33 对（`0x9A + 33*8 = 0x1A2`）。这条证据支持了
`docs/doc.md` §2 的做法：按固定索引取，不看 `cbRgFcLcb`。

### FIB 的写入口：wwlib.dll `0x1805E79D8`

```asm
0x1805E79D8  mov word ptr [rbx + 0x144], 0x147    ; nFib = 0x147（Word 2007+）
0x1805E79E1  mov word ptr [rbx + 0x238], 0xA5EC   ; wIdent
```

`0xA5EC` 在 `.text` 里以立即数出现 57 次（其中多数是 `0x0000A5EC` 的 dword 写入），
用 `tools/scan_consts.py` 或直接扫 `EC A5` 就能定位；这是找 FIB 代码最快的入口。

### 方法：不依赖 Ghidra 的三步

1. 在 `.text` 里扫立即数/字符串的字节形态（`EC A5`、UTF-16 字符串）；
2. 把命中处的 VA 换算成文件偏移，用 `tools/disasm.py` 从一个已知指令边界往后反汇编；
3. 需要找引用时，扫 `E8 rel32`（call）和 `48 8D /r rel32`（RIP 相对 lea），
   用 `va + 指令长度 + rel == 目标 VA` 判定。

`tools/disasm.py` 之前只支持 PE32（ImageBase 按 32 位读），对 64 位的 wwlib.dll 会
算错基址、报 "va not in any section"；已按 Optional Header 的 magic（`0x10B` / `0x20B`）
分支修好。

## 6. 规范 vs 实测对照表

| 项目 | 规范说法 | 实测 / 二进制证据 | 影响 |
|---|---|---|---|
| CurrentUserAtom.headerToken 偏移 | 字面理解易读成 +8 | **+12**（MSO.DLL 0x3ADCB6） | 读错→seek 到 0xE391C05F |
| 合法 headerToken | 0xE391C05F | 0xE391C05F **和** 0xF3D1C4DF | 老文件被判非法 |
| UserEditAtom 长度 | 28 | 28 或 **32**（新版多 4 字节） | 越界读/漏字段 |
| PersistDirectoryAtom 的 count | 条目数 - 1 | **就是条目数**（hello 4/4、notes 7/7、encrypted 4/4） | 少读一个 persistId |
| .doc 的 cbRgFcLcb | 0x005D (93) | WPS 写 164、Microsoft 写 183 | 按它算长度会错位 |
| .doc 表流名 | 0Table / 1Table | wwlib.dll 里两个 UTF-16 字符串相邻 | — |
| .xls SST 计数器 | 4 字节 ×2 | 4 字节 ×2（本仓库第一版按 2 字节读，全空） | 文本全部丢失 |
| .xls 纯 ASCII 字符串 | 8-bit（grbit=0） | WPS 也写 UTF-16（grbit=1） | 只按规范解码会乱码 |
| .ppt 标题页文本类型 | 0 (Title) / 1 (Body) | WPS 写 6 (CenterTitle) / 5 (CenterBody) | 标题被判成"其它" |
| .xls ptgAttr 长度 | grbit u16 + data u16（5 B） | 真实文件是 **4 B**：grbit u8 + data u16 | 读到下一个 token 中间，公式错乱 |
| .xls ptgRef 语义 | 相对偏移 | **绝对行列 + 相对标志位** | 引用整表错位 |
| .doc CHPX/PAPX 叠加 | 文档未明说 | 条目之间**不累加**，每段从空属性集开始 | 后面所有段落继承前一段的粗体 |
| .doc 段落格式位置 | PAPX 的 grpprl | WPS 写进样式表（PAPX 为空） | 读不到居中/缩进 |
| .ppt pib 属性号 | 0x0104 | 写作 **0x4104**（带 fBid 位） | 永远找不到图片 |
| .ppt BLIP 记录类型 | PNG = 0xF01D | WPS 写 0xF01E（TIFF 的编号）但内容是 PNG | 按类型判断会解错格式 |
| .doc 脚注文本 | 只在 ccpFtn 区间 | WPS 又复制一份进 ccpText 区间 | 主文档里多出脚注文字 |
| .doc 样式表 STSHI | 首字段是 `cstd` | 两个真实文件都在最前面多一个 `cbStshi`（1118 / 1062）；`STD` 数组从 `2 + cbStshi` 起才恰好走完 `lcb` | 第一步就走进别人的数据，样式名全是乱码 |
| .doc PAPX `bOffset == 0` | 文档未明说 | 表示「同上一条」，不是「无格式」 | 第 2 段之后的样式/缩进全部丢失 |
| .xls BIFF5 `ptgRef` | 与 BIFF8 同构 | **行字带相对标志 + 列只有 1 字节**：`=A2` 是 `44 01 c0 00`，按 BIFF8 读成 `$A$49154` | BIFF5 公式全部错位 |
| .xls BIFF5 `FILEPASS` | 6 字节（类型 + 密钥 + 校验） | **4 字节**：`key u16` + `verifier u16`，没有类型字段 | 解不了 Excel 95 的 XOR 文件 |

## 6.1 StyleTextPropAtom：靠"正好走到末尾"定边界

`StyleTextPropAtom`（`0x0FA1`）的 payload 是 `rgTextPFRun` 紧跟 `rgTextCFRun`，
两段都没有元素个数，也没有长度前缀（MS-PPT 2.9.35）。所以"PF 段在哪结束"这件事
在文件里根本没有标记，只能枚举 PF 段的每个可能边界、试解 CF 段，要求**CF 段正好
走到 payload 末尾**。

`samples/你好.ppt` 是微软 PowerPoint 2016 写的（`SummaryInformation` 里写着
`Microsoft Office PowerPoint`），它把这件事钉死了：标题那一段 26 字节，
`03 00 00 00 | 00 00 | 00 08 00 00 | 01 00 | 03 00 00 00 | 00 00 06 00 | 3c 00 | ff 00 00 fe`
解出「3 字符、居中、60 磅、红色」，副标题 20 字节解出「6 字符、24 磅」，两者都
恰好用完 payload。同一个文件里主母版的 `TextMasterStyleAtom`（`0x0FA3`）也印证了
掩码位（标题级 `0x000A000E` → 10 磅 + 灰色 `89 89 89 fe`）。

实测与规范的两处出入：

| 项 | 规范 | 实测 |
| --- | --- | --- |
| 两段数组的边界 | 各自的 count 之和等于字符数，可据此切分 | `hello.ppt` 的 PF 段 count 是 1，而字符数是 2（段末回车没算进去）；只能靠"CF 段走到末尾"定位 |
| 空文本框 | 未说明 | 空正文框也带 `0x0FA1`，若实现成"格式化上一个文本记录"就会覆盖前一个形状（`pptfmt.ppt` 第 2 页标题被 24 磅盖掉） |

## 7. 复现命令汇总

```console
## 8. 加密：明文边界由「读取器还要看什么」决定

三份加密样本（`samples/encrypted.*`，密码 `secret123`）是逐字节验证的：
先用 `msoffcrypto` 确认密码正确（monkeypatch `DocumentRC4.verifypw` 打印 salt/verifier），
再用本仓库实现解密，比对明文结构是否自洽。结论不是「整条流加密」，而是：

| 格式 | 保持明文的区域 | 实测依据 |
|---|---|---|
| .doc | `WordDocument[0..0x44)`（FIB 头）、表流头 52 字节 | `0Table` 前 4 字节 `01 00 01 00` = vMajor/vMinor，其后才是高熵数据 |
| .xls | **每条记录的 4 字节头** + BOF/FilePass/UsrExcl/FileLock/InterfaceHdr/RRDInfo/RRDHead 负载 + `BoundSheet8` 前 4 字节 | 未解密时仍能逐条走完整个记录流并找到 `FILEPASS@20` |
| .ppt | 持久目录、`UserEditAtom`、`Current User`、`CryptSession10Container` 自身 | 未解密时能读出 4 个持久对象偏移，第 4 个指向 `0x2F14` |

三个被真实文件纠正的点：

1. **旧式 RC4 头是 52 字节，不是 62 字节**：`encrypted.doc` 的 `0Table` = `01 00 01 00` +
   salt(16) + verifier(16) + verifierHash(16)，没有 `EncryptionHeader`，也没有
   `verifierHashSize` 字段（MS-OFFCRYPTO 2.3.6.1 的描述对应的是另一种写出方式）。
2. **`UserEditAtom.encryptSessionPersistIdRef` 在 payload[28..32]**，不是 [26..30]：
   `encrypted.ppt` 的 `UserEditAtom` recLen = `0x20`，`[24..26]` lastView=1、`[26..28]` unused=0、
   `[28..32]` = 4。按 [26..30] 读得到 `0x00040000`，持久目录里查不到，于是静默退化成「未加密」。
3. **持久目录 count 字段 = 条目数**，不是「条目数 - 1」：`encrypted.ppt` 的
   `01 00 40 00` → persistId 1、高 12 位 4，payload 20 字节正好 4 个偏移
   （0 / 0x587 / 0xA633 / 0xAD8D）。三个文件交叉验证：`hello.ppt` 4/20 字节、
   `notes.ppt` 7/32 字节（7 个对象：Document、MainMaster、Slide×2、Notes×2、NotesMaster）、
   `encrypted.ppt` 4/20 字节。全部满足 `payload = 4 + n*4`，即 count 就是 n。

`encrypted.ppt` 的解密结果（块号 = persistId，块大小按密钥位数取整）：

```text
persist 1 @0      → 0f 00 e8 03 7f 05 00 00   Document  (0x03E8, len 0x57F)
persist 2 @0x587  → 0f 00 f8 03 a4 a0 00 00   MainMaster(0x03F8, len 0xA0A4)
persist 3 @0xA633 → 0f 00 ee 03 52 07 00 00   Slide     (0x03EE, len 0x752)
persist 4 @0xAD8D → CryptSession10Container (0x2F14, len 198, 不加密)
```

顺带一个反证：`Current User` 里的 `headerToken` 在 WPS 写的加密文件里仍是
`0xE391C05F`（明文），而 MS-OFFCRYPTO 说加密后应改成 `0xF3D1C4DF`。所以读取器必须
两个 token 都接受 —— 这与 MSO.DLL `0x3ADC5D` 那个校验函数里连着比较两个常量的行为一致。

# 常量 / 字符串
$ uv run --no-project python tools/scan_consts.py MSO.DLL 0xE391C05F 0x1772 --str=WordDocument

# 反汇编（capstone）
$ uv run --with capstone python tools/disasm.py MSO.DLL 0x3adc80 120
$ uv run --with capstone python tools/disasm.py MSO.DLL --back 0x3adcc6 300

# Word：流名选择 / fcClx / FIB 写入口（wwlib.dll 是 PE32+，脚本已支持）
$ uv run --with capstone python tools/disasm.py wwlib.dll 0x180403ddc 200
$ uv run --with capstone python tools/disasm.py wwlib.dll 0x180416123 30
$ uv run --with capstone python tools/disasm.py wwlib.dll 0x1805e79d0 30

# Ghidra（headless，MCP/CLI）
$ GHIDRA_INSTALL_DIR=C:/ghidra_12.2_PUBLIC ~/re/.venv/Scripts/ghlab.exe status
$ GHIDRA_INSTALL_DIR=C:/ghidra_12.2_PUBLIC ~/re/.venv/Scripts/ghlab.exe decompile wwlib.dll 0x...
```

Ghidra 工程里已导入的二进制（`C:\Users\admin\re\proj\ghlab`）：`POWERPNT.EXE`（壳，
无格式代码）、`wwlib.dll`（54 MB，Word 引擎）。分析 54 MB 的 DLL 需要较长时间，
本文的 PPT 证据先用 capstone 静态反汇编拿到，结论与 Ghidra 反编译一致时可互相印证。

## 9. 拿到「微软写的」真值：HKCU 覆盖 HKLM

本文一半以上的结论来自「同一个格式，微软写一遍、WPS 写一遍，然后对拍」。
但在这台机器上 `New-Object -ComObject Word.Application` 永远启动 `wps.exe`，
原因不是代码写错了，而是注册表里同一个 CLSID 注册了两份：

```
HKLM\SOFTWARE\Classes\CLSID\{000209FF-…}\LocalServer32 = …\Office16\WINWORD.EXE /Automation
HKCU\Software\Classes\CLSID\{000209FF-…}\LocalServer32 = …\Kingsoft\…\office6\wps.exe /Automation
```

`HKCU\Software\Classes` 与 `HKLM\SOFTWARE\Classes` 是合并视图，**HKCU 逐值优先**，
所以 ProgID → CLSID → `LocalServer32` 一路走到 WPS。`{00024500-…}`（Excel）和
`{91493441-…}`（PowerPoint）同理，分别是 `et.exe` / `wpp.exe`。

只需临时改写 HKCU 那一个默认值（`TypeLib`、`ProgID` 等仍从 HKLM 读），就能让 COM
落到微软 Office；`tools/ms_samples.ps1` 把三个值备份到内存、生成样本、`finally` 还原，
不留下持久修改。这一步换来了三件以前拿不到的东西：

| 产物 | 以前为什么拿不到 | 换来了什么结论 |
|---|---|---|
| `samples/ms/excel95.xls`（`SaveAs(…, 39)`） | WPS 忽略 `xlExcel9795`，照样写 BIFF8 | 真 BIFF5 `0x0500`；`ptgRef` 的行字带相对标志、列只有 1 字节（§6 表末两行） |
| `samples/ms/xor.xls`（Excel 95 + 修改密码） | WPS 写 RC4 | 真 XOR：`FILEPASS` 只有 **4 字节**，`f2 db 28 96` = `key 0xDBF2` + `verifier 0x9628` |
| `samples/ms/locked.doc`（Word + 打开密码） | WPS 写的是 v1.1 的 52 字节头 | 微软写 **RC4 CryptoAPI 128 位**（`algId 0x6801`、CSP 名带 `\0` 结尾），CryptoAPI 密钥派生得到独立验证 |

另外 `samples/ms/formatted.ppt` 证明微软的实现也会把**空文本框**配上
`StyleTextPropAtom`，这正是「必须用 `pending` 记住刚压入的那一段」这个修复的来源。

## 10. 微软写的 .doc 暴露的两个「约定差异」

同一个字段在 WPS 与微软的实现里可能有不同的约定，只测一种写入器永远发现不了。
`samples/ms/formatted.doc`（Word 2016 写出：无 `WpsCustomData`、有 `\x01CompObj`）一次性暴露了两处：

### 10.1 压缩片段的 `fc` 没有折半

CLX 里的 PCD 是 `fc = 0x40001000`（`bit30` = `fCompressed`），而 CHPX/PAPX 页里的
`rgfc` 从 **2048** 开始。WPS 写的是 `fc = 0x800`——已经折半。

```text
WPS  samples/formatted.doc   PCD.fc = 0x00000800  → 字节偏移 2048   (已折半)
MS   samples/ms/formatted.doc PCD.fc = 0x40001000 → 字节偏移 2048   (需 /2)
```

按 `fc & 0x3FFFFFFF` 直接比较的后果：`char_runs` 有 53 条、`para_runs` 有 37 条，
但 `paragraphs_rich()` 每段 `spans == 0`、`in_table == None`——**格式与表格全部静默丢失**。

### 10.2 `ToggleOperand 0x81` = 取反，不是关

同一文件的 CHPX 页（`pn = 5`，`crun = 53`，`rgfc[0] = 2048`）里，加粗段落的 sprm 是 `35 08 81`：

```text
[0] fc 2048..2049  cb 29  16 68 3b4b3a00 | 30 4a 2900 | 35 08 81 | 42 2a 06 | 43 4a 2400 | 6f 28 01 | 70 68 ff000000
[1] fc 2049..2066  cb 31  15 68 … | 16 68 … | 35 08 81 | 42 2a 06 | 43 4a 2400 | 6f 28 01 | 70 68 ff000000
[2] fc 2066..2087  cb 21  15 68 … | 16 68 … | 36 08 81 | 3e 2a 01 | 6f 28 01
```

Word 自己（COM 读取）确认：第 1 段 `bold = -1, size = 18, color = 255`，第 2 段 `italic = -1`。
所以 `0x81` 必须解释成「相对默认值取反」→ 默认关，取反即开；`0x80` 是「保持原状」。
把它读成「关」的后果是所有加粗/斜体全部消失，而字号、颜色仍然正确——这种
「只丢一半属性」的现象正是最容易被误判成「解析没对齐」的地方。

| 约定 | WPS | 微软 | 读错的后果 |
|---|---|---|---|
| 压缩片段 `fc` | 已折半（`0x800`） | 未折半（`0x40001000`） | 格式/表格全丢，且不报错 |
| `ToggleOperand` | 用 `0`/`1` | 用 `0x81`（取反） | 粗体/斜体全丢，字号颜色还在 |

## 10. 切除实验：Word 到底强制要求哪几对 FIB 指针

读规范看不出"哪些字段可以不写"。做法是拿 `samples/hello.doc`，逐对把 FIB 的
`fc/lcb` 清零（其余字节不动），用自己的 CFB 写入器重打包，再让 Word 打开：

| 清掉的对 | 索引 | 结果 |
|---|---|---|
| `PlcfBteChpx` + `PlcfBtePapx` | 12, 13 | COM `Documents.Open` 90 秒无返回（挂死） |
| 只清 `PlcfBteChpx` | 12 | 挂死 |
| 只清 `PlcfBtePapx` | 13 | 挂死 |
| `StshfOrig` + `Stshf` | 0, 1 | 明确报错"文档打开失败。" |
| `PlcfSed` | 6 | 明确报错"文档打开失败。" |

对照组：同样的重打包方式、不清任何字段 → Word 正常读出 `paras=5 chars=82`。

结论：**格式页（CHPX/PAPX FKP）不是可选项**。这也解释了从零写 `.doc` 时的现象 ——
只写 FIB + CLX 会挂死，补上两页 FKP 之后 Word 立刻能打开（`docs/doc.md` §14）。

另一个必须用真应用才能发现的问题在容器层：目录项的兄弟链必须按
（名字长度，大写名）排序，否则 Word/WPS 挂死而 Excel 无感（`docs/cfb.md` §8）。

## 11. PowerPoint 的隐含校验：`TextSpecialInfoAtom` 的字符数

写入器生成的文件一度被 PowerPoint 一律判为「Office 检测到文件有问题」，而我们的读取器能完整读回、`Document`/`PersistDirectory`/`UserEditAtom` 与真文件逐字节一致。定位过程（`docs/ppt.md` §9.3）：用 COM 让 PowerPoint 写出**同样文字**的对照文件 `samples/ms/same.ppt`，再 diff 两者最后一个 `Slide` 的子树 —— 唯一差异是：

| 字段 | PowerPoint | 我们（克隆模板后） |
|---|---|---|
| `TextSpecialInfoAtom` (0x0FAA) 首字 | `0x0000000E` = 14 = 标题字符数 | `0x0000000B` = 11 = 模板文字长度 |
| 正文那条 | `0x0000001B` = 27 | `0x0000000C` = 12 |

即该原子的第一个 u32 是「这段文字覆盖的字符数」，PowerPoint 会与 `TextCharsAtom`/`TextBytesAtom` 的长度对账。MS-PPT 2.9.52 只把它记作 `cch`，没有说明它会被校验 —— 这条属于**实测得到、规范未明示**的约束。修好后 PowerPoint 直接接受，两页文字（含中文）全部正确。
