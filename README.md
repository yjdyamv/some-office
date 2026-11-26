# some-office

反向工程 Microsoft Office 97-2003 闭源二进制格式（`.doc` / `.xls` / `.ppt`），
用**零依赖 Rust** 实现读取器，并留下「结构文档 + 证据 + 可运行示例」。

```console
$ cargo test                     # 95 个测试（单元 + 真实文件集成）
$ cargo run -p officedump -- text samples/hello.doc
$ cargo run -p officedump -- html samples/formatted.doc
$ cargo run -p officedump -- text samples/ms/locked.doc --password secret123
$ cargo run -p officedump -- media samples/image.ppt /tmp/pics
$ cargo run -p officedump -- styles samples/ms/formatted.doc
$ cargo run -p officedump -- stories samples/stories.doc
```

## 文档

| 文件 | 内容 |
|---|---|
| [docs/README.md](docs/README.md) | 总览：能力表、限制、方法论、CLI |
| [docs/cfb.md](docs/cfb.md) | 复合文档容器（OLE2/CFB）：扇区、FAT、mini-stream、目录 |
| [docs/doc.md](docs/doc.md) | Word `.doc`：FIB、分片表、文本故事、CHPX/PAPX、样式表、表格、图片 |
| [docs/xls.md](docs/xls.md) | Excel `.xls`：BIFF8/BIFF5 记录、SST、单元格、公式 RPN、数字格式 |
| [docs/ppt.md](docs/ppt.md) | PowerPoint `.ppt`：记录树、持久对象、幻灯片顺序、文字格式、图片、备注 |
| [docs/encryption.md](docs/encryption.md) | 加密：XOR 混淆、RC4 v1.1、RC4 CryptoAPI，以及各格式的「明文边界」 |
| [docs/evidence.md](docs/evidence.md) | 证据：反汇编、字符串、规范与实测的差异对照表 |

## 结构

```
crates/
  office-cfb/      CFB 容器
  office-enc/      代码页 / UTF-16
  office-crypto/   MD5 / SHA1 / RC4 / XOR
  office-blip/     OfficeArt 图片嗅探（dib/png/jpeg/…）
  office-doc/      .doc
  office-xls/      .xls
  office-ppt/      .ppt
  officedump/      CLI
samples/           样本（WPS 写出的 + samples/ms/ 下微软 Office 写出的）
tools/             逆向辅助脚本与样本生成脚本
```

样本全部由 `tools/make_samples*.ps1`、`tools/ms_samples.ps1` 生成，可复现。
