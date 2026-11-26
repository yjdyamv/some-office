# 加密文档：RC4 与 XOR 混淆

Office 97-2003 的三种格式都用「流式 RC4」保护文件，但**加密的边界各不相同**：只有那些读取结构所必需的字节留在明文里，其余全部加密。这一点比密钥派生更容易踩坑 —— 拿到一份加密文件时，先想清楚「读取器还需要看到什么」。

本文记录 `office-crypto` crate 的实现，以及用真实文件（WPS 生成 + 参考实现 `msoffcrypto` 对拍）验证过的细节。

## 1. 三种加密方案

| 方案 | 出现在 | 密钥长度 | 头部长度 | 校验方式 |
|---|---|---|---|---|
| XOR 混淆（method 1） | .xls 旧文件 | 16 位密钥 | BIFF8 6 B（类型 + 密钥 + 校验）/ BIFF5 4 B（密钥 + 校验） | `verifier ^ 0xCE4B == verificationBytes` |
| 旧式 RC4（v1.1） | .doc / .xls | 40 位 | 52 B | RC4 解 `verifier‖hash` 后 `MD5(verifier) == hash` |
| RC4 CryptoAPI（v2-4） | .doc / .xls / .ppt | 40/56/128 位 | 变长 | RC4 解 `verifier‖hash` 后 `SHA1(verifier) == hash` |

`office_crypto::parse_header` 只看 `vMajor`/`vMinor`：

- `1.1` → 52 字节旧式头：`salt[16]`、`encryptedVerifier[16]`、`encryptedVerifierHash[16]`，**没有** `EncryptionHeader`、**没有** `verifierHashSize` 字段（MS-OFFCRYPTO 2.3.6.1 的 62 字节描述与实际文件不符）。
- `2.2` / `3.2` / `4.2` → CryptoAPI 头：`flags`、`sizeExtra`、`algId`、`algIdHash`、`keySize`、`providerType`、两个保留字段、UTF-16 的 CSP 名称（NUL 结尾），然后是 `saltSize`、`salt`、`encryptedVerifier[16]`、`verifierHashSize`、`encryptedVerifierHash[hashSize]`。

实测 `samples/encrypted.ppt`：`algId=0x6801`（CALG_RC4）、`algIdHash=0x8004`（CALG_SHA1）、`keySize=0x80`（128 位）、`providerType=1`、CSP = `Microsoft Enhanced Cryptographic Provider v1.0`。

## 2. 密钥派生

**旧式 RC4**（`office_crypto::legacy_key`）：

```text
h0  = MD5(password_utf16le)              # 注意：UTF-16LE，不是 UTF-8
ib  = (h0[0..5] + salt) * 16             # 336 字节
h1  = MD5(ib)
key = MD5(h1[0..5] + block_le32)[0..16]  # block = 块号，从 0 开始
```

**CryptoAPI**（`office_crypto::cryptoapi_key`）：

```text
h0  = SHA1(salt + password_utf16le)      # 顺序是 salt 在前
key = SHA1(h0 + block_le32)
key = key[0..keySize/8]                  # 40 位时取 5 字节后补 11 个 0
```

两个坑：

1. **密码必须转 UTF-16LE**。用 UTF-8 得到的哈希完全不同，而且校验会安静地失败。
2. **校验用的 RC4 流是连续的**：`encryptedVerifier` 和 `encryptedVerifierHash` 拼在一起一次性解密，不能在两个字段之间重置状态。我第一版每字段重置，两个 16 字节都解不出来。

## 3. 各格式的加密边界

### 3.1 .doc —— FIB 明文，其余按 512 字节块

`WordDocument` 流从偏移 0 开始按 **512 字节**分块加密，块号从 0 递增；但 `[0, 0x44)` 的 FIB 头在加密后被原样写回明文（否则连 `nFib`、`fEncrypted` 都读不到）。表流（`0Table`/`1Table`）同样按 512 字节块加密，头部 52 字节明文。

```rust
// crates/office-doc/src/lib.rs
let mut plain = office_crypto::decrypt(scheme, password, &word_stream, 512)?;
plain[..FIB_PLAIN_LEN].copy_from_slice(&word_stream[..FIB_PLAIN_LEN]);   // 0x44
```

加密标志是 FIB `flags` 的 **bit 8（`0x0100`，fEncrypted）**；注意 `0Table`/`1Table` 的选择仍然是 bit 9，两者互不影响。

### 3.2 .xls —— 记录头全部明文

`Workbook` 流按 **1024 字节**分块加密，块号从 0 递增。但读取器必须能逐条走记录，所以下列内容保持明文：

- **每一条记录的 4 字节头**（`id` + `len`）—— 这条最关键，漏了它就无法遍历；
- `BOF`、`FilePass`、`UsrExcl`、`FileLock`、`InterfaceHdr`、`RRDInfo`、`RRDHead` 的整个负载；
- `BoundSheet8` 负载的**前 4 字节**（`lbPlyPos`，工作表在流中的偏移）。

```rust
// crates/office-xls/src/lib.rs —— 先整体解密，再把明文区段盖回去
let mut plain = office_crypto::decrypt(scheme, password, stream, 1024)?;
for (offset, len) in plaintext_regions(stream)? {
    plain[offset..offset + len].copy_from_slice(&stream[offset..offset + len]);
}
```

`FilePass`（`0x002F`）的负载是 `wEncryptionType u16` + 加密头，所以旧式 RC4 时 `lcb = 54 = 2 + 52`。

### 3.3 .ppt —— 按持久对象分别加密

PowerPoint 不是一条长流，而是**每个持久对象（`Document`、`MainMaster`、每张 `Slide`…）各自一条 RC4 流，块号 = 持久 ID**。加密头放在一个 `CryptSession10Container`（`0x2F14`）记录里，它的偏移由持久目录给出：

```text
Current User  ──offsetToCurrentEdit──▶  UserEditAtom (0x0FF5)
UserEditAtom  ──payload[28..32]──────▶  encryptSessionPersistIdRef   // 注意不是 [26..30]
持久目录      ──persistId → offset───▶  CryptSession10Container (0x2F14)  → EncryptionInfo
持久目录      ──persistId → offset───▶  各对象（块号 = persistId）
```

两个必须照抄的怪癖：

1. **块大小取整到密钥位数**，不是字节数：`block_size = keySize * (total / keySize + 1)`（`keySize` = 128 等）。对 128 位密钥，`total/128+1` 个 128 字节块。
2. 对象长度**不能读自己的 `recLen`**（它也被加密了），只能靠「下一个持久对象的偏移」推算；持久目录按偏移排序后，第 *i* 个对象的结束位置就是第 *i+1* 个的起点。

`UserEditAtom` 的 `recLen` 有两种：`0x1C`（28 字节，没有 `encryptSessionPersistIdRef`，即未加密）和 `0x20`（32 字节，字段在 `[28..32]`）。把偏移写成 `[26..30]` 会读到 `0x00040000` 这种垃圾值，然后一路 `NotEncrypted` —— 这个 bug 我踩过。

解密后把容器记录清零，否则再解析时还会被判定为「加密」。

## 4. 用法

```rust
// .doc
let doc = office_doc::parse_with_password(&bytes, "secret123")?;
println!("{}", doc.encryption.as_ref().map(|s| s.name()).unwrap_or("none"));

// .xls
let wb = office_xls::parse_with_password(&bytes, "secret123")?;

// .ppt
let deck = office_ppt::parse_stream_with_password(&stream, user.as_deref(), "secret123")?;
```

```bash
officedump info samples/encrypted.doc                 # → encryption  protected (pass --password <pw>)
officedump info samples/encrypted.doc --password secret123
officedump text samples/encrypted.xls --password secret123
```

## 5. XOR 混淆（method 1）

旧式 `.xls` 可能只做「混淆」而不是加密（MS-OFFCRYPTO 2.3.7.1）。它没有 RC4，只有一个 16 字节的替换表：

```text
verifier = (password 的每个字符异或后查表) ... 
key      = verifier ^ 0xCE4B
array[i] = ror8(key_byte ^ PAD[(i + len(password)) % 15], 5)
```

解密时逐字节 `ror8(cipher ^ array[index], 5)`，`index` 的起点是 **`(记录负载起点 + 负载长度) % 16`**，
`BoundSheet8` 再额外 +4（`lbPlyPos` 那 4 字节是明文）。

和 RC4 一样，**4 字节记录头始终明文**——否则流没法遍历；XOR 还额外保留
`BOF` / `FilePass` / `UsrExcl` / `FileLock` / `InterfaceHdr` / `RRDInfo` / `RRDHead`
的整段负载。所以解密是「按记录逐段」而不是整块：

```rust
// crates/office-xls：走完记录后只对需要解密的负载做变换
let start = r.offset + 4;
let run = office_crypto::xor::decrypt_run(&stream[start..r.offset + total], &arr, start, false);
```

`FILEPASS` 记录解密后留下一个「类型清零、长度保留」的存根，这样后面所有偏移都不变。

### 5.1 `FILEPASS` 载荷有 4 字节和 6 字节两种

| 版本 | 载荷 | 实测 |
|---|---|---|
| BIFF8（`0x0600`） | `wEncryptionType u16 = 0` + `key u16` + `verificationBytes u16` | 6 B |
| BIFF5（`0x0500`） | `key u16` + `verificationBytes u16`，**没有类型字段** | 4 B |

微软 Excel 95 写出的 `samples/ms/xor.xls` 里，`FILEPASS` 载荷就是 `f2 db 28 96`：
`0xDBF2` 正好是 `xor::key("secret123")`，`0x9628` 是校验值。读取器按长度分派，
4 字节一律当 XOR（BIFF5 只有这一种混淆）。

```rust
// crates/office-xls/src/lib.rs
if payload.len() == 4 {
    return decrypt_xor_stream(stream, password, &recs, filepass);
}
```

`office_crypto::xor` 提供 `key`（16 位混淆密钥）/ `verifier`（文件里存的校验值）/
`array`（16 字节密钥流数组）/ `verify` / `decrypt_run`，用参考实现的向量做了单元测试
（`verify("VelvetSweatshop", 0x9A0A) == true`）。`office_xls::parse_with_password`
会自动按 `wEncryptionType`（或载荷长度）分派到 RC4 或 XOR。

**真实样本**：`samples/ms/xor.xls` 是微软 Excel 95 带「修改密码」保存的文件，
本仓库用 `secret123` 能完整读回 `XOR-SECRET` / `7.5`；`crates/office-xls` 的
`xor_tests` 另外用「按参考算法逆向加密一条合成记录流、再解回来」做不依赖文件的回归。

## 6. 与参考实现对照

| 项目 | 参考实现 `msoffcrypto` | 本文实现 | 结果 |
|---|---|---|---|
| 旧式 RC4 密钥 | `MD5(MD5((h0[:5]+salt)*16)[:5] + block)` | 同 | 一致 |
| CryptoAPI 密钥 | `SHA1(SHA1(salt+pw)+block)` | 同 | 一致 |
| .doc 块大小 | 512 | 512 | 一致 |
| .xls 块大小 | 1024 | 1024 | 一致 |
| .ppt 块大小 | `keySize*((8+recLen)/keySize+1)` | `keySize*(total/keySize+1)` | 一致（`total = 8+recLen`） |
| 加密头 v1.1 长度 | 52 B | 52 B | 一致（规范写 62 B） |
| 持久目录 count 字段 | 未使用 | 高 12 位当 count，越界则回退 | 实测 count = 条目数，规范写「条目数-1」 |

`msoffcrypto` 仅作为**验证器**使用（`uv run --with msoffcrypto-tool`），本仓库零依赖。
