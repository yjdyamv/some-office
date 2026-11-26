#!/usr/bin/env python3
"""Disassemble a range of a PE by virtual address (capstone).

    uv run --with capstone python tools/disasm.py MSO.DLL 0x3ADCC6 120
    uv run --with capstone python tools/disasm.py MSO.DLL --back 0x3ADCC6 400

`--back` walks backwards looking for a plausible function prologue so the dump
starts at the function head instead of mid-instruction.
"""
import struct
import sys

import capstone


def sections(data: bytes):
    e = struct.unpack_from("<I", data, 0x3C)[0]
    nsec, = struct.unpack_from("<H", data, e + 6)
    opt, = struct.unpack_from("<H", data, e + 20)
    base = e + 24 + opt
    # PE32 vs PE32+: ImageBase is u32 for 0x10B, u64 for 0x20B.
    magic, = struct.unpack_from("<H", data, e + 24)
    if magic == 0x20B:
        image_base, = struct.unpack_from("<Q", data, e + 24 + 24)
    else:
        image_base, = struct.unpack_from("<I", data, e + 24 + 28)
    out = []
    for i in range(nsec):
        off = base + i * 40
        name = data[off:off + 8].rstrip(b"\0").decode("latin1")
        vsize, vaddr, rsize, roff = struct.unpack_from("<IIII", data, off + 8)
        out.append((name, image_base + vaddr, vsize, roff, rsize))
    return image_base, out


def va_to_off(secs, va):
    for name, sva, vsize, roff, rsize in secs:
        if sva <= va < sva + max(vsize, rsize):
            return roff + (va - sva)
    raise SystemExit(f"va {va:#x} not in any section")


def find_function_start(data, secs, va, limit=0x800):
    """Walk back for `int3` padding followed by a prologue-ish byte."""
    off = va_to_off(secs, va)
    start = max(0, off - limit)
    for i in range(off - 1, start, -1):
        if data[i] in (0xCC, 0xC3, 0x90) and data[i + 1] not in (0xCC, 0x90):
            return va - (off - (i + 1))
    return va


def main(argv):
    back = "--back" in argv
    args = [a for a in argv[1:] if not a.startswith("--")]
    path, va = args[0], int(args[1], 0)
    count = int(args[2]) if len(args) > 2 else 200
    data = open(path, "rb").read()
    image_base, secs = sections(data)
    if back:
        va = find_function_start(data, secs, va)
    off = va_to_off(secs, va)
    md = capstone.Cs(capstone.CS_ARCH_X86, capstone.CS_MODE_64)
    md.detail = False
    code = data[off:off + count]
    print(f"{path}  va {va:#x} (file {off:#x})")
    for ins in md.disasm(code, va):
        note = ""
        m = ins.op_str
        for token in m.replace(",", " ").split():
            if token.startswith("0x") and len(token) > 4:
                try:
                    v = int(token, 16)
                except ValueError:
                    continue
                if v > 0x1000:
                    note += f"   ; {v:#x}"
        print(f"  {ins.address:#010x}  {ins.mnemonic:<7} {ins.op_str}{note}")


if __name__ == "__main__":
    main(sys.argv)
