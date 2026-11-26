#!/usr/bin/env python3
"""Scan a PE for 32-bit constants / ASCII strings and show where they live.

Used while reverse-engineering the Office binary formats: it answers "does this
constant actually appear in the shipped binary, and how often" without needing
Ghidra to be free.

    uv run --no-project python tools/scan_consts.py WORD.EXE 0xA5EC 0x0FA0
    uv run --no-project python tools/scan_consts.py WORD.EXE --str WordDocument
"""
import sys

def pe_sections(data: bytes):
    """Yield (name, vaddr, raw_offset, raw_size) for each PE section."""
    if data[:2] != b"MZ":
        return
    import struct
    e_lfanew = struct.unpack_from("<I", data, 0x3C)[0]
    if data[e_lfanew:e_lfanew + 4] != b"PE\0\0":
        return
    nsec, = struct.unpack_from("<H", data, e_lfanew + 6)
    opt_size, = struct.unpack_from("<H", data, e_lfanew + 20)
    base = e_lfanew + 24 + opt_size
    image_base, = struct.unpack_from("<I", data, e_lfanew + 24 + 28)
    for i in range(nsec):
        off = base + i * 40
        name = data[off:off + 8].rstrip(b"\0").decode("latin1")
        vsize, vaddr, rsize, roff = struct.unpack_from("<IIII", data, off + 8)
        yield name, vaddr + image_base, roff, rsize

def locate(sections, file_off):
    for name, vaddr, roff, rsize in sections:
        if roff <= file_off < roff + rsize:
            return f"{name}+{file_off - roff:#x} (va {vaddr + file_off - roff:#x})"
    return "?"

def main(argv):
    path = argv[1]
    data = open(path, "rb").read()
    secs = list(pe_sections(data))
    print(f"{path}: {len(data)} bytes, {len(secs)} sections")
    for name, vaddr, roff, rsize in secs:
        print(f"  {name:8s} va {vaddr:#010x} raw {roff:#08x} size {rsize:#08x}")
    for arg in argv[2:]:
        if arg == "--str":
            continue
        if arg.startswith("--str="):
            needle = arg.split("=", 1)[1].encode("latin1")
            print(f"\nstring {needle!r}")
            start = 0
            n = 0
            while True:
                i = data.find(needle, start)
                if i < 0:
                    break
                n += 1
                if n <= 8:
                    print(f"  @{i:#08x} {locate(secs, i)}")
                start = i + 1
            print(f"  total {n}")
            continue
        val = int(arg, 0) & 0xFFFFFFFF
        pat = val.to_bytes(4, "little")
        print(f"\nconstant {val:#010x} (bytes {pat.hex(' ')})")
        start = 0
        n = 0
        while True:
            i = data.find(pat, start)
            if i < 0:
                break
            n += 1
            if n <= 8:
                ctx = data[max(0, i - 8):i + 12].hex(" ")
                print(f"  @{i:#08x} {locate(secs, i):34s} ctx {ctx}")
            start = i + 1
        print(f"  total {n}")

if __name__ == "__main__":
    main(sys.argv)
