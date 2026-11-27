import sys, subprocess, os
sys.path.insert(0, "tools")
from cfb_list import CFB

SRC = "samples/written.xls"
GROUPS = [
    ("XFEXT",     {0x087D}),
    ("StyleExt",  {0x0892}),
    ("STYLE",     {0x0293}),
    ("XF",        {0x00E0}),
    ("FONT",      {0x0031}),
    ("FORMAT",    {0x041E}),
    ("ENV",       {0x00E1, 0x00C1, 0x00E2, 0x0161, 0x013D, 0x009C, 0x0012, 0x0013,
                   0x0040, 0x008D, 0x000E, 0x01B7, 0x00DA, 0x0160}),
    ("WRITEACCESS", {0x005C}),
    ("SUPBOOK/EXT", {0x01AE, 0x0017}),
]

def recs(d):
    p, out = 0, []
    while p + 4 <= len(d):
        rid = int.from_bytes(d[p:p + 2], "little")
        ln = int.from_bytes(d[p + 2:p + 4], "little")
        if p + 4 + ln > len(d):
            break
        out.append((rid, d[p:p + 4 + ln]))
        p += 4 + ln
    return out, p

src = CFB(SRC).read("Workbook")
rs, end = recs(src)
ge = next(i for i, (r, _) in enumerate(rs) if r == 0x000A)

def build(drop, path):
    body = [raw for i, (rid, raw) in enumerate(rs) if not (i < ge and rid in drop)]
    data = b"".join(body) + src[end:]
    open("/tmp/b.bin", "wb").write(data)
    subprocess.run(["cargo", "run", "-q", "--release", "--example", "packxls",
                    "--", "/tmp/b.bin", path], check=True, capture_output=True)
    return len(data)

for name, ids in GROUPS:
    path = "samples/bis_%s.xls" % name.replace("/", "_")
    n = build(ids, path)
    out = subprocess.run(["pwsh", "-NoProfile", "-File", "tools/check_xls.ps1",
                          "-File", os.path.abspath(path)],
                         capture_output=True, text=True, encoding="utf-8", errors="replace").stdout
    ok = "ERR" not in out
    print("%-14s drop=%-3d stream=%-6d %s" % (name, sum(1 for i,(r,_) in enumerate(rs) if i < ge and r in ids), n, "OK" if ok else "REJECT"))
