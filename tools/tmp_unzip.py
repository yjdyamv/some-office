import io, struct, sys, zipfile
path = sys.argv[1] if len(sys.argv) > 1 else "samples/ms/minimal.ppt"
data = open(path, "rb").read()
i = 0
seen = 0
while seen < 40:
    i = data.find(b"PK\x03\x04", i)
    if i < 0:
        break
    vi, kind, ln = struct.unpack("<HHI", data[i-8:i])
    z = zipfile.ZipFile(io.BytesIO(data[i:i+ln]))
    names = z.namelist()
    print(f"@{i-8} kind=0x{kind:04X} inst={vi>>4} len={ln} entries={len(names)}")
    for n in names:
        info = z.getinfo(n)
        print(f"    {n:52s} {info.file_size:6d} -> {info.compress_size:6d} m{info.compress_type}")
    i += 4
    seen += 1
