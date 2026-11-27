import sys, struct
sys.path.insert(0, "tools")
from cfb_list import CFB
c = CFB("samples/picture.doc")
wd = c.read("WordDocument"); tb = c.read("0Table") or c.read("1Table")
def u32(b,o): return struct.unpack_from("<I", b, o)[0]
fc_clx, lcb_clx = u32(wd, 0x1A2), u32(wd, 0x1A6)
clx = tb[fc_clx:fc_clx+lcb_clx]
i = 0
while clx[i] == 1:
    i += 1 + struct.unpack_from("<H", clx, i+1)[0]
lcb = u32(clx, i+1)
plc = clx[i+5:i+5+lcb]
n = (lcb - 4)//12
cps = [u32(plc, 4*k) for k in range(n+1)]
print("cps", cps, "pieces", n)
for k in range(n):
    pcd = plc[4*(n+1)+8*k : 4*(n+1)+8*k+8]
    fc = u32(pcd, 2)
    print(f"  piece {k}: cp {cps[k]}..{cps[k+1]} fc {fc:#x} comp {bool(fc>>30)}")
# chpx page
fcb = u32(wd, 0x9A+12*8); lcb2 = u32(wd, 0x9A+12*8+4)
plc2 = tb[fcb:fcb+lcb2]; n2 = (lcb2-4)//8; pn = u32(plc2, 4*(n2+1))
page = wd[pn*512:(pn+1)*512]
crun = page[511]
fcs = [u32(page, 4*k) for k in range(crun+1)]
print("chpx crun", crun, "fc0", hex(fcs[0]))
for k in range(crun):
    off = page[(crun+1)*4 + k]*2
    if off == 0:
        continue
    cb = page[off]
    g = page[off+1:off+1+cb]
    if any(g[j:j+2] == b"\x03\x6a" for j in range(len(g)-1)) or cb > 0:
        print(f"  run {k} fc {fcs[k]:#x}..{fcs[k+1]:#x} cb {cb} grpprl {g.hex(' ')}")
