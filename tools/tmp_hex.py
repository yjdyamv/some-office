import sys
sys.path.insert(0, "tools")
from cfb_list import CFB
d = CFB("samples/picture.doc").read("Data")
def rust(name, b):
    print(f"const {name}: [u8; {len(b)}] = [")
    for i in range(0, len(b), 12):
        print("    " + " ".join(f"0x{x:02X}," for x in b[i:i+12]))
    print("];")
rust("PICF_TEMPLATE", d[0:0x44])
rust("SP_CONTAINER", d[0x44:0x44+138])
rust("BLIP_UID", d[0xDE:0xDE+16])
print("picf lcb", int.from_bytes(d[0:4],'little'), "cbhdr", int.from_bytes(d[4:6],'little'))
print("dxaGoal", int.from_bytes(d[0x1c:0x1e],'little'), "dyaGoal", int.from_bytes(d[0x1e:0x20],'little'))
print("fbse at 0xd6 len", int.from_bytes(d[0xd8:0xdc],'little'), "size", int.from_bytes(d[0xf0:0xf4],'little'))
print("blip at 0x102", d[0x102:0x10a].hex(' '), "len", int.from_bytes(d[0x106:0x10a],'little'))
