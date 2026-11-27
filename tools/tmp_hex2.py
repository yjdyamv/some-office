import sys
sys.path.insert(0, "tools")
from cfb_list import CFB
d = CFB("samples/picture.doc").read("Data")
cont = d[0x44:0xD6]
assert len(cont) == 146, len(cont)
assert int.from_bytes(cont[4:8],'little') == 138
print("const SP_CONTAINER: [u8; 146] = [")
for i in range(0, 146, 12):
    print("    " + " ".join(f"0x{x:02X}," for x in cont[i:i+12]))
print("];")
print("uid", d[0xDE:0xEE].hex())
print("total check", 68+146+44+8+565, "lcb", int.from_bytes(d[0:4],'little'))
