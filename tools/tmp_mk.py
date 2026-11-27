import sys, os
sys.path.insert(0, "tools")
from cfb_list import CFB
d = CFB("samples/picture.doc").read("Data")
os.makedirs("crates/office-doc/templates", exist_ok=True)
open("crates/office-doc/templates/picture-data.bin", "wb").write(d)
print("wrote", len(d), "bytes")
# locate the FBSE and BLIP by scanning, to double-check the plan
import struct
pos = 0x44
while pos + 8 <= len(d):
    ver_inst, kind, ln = struct.unpack_from("<HHI", d, pos)
    print(f"  {pos:#06x} ver_inst {ver_inst:#06x} kind {kind:#06x} len {ln}")
    if ver_inst & 0xF == 0xF:
        pos += 8
        continue
    pos += 8 + ln
    if pos > len(d):
        break
