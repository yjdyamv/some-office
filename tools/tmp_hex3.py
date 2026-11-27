import sys
sys.path.insert(0, "tools")
from cfb_list import CFB
d = CFB("samples/picture.doc").read("Data")
for off in range(0xC8, 0x110, 8):
    print(f"{off:04x}: " + " ".join(f"{b:02x}" for b in d[off:off+8]))
