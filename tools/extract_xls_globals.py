"""Extract the globals block of a real Excel-written workbook as a template.

`crates/office-xls/templates/globals.bin` holds everything Excel writes before the
first sheet substream, *minus* the parts that depend on the workbook content:

  * `SST` (0x00FC) and `EXTSST` (0x00FF) — our writer emits its own strings,
  * `BOUNDSHEET` (0x0085) — our writer emits its own sheet list.

The trailing `EOF` (0x000A) is kept, so the writer inserts `SST`, `BOUNDSHEET` and
`EOF` itself.  Why a template at all is explained in `docs/xls.md` §15.2: our
hand-built globals are refused by Excel even though every individual record was
proven acceptable — something absent is required, and this sidesteps finding it.
"""

import sys

sys.path.insert(0, "tools")
from cfb_list import CFB

DROP = {0x00FC, 0x00FF, 0x0085}


def u16(b, i):
    return int.from_bytes(b[i:i + 2], "little")


d = CFB("samples/ms/one.xls").read("Workbook")
out, o, sheets = [], 0, 0
while o + 4 <= len(d):
    k, l = u16(d, o), u16(d, o + 2)
    if k == 0x0809 and u16(d, o + 6) == 0x0010:
        sheets += 1
        if sheets == 1:
            break  # end of the globals block
    if k == 0x000A and l == 0:
        out.append(d[o:o + 4])
        break
    if k not in DROP:
        out.append(d[o:o + 4 + l])
    o += 4 + l

data = b"".join(out)
assert data.endswith(b"\x0a\x00\x00\x00"), data[-8:]
open("crates/office-xls/templates/globals.bin", "wb").write(data)
print(f"globals template: {len(data)} bytes")
