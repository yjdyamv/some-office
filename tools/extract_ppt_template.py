"""Extract the `PowerPoint Document` stream of a blank deck as our built-in template.

`office_ppt::write::write_presentation` clones the shape tree of the first slide
in this stream, so the file has to be written by Microsoft PowerPoint (its
Escher shape tree is what PowerPoint validates most strictly).

    uv run --no-project python tools/extract_ppt_template.py \
        samples/ms/minimal.ppt crates/office-ppt/templates/minimal.bin

Trimming the stream down to "Document + MainMaster + one layout + slide" was
tried and rejected by PowerPoint even when every persist reference was made
consistent again (`docs/ppt.md` §9.4), so the whole stream is kept.
"""

import sys

sys.path.insert(0, "tools")
from cfb_list import CFB  # noqa: E402


def main():
    src, dst = sys.argv[1], sys.argv[2]
    stream = CFB(src).read("PowerPoint Document")
    with open(dst, "wb") as fh:
        fh.write(stream)
    print(f"{src} -> {dst}: {len(stream)} bytes")


if __name__ == "__main__":
    main()
