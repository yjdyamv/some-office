import re
B = chr(92)
p = "crates/office-doc/src/write.rs"
s = open(p, encoding="utf-8").read()
s = s.replace(chr(13), "")  # drop the stray CRs
s = s.replace(chr(7), "")   # and the stray BEL if any

s = re.sub(r"    /// The character that ends the paragraph: `\n` for a normal one, ``\n",
           "    /// The character that ends the paragraph: `" + B + "r` for a normal one, `" + B + "x07`\n",
           s)
s = s.replace("        Paragraph { mark: '\n', ..Paragraph::raw(text) }",
              "        Paragraph { mark: '" + B + "r', ..Paragraph::raw(text) }")
open(p, "w", encoding="utf-8", newline="\n").write(s)
print(repr(s[s.index("pub struct Paragraph"):s.index("pub struct Paragraph")+420]))
