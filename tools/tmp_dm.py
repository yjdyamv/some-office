p = "docs/doc.md"
s = open(p, encoding="utf-8").read()
add = open("tools/tmp_sec.md", encoding="utf-8").read()
i = s.index("## 15")
s = s[:i] + add + "\n" + s[i:]
open(p, "w", encoding="utf-8", newline="\n").write(s)

p = "docs/README.md"
s = open(p, encoding="utf-8").read()
old = "| 写入器不支持的细节 | `.doc` 不写表格/图片/样式表名；"
new = "| 写入器不支持的细节 | `.doc` 不写图片/样式表名（表格可以写：单元格标记 + `TDefTable`，Word 验证过）；"
assert old in s
s = s.replace(old, new, 1)
s = s.replace("| 无公式求值 | 只还原公式文本，不计算（缓存值来自文件） |",
              "| 无公式求值 | 只还原公式文本，不计算（缓存值来自文件）；写出公式后由 Excel 自己重算 |", 1)
open(p, "w", encoding="utf-8", newline="\n").write(s)
print("ok")
