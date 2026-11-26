$ErrorActionPreference = "Stop"
$dir = "C:\Users\admin\Desktop\some-office\samples"

# ---------------- Word .doc (Word 97-2003 binary) ----------------
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
$doc = $word.Documents.Add()
$doc.Content.Text = "HELLO-DOC-001`r`nSecond paragraph with ASCII only.`r`nUnicode: 中文测试 Ω≈ç√`r`nEND-MARKER-42`r`n"
$doc.SaveAs2([ref]($dir + "\hello.doc"), [ref]0)   # wdFormatDocument97
$doc.Close([ref]0)
$word.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null

# ---------------- Excel .xls (BIFF8) ----------------
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$wb = $xl.Workbooks.Add()
$ws = $wb.Worksheets.Item(1)
$ws.Name = "Data"
$ws.Range("A1").Value2 = "HELLO-XLS-001"
$ws.Range("A2").Value2 = "中文单元格"
$ws.Range("A3").Value2 = 1234.5
$ws.Range("A4").Value2 = 42
$ws.Range("A5").Value2 = -7
$ws.Range("B1").Value2 = 1
$ws.Range("B2").Value2 = 2
$ws.Range("B3").Formula = "=SUM(B1:B2)"
$ws.Range("C1").Value2 = [datetime]"2024-03-05"
$ws2 = $wb.Worksheets.Add()
$ws2.Name = "Sheet2"
$ws2.Range("A1").Value2 = "second-sheet"
$wb.SaveAs($dir + "\hello.xls", 56)   # xlExcel8
$wb.Close($false)
$xl.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null

# ---------------- PowerPoint .ppt (PowerPoint 97-2003) ----------------
$pp = New-Object -ComObject PowerPoint.Application
$pres = $pp.Presentations.Add()
$s1 = $pres.Slides.Add(1, 1)   # ppLayoutTitle
$s1.Shapes.Item(1).TextFrame.TextRange.Text = "HELLO-PPT-001"
$s1.Shapes.Item(2).TextFrame.TextRange.Text = "Slide one body text"
$s2 = $pres.Slides.Add(2, 2)   # ppLayoutText
$s2.Shapes.Item(1).TextFrame.TextRange.Text = "Second slide"
$s2.Shapes.Item(2).TextFrame.TextRange.Text = "中文幻灯片内容"
$pres.SaveAs($dir + "\hello.ppt", 1)   # ppSaveAsPresentation
$pres.Close()
$pp.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($pp) | Out-Null

Get-ChildItem $dir | Select-Object Name, Length | Format-Table
