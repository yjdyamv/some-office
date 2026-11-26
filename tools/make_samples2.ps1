# Generates the second batch of samples: formulas, number formats, BIFF5,
# character formatting in .doc, and a picture in .ppt.
# Run with pwsh (PowerShell 7), not powershell.exe: this file is UTF-8 without BOM.
$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$samples = Join-Path (Split-Path -Parent $here) 'samples'
New-Item -ItemType Directory -Force -Path $samples | Out-Null

# ---------------------------------------------------------------- formulas.xls
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$wb = $xl.Workbooks.Add()
$ws = $wb.Worksheets.Item(1)
$ws.Name = 'F'

$ws.Range('A1').Value2 = 10
$ws.Range('A2').Value2 = 20
$ws.Range('A3').Value2 = 1234.5678
$ws.Range('A4').Value2 = 0.25
$ws.Range('A5').Value2 = [datetime]'2024-03-05'
$ws.Range('A6').Value2 = 45356
$ws.Range('A3').NumberFormat = '0.00'
$ws.Range('A4').NumberFormat = '0.00%'
$ws.Range('A5').NumberFormat = 'yyyy-mm-dd'
$ws.Range('A6').NumberFormat = '$#,##0.00'

$ws.Range('B1').Formula = '=A1+A2'
$ws.Range('B2').Formula = '=A1*A2'
$ws.Range('B3').Formula = '=SUM(A1:A2)'
$ws.Range('B4').Formula = '=A1/A2'
$ws.Range('B5').Formula = '=A1^2'
$ws.Range('B6').Formula = '=IF(A1>5,"big","small")'
$ws.Range('B7').Formula = '=AVERAGE(A1:A2)'
$ws.Range('B8').Formula = '=MAX(A1:A2)'
$ws.Range('B9').Formula = '=-A1'
$ws.Range('B10').Formula = '=A1&"x"'
$ws.Range('B11').Formula = '=A1>5'
$ws.Range('B12').Formula = '=ROUND(A3,1)'
$ws.Range('B13').Formula = '=SUM(A1:A2)*2'
$ws.Range('B14').Formula = '=ABS(A1-A2)'
$ws.Range('B15').Formula = '=A1<>A2'
$ws.Range('B16').Formula = '=SUM(A1,A2)'

$ws2 = $wb.Worksheets.Add()
$ws2.Name = 'G'
$ws2.Range('A1').Value2 = 7
$ws.Range('C1').Formula = "='G'!A1"
$ws.Range('C2').Formula = '=SUM(G!A1,3)'

$path = Join-Path $samples 'formulas.xls'
$wb.SaveAs($path, 56)          # xlExcel8
$wb.Close($false)

# ------------------------------------------- excel95-request.xls (Excel 95 request)
# NOTE: WPS ignores xlExcel9795 and still writes BIFF8 (0x0600), so this file is
# *not* a BIFF5 sample. The BIFF5 branch is covered by synthetic streams instead.
$wb = $xl.Workbooks.Add()
$ws = $wb.Worksheets.Item(1)
$ws.Name = 'Old'
$ws.Range('A1').Value2 = 'BIFF5-TEXT'
$ws.Range('A2').Value2 = 1234.5
$ws.Range('A3').Formula = '=A2*2'
$path5 = Join-Path $samples 'excel95-request.xls'
$wb.SaveAs($path5, 39)         # xlExcel9795 = Excel 95/97
$wb.Close($false)
$xl.Quit()
[void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)

# ---------------------------------------------------------- formatted.doc
$wd = New-Object -ComObject Word.Application
$wd.Visible = $false
$doc = $wd.Documents.Add()
$doc.Content.Text = ''
$r = $doc.Content
$r.InsertAfter("Bold red heading`r")
$r.Paragraphs.Item(1).Range.Font.Bold = $true
$r.Paragraphs.Item(1).Range.Font.Size = 18
$r.Paragraphs.Item(1).Range.Font.Color = 255   # wdColorRed
$r.InsertAfter("Italic underline run with mixed ")
$r.InsertAfter("bold tail")
$p2 = $doc.Paragraphs.Item(2).Range
$p2.Font.Italic = $true
$p2.Font.Underline = 1
$doc.Paragraphs.Item(2).Range.Words.Item(1).Font.Bold = $true
$r.InsertAfter("Centered paragraph`r")
$doc.Paragraphs.Item(3).Alignment = 1          # wdAlignParagraphCenter
$r.InsertAfter("Indented paragraph`r")
$doc.Paragraphs.Item(4).LeftIndent = 36
$doc.Paragraphs.Item(4).FirstLineIndent = 18
$r.InsertAfter("Plain tail paragraph`r")
$pathd = Join-Path $samples 'formatted.doc'
$doc.SaveAs2($pathd, 0)
$doc.Close($false)
$wd.Quit()
[void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($wd)

# ------------------------------------------------------------- image.ppt
$pp = New-Object -ComObject PowerPoint.Application
$pres = $pp.Presentations.Add()
$slide = $pres.Slides.Add(1, 1)   # ppLayoutTitle
$slide.Shapes.Item(1).TextFrame.TextRange.Text = 'IMAGE-PPT-001'
$slide.Shapes.Item(2).TextFrame.TextRange.Text = 'A picture below'
$png = Join-Path $samples 'pixel.png'
$slide.Shapes.AddPicture($png, 0, -1, 60, 200, 120, 120) | Out-Null
$pathp = Join-Path $samples 'image.ppt'
$pres.SaveAs($pathp, 1)
$pres.Close()
$pp.Quit()
[void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($pp)

Get-ChildItem $samples | Select-Object Name, Length | Format-Table -AutoSize
