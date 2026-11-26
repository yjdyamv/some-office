# Samples that push the readers into the paths the simple ones never reach:
#   long.doc    -> several pieces, 8-bit (fCompressed) piece table
#   notes.ppt   -> speaker notes (SlideListWithText instance 2)
#   longstr.xls -> SST > 8224 bytes (CONTINUE records), merged cells, 3 sheets
param([string]$Out = "$PSScriptRoot\..\samples")
$Out = (Resolve-Path -LiteralPath $Out).Path
$ErrorActionPreference = "Stop"

# ---- Word: ~3 pages of ASCII, some bold, to force multiple pieces ----------
$wd = New-Object -ComObject Word.Application
$wd.Visible = $false
$doc = $wd.Documents.Add()
$doc.Content.Text = ""
for ($i = 1; $i -le 60; $i++) {
    $r = $doc.Content
    $r.Collapse(0)
    $r.InsertAfter("Paragraph $i of the long document. " + ("lorem ipsum dolor sit amet " * 3) + "`r")
    if ($i % 7 -eq 0) {
        $r2 = $doc.Content
        $r2.Collapse(0)
        $r2.MoveStart(-1) -or $null
        $doc.Paragraphs.Item($doc.Paragraphs.Count).Range.Font.Bold = $true
    }
}
$path = Join-Path $Out "long.doc"
$doc.SaveAs2($path, 0)
$doc.Close(0)
try { $wd.Quit() } catch {}
Write-Host "wrote $path"

# ---- PowerPoint: two slides with notes ------------------------------------
$pp = New-Object -ComObject PowerPoint.Application
$pres = $pp.Presentations.Add(0)
foreach ($n in 1..2) {
    $slide = $pres.Slides.Add($n, 2)          # ppLayoutText
    $slide.Shapes.Item(1).TextFrame.TextRange.Text = "NOTES-SLIDE-$n"
    $slide.Shapes.Item(2).TextFrame.TextRange.Text = "Body of slide $n"
    $slide.NotesPage.Shapes.Item(2).TextFrame.TextRange.Text = "SPEAKER-NOTE-$n hidden text"
}
$path = Join-Path $Out "notes.ppt"
$pres.SaveAs($path, 1)
$pres.Close()
try { $pp.Quit() } catch {}
Write-Host "wrote $path"

# ---- Excel: >8224-byte string (SST CONTINUE), merged cells, 3 sheets ------
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$wb = $xl.Workbooks.Add()
$s1 = $wb.Worksheets.Item(1)
$s1.Name = "Long"
$long = ("ABCDEFGHIJ" * 900)          # 9000 chars -> needs CONTINUE
$s1.Range("A1").Value2 = $long
$s1.Range("A2").Value2 = "tail-after-continue"
$s1.Range("A4:B5").Merge()
$s1.Range("A4").Value2 = "merged"
$s2 = $wb.Worksheets.Add()
$s2.Name = "Second"
$s2.Range("A1").Value2 = 1
$s2.Range("A2").Value2 = "=A1*2"
$s3 = $wb.Worksheets.Add()
$s3.Name = "Third"
$s3.Range("A1").Value2 = "third sheet"
$path = Join-Path $Out "longstr.xls"
$wb.SaveAs($path, 56)
$wb.Close($false)
try { $xl.Quit() } catch {}
Write-Host "wrote $path"
