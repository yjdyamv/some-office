# Password-protected samples: RC4 CryptoAPI encryption, the Office 97-2003 way.
$ErrorActionPreference = "Continue"
$out = (Resolve-Path (Join-Path $PSScriptRoot "..\samples")).Path
$pw = "secret123"

$wd = New-Object -ComObject Word.Application
$wd.Visible = $false
$doc = $wd.Documents.Add()
$doc.Content.Text = "ENCRYPTED-DOC-777`rSecond secret line.`r"
$doc.Password = $pw
$doc.SaveAs2("$out\encrypted.doc", 0)
Write-Host ("doc {0} bytes" -f (Get-Item "$out\encrypted.doc").Length)
try { $doc.Close(0); $wd.Quit() } catch { Write-Host "word quit: $($_.Exception.Message)" }

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$wb = $xl.Workbooks.Add()
$ws = $wb.Worksheets.Item(1)
$ws.Name = "Secret"
$ws.Cells.Item(1,1) = "ENCRYPTED-XLS-888"
$ws.Cells.Item(2,1) = 1234.5
$ws.Cells.Item(3,1) = "=A2*2"
$wb.Password = $pw
$wb.SaveAs("$out\encrypted.xls", 56)
Write-Host ("xls {0} bytes" -f (Get-Item "$out\encrypted.xls").Length)
try { $wb.Close(0); $xl.Quit() } catch { Write-Host "excel quit: $($_.Exception.Message)" }

$pp = New-Object -ComObject PowerPoint.Application
$pres = $pp.Presentations.Add()
$slide = $pres.Slides.Add(1, 1)
$slide.Shapes.Title.TextFrame.TextRange.Text = "ENCRYPTED-PPT-999"
$slide.Shapes.Item(2).TextFrame.TextRange.Text = "secret body"
$pres.Password = $pw
$pres.SaveAs("$out\encrypted.ppt", 1)
Write-Host ("ppt {0} bytes" -f (Get-Item "$out\encrypted.ppt").Length)
try { $pres.Close(); $pp.Quit() } catch { Write-Host "ppt quit: $($_.Exception.Message)" }
