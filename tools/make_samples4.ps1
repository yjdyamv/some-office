# Build a .doc that contains an inline picture, plus a text box and a table.
$ErrorActionPreference = 'Stop'
$out = Join-Path $PSScriptRoot '..\samples'
$out = (Resolve-Path $out).Path

$wd = New-Object -ComObject Word.Application
$wd.Visible = $false
$doc = $wd.Documents.Add()
$doc.Content.Text = "PICTURE-DOC-555`rBefore image.`r"

# Inline picture at the end of the document (wdCollapseEnd = 0).
$end = $doc.Content
$end.Collapse(0)
$shape = $doc.InlineShapes.AddPicture((Join-Path $out 'pixel.png'), $false, $true, $end)

$end2 = $doc.Content
$end2.Collapse(0)
$end2.InsertAfter("After image.`r")

$path = Join-Path $out 'picture.doc'
$doc.SaveAs2($path, 0)
Write-Host "wrote $path"
try { $doc.Close(0) } catch {}
try { $wd.Quit() } catch {}
