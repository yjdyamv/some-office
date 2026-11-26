# Header / footer / footnote sample for the .doc story reader.
$ErrorActionPreference = "Continue"
$out = Join-Path $PSScriptRoot "..\samples"
$path = (Resolve-Path $out).Path + "\stories.doc"

$wd = New-Object -ComObject Word.Application
$wd.Visible = $false
$doc = $wd.Documents.Add()
$doc.Content.Text = "Body paragraph one.`rBody paragraph two.`r"

# A header, a footer and a footnote, each with a distinctive marker.
$section = $doc.Sections.Item(1)
$section.Headers.Item(1).Range.Text = "HEADER-MARK-11"
$section.Footers.Item(1).Range.Text = "FOOTER-MARK-22"

$rng = $doc.Content
$rng.Collapse(0) # wdCollapseEnd
$rng.MoveStart(-1)
$rng.Footnotes.Add($rng, "FOOTNOTE-MARK-33")
$doc.SaveAs2($path, 0)
Write-Host ("wrote {0} ({1} bytes)" -f $path, (Get-Item $path).Length)
try { $doc.Close(0); $wd.Quit() } catch { Write-Host "close/quit warning: $($_.Exception.Message)" }
