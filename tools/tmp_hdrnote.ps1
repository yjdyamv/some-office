# Make a Word sample with a header, a footer, a footnote and an endnote, so the
# four stories' `ccp*` counts can be read off a file Word wrote itself.
$ErrorActionPreference = 'Continue'
$out = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\samples\ms\header-note.doc'))
$clsid = 'HKCU:\Software\Classes\CLSID\{000209FF-0000-0000-C000-000000000046}\LocalServer32'
$exe = '"C:\Program Files\Microsoft Office\Root\Office16\WINWORD.EXE" /Automation'
$backup = $null
if (Test-Path $clsid) { $backup = (Get-ItemProperty $clsid).'(default)' }
New-Item -Path $clsid -Force | Out-Null
Set-ItemProperty -Path $clsid -Name '(default)' -Value $exe
try {
    $wd = New-Object -ComObject Word.Application
    $wd.Visible = $false
    $wd.DisplayAlerts = 0
    $doc = $wd.Documents.Add()
    $doc.Content.Text = "Body one carries a footnote.`rBody two carries an endnote.`rThird paragraph.`r"
    $range = $doc.Paragraphs.Item(1).Range
    $range.Collapse(0) | Out-Null
    $doc.Footnotes.Add($range, "footnote text") | Out-Null
    $range = $doc.Paragraphs.Item(2).Range
    $range.Collapse(0) | Out-Null
    $doc.Endnotes.Add($range, "endnote text") | Out-Null
    $doc.Sections.Item(1).Headers.Item(1).Range.Text = "Draft header"
    $doc.Sections.Item(1).Footers.Item(1).Range.Text = "Draft footer"
    if (Test-Path $out) { Remove-Item $out -Force }
    $doc.SaveAs2($out, 0)
    $doc.Close($false)
    $wd.Quit()
    Write-Host "wrote $out"
} catch {
    Write-Host "ERR $($_.Exception.Message)"
} finally {
    if ($backup) { Set-ItemProperty -Path $clsid -Name '(default)' -Value $backup }
    Get-Process WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force
}
