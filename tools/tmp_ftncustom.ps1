# Footnotes: one custom-marked ("*"), one auto-numbered -- enough to see FRD.
param([string]$Path = "samples/ms/footnote-custom.doc")
$ErrorActionPreference = "Stop"
$key = "HKCU:\Software\Classes\CLSID\{000209FF-0000-0000-C000-000000000046}\LocalServer32"
$old = (Get-ItemProperty -Path $key -ErrorAction SilentlyContinue)."(default)"
Remove-Item -Path $key -Force -ErrorAction SilentlyContinue
Remove-Item -Path "HKCU:\Software\Microsoft\Office\16.0\Word\Resiliency" -Recurse -Force -ErrorAction SilentlyContinue
$full = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot "..\$Path"))
$wd = New-Object -ComObject Word.Application
try {
    Write-Output ("app: " + $wd.Path)
    $wd.Visible = $false
    $doc = $wd.Documents.Add()
    $doc.Content.Text = "Custom marked note here.`rAuto numbered note here.`r"
    $r2 = $doc.Paragraphs.Item(2).Range
    $r2.End = $r2.End - 1
    $doc.Footnotes.Add($r2, "", "auto note")
    $r1 = $doc.Paragraphs.Item(1).Range
    $r1.End = $r1.End - 1
    $doc.Footnotes.Add($r1, "*", "star note")
    $doc.SaveAs2($full, 0)
    foreach ($f in $doc.Footnotes) {
        Write-Output ("note {0}: ref='{1}' text='{2}'" -f $f.Index, $f.Reference.Text, ($f.Range.Text -replace "`r","\n"))
    }
    $doc.Close(0)
} finally {
    $wd.Quit()
    Get-Process WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force
    if ($old) { New-Item -Path $key -Force | Out-Null; Set-ItemProperty -Path $key -Name "(default)" -Value $old }
}
