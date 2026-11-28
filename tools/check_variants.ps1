# Read the three kinds of header back out of a document with Word itself.
param([string]$Path = "samples/written-variants.doc")
$ErrorActionPreference = "Stop"
$full = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot "..\$Path"))
$key = "HKCU:\Software\Classes\CLSID\{000209FF-0000-0000-C000-000000000046}\LocalServer32"
$old = (Get-ItemProperty -Path $key -ErrorAction SilentlyContinue)."(default)"
Remove-Item -Path $key -Force -ErrorAction SilentlyContinue
Remove-Item -Path 'HKCU:\Software\Microsoft\Office\16.0\Word\Resiliency' -Recurse -Force -ErrorAction SilentlyContinue
$word = New-Object -ComObject Word.Application
try {
    Write-Output "app: $($word.Path)"
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $doc = $word.Documents.Open($full, $false, $true)
    $section = $doc.Sections.Item(1)
    Write-Output ("first page header: {0}" -f $section.PageSetup.DifferentFirstPageHeaderFooter)
    Write-Output ("odd and even:      {0}" -f $section.PageSetup.OddAndEvenPagesHeaderFooter)
    Write-Output ("primary header: [{0}]" -f $section.Headers.Item(1).Range.Text.TrimEnd("`r", "`a"))
    Write-Output ("first header:   [{0}]" -f $section.Headers.Item(2).Range.Text.TrimEnd("`r", "`a"))
    Write-Output ("even header:    [{0}]" -f $section.Headers.Item(3).Range.Text.TrimEnd("`r", "`a"))
    $doc.Close(0)
} finally {
    $word.Quit()
    Start-Sleep -Milliseconds 300
    Get-Process WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force
    if ($old) { New-Item -Path $key -Force | Out-Null; Set-ItemProperty -Path $key -Name "(default)" -Value $old }
}
