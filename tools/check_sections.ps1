# Read the sections back through Word itself.
param([string]$Path = "samples/written-sections.doc")
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
    $doc = $wd.Documents.Open($full, $false, $true)
    Write-Output ("sections: {0}" -f $doc.Sections.Count)
    for ($i = 1; $i -le $doc.Sections.Count; $i++) {
        $ps = $doc.Sections.Item($i).PageSetup
        Write-Output ("  {0}: {1}x{2} pt orient={3} text=[{4}]" -f $i, [math]::Round($ps.PageWidth,1), [math]::Round($ps.PageHeight,1), $ps.Orientation, ($doc.Sections.Item($i).Range.Text -replace "`r","\n" -replace "`f","<sect>"))
    }
    $doc.Close(0)
} finally {
    $wd.Quit()
    Get-Process WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force
    if ($old) { New-Item -Path $key -Force | Out-Null; Set-ItemProperty -Path $key -Name "(default)" -Value $old }
}
