$ErrorActionPreference = "Stop"
$key = "HKCU:\Software\Classes\CLSID\{000209FF-0000-0000-C000-000000000046}\LocalServer32"
$old = (Get-ItemProperty -Path $key -ErrorAction SilentlyContinue)."(default)"
Remove-Item -Path $key -Force -ErrorAction SilentlyContinue
Remove-Item -Path 'HKCU:\Software\Microsoft\Office\16.0\Word\Resiliency' -Recurse -Force -ErrorAction SilentlyContinue
$full = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot "..\samples\written-fields.doc"))
$wd = New-Object -ComObject Word.Application
try {
    Write-Output ("app: " + $wd.Path)
    $wd.Visible = $false
    $doc = $wd.Documents.Open($full, $false, $true)
    $header = $doc.Sections.Item(1).Headers.Item(1)
    $fields = $header.Range.Fields
    Write-Output ("header text: [{0}]" -f $header.Range.Text)
    Write-Output ("fields: {0}" -f $fields.Count)
    for ($i = 1; $i -le $fields.Count; $i++) {
        Write-Output ("  field {0}: code=[{1}] result=[{2}]" -f $i, $fields.Item($i).Code.Text, $fields.Item($i).Result.Text)
    }
    $fields.Update() | Out-Null
    Write-Output ("after update: [{0}]" -f $header.Range.Text)
    $doc.Close($false)
} finally {
    $wd.Quit()
    Start-Sleep -Milliseconds 300
    Get-Process WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force
    if ($old) { New-Item -Path $key -Force | Out-Null; Set-ItemProperty -Path $key -Name "(default)" -Value $old }
}
