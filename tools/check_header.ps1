# Check the header and footer of a .doc written by office-doc.
$ErrorActionPreference = 'Continue'
$path = (Resolve-Path $args[0]).Path
$clsid = 'HKCU:\Software\Classes\CLSID\{000209FF-0000-0000-C000-000000000046}\LocalServer32'
$exe = '"C:\Program Files\Microsoft Office\Root\Office16\WINWORD.EXE" /Automation'
$backup = $null
if (Test-Path $clsid) { $backup = (Get-ItemProperty $clsid).'(default)' }
New-Item -Path $clsid -Force | Out-Null
Set-ItemProperty -Path $clsid -Name '(default)' -Value $exe
try {
    $wd = New-Object -ComObject Word.Application
    Write-Host ("app: {0}" -f $wd.Path)
    $wd.Visible = $false
    $wd.DisplayAlerts = 0
    $doc = $wd.Documents.Open($path, $false, $true)
    Write-Host ("sections={0} pages={1}" -f $doc.Sections.Count, $doc.ComputeStatistics(2))
    foreach ($p in $doc.Paragraphs) {
        Write-Host ("  body: {0}" -f $p.Range.Text.Trim())
    }
    $i = 0
    foreach ($s in $doc.Sections) {
        $i++
        foreach ($hf in @(
            @('header', $s.Headers.Item(1)),
            @('footer', $s.Footers.Item(1)))) {
            $text = $hf[1].Range.Text
            Write-Host ("  section {0} {1}: [{2}]" -f $i, $hf[0], ($text -replace "`r", '\r').Trim())
        }
    }
    $doc.Close($false)
    $wd.Quit()
} catch {
    Write-Host "ERR $($_.Exception.Message)"
} finally {
    if ($backup) { Set-ItemProperty -Path $clsid -Name '(default)' -Value $backup }
    Get-Process WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force
}
