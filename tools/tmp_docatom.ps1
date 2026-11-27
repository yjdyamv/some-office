$ErrorActionPreference = "Stop"
$key = "HKCU:\Software\Classes\CLSID\{91493441-5A91-11CF-8700-00AA0060263B}\LocalServer32"
$old = (Get-ItemProperty -Path $key -ErrorAction SilentlyContinue)."(default)"
function Save($pres, $name) {
    $path = (Join-Path (Get-Location) "samples/ms/$name")
    if (Test-Path $path) { Remove-Item $path -Force }
    $pres.SaveAs($path, 1)
    "$name  slides=$($pres.Slides.Count) layouts=$($pres.SlideMaster.CustomLayouts.Count) notes=$($pres.HasNotesMaster) handout=$($pres.HandoutMaster.Name)"
    $pres.Close()
}
try {
    Set-ItemProperty -Path $key -Name "(default)" -Value '"C:\Program Files\Microsoft Office\Root\Office16\POWERPNT.EXE" /automation'
    $ppt = New-Object -ComObject PowerPoint.Application
    $ppt.Visible = -1
    # baseline
    $p = $ppt.Presentations.Add(1); Save $p "atom-base.ppt"
    # zoom
    $p = $ppt.Presentations.Add(1)
    $p.Windows.Item(1).View.Zoom = 66
    Save $p "atom-zoom.ppt"
    # notes master present
    $p = $ppt.Presentations.Add(1)
    $null = $p.NotesMaster
    Save $p "atom-notes.ppt"
    # fewer layouts
    $p = $ppt.Presentations.Add(1)
    while ($p.SlideMaster.CustomLayouts.Count -gt 2) { $p.SlideMaster.CustomLayouts.Item(3).Delete() }
    Save $p "atom-fewlayouts.ppt"
    # notes size changed
    $p = $ppt.Presentations.Add(1)
    $p.PageSetup.NotesOrientation = 2
    Save $p "atom-notessize.ppt"
    $ppt.Quit()
} finally {
    Set-ItemProperty -Path $key -Name "(default)" -Value $old
    Get-Process POWERPNT -ErrorAction SilentlyContinue | Stop-Process -Force
}
