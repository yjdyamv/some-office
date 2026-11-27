param([string]$Out = "")

# Author a deck with speaker notes on every slide, saved by PowerPoint itself.
$here = (Get-Location).Path
if ($Out -eq "") { $Out = Join-Path $here "samples\ms\notes2.ppt" }
$Out = [System.IO.Path]::GetFullPath($Out)

$app = New-Object -ComObject PowerPoint.Application
try {
    $pres = $app.Presentations.Add()
    while ($pres.Slides.Count -gt 0) { $pres.Slides.Item(1).Delete() }
    foreach ($i in 1..3) {
        $slide = $pres.Slides.Add($i, 2)
        $slide.Shapes.Item(1).TextFrame.TextRange.Text = "NOTE-TITLE-$i"
        $slide.Shapes.Item(2).TextFrame.TextRange.Text = "note body $i"
        $slide.NotesPage.Shapes.Item(2).TextFrame.TextRange.Text = "speaker note for slide $i"
    }
    try {
        $pres.SaveAs($Out, 1)
        Write-Output "saved $Out"
    } catch {
        Write-Output ("save failed: " + $_.Exception.Message)
    }
    Write-Output ("path: " + $app.Path)
    Write-Output ("slides: " + $pres.Slides.Count)
} finally {
    $pres = $null
    $app.Quit()
}
