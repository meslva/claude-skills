# PPTX 장표를 PNG로 렌더링한다 (Windows + PowerPoint 필요).
# 사용: powershell -ExecutionPolicy Bypass -File render_pptx.ps1 -Path deck.pptx -OutDir _render [-Width 1600] [-Slides "1,3,5"]
param(
    [Parameter(Mandatory = $true)][string]$Path,
    [Parameter(Mandatory = $true)][string]$OutDir,
    [int]$Width = 1600,
    [string]$Slides = ""
)
$ErrorActionPreference = "Stop"
$full = (Resolve-Path $Path).Path
New-Item -ItemType Directory -Force -Path $OutDir | Out-Null
$out = (Resolve-Path $OutDir).Path
$app = New-Object -ComObject PowerPoint.Application
try {
    # Open(FileName, ReadOnly, Untitled, WithWindow)
    $pres = $app.Presentations.Open($full, $true, $false, $false)
    $ratio = $pres.PageSetup.SlideHeight / $pres.PageSetup.SlideWidth
    $height = [int]($Width * $ratio)
    $want = @()
    if ($Slides -ne "") { $want = $Slides.Split(",") | ForEach-Object { [int]$_ } }
    $n = 0
    foreach ($s in $pres.Slides) {
        if ($want.Count -gt 0 -and -not ($want -contains $s.SlideIndex)) { continue }
        $file = Join-Path $out ("slide_{0:D2}.png" -f $s.SlideIndex)
        $s.Export($file, "PNG", $Width, $height)
        $n++
    }
    $pres.Close()
    Write-Output "rendered $n slide(s) to $out"
}
finally {
    $app.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null
}
