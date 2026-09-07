# ファイルを PowerPoint で(ウィンドウを出さずに)開き、全スライドを PNG に書き出して閉じる
# usage: powershell.exe -File scripts/ppt-export-file.ps1 <deck.pptx> <outDir> [width]
param([string]$file, [string]$outDir, [int]$width = 1920)
$file = (Resolve-Path $file).Path
New-Item -ItemType Directory -Force $outDir | Out-Null
$outDir = (Resolve-Path $outDir).Path
try { $app = [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application') } catch { $app = New-Object -ComObject PowerPoint.Application }
$p = $app.Presentations.Open($file, $true, $false, $false)  # ReadOnly, Untitled=false, WithWindow=false
$h = [int]($width * $p.PageSetup.SlideHeight / $p.PageSetup.SlideWidth)
$n = $p.Slides.Count
for ($i = 1; $i -le $n; $i++) {
  $out = Join-Path $outDir ("{0:d2}.png" -f $i)
  $p.Slides.Item($i).Export($out, "PNG", $width, $h)
}
$p.Close()
Write-Output "exported $n slides -> $outDir"
