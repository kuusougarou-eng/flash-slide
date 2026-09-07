# 指定プレゼン(名前部分一致)のウィンドウを前面化し、必要ならスライドを選択
# usage: powershell.exe -File scripts/ppt-activate.ps1 "<name like>" [slideIndex]
param([string]$nameLike = "PowerPoint add-in", [int]$slide = 0)
$app = [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')
$p = $null
foreach ($x in @($app.Presentations)) { if ($x.Name -like "*$nameLike*") { $p = $x } }
if (-not $p) { Write-Output "presentation not found: $nameLike"; exit 1 }
$w = $p.Windows.Item(1)
if (-not $env:PPT_NO_ACTIVATE) { $w.Activate() }
if (-not $env:PPT_NO_ACTIVATE) { $app.Activate() }
if ($slide -gt 0) { $w.View.GotoSlide($slide) }
Write-Output ("activated {0} (slides={1}, selected={2})" -f $p.Name, $p.Slides.Count, $w.View.Slide.SlideIndex)
