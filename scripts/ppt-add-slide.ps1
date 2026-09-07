# 検証用: 指定プレゼン(名前部分一致)の末尾に、名前部分一致のレイアウトでスライドを追加して選択する
# usage: powershell.exe -File scripts/ppt-add-slide.ps1 "<presentation like>" "<layout name like>"
param([string]$nameLike, [string]$layoutLike = "Blank")
$app = [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')
$p = $null
foreach ($x in @($app.Presentations)) { if ($x.Name -like "*$nameLike*") { $p = $x } }
if (-not $p) { Write-Output "presentation not found: $nameLike"; exit 1 }
$lay = $null
foreach ($l in $p.SlideMaster.CustomLayouts) { if ($l.Name -like "*$layoutLike*") { $lay = $l; break } }
if (-not $lay) { Write-Output "layout not found: $layoutLike"; foreach ($l in $p.SlideMaster.CustomLayouts) { Write-Output ("  " + $l.Name) }; exit 1 }
$s = $p.Slides.AddSlide($p.Slides.Count + 1, $lay)
$p.Windows.Item(1).View.GotoSlide($s.SlideIndex)
Write-Output ("added slide {0} layout='{1}' shapes={2} and selected" -f $s.SlideIndex, $lay.Name, $s.Shapes.Count)
