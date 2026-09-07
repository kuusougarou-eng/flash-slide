# 指定プレゼン(名前部分一致)のスライドを PNG に書き出す(レイアウト確定後の実描画。タスクペインの即時スナップショットとの比較用)
# usage: powershell.exe -File scripts/ppt-export-slide.ps1 "<presentation like>" <slideIndex> <out.png>
param([string]$nameLike, [int]$index, [string]$out)
$app = [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')
$p = $null
foreach ($x in @($app.Presentations)) { if ($x.Name -like "*$nameLike*") { $p = $x } }
if (-not $p) { Write-Output "presentation not found: $nameLike"; exit 1 }
$p.Slides.Item($index).Export($out, "PNG", 1280, 720)
Write-Output "exported slide $index -> $out"
