# Flash Slide が生成したスライド(シェイプ名がすべて FS_* / 画像)だけを削除する安全弁付きの削除
# usage: powershell.exe -File scripts/ppt-delete-slide.ps1 "<presentation name>" <slideIndex>
param([string]$name, [int]$index)
$app = [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')
$p = $null
foreach ($x in @($app.Presentations)) { if ($x.Name -eq $name) { $p = $x } }
if (-not $p) { Write-Output "presentation not found: $name"; exit 1 }
$s = $p.Slides.Item($index)
$foreign = @()
foreach ($sh in $s.Shapes) { if (-not ($sh.Name -like 'FS_*' -or $sh.Name -like 'Picture*' -or $sh.Name -like '図*')) { $foreign += $sh.Name } }
if ($foreign.Count -gt 0) { Write-Output ("refused: slide {0} has non-FlashSlide shapes: {1}" -f $index, ($foreign -join ', ')); exit 2 }
$s.Delete()
Write-Output ("deleted slide {0} of {1} (now {2} slides)" -f $index, $p.Name, $p.Slides.Count)
