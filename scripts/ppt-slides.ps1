# 指定プレゼン(名前完全一致)の各スライドのシェイプ名を列挙(usage: powershell.exe -File scripts/ppt-slides.ps1 "<name>")
param([string]$name)
$app = [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')
$p = $null
foreach ($x in @($app.Presentations)) { if ($x.Name -eq $name) { $p = $x } }
if (-not $p) { Write-Output "presentation not found: $name"; exit 1 }
foreach ($s in $p.Slides) {
  $names = @()
  foreach ($sh in $s.Shapes) { $names += $sh.Name }
  Write-Output ('slide {0} (id={1}): shapes={2} [{3}]' -f $s.SlideIndex, $s.SlideID, $s.Shapes.Count, ($names -join ', '))
}
