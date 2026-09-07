# Office.js の slideId("<SlideID>#<...>")から該当スライドを探して PNG 書き出し(タスクペインのログと突き合わせる用)
# usage: powershell.exe -File scripts/ppt-export-by-id.ps1 "<presentation like>" "<slideId>" <out.png>
param([string]$nameLike, [string]$slideId, [string]$out)
$app = [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')
$p = $null
foreach ($x in @($app.Presentations)) { if ($x.Name -like "*$nameLike*") { $p = $x } }
if (-not $p) { Write-Output "presentation not found: $nameLike"; exit 1 }
$id = [int64]($slideId.Split('#')[0])
foreach ($s in $p.Slides) {
  if ([int64]$s.SlideID -eq $id) { $s.Export($out, "PNG", 1280, 720); Write-Output ("exported slide {0} (id {1}) -> {2}" -f $s.SlideIndex, $id, $out); exit 0 }
}
Write-Output "slide id not found: $id"; exit 2
