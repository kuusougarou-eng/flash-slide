# 検証用: 名前部分一致のプレゼンを閉じる。保存はしない(必要なら先に backup へコピー)
# usage: powershell.exe -File scripts/ppt-close.ps1 "<name like>" [backupPath]
param([string]$nameLike, [string]$backup = "")
$app = [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')
$p = $null
foreach ($x in @($app.Presentations)) { if ($x.Name -like "*$nameLike*") { $p = $x } }
if (-not $p) { Write-Output "presentation not found: $nameLike"; exit 1 }
if ($backup) { $p.SaveCopyAs($backup); Write-Output "backup -> $backup" }
$name = $p.Name
$p.Saved = -1
$p.Close()
Write-Output "closed $name (not saved)"
