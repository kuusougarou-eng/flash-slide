# 検証用: 指定 pptx を既存の PowerPoint インスタンスで開く(usage: powershell.exe -File scripts/ppt-open.ps1 <path>)
param([string]$path)
$app = [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')
$full = (Resolve-Path $path).Path
$p = $app.Presentations.Open($full, 0, 0, -1)
Write-Output ("opened {0} slides={1} size={2}x{3}" -f $p.Name, $p.Slides.Count, $p.PageSetup.SlideWidth, $p.PageSetup.SlideHeight)
