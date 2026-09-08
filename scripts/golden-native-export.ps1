param([string]$Root='debug/golden-native',[string]$Source='compiled',[string]$Out='rendered')
$ErrorActionPreference='Stop'
$rootPath=(Resolve-Path -LiteralPath $Root).Path
try {$app=[Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')} catch {$app=New-Object -ComObject PowerPoint.Application}
$outPath=Join-Path $rootPath $Out
New-Item -ItemType Directory -Force -Path $outPath | Out-Null
foreach($file in Get-ChildItem -LiteralPath (Join-Path $rootPath $Source) -Filter '*.pptx'){
 $p=$app.Presentations.Open($file.FullName,-1,0,0)
 try {
  if($p.Slides.Count -ne 1){throw 'Compiled package contains multiple slides'}
  $p.Slides.Item(1).Export((Join-Path $outPath ($file.BaseName+'.png')),'PNG',1280,720)
 }finally{$p.Close()}
}
Write-Output ('Exported to '+$outPath)
