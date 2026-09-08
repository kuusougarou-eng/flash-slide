param([string]$Deck,[string]$Output)
$deckPath=(Resolve-Path -LiteralPath $Deck).Path
$outputPath=[IO.Path]::GetFullPath($Output)
New-Item -ItemType Directory -Force -Path $outputPath | Out-Null
$app=[Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')
$p=$app.Presentations.Open($deckPath,-1,0,0)
try {
  foreach($s in @($p.Slides)) {
    $s.Export((Join-Path $outputPath ('slide-{0:d2}.png' -f $s.SlideIndex)),'PNG',1280,720)
  }
  Write-Output ('Exported {0} slides' -f $p.Slides.Count)
} finally { $p.Close() }
