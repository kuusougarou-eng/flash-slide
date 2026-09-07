# 指定プレゼン(名前部分一致)の指定スライドのシェイプを列挙(UTF-8 でファイルに書く)
# usage: powershell.exe -File scripts/ppt-slide-shapes.ps1 "<name like>" <slideIndex> <out.txt>
param([string]$nameLike, [int]$index, [string]$out)
$app = [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')
$p = $null
foreach ($x in @($app.Presentations)) { if ($x.Name -like "*$nameLike*") { $p = $x } }
$lines = @()
if (-not $p) { $lines += "presentation not found: $nameLike" }
else {
  $s = $p.Slides.Item($index)
  $lines += ("{0} slide {1} layout='{2}' shapes={3}" -f $p.Name, $s.SlideIndex, $s.CustomLayout.Name, $s.Shapes.Count)
  foreach ($sh in $s.Shapes) {
    $info = "  [{0}] type={1} L={2:N0} T={3:N0} W={4:N0} H={5:N0} vis={6}" -f $sh.Name, $sh.Type, $sh.Left, $sh.Top, $sh.Width, $sh.Height, $sh.Visible
    if ($sh.HasTextFrame) {
      $tr = $sh.TextFrame.TextRange
      $t = $tr.Text; if ($t.Length -gt 40) { $t = $t.Substring(0,40) }
      $info += (" font='{0}' {1}pt color={2} text='{3}'" -f $tr.Font.Name, $tr.Font.Size, $tr.Font.Color.RGB, ($t -replace "`r|`n"," "))
    }
    $lines += $info
  }
}
[IO.File]::WriteAllLines($out, $lines, (New-Object Text.UTF8Encoding($false)))
