# ランチャーのスライド1を「既存デッキ」風(ネイビー Meiryo, リード, 罫線, フッター)に仕立てる。
# 文字化け回避のためテキストは UTF-8 JSON (scripts/style-strings.json) から読む。
$app = [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')
$p = $app.Presentations.Item(1)
$s = $p.Slides.Item(1)
$json = [System.IO.File]::ReadAllText("C:\Users\nakaj\PPTADDINCLAUDE\scripts\style-strings.json", [System.Text.Encoding]::UTF8) | ConvertFrom-Json
foreach ($sh in @($s.Shapes)) { $sh.Delete() }
$navy = 0x5F3A1F; $gray = 0x595959; $line = 0xBFBFBF
function AddText($l,$t,$w,$h,$txt,$sz,$bold,$col,$align){
  $tb = $s.Shapes.AddTextbox(1, $l, $t, $w, $h)
  $tb.TextFrame.WordWrap = -1
  $tr = $tb.TextFrame.TextRange
  $tr.Text = $txt; $tr.Font.Size = $sz; $tr.Font.Bold = $bold
  $tr.Font.Name = "Meiryo UI"; $tr.Font.Color.RGB = $col
  $tr.ParagraphFormat.Alignment = $align
}
AddText 48 28 864 44 $json.title 24 -1 $navy 1
$rule = $s.Shapes.AddShape(1, 48, 76, 864, 2); $rule.Fill.ForeColor.RGB=$navy; $rule.Line.Visible=0
AddText 48 84 864 40 $json.lead 16 0 $gray 1
$b1=$s.Shapes.AddShape(1,48,140,270,300); $b1.Fill.ForeColor.RGB=0xF2F2F2; $b1.Line.Visible=0
AddText 60 150 246 40 $json.body1 14 0 $gray 1
$b2=$s.Shapes.AddShape(1,345,140,567,300); $b2.Fill.ForeColor.RGB=0xF2F2F2; $b2.Line.Visible=0
AddText 357 150 543 40 $json.body2 14 0 $gray 1
AddText 48 470 600 20 $json.source 10 0 $gray 1
$r2=$s.Shapes.AddShape(1,48,500,864,1); $r2.Fill.ForeColor.RGB=$line; $r2.Line.Visible=0
AddText 48 506 400 20 $json.copyright 9 0 $gray 1
AddText 880 506 32 20 "1" 9 0 $gray 3
$p.Windows.Item(1).View.GotoSlide(1)
Write-Output ("styled slide1: shapes=" + $s.Shapes.Count)
