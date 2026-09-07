# 開いている PowerPoint の状態を列挙(Windows PowerShell 5.1 で実行: powershell.exe -File scripts/ppt-state.ps1)
$app = [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')
foreach ($p in @($app.Presentations)) { '{0} | saved={1} | slides={2} | path={3}' -f $p.Name, $p.Saved, $p.Slides.Count, $p.FullName }
'windows=' + $app.Windows.Count
if ($app.Windows.Count -gt 0) {
  'active=' + $app.ActiveWindow.Presentation.Name
  try { 'selectedSlide=' + $app.ActiveWindow.View.Slide.SlideIndex } catch { 'selectedSlide=?' }
}
