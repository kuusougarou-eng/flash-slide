# UI Automation で PowerPoint ウィンドウ内の要素を名前で探す(クリックしない・前面化しない)。
# usage: pwsh -File scripts/uia-find.ps1 "<name1>" "<name2>" ...   → 各名前の件数と ControlType を UTF-8 で出力
param([Parameter(ValueFromRemainingArguments = $true)][string[]]$names)
[Console]::OutputEncoding = [Text.Encoding]::UTF8
Add-Type -AssemblyName UIAutomationClient; Add-Type -AssemblyName UIAutomationTypes
$root = [System.Windows.Automation.AutomationElement]::RootElement
$cond = New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::ClassNameProperty, 'PPTFrameClass')
$win = $root.FindFirst([System.Windows.Automation.TreeScope]::Children, $cond)
if (-not $win) { Write-Output 'PowerPoint window not found'; exit 1 }
foreach ($name in $names) {
  $c = New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::NameProperty, $name)
  $els = $win.FindAll([System.Windows.Automation.TreeScope]::Descendants, $c)
  $types = @(); foreach ($e in $els) { $types += $e.Current.ControlType.ProgrammaticName }
  Write-Output ("{0}: {1} {2}" -f $name, $els.Count, ($types -join ','))
}
