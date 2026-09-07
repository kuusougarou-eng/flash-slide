# UI Automation で、指定ウィンドウ(タイトル部分一致)内のリボンタブを選択し、ボタン名でクリック(Invoke)する。
# 前面化やマウス移動を伴わないので、他のモニタでユーザーが作業中でも邪魔をしない。
# usage: pwsh scripts/uia-click.ps1 "<window title like>" "<tab name>" "<button name>"
#        pwsh scripts/uia-click.ps1 "<window title like>" list        … ウィンドウ直下の要素名を列挙(デバッグ)
param([string]$win, [string]$tab, [string]$button)
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
$root = [System.Windows.Automation.AutomationElement]::RootElement
$cond = New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::ControlTypeProperty, [System.Windows.Automation.ControlType]::Window)
$w = $null
foreach ($e in $root.FindAll([System.Windows.Automation.TreeScope]::Children, $cond)) {
  if ($e.Current.Name -like "*$win*") { $w = $e; break }
}
if (-not $w) { Write-Output "window not found: $win"; exit 1 }
Write-Output ("window: " + $w.Current.Name)
function FindByName($el, $name, $type) {
  $c1 = New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::NameProperty, $name)
  if ($type) {
    $c2 = New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::ControlTypeProperty, $type)
    $c1 = New-Object System.Windows.Automation.AndCondition($c1, $c2)
  }
  return $el.FindFirst([System.Windows.Automation.TreeScope]::Descendants, $c1)
}
if ($tab -eq "list") {
  $all = $w.FindAll([System.Windows.Automation.TreeScope]::Descendants, [System.Windows.Automation.Condition]::TrueCondition)
  foreach ($e in $all) { if ($e.Current.Name) { Write-Output ("{0} | {1}" -f $e.Current.ControlType.ProgrammaticName, $e.Current.Name) } }
  exit 0
}
$scope = $w
if ($tab -like "in:*") {
  # "in:<name>" … 指定名の子要素(タスクペイン等)の中だけを探す(ウィンドウの「閉じる」を誤爆しない)
  $scope = FindByName $w $tab.Substring(3) $null
  if (-not $scope) { Write-Output ("container not found: " + $tab.Substring(3)); exit 1 }
  Write-Output ("scope: " + $scope.Current.Name)
  $tab = ""
}
if ($tab) {
  $t = FindByName $w $tab ([System.Windows.Automation.ControlType]::TabItem)
  if (-not $t) { Write-Output "tab not found: $tab" } else {
    try { $t.GetCurrentPattern([System.Windows.Automation.SelectionItemPattern]::Pattern).Select(); Write-Output "tab selected: $tab" } catch { Write-Output "tab select failed: $_" }
    Start-Sleep -Milliseconds 800
  }
}
if ($button) {
  $b = FindByName $scope $button ([System.Windows.Automation.ControlType]::Button)
  if (-not $b) { $b = FindByName $scope $button $null }
  if (-not $b) { Write-Output "button not found: $button"; exit 2 }
  try { $b.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern).Invoke(); Write-Output "invoked: $button" }
  catch {
    try { $b.GetCurrentPattern([System.Windows.Automation.TogglePattern]::Pattern).Toggle(); Write-Output "toggled: $button" }
    catch { Write-Output "invoke failed: $_"; exit 3 }
  }
}
