# 開発検証用 UI ヘルパ: スクリーンショット / クリック / キー送信
# usage:
#   pwsh scripts/ui.ps1 shot <out.png> [display]   display: 省略=前面ウィンドウのあるモニタ / 1..n / all
#   pwsh scripts/ui.ps1 focuswin <hwnd> / max <hwnd> / shotwin <hwnd> <out.png>(前面化せずに撮る)
#   負の座標(サブモニタ)は pwsh 内から & scripts/ui.ps1 click 1716 -1377 のように呼ぶ(-File だと -1377 がパラメータ扱いになる)
#   pwsh scripts/ui.ps1 click <x> <y>
#   pwsh scripts/ui.ps1 type <text>
#   pwsh scripts/ui.ps1 keys <SendKeys 文字列>   例: "^{ENTER}"
#   pwsh scripts/ui.ps1 focus <window title substring>
param([string]$cmd, [string]$a, [string]$b)
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type @"
using System; using System.Runtime.InteropServices;
public static class U {
  [DllImport("user32.dll")] public static extern bool SetCursorPos(int x, int y);
  [DllImport("user32.dll")] public static extern void mouse_event(uint f, uint dx, uint dy, uint d, UIntPtr e);
  [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
  [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr h, int n);
  [DllImport("user32.dll")] public static extern bool SetProcessDPIAware();
  [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
  [DllImport("user32.dll")] public static extern bool PrintWindow(IntPtr h, IntPtr hdc, uint flags);
  [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
  [StructLayout(LayoutKind.Sequential)] public struct RECT { public int L, T, R, B; }
}
"@
[U]::SetProcessDPIAware() | Out-Null
switch ($cmd) {
  "shot" {
    $screens = [System.Windows.Forms.Screen]::AllScreens
    if ($b -eq "all") {
      $b0 = [System.Windows.Forms.SystemInformation]::VirtualScreen
    } elseif ($b -match '^\d+$') {
      $b0 = $screens[[int]$b - 1].Bounds
    } else {
      $fg = [U]::GetForegroundWindow(); $r = New-Object U+RECT; [U]::GetWindowRect($fg, [ref]$r) | Out-Null
      $cx = [int](($r.L + $r.R) / 2); $cy = [int](($r.T + $r.B) / 2)
      $hit = $screens | Where-Object { $_.Bounds.Contains($cx, $cy) } | Select-Object -First 1
      $b0 = if ($hit) { $hit.Bounds } else { [System.Windows.Forms.Screen]::PrimaryScreen.Bounds }
    }
    $bmp = New-Object System.Drawing.Bitmap $b0.Width, $b0.Height
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.CopyFromScreen($b0.Location, [System.Drawing.Point]::Empty, $b0.Size)
    $bmp.Save($a, [System.Drawing.Imaging.ImageFormat]::Png)
    Write-Output "saved $a $($b0.Width)x$($b0.Height)"
  }
  "click" {
    [U]::SetCursorPos([int]$a, [int]$b) | Out-Null; Start-Sleep -Milliseconds 120
    [U]::mouse_event(2, 0, 0, 0, [UIntPtr]::Zero); [U]::mouse_event(4, 0, 0, 0, [UIntPtr]::Zero)
    Write-Output "clicked $a,$b"
  }
  "type" { [System.Windows.Forms.SendKeys]::SendWait($a); Write-Output "typed" }
  "keys" { [System.Windows.Forms.SendKeys]::SendWait($a); Write-Output "sent $a" }
  "shotwin" {
    # 前面化せずにウィンドウ内容を撮る(PrintWindow + PW_RENDERFULLCONTENT)。usage: shotwin <hwnd|タイトル部分一致> <out.png>
    if ($a -match '^\d+$') { $h = [IntPtr][int64]$a } else {
      $h = [IntPtr]::Zero
      $procs = Get-Process | Where-Object { $_.MainWindowTitle -like "*$a*" }
      if ($procs) { $h = ($procs | Select-Object -First 1).MainWindowHandle }
      if ($h -eq [IntPtr]::Zero) {
        Add-Type @"
using System; using System.Runtime.InteropServices; using System.Text;
public static class WF {
  public delegate bool EnumWindowsProc(IntPtr h, IntPtr l);
  [DllImport("user32.dll")] public static extern bool EnumWindows(EnumWindowsProc cb, IntPtr l);
  [DllImport("user32.dll")] public static extern int GetWindowText(IntPtr h, StringBuilder s, int n);
  [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
  public static IntPtr Find(string like) { IntPtr found = IntPtr.Zero; EnumWindows((h, l) => { if (!IsWindowVisible(h)) return true; var sb = new StringBuilder(512); GetWindowText(h, sb, 512); if (sb.ToString().Contains(like)) { found = h; return false; } return true; }, IntPtr.Zero); return found; }
}
"@
        $h = [WF]::Find($a)
      }
      if ($h -eq [IntPtr]::Zero) { Write-Output "window not found: $a"; exit 1 }
    }
    $r = New-Object U+RECT; [U]::GetWindowRect($h, [ref]$r) | Out-Null
    $w = $r.R - $r.L; $hh = $r.B - $r.T
    $bmp = New-Object System.Drawing.Bitmap $w, $hh
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $hdc = $g.GetHdc()
    [U]::PrintWindow($h, $hdc, 2) | Out-Null
    $g.ReleaseHdc($hdc)
    $bmp.Save($b, [System.Drawing.Imaging.ImageFormat]::Png)
    Write-Output "saved $b ${w}x${hh}"
  }
  "max" {
    $h = [IntPtr][int64]$a
    [U]::ShowWindow($h, 3) | Out-Null; [U]::SetForegroundWindow($h) | Out-Null; Write-Output "maximized hwnd $a"
  }
  "focuswin" {
    $h = [IntPtr][int64]$a
    [U]::ShowWindow($h, 9) | Out-Null; [U]::SetForegroundWindow($h) | Out-Null; Write-Output "focused hwnd $a"
  }
  "focus" {
    $p = Get-Process | Where-Object { $_.MainWindowTitle -like "*$a*" } | Select-Object -First 1
    if ($p) { [U]::ShowWindow($p.MainWindowHandle, 3) | Out-Null; [U]::SetForegroundWindow($p.MainWindowHandle) | Out-Null; Write-Output "focused $($p.MainWindowTitle)" } else { Write-Output "not found" }
  }
}
