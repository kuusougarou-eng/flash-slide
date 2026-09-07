# ウィンドウを前面化せずに連続キャプチャして PNG 連番にする(マーケ動画・操作記録用)
# usage: pwsh scripts/record-window.ps1 "<title substring>" <outDir> <seconds> [fps=6]
#   → <outDir>/f000001.png … を書く。動画化は例:
#   ffmpeg -y -framerate 6 -i <outDir>/f%06d.png -vf "scale=1920:-2" -c:v libx264 -pix_fmt yuv420p -crf 20 out.mp4
param([string]$title, [string]$outDir, [int]$seconds = 20, [int]$fps = 6)
Add-Type -AssemblyName System.Drawing
Add-Type @"
using System; using System.Runtime.InteropServices; using System.Text;
public static class RW {
  public delegate bool EnumWindowsProc(IntPtr h, IntPtr l);
  [DllImport("user32.dll")] public static extern bool EnumWindows(EnumWindowsProc cb, IntPtr l);
  [DllImport("user32.dll")] public static extern int GetWindowText(IntPtr h, StringBuilder s, int n);
  [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
  [DllImport("user32.dll")] public static extern bool PrintWindow(IntPtr h, IntPtr hdc, uint flags);
  [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
  [DllImport("user32.dll")] public static extern bool SetProcessDPIAware();
  [StructLayout(LayoutKind.Sequential)] public struct RECT { public int L, T, R, B; }
  public static IntPtr Find(string like) { IntPtr found = IntPtr.Zero; EnumWindows((h, l) => { if (!IsWindowVisible(h)) return true; var sb = new StringBuilder(512); GetWindowText(h, sb, 512); if (sb.ToString().Contains(like)) { found = h; return false; } return true; }, IntPtr.Zero); return found; }
}
"@
[RW]::SetProcessDPIAware() | Out-Null
$h = [RW]::Find($title)
if ($h -eq [IntPtr]::Zero) { Write-Output "window not found: $title"; exit 1 }
New-Item -ItemType Directory -Force $outDir | Out-Null
$r = New-Object RW+RECT; [RW]::GetWindowRect($h, [ref]$r) | Out-Null
$w = $r.R - $r.L; $hh = $r.B - $r.T
$interval = [int](1000 / $fps)
$n = 0
$sw = [Diagnostics.Stopwatch]::StartNew()
while ($sw.Elapsed.TotalSeconds -lt $seconds) {
  $t0 = $sw.ElapsedMilliseconds
  $bmp = New-Object System.Drawing.Bitmap $w, $hh
  $g = [System.Drawing.Graphics]::FromImage($bmp)
  $hdc = $g.GetHdc(); [RW]::PrintWindow($h, $hdc, 2) | Out-Null; $g.ReleaseHdc($hdc)
  $n++
  $bmp.Save((Join-Path $outDir ("f{0:D6}.png" -f $n)), [System.Drawing.Imaging.ImageFormat]::Png)
  $g.Dispose(); $bmp.Dispose()
  $rest = $interval - ($sw.ElapsedMilliseconds - $t0)
  if ($rest -gt 0) { Start-Sleep -Milliseconds $rest }
}
Write-Output ("frames={0} size={1}x{2} dir={3}" -f $n, $w, $hh, $outDir)
