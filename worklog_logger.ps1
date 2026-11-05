# worklog_logger.ps1 — 前面ウィンドウを定期記録（常駐・低負荷）
# 記録先: %LOCALAPPDATA%\WorkLog\yyyy-MM-dd.csv

$IntervalSec = 60   # 記録間隔（秒）。report側の設定と合わせる

$ErrorActionPreference = 'SilentlyContinue'

# 二重起動防止
$created = $false
$mutex = New-Object System.Threading.Mutex($true, 'Global\WorkLogLogger', [ref]$created)
if (-not $created) { exit }

$dir = Join-Path $env:LOCALAPPDATA 'WorkLog'
New-Item -ItemType Directory -Force -Path $dir | Out-Null

Add-Type @"
using System;
using System.Runtime.InteropServices;
using System.Text;
public static class WL {
    [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern int GetWindowText(IntPtr h, StringBuilder s, int n);
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
    [StructLayout(LayoutKind.Sequential)] public struct LII { public uint cbSize; public uint dwTime; }
    [DllImport("user32.dll")] static extern bool GetLastInputInfo(ref LII p);
    public static uint IdleSec() {
        LII l = new LII(); l.cbSize = (uint)Marshal.SizeOf(l);
        GetLastInputInfo(ref l);
        return ((uint)Environment.TickCount - l.dwTime) / 1000;
    }
}
"@

while ($true) {
    $h  = [WL]::GetForegroundWindow()
    $sb = New-Object System.Text.StringBuilder 512
    [void][WL]::GetWindowText($h, $sb, 512)
    $procId = [uint32]0
    [void][WL]::GetWindowThreadProcessId($h, [ref]$procId)
    $proc  = (Get-Process -Id $procId -ErrorAction SilentlyContinue).ProcessName
    $title = $sb.ToString() -replace '[\r\n"]', ' '

    $now  = Get-Date
    $file = Join-Path $dir ($now.ToString('yyyy-MM-dd') + '.csv')
    if (-not (Test-Path $file)) {
        Set-Content -Path $file -Value '"time","process","title","idle"' -Encoding UTF8
    }
    $line = '"{0}","{1}","{2}",{3}' -f $now.ToString('yyyy-MM-dd HH:mm:ss'), $proc, $title, [WL]::IdleSec()
    Add-Content -Path $file -Value $line -Encoding UTF8

    Start-Sleep -Seconds $IntervalSec
}
