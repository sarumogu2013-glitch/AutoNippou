# worklog_setup.ps1 — タスクスケジューラに登録（ログオン時に記録開始／毎日17:30に日報作成）
# このファイルと同じフォルダに worklog_logger.ps1 / worklog_report.ps1 を置いて実行

$ReportTime = '17:30'
$here = $PSScriptRoot
$ps   = "$env:WINDIR\System32\WindowsPowerShell\v1.0\powershell.exe"
$user = "$env:USERDOMAIN\$env:USERNAME"

# 記録（ログオン時に起動・時間制限なし）
$a1 = New-ScheduledTaskAction -Execute $ps -Argument "-NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$here\worklog_logger.ps1`""
$t1 = New-ScheduledTaskTrigger -AtLogOn -User $user
$s1 = New-ScheduledTaskSettingsSet -ExecutionTimeLimit ([TimeSpan]::Zero) -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -MultipleInstances IgnoreNew
Register-ScheduledTask -TaskName 'WorkLog_Logger' -Action $a1 -Trigger $t1 -Settings $s1 -Force | Out-Null

# 日報（平日の指定時刻・PCが起動していなければ次回起動時）
$a2 = New-ScheduledTaskAction -Execute $ps -Argument "-NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$here\worklog_report.ps1`""
$t2 = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Monday,Tuesday,Wednesday,Thursday,Friday -At $ReportTime
$s2 = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries
Register-ScheduledTask -TaskName 'WorkLog_Report' -Action $a2 -Trigger $t2 -Settings $s2 -Force | Out-Null

# 今すぐ記録を開始
Start-ScheduledTask -TaskName 'WorkLog_Logger'
Write-Host "登録完了: 記録=ログオン時 / 日報=平日 $ReportTime"
Write-Host "ログ保存先: $env:LOCALAPPDATA\WorkLog"
