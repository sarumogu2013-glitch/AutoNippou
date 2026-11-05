# worklog_report.ps1 — 1日分のログを集計し、Claudeで日報（箇条書き）を作成
# 使い方: powershell -ExecutionPolicy Bypass -File worklog_report.ps1 [-Date 2026-10-29]
param([string]$Date = (Get-Date -Format 'yyyy-MM-dd'))

# ===== 設定 =====
$IntervalSec = 60                 # logger と同じ値
$IdleLimit   = 300                # 無操作がこの秒数以上なら離席扱い
$MinMinutes  = 2                  # これ未満の作業は集計から除外
$Model       = 'claude-haiku-5-5' # 安価なモデルで十分
$KeepDays    = 90                 # これより古いログCSVは削除
$ExcludeProc = @('LockApp', 'ScreenSaver', 'explorer', '')   # 集計しないプロセス
$MaskProc    = @{                                             # タイトルを伏せるプロセス
    'OUTLOOK' = 'メール対応'
    'olk'     = 'メール対応'
}
$MaskWords   = @()                # タイトル中で伏せたい語（例: @('〇〇建設','△△様')）
$OpenNotepad = $true              # 作成後にメモ帳で開く
# ================

$ErrorActionPreference = 'Stop'
$dir = Join-Path $env:LOCALAPPDATA 'WorkLog'
$csvPath = Join-Path $dir "$Date.csv"
if (-not (Test-Path $csvPath)) { Write-Host "ログがありません: $csvPath"; exit 1 }

$log    = Import-Csv -Path $csvPath -Encoding UTF8
$active = @($log | Where-Object { [int]$_.idle -lt $IdleLimit -and $ExcludeProc -notcontains $_.process })
if ($active.Count -eq 0) { Write-Host '作業記録がありません'; exit 1 }

$start = ([datetime]$active[0].time).ToString('HH:mm')
$end   = ([datetime]$active[-1].time).ToString('HH:mm')
$totalMin = [int]($active.Count * $IntervalSec / 60)

# アプリ×タイトル別に集計
$rows = $active | ForEach-Object {
    $t = $_.title
    if ($MaskProc.ContainsKey($_.process)) { $t = $MaskProc[$_.process] }
    foreach ($w in $MaskWords) { $t = $t.Replace($w, '＊＊') }
    [pscustomobject]@{ p = $_.process; t = $t }
} | Group-Object p, t | ForEach-Object {
    [pscustomobject]@{
        Min  = [int]($_.Count * $IntervalSec / 60)
        Proc = $_.Group[0].p
        Title= $_.Group[0].t
    }
} | Where-Object { $_.Min -ge $MinMinutes } | Sort-Object Min -Descending | Select-Object -First 80

$summary = ($rows | ForEach-Object { '{0}分 [{1}] {2}' -f $_.Min, $_.Proc, $_.Title }) -join "`n"

$prompt = @"
以下はPCの作業ログです（使用時間・アプリ名・ウィンドウタイトル）。
これをもとに、今日行った作業を日報用の箇条書きにしてください。

条件:
- 1行20文字程度、行頭は「・」
- 5〜10行。時間の長い作業を優先し、同じ業務はまとめる
- ファイル名やページ名から業務内容を具体的に推定する（例:「〇〇工区の数量計算」）
- 推定できないもの、私用と思われるものは省く
- 箇条書きのみを出力（前置き・説明は不要）

作業ログ:
$summary
"@

$outPath = Join-Path $dir "日報_$Date.txt"
$header  = "日報 $Date（PC稼働 $start〜$end／操作時間 約$([math]::Round($totalMin/60,1))時間）`r`n"

$key = $env:ANTHROPIC_API_KEY
if (-not $key) { $key = [Environment]::GetEnvironmentVariable('ANTHROPIC_API_KEY', 'User') }

if (-not $key) {
    # APIキー未設定時: プロンプトをクリップボードへ（Claudeに貼り付けて使う）
    Set-Clipboard -Value $prompt
    Set-Content -Path $outPath -Value ($header + "`r`n（APIキー未設定。プロンプトをクリップボードにコピーしました）`r`n`r`n" + $prompt) -Encoding UTF8
} else {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $body = @{
        model      = $Model
        max_tokens = 600
        messages   = @(@{ role = 'user'; content = $prompt })
    } | ConvertTo-Json -Depth 5
    $res = Invoke-WebRequest -Uri 'https://api.anthropic.com/v1/messages' -Method Post -UseBasicParsing `
        -Headers @{ 'x-api-key' = $key; 'anthropic-version' = '2023-06-01' } `
        -ContentType 'application/json; charset=utf-8' `
        -Body ([Text.Encoding]::UTF8.GetBytes($body))
    $json = [Text.Encoding]::UTF8.GetString($res.RawContentStream.ToArray()) | ConvertFrom-Json
    $text = ($json.content | Where-Object { $_.type -eq 'text' } | ForEach-Object { $_.text }) -join "`n"
    Set-Content -Path $outPath -Value ($header + "`r`n" + $text.Trim() + "`r`n") -Encoding UTF8
}

# 古いログの整理
Get-ChildItem $dir -Filter '*.csv' | Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-$KeepDays) } | Remove-Item

Write-Host "作成: $outPath"
if ($OpenNotepad) { Start-Process notepad.exe $outPath }
