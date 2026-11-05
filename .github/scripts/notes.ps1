# CHANGELOG.md の「## 版（日付）」の節から、その版のリリースノートを作る（節がなければ失敗）
param([Parameter(Mandatory)][string]$Version, [Parameter(Mandatory)][string]$Out)
$md = Get-Content (Join-Path $PSScriptRoot '../../CHANGELOG.md') -Raw -Encoding utf8
$m = [regex]::Match($md, '(?ms)^## ' + [regex]::Escape($Version) + '（[^\r\n]*\r?\n(.*?)(?=^## |\z)')
if (-not $m.Success -or $m.Groups[1].Value.Trim().Length -eq 0) {
  [Console]::Error.WriteLine("CHANGELOG.md に「## $Version（日付）」の節がありません")
  exit 1
}
$notes = @"
> [!IMPORTANT]
> このソフトは AI（Anthropic の Claude / Claude Code）が作成したソフトウェアです。無保証です。

## 変わったところ

$($m.Groups[1].Value.Trim())

## 入れ方

- 初めて入れるとき: WorkLog-$Version.zip を展開し、中の「インストール手順.txt」をご覧ください
- 0.10.0 以降を使っているとき: 自動で更新されます（日報画面のメニュー「ヘルプ → 更新を確認」ですぐ確かめられます）
"@
Set-Content -Path $Out -Value $notes -Encoding utf8NoBOM
