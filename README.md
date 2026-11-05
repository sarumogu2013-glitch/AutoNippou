# AutoNippou（WorkLog）

> [!IMPORTANT]
> **このソフトは AI が作成したソフトウェアです。**
> ソースコード・ビルドとリリースの仕組み・説明文は、AI（Anthropic の Claude / Claude Code）が書きました。
> 人の手による網羅的なレビューやテストは受けていません。**無保証です。** 内容を確かめたうえで、自己責任でお使いください。

PC の前面ウィンドウを低負荷で記録し、Claude Code（Pro プラン）で日報を自動で作る Windows のタスクトレイ常駐アプリ。

## ダウンロード・インストール

[最新のリリース](https://github.com/sarumogu2013-glitch/AutoNippou/releases/latest) から `WorkLog-<版>.zip` を落として展開し、中の「インストール手順.txt」を見てください。

- 0.10.0 からは **自動で更新** されます（12時間ごとに GitHub のリリースを確かめ、日報画面を閉じているときに入れ替えて起動し直す）。設定のタブ「起動」で「知らせるだけ」にもできます
- **Windows 起動時の自動起動**: 初めて起動したときに聞きます。あとからトレイのアイコンの右クリックメニューか、設定のタブ「起動」で登録・解除できます

## 注意

- 作業の記録（ウィンドウのタイトルなど）は、この PC の `%LOCALAPPDATA%\WorkLog` に保存します
- 日報を作るときは、その要約を Claude Code 経由で Anthropic の AI に送ります（伏せ字の設定ができます）
- AI が書いた日報は間違うことがあります。提出する前に必ず確かめてください

## 開発

- ソース: `WorkLogTray/WorkLog.cs`（C# 5・WinForms）。アイコン: `WorkLogTray/WorkLog.ico`
- ビルド: `WorkLogTray/build.ps1`（Windows 標準の .NET Framework 4.x の csc.exe。追加インストール不要）
- リリース: `WorkLog.cs` の `Core.Version` を上げて main にマージすると、GitHub Actions がビルド・自動アップデートの動作確認をして、zip をリリースに載せる（`.github/workflows/release.yml`）
- 詳しい仕様・決まり・確認の記録: `引き継ぎ.md`
- `worklog_*.ps1` は作り替える前の PowerShell 版
