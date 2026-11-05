# 🤖 このソフトは AI が作成したソフトウェアです

> [!CAUTION]
> ## ソースコード・ビルドとリリースの仕組み・説明文・アイコンは、すべて AI（Anthropic の Claude / Claude Code）が作成しました
>
> - 人の手による網羅的なレビューやテストは受けていません
> - **無保証です。** 不具合で記録や日報が失われても責任を負えません
> - 内容を確かめたうえで、**自己責任で** お使いください
> - AI が書いた日報は間違うことがあります。提出する前に必ず確かめてください

---

# AutoNippou（WorkLog）

PC の前面ウィンドウを低負荷で記録し、Claude Code（Pro プラン）で日報を自動で作る Windows のタスクトレイ常駐アプリ。

## ダウンロード・インストール

[最新のリリース](https://github.com/sarumogu2013-glitch/AutoNippou/releases/latest) から `WorkLog-<版>.zip` を落として展開し、中の「インストール手順.txt」を見てください。

- 0.10.0 からは **自動で更新** されます（12時間ごとに GitHub のリリースを確かめ、日報画面を閉じているときに入れ替えて起動し直す）。設定のタブ「起動」で「知らせるだけ」にもできます
- **Windows 起動時の自動起動**: 初めて起動したときに聞きます。あとからトレイのアイコンの右クリックメニューか、設定のタブ「起動」で登録・解除できます

## 注意

- 作業の記録（ウィンドウのタイトルなど）は、この PC の `%LOCALAPPDATA%\WorkLog` に保存します
- 日報を作るときは、その要約を Claude Code 経由で Anthropic の AI に送ります（伏せ字の設定ができます）
- AI が書いた日報は間違うことがあります。提出する前に必ず確かめてください

## ライセンス

[MIT No Attribution（MIT-0）](LICENSE) © 2026 sarumogu2013-glitch

- 使う・直す・再配布する・売るのは自由です。著作権表示を残す義務もありません
- **無保証です。** 使ったことで生じた損害について、作者は責任を負いません

### ほかのソフトウェア・名前について

- **Claude Code**: このソフトは、PC に入っている Claude Code（Anthropic）を呼び出して使うだけで、同梱も再配布もしていません。Claude Code を使うときは、[Anthropic の利用規約](https://www.anthropic.com/legal) に従ってください。このソフトを再配布するときも、Claude Code を同梱しないでください
- **商標**: 「Claude」「Anthropic」は Anthropic の商標です。このソフトは Anthropic の公式ソフトウェアではなく、Anthropic が推奨・保証するものでもありません
- **アイコン**: アイコンの「報」の字は、IPA ゴシック（[IPA フォントライセンス v1.0](https://moji.or.jp/ipafont/license/)）で描いた画像です。フォント自体は含んでいません
- **.NET Framework / Windows**: Windows に標準で入っているものを使います。同梱していません

## 開発

- ソース: `WorkLogTray/WorkLog.cs`（C# 5・WinForms）。アイコン: `WorkLogTray/WorkLog.ico`
- ビルド: `WorkLogTray/build.ps1`（Windows 標準の .NET Framework 4.x の csc.exe。追加インストール不要）
- リリース: `WorkLog.cs` の `Core.Version` を上げて main にマージすると、GitHub Actions がビルド・自動アップデートの動作確認をして、zip をリリースに載せる（`.github/workflows/release.yml`）
- 詳しい仕様・決まり・確認の記録: `引き継ぎ.md`
- `worklog_*.ps1` は作り替える前の PowerShell 版
