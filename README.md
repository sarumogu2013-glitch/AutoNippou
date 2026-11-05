# AutoNippou（WorkLog）

PC の前面ウィンドウを低負荷で記録し、Claude Code（Pro プラン）で日報を自動で作る Windows のタスクトレイ常駐アプリ。

- ソース: `WorkLogTray/WorkLog.cs`（C# 5・WinForms）
- ビルド: `WorkLogTray/build.ps1`（Windows 標準の .NET Framework 4.x の csc.exe。追加インストール不要）
- 詳しい仕様・決まり・確認の記録: `引き継ぎ.md`
- `worklog_*.ps1` は作り替える前の PowerShell 版
