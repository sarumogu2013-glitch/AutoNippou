# build.ps1 — WorkLog.exe を作る（Windows 標準の .NET Framework 4.x の C# コンパイラを使用）
$csc = "$env:WINDIR\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
& $csc /nologo /target:winexe /codepage:65001 /optimize+ /win32manifest:"$PSScriptRoot\app.manifest" /win32icon:"$PSScriptRoot\WorkLog.ico" `
    /r:System.Windows.Forms.dll /r:System.Drawing.dll /r:System.Web.Extensions.dll /r:System.IO.Compression.dll /r:System.IO.Compression.FileSystem.dll `
    /out:"$PSScriptRoot\WorkLog.exe" "$PSScriptRoot\WorkLog.cs"
if ($LASTEXITCODE -eq 0) { Write-Host "作成: $PSScriptRoot\WorkLog.exe" }

exit $LASTEXITCODE
