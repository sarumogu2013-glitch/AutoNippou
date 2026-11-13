// WorkLog.cs — PCの作業ログを記録し、Claudeで日報（箇条書き）を作るタスクトレイ常駐アプリ
// ビルド: build.ps1（Windows 標準の .NET Framework 4.x の csc.exe。C# 5 の書き方で書く）
// 保存先: %LOCALAPPDATA%\WorkLog（yyyy-MM-dd.csv / 日報_yyyy-MM-dd.txt / settings.ini / error.log / ai_calls.log / old\）
// AI: この PC の Claude Code（claude.exe）を、claude.ai のプランのログインで呼ぶ（API の従量課金は使わない）
// 起動オプション:
//   （なし）             トレイに常駐（すでに起動していれば日報画面を開く）
//   --show               常駐して日報画面も開く
//   --ai                 常駐して「AIの接続先」も開く
//   --settings           常駐して「設定」も開く
//   --report yyyy-MM-dd  その日の日報を作り直して終了（画面なし）
//   --auto               未作成の日報をまとめて作って終了（画面なし）
//   --updated <pid>      自動アップデートで起動された（前の版の終了を待ってから常駐する）
//   --screenshot-about <png>  テスト中だけ: バージョン情報の画面を画像にして終わる
//   --screenshot-settings <dir>  テスト中だけ: 設定画面の各タブを画像にして終わる
//   --screenshot-timeline <png>  テスト中だけ: 日報画面の帯（テスト用の日）を画像にして終わる
//   --selftest           テスト中だけ: 日報の書式を直す処理を確かめる（0=期待どおり 1=違う）
//   --check-update [--apply]  テスト中（WORKLOG_TEST）だけ: 更新を確かめて落とす（--apply で入れ替えも）。0=最新 10=更新あり 2=失敗
// 動作確認用の環境変数: WORKLOG_TEST=保存先フォルダ（二重起動防止を本番と分け、スタートアップ登録を無効にする）
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Text;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Web.Script.Serialization;
using System.Windows.Forms;
using Microsoft.Win32;
using System.Reflection;

// 版の番号は Core.Version だけを直す（exe のプロパティの「詳細」にも出る）
[assembly: AssemblyTitle("作業ログ・日報")]
[assembly: AssemblyProduct("作業ログ・日報")]
[assembly: AssemblyVersion(WorkLog.Core.Version)]
[assembly: AssemblyFileVersion(WorkLog.Core.Version)]
[assembly: AssemblyInformationalVersion(WorkLog.Core.Version)]

namespace WorkLog
{
    static class Core
    {
        public const int IntervalSec = 60;   // 記録間隔（秒）
        public const string Version = "0.22.0";   // 主.機能.修正（変わったところは CHANGELOG.md に書く。リリースのノートになる）
        public static readonly bool Test = !string.IsNullOrEmpty(Environment.GetEnvironmentVariable("WORKLOG_TEST"));
        public static readonly string Dir = Test ? Environment.GetEnvironmentVariable("WORKLOG_TEST")
            : Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "WorkLog");
        public static readonly string Suffix = Test ? "_test" : "";

        static readonly object errLock = new object();
        static string lastErr;
        static DateTime lastErrTime;

        // error.log に追記（同じエラーが続くときは1時間に1回だけ）
        public static void Error(string where, string msg)
        {
            lock (errLock)
            {
                string key = where + msg;
                if (key == lastErr && DateTime.Now - lastErrTime < TimeSpan.FromHours(1)) return;
                lastErr = key; lastErrTime = DateTime.Now;
                try
                {
                    File.AppendAllText(Path.Combine(Dir, "error.log"),
                        string.Format("{0} [{1}] {2}\r\n", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), where, msg), new UTF8Encoding(true));
                }
                catch { }
            }
        }

        // 記録（CSV）は 保存先\ログ、日報は 保存先\日報（0.22.0 から。前は保存先の直下）
        public static string LogDir { get { string d = Path.Combine(Dir, "ログ"); Directory.CreateDirectory(d); return d; } }
        public static string ReportDir { get { string d = Path.Combine(Dir, "日報"); Directory.CreateDirectory(d); return d; } }
        public static string ReportPath(string date) { return Path.Combine(ReportDir, "日報_" + date + ".txt"); }
        public static string CsvPath(string date) { return Path.Combine(LogDir, date + ".csv"); }
        public static string WindowsPath(string date) { return Path.Combine(LogDir, "窓_" + date + ".csv"); }

        // 0.22.0 より前の置き方（保存先の直下）の記録と日報を、ログ・日報のフォルダへ移す（移す先にあるものは上書きしない）
        public static void MoveToFolders()
        {
            try
            {
                foreach (string f in Directory.GetFiles(Dir, "*.csv"))
                    if (Regex.IsMatch(Path.GetFileName(f), @"^\d{4}-\d{2}-\d{2}\.csv$") && !File.Exists(Path.Combine(LogDir, Path.GetFileName(f))))
                        File.Move(f, Path.Combine(LogDir, Path.GetFileName(f)));
                foreach (string f in Directory.GetFiles(Dir, "日報_*.txt"))
                    if (Regex.IsMatch(Path.GetFileName(f), @"^日報_\d{4}-\d{2}-\d{2}\.txt$") && !File.Exists(Path.Combine(ReportDir, Path.GetFileName(f))))
                        File.Move(f, Path.Combine(ReportDir, Path.GetFileName(f)));
            }
            catch (Exception ex) { Error("フォルダの整理", ex.Message); }
        }
        public static void WriteText(string path, string text) { File.WriteAllText(path, text, new UTF8Encoding(true)); }

        // 記録中のファイルも読めるように共有モードで開く
        public static string ReadText(string path)
        {
            using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
            using (var r = new StreamReader(fs, Encoding.UTF8))
                return r.ReadToEnd();
        }
    }

    // settings.ini（保存先フォルダ）。書き換えると次の記録・日報作成から反映される
    class Settings
    {
        public TimeSpan ReportTime = new TimeSpan(17, 30, 0);
        public int IdleLimit = 300, MinMinutes = 2, MinDayMinutes = 30, RedoMinutes = 15, CatchUpDays = 7, KeepDays = 90;
        public string ClaudePath = "";
        public string Model = "claude-haiku-5-5";
        public string FixModel = "claude-opus-5-5";   // 「AIで直す」は高いモデルで
        public HashSet<string> ExcludeProc = new HashSet<string>(new[] { "LockApp", "ScreenSaver", "explorer" }, StringComparer.OrdinalIgnoreCase);
        public Dictionary<string, string> MaskProc = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) { { "OUTLOOK", "メール対応" }, { "olk", "メール対応" } };
        public List<string> MaskWords = new List<string>();
        public bool OpenWindowOnReport = true;
        public int AfkMinutes = 15;     // 入力がこの分数なかったら離席（0 で記録しない）
        public bool AfkAsk = true;      // 離席から戻ったときに何をしていたか聞く
        public TimeSpan BreakStart = new TimeSpan(12, 0, 0), BreakEnd = new TimeSpan(13, 0, 0);   // 休憩時間（同じなら休憩なし）
        public int AppendMinutes = 60;  // 今日の日報をこの分数ごとに追記する（0 で ReportTime にまとめて作る）
        public TimeSpan WorkStart = new TimeSpan(8, 0, 0);   // 退勤中でも、この時刻を過ぎて操作があれば出勤とみなす
        public TimeSpan PlanFrom = new TimeSpan(8, 30, 0), PlanTo = new TimeSpan(17, 30, 0);   // 勤務予定（日報の時刻はこの中に収める）
        public HashSet<DayOfWeek> PlanDays = new HashSet<DayOfWeek> { DayOfWeek.Monday, DayOfWeek.Tuesday, DayOfWeek.Wednesday, DayOfWeek.Thursday, DayOfWeek.Friday };
        public const string DayChars = "日月火水木金土";   // DayOfWeek の順

        public bool HasPlan { get { return PlanTo > PlanFrom && PlanDays.Count > 0; } }
        public bool PlanOn(DateTime day) { return HasPlan && PlanDays.Contains(day.DayOfWeek); }
        public string PlanText { get { return PlanFrom.ToString(@"hh\:mm") + "〜" + PlanTo.ToString(@"hh\:mm"); } }
        public string PlanDaysText { get { return string.Concat(Enumerable.Range(0, 7).Where(i => PlanDays.Contains((DayOfWeek)i)).Select(i => DayChars[i])); } }
        public int OffIdleMinutes = 30;  // ReportTime を過ぎて、操作がこの分数なかったら退勤とみなす（0 で自動で退勤しない）
        public bool AutoUpdate = true;   // 新しい版が出たら自動で入れ替える（false は知らせるだけ）
        public string UserName = "";     // 要望を送るときの名前（日報画面の「要望」タブで入れる）
        public string LinkName = "";   // 日報リンクのボタン名（空なら「日報リンク」）。名前と開く先はリポジトリに置かないので既定は空
        public string LinkUrl = "";    // 日報リンクで開く URL かファイル（Excel の日報など）。空ならボタンを出さない

        public bool HasBreak { get { return BreakEnd > BreakStart; } }
        public string BreakText { get { return BreakStart.ToString(@"hh\:mm") + "〜" + BreakEnd.ToString(@"hh\:mm"); } }
        public bool InBreak(DateTime t) { return HasBreak && t.TimeOfDay >= BreakStart && t.TimeOfDay < BreakEnd; }

        // settings.ini の中身（設定画面で保存するときもこの形で書く）
        public string ToText()
        {
            var b = new StringBuilder("# 作業ログ・日報の設定（日報画面のメニュー「ツール → 設定」で直せます。保存すると次の記録・日報作成から反映。再起動は不要）\r\n");
            Action<string, string, string> add = (comment, key, value) => b.Append("# " + comment + "\r\n" + key + "=" + value + "\r\n");
            add("今日の日報を仕上げて知らせる時刻（追記しないときは、この時刻を過ぎてから作る）", "ReportTime", ReportTime.ToString(@"hh\:mm"));
            add("今日の日報をこの分数ごとに追記する（0 で追記しない）", "AppendMinutes", AppendMinutes.ToString());
            add("休憩時間（この時間の PC 操作は日報に入れない。空で休憩なし）", "Break", HasBreak ? BreakStart.ToString(@"hh\:mm") + "-" + BreakEnd.ToString(@"hh\:mm") : "");
            add("無操作がこの秒数以上なら離席扱い", "IdleLimit", IdleLimit.ToString());
            add("これ未満（分）の作業は集計しない", "MinMinutes", MinMinutes.ToString());
            add("自動で作るのは、操作時間がこの分数以上の日だけ", "MinDayMinutes", MinDayMinutes.ToString());
            add("日報を作ったあとにこの分数以上作業した日は、翌日以降に作り直す（前の版は old フォルダへ）", "RedoMinutes", RedoMinutes.ToString());
            add("さかのぼって未作成の日報を探す日数", "CatchUpDays", CatchUpDays.ToString());
            add("これより古いログ・旧版の日報は削除（日）", "KeepDays", KeepDays.ToString());
            add("日報を作るモデル（略称ではなく正式な名前で書く）", "Model", Model);
            add("「AIで直す」で使うモデル（略称ではなく正式な名前で書く）", "FixModel", FixModel);
            add(@"Claude Code（claude.exe）の場所。空なら PATH・%USERPROFILE%\.local\bin・Claude デスクトップアプリの中から探す", "ClaudePath", ClaudePath);
            add("集計しないプロセス（カンマ区切り）", "ExcludeProc", string.Join(",", ExcludeProc));
            add("タイトルを伏せるプロセス（プロセス名:置き換える言葉、カンマ区切り）", "MaskProc", string.Join(",", MaskProc.Select(kv => kv.Key + ":" + kv.Value)));
            add("タイトル中で伏せる語（カンマ区切り。例: 〇〇建設,△△様）", "MaskWords", string.Join(",", MaskWords));
            add("日報ができたら画面を開く（1=開く 0=通知だけ）", "OpenWindowOnReport", OpenWindowOnReport ? "1" : "0");
            add("入力がこの分数なかったら離席として記録する（0 で記録しない）", "AfkMinutes", AfkMinutes.ToString());
            add("離席から戻ったときに何をしていたか聞く（1=聞く 0=日報画面の離席メモに書くだけ）", "AfkAsk", AfkAsk ? "1" : "0");
            add("退勤中（待機）でも、この時刻を過ぎて操作があれば出勤とみなす", "WorkStart", WorkStart.ToString(@"hh\:mm"));
            add("勤務予定の時刻（日報の時刻はこの中に収める。始業・終業までに操作がない時間は「現場（外勤）」）", "PlanHours", PlanFrom.ToString(@"hh\:mm") + "-" + PlanTo.ToString(@"hh\:mm"));
            add("勤務予定の曜日（例: 月火水木金。空なら勤務予定を使わない）", "PlanDays", PlanDaysText);
            add("ReportTime を過ぎて、操作がこの分数なかったら退勤とみなす（0 で自動で退勤しない）", "OffIdleMinutes", OffIdleMinutes.ToString());
            add("新しい版が出たら自動で更新する（1=自動で更新 0=知らせるだけ）", "AutoUpdate", AutoUpdate ? "1" : "0");
            add("要望を送るときの名前（日報画面の「要望」タブで入れる）", "UserName", UserName);
            add("日報リンクのボタン名（空なら「日報リンク」）", "LinkName", LinkName);
            add("日報リンクで開く URL かファイル（空ならボタンを出さない。{yyyy} {MM} {dd} は今日の年・月・日）", "LinkUrl", LinkUrl);
            return b.ToString();
        }

        public static string FilePath { get { return Path.Combine(Core.Dir, "settings.ini"); } }
        static readonly object cacheLock = new object();
        static Settings cache;
        static DateTime cacheTime;

        // 前の中身は settings.ini.bak に残す
        public static void Save(Settings s)
        {
            lock (cacheLock)
            {
                if (File.Exists(FilePath)) File.Copy(FilePath, FilePath + ".bak", true);
                Core.WriteText(FilePath, s.ToText());
                cache = null;
            }
        }

        // Claude Code の場所だけを書き換える（「AIの接続先」の「詳しく探す」で見つけたとき）
        public static void SetClaudePath(string path)
        {
            Settings s;
            lock (cacheLock) s = Parse(File.Exists(FilePath) ? Core.ReadText(FilePath) : "");
            s.ClaudePath = path ?? "";
            Save(s);
        }

        // 要望の名前だけを書き換える（ほかの項目は settings.ini のまま）
        public static void SetUserName(string name)
        {
            name = (name ?? "").Replace("\r", " ").Replace("\n", " ").Trim();
            Settings s;
            lock (cacheLock) s = Parse(File.Exists(FilePath) ? Core.ReadText(FilePath) : "");
            if (s.UserName == name) return;
            s.UserName = name;
            Save(s);
        }

        public static Settings Get()
        {
            lock (cacheLock)
            {
                try
                {
                    if (!File.Exists(FilePath)) Core.WriteText(FilePath, new Settings().ToText());
                    DateTime t = File.GetLastWriteTime(FilePath);
                    if (cache == null || t != cacheTime) { cache = Parse(Core.ReadText(FilePath)); cacheTime = t; }
                }
                catch (Exception ex)
                {
                    Core.Error("設定", ex.Message);
                    if (cache == null) cache = new Settings();
                }
                return cache;
            }
        }

        static List<string> Items(string v)
        {
            return v.Split(',', '、').Select(x => x.Trim()).Where(x => x.Length > 0).ToList();
        }

        public static Settings FromText(string text) { return Parse(text); }   // 自己テスト用

        static Settings Parse(string text)
        {
            var s = new Settings();
            foreach (string raw in text.Split('\n'))
            {
                string line = raw.Trim();
                if (line.Length == 0 || line[0] == '#' || line[0] == ';') continue;
                int eq = line.IndexOf('=');
                if (eq < 0) continue;
                string k = line.Substring(0, eq).Trim(), v = line.Substring(eq + 1).Trim();
                try
                {
                    switch (k)
                    {
                        case "ReportTime": s.ReportTime = TimeSpan.Parse(v); break;
                        case "IdleLimit": s.IdleLimit = int.Parse(v); break;
                        case "MinMinutes": s.MinMinutes = int.Parse(v); break;
                        case "MinDayMinutes": s.MinDayMinutes = int.Parse(v); break;
                        case "RedoMinutes": s.RedoMinutes = int.Parse(v); break;
                        case "CatchUpDays": s.CatchUpDays = int.Parse(v); break;
                        case "KeepDays": s.KeepDays = int.Parse(v); break;
                        case "Model": if (v.Length > 0) s.Model = v; break;
                        case "FixModel": if (v.Length > 0) s.FixModel = v; break;
                        case "ClaudePath": s.ClaudePath = v.Trim('"'); break;
                        case "ExcludeProc": s.ExcludeProc = new HashSet<string>(Items(v), StringComparer.OrdinalIgnoreCase); break;
                        case "MaskProc":
                            s.MaskProc = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                            foreach (string item in Items(v))
                            {
                                int c = item.IndexOf(':');
                                if (c > 0) s.MaskProc[item.Substring(0, c).Trim()] = item.Substring(c + 1).Trim();
                            }
                            break;
                        case "MaskWords": s.MaskWords = Items(v); break;
                        case "OpenWindowOnReport": s.OpenWindowOnReport = v == "1" || v.ToLower() == "true"; break;
                        case "AfkMinutes": s.AfkMinutes = int.Parse(v); break;
                        case "AfkAsk": s.AfkAsk = v == "1" || v.ToLower() == "true"; break;
                        case "AppendMinutes": s.AppendMinutes = int.Parse(v); break;
                        case "WorkStart": s.WorkStart = TimeSpan.Parse(v); break;
                        case "PlanHours":
                            { string[] ph = v.Split('-'); s.PlanFrom = TimeSpan.Parse(ph[0].Trim()); s.PlanTo = TimeSpan.Parse(ph[1].Trim()); }
                            break;
                        case "PlanDays":
                            s.PlanDays = new HashSet<DayOfWeek>(v.Where(ch => DayChars.IndexOf(ch) >= 0).Select(ch => (DayOfWeek)DayChars.IndexOf(ch)));
                            break;
                        case "OffIdleMinutes": s.OffIdleMinutes = int.Parse(v); break;
                        case "AutoUpdate": s.AutoUpdate = v == "1" || v.ToLower() == "true"; break;
                        case "UserName": s.UserName = v; break;
                        case "LinkName":
                            // 古い設定ツールが「LinkName=名前 LinkUrl=URL」と1行に書いたものを直す
                            int u = v.IndexOf("LinkUrl=", StringComparison.Ordinal);
                            if (u >= 0) { if (s.LinkUrl.Length == 0) s.LinkUrl = v.Substring(u + 8).Trim(); v = v.Substring(0, u).Trim(); }
                            s.LinkName = v;
                            break;
                        case "LinkUrl": if (v.Length > 0) s.LinkUrl = v; break;
                        case "Break":
                            if (v.Length == 0) s.BreakStart = s.BreakEnd = TimeSpan.Zero;
                            else { string[] b = v.Split('-'); s.BreakStart = TimeSpan.Parse(b[0].Trim()); s.BreakEnd = TimeSpan.Parse(b[1].Trim()); }
                            break;
                    }
                }
                catch { Core.Error("設定", k + " の値が読めません: " + v); }
            }
            return s;
        }
    }

    // 前面ウィンドウの記録
    class Logger
    {
        [DllImport("user32.dll")] static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetWindowText(IntPtr h, StringBuilder s, int n);
        [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
        [StructLayout(LayoutKind.Sequential)] struct LII { public uint cbSize; public uint dwTime; }
        [DllImport("user32.dll")] static extern bool GetLastInputInfo(ref LII p);
        delegate bool EnumProc(IntPtr h, IntPtr l);
        [DllImport("user32.dll")] static extern bool EnumWindows(EnumProc cb, IntPtr l);
        [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr h);
        [DllImport("user32.dll")] static extern IntPtr GetWindow(IntPtr h, uint cmd);
        [DllImport("user32.dll")] static extern int GetWindowLong(IntPtr h, int index);
        [DllImport("user32.dll")] static extern int GetWindowTextLength(IntPtr h);
        [DllImport("dwmapi.dll")] static extern int DwmGetWindowAttribute(IntPtr h, int attr, out int value, int size);

        // ---- 開いているウィンドウ（開いた・閉じたときを ログ\窓_日付.csv に書く。閉じるまでその作業をしていたとみなす）----
        static readonly TimeSpan WinGap = TimeSpan.FromMinutes(5);   // これ以上記録があいたら（スリープ・電源オフ・再起動）、いったん閉じたとみなす
        HashSet<string> openWins;   // 前の記録のときに開いていたウィンドウ（アプリ<TAB>タイトル）
        DateTime openAt = DateTime.MinValue;

        static string StatePath { get { return Path.Combine(Core.LogDir, "窓_開いている.txt"); } }

        static string CleanTitle(string t) { return (t ?? "").Replace('\r', ' ').Replace('\n', ' ').Replace('"', ' ').Replace('\t', ' ').Trim(); }

        // 開いているウィンドウ: 見えている・持ち主のない・ツール用でない・隠されていない・タイトルのあるもの（集計しないプロセスと作業ログ自身は除く）
        public static HashSet<string> OpenWindows(Settings s)
        {
            var set = new HashSet<string>();
            var names = new Dictionary<uint, string>();
            EnumWindows((h, l) =>
            {
                if (!IsWindowVisible(h) || GetWindow(h, 4) != IntPtr.Zero) return true;            // GW_OWNER
                if ((GetWindowLong(h, -20) & 0x80) != 0) return true;                              // WS_EX_TOOLWINDOW
                int cloaked;
                if (DwmGetWindowAttribute(h, 14, out cloaked, 4) == 0 && cloaked != 0) return true; // DWMWA_CLOAKED
                int n = GetWindowTextLength(h);
                if (n <= 0) return true;
                var sb = new StringBuilder(n + 1);
                GetWindowText(h, sb, n + 1);
                string title = CleanTitle(sb.ToString());
                if (title.Length == 0 || title == "Program Manager") return true;
                uint pid;
                GetWindowThreadProcessId(h, out pid);
                string proc;
                if (!names.TryGetValue(pid, out proc))
                {
                    try { using (var p = Process.GetProcessById((int)pid)) proc = p.ProcessName; } catch { proc = ""; }
                    names[pid] = proc;
                }
                if (proc.Length == 0 || s.ExcludeProc.Contains(proc) || proc.Equals("WorkLog", StringComparison.OrdinalIgnoreCase)) return true;
                set.Add(proc + "\t" + title);
                return true;
            }, IntPtr.Zero);
            return set;
        }

        static void WriteWinEvents(DateTime at, string kind, IEnumerable<string> keys)
        {
            List<string> list = keys.ToList();
            if (list.Count == 0) return;
            string file = Core.WindowsPath(at.ToString("yyyy-MM-dd"));
            bool isNew = !File.Exists(file);
            using (var fs = new FileStream(file, FileMode.Append, FileAccess.Write, FileShare.ReadWrite))
            using (var w = new StreamWriter(fs, new UTF8Encoding(isNew)))
            {
                if (isNew) w.WriteLine("\"time\",\"event\",\"process\",\"title\"");
                foreach (string k in list)
                {
                    string[] c = k.Split(new[] { '\t' }, 2);
                    w.WriteLine(string.Format("\"{0}\",\"{1}\",\"{2}\",\"{3}\"", at.ToString("yyyy-MM-dd HH:mm:ss"), kind, c[0], c.Length > 1 ? c[1] : ""));
                }
            }
        }

        void SaveWinState()
        {
            try { File.WriteAllText(StatePath, openAt.ToString("yyyy-MM-dd HH:mm:ss") + "\r\n" + string.Join("\r\n", openWins), new UTF8Encoding(false)); }
            catch { }
        }

        void LoadWinState()
        {
            openWins = new HashSet<string>();
            try
            {
                if (!File.Exists(StatePath)) return;
                string[] lines = File.ReadAllLines(StatePath, Encoding.UTF8);
                DateTime t;
                if (lines.Length == 0 || !DateTime.TryParseExact(lines[0].Trim(), "yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture, DateTimeStyles.None, out t)) return;
                openAt = t;
                foreach (string l in lines.Skip(1)) if (l.Contains("\t")) openWins.Add(l);
            }
            catch { }
        }

        // 開いていたウィンドウをすべて、at に閉じたことにする（退勤・終了・記録があいたとき）
        public void CloseWindows(DateTime at)
        {
            try
            {
                if (openWins == null) LoadWinState();
                WriteWinEvents(at, "close", openWins);
                openWins.Clear();
                openAt = at;
                SaveWinState();
            }
            catch (Exception ex) { Core.Error("ウィンドウの記録", ex.Message); }
        }

        void TickWindows(DateTime now)
        {
            if (openWins == null) LoadWinState();
            // 日付が変わった・記録があいた: 前に開いていたものは、最後に確かめた時刻に閉じたことにする
            if (openWins.Count > 0 && (now.Date != openAt.Date || now - openAt > WinGap)) CloseWindows(openAt);
            HashSet<string> cur = OpenWindows(Settings.Get());
            WriteWinEvents(now, "close", openWins.Where(k => !cur.Contains(k)));
            WriteWinEvents(now, "open", cur.Where(k => !openWins.Contains(k)));
            openWins = cur;
            openAt = now;
            SaveWinState();
        }

        public DateTime LastOk = DateTime.MinValue;
        public string LastError;

        public static uint IdleSec()
        {
            var l = new LII();
            l.cbSize = (uint)Marshal.SizeOf(l);
            GetLastInputInfo(ref l);
            return unchecked((uint)Environment.TickCount - l.dwTime) / 1000;
        }

        public void Tick()
        {
            try
            {
                IntPtr h = GetForegroundWindow();
                var sb = new StringBuilder(512);
                GetWindowText(h, sb, 512);
                uint pid;
                GetWindowThreadProcessId(h, out pid);
                string proc = "";
                try { using (var p = Process.GetProcessById((int)pid)) proc = p.ProcessName; } catch { }
                string title = sb.ToString().Replace('\r', ' ').Replace('\n', ' ').Replace('"', ' ');

                DateTime now = DateTime.Now;
                string file = Core.CsvPath(now.ToString("yyyy-MM-dd"));
                bool isNew = !File.Exists(file);
                using (var fs = new FileStream(file, FileMode.Append, FileAccess.Write, FileShare.ReadWrite))
                using (var w = new StreamWriter(fs, new UTF8Encoding(isNew)))
                {
                    if (isNew) w.WriteLine("\"time\",\"process\",\"title\",\"idle\"");
                    w.WriteLine(string.Format("\"{0}\",\"{1}\",\"{2}\",{3}", now.ToString("yyyy-MM-dd HH:mm:ss"), proc, title, IdleSec()));
                }
                LastOk = now;
                LastError = null;
                try { TickWindows(now); }
                catch (Exception ex) { Core.Error("ウィンドウの記録", ex.Message); }
            }
            catch (Exception ex)
            {
                LastError = ex.Message;
                Core.Error("記録", ex.Message);
            }
        }
    }

    class Row { public DateTime Time; public string Proc, Title; }
    class Seg { public string Key; public DateTime Start, End; public int Rows; }
    class Result { public string Date, Message; public bool Ok, NoData, NeedLogin, Appended; }
    class NeedLoginException : Exception { public NeedLoginException(string m) : base(m) { } }
    class StateEntry { public DateTime Last; public bool Ok; }

    // 今取り掛かっているタスク（日報画面の右の欄「タスク」。タスク.txt: 名前<TAB>語（カンマ区切り）<TAB>完了なら done）
    class TaskItem
    {
        public string Name = "", Keywords = "";
        public bool Done;

        // ウィンドウタイトル（かアプリ名）に出る語。空ならタスクの名前で探す
        public List<string> Words
        {
            get
            {
                List<string> w = Keywords.Split(',', '、', '，').Select(x => x.Trim()).Where(x => x.Length > 0).ToList();
                if (w.Count == 0 && Name.Trim().Length > 0) w.Add(Name.Trim());
                return w;
            }
        }
    }

    static class Tasks
    {
        static string PathOf { get { return Path.Combine(Core.Dir, "タスク.txt"); } }

        static string Clean(string s) { return (s ?? "").Replace('\t', ' ').Replace('\r', ' ').Replace('\n', ' ').Trim(); }

        public static List<TaskItem> Load()
        {
            var list = new List<TaskItem>();
            try
            {
                if (!File.Exists(PathOf)) return list;
                foreach (string raw in Core.ReadText(PathOf).Split('\n'))
                {
                    string[] c = raw.TrimEnd('\r').Split('\t');
                    if (c[0].Trim().Length == 0 || c[0].StartsWith("#")) continue;
                    list.Add(new TaskItem { Name = c[0].Trim(), Keywords = c.Length > 1 ? c[1].Trim() : "", Done = c.Length > 2 && c[2].Trim() == "done" });
                }
            }
            catch (Exception ex) { Core.Error("タスク", ex.Message); }
            return list;
        }

        public static void Save(List<TaskItem> list)
        {
            var b = new StringBuilder("# 今取り掛かっているタスク（名前<TAB>ウィンドウタイトルに出る語（カンマ区切り）<TAB>終わったら done）。日報画面の右の欄「タスク」で直せます\r\n");
            foreach (TaskItem t in list.Where(x => Clean(x.Name).Length > 0))
                b.Append(Clean(t.Name) + "\t" + Clean(t.Keywords) + "\t" + (t.Done ? "done" : "") + "\r\n");
            Core.WriteText(PathOf, b.ToString());
        }

        // 前面のウィンドウのタイトルかアプリ名から、当てはまるタスクの名前（上にあるものから。なければ null）
        public static string Match(List<TaskItem> list, string proc, string title)
        {
            foreach (TaskItem t in list)
            {
                if (t.Done || t.Name.Trim().Length == 0) continue;
                foreach (string w in t.Words)
                    if ((title ?? "").IndexOf(w, StringComparison.OrdinalIgnoreCase) >= 0 || string.Equals(proc, w, StringComparison.OrdinalIgnoreCase))
                        return t.Name.Trim();
            }
            return null;
        }

        // ---- ウィンドウタイトルとタスクの対応（タスク_タイトル.txt: アプリ<TAB>タイトル<TAB>タスク<TAB>user か ai）----
        // 「タイトルとタスク」の画面で利用者が付ける。手で付けた対応が LearnMin 件たまったら、残りは AI が手本から予測して付ける
        public const string NoTask = "なし";   // タスクではないと決めたタイトル
        public const int LearnMin = 10;
        static string MapPath { get { return Path.Combine(Core.Dir, "タスク_タイトル.txt"); } }
        static readonly object mapLock = new object();

        public class TitleTask { public string Proc = "", Title = "", Task = ""; public bool ByAi; public DateTime Last; }

        // 表記ゆれ（保存していない印の * ● と前後の空白、アプリ名の大文字小文字）をそろえた見出し
        public static string Key(string proc, string title)
        {
            return (proc ?? "").Trim().ToLowerInvariant() + "\t" + (title ?? "").Trim().TrimStart('*', '●', '•').Trim();
        }

        public static Dictionary<string, TitleTask> LoadMap()
        {
            var map = new Dictionary<string, TitleTask>();
            try
            {
                if (!File.Exists(MapPath)) return map;
                foreach (string raw in Core.ReadText(MapPath).Split('\n'))
                {
                    string[] c = raw.TrimEnd('\r').Split('\t');
                    if (c.Length < 3 || c[0].StartsWith("#") || c[2].Trim().Length == 0) continue;
                    map[Key(c[0], c[1])] = new TitleTask { Proc = c[0].Trim(), Title = c[1].Trim(), Task = c[2].Trim(), ByAi = c.Length > 3 && c[3].Trim() == "ai" };
                }
            }
            catch (Exception ex) { Core.Error("タスク", ex.Message); }
            return map;
        }

        static void SaveMap(Dictionary<string, TitleTask> map)
        {
            var b = new StringBuilder("# ウィンドウタイトルとタスクの対応（アプリ<TAB>タイトル<TAB>タスク<TAB>手なら user、AI の予測なら ai）。「なし」はタスクではないタイトル\r\n");
            foreach (TitleTask t in map.Values)
                b.Append(Clean(t.Proc) + "\t" + Clean(t.Title) + "\t" + Clean(t.Task) + "\t" + (t.ByAi ? "ai" : "user") + "\r\n");
            Core.WriteText(MapPath, b.ToString());
        }

        // 1つのタイトルのタスクを決める（空なら対応を消す）
        public static void SetTitle(string proc, string title, string task, bool byAi)
        {
            lock (mapLock)
            {
                Dictionary<string, TitleTask> map = LoadMap();
                string k = Key(proc, title);
                if (string.IsNullOrWhiteSpace(task)) map.Remove(k);
                else map[k] = new TitleTask { Proc = proc, Title = title, Task = task.Trim(), ByAi = byAi };
                SaveMap(map);
            }
        }

        public static int UserCount(Dictionary<string, TitleTask> map) { return map.Values.Count(x => !x.ByAi); }

        // タスクの名前: タイトルの対応が先（「なし」ならタスクではない）。なければタイトルに出る語で
        public static string Resolve(Dictionary<string, TitleTask> map, List<TaskItem> tasks, string proc, string title)
        {
            TitleTask t;
            if (map.TryGetValue(Key(proc, title), out t)) return t.Task == NoTask ? null : t.Task;
            return Match(tasks, proc, title);
        }

        // 最近 days 日の記録の、アプリ×タイトルごとの使った時間（分。多い順）
        public static List<KeyValuePair<TitleTask, int>> RecentTitles(int days, Settings s)
        {
            var count = new Dictionary<string, KeyValuePair<TitleTask, int>>();
            for (int i = 0; i < days; i++)
            {
                List<Row> rows = Reporter.ReadActive(DateTime.Today.AddDays(-i).ToString("yyyy-MM-dd"), s);
                if (rows == null) continue;
                foreach (Row r in rows)
                {
                    string k = Key(r.Proc, r.Title);
                    KeyValuePair<TitleTask, int> v;
                    if (count.TryGetValue(k, out v)) { if (r.Time > v.Key.Last) v.Key.Last = r.Time; count[k] = new KeyValuePair<TitleTask, int>(v.Key, v.Value + 1); }
                    else count[k] = new KeyValuePair<TitleTask, int>(new TitleTask { Proc = r.Proc, Title = r.Title.Trim(), Last = r.Time }, 1);
                }
            }
            return count.Values.Select(v => new KeyValuePair<TitleTask, int>(v.Key, Reporter.Minutes(v.Value))).OrderByDescending(v => v.Value).ToList();
        }

        // AI に見せるタイトル（伏せ字の設定を反映）
        static string Shown(TitleTask t, Settings s)
        {
            string title;
            if (!s.MaskProc.TryGetValue(t.Proc, out title)) title = t.Title;
            foreach (string w in s.MaskWords) title = title.Replace(w, "＊＊");
            return t.Proc + " | " + title;
        }

        // 手で付けた対応を手本に、対応のないタイトルのタスクを AI（ふだんのモデル）が予測して付ける。
        // 手で付けた対応が LearnMin 件に満たないとき・タスクがないときは何もしない。付けた数を返す
        public static int Predict(List<TitleTask> titles, Settings s)
        {
            Dictionary<string, TitleTask> map = LoadMap();
            List<string> names = Load().Where(t => !t.Done && t.Name.Trim().Length > 0).Select(t => t.Name.Trim()).ToList();
            if (UserCount(map) < LearnMin || names.Count == 0) return 0;
            List<TitleTask> todo = titles.Where(t => t.Title.Trim().Length > 0 && !map.ContainsKey(Key(t.Proc, t.Title)))
                .GroupBy(t => Key(t.Proc, t.Title)).Select(g => g.First()).Take(80).ToList();
            if (todo.Count == 0) return 0;
            List<TitleTask> examples = map.Values.Where(x => !x.ByAi).ToList();
            examples = examples.Skip(Math.Max(0, examples.Count - 150)).ToList();
            var allowed = new HashSet<string>(names.Concat(examples.Select(x => x.Task)).Concat(new[] { NoTask }));
            string prompt =
                "ウィンドウタイトルから、どのタスクの作業かを当ててください。\n" +
                "下の「手本」は、利用者がアプリとウィンドウタイトルにタスクを付けたものです。同じ考え方で「当てるもの」にタスクを付けてください。\n\n" +
                "条件:\n" +
                "- タスクは「タスクの一覧」か手本にある名前から選ぶ。どれでもない・判断できないときは「" + NoTask + "」\n" +
                "- 出力は「番号: タスク名」を1行ずつだけ（前置き・説明は不要）\n\n" +
                "タスクの一覧: " + string.Join("、", allowed.Where(x => x != NoTask)) + "\n\n" +
                "手本（アプリ | ウィンドウタイトル → タスク）:\n" + string.Concat(examples.Select(x => Shown(x, s) + " → " + x.Task + "\n")) + "\n" +
                "当てるもの（番号: アプリ | ウィンドウタイトル）:\n" + string.Concat(todo.Select((x, i) => (i + 1) + ": " + Shown(x, s) + "\n"));
            Dictionary<int, string> got = ParsePredict(ClaudeCli.Ask(prompt, s, DateTime.Today.ToString("yyyy-MM-dd"), "（タスクの予測）"));
            int n = 0;
            lock (mapLock)
            {
                map = LoadMap();
                for (int i = 0; i < todo.Count; i++)
                {
                    string task, k = Key(todo[i].Proc, todo[i].Title);
                    if (!got.TryGetValue(i + 1, out task) || !allowed.Contains(task) || map.ContainsKey(k)) continue;
                    map[k] = new TitleTask { Proc = todo[i].Proc, Title = todo[i].Title, Task = task, ByAi = true };
                    n++;
                }
                if (n > 0) SaveMap(map);
            }
            return n;
        }

        // 「番号: タスク名」の行を読む
        public static Dictionary<int, string> ParsePredict(string answer)
        {
            var d = new Dictionary<int, string>();
            foreach (string raw in (answer ?? "").Replace("\r\n", "\n").Split('\n'))
            {
                Match m = Regex.Match(raw.Trim(), @"^(\d+)\s*[:：.．)）]\s*(.+)$");
                if (m.Success) d[int.Parse(m.Groups[1].Value)] = m.Groups[2].Value.Trim().Trim('「', '」').Trim();
            }
            return d;
        }

        // 日報を作る前に、その日の（MinMinutes 以上使った）タイトルで対応のないものを予測しておく。失敗しても日報は作る
        public static void PredictFor(List<Row> active, Settings s)
        {
            try
            {
                List<TitleTask> titles = active.GroupBy(r => Key(r.Proc, r.Title)).Where(g => Reporter.Minutes(g.Count()) >= Math.Max(1, s.MinMinutes))
                    .Select(g => new TitleTask { Proc = g.First().Proc, Title = g.First().Title.Trim() }).ToList();
                Predict(titles, s);
            }
            catch (Exception ex) { Core.Error("タスクの予測", ex.Message); }
        }

        // AI に渡す決まり（終わっていないタスクがあるときだけ）
        public static string Prompt()
        {
            List<string> names = Load().Where(t => !t.Done && t.Name.Trim().Length > 0).Select(t => t.Name.Trim()).ToList();
            if (names.Count == 0) return "";
            return "今取り掛かっているタスク（利用者が登録したもの）: " + string.Join("、", names) + "\n" +
                   "作業ログの〔タスク: …〕は、ウィンドウタイトルからそのタスクと判定した作業。その作業は、作業内容にタスク名を使って書く" +
                   "（例:「" + names[0] + "」「" + names[0] + "（図面修正）」）。同じタスクが続いているときは1つの工程にまとめる\n\n";
        }
    }

    // 離席（入力のなかった時間）と、そのあいだに何をしていたかのメモ
    class AfkEntry
    {
        public DateTime Start, End;
        public string Memo = "";
        public bool Skip;   // 休憩・私用など、日報に入れない
        public string Line = "";   // 日報に足した行（番号なし）。メモを直したときに入れ替える
        public bool Answered { get { return Skip || Memo.Length > 0; } }
        public string Range { get { return Start.ToString("HH:mm") + "〜" + End.ToString("HH:mm"); } }
        public string Duration
        {
            get
            {
                int m = (int)Math.Round((End - Start).TotalMinutes);
                return m >= 60 ? m / 60 + "時間" + (m % 60 > 0 ? m % 60 + "分" : "") : m + "分";
            }
        }
    }

    static class Afk
    {
        const string F = "yyyy-MM-dd HH:mm:ss";
        static readonly object fileLock = new object();

        public static string PathOf(string date) { return Path.Combine(Core.Dir, "離席_" + date + ".txt"); }
        public static string FreePath(string date) { return Path.Combine(Core.Dir, "メモ_" + date + ".txt"); }

        // 離席_yyyy-MM-dd.txt: 始め<TAB>終わり<TAB>1=日報に入れない<TAB>メモ<TAB>日報に足した行
        public static List<AfkEntry> Load(string date)
        {
            var list = new List<AfkEntry>();
            string p = PathOf(date);
            if (!File.Exists(p)) return list;
            foreach (string raw in Core.ReadText(p).Split('\n'))
            {
                string[] c = raw.TrimEnd('\r').Split('\t');
                DateTime a, b;
                if (c.Length < 4 || !DateTime.TryParseExact(c[0], F, CultureInfo.InvariantCulture, DateTimeStyles.None, out a) ||
                    !DateTime.TryParseExact(c[1], F, CultureInfo.InvariantCulture, DateTimeStyles.None, out b)) continue;
                list.Add(new AfkEntry { Start = a, End = b, Skip = c[2] == "1", Memo = c[3].Trim(), Line = c.Length > 4 ? c[4].Trim() : "" });
            }
            return list.OrderBy(e => e.Start).ToList();
        }

        static DateTime Seconds(DateTime d) { return new DateTime(d.Ticks - d.Ticks % TimeSpan.TicksPerSecond); }

        static string Clean(string s) { return (s ?? "").Replace('\t', ' ').Replace('\r', ' ').Replace('\n', ' ').Trim(); }

        // 始めた日の分に足す。同じ始まりのものは入れ替える
        public static void Put(AfkEntry e)
        {
            lock (fileLock)
            {
                e.Start = Seconds(e.Start);
                e.End = Seconds(e.End);
                string date = e.Start.ToString("yyyy-MM-dd");
                List<AfkEntry> list = Load(date);
                list.RemoveAll(x => x.Start == e.Start);
                list.Add(e);
                Core.WriteText(PathOf(date), string.Concat(list.OrderBy(x => x.Start).Select(x =>
                    x.Start.ToString(F) + "\t" + x.End.ToString(F) + "\t" + (x.Skip ? "1" : "0") + "\t" + Clean(x.Memo) + "\t" + Clean(x.Line) + "\r\n")));
            }
        }

        public static string ReadFree(string date)
        {
            string p = FreePath(date);
            return File.Exists(p) ? Core.ReadText(p).Trim() : "";
        }

        public static void SaveFree(string date, string text)
        {
            string p = FreePath(date);
            if (text.Trim().Length > 0) Core.WriteText(p, text.Trim() + "\r\n");
            else if (File.Exists(p)) File.Delete(p);
        }

        // 日報を作る AI に渡す分。メモのある離席はプログラムが日報に足すので、AI には「その時間はほかの作業」とだけ伝える
        public static string Prompt(string date)
        {
            List<AfkEntry> list = Load(date).Where(e => !e.Skip).ToList();
            List<AfkEntry> memo = list.Where(e => e.Memo.Length > 0).ToList();
            string free = ReadFree(date);
            if (memo.Count == 0 && free.Length == 0) return "";
            var sb = new StringBuilder();
            if (memo.Count > 0)
            {
                sb.Append("次の時間は PC から離れてほかの作業をしていた（その作業は別に日報へ足すので、ここでは書かない。工程はこの時間をまたがず前後で分ける）:\n");
                foreach (AfkEntry e in memo) sb.Append(e.Start.ToString("HH:mm") + "-" + e.End.ToString("HH:mm") + "\n");
                sb.Append("\n");
            }
            if (free.Length > 0)
            {
                sb.Append("時刻のない本人のメモ（下の「メモのない離席時間」のうち内容に合うものに当てはめて時刻を決め、工程か軽作業に入れる。" +
                          "合うものがなければ、作業ログから最もありそうな時刻を決めて軽作業に入れる）:\n" + free.Replace("\r\n", "\n") + "\n\n");
                List<AfkEntry> rest = list.Where(e => e.Memo.Length == 0).ToList();
                sb.Append("メモのない離席時間:\n" + (rest.Count == 0 ? "（なし）\n" :
                    string.Concat(rest.Select(e => e.Start.ToString("HH:mm") + "-" + e.End.ToString("HH:mm") + "（" + e.Duration + "）\n"))) + "\n");
            }
            return sb.ToString();
        }
    }

    // Claude Code（claude.exe）を、claude.ai のプランのログインで呼ぶ
    static class ClaudeCli
    {
        public class Auth { public bool LoggedIn; public string Method, Email, Org, Plan; }

        // これがあると Claude Code は API の従量課金で動くので、子プロセスには渡さない
        public static readonly string[] StripEnv = { "ANTHROPIC_API_KEY", "ANTHROPIC_AUTH_TOKEN" };
        const string SystemPrompt = "あなたは日報作成アシスタントです。指示された形だけを出力します。";
        static string LogPath { get { return Path.Combine(Core.Dir, "ai_calls.log"); } }
        static string WorkDir
        {
            get { string d = Path.Combine(Core.Dir, "claude_work"); Directory.CreateDirectory(d); return d; }
        }

        // デスクトップアプリ本体（Claude Code ではない）: ストア版の起動用の別名（WindowsApps\Claude.exe）と、通常版の本体（AnthropicClaude）
        public static bool IsDesktopApp(string p)
        {
            return p.IndexOf(@"\WindowsApps\", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   p.IndexOf(@"\AnthropicClaude\", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        public static string FindExe(Settings s)
        {
            if (s.ClaudePath.Length > 0) return File.Exists(s.ClaudePath) ? s.ClaudePath : null;
            string path = Environment.GetEnvironmentVariable("PATH") + ";" + Environment.GetEnvironmentVariable("PATH", EnvironmentVariableTarget.User);
            foreach (string d in path.Split(';'))
            {
                try
                {
                    string p = Path.Combine(Environment.ExpandEnvironmentVariables(d.Trim().Trim('"')), "claude.exe");
                    if (d.Trim().Length > 0 && File.Exists(p) && !IsDesktopApp(p)) return p;
                }
                catch { }
            }
            string home = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ".local", "bin", "claude.exe");
            if (File.Exists(home)) return home;
            return FindDesktop();
        }

        // Claude デスクトップアプリが入れる Claude Code の置き場所（claude-code\<版>\claude.exe か claude-code\<版>\<build>\claude.exe）
        public static List<string> DesktopRoots()
        {
            string roam = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string local = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            var roots = new List<string>();
            foreach (string app in new[] { "Claude", "Claude-3p" })
            {
                roots.Add(Path.Combine(roam, app, "claude-code"));
                roots.Add(Path.Combine(local, app, "claude-code"));
            }
            try
            {
                // ストア版（MSIX）は AppData が Packages\Claude_…\LocalCache の中に移される
                string pkgs = Path.Combine(local, "Packages");
                if (Directory.Exists(pkgs))
                    foreach (string d in Directory.GetDirectories(pkgs, "*Claude*"))
                        foreach (string app in new[] { "Claude", "Claude-3p" })
                        {
                            roots.Add(Path.Combine(d, "LocalCache", "Roaming", app, "claude-code"));
                            roots.Add(Path.Combine(d, "LocalCache", "Local", app, "claude-code"));
                        }
            }
            catch { }
            return roots;
        }

        public static System.Version PathVersion(string p)
        {
            System.Version best = new System.Version(0, 0);
            foreach (string part in p.Split('\\', '/'))
            {
                System.Version v;
                if (System.Version.TryParse(part.TrimStart('v', 'V'), out v) && v > best) best = v;
            }
            return best;
        }

        // デスクトップアプリの中の Claude Code。いくつかあれば版の新しいもの（同じなら新しく置かれたもの）
        public static string FindDesktop()
        {
            var found = new List<string>();
            foreach (string r in DesktopRoots())
            {
                try
                {
                    if (!Directory.Exists(r)) continue;
                    foreach (string f in Directory.GetFiles(r, "claude.exe", SearchOption.AllDirectories))
                        if (f.Substring(r.Length).Count(c => c == '\\') <= 3) found.Add(f);
                }
                catch { }
            }
            return found.OrderByDescending(PathVersion).ThenByDescending(f => File.GetLastWriteTime(f)).FirstOrDefault();
        }

        // 詳しく探す: ユーザーのフォルダと Program Files の中の claude.exe を探し、本当に Claude Code か（--version に「Claude Code」と出るか）を確かめる。
        // デスクトップアプリ本体（Electron のアプリ。resources\app.asar がある）は動かさない。見つかったもの（場所と --version）を返す
        public static List<KeyValuePair<string, string>> DeepSearch(Action<string> progress)
        {
            var roots = new List<string>
            {
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ".local"),
                Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles),
                Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86),
            };
            var skip = new[] { "Temp", "INetCache", "Microsoft", "Google", "Mozilla", "WorkLog", "CrashDumps", "D3DSCache", "pip", "npm-cache" };
            var files = new List<string>();
            DateTime until = DateTime.Now.AddSeconds(60);
            Action<string, int> walk = null;
            walk = (dir, depth) =>
            {
                if (depth > 7 || DateTime.Now > until) return;
                try
                {
                    foreach (string f in Directory.GetFiles(dir, "claude.exe")) files.Add(f);
                    foreach (string d in Directory.GetDirectories(dir))
                    {
                        string name = Path.GetFileName(d);
                        if (skip.Contains(name, StringComparer.OrdinalIgnoreCase) && depth <= 1) continue;
                        if ((File.GetAttributes(d) & FileAttributes.ReparsePoint) != 0) continue;
                        walk(d, depth + 1);
                    }
                }
                catch { }
            };
            foreach (string r in roots.Where(x => !string.IsNullOrEmpty(x) && Directory.Exists(x)).Distinct(StringComparer.OrdinalIgnoreCase))
            {
                if (progress != null) progress("探しています: " + r);
                walk(r, 0);
            }
            // ストア版のデスクトップアプリの中（Packages は大きいので Claude の名前のものだけ、深めに）
            try
            {
                string pkgs = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Packages");
                if (Directory.Exists(pkgs))
                    foreach (string d in Directory.GetDirectories(pkgs, "*Claude*")) walk(d, -3);
            }
            catch { }
            var found = new List<KeyValuePair<string, string>>();
            foreach (string f in files.Distinct(StringComparer.OrdinalIgnoreCase))
            {
                if (IsDesktopApp(f) || File.Exists(Path.Combine(Path.GetDirectoryName(f), "resources", "app.asar"))) continue;
                if (progress != null) progress("確かめています: " + f);
                string v;
                try { v = RunExe(f, "--version", null, 15000).Trim(); } catch { continue; }
                if (v.IndexOf("Claude Code", StringComparison.OrdinalIgnoreCase) >= 0) found.Add(new KeyValuePair<string, string>(f, v));
            }
            return found.OrderByDescending(x => PathVersion(x.Value + "\\" + x.Key)).ToList();
        }

        // 見つからないときの案内（「AIの接続先」に出す）
        public static string NotFoundHelp()
        {
            return "「詳しく探す」で検索、または「Claude Code の入れ方」を参照。";
        }

        // どこの Claude Code か（「AIの接続先」に出す）
        public static string Source(string exe)
        {
            if (exe == null) return "";
            string e = exe.Replace('/', '\\');
            if (e.IndexOf("\\claude-code\\", StringComparison.OrdinalIgnoreCase) >= 0) return "Claude デスクトップアプリに入っている Claude Code";
            return "Claude Code（単体でインストールしたもの）";
        }

        static string RunExe(string exe, string args, string stdin, int timeoutMs)
        {
            var psi = new ProcessStartInfo(exe, args)
            {
                UseShellExecute = false, CreateNoWindow = true, WorkingDirectory = WorkDir,
                RedirectStandardInput = true, RedirectStandardOutput = true, RedirectStandardError = true,
                StandardOutputEncoding = Encoding.UTF8, StandardErrorEncoding = Encoding.UTF8
            };
            foreach (string n in StripEnv) psi.EnvironmentVariables.Remove(n);
            using (Process p = Process.Start(psi))
            {
                var outTask = p.StandardOutput.ReadToEndAsync();
                var errTask = p.StandardError.ReadToEndAsync();
                using (var w = new StreamWriter(p.StandardInput.BaseStream, new UTF8Encoding(false)))
                    if (stdin != null) w.Write(stdin);
                if (!p.WaitForExit(timeoutMs))
                {
                    try { p.Kill(); } catch { }
                    throw new Exception("Claude Code が " + timeoutMs / 1000 + " 秒たっても終わりませんでした");
                }
                p.WaitForExit();
                string o = outTask.Result.Trim();
                return o.Length > 0 ? o : errTask.Result.Trim();
            }
        }

        static Dictionary<string, object> ParseJson(string text)
        {
            int i = text.IndexOf('{');
            if (i < 0) throw new Exception(text.Length > 200 ? text.Substring(0, 200) : text);
            return new JavaScriptSerializer().Deserialize<Dictionary<string, object>>(text.Substring(i));
        }

        static string Str(Dictionary<string, object> d, string k)
        {
            object v;
            return d.TryGetValue(k, out v) && v != null ? v.ToString() : "";
        }

        public static string Version(string exe)
        {
            try { return RunExe(exe, "--version", null, 30000); } catch (Exception ex) { return "（確認できません: " + ex.Message + "）"; }
        }

        public static Auth Status(string exe)
        {
            Dictionary<string, object> j = ParseJson(RunExe(exe, "auth status --json", null, 30000));
            return new Auth
            {
                LoggedIn = Str(j, "loggedIn") == "True",
                Method = Str(j, "authMethod"), Email = Str(j, "email"), Org = Str(j, "orgName"), Plan = Str(j, "subscriptionType")
            };
        }

        public static string Describe(Auth a)
        {
            return a.Method + (a.Plan.Length > 0 ? "（" + a.Plan + "）" : "") + (a.Email.Length > 0 ? " " + a.Email : "");
        }

        // 日報を作れない理由（作れるなら null）
        public static string Problem(Settings s, out string exe, out Auth auth)
        {
            auth = null;
            exe = FindExe(s);
            if (exe == null) return "Claude Code（claude.exe）が見つかりません";
            try { auth = Status(exe); }
            catch (Exception ex) { return "Claude Code のログイン状態を確かめられません: " + ex.Message; }
            if (!auth.LoggedIn) return "Claude Code にログインしていません";
            if (auth.Method != "claude.ai") return "Claude Code のログインが claude.ai のプランではありません（" + auth.Method + "）。API の従量課金になるため使いません";
            return null;
        }

        // 日報の文を作る。呼び出しは ai_calls.log に残す
        public static string Ask(string prompt, Settings s, string date, string note = "", string model = null)
        {
            string exe;
            Auth auth;
            string problem = Problem(s, out exe, out auth);
            if (problem != null) { WriteLog(date, exe, auth, null, 0, "失敗" + note + ": " + problem); throw new NeedLoginException(problem); }

            string args = "-p --safe-mode --tools \"\" --no-session-persistence --output-format json --model \"" + (model ?? s.Model) +
                          "\" --system-prompt \"" + SystemPrompt + "\"";
            Exception last = null;
            for (int i = 1; i <= 2; i++)
            {
                try
                {
                    Dictionary<string, object> j = ParseJson(RunExe(exe, args, prompt, 300000));
                    string text = Str(j, "result").Trim();
                    var models = new List<string>();
                    long tokens = 0;
                    object mu;
                    if (j.TryGetValue("modelUsage", out mu) && mu is Dictionary<string, object>)
                        foreach (KeyValuePair<string, object> kv in (Dictionary<string, object>)mu)
                        {
                            models.Add(kv.Key);
                            var u = kv.Value as Dictionary<string, object>;
                            if (u != null)
                                foreach (string k in new[] { "inputTokens", "outputTokens", "cacheReadInputTokens", "cacheCreationInputTokens" })
                                {
                                    object n;
                                    if (u.TryGetValue(k, out n) && n != null) tokens += Convert.ToInt64(n);
                                }
                        }
                    if (Str(j, "is_error") == "True" || text.Length == 0)
                    {
                        string msg = text.Length > 0 ? text : Str(j, "subtype");
                        WriteLog(date, exe, auth, string.Join(",", models), tokens, "失敗" + note + ": " + msg);
                        if (msg.IndexOf("login", StringComparison.OrdinalIgnoreCase) >= 0) throw new NeedLoginException("Claude Code: " + msg);
                        throw new Exception("Claude Code: " + msg);
                    }
                    WriteLog(date, exe, auth, string.Join(",", models), tokens, "成功" + note);
                    return text;
                }
                catch (NeedLoginException) { throw; }
                catch (Exception ex) { last = ex; }
                if (i < 2) Thread.Sleep(15000);
            }
            throw last;
        }

        // その日に AI を呼んだ時刻と、うまくいったか（日報画面の帯に印を付ける）
        public static List<KeyValuePair<DateTime, bool>> CallsOn(DateTime day)
        {
            var list = new List<KeyValuePair<DateTime, bool>>();
            try
            {
                if (!File.Exists(LogPath)) return list;
                string prefix = day.ToString("yyyy-MM-dd");
                foreach (string raw in Core.ReadText(LogPath).Split('\n'))
                {
                    if (!raw.StartsWith(prefix) && !raw.TrimStart('\uFEFF').StartsWith(prefix)) continue;
                    string[] c = raw.TrimStart('\uFEFF').TrimEnd('\r').Split('\t');
                    DateTime t;
                    if (c.Length >= 8 && DateTime.TryParseExact(c[0], "yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture, DateTimeStyles.None, out t))
                        list.Add(new KeyValuePair<DateTime, bool>(t, c[7].StartsWith("成功")));
                }
            }
            catch { }
            return list;
        }

        // ai_calls.log: 日時<TAB>日報の日付<TAB>元のログ<TAB>claude.exe<TAB>ログイン<TAB>モデル<TAB>トークン<TAB>結果
        static void WriteLog(string date, string exe, Auth auth, string models, long tokens, string result)
        {
            try
            {
                string line = string.Join("\t", new[] {
                    DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), date, Path.GetFileName(Core.CsvPath(date)), exe ?? "",
                    auth != null ? Describe(auth) : "", models ?? "", tokens.ToString(), result.Replace('\t', ' ').Replace('\r', ' ').Replace('\n', ' ') });
                File.AppendAllText(LogPath, line + "\r\n", new UTF8Encoding(true));
            }
            catch (Exception ex) { Core.Error("AI履歴", ex.Message); }
        }

        public static List<string[]> RecentLog(int n)
        {
            if (!File.Exists(LogPath)) return new List<string[]>();
            return Core.ReadText(LogPath).Split('\n').Select(l => l.TrimEnd('\r').Split('\t'))
                .Where(p => p.Length >= 8).Reverse().Take(n).ToList();
        }
    }

    // 日報の作成
    static class Reporter
    {
        static readonly Regex CsvRe = new Regex("^\"([^\"]*)\",\"([^\"]*)\",\"([^\"]*)\",(\\d+)");
        static readonly Regex DateRe = new Regex(@"^\d{4}-\d{2}-\d{2}$");
        static readonly CultureInfo Inv = CultureInfo.InvariantCulture;
        static string StatePath { get { return Path.Combine(Core.Dir, "report_state.txt"); } }

        public static bool IsDate(string s) { return s != null && DateRe.IsMatch(s); }
        public static int Minutes(int rows) { return rows * Core.IntervalSec / 60; }

        // 日報作成は同時に1つだけ（常駐側と --report が重ならないように）
        public static List<Result> Locked(Func<List<Result>> work)
        {
            using (var m = new Mutex(false, "WorkLogReport" + Core.Suffix))
            {
                bool got;
                try { got = m.WaitOne(TimeSpan.FromMinutes(3)); }
                catch (AbandonedMutexException) { got = true; }
                if (!got) throw new Exception("別の日報作成が終わらないため中止しました");
                try { return work(); }
                finally { m.ReleaseMutex(); }
            }
        }

        // 離席・除外プロセスを除いた記録
        // 記録の時刻と無操作の秒数（日報画面の帯に出す）
        public static List<KeyValuePair<DateTime, long>> ReadIdle(string date)
        {
            var list = new List<KeyValuePair<DateTime, long>>();
            string path = Core.CsvPath(date);
            if (!File.Exists(path)) return list;
            foreach (string raw in Core.ReadText(path).Split('\n'))
            {
                Match m = CsvRe.Match(raw.TrimEnd('\r'));
                DateTime t;
                if (m.Success && DateTime.TryParseExact(m.Groups[1].Value, "yyyy-MM-dd HH:mm:ss", Inv, DateTimeStyles.None, out t))
                    list.Add(new KeyValuePair<DateTime, long>(t, long.Parse(m.Groups[4].Value)));
            }
            return list;
        }

        public static List<Row> ReadActive(string date, Settings s)
        {
            string path = Core.CsvPath(date);
            if (!File.Exists(path)) return null;
            var list = new List<Row>();
            bool rest = BreakOn(date, s);
            foreach (string raw in Core.ReadText(path).Split('\n'))
            {
                Match m = CsvRe.Match(raw.TrimEnd('\r'));
                if (!m.Success) continue;
                DateTime t;
                if (!DateTime.TryParseExact(m.Groups[1].Value, "yyyy-MM-dd HH:mm:ss", Inv, DateTimeStyles.None, out t)) continue;
                string proc = m.Groups[2].Value;
                if (long.Parse(m.Groups[4].Value) >= s.IdleLimit) continue;
                if (rest && s.InBreak(t)) continue;   // 休憩時間の操作は日報に入れない（休憩を外した日は入れる）
                if (proc.Length == 0 || s.ExcludeProc.Contains(proc)) continue;
                list.Add(new Row { Time = t, Proc = proc, Title = m.Groups[3].Value });
            }
            return list;
        }

        // 利用者が「AIで直す」で出した訂正の指示（作り直すときにも AI に渡す）
        public static string NotesPath(string date) { return Path.Combine(Core.Dir, "訂正_" + date + ".txt"); }

        public static string ReadNotes(string date)
        {
            string p = NotesPath(date);
            return File.Exists(p) ? Core.ReadText(p).Trim() : "";
        }

        static string NotesPrompt(string date)
        {
            string notes = ReadNotes(date);
            return notes.Length == 0 ? "" :
                "利用者からの訂正（以前この日の日報に対して受けたもの。作業ログより優先して必ず反映する）:\n" + notes.Replace("\r\n", "\n") + "\n\n";
        }

        // アプリ×タイトル別の合計と、時刻順のまとまり（同じアプリ×タイトルが続いた区間）
        // ---- 開いていたウィンドウ（ログ\窓_日付.csv）----
        public class WinSpan { public string Proc, Title; public DateTime Start, End; }
        static readonly Regex WinRe = new Regex("^\"([^\"]*)\",\"(open|close)\",\"([^\"]*)\",\"([^\"]*)\"");
        public const int WinMinMinutes = 10;   // これより短く開いていたものは AI に渡さない

        // 開いていた区間（開いた〜閉じた）。閉じていないものは、今日なら今まで、ほかの日はその日の最後の記録まで
        public static List<WinSpan> WindowSpans(string date)
        {
            var list = new List<WinSpan>();
            string path = Core.WindowsPath(date);
            if (!File.Exists(path)) return list;
            var open = new Dictionary<string, KeyValuePair<DateTime, string[]>>();
            DateTime last = DateTime.MinValue;
            foreach (string raw in Core.ReadText(path).Split('\n'))
            {
                Match m = WinRe.Match(raw.TrimEnd('\r').TrimStart('\uFEFF'));
                DateTime t;
                if (!m.Success || !DateTime.TryParseExact(m.Groups[1].Value, "yyyy-MM-dd HH:mm:ss", Inv, DateTimeStyles.None, out t)) continue;
                if (t > last) last = t;
                string key = m.Groups[3].Value + "\t" + m.Groups[4].Value;
                KeyValuePair<DateTime, string[]> o;
                if (m.Groups[2].Value == "open") { if (!open.ContainsKey(key)) open[key] = new KeyValuePair<DateTime, string[]>(t, new[] { m.Groups[3].Value, m.Groups[4].Value }); }
                else if (open.TryGetValue(key, out o))
                {
                    if (t > o.Key) list.Add(new WinSpan { Proc = o.Value[0], Title = o.Value[1], Start = o.Key, End = t });
                    open.Remove(key);
                }
            }
            DateTime end = date == DateTime.Today.ToString("yyyy-MM-dd") ? DateTime.Now : last;
            foreach (var kv in open)
                if (end > kv.Value.Key) list.Add(new WinSpan { Proc = kv.Value.Value[0], Title = kv.Value.Value[1], Start = kv.Value.Key, End = end });
            return list.OrderBy(x => x.Start).ToList();
        }

        // AI に渡す「開いていたウィンドウ」（since があればそれより後の分だけ）
        static string WindowText(string date, Settings s, Dictionary<string, Tasks.TitleTask> map, List<TaskItem> tasks, DateTime? since)
        {
            List<WinSpan> spans = WindowSpans(date);
            if (since.HasValue)
                spans = spans.Where(x => x.End > since.Value).Select(x => new WinSpan { Proc = x.Proc, Title = x.Title, Start = x.Start < since.Value ? since.Value : x.Start, End = x.End }).ToList();
            spans = spans.Where(x => (x.End - x.Start).TotalMinutes >= Math.Max(WinMinMinutes, s.MinMinutes) && !s.ExcludeProc.Contains(x.Proc))
                .OrderByDescending(x => x.End - x.Start).Take(60).OrderBy(x => x.Start).ToList();
            if (spans.Count == 0) return "";
            var b = new StringBuilder("\n\n開いていたウィンドウ（開いた〜閉じた。閉じるまでその作業をしていたとみなす。前面で使っていた作業を優先し、ほかに作業のない時間をこの作業として書く）:\n");
            foreach (WinSpan x in spans)
            {
                string t;
                if (!s.MaskProc.TryGetValue(x.Proc, out t)) t = x.Title;
                foreach (string w in s.MaskWords) t = t.Replace(w, "＊＊");
                string task = Tasks.Resolve(map, tasks, x.Proc, x.Title);
                b.Append(x.Start.ToString("HH:mm") + "-" + x.End.ToString("HH:mm") + " [" + x.Proc + "] " + t + (task != null ? " 〔タスク: " + task + "〕" : "") + "\n");
            }
            return b.ToString().TrimEnd('\n');
        }

        static string LogText(List<Row> active, Settings s, DateTime? since = null)
        {
            var groups = new Dictionary<string, int>();
            var segs = new List<Seg>();
            List<TaskItem> tasks = Tasks.Load();
            if (tasks.Any(x => !x.Done)) Tasks.PredictFor(active, s);   // 対応のないタイトルは AI が予測しておく（手本が足りなければ何もしない）
            Dictionary<string, Tasks.TitleTask> map = Tasks.LoadMap();
            foreach (Row a in active)
            {
                string t;
                if (!s.MaskProc.TryGetValue(a.Proc, out t)) t = a.Title;
                foreach (string w in s.MaskWords) t = t.Replace(w, "＊＊");
                string task = Tasks.Resolve(map, tasks, a.Proc, a.Title);   // 伏せる前のタイトルで判定
                string k = a.Proc + "\t" + t + (task != null ? " 〔タスク: " + task + "〕" : "");
                int n;
                groups.TryGetValue(k, out n);
                groups[k] = n + 1;

                Seg last = segs.Count > 0 ? segs[segs.Count - 1] : null;
                if (last != null && last.Key == k && a.Time <= last.End.AddSeconds(Core.IntervalSec)) { last.End = a.Time.AddSeconds(Core.IntervalSec); last.Rows++; }
                else segs.Add(new Seg { Key = k, Start = a.Time, End = a.Time.AddSeconds(Core.IntervalSec), Rows = 1 });
            }
            string summary = string.Join("\n", groups
                .Select(g => new { Min = Minutes(g.Value), Key = g.Key })
                .Where(g => g.Min >= s.MinMinutes)
                .OrderByDescending(g => g.Min).Take(80)
                .Select(g => { string[] p = g.Key.Split(new[] { '\t' }, 2); return string.Format("{0}分 [{1}] {2}", g.Min, p[0], p[1]); }));
            // 長すぎるときは1分だけの切り替わりを省く
            if (segs.Count > 150) segs = segs.Where(x => x.Rows >= 2).ToList();
            if (segs.Count > 200) segs = segs.Where(x => Minutes(x.Rows) >= s.MinMinutes).ToList();
            string timeline = string.Join("\n", segs.Select(x =>
            {
                string[] p = x.Key.Split(new[] { '\t' }, 2);
                return string.Format("{0}-{1} [{2}] {3}", x.Start.ToString("HH:mm"), x.End.ToString("HH:mm"), p[0], p[1]);
            }));
            string wins = active.Count > 0 ? WindowText(active[0].Time.ToString("yyyy-MM-dd"), s, map, tasks, since) : "";
            return "合計時間:\n" + summary + "\n\n時刻順:\n" + timeline + wins;
        }

        // 作業の行（「・」で始まる行。「・なし」は除く）に通し番号 [1] [2] … を付け直す。「3番を直して」と指示できるように
        static readonly Regex ItemRe = new Regex(@"^(\s*)・\s*(?:\[\d+\]\s*)?(.*)$");
        static readonly Regex NumberedRe = new Regex(@"^\s*・\s*\[(\d+)\]\s*(.*)$");

        public static string Number(string text)
        {
            int n = 0;
            string[] lines = text.Replace("\r\n", "\n").Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                Match m = ItemRe.Match(lines[i]);
                if (!m.Success) continue;
                string rest = m.Groups[2].Value.TrimEnd();
                lines[i] = m.Groups[1].Value + "・" + (rest.Length == 0 || rest.StartsWith("なし") ? rest : "[" + ++n + "] " + rest);
            }
            return string.Join("\r\n", lines);
        }

        // コピーするときは番号を外す
        public static string Unnumber(string text)
        {
            return Regex.Replace(text, @"(?m)^(\s*)・\s*\[\d+\]\s*", "$1・");
        }

        // 指示の「3番」「[3]」を、その行の中身に置き換える（訂正の記録に残すとき。あとで番号が変わっても指す行がわかるように）
        public static string ResolveNumbers(string instruction, string report)
        {
            var map = new Dictionary<int, string>();
            foreach (string l in report.Replace("\r\n", "\n").Split('\n'))
            {
                Match m = NumberedRe.Match(l);
                if (m.Success) map[int.Parse(m.Groups[1].Value)] = m.Groups[2].Value.Trim();
            }
            return Regex.Replace(instruction, @"\[(\d+)\]|(\d+)\s*番目?", m =>
            {
                string c;
                int k = int.Parse(m.Groups[1].Success ? m.Groups[1].Value : m.Groups[2].Value);
                return map.TryGetValue(k, out c) ? "「" + c + "」" : m.Value;
            });
        }

        // ---- 作成済みの日報（md）: AI が作った・追記した・作り直した日報と、「AIで直す」で確定した日報。手で保存したときは書き換えない ----
        // 作り直す・AIで直すときに AI に渡す。「md にあって真ん中の日報にない作業 = 利用者が消したもの」とわかるように持っておく
        static string MdDir { get { string d = Path.Combine(Core.Dir, "md"); Directory.CreateDirectory(d); return d; } }

        public static string MdPath(string date) { return Path.Combine(MdDir, "日報_" + date + ".md"); }

        public static void SaveMd(string date, string report)
        {
            try
            {
                var b = new StringBuilder("<!-- AI が作った日報（" + DateTime.Now.ToString("yyyy-MM-dd HH:mm") + "）。作り直す・AIで直すときの参考 -->\r\n");
                foreach (string raw in SplitLines(Unnumber(report.Trim())))
                {
                    string l = raw.Trim();
                    if (l.StartsWith("日報")) b.Append("# " + l + "\r\n");
                    else if (l.StartsWith("■")) b.Append("\r\n## " + l.Substring(1).Trim() + "\r\n");
                    else if (l.StartsWith("・")) b.Append("- " + l.Substring(1).Trim() + "\r\n");
                    else if (l.Length > 0) b.Append(l + "\r\n");
                }
                Core.WriteText(MdPath(date), b.ToString());
            }
            catch (Exception ex) { Core.Error("md", ex.Message); }
        }

        public static string LoadMd(string date)
        {
            string p = MdPath(date);
            return File.Exists(p) ? Core.ReadText(p).Trim() : "";
        }

        // 作り直す・AIで直すで共通の決まり（真ん中の日報がメイン）
        static string MainRules()
        {
            return
                "- 「真ん中の日報」は利用者が画面で書いた・直した日報で、これがメイン。書いてある作業はすべて残す。" +
                "書式の正しい行（「■工程」の下の「・HH:MM〜HH:MM 作業内容」、「■軽作業」の下の「・HH:MM 作業内容」）は一字も変えない\n" +
                "- 時刻と内容が適当に書かれた行（例:「9時から10時半 資料作成」「資料作成 13:00」「13:00-14:00 打合せ」「１３：００ 電話」）は、内容の言葉をできるだけそのまま使って正しい形に直す\n" +
                "- 内容だけで時刻のない行は、作業ログからその作業の時刻を推定して入れる。ログで見つからなければ前後の行の時刻から決める\n" +
                "- 時間の幅があり20分程度を超える作業は「■工程」、短い作業は「■軽作業」に入れる。それぞれ時刻の早い順\n" +
                "- 工程は時間が途切れないようにつなげる（前の工程の終わり＝次の工程の始め）。軽作業しかしていない時間は、工程に「・HH:MM〜HH:MM 軽作業」と書く\n" +
                "- 「作成済みの日報（md）」は前に AI が作った日報。md にあって真ん中の日報にない作業は、利用者が要らないので消したもの。足し直さない\n" +
                "- 行頭の [番号] は付けない（番号はプログラムが振り直す）\n";
        }

        static string MdPrompt(string date)
        {
            string md = LoadMd(date);
            return md.Length > 0 ? "作成済みの日報（md）:\n" + md + "\n\n" : "（作成済みの日報（md）はありません）\n\n";
        }

        // 真ん中の日報の、書式の正しい行（AI が落としたらプログラムが戻す）。値は軽作業か
        public static List<KeyValuePair<string, bool>> CorrectLines(string body)
        {
            var list = new List<KeyValuePair<string, bool>>();
            string section = null;
            foreach (string raw in SplitLines(Unnumber(body)))
            {
                string l = raw.Trim();
                if (l.StartsWith("■")) { section = l == "■工程" ? "工程" : l == "■軽作業" ? "軽作業" : null; continue; }
                if (section == "工程" && GoodRange.IsMatch(l)) list.Add(new KeyValuePair<string, bool>(l, false));
                else if (section == "軽作業" && GoodLight.IsMatch(l)) list.Add(new KeyValuePair<string, bool>(l, true));
            }
            return list;
        }

        // AI の出した日報に、真ん中の日報の正しい行がなければ時刻の位置へ戻す。戻した数を返す
        public static int KeepCorrect(string body, List<string> lines)
        {
            int n = 0;
            foreach (KeyValuePair<string, bool> k in CorrectLines(body))
                if (!lines.Any(x => Unnumber(x).Trim() == k.Key)) { InsertItem(lines, k.Value, k.Key); n++; }
            return n;
        }

        // 日報リンクの開く先: {yyyy} {MM} {dd} を日付に置き換え、URL・ファイル・フォルダのどれでもなければ https:// を付ける
        public static string LinkTarget(string link, DateTime day)
        {
            string t = (link ?? "").Trim().Trim('"');
            if (t.Length == 0) return "";
            t = t.Replace("{yyyy}", day.ToString("yyyy")).Replace("{MM}", day.ToString("MM")).Replace("{dd}", day.ToString("dd"));
            if (Regex.IsMatch(t, @"^[a-zA-Z][a-zA-Z0-9+.-]*://")) return t;
            string path = Environment.ExpandEnvironmentVariables(t);
            if (Regex.IsMatch(path, @"^([a-zA-Z]:\\|\\\\)") || File.Exists(path) || Directory.Exists(path)) return path;
            return "https://" + t;
        }

        // AI の返事を「返事」と「日報の案」に分ける（案がなければ null）
        public static string ParseChat(string answer, out string reply)
        {
            answer = (answer ?? "").Trim();
            Match r = Regex.Match(answer, @"<返事>\s*(.*?)\s*</返事>", RegexOptions.Singleline);
            Match d = Regex.Match(answer, @"<日報>\s*(.*?)\s*</日報>", RegexOptions.Singleline);
            string report = d.Success ? d.Groups[1].Value.Trim() : null;
            if (report != null && !SplitLines(report).Any(l => l.Trim().StartsWith("■"))) report = null;
            if (r.Success) reply = r.Groups[1].Value.Trim();
            else
            {
                string rest = Regex.Replace(answer, @"<日報>.*?</日報>", "", RegexOptions.Singleline).Trim();
                // 形を守らずに日報だけを返したとき
                if (!d.Success && SplitLines(rest).Any(l => l.Trim().StartsWith("■"))) { report = rest; rest = ""; }
                reply = rest;
            }
            if (reply.Length == 0) reply = report != null ? "日報を直しました。青字のところを確かめてください。" : "（返事がありませんでした）";
            return report;
        }

        // 「AIで直す」（会話）: 利用者の頼みと、これまでの会話・真ん中の日報・md・作業ログを渡す。返事と、直した日報の案（直さないときは null）を返す
        public static string Chat(string date, string current, List<KeyValuePair<bool, string>> history, string message, string notes, Settings s, out string reply)
        {
            string[] lines = current.Replace("\r\n", "\n").Split('\n');
            bool hasHeader = lines.Length > 0 && lines[0].StartsWith("日報");
            string header = hasHeader ? lines[0] : null;
            string body = string.Join("\n", hasHeader ? lines.Skip(1) : lines).Trim();
            List<Row> active = ReadActive(date, s);
            var talk = new StringBuilder();
            foreach (KeyValuePair<bool, string> h in history.Skip(Math.Max(0, history.Count - 12)))
                talk.Append(h.Key ? "利用者: " : "あなた: ").Append(h.Value.Replace("\r\n", "\n").Trim()).Append("\n");

            string prompt =
                "あなたは日報作りを手伝うアシスタントです。利用者と会話しながら、画面の真ん中に出ている日報を直します。\n" +
                "利用者の頼みに合わせて日報を直してください。頼まれていなくても、次の決まりに合わない行があれば一緒に直してよい。\n\n" +
                "決まり:\n" + MainRules() +
                "- 利用者の頼みは決まりより優先する（消してと言われた行は消す、など）\n" +
                "- 行頭の [番号] は作業の通し番号。利用者の「3番」などはこの番号の行のこと\n" +
                "- 足す作業は、利用者に頼まれたもの、または作業ログにあって真ん中の日報にも md にもないものだけ\n" +
                "- 「休憩」の行は、頼まれない限りそのまま残す\n" +
                BreakRule(s, date) + PlanRule(date, s) + "\n" +
                "返事の形（この形だけを出力）:\n" +
                "<返事>\n利用者への短い返事（何をどう直したか。わからないことがあれば質問する）\n</返事>\n" +
                "<日報>\n直した日報の本文（日報を直すときだけ書く。「■工程」「■軽作業」の形。見出しの日付行は不要）\n</日報>\n\n" +
                (notes.Trim().Length > 0 ? "これまでの訂正（これも守る）:\n" + notes.Trim().Replace("\r\n", "\n") + "\n\n" : "") +
                Tasks.Prompt() + Afk.Prompt(date) +
                (talk.Length > 0 ? "これまでの会話:\n" + talk + "\n" : "") +
                "利用者の今の頼み:\n" + message.Trim() + "\n\n" +
                "真ん中の日報（メイン。番号つき）:\n" + body + "\n\n" +
                MdPrompt(date) +
                (active != null && active.Count > 0 ? "作業ログ（離席時間は除いています）\n" + LogText(active, s) : "（この日の作業ログはありません）");

            string answer = ClaudeCli.Ask(prompt, s, date, "（AIで直す）", s.FixModel);
            string report = ParseChat(answer, out reply);
            if (report == null) return null;
            List<string> result = SplitLines((header != null ? header + "\n\n" : "") + Unnumber(StripHeader(report)));
            if (!message.Contains("休憩")) EnsureBreak(result, date, s);   // AI がうっかり休憩の行を落としたとき
            Tidy(result, date, s);
            return Number(string.Join("\n", result).TrimEnd() + "\n");
        }

        // ---- 会話の記録（md\会話_日付.md）----
        static string ChatPath(string date) { return Path.Combine(MdDir, "会話_" + date + ".md"); }

        // 会話（true が利用者）
        public static List<KeyValuePair<bool, string>> LoadChat(string date)
        {
            var list = new List<KeyValuePair<bool, string>>();
            string p = ChatPath(date);
            if (!File.Exists(p)) return list;
            foreach (Match m in Regex.Matches(Core.ReadText(p).Replace("\r\n", "\n"), @"^### (あなた|AI)[^\n]*\n(.*?)(?=^### |\z)", RegexOptions.Singleline | RegexOptions.Multiline))
                list.Add(new KeyValuePair<bool, string>(m.Groups[1].Value == "あなた", m.Groups[2].Value.Trim()));
            return list;
        }

        public static void AddChat(string date, bool user, string text)
        {
            try
            {
                string p = ChatPath(date);
                string head = File.Exists(p) ? "" : "# " + date + " の日報についての会話（作業ログ・日報の「AIで直す」）\r\n\r\n";
                File.AppendAllText(p, head + "### " + (user ? "あなた" : "AI") + "（" + DateTime.Now.ToString("HH:mm") + "）\r\n" +
                    text.Trim().Replace("\r\n", "\n").Replace("\n", "\r\n") + "\r\n\r\n", new UTF8Encoding(true));
            }
            catch (Exception ex) { Core.Error("会話", ex.Message); }
        }

        public static void ClearChat(string date)
        {
            try { string p = ChatPath(date); if (File.Exists(p)) File.Delete(p); }
            catch (Exception ex) { Core.Error("会話", ex.Message); }
        }

        // 直した日報と訂正の記録を保存する（前の日報は old フォルダへ）
        public static void SaveRevision(string date, string text, string notes)
        {
            string path = Core.ReportPath(date);
            Backup(path);
            Core.WriteText(path, text);
            SaveNotes(date, notes);
            SaveMd(date, text);
        }

        public static void SaveNotes(string date, string notes)
        {
            string p = NotesPath(date);
            if (notes.Trim().Length > 0) Core.WriteText(p, notes.Trim() + "\r\n");
            else if (File.Exists(p)) File.Delete(p);
        }

        // 1日分の日報を作る。日報がもうあるときは、今ある行を変えずに書式の正しくない行だけ直し、作業ログにしかない作業を足す（Combine）
        public static Result Make(string date, Settings s)
        {
            var r = new Result { Date = date };
            List<Row> active = ReadActive(date, s);
            if (active == null || active.Count == 0) { r.NoData = true; r.Message = date + " の作業記録がありません"; return r; }

            string prompt =
                "以下はPCの作業ログです。「合計時間」はアプリ×ウィンドウタイトルごとの使用時間、\n" +
                "「時刻順」は使っていたアプリとウィンドウタイトルの移り変わりです（離席時間は除いています）。\n" +
                "これをもとに、今日の日報を作ってください。\n\n" +
                "出力の形（この形だけを出力。前置き・説明は不要）:\n" +
                "■工程\n" +
                "・HH:MM〜HH:MM 作業内容\n" +
                "■軽作業\n" +
                "・HH:MM 作業内容\n\n" +
                "条件:\n" +
                "- 同じ業務のまとまりを1つの工程にし、始めた時刻と終えた時刻を書く。途中に短い別作業が挟まっても、同じ業務が続いていれば1つの工程にまとめる\n" +
                "- 合計20分程度以下の作業は工程にせず軽作業に入れ、終えた時刻だけを書く\n" +
                "- 工程は時間が途切れないようにつなげる（前の工程の終わり＝次の工程の始め）。軽作業しかしていない時間は、工程に「・HH:MM〜HH:MM 軽作業」と書く\n" +
                "- 工程・軽作業とも時刻の早い順。時刻は作業ログに出てくる時刻を使う\n" +
                "- 作業内容は20文字程度。ファイル名やページ名から業務内容を具体的に推定する（例:「〇〇工区の数量計算」）\n" +
                "- 推定できないもの、私用と思われるものは省く\n" +
                "- 「時刻のない本人のメモ」があれば、作業ログになくても業務として必ず入れる\n" +
                BreakRule(s, date) + PlanRule(date, s) +
                "- 該当がなければ「・なし」\n\n" +
                NotesPrompt(date) + Tasks.Prompt() + Afk.Prompt(date) + ManualPrompt(date) + LogText(active, s);

            string header = Header(date, active);
            string path = Core.ReportPath(date);
            string existing = File.Exists(path) ? Core.ReadText(path) : "";
            // 日報がある（AI に失敗したときの案内だけの日は、日報がないものとして作る）
            bool merge = existing.Contains("■") && !existing.Contains("急ぐときは下の文を Claude に貼ってください");
            // 日報がもうある日は、真ん中の日報（保存済み）をメインに、作成済みの日報（md）と作業ログを見て整える
            if (merge) return Combine(date, s, r, existing, header, active);

            string body;
            try
            {
                body = Number(ClaudeCli.Ask(prompt, s, date, merge ? "（作り直し）" : ""));
                r.Ok = true;
            }
            catch (NeedLoginException ex)
            {
                // 失敗扱いにして、ログインを直したあとの自動作成で作り直す（今の日報があれば消さない）
                r.Ok = false;
                r.NeedLogin = true;
                r.Message = ex.Message + "（AIの接続先）";
                if (merge) return r;
                body = "（" + r.Message + "。直すと自動で作り直します。急ぐときは下の文を Claude に貼ってください）\n\n" + prompt;
            }
            catch (Exception ex)
            {
                r.Ok = false;
                r.Message = ex.Message;
                Core.Error("日報 " + date, "AI失敗: " + ex.Message);
                if (merge) return r;
                body = "（AIで日報を作れませんでした: " + ex.Message + "）\n" +
                       "（あとで自動でもう一度試します。急ぐときは下の文を Claude に貼ってください）\n\n" + prompt;
            }


            lock (reportLock)
            {
                Backup(path);
                List<string> lines = SplitLines(header + "\n\n" + body.Trim());
                if (r.Ok) { ApplyManual(lines, date); EnsureBreak(lines, date, s); Tidy(lines, date, s); }
                string text = string.Join("\n", lines).TrimEnd() + "\n";
                Core.WriteText(path, r.Ok ? Number(text) : Regex.Replace(text, "\r?\n", "\r\n"));
            }
            if (r.Ok) SaveMd(date, Core.ReadText(path));
            if (r.Ok) ApplyAllAfk(date);   // メモのある離席は、プログラムが時刻どおりに足す
            SaveState(date, active[active.Count - 1].Time, r.Ok);
            return r;
        }

        // ---- 作り直し（日報がもうあるとき）----

        static readonly Regex GoodRange = new Regex(@"^・\d{2}:\d{2}〜\d{2}:\d{2} \S");
        static readonly Regex GoodLight = new Regex(@"^・\d{2}:\d{2} \S");
        // 作り直す・今日の分を今すぐ作る（日報がもうあるとき）: 真ん中の日報（保存済みの今の日報）をメインに、作成済みの日報（md）と作業ログを見て、
        // FixModel が書式を直し・時刻のない行に時刻を入れ・作業ログにしかない作業を足す。書式の正しい行は AI が落としてもプログラムが戻し、番号は振り直す
        static Result Combine(string date, Settings s, Result r, string existing, string header, List<Row> active)
        {
            string path = Core.ReportPath(date);
            string body = Unnumber(StripHeader(existing));
            string prompt =
                "日報を整えて、足りない作業を足してください。出力は整えた日報の本文だけ（前置き・説明・見出しの日付行は不要）。\n\n" +
                "出力の形:\n■工程\n・HH:MM〜HH:MM 作業内容\n■軽作業\n・HH:MM 作業内容\n\n" +
                "条件:\n" + MainRules() +
                "- 作業ログにあって、真ん中の日報にも作成済みの日報（md）にもない作業だけを足す。推定できないもの・私用と思われるものは足さない。足す作業は今ある工程と時間が重ならないようにする\n" +
                "- 「時刻のない本人のメモ」のうち日報にないものは足す\n" +
                BreakRule(s, date) + PlanRule(date, s) + "\n" +
                NotesPrompt(date) + Tasks.Prompt() + Afk.Prompt(date) + ManualPrompt(date) +
                "真ん中の日報（メイン）:\n" + body + "\n\n" +
                MdPrompt(date) +
                "作業ログ（離席時間は除いています）\n" + LogText(active, s);

            string answer;
            try
            {
                answer = StripHeader(ClaudeCli.Ask(prompt, s, date, "（作り直し）", s.FixModel));
                if (!SplitLines(answer).Any(l => l.Trim().StartsWith("■"))) throw new Exception("AI が日報の形で返しませんでした");
            }
            catch (NeedLoginException ex)
            {
                r.Ok = false;
                r.NeedLogin = true;
                r.Message = ex.Message + "（日報は変更なし）";
                return r;
            }
            catch (Exception ex)
            {
                r.Ok = false;
                r.Message = "作り直しエラー: " + ex.Message + "（日報は変更なし）";
                Core.Error("作り直し " + date, ex.Message);
                return r;
            }

            List<string> lines = SplitLines(Unnumber(answer)).Select(l => l.TrimEnd()).ToList();
            int restored = KeepCorrect(body, lines);
            var before = new HashSet<string>(SplitLines(body).Select(l => l.Trim()));
            int changed = lines.Count(l => l.Trim().StartsWith("・") && !before.Contains(l.Trim()));

            lock (reportLock)
            {
                List<string> result = SplitLines(header + "\n\n" + string.Join("\n", lines).Trim());
                EnsureBreak(result, date, s);
                Tidy(result, date, s);
                string text = Number(string.Join("\n", result).TrimEnd() + "\n");
                // AI を待つあいだに手で直された・離席メモが足されたときは、上書きせずに案を old に残す
                if (!File.Exists(path) || Norm(Core.ReadText(path)) != Norm(existing))
                {
                    SaveOld(date, "作り直し案", text);
                    r.Ok = false;
                    r.Message = "日報が別に更新されたため変更なし（案は old フォルダ）";
                    return r;
                }
                if (Norm(text) != Norm(existing))
                {
                    Backup(path);
                    Core.WriteText(path, text);
                }
                SaveMd(date, text);
            }
            ApplyAllAfk(date);
            SaveState(date, active[active.Count - 1].Time, true);
            r.Ok = true;
            r.Message = changed == 0 ? "変更なし" : "変更 " + changed + " 行" + (restored > 0 ? "（元の行 " + restored + " 件を復元）" : "");
            return r;
        }

        static string Norm(string t) { return t.Replace("\r\n", "\n").Trim(); }

        // 先頭の「日報 …」の見出し行を外した本文（改行は \n）
        static string StripHeader(string t)
        {
            List<string> lines = SplitLines(t.Trim());
            if (lines.Count > 0 && lines[0].StartsWith("日報")) lines.RemoveAt(0);
            return string.Join("\n", lines).Trim();
        }

        static void SaveOld(string date, string kind, string text)
        {
            string dir = Path.Combine(Core.Dir, "old");
            Directory.CreateDirectory(dir);
            Core.WriteText(Path.Combine(dir, "日報_" + date + "_" + kind + "_" + DateTime.Now.ToString("yyyyMMdd-HHmmss") + ".txt"), text);
        }

        // ---- 休憩: 日報に「・12:00〜13:00 休憩」の行を残す。日報画面のチェックを外した日は 休憩なし_yyyy-MM-dd.txt を置く ----

        static string NoBreakPath(string date) { return Path.Combine(Core.Dir, "休憩なし_" + date + ".txt"); }

        public static bool BreakOn(string date, Settings s) { return s.HasBreak && !File.Exists(NoBreakPath(date)); }

        public static string BreakLine(Settings s) { return "・" + s.BreakText + " 休憩"; }

        public static bool HasBreakLine(string text, Settings s)
        {
            string b = BreakLine(s);
            return SplitLines(text).Any(l => Unnumber(l).Trim() == b);
        }

        public static void SetBreakFlag(string date, bool on)
        {
            if (on) { if (File.Exists(NoBreakPath(date))) File.Delete(NoBreakPath(date)); }
            else Core.WriteText(NoBreakPath(date), "この日は休憩なし（日報画面の「休憩」のチェックを外した）\r\n");
        }

        // チェックを付けた・外した: 日報の休憩の行を足す・消す
        public static void SetBreak(string date, bool on, Settings s)
        {
            SetBreakFlag(date, on);
            string path = Core.ReportPath(date);
            lock (reportLock)
            {
                if (!File.Exists(path)) return;
                string text = Core.ReadText(path);
                if (!text.Contains("■")) return;
                List<string> lines = SplitLines(text);
                if (on) EnsureBreak(lines, date, s);
                else lines.RemoveAll(l => Unnumber(l).Trim() == BreakLine(s));
                Core.WriteText(path, Number(string.Join("\n", lines).TrimEnd() + "\n"));
            }
        }

        static readonly Regex LightRe = new Regex(@"^・(\d{1,2}:\d{2})\s");

        // 休憩の行がなければ足す。休憩の時間に作業が書いてあるとき（利用者が書いた）は足さない
        // ---- 工程の時間をつなげる: 時刻の順に並べ、前の工程の終わり＝次の工程の始めにする。すき間が GapLight 分以上なら「軽作業」の工程で埋め、
        // それより短いすき間・重なりは前の工程の終わりを動かす（前が動かさない行＝休憩・離席メモ・手で追記した行なら、次の工程の始めを動かす）----
        public const int GapLight = 15;
        public const string LightOnly = "軽作業";

        class Proc { public TimeSpan A, B; public string Rest = "", Orig; public bool Fixed; }

        static string Hm(TimeSpan t) { return ((int)t.TotalHours).ToString("00") + ":" + t.Minutes.ToString("00"); }

        // 動かさない工程の行（休憩・離席メモ・手で追記した行）
        static HashSet<string> FixedLines(string date, Settings s, string extra = null)
        {
            var set = new HashSet<string>(Afk.Load(date).Select(e => e.Line).Where(x => x.Length > 0));
            foreach (KeyValuePair<bool, string> m in LoadManual(date)) set.Add(m.Value);
            set.Add(BreakLine(s));
            if (!string.IsNullOrEmpty(extra)) set.Add(extra);
            return set;
        }

        // 直した行の数を返す
        public const string OutsideWork = "現場（外勤）";

        // 日報の工程をそろえる（勤務予定の日は、時刻を勤務予定の中に収め、始業・終業までのすき間を「現場（外勤）」にする）
        static int Tidy(List<string> lines, string date, Settings s, string extra = null)
        {
            DateTime day = DateTime.ParseExact(date, "yyyy-MM-dd", CultureInfo.InvariantCulture);
            bool plan = s.PlanOn(day);
            return ConnectProcesses(lines, FixedLines(date, s, extra), plan ? s.PlanFrom : (TimeSpan?)null, plan ? s.PlanTo : (TimeSpan?)null);
        }

        // AI への指示（勤務予定の日だけ）
        static string PlanRule(string date, Settings s)
        {
            DateTime day = DateTime.ParseExact(date, "yyyy-MM-dd", CultureInfo.InvariantCulture);
            return s.PlanOn(day)
                ? "- 勤務予定は " + s.PlanText + "。この外の時刻は書かない。始業から最初の PC の作業まで・最後の作業から終業までに操作がない時間は「" + OutsideWork + "」の工程として書く\n"
                : "";
        }

        // from・to（勤務予定）があれば、工程・軽作業の時刻をその中に収め、始業・終業までのすき間を埋める
        public static int ConnectProcesses(List<string> lines, ICollection<string> fixedLines, TimeSpan? from = null, TimeSpan? to = null)
        {
            int clamped = from.HasValue && to.HasValue ? ClampLight(lines, from.Value, to.Value) : 0;
            int h = lines.FindIndex(l => l.Trim().StartsWith("■") && l.Contains("工程"));
            if (h < 0) return 0;
            int end = h + 1;
            while (end < lines.Count && !lines[end].Trim().StartsWith("■")) end++;
            var items = new List<Proc>();
            var others = new List<string>();
            int blanks = 0;
            for (int i = h + 1; i < end; i++)
            {
                string l = Unnumber(lines[i]).Trim();
                Match m = RangeRe.Match(l);
                if (m.Success) items.Add(new Proc { A = Clock(m.Groups[1].Value), B = Clock(m.Groups[2].Value), Rest = m.Groups[3].Value, Orig = l, Fixed = fixedLines.Contains(l) });
                else if (l.Length == 0) blanks++;
                else if (l != "・なし") others.Add(l);
            }
            if (from.HasValue && to.HasValue)
            {
                foreach (Proc x in items) { if (x.A < from.Value) x.A = from.Value; if (x.B > to.Value) x.B = to.Value; }
                items = items.Where(x => x.B > x.A).ToList();
            }
            if (items.Count == 0)
            {
                if (clamped == 0) return 0;
                lines.RemoveRange(h + 1, end - h - 1);
                lines.InsertRange(h + 1, others.Concat(Enumerable.Repeat("", blanks)));
                return clamped;
            }
            items = items.OrderBy(x => x.A).ThenBy(x => x.B).ToList();
            var result = new List<Proc> { items[0] };
            for (int i = 1; i < items.Count; i++)
            {
                Proc prev = result[result.Count - 1], next = items[i];
                TimeSpan gap = next.A - prev.B;
                if (gap >= TimeSpan.FromMinutes(GapLight))
                    result.Add(new Proc { A = prev.B, B = next.A, Rest = " " + LightOnly });
                else if (gap != TimeSpan.Zero)
                {
                    if (!prev.Fixed && next.A > prev.A) prev.B = next.A;
                    else if (!next.Fixed && prev.B < next.B) next.A = prev.B;
                }
                result.Add(next);
            }
            // 始業から最初の工程・最後の工程から終業まで: 短ければ工程を延ばし、GapLight 分以上なら「現場（外勤）」
            if (from.HasValue && to.HasValue)
            {
                Proc first = result[0], last = result[result.Count - 1];
                if (first.A > from.Value)
                {
                    if (first.A - from.Value >= TimeSpan.FromMinutes(GapLight)) result.Insert(0, new Proc { A = from.Value, B = first.A, Rest = " " + OutsideWork });
                    else if (!first.Fixed) first.A = from.Value;
                }
                if (last.B < to.Value)
                {
                    if (to.Value - last.B >= TimeSpan.FromMinutes(GapLight)) result.Add(new Proc { A = last.B, B = to.Value, Rest = " " + OutsideWork });
                    else if (!last.Fixed) last.B = to.Value;
                }
            }
            var built = result.Select(x => "・" + Hm(x.A) + "〜" + Hm(x.B) + x.Rest).ToList();
            int changed = clamped + result.Count(x => x.Orig == null) + result.Count(x => x.Orig != null && x.Orig != "・" + Hm(x.A) + "〜" + Hm(x.B) + x.Rest);
            lines.RemoveRange(h + 1, end - h - 1);
            lines.InsertRange(h + 1, built.Concat(others).Concat(Enumerable.Repeat("", blanks)));
            return changed;
        }

        // 軽作業の時刻を勤務予定の中に収める（直した数）
        static int ClampLight(List<string> lines, TimeSpan from, TimeSpan to)
        {
            int n = 0;
            bool light = false;
            for (int i = 0; i < lines.Count; i++)
            {
                string l = Unnumber(lines[i]).Trim();
                if (l.StartsWith("■")) { light = l.Contains("軽作業"); continue; }
                if (!light) continue;
                Match m = LightRe.Match(l);
                if (!m.Success) continue;
                TimeSpan t = Clock(m.Groups[1].Value), c = t < from ? from : t > to ? to : t;
                if (c == t) continue;
                lines[i] = "・" + Hm(c) + l.Substring(m.Groups[1].Index + m.Groups[1].Length);
                n++;
            }
            return n;
        }

        static bool EnsureBreak(List<string> lines, string date, Settings s)
        {
            string b = BreakLine(s);
            if (!BreakOn(date, s) || lines.Any(l => Unnumber(l).Trim() == b)) return false;
            if (!lines.Any(l => l.Trim().StartsWith("■"))) return false;
            foreach (string raw in lines)
            {
                string l = Unnumber(raw).Trim();
                Match m = RangeRe.Match(l);
                if (m.Success)
                {
                    if (Clock(m.Groups[1].Value) < s.BreakEnd && Clock(m.Groups[2].Value) > s.BreakStart) return false;
                    continue;
                }
                m = LightRe.Match(l);
                if (m.Success && Clock(m.Groups[1].Value) > s.BreakStart && Clock(m.Groups[1].Value) < s.BreakEnd) return false;
            }
            InsertItem(lines, false, b);
            return true;
        }

        // ---- 手で追記: 手入力_yyyy-MM-dd.txt（1=軽作業 0=工程<TAB>行）。日報がまだない日は、作るときに入れる ----

        static string ManualPath(string date) { return Path.Combine(Core.Dir, "手入力_" + date + ".txt"); }

        static List<KeyValuePair<bool, string>> LoadManual(string date)
        {
            var list = new List<KeyValuePair<bool, string>>();
            string p = ManualPath(date);
            if (!File.Exists(p)) return list;
            foreach (string raw in SplitLines(Core.ReadText(p)))
            {
                string[] c = raw.Split(new[] { '\t' }, 2);
                if (c.Length == 2 && c[1].Trim().StartsWith("・")) list.Add(new KeyValuePair<bool, string>(c[0] == "1", c[1].Trim()));
            }
            return list;
        }

        // 開始と終了があれば工程、片方だけ（開始だけ）なら軽作業
        public static string ManualItem(TimeSpan? start, TimeSpan? end, string what, out bool light)
        {
            light = !(start.HasValue && end.HasValue);
            what = what.Replace('\t', ' ').Replace('\r', ' ').Replace('\n', ' ').Trim();
            if (!light) return "・" + start.Value.ToString(@"hh\:mm") + "〜" + end.Value.ToString(@"hh\:mm") + " " + what;
            return "・" + (start ?? end).Value.ToString(@"hh\:mm") + " " + what;
        }

        // 日報に足したら true。日報がまだなければ控えだけ残し、作るときに入れる
        public static bool AddManual(string date, string item, bool light)
        {
            File.AppendAllText(ManualPath(date), (light ? "1" : "0") + "\t" + item + "\r\n", new UTF8Encoding(true));
            string path = Core.ReportPath(date);
            lock (reportLock)
            {
                if (!File.Exists(path)) return false;
                string text = Core.ReadText(path);
                if (!text.Contains("■")) return false;
                List<string> lines = SplitLines(text);
                if (!lines.Any(l => Unnumber(l).Trim() == item)) InsertItem(lines, light, item);
                Tidy(lines, date, Settings.Get(), item);
                Core.WriteText(path, Number(string.Join("\n", lines).TrimEnd() + "\n"));
            }
            return true;
        }

        static void ApplyManual(List<string> lines, string date)
        {
            foreach (KeyValuePair<bool, string> m in LoadManual(date))
                if (!lines.Any(l => Unnumber(l).Trim() == m.Value)) InsertItem(lines, m.Key, m.Value);
        }

        static string ManualPrompt(string date)
        {
            List<KeyValuePair<bool, string>> list = LoadManual(date);
            return list.Count == 0 ? "" :
                "本人が手で入れた作業（プログラムが日報に足すので、ここでは書かない。工程はこの時間をまたがず前後で分ける）:\n" +
                string.Concat(list.Select(m => m.Value.Substring(1) + "\n")) + "\n";
        }

        static readonly object reportLock = new object();

        static string Header(string date, List<Row> active)
        {
            return string.Format("日報 {0}（PC稼働 {1}〜{2}／操作時間 約{3}時間）", date,
                active[0].Time.ToString("HH:mm"), active[active.Count - 1].Time.ToString("HH:mm"), (Minutes(active.Count) / 60.0).ToString("0.#"));
        }

        static string BreakRule(Settings s, string date)
        {
            return BreakOn(date, s) ? "- " + s.BreakText + " は休憩時間。工程は休憩をまたがず前後で分ける。休憩の行（" + BreakLine(s) + "）はプログラムが入れるので書かなくてよい\n" : "";
        }

        // ---- 逐次追記: 今ある行（手で直した行も）は変えずに、行を足す ----

        static readonly Regex TimeRe = new Regex(@"(\d{1,2}):(\d{2})");

        static TimeSpan ItemTime(string line)
        {
            Match m = TimeRe.Match(Unnumber(line));
            return m.Success ? new TimeSpan(int.Parse(m.Groups[1].Value), int.Parse(m.Groups[2].Value), 0) : TimeSpan.MaxValue;
        }

        // 「■工程」か「■軽作業」の見出しの中に、時刻の順になる位置へ 1 行足す（「・なし」は消す）
        static void InsertItem(List<string> lines, bool light, string item)
        {
            Func<string, bool> isHead = l => l.Trim().StartsWith("■") && (light ? l.Contains("軽作業") : l.Contains("工程"));
            int h = lines.FindIndex(l => isHead(l));
            if (h < 0)
            {
                // 見出しがなければ作る（工程は軽作業の前）
                int lightHead = lines.FindIndex(l => l.Trim().StartsWith("■") && l.Contains("軽作業"));
                h = !light && lightHead >= 0 ? lightHead : lines.Count;
                lines.Insert(h, light ? "■軽作業" : "■工程");
            }
            int end = h + 1;
            while (end < lines.Count && !lines[end].Trim().StartsWith("■")) end++;
            for (int i = end - 1; i > h; i--)
                if (Unnumber(lines[i]).Trim() == "・なし") { lines.RemoveAt(i); end--; }
            int lastItem = h, pos = -1;
            TimeSpan time = ItemTime(item);
            for (int i = h + 1; i < end; i++)
            {
                if (!lines[i].Trim().StartsWith("・")) continue;
                if (pos < 0 && ItemTime(lines[i]) > time) pos = i;
                lastItem = i;
            }
            lines.Insert(pos >= 0 ? pos : lastItem + 1, item);
        }

        static List<string> SplitLines(string text) { return text.Replace("\r\n", "\n").Split('\n').ToList(); }

        static readonly Regex RangeRe = new Regex(@"^・(\d{1,2}:\d{2})〜(\d{1,2}:\d{2})(.*)$");

        static TimeSpan Clock(string hhmm) { string[] p = hhmm.Split(':'); return new TimeSpan(int.Parse(p[0]), int.Parse(p[1]), 0); }

        // AI が返した「■工程／■軽作業」の行を、今の日報に足す。足した・延ばした件数を返す
        // 続いている工程（始めの時刻が同じ行がある）は、行を足さずに終わりの時刻だけ延ばす（文言は今の行のまま）。離席メモの行は延ばさない
        static int MergeItems(List<string> lines, string added, ICollection<string> afkLines)
        {
            bool light = false;
            int n = 0;
            foreach (string raw in SplitLines(added))
            {
                string l = Unnumber(raw).Trim();
                if (l.StartsWith("■")) { light = l.Contains("軽作業"); continue; }
                if (!l.StartsWith("・") || l == "・なし" || l.Length < 3) continue;
                if (lines.Any(x => Unnumber(x).Trim() == l)) continue;   // 同じ行はもうある
                Match m = RangeRe.Match(l);
                if (!light && m.Success)
                {
                    string start = "・" + m.Groups[1].Value + "〜";
                    int i = lines.FindIndex(x => Unnumber(x).Trim().StartsWith(start) && !afkLines.Contains(Unnumber(x).Trim()));
                    if (i >= 0)
                    {
                        Match old = RangeRe.Match(Unnumber(lines[i]).Trim());
                        if (old.Success && Clock(m.Groups[2].Value) > Clock(old.Groups[2].Value))
                        {
                            lines[i] = "・" + old.Groups[1].Value + "〜" + m.Groups[2].Value + old.Groups[3].Value;
                            n++;
                        }
                        continue;
                    }
                }
                InsertItem(lines, light, l);
                n++;
            }
            return n;
        }

        // 離席メモの行: 20分を超えるなら工程（始め〜終わり）、それ以下は軽作業（終わりの時刻）
        static bool AfkIsLight(AfkEntry e) { return (e.End - e.Start).TotalMinutes <= 20; }

        static string AfkItem(AfkEntry e)
        {
            return AfkIsLight(e) ? "・" + e.End.ToString("HH:mm") + " " + e.Memo
                                 : "・" + e.Start.ToString("HH:mm") + "〜" + e.End.ToString("HH:mm") + " " + e.Memo;
        }

        // 離席メモを日報へ反映する（前に足した行は入れ替え、「除く」やメモなしなら消す）。日報がまだなければ何もしない
        public static bool ApplyAfk(AfkEntry e)
        {
            string date = e.Start.ToString("yyyy-MM-dd");
            string path = Core.ReportPath(date);
            if (!File.Exists(path)) return false;
            lock (reportLock)
            {
                List<string> lines = SplitLines(Core.ReadText(path));
                // 前に足した行を探す。記録がなければ（前の版の日報・AI が書いた行）、前のメモの行か同じ時間帯の工程の行
                string prev = e.Line;
                AfkEntry old = Afk.Load(date).FirstOrDefault(x => x.Start == e.Start);
                if (prev.Length == 0 && old != null && old.Memo.Length > 0) prev = AfkItem(old);
                int i = prev.Length > 0 ? lines.FindIndex(l => Unnumber(l).Trim() == prev) : -1;
                if (i < 0 && e.Line.Length == 0 && !AfkIsLight(e))
                {
                    string range = "・" + e.Start.ToString("HH:mm") + "〜" + e.End.ToString("HH:mm") + " ";
                    i = lines.FindIndex(l => Unnumber(l).Trim().StartsWith(range));
                }
                if (i >= 0) lines.RemoveAt(i);
                e.Line = "";
                if (!e.Skip && e.Memo.Length > 0)
                {
                    string item = AfkItem(e);
                    InsertItem(lines, AfkIsLight(e), item);
                    e.Line = item;
                }
                Tidy(lines, e.Start.ToString("yyyy-MM-dd"), Settings.Get(), e.Line);
                Core.WriteText(path, Number(string.Join("\n", lines).TrimEnd() + "\n"));
            }
            Afk.Put(e);
            return true;
        }

        static void ApplyAllAfk(string date)
        {
            foreach (AfkEntry e in Afk.Load(date))
            {
                e.Line = "";   // 作り直した日報には前の行はない
                if (!e.Skip && e.Memo.Length > 0) ApplyAfk(e); else Afk.Put(e);
            }
        }

        // 前に日報へ入れた時刻（since）よりあとの作業を、AI に新しい行だけ書いてもらって足す
        public static Result Append(string date, Settings s, DateTime since)
        {
            var r = new Result { Date = date };
            string path = Core.ReportPath(date);
            List<Row> all = ReadActive(date, s);
            List<Row> fresh = all == null ? new List<Row>() : all.Where(x => x.Time > since).ToList();
            if (fresh.Count == 0) { r.NoData = true; r.Message = "新しい作業はありません"; return r; }

            string current = Number(Core.ReadText(path));
            string prompt =
                "以下は今日の日報（ここまでの分）と、そのあとの PC の作業ログです。\n" +
                "日報にまだ書かれていない、" + since.ToString("HH:mm") + " より後の作業だけを、次の形で出力してください（前置き・説明は不要）:\n" +
                "■工程\n・HH:MM〜HH:MM 作業内容\n■軽作業\n・HH:MM 作業内容\n\n" +
                "条件:\n" +
                "- 日報にすでにある作業は書かない\n" +
                "- 日報の工程の作業がそのまま続いているときは、新しい行にせず、その行と同じ始めの時刻で終わりの時刻だけを延ばした行を書く" +
                "（例: 日報に「・13:00〜15:30 A工区の数量計算」があり 17:10 まで続いたら「・13:00〜17:10 A工区の数量計算」）。" +
                since.ToString("HH:mm") + " は区切りではないので、ここで作業を分けない\n" +
                "- 合計20分程度以下の作業は工程にせず軽作業に入れ、終えた時刻だけを書く\n" +
                "- 工程は時間が途切れないようにつなげる（前の工程の終わり＝次の工程の始め）。軽作業しかしていない時間は、工程に「・HH:MM〜HH:MM 軽作業」と書く\n" +
                "- 時刻は作業ログに出てくる時刻を使う。作業内容は20文字程度で、日報の書き方に合わせる\n" +
                "- 推定できないもの、私用と思われるものは省く\n" +
                "- 「時刻のない本人のメモ」のうち日報にまだないものがあれば、それも入れる\n" +
                BreakRule(s, date) + PlanRule(date, s) +
                "- 新しい作業がなければ「・なし」とだけ書く\n\n" +
                NotesPrompt(date) + Tasks.Prompt() + Afk.Prompt(date) +
                "今の日報:\n" + current.Replace("\r\n", "\n") + "\n\n" +
                since.ToString("HH:mm") + " より後の作業ログ:\n" + LogText(fresh, s, since);

            string added;
            try { added = ClaudeCli.Ask(prompt, s, date, "（追記）"); }
            catch (NeedLoginException ex) { r.NeedLogin = true; r.Message = ex.Message; return r; }   // 日報はそのまま。次の追記で試す
            catch (Exception ex) { r.Message = ex.Message; Core.Error("追記 " + date, ex.Message); return r; }

            int n;
            lock (reportLock)
            {
                // AI を待つあいだに離席メモなどが足されているかもしれないので、読み直してから足す
                List<string> lines = SplitLines(Core.ReadText(path));
                var fixedLines = new HashSet<string>(Afk.Load(date).Select(e => e.Line).Where(x => x.Length > 0));
                fixedLines.Add(BreakLine(s));   // 休憩の行は延ばさない
                n = MergeItems(lines, added, fixedLines);
                EnsureBreak(lines, date, s);
                Tidy(lines, date, s);
                if (lines.Count > 0 && lines[0].StartsWith("日報 ")) lines[0] = Header(date, all);
                Core.WriteText(path, Number(string.Join("\n", lines).TrimEnd() + "\n"));
                SaveMd(date, Core.ReadText(path));
            }
            SaveState(date, all[all.Count - 1].Time, true);
            r.Ok = true;
            r.Message = n > 0 ? "追記 " + n + " 件" : "追記なし";
            r.Appended = true;
            return r;
        }

        // 自動作成: 未作成の日・AI に失敗した日を作り、作ったあとに作業が続いた日（今日も）は追記する
        // final: 退勤したとき。今日の分は、新しい作業が少しでもあれば追記して仕上げる
        public static List<Result> Auto(Settings s, bool final = false)
        {
            var results = new List<Result>();
            DateTime now = DateTime.Now, today = now.Date;
            Dictionary<string, StateEntry> state = LoadState();
            for (int i = s.CatchUpDays; i >= 0; i--)
            {
                string d = today.AddDays(-i).ToString("yyyy-MM-dd");
                try
                {
                    // 追記しない設定のときは、今日の分は ReportTime を過ぎてから
                    bool append = s.AppendMinutes > 0;
                    if (i == 0 && !append && now.TimeOfDay < s.ReportTime) continue;
                    if (!File.Exists(Core.CsvPath(d))) continue;
                    StateEntry st;
                    if (!state.TryGetValue(d, out st))
                    {
                        if (File.Exists(Core.ReportPath(d))) continue;   // 記録を始める前に作った日報はそのまま
                        List<Row> a = ReadActive(d, s);
                        if (a != null && Minutes(a.Count) >= s.MinDayMinutes) results.Add(Make(d, s));
                    }
                    else if (!st.Ok || !File.Exists(Core.ReportPath(d))) results.Add(Make(d, s));
                    else if (i > 0 || append)
                    {
                        // 作ったあとに作業が続いた日は、作り直さずに追記する（手で直した行はそのまま）
                        List<Row> a = ReadActive(d, s);
                        int need = final && i == 0 ? 1 : s.RedoMinutes;
                        if (a != null && Minutes(a.Count(x => x.Time > st.Last)) >= need) results.Add(Append(d, s, st.Last));
                    }
                }
                catch (Exception ex)
                {
                    Core.Error("日報 " + d, ex.Message);
                    results.Add(new Result { Date = d, Ok = false, Message = ex.Message });
                }
            }
            Cleanup(s);
            return results;
        }

        static void Backup(string path)
        {
            if (!File.Exists(path)) return;
            string dir = Path.Combine(Core.Dir, "old");
            Directory.CreateDirectory(dir);
            File.Copy(path, Path.Combine(dir, Path.GetFileNameWithoutExtension(path) + "_" + DateTime.Now.ToString("yyyyMMdd-HHmmss") + ".txt"), true);
        }

        static void Cleanup(Settings s)
        {
            try
            {
                DateTime limit = DateTime.Now.AddDays(-s.KeepDays);
                foreach (string f in Directory.GetFiles(Core.LogDir, "*.csv"))
                {
                    string n = Path.GetFileNameWithoutExtension(f);
                    if ((IsDate(n) || (n.StartsWith("窓_") && IsDate(n.Substring(2)))) && File.GetLastWriteTime(f) < limit) File.Delete(f);
                }
                string old = Path.Combine(Core.Dir, "old");
                if (Directory.Exists(old))
                    foreach (string f in Directory.GetFiles(old, "日報_*.txt"))
                        if (File.GetLastWriteTime(f) < limit) File.Delete(f);
            }
            catch (Exception ex) { Core.Error("整理", ex.Message); }
        }

        // report_state.txt: 日付<TAB>日報に含めた最後の記録時刻<TAB>1=成功 0=失敗
        static Dictionary<string, StateEntry> LoadState()
        {
            var d = new Dictionary<string, StateEntry>();
            if (!File.Exists(StatePath)) return d;
            foreach (string raw in Core.ReadText(StatePath).Split('\n'))
            {
                string[] p = raw.TrimEnd('\r').Split('\t');
                DateTime t;
                if (p.Length == 3 && IsDate(p[0]) && DateTime.TryParseExact(p[1], "yyyy-MM-dd HH:mm:ss", Inv, DateTimeStyles.None, out t))
                    d[p[0]] = new StateEntry { Last = t, Ok = p[2] == "1" };
            }
            return d;
        }

        static void SaveState(string date, DateTime last, bool ok)
        {
            Dictionary<string, StateEntry> d = LoadState();
            d[date] = new StateEntry { Last = last, Ok = ok };
            Core.WriteText(StatePath, string.Concat(d.OrderBy(x => x.Key)
                .Select(x => x.Key + "\t" + x.Value.Last.ToString("yyyy-MM-dd HH:mm:ss") + "\t" + (x.Value.Ok ? "1" : "0") + "\r\n")));
        }
    }

    // 出勤・退勤（勤務.txt: on か off<TAB>その時刻）。退勤中は記録も AI も止めて待つ
    static class Duty
    {
        static string PathOf { get { return Path.Combine(Core.Dir, "勤務.txt"); } }

        // 起動したとき: 退勤のままか、前の日に出勤したまま（退勤を押さずに PC を切った）なら退勤中から始める
        public static bool LoadOff(out DateTime at)
        {
            at = DateTime.MinValue;
            try
            {
                if (!File.Exists(PathOf)) return false;
                string[] c = Core.ReadText(PathOf).Trim().Split('\t');
                if (c.Length < 2 || !DateTime.TryParseExact(c[1], "yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture, DateTimeStyles.None, out at)) return false;
                return c[0] == "off" || at.Date < DateTime.Today;
            }
            catch (Exception ex) { Core.Error("勤務", ex.Message); return false; }
        }

        public static void Save(bool off, DateTime at)
        {
            try
            {
                string line = (off ? "off" : "on") + "\t" + at.ToString("yyyy-MM-dd HH:mm:ss") + "\r\n";
                Core.WriteText(PathOf, line);
                File.AppendAllText(HistoryPath, line, new UTF8Encoding(false));   // 退勤していた時間帯を日報画面の帯に出すため
            }
            catch (Exception ex) { Core.Error("勤務", ex.Message); }
        }

        static string HistoryPath { get { return Path.Combine(Core.Dir, "勤務履歴.log"); } }

        // 出勤・退勤の履歴（時刻の順）。true が退勤（0.12.0 から残す）
        public static List<KeyValuePair<DateTime, bool>> History()
        {
            var list = new List<KeyValuePair<DateTime, bool>>();
            try
            {
                if (!File.Exists(HistoryPath)) return list;
                foreach (string raw in Core.ReadText(HistoryPath).Split('\n'))
                {
                    string[] c = raw.Trim().Split('\t');
                    DateTime t;
                    if (c.Length >= 2 && DateTime.TryParseExact(c[1], "yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture, DateTimeStyles.None, out t))
                        list.Add(new KeyValuePair<DateTime, bool>(t, c[0] == "off"));
                }
            }
            catch (Exception ex) { Core.Error("勤務", ex.Message); }
            return list.OrderBy(x => x.Key).ToList();
        }

        // 退勤中に出勤とみなすか: 退勤した日より後の日で、出勤の時刻を過ぎて操作があった
        public static bool AutoIn(DateTime now, DateTime offAt, uint idleSec, Settings s)
        {
            return idleSec < 10 && now.TimeOfDay >= s.WorkStart && offAt.Date < now.Date;
        }

        // 出勤中に退勤とみなすか: 仕上げの時刻（ReportTime）を過ぎてから OffIdleMinutes 操作がない
        public static bool AutoOut(DateTime now, DateTime lastInput, Settings s)
        {
            if (s.OffIdleMinutes <= 0 || now.TimeOfDay < s.ReportTime) return false;
            DateTime from = lastInput > now.Date + s.ReportTime ? lastInput : now.Date + s.ReportTime;
            return now - from >= TimeSpan.FromMinutes(s.OffIdleMinutes);
        }
    }

    // トレイ常駐の本体
    // 自動アップデート: GitHub の最新リリースを見て、新しい版があれば zip を落とし、exe を入れ替えて起動し直す
    // 動いている exe は消せないが名前は変えられるので、WorkLog.exe → WorkLog.exe.old にして新しい exe を置く（.old は次の起動で消す）
    static class Updater
    {
        public const string Repo = "sarumogu2013-glitch/AutoNippou";
        public const string Page = "https://github.com/" + Repo + "/releases/latest";
        public const string Releases = "https://github.com/" + Repo + "/releases";
        public const string Help = "https://github.com/" + Repo + "/blob/main/HELP.md";
        public class Release { public string Tag, Version, Notes, Url, ZipName, ZipUrl, Digest; }

        static string Dir { get { return Path.Combine(Core.Dir, "update"); } }

        // 今の版。テスト中（WORKLOG_TEST）だけ WORKLOG_FAKE_VERSION で古い版のふりができる（更新の動作確認用）
        public static string Current
        {
            get
            {
                string f = Core.Test ? Environment.GetEnvironmentVariable("WORKLOG_FAKE_VERSION") : null;
                return string.IsNullOrEmpty(f) ? Core.Version : f;
            }
        }

        // update.log に1行ずつ残す
        public static void Log(string msg)
        {
            try
            {
                File.AppendAllText(Path.Combine(Core.Dir, "update.log"),
                    DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " " + msg + "\r\n", new UTF8Encoding(true));
            }
            catch { }
        }

        static Version Parse(string v)
        {
            Version r;
            return Version.TryParse((v ?? "").Trim().TrimStart('v', 'V'), out r) ? r : null;
        }

        public static bool IsNewer(string latest, string current)
        {
            Version a = Parse(latest), b = Parse(current);
            return a != null && b != null && a > b;
        }

        static System.Net.HttpWebRequest Request(string url)
        {
            System.Net.ServicePointManager.SecurityProtocol |= System.Net.SecurityProtocolType.Tls12;
            var req = (System.Net.HttpWebRequest)System.Net.WebRequest.Create(url);
            req.UserAgent = "WorkLog/" + Core.Version;
            req.Timeout = 30000;
            req.ReadWriteTimeout = 120000;
            return req;
        }

        static string Str(IDictionary<string, object> d, string k)
        {
            object v;
            return d != null && d.TryGetValue(k, out v) && v != null ? v.ToString() : "";
        }

        // 最新のリリースを調べる（インストール用 zip がないリリースは例外）
        // 最新のリリース。GitHub の API が回数の上限（403・429）のときは、リリースのページの転送先から版を読む
        public static Release Latest()
        {
            try { return LatestByApi(); }
            catch (System.Net.WebException ex)
            {
                var res = ex.Response as System.Net.HttpWebResponse;
                if (res == null || ((int)res.StatusCode != 403 && (int)res.StatusCode != 429)) throw;
                Log("GitHub の API が回数の上限（" + (int)res.StatusCode + "）のため、リリースのページから確かめます");
                return LatestByRedirect();
            }
        }

        // github.com/…/releases/latest は最新のタグのページへ転送されるので、その行き先から版を読む（zip の SHA-256 はわからない）
        static Release LatestByRedirect()
        {
            System.Net.HttpWebRequest req = Request(Page);
            req.AllowAutoRedirect = false;
            string location;
            using (var res = (System.Net.HttpWebResponse)req.GetResponse()) location = res.Headers["Location"] ?? "";
            Match m = Regex.Match(location, @"/releases/tag/(v?[0-9][0-9.]*)$");
            if (!m.Success) throw new Exception("最新のリリースがわかりません（" + location + "）");
            var rel = new Release { Tag = m.Groups[1].Value, Notes = "", Url = "https://github.com/" + Repo + "/releases/tag/" + m.Groups[1].Value, Digest = "" };
            rel.Version = rel.Tag.TrimStart('v', 'V');
            rel.ZipName = "WorkLog-" + rel.Version + ".zip";
            rel.ZipUrl = "https://github.com/" + Repo + "/releases/download/" + rel.Tag + "/" + rel.ZipName;
            return rel;
        }

        static Release LatestByApi()
        {
            System.Net.HttpWebRequest req = Request("https://api.github.com/repos/" + Repo + "/releases/latest");
            req.Accept = "application/vnd.github+json";
            // テスト（CI）のときだけ、回数の上限に当たらないようにトークンを付ける
            string token = Core.Test ? Environment.GetEnvironmentVariable("WORKLOG_GITHUB_TOKEN") : null;
            if (!string.IsNullOrEmpty(token)) req.Headers["Authorization"] = "Bearer " + token;
            string json;
            using (var res = req.GetResponse())
            using (var r = new StreamReader(res.GetResponseStream(), Encoding.UTF8))
                json = r.ReadToEnd();
            var d = new JavaScriptSerializer { MaxJsonLength = int.MaxValue }.Deserialize<Dictionary<string, object>>(json);
            var rel = new Release { Tag = Str(d, "tag_name"), Notes = Str(d, "body"), Url = Str(d, "html_url") };
            rel.Version = rel.Tag.TrimStart('v', 'V');
            if (Parse(rel.Version) == null) throw new Exception("最新のリリースの版が読めません: " + rel.Tag);
            object assets;
            if (d.TryGetValue("assets", out assets) && assets is IEnumerable)
                foreach (object o in (IEnumerable)assets)
                {
                    var a = o as IDictionary<string, object>;
                    string name = Str(a, "name");
                    if (!Regex.IsMatch(name, @"^WorkLog-[0-9.]+\.zip$", RegexOptions.IgnoreCase)) continue;
                    rel.ZipName = name;
                    rel.ZipUrl = Str(a, "browser_download_url");
                    rel.Digest = Str(a, "digest");
                    break;
                }
            if (rel.ZipUrl == null) throw new Exception(rel.Tag + " にインストール用 zip がありません");
            return rel;
        }

        // zip を落として確かめ（SHA-256・中の exe の版）、新しい exe の場所を返す
        public static string Download(Release rel)
        {
            Directory.CreateDirectory(Dir);
            foreach (string f in Directory.GetFiles(Dir)) { try { File.Delete(f); } catch { } }
            string zip = Path.Combine(Dir, rel.ZipName);
            using (var res = Request(rel.ZipUrl).GetResponse())
            using (Stream s = res.GetResponseStream())
            using (FileStream f = File.Create(zip))
                s.CopyTo(f);
            if (rel.Digest.StartsWith("sha256:", StringComparison.OrdinalIgnoreCase))
            {
                string hash;
                using (var sha = System.Security.Cryptography.SHA256.Create())
                using (FileStream f = File.OpenRead(zip))
                    hash = BitConverter.ToString(sha.ComputeHash(f)).Replace("-", "");
                if (!hash.Equals(rel.Digest.Substring(7), StringComparison.OrdinalIgnoreCase))
                    throw new Exception("ダウンロードした zip が壊れています（SHA-256 が合いません）");
            }
            string exe = Path.Combine(Dir, "WorkLog.new.exe");
            using (System.IO.Compression.ZipArchive a = System.IO.Compression.ZipFile.OpenRead(zip))
            {
                System.IO.Compression.ZipArchiveEntry e = a.Entries.FirstOrDefault(x => x.Name.Equals("WorkLog.exe", StringComparison.OrdinalIgnoreCase));
                if (e == null) throw new Exception("zip の中に WorkLog.exe がありません");
                System.IO.Compression.ZipFileExtensions.ExtractToFile(e, exe, true);
            }
            string v = FileVersionInfo.GetVersionInfo(exe).FileVersion;
            if (Parse(v) == null || Parse(v) != Parse(rel.Version))
                throw new Exception("zip の中の exe の版（" + v + "）がリリース（" + rel.Version + "）と違います");
            return exe;
        }

        // exe を入れ替えて新しい exe を起動する（呼んだ側はこのあとすぐ終わる）。うまくいかなければ元に戻して例外
        public static void Apply(string newExe)
        {
            string exe = Application.ExecutablePath, old = exe + ".old";
            if (File.Exists(old)) File.Delete(old);
            File.Move(exe, old);
            try
            {
                File.Copy(newExe, exe);
                Process.Start(new ProcessStartInfo(exe, "--updated " + Process.GetCurrentProcess().Id)
                    { UseShellExecute = false, WorkingDirectory = Path.GetDirectoryName(exe) });
            }
            catch
            {
                try { if (File.Exists(exe)) File.Delete(exe); File.Move(old, exe); } catch { }
                throw;
            }
        }

        // 更新で起動されたとき: 前の版が終わるのを待ってから、前の exe と落とした zip を消す
        public static void AfterUpdate(int pid)
        {
            try { using (Process p = Process.GetProcessById(pid)) p.WaitForExit(30000); }
            catch { }   // もう終わっている
            string old = Application.ExecutablePath + ".old";
            for (int i = 0; i < 10 && File.Exists(old); i++)
            {
                try { File.Delete(old); }
                catch { Thread.Sleep(500); }
            }
            try { if (Directory.Exists(Dir)) Directory.Delete(Dir, true); } catch { }
            Log(Core.Version + " を起動しました（更新）");
        }

        // ふだんの起動: 前の更新の残りがあれば消す
        public static void Cleanup()
        {
            try
            {
                string old = Application.ExecutablePath + ".old";
                if (File.Exists(old)) File.Delete(old);
            }
            catch { }
        }
    }

    class TrayApp : ApplicationContext
    {
        const string RunKey = @"Software\Microsoft\Windows\CurrentVersion\Run";
        const string RunName = "WorkLog";

        readonly NotifyIcon tray = new NotifyIcon();
        readonly System.Windows.Forms.Timer timer = new System.Windows.Forms.Timer();
        readonly Logger logger = new Logger();
        readonly DateTime started = DateTime.Now;
        readonly Icon icon = MakeIcon(Color.FromArgb(30, 111, 217)), offIcon = MakeIcon(Color.FromArgb(150, 150, 150));
        readonly SynchronizationContext ui;
        readonly System.Windows.Forms.Timer afkTimer = new System.Windows.Forms.Timer();
        DateTime lastAuto = DateTime.MinValue;
        DateTime lastInput;
        MainForm form;
        AiForm aiForm;
        SettingsForm settingsForm;
        AfkAskForm askForm;
        bool busy;
        Action balloonClick;
        bool off;          // 退勤中（待機）
        DateTime offAt;    // 退勤した時刻
        ToolStripMenuItem dutyItem, startupItem;
        DateTime lastUpdateCheck = DateTime.MinValue;
        bool checking;                     // 更新を確かめている・落としている最中
        Updater.Release pendingRelease;    // 落とし終えて入れ替え待ちの版
        string pendingExe;
        bool applyNow;                     // 手で「更新する」と言われた（日報画面が開いていても入れ替える）
        string failedVersion, notifiedVersion;

        public TrayApp(bool show, bool updated, bool firstRun)
        {
            // トレイには、画面を開かずに済ませたいものだけを置く（ほかは日報画面のメニューバー）
            var menu = new ContextMenuStrip();
            ui = SynchronizationContext.Current ?? new WindowsFormsSynchronizationContext();
            menu.Items.Add("日報を開く", null, (o, e) => ShowForm(null));
            menu.Items.Add("今日の日報を今すぐ作る", null, (o, e) => Generate(DateTime.Today.ToString("yyyy-MM-dd")));
            menu.Items.Add(new ToolStripSeparator());
            dutyItem = new ToolStripMenuItem("退勤", null, (o, e) => ToggleDuty());
            menu.Items.Add(dutyItem);
            menu.Items.Add(new ToolStripSeparator());
            startupItem = new ToolStripMenuItem("Windows 起動時に自動で起動する", null, (o, e) => ToggleStartup()) { Enabled = !Core.Test };
            menu.Items.Add(startupItem);
            menu.Items.Add("更新を確認", null, (o, e) => CheckUpdate(true));
            menu.Items.Add(new ToolStripSeparator());
            menu.Items.Add("終了", null, (o, e) => Quit());
            menu.Opening += (o, e) => { try { startupItem.Checked = StartupPath() != null; } catch { } };

            tray.Icon = icon;
            tray.Text = "作業ログ";
            tray.ContextMenuStrip = menu;
            tray.MouseClick += (o, e) => { if (e.Button == MouseButtons.Left) ShowForm(null); };
            tray.BalloonTipClicked += (o, e) => { if (balloonClick != null) balloonClick(); };
            tray.Visible = true;

            off = Duty.LoadOff(out offAt);
            UpdateDuty();
            timer.Interval = Core.IntervalSec * 1000;
            timer.Tick += (o, e) => OnTick();
            timer.Start();
            OnTick();
            lastInput = DateTime.Now.AddSeconds(-Logger.IdleSec());
            afkTimer.Interval = 5000;
            afkTimer.Tick += (o, e) => CheckAfk();
            afkTimer.Start();
            if (show) ShowForm(null);
            if (updated) Balloon(Core.Version + " に更新", "更新履歴: ヘルプ → バージョン情報", ToolTipIcon.Info, () => ShowForm(null));
            if (firstRun && !Core.Test) { AskStartup(); FirstAiSetup(); }
            else CheckAi();
        }

        // 初めての起動: Claude Code が見つからなければ「詳しく探す」も行い、見つけたものを使う。そのあと AI の接続を確かめる
        void FirstAiSetup()
        {
            var t = new Thread(() =>
            {
                string found = null;
                try
                {
                    Settings s = Settings.Get();
                    if (ClaudeCli.FindExe(s) == null)
                    {
                        List<KeyValuePair<string, string>> list = ClaudeCli.DeepSearch(null);
                        if (list.Count > 0) { found = list[0].Key; Settings.SetClaudePath(found); }
                    }
                }
                catch (Exception ex) { Core.Error("初回の AI の確認", ex.Message); }
                ui.Post(_ =>
                {
                    if (found != null) Balloon("Claude Code を設定", found, ToolTipIcon.Info, ShowAiForm);
                    CheckAi();
                }, null);
            });
            t.IsBackground = true;
            t.Start();
        }

        // 初めての起動: スタートアップに登録するか聞く（あとからトレイのメニューか設定のタブ「起動」で変えられる）
        void AskStartup()
        {
            try
            {
                if (StartupPath() != null) return;
                if (MessageBox.Show("Windows 起動時に自動で起動しますか？", "作業ログ", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return;
                SetStartup(true);
            }
            catch (Exception ex) { Core.Error("スタートアップ", ex.Message); }
        }

        void ToggleStartup()
        {
            if (Core.Test) return;
            try
            {
                bool on = StartupPath() == null;
                SetStartup(on);
                Balloon(on ? "自動で起動するように登録しました" : "自動で起動しないようにしました",
                    on ? "登録した場所: " + Application.ExecutablePath : "Windows を起動しても作業ログは起動しません（記録も取りません）", ToolTipIcon.Info, null);
            }
            catch (Exception ex)
            {
                Core.Error("スタートアップ", ex.Message);
                Balloon("自動起動を変えられませんでした", ex.Message, ToolTipIcon.Warning, null);
            }
        }

        // 自動アップデート: 起動の5分後と、そのあと12時間ごとに確かめる。落とした新しい版は、日報を作っていなくて
        // 画面を開いていないとき（記録の1分ごとに確かめる）に入れ替えて起動し直す
        void UpdateTick()
        {
            if (pendingExe != null) { TryApply(); return; }
            DateTime now = DateTime.Now;
            if (Core.Test || checking || now < started.AddMinutes(5) || now - lastUpdateCheck < TimeSpan.FromHours(12)) return;
            lastUpdateCheck = now;
            CheckUpdate(false);
        }

        // manual: メニューの「更新を確認」（結果を必ず知らせる）
        public void CheckUpdate(bool manual)
        {
            if (checking) { if (manual) Balloon("更新", "確認中", ToolTipIcon.Info, null); return; }
            if (pendingExe != null) { if (manual) { applyNow = true; TryApply(); } return; }
            checking = true;
            bool auto = Settings.Get().AutoUpdate;
            var t = new Thread(() =>
            {
                Updater.Release rel = null;
                string exe = null;
                Exception err = null;
                try
                {
                    rel = Updater.Latest();
                    if (Updater.IsNewer(rel.Version, Updater.Current) && auto && !manual && rel.Version != failedVersion)
                        exe = Updater.Download(rel);
                }
                catch (Exception ex) { err = ex; }
                ui.Post(_ => CheckDone(rel, exe, err, manual), null);
            });
            t.IsBackground = true;
            t.Start();
        }

        void CheckDone(Updater.Release rel, string exe, Exception err, bool manual)
        {
            checking = false;
            if (err != null)
            {
                Updater.Log("確かめられませんでした: " + err.Message);
                if (manual) MessageBox.Show("新しい版があるか確かめられませんでした。\n" + err.Message, "更新を確認", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            bool newer = Updater.IsNewer(rel.Version, Updater.Current);
            Updater.Log("確認: 最新 " + rel.Version + " / 今 " + Updater.Current + (newer ? "（新しい版あり）" : ""));
            if (!newer)
            {
                if (manual) MessageBox.Show("最新の版です（" + Updater.Current + "）。", "更新を確認", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (manual) { ConfirmUpdate(rel); return; }
            if (exe != null)
            {
                Updater.Log(rel.Version + " を落としました。日報画面を閉じているときに入れ替えます");
                pendingRelease = rel;
                pendingExe = exe;
                TryApply();
                return;
            }
            if (rel.Version == failedVersion || rel.Version == notifiedVersion) return;
            notifiedVersion = rel.Version;
            Balloon("新しい版 " + rel.Version, "クリックで更新", ToolTipIcon.Info, () => ConfirmUpdate(rel));
        }

        void ConfirmUpdate(Updater.Release rel)
        {
            string notes = (rel.Notes ?? "").Trim();
            if (notes.Length > 600) notes = notes.Substring(0, 600) + "…";
            if (MessageBox.Show("新しい版 " + rel.Version + "（今は " + Updater.Current + "）\n\n" + (notes.Length > 0 ? notes + "\n\n" : "") + "更新しますか？",
                    "更新を確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return;
            checking = true;
            var t = new Thread(() =>
            {
                string exe = null;
                Exception err = null;
                try { exe = Updater.Download(rel); }
                catch (Exception ex) { err = ex; }
                ui.Post(_ =>
                {
                    checking = false;
                    if (err != null)
                    {
                        Updater.Log(rel.Version + " を落とせませんでした: " + err.Message);
                        MessageBox.Show("新しい版を落とせませんでした。\n" + err.Message, "更新", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                    pendingRelease = rel;
                    pendingExe = exe;
                    applyNow = true;
                    TryApply();
                }, null);
            });
            t.IsBackground = true;
            t.Start();
        }

        bool Shown(Form f) { return f != null && !f.IsDisposed && f.Visible; }

        void TryApply()
        {
            if (pendingExe == null) return;
            if (busy)
            {
                if (applyNow) Balloon("更新", "日報の作成後に更新", ToolTipIcon.Info, null);
                return;   // 次の記録のときにまた試す
            }
            if (!applyNow && (Shown(form) || Shown(askForm) || Shown(settingsForm) || Shown(aiForm))) return;
            Updater.Release rel = pendingRelease;
            string exe = pendingExe;
            pendingRelease = null;
            pendingExe = null;
            applyNow = false;
            try
            {
                if (form != null && !form.IsDisposed) form.SaveIfDirty();
                Updater.Apply(exe);
                Updater.Log(Core.Version + " → " + rel.Version + " に入れ替えました（起動し直します）");
            }
            catch (Exception ex)
            {
                failedVersion = rel.Version;
                Updater.Log(rel.Version + " に入れ替えられませんでした: " + ex.Message);
                Balloon("更新できません", ex.Message + "\nクリックでリリースのページ",
                    ToolTipIcon.Warning, () => { try { Process.Start(rel.Url); } catch { } });
                return;
            }
            Shutdown();
        }

        // 一晩つけっぱなしなどは離席として扱わない
        static readonly TimeSpan AfkMax = TimeSpan.FromHours(8);

        // 入力のあいだが AfkMinutes 以上あいたら、戻ってきたところで離席として記録し、何をしていたか聞く
        void CheckAfk()
        {
            try
            {
                uint idle = Logger.IdleSec();
                DateTime now = DateTime.Now, input = now.AddSeconds(-idle);
                Settings s = Settings.Get();
                if (off)
                {
                    if (Duty.AutoIn(now, offAt, idle, s)) ClockIn(false);
                    return;
                }
                if (Duty.AutoOut(now, input > lastInput ? input : lastInput, s)) { ClockOut(input > lastInput ? input : lastInput, false); return; }
                TimeSpan gap = input - lastInput;
                if (s.AfkMinutes > 0 && gap >= TimeSpan.FromMinutes(s.AfkMinutes) && gap <= AfkMax && input.Date == lastInput.Date)
                {
                    var e = new AfkEntry { Start = lastInput, End = input };
                    // 半分以上が休憩時間なら休憩として記録し、聞かない
                    bool rest = Reporter.BreakOn(input.ToString("yyyy-MM-dd"), s) && Overlap(e, s).TotalMinutes * 2 >= gap.TotalMinutes;
                    if (rest) { e.Skip = true; e.Memo = "休憩（自動）"; }
                    AddAfk(e, s.AfkAsk && !rest);
                }
                if (input > lastInput) lastInput = input;
            }
            catch (Exception ex) { Core.Error("離席", ex.Message); }
        }

        static TimeSpan Overlap(AfkEntry e, Settings s)
        {
            DateTime bs = e.Start.Date + s.BreakStart, be = e.Start.Date + s.BreakEnd;
            DateTime a = e.Start > bs ? e.Start : bs, b = e.End < be ? e.End : be;
            return b > a ? b - a : TimeSpan.Zero;
        }

        void AddAfk(AfkEntry e, bool ask)
        {
            Afk.Put(e);
            OnAfkSaved(e.Start.ToString("yyyy-MM-dd"));
            if (!ask) return;
            if (askForm == null || askForm.IsDisposed) askForm = new AfkAskForm(icon, AfkChanged);
            askForm.SetEntry(e);
            if (!askForm.Visible) askForm.Show();
        }

        void OnAfkSaved(string date)
        {
            if (form != null && !form.IsDisposed) form.RefreshAfk(date);
        }

        // 離席メモが書かれた・直された: 日報があればすぐその時刻の行として足す（なければ日報を作るときに入る）
        public void AfkChanged(AfkEntry e)
        {
            string date = e.Start.ToString("yyyy-MM-dd");
            bool shown = form != null && !form.IsDisposed;
            if (shown) form.SaveQuiet();   // 手で直していた本文を先に保存
            bool applied = false;
            try { applied = Reporter.ApplyAfk(e); }
            catch (Exception ex) { Core.Error("離席メモ", ex.Message); Balloon("離席メモ", "日報に追加できません: " + ex.Message, ToolTipIcon.Warning, null); }
            if (!applied) Afk.Put(e);
            if (shown) form.ReportUpdated(date);
        }

        public bool Off { get { return off; } }

        public void ToggleDuty()
        {
            if (off) ClockIn(true); else ClockOut(DateTime.Now, true);
        }

        // 出勤: 記録を始め、AI の接続を確かめて、残っている日報（前の日の仕上げなど）を作る
        void ClockIn(bool manual)
        {
            off = false;
            DateTime now = DateTime.Now;
            Duty.Save(false, now);
            lastInput = now;   // 退勤中のあいだを離席にしない
            logger.Tick();
            UpdateDuty();
            Balloon("出勤（" + now.ToString("HH:mm") + "）", manual ? "記録開始" : "記録開始（操作を検知）", ToolTipIcon.Info, null);
            CheckAi();
            if (!busy && now >= started.AddMinutes(2)) { lastAuto = now; Run(() => Reporter.Auto(Settings.Get()), false, true); }
        }

        // 退勤: 今日の日報を仕上げて（残りの作業を追記）、記録を止めて待つ
        void ClockOut(DateTime at, bool manual)
        {
            off = true;
            offAt = at;
            Duty.Save(true, at);
            logger.CloseWindows(at);   // 退勤中は開いていても作業していない
            UpdateDuty();
            Balloon("退勤（" + at.ToString("HH:mm") + "）", manual ? "記録停止" : "記録停止（操作なし）", ToolTipIcon.Info, null);
            if (askForm != null && !askForm.IsDisposed && askForm.Visible) askForm.Hide();
            if (!busy) { lastAuto = DateTime.Now; Run(() => Reporter.Auto(Settings.Get(), true), false, true); }
        }

        void UpdateDuty()
        {
            dutyItem.Text = off ? "出勤する（記録と AI を始める）" : "退勤する（日報を仕上げて待機）";
            tray.Icon = off ? offIcon : icon;
            string st = "作業ログ " + RecordStatus();
            tray.Text = st.Length > 63 ? st.Substring(0, 63) : st;
            if (form != null && !form.IsDisposed) { form.SetDuty(off); form.SetRecordStatus(RecordStatus()); }
        }

        // 動作確認用（WORKLOG_TEST のときの --test-afk）: 70分前から5分前まで離席していたことにする
        public void TestAfk()
        {
            DateTime now = DateTime.Now;
            AddAfk(new AfkEntry { Start = now.AddMinutes(-70), End = now.AddMinutes(-5) }, true);
        }

        // Claude Code を使えないときに知らせる（確かめるのに数秒かかるので別スレッドで）
        void CheckAi()
        {
            var t = new Thread(() =>
            {
                string exe;
                ClaudeCli.Auth auth;
                string problem = ClaudeCli.Problem(Settings.Get(), out exe, out auth);
                if (problem != null)
                    ui.Post(_ => Balloon("AI（Claude Code）を使えません", problem, ToolTipIcon.Warning, ShowAiForm), null);
            });
            t.IsBackground = true;
            t.Start();
        }

        public void ShowAiForm()
        {
            if (aiForm == null || aiForm.IsDisposed) aiForm = new AiForm(icon, OnAiReady);
            aiForm.Show();
            if (aiForm.WindowState == FormWindowState.Minimized) aiForm.WindowState = FormWindowState.Normal;
            aiForm.Activate();
        }

        public void ShowSettingsForm()
        {
            if (settingsForm == null || settingsForm.IsDisposed) settingsForm = new SettingsForm(icon, OnSettingsSaved, ShowAiForm);
            settingsForm.Show();
            if (settingsForm.WindowState == FormWindowState.Minimized) settingsForm.WindowState = FormWindowState.Normal;
            settingsForm.Activate();
        }

        // 日報の時刻の案内を直し、AI の接続先が変わっていれば知らせる
        void OnSettingsSaved()
        {
            if (form != null && !form.IsDisposed) form.UpdateLink();
            if (form != null && !form.IsDisposed && form.Visible) form.RefreshList(null);
            CheckAi();
        }

        // 接続先を確かめて使える状態なら、ログイン待ちだった日報をすぐ作り直す
        void OnAiReady()
        {
            if (busy) return;
            lastAuto = DateTime.Now;
            Run(() => Reporter.Auto(Settings.Get()), false, true);
        }

        public string RecordStatus()
        {
            if (off) return "退勤中（" + offAt.ToString("M/d HH:mm") + "から待機。記録していません）";
            if (logger.LastError != null) return "記録エラー: " + logger.LastError;
            if (logger.LastOk == DateTime.MinValue) return "記録待ち";
            return "記録中（最後の記録 " + logger.LastOk.ToString("HH:mm") + "）";
        }

        void OnTick()
        {
            if (!off) logger.Tick();
            string st = "作業ログ " + RecordStatus();
            tray.Text = st.Length > 63 ? st.Substring(0, 63) : st;
            if (form != null && !form.IsDisposed && form.Visible) form.SetRecordStatus(RecordStatus());
            UpdateTick();

            // 起動直後はネットワークを待つ。日付が変わったとき・仕上げの時刻を過ぎたとき（知らせる）と、AppendMinutes ごと（静かに追記）
            Settings s = Settings.Get();
            DateTime now = DateTime.Now;
            if (off || busy || now < started.AddMinutes(2)) return;
            bool newDay = lastAuto.Date != now.Date;
            bool finish = now.TimeOfDay >= s.ReportTime && lastAuto < now.Date + s.ReportTime;
            bool interval = s.AppendMinutes > 0 && now - lastAuto >= TimeSpan.FromMinutes(s.AppendMinutes);
            if (!newDay && !finish && !interval) return;
            if (!newDay && !finish && Resting(s, now) != null) return;   // 休憩・離席のあいだは追記しない（戻ってから追記する）
            if (newDay) CheckAi();   // 1日1回
            lastAuto = now;
            Run(() => Reporter.Auto(s), false, newDay || finish);
        }

        public bool Busy { get { return busy; } }

        // 休憩中・離席中なら、その説明（AI の自動追記を止めている）
        string Resting(Settings s, DateTime now)
        {
            if (s.InBreak(now) && Reporter.BreakOn(now.ToString("yyyy-MM-dd"), s))
                return "休憩中（" + s.BreakEnd.ToString(@"hh\:mm") + " まで）";
            uint idle = Logger.IdleSec();
            if (s.AfkMinutes > 0 && idle >= s.AfkMinutes * 60)
                return "離席中（" + now.AddSeconds(-idle).ToString("HH:mm") + " から。戻ると再開）";
            return null;
        }

        // AI の今の状態（日報画面の帯の上に出す）: 休止中ならその理由、動くなら次にいつ何をするか
        public string AiState()
        {
            Settings s = Settings.Get();
            DateTime now = DateTime.Now;
            if (busy) return "AI 実行中（日報を作っています）";
            if (off) return "AI 休止中: 退勤中（" + offAt.ToString("M/d HH:mm") + " から。出勤すると再開）";
            string rest = Resting(s, now);
            if (rest != null) return "AI 休止中: " + rest;
            DateTime first = started.AddMinutes(2);
            if (now < first || lastAuto.Date != now.Date) return "次の AI: まもなく（まだ作っていない日報・追記を確かめる）";
            DateTime finish = now.Date + s.ReportTime;
            var next = new List<KeyValuePair<DateTime, string>>();
            if (lastAuto < finish) next.Add(new KeyValuePair<DateTime, string>(finish, "今日の日報を仕上げる"));
            if (s.AppendMinutes > 0) next.Add(new KeyValuePair<DateTime, string>(lastAuto.AddMinutes(s.AppendMinutes), "今日の日報に追記"));
            if (next.Count == 0) return "次の AI: 明日（日付が変わったら前の日を仕上げる）";
            KeyValuePair<DateTime, string> n = next.OrderBy(x => x.Key).First();
            DateTime at = n.Key;
            // 休憩中にあたるなら、休憩が終わってから
            if (Reporter.BreakOn(now.ToString("yyyy-MM-dd"), s) && s.InBreak(at) && at.Date == now.Date) at = now.Date + s.BreakEnd;
            if (at <= now) return "次の AI: まもなく（" + n.Value + "）";
            int min = (int)Math.Ceiling((at - now).TotalMinutes);
            return "次の AI: " + at.ToString("HH:mm") + " " + n.Value + "（あと " + (min >= 60 ? min / 60 + "時間" + (min % 60 > 0 ? min % 60 + "分" : "") : min + "分") + "）";
        }

        public void Generate(string date)
        {
            if (busy) { Balloon("日報", "作成中", ToolTipIcon.Info, null); return; }
            Run(() => new List<Result> { Reporter.Make(date, Settings.Get()) }, true, true);
        }

        // notify: 終わったら通知・画面を開く（途中の追記は静かに、開いている画面だけ更新する）
        void Run(Func<List<Result>> work, bool manual, bool notify)
        {
            busy = true;
            if (form != null && !form.IsDisposed) { form.SaveQuiet(); form.SetBusy(true); }
            var t = new Thread(() =>
            {
                List<Result> res = null;
                Exception err = null;
                try { res = Reporter.Locked(work); }
                catch (Exception ex) { err = ex; Core.Error("日報", ex.Message); }
                ui.Post(_ => Done(res, err, manual, notify), null);
            });
            t.IsBackground = true;
            t.Start();
        }

        void Done(List<Result> res, Exception err, bool manual, bool notify)
        {
            busy = false;
            if (form != null && !form.IsDisposed) form.SetBusy(false);
            if (aiForm != null && !aiForm.IsDisposed && aiForm.Visible) aiForm.RefreshHistory();
            if (err != null) { if (notify) Balloon("日報を作れませんでした", err.Message, ToolTipIcon.Error, null); return; }
            if (!notify)
            {
                // 途中の追記: 知らせず、開いている画面だけ新しくする
                if (form != null && !form.IsDisposed)
                    foreach (Result r in res.Where(x => !x.NoData)) form.ReportUpdated(r.Date);
                return;
            }
            string open = null;
            Result needLogin = null;
            foreach (Result r in res)
            {
                if (r.NoData) { if (manual) Balloon("日報", r.Message, ToolTipIcon.Info, null); continue; }
                open = r.Date;
                string d = r.Date;
                if (r.NeedLogin) needLogin = r;
                else if (!r.Ok) Balloon("日報 " + d + " を作れません", r.Message, ToolTipIcon.Warning, () => ShowForm(d));
                else Balloon("日報 " + d + (r.Appended ? " 追記" : " 作成"), r.Message ?? "", ToolTipIcon.Info, () => ShowForm(d));
            }
            // ログインが原因のときは、日報画面ではなく接続先の画面へ案内する
            if (needLogin != null) Balloon("AI（Claude Code）を使えません", needLogin.Message, ToolTipIcon.Warning, ShowAiForm);
            if (open == null) return;
            if (manual || (needLogin == null && Settings.Get().OpenWindowOnReport)) ShowForm(open);
            else if (form != null && !form.IsDisposed) form.RefreshList(open);
        }

        void Balloon(string title, string text, ToolTipIcon kind, Action click)
        {
            balloonClick = click;
            tray.ShowBalloonTip(10000, title, text.Length > 200 ? text.Substring(0, 200) : text, kind);
        }

        public void ShowForm(string date)
        {
            if (form == null || form.IsDisposed) form = new MainForm(this, icon);
            form.UpdateLink();   // 設定が外から書き換えられたときも反映する
            form.RefreshList(date);
            form.SetRecordStatus(RecordStatus());
            form.SetDuty(off);
            form.SetBusy(busy);
            form.Show();
            if (form.WindowState == FormWindowState.Minimized) form.WindowState = FormWindowState.Normal;
            form.Activate();
        }

        // Windows 起動時の自動起動（HKCU\...\Run）。登録されている場所も返す
        public static string StartupPath()
        {
            using (RegistryKey k = Registry.CurrentUser.OpenSubKey(RunKey))
                return k == null ? null : k.GetValue(RunName) as string;
        }

        public static void SetStartup(bool on)
        {
            using (RegistryKey k = Registry.CurrentUser.CreateSubKey(RunKey))
            {
                if (on) k.SetValue(RunName, "\"" + Application.ExecutablePath + "\"");
                else k.DeleteValue(RunName, false);
            }
        }

        public void Quit()
        {
            if (MessageBox.Show("終了すると作業の記録が止まります。終了しますか？", "作業ログ", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return;
            Shutdown();
        }

        // 聞かずに終わる（更新で起動し直すときも）
        void Shutdown()
        {
            if (!off) logger.CloseWindows(DateTime.Now);
            if (form != null && !form.IsDisposed) { form.SaveIfDirty(); form.Dispose(); }
            if (aiForm != null && !aiForm.IsDisposed) aiForm.Dispose();
            if (settingsForm != null && !settingsForm.IsDisposed) settingsForm.Dispose();
            if (askForm != null && !askForm.IsDisposed) askForm.Dispose();
            afkTimer.Stop();
            timer.Stop();
            tray.Visible = false;
            tray.Dispose();
            offIcon.Dispose();
            ExitThread();
        }

        static Icon MakeIcon(Color back)
        {
            using (Bitmap bmp = Logo(32, back))
                return Icon.FromHandle(bmp.GetHicon());
        }

        // アイコンの絵（角の丸い四角に白の「報」）。size は 32 を基準に拡大する
        public static Bitmap Logo(int size, Color back)
        {
            var bmp = new Bitmap(size, size);
            float k = size / 32f;
            using (Graphics g = Graphics.FromImage(bmp))
            using (var path = new GraphicsPath())
            using (var bg = new SolidBrush(back))
            using (var f = new Font("Meiryo UI", 20f * k, FontStyle.Bold, GraphicsUnit.Pixel))
            using (var sf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center })
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.TextRenderingHint = TextRenderingHint.AntiAlias;
                float r = 14 * k, a = 1 * k, b = 31 * k;
                path.AddArc(a, a, r, r, 180, 90);
                path.AddArc(b - r, a, r, r, 270, 90);
                path.AddArc(b - r, b - r, r, r, 0, 90);
                path.AddArc(a, b - r, r, r, 90, 90);
                path.CloseFigure();
                g.FillPath(bg, path);
                g.DrawString("報", f, Brushes.White, new RectangleF(0, k, size, size), sf);
            }
            return bmp;
        }
    }

    // 要望・不具合の報告: GitHub の Issue の作成画面を、書いた内容を入れた状態でブラウザに開く
    // （送るには GitHub のアカウントが要る。アカウントなしで送れる窓口はないので、トークンを exe に入れて代わりに送ることはしない）
    static class Feedback
    {
        static string E(string s) { return Uri.EscapeDataString(s ?? ""); }

        static string FirstLine(string text, int max)
        {
            string l = (text ?? "").Replace("\r\n", "\n").Split('\n').Select(x => x.Trim()).FirstOrDefault(x => x.Length > 0) ?? "";
            return l.Length > max ? l.Substring(0, max) + "…" : l;
        }

        public static string Env() { return "作業ログ・日報 " + Core.Version + " / " + Environment.OSVersion.VersionString; }

        // 作成画面の URL（.github/ISSUE_TEMPLATE/request.yml の欄に入れる）。長すぎるときは内容を途中で切り、truncated を立てる
        public static string Url(string kind, string name, string body, bool withEnv, out bool truncated)
        {
            string title = "[" + kind + "] " + FirstLine(body, 40);
            string env = withEnv ? Env() : "";
            string b = body.Trim();
            truncated = false;
            Func<string, string> build = text => "https://github.com/" + Updater.Repo + "/issues/new?template=request.yml" +
                "&title=" + E(title) + "&name=" + E(name) + "&kind=" + E(kind) + "&details=" + E(text) + "&version=" + E(env);
            string url = build(b);
            while (url.Length > 7000 && b.Length > 20)
            {
                b = b.Substring(0, b.Length * 3 / 4);
                truncated = true;
                url = build(b + "…\n\n（長いので途中までです。続きはクリップボードにあるので貼ってください）");
            }
            return url;
        }

        // 送った内容は保存フォルダの「要望」フォルダにも残す
        public static void SaveCopy(string kind, string name, string body, bool withEnv)
        {
            string dir = Path.Combine(Core.Dir, "要望");
            Directory.CreateDirectory(dir);
            Core.WriteText(Path.Combine(dir, DateTime.Now.ToString("yyyyMMdd-HHmmss") + ".txt"),
                "種類: " + kind + "\r\n名前: " + name + "\r\n" + (withEnv ? "版: " + Env() + "\r\n" : "") + "\r\n" + body.Trim().Replace("\r\n", "\n").Replace("\n", "\r\n") + "\r\n");
        }
    }

    // 1日の状態（1分ごと）。AI が日報の材料にするのは「作業」だけで、休憩・離席・退勤のあいだは AI を止めている
    enum Slot { None, Work, Idle, Break, Afk, Off }

    static class DayStatus
    {
        public static readonly string[] Names = { "記録なし", "作業", "無操作", "休憩", "離席", "退勤" };
        public static readonly Color[] Colors =
        {
            Color.FromArgb(232, 234, 238), Color.FromArgb(30, 111, 217), Color.FromArgb(160, 196, 238),
            Color.FromArgb(240, 160, 60), Color.FromArgb(150, 120, 200), Color.FromArgb(95, 99, 108)
        };

        // 0:00〜24:00 の 1440 分。あとに書いたものほど優先（記録 → 休憩 → 離席 → 退勤）
        public static Slot[] Build(string date, Settings s)
        {
            var a = new Slot[1440];
            DateTime day = DateTime.ParseExact(date, "yyyy-MM-dd", CultureInfo.InvariantCulture);
            Action<DateTime, DateTime, Slot> fill = (t0, t1, v) =>
            {
                int i0 = Math.Max(0, (int)Math.Floor((t0 - day).TotalMinutes)), i1 = Math.Min(1440, (int)Math.Ceiling((t1 - day).TotalMinutes));
                for (int i = i0; i < i1; i++) a[i] = v;
            };
            foreach (KeyValuePair<DateTime, long> r in Reporter.ReadIdle(date))
            {
                int i = (int)(r.Key - day).TotalMinutes;
                if (i >= 0 && i < 1440) a[i] = r.Value >= s.IdleLimit ? Slot.Idle : Slot.Work;
            }
            if (Reporter.BreakOn(date, s)) fill(day + s.BreakStart, day + s.BreakEnd, Slot.Break);
            foreach (AfkEntry e in Afk.Load(date))
                fill(e.Start, e.End, e.Skip && e.Memo.StartsWith("休憩") ? Slot.Break : Slot.Afk);
            // 退勤していた時間帯（勤務履歴.log）。今日なら今まで
            DateTime end = day.AddDays(1), now = DateTime.Now;
            if (end > now) end = now;
            bool offNow = false;
            DateTime offFrom = day;
            foreach (KeyValuePair<DateTime, bool> h in Duty.History())
            {
                if (h.Key >= end) break;
                if (h.Key < day) { offNow = h.Value; offFrom = day; continue; }
                if (h.Value && !offNow) { offNow = true; offFrom = h.Key; }
                else if (!h.Value && offNow) { fill(offFrom, h.Key, Slot.Off); offNow = false; }
            }
            if (offNow && Duty.History().Count > 0) fill(offFrom, end, Slot.Off);
            return a;
        }

        public static string Span(int minutes)
        {
            return minutes >= 60 ? minutes / 60 + "時間" + (minutes % 60 > 0 ? minutes % 60 + "分" : "") : minutes + "分";
        }
    }

    // 日報画面の上の帯: その日の作業・休憩・離席・退勤を色で、AI を呼んだ時刻を印で出す。今日は AI の状態（休止中か、次にいつ動くか）も
    class DayTimeline : Control
    {
        readonly Func<string> aiState;
        readonly ToolTip tip = new ToolTip();
        string date;
        Slot[] slots = new Slot[1440];
        List<KeyValuePair<DateTime, bool>> calls = new List<KeyValuePair<DateTime, bool>>();
        int v0 = 8 * 60, v1 = 18 * 60;   // 見せる範囲（分。v0〜v1）
        string head = "", sum = "", lastTip;

        public DayTimeline(Func<string> aiState)
        {
            this.aiState = aiState;
            DoubleBuffered = true;
            SetStyle(ControlStyles.ResizeRedraw, true);
            Dock = DockStyle.Top;
            Height = 80;   // 画面の拡大率に合わせて、日報画面が大きさを直す
            BackColor = Color.FromArgb(244, 247, 252);
        }

        public string Date { get { return date; } }

        public void SetDate(string d)
        {
            date = d;
            Reload();
        }

        // 記録・離席メモ・休憩・勤務の履歴と AI の呼び出しを読み直す（今日の分は記録のたびに呼ぶ）
        public void Reload()
        {
            if (date == null) return;
            Settings s = Settings.Get();
            try
            {
                slots = DayStatus.Build(date, s);
                calls = ClaudeCli.CallsOn(DateTime.ParseExact(date, "yyyy-MM-dd", CultureInfo.InvariantCulture));
            }
            catch (Exception ex) { Core.Error("帯", ex.Message); slots = new Slot[1440]; calls.Clear(); }
            bool today = date == DateTime.Today.ToString("yyyy-MM-dd");
            int first = Array.FindIndex(slots, x => x == Slot.Work || x == Slot.Idle);
            int last = Array.FindLastIndex(slots, x => x == Slot.Work || x == Slot.Idle || x == Slot.Afk);
            int lo = (int)s.WorkStart.TotalMinutes, hi = (int)s.ReportTime.TotalMinutes;
            if (first >= 0) lo = Math.Min(lo, first);
            if (last >= 0) hi = Math.Max(hi, last + 1);
            if (today) hi = Math.Max(hi, (int)DateTime.Now.TimeOfDay.TotalMinutes + 1);
            v0 = Math.Max(0, lo / 60 * 60);
            v1 = Math.Min(1440, (hi + 59) / 60 * 60);
            if (v1 - v0 < 120) v1 = Math.Min(1440, v0 + 120);

            int[] n = new int[6];
            for (int i = v0; i < v1; i++) n[(int)slots[i]]++;
            int rest = n[(int)Slot.Break] + n[(int)Slot.Afk] + n[(int)Slot.Off];
            var parts = new List<string>();
            foreach (Slot k in new[] { Slot.Break, Slot.Afk, Slot.Off })
                if (n[(int)k] > 0) parts.Add(DayStatus.Names[(int)k] + " " + DayStatus.Span(n[(int)k]));
            sum = "作業 " + DayStatus.Span(n[(int)Slot.Work]) + "　AI 休止 " + DayStatus.Span(rest) +
                  (parts.Count > 0 ? "（" + string.Join("・", parts) + "）" : "") + "　AI 呼び出し " + calls.Count + "回";
            head = today && aiState != null ? aiState() : date + " の記録";
            FitHeight();
            Invalidate();
        }

        int Scale(int px) { return (int)Math.Round(px * DeviceDpiOf() / 96f); }

        float dpi;
        float DeviceDpiOf()
        {
            if (dpi == 0) using (Graphics g = CreateGraphics()) dpi = g.DpiX;
            return dpi;
        }

        // 行の並び（上から）: AI の状態 / 合計 / ▼の列と帯 / 時刻 / 凡例。文字の高さを測って重ならないように積む
        Font boldFont, smallFont;
        Font Bold { get { return boldFont ?? (boldFont = new Font(Font, FontStyle.Bold)); } }
        Font Small { get { return smallFont ?? (smallFont = new Font(Font.FontFamily, Font.Size * 0.9f)); } }
        const TextFormatFlags Flags = TextFormatFlags.NoPadding | TextFormatFlags.SingleLine | TextFormatFlags.NoPrefix;

        int Pad { get { return Scale(10); } }
        int HeadH { get { return TextRenderer.MeasureText("あ", Bold, Size.Empty, Flags).Height; } }
        int SmallH { get { return TextRenderer.MeasureText("あ", Small, Size.Empty, Flags).Height; } }
        int SumTop { get { return Scale(6) + HeadH + Scale(2); } }
        Rectangle Bar { get { return new Rectangle(Pad, SumTop + SmallH + Scale(4) + Scale(9), Math.Max(10, Width - Pad * 2), Scale(16)); } }
        int HourTop { get { return Bar.Bottom + Scale(4); } }
        int LegendTop { get { return HourTop + SmallH + Scale(4); } }

        // 凡例の項目（色の四角と名前）を左から並べ、入りきらなければ次の行へ。各項目の位置を返す
        List<KeyValuePair<Point, int>> LegendLayout(out int lines)
        {
            var pos = new List<KeyValuePair<Point, int>>();
            int x = Pad, y = LegendTop, box = Scale(11), gap = Scale(14);
            lines = 1;
            for (int i = 0; i < LegendItems.Length; i++)
            {
                int w = box + Scale(4) + TextRenderer.MeasureText(LegendItems[i].Key, Small, Size.Empty, Flags).Width;
                if (x + w > Width - Pad && x > Pad) { x = Pad; y += SmallH + Scale(3); lines++; }
                pos.Add(new KeyValuePair<Point, int>(new Point(x, y), i));
                x += w + gap;
            }
            return pos;
        }

        static readonly KeyValuePair<string, Color>[] LegendItems =
        {
            new KeyValuePair<string, Color>("作業（日報の材料）", DayStatus.Colors[(int)Slot.Work]),
            new KeyValuePair<string, Color>("無操作", DayStatus.Colors[(int)Slot.Idle]),
            new KeyValuePair<string, Color>("休憩", DayStatus.Colors[(int)Slot.Break]),
            new KeyValuePair<string, Color>("離席", DayStatus.Colors[(int)Slot.Afk]),
            new KeyValuePair<string, Color>("退勤", DayStatus.Colors[(int)Slot.Off]),
            new KeyValuePair<string, Color>("▼ AI を呼んだ時刻", Color.FromArgb(30, 150, 70)),
        };

        // 中身に合わせて高さを決める（幅が変わると凡例の折り返しが変わる）
        void FitHeight()
        {
            if (!IsHandleCreated || Width <= 0) return;
            int lines;
            LegendLayout(out lines);
            int h = LegendTop + lines * (SmallH + Scale(3)) + Scale(4);
            if (Height != h) Height = h;
        }

        protected override void OnHandleCreated(EventArgs e) { base.OnHandleCreated(e); FitHeight(); }
        protected override void OnResize(EventArgs e) { base.OnResize(e); FitHeight(); }
        protected override void OnFontChanged(EventArgs e) { boldFont = smallFont = null; base.OnFontChanged(e); FitHeight(); }

        float X(int minute) { Rectangle b = Bar; return b.Left + (minute - v0) * (float)b.Width / (v1 - v0); }

        static string Fit(string text, Font f, int width)
        {
            if (TextRenderer.MeasureText(text, f, Size.Empty, Flags).Width <= width) return text;
            while (text.Length > 1 && TextRenderer.MeasureText(text + "…", f, Size.Empty, Flags).Width > width) text = text.Substring(0, text.Length - 1);
            return text + "…";
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            Rectangle b = Bar;
            int width = Width - Pad * 2;
            Color gray = Color.FromArgb(80, 80, 80);
            bool pause = head.StartsWith("AI 休止中");
            TextRenderer.DrawText(g, Fit(head, Bold, width), Bold, new Point(Pad, Scale(6)),
                pause ? Color.FromArgb(190, 100, 10) : Color.FromArgb(30, 90, 170), Flags);
            TextRenderer.DrawText(g, Fit(sum, Small, width), Small, new Point(Pad, SumTop), gray, Flags);

            g.SmoothingMode = SmoothingMode.None;
            using (var bg = new SolidBrush(DayStatus.Colors[0])) g.FillRectangle(bg, b);
            for (int i = v0; i < v1; )
            {
                int j = i;
                while (j < v1 && slots[j] == slots[i]) j++;
                if (slots[i] != Slot.None)
                    using (var br = new SolidBrush(DayStatus.Colors[(int)slots[i]]))
                        g.FillRectangle(br, X(i), b.Top, X(j) - X(i) + 0.5f, b.Height);
                i = j;
            }
            g.SmoothingMode = SmoothingMode.AntiAlias;
            // 1時間ごとの細い線（帯の上）
            using (var pen = new Pen(Color.FromArgb(90, 255, 255, 255)))
                for (int m = v0 + 60; m < v1; m += 60) g.DrawLine(pen, X(m), b.Top, X(m), b.Bottom - 1);
            // 今の時刻
            if (date == DateTime.Today.ToString("yyyy-MM-dd"))
            {
                float x = X((int)DateTime.Now.TimeOfDay.TotalMinutes);
                using (var pen = new Pen(Color.FromArgb(210, 40, 40), Scale(2))) g.DrawLine(pen, x, b.Top - Scale(2), x, b.Bottom + Scale(2));
            }
            // AI を呼んだ時刻（帯のすぐ上の列。緑は成功、赤は失敗）
            foreach (KeyValuePair<DateTime, bool> c in calls)
            {
                int m = (int)c.Key.TimeOfDay.TotalMinutes;
                if (m < v0 || m >= v1) continue;
                float x = X(m), top = b.Top - Scale(9);
                using (var br = new SolidBrush(c.Value ? Color.FromArgb(30, 150, 70) : Color.FromArgb(210, 40, 40)))
                    g.FillPolygon(br, new[] { new PointF(x - Scale(4), top), new PointF(x + Scale(4), top), new PointF(x, b.Top - Scale(2)) });
            }
            // 時刻: 文字が重ならない間隔（1・2・3・4・6時間ごと）で出す
            int labelW = TextRenderer.MeasureText("00:00", Small, Size.Empty, Flags).Width + Scale(10);
            float perHour = b.Width * 60f / (v1 - v0);
            int step = new[] { 1, 2, 3, 4, 6, 12 }.FirstOrDefault(h => perHour * h >= labelW);
            if (step == 0) step = 24;
            using (var pen = new Pen(Color.FromArgb(150, 150, 150)))
                for (int m = v0; m <= v1; m += 60)
                {
                    float x = X(m);
                    g.DrawLine(pen, x, b.Bottom, x, b.Bottom + Scale(3));
                    if ((m / 60) % step != 0) continue;
                    string t = (m / 60) + ":00";
                    int tw = TextRenderer.MeasureText(t, Small, Size.Empty, Flags).Width;
                    int tx = (int)Math.Min(Math.Max(0, x - tw / 2f), Width - tw);
                    TextRenderer.DrawText(g, t, Small, new Point(tx, HourTop), gray, Flags);
                }
            // 凡例
            int lines;
            foreach (KeyValuePair<Point, int> p in LegendLayout(out lines))
            {
                KeyValuePair<string, Color> it = LegendItems[p.Value];
                int box = Scale(11), ty = p.Key.Y, by = ty + (SmallH - box) / 2;
                if (it.Key.StartsWith("▼"))
                {
                    using (var br = new SolidBrush(it.Value))
                        g.FillPolygon(br, new[] { new PointF(p.Key.X, by + Scale(2)), new PointF(p.Key.X + box, by + Scale(2)), new PointF(p.Key.X + box / 2f, by + box) });
                    TextRenderer.DrawText(g, it.Key.Substring(1).Trim(), Small, new Point(p.Key.X + box + Scale(4), ty), gray, Flags);
                }
                else
                {
                    using (var br = new SolidBrush(it.Value)) g.FillRectangle(br, p.Key.X, by, box, box);
                    TextRenderer.DrawText(g, it.Key, Small, new Point(p.Key.X + box + Scale(4), ty), gray, Flags);
                }
            }
        }

        // マウスを載せた時刻の状態を出す
        protected override void OnMouseMove(MouseEventArgs e)
        {
            base.OnMouseMove(e);
            Rectangle b = Bar;
            string t = null;
            if (e.X >= b.Left && e.X < b.Right && e.Y >= b.Top - Scale(10) && e.Y <= b.Bottom + Scale(4))
            {
                int m = v0 + (int)((e.X - b.Left) * (v1 - v0) / (float)b.Width);
                m = Math.Max(v0, Math.Min(v1 - 1, m));
                Slot v = slots[m];
                t = (m / 60).ToString("00") + ":" + (m % 60).ToString("00") + "　" + DayStatus.Names[(int)v] +
                    (v == Slot.Break || v == Slot.Afk || v == Slot.Off ? "（AI 休止）" : v == Slot.Work ? "（日報の材料）" : "");
                KeyValuePair<DateTime, bool>[] near = calls.Where(c => Math.Abs((int)c.Key.TimeOfDay.TotalMinutes - m) <= 2).ToArray();
                if (near.Length > 0) t += "\n" + string.Join("\n", near.Select(c => "AI " + c.Key.ToString("HH:mm") + (c.Value ? " 成功" : " 失敗")));
            }
            if (t == lastTip) return;
            lastTip = t;
            if (t == null) tip.Hide(this); else tip.Show(t, this, e.X + Scale(12), e.Y + Scale(16));
        }

        protected override void OnMouseLeave(EventArgs e) { base.OnMouseLeave(e); lastTip = null; tip.Hide(this); }

        protected override void Dispose(bool disposing)
        {
            if (disposing) { tip.Dispose(); if (boldFont != null) boldFont.Dispose(); if (smallFont != null) smallFont.Dispose(); }
            base.Dispose(disposing);
        }
    }

    // タイトルとタスク: 最近のウィンドウタイトルの一覧に、正しいタスク名を付ける。手で付けた対応が Tasks.LearnMin 件たまったら、残りは AI が予測する
    class TaskTitleForm : Form
    {
        static readonly Color AiColor = Color.FromArgb(30, 90, 170), WordColor = Color.Gray;
        const int TaskCol = 5, HowCol = 6;
        readonly DataGridView grid = new DataGridView();
        readonly Label info = new Label { Dock = DockStyle.Top, AutoSize = false, Height = 30, Padding = new Padding(8, 8, 8, 4) };
        readonly Button predictBtn;
        readonly AutoCompleteStringCollection names = new AutoCompleteStringCollection();
        List<Tasks.TitleTask> titles = new List<Tasks.TitleTask>();
        bool loading, predicting, autoTried;

        public TaskTitleForm(Icon icon)
        {
            Text = "タイトルとタスク" + (Core.Test ? "（テスト）" : "");
            Icon = icon;
            AutoScaleDimensions = new SizeF(96f, 96f);
            AutoScaleMode = AutoScaleMode.Dpi;
            Font = new Font("Meiryo UI", 10f);
            ClientSize = new Size(1000, 640);
            MinimumSize = new Size(700, 420);
            StartPosition = FormStartPosition.CenterScreen;

            grid.Dock = DockStyle.Fill;
            grid.BackgroundColor = SystemColors.Window;
            grid.RowHeadersVisible = false;
            grid.AllowUserToAddRows = false;
            grid.AllowUserToDeleteRows = false;
            grid.AllowUserToResizeRows = false;
            grid.SelectionMode = DataGridViewSelectionMode.CellSelect;
            grid.Columns.Add(GridSort.NoColumn());
            grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "使った時間", ReadOnly = true, Width = 90, DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleRight } });
            grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "最近", ReadOnly = true, Width = 95 });
            grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "アプリ", ReadOnly = true, Width = 110 });
            grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "ウィンドウタイトル", ReadOnly = true, AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill });
            grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "タスク", Width = 220 });
            grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "付け方", ReadOnly = true, Width = 90 });
            GridSort.Enable(grid);   // 行の Tag に TitleTask を持つので、並べ替えても対応はずれない
            grid.EditingControlShowing += (o, e) =>
            {
                var tb = e.Control as TextBox;
                if (tb == null) return;
                bool task = grid.CurrentCell != null && grid.CurrentCell.ColumnIndex == TaskCol;
                tb.AutoCompleteMode = task ? AutoCompleteMode.SuggestAppend : AutoCompleteMode.None;
                tb.AutoCompleteSource = task ? AutoCompleteSource.CustomSource : AutoCompleteSource.None;
                if (task) tb.AutoCompleteCustomSource = names;
            };
            grid.CellValueChanged += (o, e) => { if (!loading && e.RowIndex >= 0 && e.ColumnIndex == TaskCol) Edited(e.RowIndex); };

            var bar = new FlowLayoutPanel { Dock = DockStyle.Bottom, AutoSize = true, Padding = new Padding(6, 6, 6, 4) };
            AddButton(bar, "読み直す", (o, e) => Reload());
            predictBtn = AddButton(bar, "AI で残りを予測する", (o, e) => RunPredict(true));
            AddButton(bar, "閉じる", (o, e) => Close());

            Controls.Add(grid);
            Controls.Add(info);
            Controls.Add(bar);
            Shown += (o, e) => Reload();
        }

        static Button AddButton(FlowLayoutPanel bar, string text, EventHandler click)
        {
            var b = new Button { Text = text, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Padding = new Padding(8, 2, 8, 2), Margin = new Padding(3) };
            b.Click += click;
            bar.Controls.Add(b);
            return b;
        }

        // 最近14日のタイトル（使った時間の長い順に400件まで）と、今の対応を出す
        void Reload()
        {
            Settings s = Settings.Get();
            List<KeyValuePair<Tasks.TitleTask, int>> recent = Tasks.RecentTitles(14, s).Take(400).ToList();
            Dictionary<string, Tasks.TitleTask> map = Tasks.LoadMap();
            List<TaskItem> tasks = Tasks.Load();
            names.Clear();
            names.AddRange(tasks.Where(t => !t.Done).Select(t => t.Name.Trim()).Concat(map.Values.Select(x => x.Task)).Concat(new[] { Tasks.NoTask }).Distinct().ToArray());
            loading = true;
            grid.Rows.Clear();
            titles = recent.Select(r => r.Key).ToList();
            foreach (KeyValuePair<Tasks.TitleTask, int> r in recent)
            {
                Tasks.TitleTask m;
                string task = "", how = "";
                Color color = SystemColors.WindowText;
                if (map.TryGetValue(Tasks.Key(r.Key.Proc, r.Key.Title), out m)) { task = m.Task; how = m.ByAi ? "AI" : "手"; if (m.ByAi) color = AiColor; }
                else
                {
                    string w = Tasks.Match(tasks, r.Key.Proc, r.Key.Title);
                    if (w != null) { task = w; how = "語"; color = WordColor; }
                }
                int i = grid.Rows.Add(null, null, null, r.Key.Proc, r.Key.Title, task, how);
                grid.Rows[i].Tag = r.Key;
                GridSort.Set(grid.Rows[i].Cells[0], i + 1, i + 1);
                GridSort.Set(grid.Rows[i].Cells[1], DayStatus.Span(r.Value), r.Value);
                GridSort.Set(grid.Rows[i].Cells[2], r.Key.Last == DateTime.MinValue ? "" : r.Key.Last.ToString("MM/dd HH:mm"), r.Key.Last);
                grid.Rows[i].Cells[TaskCol].Style.ForeColor = color;
            }
            loading = false;
            UpdateInfo(map);
            // 手本がたまっていて、対応のないタイトルがあれば、開いたときに1回だけ AI が予測する
            if (!autoTried && Tasks.UserCount(map) >= Tasks.LearnMin && titles.Any(t => !map.ContainsKey(Tasks.Key(t.Proc, t.Title))))
            {
                autoTried = true;
                RunPredict(false);
            }
        }

        void UpdateInfo(Dictionary<string, Tasks.TitleTask> map)
        {
            int user = Tasks.UserCount(map);
            int ai = titles.Count(t => { Tasks.TitleTask m; return map.TryGetValue(Tasks.Key(t.Proc, t.Title), out m) && m.ByAi; });
            int none = titles.Count(t => !map.ContainsKey(Tasks.Key(t.Proc, t.Title)));
            info.Text = "最近14日　手 " + user + " 件・AI " + ai + " 件・未入力 " + none + " 件" +
                        (user < Tasks.LearnMin ? "　（AI の予測まで あと " + (Tasks.LearnMin - user) + " 件）" : "");
            predictBtn.Enabled = !predicting && user >= Tasks.LearnMin && none > 0;
        }

        // タスク名が書かれた・直された: 手で付けた対応として残す（空なら対応を消す）。新しい名前はタスクの一覧にも足す
        void Edited(int row)
        {
            var t = (Tasks.TitleTask)grid.Rows[row].Tag;
            string task = Convert.ToString(grid.Rows[row].Cells[TaskCol].Value).Trim();
            try
            {
                Tasks.SetTitle(t.Proc, t.Title, task, false);
                List<TaskItem> list = Tasks.Load();
                if (task.Length > 0 && task != Tasks.NoTask && !list.Any(x => x.Name.Trim() == task))
                {
                    list.Add(new TaskItem { Name = task });
                    Tasks.Save(list);
                    names.Add(task);
                }
            }
            catch (Exception ex) { MessageBox.Show(this, "保存できません: " + ex.Message, "タイトルとタスク"); return; }
            loading = true;
            grid.Rows[row].Cells[TaskCol].Style.ForeColor = SystemColors.WindowText;
            grid.Rows[row].Cells[HowCol].Value = task.Length > 0 ? "手" : "";
            loading = false;
            UpdateInfo(Tasks.LoadMap());
        }

        void RunPredict(bool byUser)
        {
            if (predicting) return;
            predicting = true;
            predictBtn.Enabled = false;
            info.Text = "AI で予測中…";
            List<Tasks.TitleTask> list = titles.ToList();
            var th = new Thread(() =>
            {
                int n = 0;
                string error = null;
                try { n = Tasks.Predict(list, Settings.Get()); }
                catch (Exception ex) { error = ex.Message; Core.Error("タスクの予測", ex.Message); }
                if (IsDisposed) return;
                BeginInvoke((Action)(() =>
                {
                    predicting = false;
                    Reload();
                    if (error != null) MessageBox.Show(this, "予測できません: " + error, "タイトルとタスク", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    else if (byUser) MessageBox.Show(this, "AI の予測: " + n + " 件", "タイトルとタスク");
                }));
            });
            th.IsBackground = true;
            th.Start();
        }
    }

    // 表の並べ替え: 見出しのクリックで並べ替える。No 列（元の並び順）の見出しで元に戻せる。
    // セルの Tag に値（分・日時など）があればそれで比べる。チェックの列は印の有無で比べる
    static class GridSort
    {
        public static DataGridViewTextBoxColumn NoColumn()
        {
            var c = new DataGridViewTextBoxColumn { HeaderText = "No", ReadOnly = true, AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells };
            c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            c.DefaultCellStyle.ForeColor = Color.DimGray;
            return c;
        }

        public static void Enable(DataGridView g)
        {
            foreach (DataGridViewColumn c in g.Columns) c.SortMode = DataGridViewColumnSortMode.Automatic;
            g.SortCompare += (o, e) =>
            {
                object a = g.Rows[e.RowIndex1].Cells[e.Column.Index].Tag, b = g.Rows[e.RowIndex2].Cells[e.Column.Index].Tag;
                if (a is IComparable && b != null && a.GetType() == b.GetType()) e.SortResult = ((IComparable)a).CompareTo(b);
                else if (a is IComparable && b == null) e.SortResult = 1;
                else if (a == null && b is IComparable) e.SortResult = -1;
                else if (e.CellValue1 is bool || e.CellValue2 is bool) e.SortResult = (e.CellValue1 is bool && (bool)e.CellValue1).CompareTo(e.CellValue2 is bool && (bool)e.CellValue2);
                else e.SortResult = string.Compare(Convert.ToString(e.CellValue1), Convert.ToString(e.CellValue2), StringComparison.CurrentCulture);
                e.Handled = true;
            };
        }

        // セルに値と、並べ替えに使う値を入れる
        public static void Set(DataGridViewCell c, object shown, IComparable key) { c.Value = shown; c.Tag = key; }
    }

    // ListView の並べ替え（見出しのクリック。もう一度で逆順）。数字の列は数で比べる
    class ListSorter : IComparer
    {
        public int Column;
        public bool Desc;

        public static void Enable(ListView v)
        {
            var s = new ListSorter();
            v.ListViewItemSorter = s;
            v.ColumnClick += (o, e) =>
            {
                if (s.Column == e.Column) s.Desc = !s.Desc; else { s.Column = e.Column; s.Desc = false; }
                v.Sort();
            };
        }

        public int Compare(object x, object y)
        {
            var a = (ListViewItem)x; var b = (ListViewItem)y;
            string sa = Column < a.SubItems.Count ? a.SubItems[Column].Text : "", sb = Column < b.SubItems.Count ? b.SubItems[Column].Text : "";
            long na, nb;
            int r = long.TryParse(sa.Replace(",", ""), out na) && long.TryParse(sb.Replace(",", ""), out nb) ? na.CompareTo(nb) : string.Compare(sa, sb, StringComparison.CurrentCulture);
            return Desc ? -r : r;
        }
    }

    // 時刻の入力欄。4桁（1430）で 14:30。ホイールか ↑↓ で、時の上（左側）なら1時間、分の上（右側）なら1分ずつ変わる。
    // Shift か Ctrl を押しながらのホイールは、どこでも1時間ずつ。空なら時刻なし
    class TimeBox : TextBox
    {
        static readonly Regex Re = new Regex(@"^(\d{1,2}):?(\d{2})$");
        bool formatting;

        public TimeBox()
        {
            Width = 62;
            MaxLength = 5;
            TextAlign = HorizontalAlignment.Center;
            ImeMode = ImeMode.Disable;   // 全角の数字にならないように
        }

        public static bool TryParse(string text, out TimeSpan t)
        {
            t = TimeSpan.Zero;
            Match m = Re.Match((text ?? "").Trim());
            if (!m.Success) return false;
            int h = int.Parse(m.Groups[1].Value), mi = int.Parse(m.Groups[2].Value);
            if (h > 23 || mi > 59) return false;
            t = new TimeSpan(h, mi, 0);
            return true;
        }

        public TimeSpan? Value
        {
            get { TimeSpan t; return TryParse(Text, out t) ? t : (TimeSpan?)null; }
            set { Text = value.HasValue ? value.Value.ToString(@"hh\:mm") : ""; }
        }

        public bool Valid { get { return Text.Trim().Length == 0 || Value.HasValue; } }

        public void SetNow() { Value = new TimeSpan(DateTime.Now.Hour, DateTime.Now.Minute, 0); }

        protected override void OnTextChanged(EventArgs e)
        {
            base.OnTextChanged(e);
            if (formatting) return;
            // 4桁そろったら 14:30 の形にする
            if (Regex.IsMatch(Text, @"^\d{4}$") && Value.HasValue)
            {
                formatting = true;
                Value = Value;
                SelectionStart = TextLength;
                formatting = false;
            }
            ForeColor = Valid ? SystemColors.WindowText : Color.Firebrick;
        }

        protected override void OnLeave(EventArgs e)
        {
            base.OnLeave(e);
            if (Value.HasValue) Value = Value;   // 930 → 09:30
        }

        protected override void OnKeyPress(KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && e.KeyChar != ':') e.Handled = true;
            base.OnKeyPress(e);
        }

        void Step(int n, bool hour)
        {
            int caret = SelectionStart;
            TimeSpan t = Value ?? new TimeSpan(DateTime.Now.Hour, DateTime.Now.Minute, 0);
            int m = ((int)t.TotalMinutes + n * (hour ? 60 : 1)) % 1440;
            if (m < 0) m += 1440;
            Value = TimeSpan.FromMinutes(m);
            SelectionStart = Math.Min(caret, TextLength);
        }

        // マウスが「時」の上にあるか（「:」より左）。文字の位置がわからなければ欄の左半分
        bool OverHour()
        {
            int x = PointToClient(Cursor.Position).X;
            if (TextLength == 5)
            {
                Point colon = GetPositionFromCharIndex(2);
                if (colon.X > 0) return x < colon.X + 2;
            }
            return x < ClientSize.Width / 2;
        }

        protected override void OnMouseWheel(MouseEventArgs e)
        {
            bool hour = (ModifierKeys & (Keys.Shift | Keys.Control)) != 0 || OverHour();
            Step(e.Delta > 0 ? 1 : -1, hour);
            var h = e as HandledMouseEventArgs;
            if (h != null) h.Handled = true;
        }

        protected override void OnKeyDown(KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Up || e.KeyCode == Keys.Down)
            {
                Step(e.KeyCode == Keys.Up ? 1 : -1, e.Shift || e.Control || (Text.Length == 5 && SelectionStart <= 2));
                e.Handled = e.SuppressKeyPress = true;
                return;
            }
            base.OnKeyDown(e);
        }
    }

    // 日報画面（右側に「AIで直す」欄。AI の案は本文にそのまま出し、変わった行を青字にする）
    class MainForm : Form
    {
        class DateItem
        {
            public string Date;
            public bool HasReport;
            public override string ToString()
            {
                DateTime d = DateTime.ParseExact(Date, "yyyy-MM-dd", CultureInfo.InvariantCulture);
                return Date + "（" + d.ToString("ddd", new CultureInfo("ja-JP")) + "）" + (HasReport ? "" : " 未作成");
            }
        }

        static readonly Color AiBlue = Color.FromArgb(0, 84, 210);

        readonly TrayApp app;
        readonly ListBox list = new ListBox();
        // 日本語の文字で自動でフォントを切り替えると行の高さ・描画が乱れるので、UI のフォントのままにする
        readonly RichTextBox text = new RichTextBox { DetectUrls = false, LanguageOption = RichTextBoxLanguageOptions.UIFonts };
        readonly ToolStripStatusLabel message = new ToolStripStatusLabel();
        readonly ToolStripStatusLabel record = new ToolStripStatusLabel();
        DayTimeline timeline;
        readonly Button copyBullets, copyAll, saveBtn, redoBtn, todayBtn, linkBtn;
        readonly ToolStripMenuItem mSave, mCopyBullets, mCopyAll, mAiFix, mRedo, mToday, mError;

        // 「AIで直す」欄
        readonly Label aiTitle = new Label { AutoSize = true, Margin = new Padding(3, 6, 3, 2) };
        readonly TextBox instruction = new TextBox { Multiline = true, AcceptsReturn = true, ScrollBars = ScrollBars.Vertical, Dock = DockStyle.Fill };
        readonly Button askBtn = new Button { Text = "送る（Enter）", AutoSize = true, Margin = new Padding(3, 2, 3, 2) };
        readonly Button clearChatBtn = new Button { Text = "会話を消す", AutoSize = true, Margin = new Padding(3, 2, 3, 2) };
        readonly RichTextBox chatView = new RichTextBox { ReadOnly = true, Dock = DockStyle.Fill, BorderStyle = BorderStyle.None, BackColor = Color.White, DetectUrls = false };
        List<KeyValuePair<bool, string>> chat = new List<KeyValuePair<bool, string>>();   // この日の会話（true が利用者）
        readonly Label aiStatus = new Label { AutoSize = true, Margin = new Padding(3, 4, 3, 4) };
        readonly Button acceptBtn = new Button { Text = "確定して保存", AutoSize = true, Margin = new Padding(3, 2, 3, 2) };
        readonly Button revertBtn = new Button { Text = "元に戻す", AutoSize = true, Margin = new Padding(3, 2, 3, 2) };
        readonly FlowLayoutPanel previewBar = new FlowLayoutPanel { AutoSize = true, WrapContents = false, Margin = new Padding(0) };
        readonly TextBox notes = new TextBox { Multiline = true, AcceptsReturn = true, ScrollBars = ScrollBars.Vertical, Dock = DockStyle.Fill };

        // 「離席メモ」欄
        readonly TabPage aiPage = new TabPage("AIで直す"), afkPage = new TabPage("離席メモ"), taskPage = new TabPage("タスク"), reqPage = new TabPage("要望");
        readonly DataGridView taskGrid = new DataGridView();
        readonly Label taskInfo = new Label { AutoSize = true, ForeColor = Color.DimGray, Margin = new Padding(3, 6, 3, 3) };
        bool taskLoading;
        readonly TextBox reqName = new TextBox { Dock = DockStyle.Fill, MaxLength = 50 };
        readonly ComboBox reqKind = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 140 };
        readonly TextBox reqText = new TextBox { Multiline = true, AcceptsReturn = true, ScrollBars = ScrollBars.Vertical, Dock = DockStyle.Fill };
        readonly CheckBox reqEnv = new CheckBox { Text = "アプリの版と Windows の版を添える", AutoSize = true, Checked = true, Margin = new Padding(3, 6, 3, 2) };
        readonly Button reqSend = new Button { Text = "GitHub に送る", AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Padding = new Padding(8, 2, 8, 2), Margin = new Padding(3, 6, 3, 3) };
        readonly Label reqStatus = new Label { AutoSize = true, ForeColor = Color.DimGray, Margin = new Padding(3, 4, 3, 3) };
        readonly Label afkHead = new Label { AutoSize = true, Margin = new Padding(3, 6, 3, 4) };
        readonly DataGridView afkGrid = new DataGridView();
        readonly TextBox freeMemo = new TextBox { Multiline = true, AcceptsReturn = true, ScrollBars = ScrollBars.Vertical, Dock = DockStyle.Fill };
        List<AfkEntry> afkList = new List<AfkEntry>();
        bool freeDirty;

        // 本文の下の「手で追記」欄と、その日の休憩
        readonly TimeBox startBox = new TimeBox(), endBox = new TimeBox();
        readonly TextBox addText = new TextBox { Dock = DockStyle.Fill, Margin = new Padding(3, 3, 3, 4) };
        readonly Button addBtn = new Button { Text = "追記", AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Padding = new Padding(10, 0, 10, 0), Margin = new Padding(3, 2, 3, 4) };
        readonly CheckBox breakCheck = new CheckBox { AutoSize = true, Anchor = AnchorStyles.Right, Margin = new Padding(12, 6, 3, 0) };
        readonly ToolTip tip = new ToolTip();
        readonly Button dutyBtn;
        readonly ToolStripMenuItem mDuty;

        [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern IntPtr SendMessage(IntPtr h, int msg, IntPtr w, string l);

        string current;
        bool hasReport, dirty, loading, busy, sample, notesDirty, aiRunning;
        DateTime loadedStamp;
        // AI の案を本文に出しているあいだ: 出す前の本文・訂正の記録と、出す前に未保存だったか
        string previewBase, previewNotes;
        bool previewDirty;

        // 見本の日報（保存しない・APIも使わない）
        const string SampleText =
            "日報 サンプル（PC稼働 08:45〜17:40／操作時間 約7.2時間）\r\n\r\n" +
            "■工程\r\n" +
            "・08:45〜10:30 A工区の数量計算\r\n" +
            "・10:40〜12:00 〇〇橋の図面確認\r\n" +
            "・13:00〜15:10 協議資料の作成\r\n" +
            "・15:20〜17:00 工程表の更新\r\n" +
            "■軽作業\r\n" +
            "・09:20 メール対応\r\n" +
            "・12:15 現場写真の整理\r\n" +
            "・15:15 見積書の確認\r\n" +
            "・17:35 メール対応\r\n\r\n" +
            "［これは見本です。保存されません。「箇条書きをコピー」を押すと「■」と「・」の行だけがコピーされます］";

        public MainForm(TrayApp app, Icon icon)
        {
            this.app = app;
            timeline = new DayTimeline(app.AiState);
            Text = "日報" + (Core.Test ? "（テスト）" : "");
            Icon = icon;
            AutoScaleDimensions = new SizeF(96f, 96f);
            AutoScaleMode = AutoScaleMode.Dpi;
            Font = new Font("Meiryo UI", 10f);
            ClientSize = new Size(1100, 560);
            MinimumSize = new Size(820, 420);
            StartPosition = FormStartPosition.CenterScreen;

            text.Dock = DockStyle.Fill;
            text.Font = new Font("Meiryo UI", 11f);
            text.ScrollBars = RichTextBoxScrollBars.Vertical;
            text.TextChanged += (o, e) => { if (!loading) { dirty = true; UpdateButtons(); } };
            // 貼り付けは文字だけ（ほかのアプリの書式を持ち込まない）
            text.KeyDown += (o, e) =>
            {
                if ((e.Control && e.KeyCode == Keys.V) || (e.Shift && e.KeyCode == Keys.Insert))
                {
                    text.Paste(DataFormats.GetFormat(DataFormats.UnicodeText));
                    e.Handled = true;
                }
            };

            list.Dock = DockStyle.Left;
            list.Width = 170;
            list.IntegralHeight = false;
            list.SelectedIndexChanged += (o, e) =>
            {
                var it = list.SelectedItem as DateItem;
                if (!loading && it != null && it.Date != current) ShowDate(it.Date);
            };

            var bar = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = 50, Padding = new Padding(6, 6, 6, 4), WrapContents = false };
            copyBullets = AddButton(bar, "箇条書きをコピー", (o, e) => CopyBullets());
            copyBullets.Font = new Font(Font, FontStyle.Bold);
            copyAll = AddButton(bar, "全文をコピー", (o, e) => Copy(Reporter.Unnumber(Body()), "全文をコピーしました"));
            saveBtn = AddButton(bar, "保存", (o, e) => SaveIfDirty());
            redoBtn = AddButton(bar, "作り直す", (o, e) => Redo(current));
            todayBtn = AddButton(bar, "今日の分を今すぐ作る", (o, e) => Redo(DateTime.Today.ToString("yyyy-MM-dd")));
            AddButton(bar, "サンプル", (o, e) => ShowSample());
            dutyBtn = AddButton(bar, "退勤", (o, e) => app.ToggleDuty());
            linkBtn = AddButton(bar, "", (o, e) => OpenLink());
            UpdateLink();

            var strip = new StatusStrip { SizingGrip = false };
            message.Spring = true;
            message.TextAlign = ContentAlignment.MiddleLeft;
            record.ForeColor = Color.DimGray;
            strip.Items.Add(message);
            strip.Items.Add(record);

            var menu = new MenuStrip();
            ToolStripMenuItem file = TopMenu(menu, "ファイル(&F)");
            mSave = Item(file, "保存(&S)", SaveIfDirty, Keys.Control | Keys.S);
            Item(file, "保存フォルダを開く(&O)", () => Process.Start("explorer.exe", "\"" + Core.Dir + "\""));
            file.DropDownItems.Add(new ToolStripSeparator());
            mDuty = Item(file, "退勤する(&Q)", app.ToggleDuty);
            file.DropDownItems.Add(new ToolStripSeparator());
            Item(file, "閉じる（トレイに戻す）(&C)", Close);
            Item(file, "終了（記録も止める）(&X)", app.Quit);
            ToolStripMenuItem report = TopMenu(menu, "日報(&R)");
            mCopyBullets = Item(report, "箇条書きをコピー(&B)", CopyBullets);
            mCopyAll = Item(report, "全文をコピー(&A)", () => Copy(Reporter.Unnumber(Body()), "全文をコピーしました"));
            report.DropDownItems.Add(new ToolStripSeparator());
            mAiFix = Item(report, "AIで直す（右の欄へ）(&E)", () => instruction.Focus(), Keys.Control | Keys.E);
            mRedo = Item(report, "この日を作り直す(&R)", () => Redo(current));
            mToday = Item(report, "今日の分を今すぐ作る(&T)", () => Redo(DateTime.Today.ToString("yyyy-MM-dd")));
            report.DropDownItems.Add(new ToolStripSeparator());
            Item(report, "サンプルを表示(&P)", ShowSample);
            ToolStripMenuItem tools = TopMenu(menu, "ツール(&T)");
            Item(tools, "設定(&S)…", app.ShowSettingsForm);
            Item(tools, "AIの接続先（Claude Code）(&A)…", app.ShowAiForm);
            tools.DropDownItems.Add(new ToolStripSeparator());
            mError = Item(tools, "エラーの記録を開く(&E)", () => Process.Start("notepad.exe", "\"" + Path.Combine(Core.Dir, "error.log") + "\""));
            tools.DropDownOpening += (o, e) => mError.Enabled = File.Exists(Path.Combine(Core.Dir, "error.log"));
            ToolStripMenuItem help = TopMenu(menu, "ヘルプ(&H)");
            Item(help, "使い方(&H)", () => { try { Process.Start(Updater.Help); } catch { } });
            Item(help, "更新を確認(&U)…", () => app.CheckUpdate(true));
            Item(help, "リリースのページを開く(&R)", () => Process.Start(Updater.Page));
            help.DropDownItems.Add(new ToolStripSeparator());
            Item(help, "バージョン情報(&A)", ShowAbout);
            MainMenuStrip = menu;
            // Dock は後から追加したものが外側になる
            var center = new Panel { Dock = DockStyle.Fill };
            center.Controls.Add(text);
            center.Controls.Add(BuildAddBar());
            center.Controls.Add(timeline);
            Controls.Add(center);
            Controls.Add(new Splitter { Dock = DockStyle.Right, Width = 5, MinExtra = 300, MinSize = W });   // 右の欄の幅をドラッグで変える
            Controls.Add(BuildRight());
            Controls.Add(list);
            Controls.Add(bar);
            Controls.Add(strip);
            Controls.Add(menu);

            FormClosing += (o, e) =>
            {
                if (e.CloseReason != CloseReason.UserClosing) { SaveIfDirty(); return; }
                e.Cancel = true;   // ×はトレイに戻すだけ
                SaveIfDirty();
                Hide();
            };
        }

        const int W = 300;
        static readonly Color PanelBack = Color.FromArgb(244, 247, 252);

        // 右側: タブ「AIで直す」「離席メモ」
        Control BuildRight()
        {
            var tabs = new TabControl { Dock = DockStyle.Right, Width = W + 22 };
            aiPage.Controls.Add(BuildAiPanel());
            afkPage.Controls.Add(BuildAfkPanel());
            reqPage.Controls.Add(BuildRequestPanel());
            taskPage.Controls.Add(BuildTaskPanel());
            aiPage.BackColor = afkPage.BackColor = taskPage.BackColor = reqPage.BackColor = PanelBack;
            tabs.TabPages.Add(aiPage);
            tabs.TabPages.Add(afkPage);
            tabs.TabPages.Add(taskPage);
            tabs.TabPages.Add(reqPage);
            tabs.SelectedIndexChanged += (o, e) => { if (tabs.SelectedTab == taskPage) UpdateTaskInfo(); };
            return tabs;
        }

        // 右の欄のタブ「タスク」: タスクとタイトルに出る語。今日当てはまった時間と、最近当てはまった日時を出す
        Control BuildTaskPanel()
        {
            var p = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, Padding = new Padding(4, 2, 2, 2), BackColor = PanelBack };
            p.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            var title = new Label { Text = "タスク", AutoSize = true, Font = new Font(Font, FontStyle.Bold), Margin = new Padding(3, 6, 3, 2) };
            var titlesBtn = new Button { Text = "タイトル一覧でタスクを付ける…", AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Padding = new Padding(6, 2, 6, 2), Margin = new Padding(3, 2, 3, 4) };
            titlesBtn.Click += (o, e) =>
            {
                using (var f = new TaskTitleForm(Icon)) f.ShowDialog(this);
                LoadTasks();
            };
            taskGrid.Dock = DockStyle.Fill;
            taskGrid.BackgroundColor = SystemColors.Window;
            taskGrid.RowHeadersWidth = 24;
            taskGrid.AllowUserToResizeRows = false;
            taskGrid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            taskGrid.Columns.Add(GridSort.NoColumn());
            taskGrid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "タスク", FillWeight = 40, MinimumWidth = 60 });
            taskGrid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "タイトルに出る語", FillWeight = 45, MinimumWidth = 60, ToolTipText = "カンマ区切り" });
            taskGrid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "今日", ReadOnly = true, AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells });
            taskGrid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "最近", ReadOnly = true, AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells });
            taskGrid.Columns.Add(new DataGridViewCheckBoxColumn { HeaderText = "完了", AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader });
            taskGrid.Columns[3].DefaultCellStyle.ForeColor = taskGrid.Columns[4].DefaultCellStyle.ForeColor = Color.DimGray;
            GridSort.Enable(taskGrid);
            taskGrid.CurrentCellDirtyStateChanged += (o, e) => { if (taskGrid.IsCurrentCellDirty && taskGrid.CurrentCell is DataGridViewCheckBoxCell) taskGrid.CommitEdit(DataGridViewDataErrorContexts.Commit); };
            taskGrid.CellValueChanged += (o, e) => { if (e.ColumnIndex == 1 || e.ColumnIndex == 2 || e.ColumnIndex == 5) SaveTasks(); };
            taskGrid.UserDeletedRow += (o, e) => SaveTasks();
            taskInfo.MaximumSize = new Size(0, 0);
            Control[] rows = { title, titlesBtn, taskGrid, taskInfo };
            foreach (Control c in rows)
            {
                p.RowStyles.Add(c == taskGrid ? new RowStyle(SizeType.Percent, 100) : new RowStyle(SizeType.AutoSize));
                p.Controls.Add(c, 0, p.RowCount++);
            }
            LoadTasks();
            return p;
        }

        void LoadTasks()
        {
            taskLoading = true;
            taskGrid.Rows.Clear();
            int no = 0;
            foreach (TaskItem t in Tasks.Load())
            {
                int i = taskGrid.Rows.Add(null, t.Name, t.Keywords, "", "", t.Done);
                GridSort.Set(taskGrid.Rows[i].Cells[0], ++no, no);
            }
            taskLoading = false;
            UpdateTaskInfo();
        }

        void SaveTasks()
        {
            if (taskLoading) return;
            // 並べ替えていても、保存は No の順（上にあるタスクが先）。足した行は最後
            var list = new List<KeyValuePair<int, TaskItem>>();
            int max = taskGrid.Rows.Cast<DataGridViewRow>().Select(r => r.Cells[0].Tag is int ? (int)r.Cells[0].Tag : 0).DefaultIfEmpty(0).Max();
            foreach (DataGridViewRow r in taskGrid.Rows)
            {
                if (r.IsNewRow) continue;
                string name = Convert.ToString(r.Cells[1].Value).Trim(), words = Convert.ToString(r.Cells[2].Value).Trim();
                if (name.Length == 0 && words.Length == 0) continue;
                if (name.Length == 0) name = words.Split(',', '、')[0].Trim();   // 名前が空なら最初の語を名前に
                if (!(r.Cells[0].Tag is int))
                {
                    taskLoading = true;
                    GridSort.Set(r.Cells[0], ++max, max);
                    taskLoading = false;
                }
                list.Add(new KeyValuePair<int, TaskItem>((int)r.Cells[0].Tag, new TaskItem { Name = name, Keywords = words, Done = r.Cells[5].Value is bool && (bool)r.Cells[5].Value }));
            }
            try { Tasks.Save(list.OrderBy(x => x.Key).Select(x => x.Value).ToList()); }
            catch (Exception ex) { taskInfo.ForeColor = Color.Firebrick; taskInfo.Text = "保存できません: " + ex.Message; return; }
            UpdateTaskInfo();
        }

        // 今日当てはまった時間（列「今日」）と、最近14日で最後に当てはまった日時（列「最近」mm/dd HH:mm）
        void UpdateTaskInfo()
        {
            taskInfo.ForeColor = Color.DimGray;
            try
            {
                Settings s = Settings.Get();
                List<TaskItem> tasks = Tasks.Load();
                Dictionary<string, Tasks.TitleTask> map = Tasks.LoadMap();
                var today = new Dictionary<string, int>();
                var last = new Dictionary<string, DateTime>();
                int none = 0;
                for (int i = 0; i < 14; i++)
                {
                    List<Row> rows = Reporter.ReadActive(DateTime.Today.AddDays(-i).ToString("yyyy-MM-dd"), s);
                    if (rows == null) continue;
                    foreach (Row r in rows)
                    {
                        string t = Tasks.Resolve(map, tasks, r.Proc, r.Title);
                        if (t == null) { if (i == 0) none++; continue; }
                        DateTime d;
                        if (!last.TryGetValue(t, out d) || r.Time > d) last[t] = r.Time;
                        if (i == 0) { int n; today.TryGetValue(t, out n); today[t] = n + 1; }
                    }
                }
                taskLoading = true;
                foreach (DataGridViewRow row in taskGrid.Rows)
                {
                    if (row.IsNewRow) continue;
                    string name = Convert.ToString(row.Cells[1].Value).Trim();
                    int n;
                    DateTime d;
                    if (today.TryGetValue(name, out n)) GridSort.Set(row.Cells[3], DayStatus.Span(Reporter.Minutes(n)), n); else GridSort.Set(row.Cells[3], "", null);
                    if (last.TryGetValue(name, out d)) GridSort.Set(row.Cells[4], d.ToString("MM/dd HH:mm"), d); else GridSort.Set(row.Cells[4], "", null);
                }
                taskLoading = false;
                taskInfo.Text = "今日 タスク外: " + DayStatus.Span(Reporter.Minutes(none));
            }
            catch (Exception ex) { taskLoading = false; taskInfo.Text = "集計できません: " + ex.Message; }
        }

        // 右の欄のタブ「要望」: 名前・種類・内容を書いて、GitHub の Issue の作成画面を開く
        Control BuildRequestPanel()
        {
            var p = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, Padding = new Padding(4, 2, 2, 2), BackColor = PanelBack };
            p.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            Func<string, Label> lbl = t => new Label { Text = t, AutoSize = true, MaximumSize = new Size(W - 16, 0), Margin = new Padding(3, 8, 3, 2) };
            var title = new Label { Text = "要望・不具合を送る", AutoSize = true, Font = new Font(Font, FontStyle.Bold), Margin = new Padding(3, 6, 3, 2) };
            var hint = new Label
            {
                AutoSize = true, MaximumSize = new Size(W - 16, 0), ForeColor = Color.DimGray, Margin = new Padding(3, 0, 3, 2),
                Text = "GitHub の Issues に送る（GitHub のアカウントが必要・内容は公開）"
            };
            foreach (string k in new[] { "要望", "不具合", "質問", "その他" }) reqKind.Items.Add(k);
            reqKind.SelectedIndex = 0;
            reqName.Text = Settings.Get().UserName;
            reqText.TextChanged += (o, e) => reqSend.Enabled = reqText.Text.Trim().Length > 0;
            reqSend.Enabled = false;
            reqSend.Font = new Font(Font, FontStyle.Bold);
            reqSend.Click += (o, e) => SendRequest();
            reqStatus.MaximumSize = new Size(W - 16, 0);
            Control[] rows = { title, hint, lbl("名前"), reqName, lbl("種類"), reqKind, lbl("内容"), reqText, reqEnv, reqSend, reqStatus };
            foreach (Control c in rows)
            {
                p.RowStyles.Add(c == reqText ? new RowStyle(SizeType.Percent, 100) : new RowStyle(SizeType.AutoSize));
                p.Controls.Add(c, 0, p.RowCount++);
            }
            return p;
        }

        void SendRequest()
        {
            string body = reqText.Text.Trim(), name = reqName.Text.Trim(), kind = reqKind.Text;
            if (body.Length == 0) { reqStatus.Text = "内容を書いてください"; reqText.Focus(); return; }
            try { Settings.SetUserName(name); } catch (Exception ex) { Core.Error("要望", "名前を保存できません: " + ex.Message); }
            try { Feedback.SaveCopy(kind, name, body, reqEnv.Checked); } catch (Exception ex) { Core.Error("要望", "写しを保存できません: " + ex.Message); }
            bool truncated;
            string url = Feedback.Url(kind, name, body, reqEnv.Checked, out truncated);
            if (truncated) { try { Clipboard.SetText(body); } catch { } }
            try { Process.Start(url); }
            catch (Exception ex) { reqStatus.ForeColor = Color.Firebrick; reqStatus.Text = "ブラウザを開けませんでした: " + ex.Message; return; }
            reqStatus.ForeColor = Color.DimGray;
            reqStatus.Text = "作成画面を開きました" + (truncated ? "（全文はクリップボード）" : "");
            reqText.Text = "";
        }

        // 本文の下: 開始・終了の時刻と内容を入れて「追記」。開始だけなら軽作業。右端はその日の休憩
        Control BuildAddBar()
        {
            var p = new TableLayoutPanel { Dock = DockStyle.Bottom, AutoSize = true, ColumnCount = 2, BackColor = PanelBack, Padding = new Padding(4, 4, 4, 2) };
            p.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            p.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            var row = new FlowLayoutPanel { AutoSize = true, Dock = DockStyle.Fill, WrapContents = true, Margin = new Padding(0) };
            Func<string, Label> lbl = t => new Label { Text = t, AutoSize = true, Margin = new Padding(3, 7, 2, 0) };
            Func<TimeBox, Button> now = box =>
            {
                var b = new Button { Text = "今", AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Padding = new Padding(4, 0, 4, 0), Margin = new Padding(1, 2, 8, 2) };
                b.Click += (o, e) => { box.SetNow(); box.Focus(); box.SelectionStart = box.TextLength; };
                tip.SetToolTip(b, "今の時刻を入れる");
                return b;
            };
            Label head = lbl("手で追記");
            head.Font = new Font(Font, FontStyle.Bold);
            head.Margin = new Padding(3, 7, 10, 0);
            row.Controls.Add(head);
            row.Controls.Add(lbl("開始"));
            row.Controls.Add(startBox);
            row.Controls.Add(now(startBox));
            row.Controls.Add(lbl("終了"));
            row.Controls.Add(endBox);
            row.Controls.Add(now(endBox));
            Label hint = lbl("開始だけなら軽作業");
            hint.ForeColor = Color.DimGray;
            row.Controls.Add(hint);
            string boxTip = "4桁で入力（1430 → 14:30）\nホイール・↑↓: 時の上で1時間、分の上で1分\nShift+ホイール: 1時間";
            tip.SetToolTip(startBox, boxTip);
            tip.SetToolTip(endBox, boxTip);
            tip.SetToolTip(breakCheck, "印を外すと、この日は休憩なし（休憩の時間の操作も日報に入れ、休憩の行を消す）");
            addText.KeyDown += (o, e) => { if (e.KeyCode == Keys.Enter) { e.Handled = e.SuppressKeyPress = true; AddManual(); } };
            addText.HandleCreated += (o, e) => SendMessage(addText.Handle, 0x1501, (IntPtr)1, "作業の内容（例: 〇〇さんと打合せ）を書いて Enter か「追記」");   // EM_SETCUEBANNER
            addBtn.Font = new Font(Font, FontStyle.Bold);
            addBtn.Click += (o, e) => AddManual();
            breakCheck.CheckedChanged += (o, e) => { if (!loading) BreakToggled(); };
            p.Controls.Add(row, 0, 0);
            p.Controls.Add(breakCheck, 1, 0);
            p.Controls.Add(addText, 0, 1);
            p.Controls.Add(addBtn, 1, 1);
            return p;
        }

        void AddManual()
        {
            if (current == null || sample) return;
            if (Previewing) { message.Text = "AI の案を「確定して保存」か「元に戻す」をしてから追記してください"; return; }
            if (!startBox.Valid || !endBox.Valid) { message.Text = "時刻は 4桁（例: 1430）か 14:30 の形で入れてください"; return; }
            TimeSpan? a = startBox.Value, b = endBox.Value;
            string what = addText.Text.Trim();
            if (a == null && b == null) { message.Text = "開始の時刻を入れてください（「今」で今の時刻が入ります）"; startBox.Focus(); return; }
            if (a != null && b != null && b <= a) { message.Text = "終了は開始より後にしてください"; endBox.Focus(); return; }
            if (what.Length == 0) { message.Text = "作業の内容を書いてください"; addText.Focus(); return; }
            SaveQuiet();   // 手で直していた本文を先に保存
            bool light;
            string item = Reporter.ManualItem(a, b, what, out light);
            bool inReport;
            try { inReport = Reporter.AddManual(current, item, light); }
            catch (Exception ex) { message.Text = "追記できませんでした: " + ex.Message; return; }
            startBox.Text = endBox.Text = addText.Text = "";
            ShowDate(current);
            if (inReport)
            {
                int i = Array.FindIndex(text.Lines, l => Reporter.Unnumber(l).Trim() == item);
                if (i >= 0) { text.Select(text.GetFirstCharIndexFromLine(i), text.Lines[i].Length); text.ScrollToCaret(); }
            }
            message.Text = inReport ? (light ? "軽作業" : "工程") + "に追記しました: " + item.Substring(1)
                                    : "日報がまだないので控えておき、日報を作るときに入れます: " + item.Substring(1);
        }

        void LoadBreak()
        {
            Settings s = Settings.Get();
            loading = true;
            breakCheck.Text = "休憩 " + s.BreakText;
            breakCheck.Checked = current != null && !sample && Reporter.BreakOn(current, s);
            loading = false;
        }

        // 休憩のチェックを付けた・外した: その日の休憩の行を足す・消す
        void BreakToggled()
        {
            if (current == null || sample) return;
            if (Previewing) { LoadBreak(); message.Text = "AI の案を「確定して保存」か「元に戻す」をしてから変えてください"; return; }
            bool on = breakCheck.Checked;
            SaveQuiet();
            Settings s = Settings.Get();
            try { Reporter.SetBreak(current, on, s); }
            catch (Exception ex) { LoadBreak(); message.Text = "休憩を変えられませんでした: " + ex.Message; return; }
            ShowDate(current);
            message.Text = on ? "休憩（" + s.BreakText + "）の行を入れました。この時間の操作は日報に入れません"
                              : "この日は休憩なしにしました。" + s.BreakText + " の操作も日報に入ります";
        }

        // 休憩の行を手で消した・AI に消してもらったら、その日は休憩なしにする（チェックも外す）
        void BreakRemoved(string before, string after)
        {
            Settings s = Settings.Get();
            if (current == null || !Reporter.BreakOn(current, s)) return;
            if (!Reporter.HasBreakLine(before, s) || Reporter.HasBreakLine(after, s)) return;
            try { Reporter.SetBreakFlag(current, false); }
            catch (Exception ex) { message.Text = "休憩なしにできませんでした: " + ex.Message; return; }
            LoadBreak();
            message.Text = "この日は休憩なし";
        }

        public void SetDuty(bool off)
        {
            timeline.Reload();
            dutyBtn.Text = off ? "出勤" : "退勤";
            mDuty.Text = off ? "出勤する（記録と AI を始める）(&Q)" : "退勤する（日報を仕上げて待機）(&Q)";
        }

        Control BuildAfkPanel()
        {
            var p = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, Padding = new Padding(4, 2, 2, 2), BackColor = PanelBack };
            p.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            afkHead.MaximumSize = new Size(W - 12, 0);
            afkGrid.Dock = DockStyle.Fill;
            afkGrid.AllowUserToAddRows = afkGrid.AllowUserToDeleteRows = afkGrid.AllowUserToResizeRows = false;
            afkGrid.RowHeadersVisible = false;
            afkGrid.BackgroundColor = SystemColors.Window;
            afkGrid.Columns.Add(GridSort.NoColumn());
            afkGrid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "時間", ReadOnly = true, AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells });
            afkGrid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "何をしていたか", AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill });
            afkGrid.Columns.Add(new DataGridViewCheckBoxColumn { HeaderText = "除く", AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader, ToolTipText = "休憩・私用など、日報に入れないとき" });
            afkGrid.CurrentCellDirtyStateChanged += (o, e) =>
            {
                if (afkGrid.IsCurrentCellDirty && afkGrid.CurrentCell is DataGridViewCheckBoxCell) afkGrid.CommitEdit(DataGridViewDataErrorContexts.Commit);
            };
            GridSort.Enable(afkGrid);
            afkGrid.CellValueChanged += (o, e) => { if (!loading && e.RowIndex >= 0 && afkGrid.Rows[e.RowIndex].Tag is AfkEntry) AfkEdited(e.RowIndex); };
            var freeHead = new Label
            {
                AutoSize = true, MaximumSize = new Size(W - 12, 0), Margin = new Padding(3, 10, 3, 2),
                Text = "時刻のわからないメモ（1行に1つ）"
            };
            freeMemo.TextChanged += (o, e) => { if (!loading) freeDirty = true; };
            p.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            p.Controls.Add(afkHead, 0, p.RowCount++);
            p.RowStyles.Add(new RowStyle(SizeType.Percent, 60));
            p.Controls.Add(afkGrid, 0, p.RowCount++);
            p.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            p.Controls.Add(freeHead, 0, p.RowCount++);
            p.RowStyles.Add(new RowStyle(SizeType.Percent, 40));
            p.Controls.Add(freeMemo, 0, p.RowCount++);
            return p;
        }

        void LoadAfk()
        {
            if (afkGrid.IsCurrentCellInEditMode) afkGrid.EndEdit();
            loading = true;
            afkGrid.Rows.Clear();
            try { afkList = current != null ? Afk.Load(current) : new List<AfkEntry>(); }
            catch (Exception ex) { afkList = new List<AfkEntry>(); message.Text = "離席メモを読めませんでした: " + ex.Message; }
            int no = 0;
            foreach (AfkEntry e in afkList)
            {
                int i = afkGrid.Rows.Add(null, e.Range, e.Memo, e.Skip);
                afkGrid.Rows[i].Tag = e;
                GridSort.Set(afkGrid.Rows[i].Cells[0], ++no, no);
                GridSort.Set(afkGrid.Rows[i].Cells[1], e.Range, e.Start);
                PaintAfkRow(i);
            }
            afkGrid.ClearSelection();
            try { freeMemo.Text = current != null ? Afk.ReadFree(current) : ""; }
            catch (Exception ex) { freeMemo.Text = ""; message.Text = "メモを読めませんでした: " + ex.Message; }
            loading = false;
            freeDirty = false;
            Settings s = Settings.Get();
            afkHead.Text = current == null ? "日付を選んでください" :
                afkList.Count == 0 ? "この日は、" + s.AfkMinutes + "分以上 PC から離れた時間はありません" :
                "PC から離れていた時間（" + s.AfkMinutes + "分以上）。何をしていたか書くと、その時間の作業として日報に入ります。休憩・私用は「除く」に印";
            UpdateAfkTab();
        }

        void PaintAfkRow(int i)
        {
            var e = (AfkEntry)afkGrid.Rows[i].Tag;
            afkGrid.Rows[i].DefaultCellStyle.BackColor = e.Answered ? SystemColors.Window : Color.FromArgb(255, 246, 214);
            afkGrid.Rows[i].Cells[1].ToolTipText = e.Duration;
        }

        void UpdateAfkTab()
        {
            int n = afkList.Count(e => !e.Answered);
            afkPage.Text = n > 0 ? "離席メモ（" + n + "）" : "離席メモ";
        }

        void AfkEdited(int i)
        {
            var e = (AfkEntry)afkGrid.Rows[i].Tag;
            e.Memo = Convert.ToString(afkGrid.Rows[i].Cells[2].Value).Trim();
            e.Skip = Convert.ToBoolean(afkGrid.Rows[i].Cells[3].Value);
            PaintAfkRow(i);
            UpdateAfkTab();
            // 保存して日報へ足す（表のイベントの中で表を作り直さないよう、あとで）
            BeginInvoke((Action)(() => app.AfkChanged(e)));
        }

        // 日報が追記された・離席メモが足された: 見ている日なら読み直す（AI の案を出しているときは読み直さない）
        public void ReportUpdated(string date)
        {
            if (!Visible || sample) return;   // 次に開いたときに読み直す
            if (Previewing)
            {
                if (date == current) message.Text = "日報に追記あり";
                return;
            }
            RefreshList(current);
            if (date == current) message.Text = "日報を更新しました（" + DateTime.Now.ToString("HH:mm") + "）";
        }

        // 離席が記録された・戻ったときの画面で書かれたとき
        public void RefreshAfk(string date)
        {
            if (date != current) return;
            SaveFreeIfDirty();
            LoadAfk();
        }

        void SaveFreeIfDirty()
        {
            if (afkGrid.IsCurrentCellInEditMode) afkGrid.EndEdit();
            if (!freeDirty || current == null) return;
            try { Afk.SaveFree(current, freeMemo.Text); freeDirty = false; }
            catch (Exception ex) { message.Text = "メモを保存できませんでした: " + ex.Message; }
        }

        // 右の欄「AIで直す」: AI と会話しながら真ん中の日報を直す。AI の案は本文に青字で出し、「確定して保存」か「元に戻す」
        Control BuildAiPanel()
        {
            var p = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, Padding = new Padding(4, 2, 2, 2), BackColor = PanelBack };
            p.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            aiTitle.Font = new Font(Font, FontStyle.Bold);
            aiStatus.MaximumSize = new Size(W - 16, 0);
            var hint = new Label
            {
                AutoSize = true, MaximumSize = new Size(W - 16, 0), ForeColor = Color.DimGray, Margin = new Padding(3, 0, 3, 4),
                Text = "日報の直し方を AI に頼む（行は番号で指定）"
            };
            var notesHead = new Label
            {
                AutoSize = true, MaximumSize = new Size(W - 16, 0), Margin = new Padding(3, 8, 3, 2), ForeColor = Color.DimGray,
                Text = "訂正の記録"
            };
            chatView.Font = new Font(Font.FontFamily, Font.Size * 0.95f);
            askBtn.Font = new Font(Font, FontStyle.Bold);
            askBtn.Click += (o, e) => AskAi();
            clearChatBtn.Click += (o, e) => ClearChat();
            acceptBtn.Click += (o, e) => Accept();
            revertBtn.Click += (o, e) => Revert();
            previewBar.Controls.Add(acceptBtn);
            previewBar.Controls.Add(revertBtn);
            var sendBar = new FlowLayoutPanel { AutoSize = true, WrapContents = false, Margin = new Padding(0) };
            sendBar.Controls.Add(askBtn);
            sendBar.Controls.Add(clearChatBtn);
            instruction.TextChanged += (o, e) => UpdateButtons();
            // Enter で送る。Shift+Enter で改行
            instruction.KeyDown += (o, e) => { if (e.KeyCode == Keys.Enter && !e.Shift) { e.Handled = e.SuppressKeyPress = true; AskAi(); } };
            instruction.HandleCreated += (o, e) => SendMessage(instruction.Handle, 0x1501, (IntPtr)1, "AI に頼むことを書いて Enter（Shift+Enter で改行）");   // EM_SETCUEBANNER（1行のときだけ出る）
            notes.TextChanged += (o, e) => { if (!loading) notesDirty = true; };

            Control[] rows = { aiTitle, hint, chatView, aiStatus, previewBar, instruction, sendBar, notesHead, notes };
            foreach (Control c in rows)
            {
                RowStyle st = c == chatView ? new RowStyle(SizeType.Percent, 100)
                            : c == instruction ? new RowStyle(SizeType.Absolute, 72)
                            : c == notes ? new RowStyle(SizeType.Absolute, 70) : new RowStyle(SizeType.AutoSize);
                p.RowStyles.Add(st);
                p.Controls.Add(c, 0, p.RowCount++);
            }
            return p;
        }

        // 会話を表示する（利用者は青、AI は黒。名前は太字）
        void ShowChat()
        {
            chatView.Clear();
            if (chat.Count == 0)
            {
                chatView.SelectionColor = Color.DimGray;
                chatView.AppendText("まだ会話はありません。下に頼むことを書いて送ってください。");
                return;
            }
            foreach (KeyValuePair<bool, string> m in chat) AppendChat(m.Key, m.Value);
        }

        void AppendChat(bool user, string text)
        {
            if (chatView.TextLength > 0 && chat.Count > 0) chatView.AppendText("\n");
            if (chat.Count == 1 && user && chatView.Text.StartsWith("まだ会話")) chatView.Clear();
            chatView.SelectionStart = chatView.TextLength;
            chatView.SelectionFont = new Font(chatView.Font, FontStyle.Bold);
            chatView.SelectionColor = user ? Color.FromArgb(30, 90, 170) : Color.FromArgb(40, 40, 40);
            chatView.AppendText((user ? "あなた" : "AI") + "\n");
            chatView.SelectionFont = chatView.Font;
            chatView.SelectionColor = user ? Color.FromArgb(30, 90, 170) : Color.FromArgb(40, 40, 40);
            chatView.AppendText(text.Trim().Replace("\r\n", "\n") + "\n");
            chatView.SelectionStart = chatView.TextLength;
            chatView.ScrollToCaret();
        }

        void LoadChat()
        {
            try { chat = current != null && !sample ? Reporter.LoadChat(current) : new List<KeyValuePair<bool, string>>(); }
            catch (Exception ex) { chat = new List<KeyValuePair<bool, string>>(); message.Text = "会話を読めませんでした: " + ex.Message; }
            ShowChat();
        }

        void ClearChat()
        {
            if (current == null || chat.Count == 0) return;
            if (MessageBox.Show(this, current + " の会話を消しますか？（日報と訂正の記録は消えません）", "AIで直す", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return;
            Reporter.ClearChat(current);
            chat.Clear();
            ShowChat();
        }

        static Button AddButton(FlowLayoutPanel bar, string label, EventHandler click)
        {
            var b = new Button { Text = label, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Padding = new Padding(8, 2, 8, 2), Margin = new Padding(3) };
            b.Click += click;
            bar.Controls.Add(b);
            return b;
        }

        // 下の日報リンクのボタン（設定の LinkName・LinkUrl）
        public void UpdateLink()
        {
            Settings s = Settings.Get();
            linkBtn.Text = s.LinkName.Trim().Length > 0 ? s.LinkName.Trim() : "日報リンク";
            linkBtn.Visible = s.LinkUrl.Trim().Length > 0;
        }

        // 日報リンクを開く: URL は既定のブラウザ、ファイル（Excel の日報など）・フォルダは関連付けられたアプリで
        void OpenLink()
        {
            string target = Reporter.LinkTarget(Settings.Get().LinkUrl, DateTime.Today);
            if (target.Length == 0) return;
            try { Process.Start(new ProcessStartInfo(target) { UseShellExecute = true }); }
            catch (Exception ex) { message.Text = "開けません: " + target + "（" + ex.Message + "）"; }
        }

        void ShowAbout()
        {
            using (var f = new AboutForm(Icon)) f.ShowDialog(this);
        }

        static ToolStripMenuItem TopMenu(MenuStrip menu, string label)        {
            var m = new ToolStripMenuItem(label);
            menu.Items.Add(m);
            return m;
        }

        static ToolStripMenuItem Item(ToolStripMenuItem parent, string label, Action click, Keys keys = Keys.None)
        {
            var m = new ToolStripMenuItem(label, null, (o, e) => click()) { ShortcutKeys = keys };
            parent.DropDownItems.Add(m);
            return m;
        }

        // RichTextBox は改行を \n にするので、保存・コピーのときは \r\n にそろえる
        string Body() { return Regex.Replace(text.Text, "\r?\n", "\r\n"); }

        static string Norm(string s) { return s.Replace("\r\n", "\n").TrimEnd(); }

        bool Previewing { get { return previewBase != null; } }

        void UpdateButtons()
        {
            bool has = current != null && hasReport;
            copyBullets.Enabled = copyAll.Enabled = has || sample;
            saveBtn.Enabled = has && (dirty || Previewing);
            redoBtn.Enabled = current != null && !busy && !Previewing;
            todayBtn.Enabled = !busy && !Previewing;
            mCopyBullets.Enabled = copyBullets.Enabled;
            mCopyAll.Enabled = copyAll.Enabled;
            mSave.Enabled = saveBtn.Enabled;
            mRedo.Enabled = redoBtn.Enabled;
            mAiFix.Enabled = has;
            mToday.Enabled = todayBtn.Enabled;
            // 本文はいつでも手で直せる（日報がない画面の案内文だけは直せない）
            text.ReadOnly = !has && !sample;
            instruction.Enabled = notes.Enabled = has;
            askBtn.Enabled = has && !aiRunning && instruction.Text.Trim().Length > 0;
            clearChatBtn.Enabled = has && !aiRunning && chat.Count > 0;
            previewBar.Visible = Previewing;
            Settings st = Settings.Get();
            aiTitle.Text = "AIで直す（" + st.FixModel + "）";
            bool day = current != null && !sample;
            startBox.Enabled = endBox.Enabled = addText.Enabled = day;
            addBtn.Enabled = day && !Previewing;
            breakCheck.Visible = st.HasBreak;
            breakCheck.Enabled = day && !Previewing;
        }

        public void ShowSample()
        {
            SaveIfDirty();
            current = null;
            hasReport = false;
            sample = true;
            SetText(Reporter.Number(SampleText));
            list.SelectedIndex = -1;
            dirty = false;
            LoadNotes();
            LoadChat();
            LoadAfk();
            LoadBreak();
            Text = "日報 サンプル" + (Core.Test ? "（テスト）" : "");
            if (!busy) message.Text = "見本";
            SetAiStatus("", false);
            UpdateButtons();
        }

        const string BusyText = "日報を作成中…";

        public void SetBusy(bool b)
        {
            busy = b;
            timeline.Reload();
            if (b) message.Text = BusyText;
            else if (message.Text == BusyText) message.Text = "";
            UpdateButtons();
        }

        // 記録のたび（1分ごと）に呼ばれる。今日を開いていれば帯も新しくする
        public void SetRecordStatus(string s)
        {
            record.Text = s;
            UpdateDatesIfStale();
            if (timeline.Date == DateTime.Today.ToString("yyyy-MM-dd")) timeline.Reload();
        }

        string listSig;   // 日付の一覧を作ったときの今日と日報のファイル（変わったら一覧を作り直す）

        static List<string> ReportDates()
        {
            return Directory.Exists(Core.Dir)
                ? Directory.GetFiles(Core.ReportDir, "日報_*.txt").Select(f => Path.GetFileNameWithoutExtension(f).Substring(3))
                    .Where(Reporter.IsDate).OrderByDescending(d => d).ToList()
                : new List<string>();
        }

        static string Sig(List<string> reports) { return DateTime.Today.ToString("yyyy-MM-dd") + "|" + string.Join(",", reports); }

        // 日付が変わった・日報が増えたら、日付の一覧だけを作り直す（選んでいる日と本文はそのまま）。記録のたび（1分ごと）に呼ぶ
        void UpdateDatesIfStale()
        {
            List<string> reports;
            try { reports = ReportDates(); } catch { return; }
            if (Sig(reports) == listSig || list.Items.Count == 0) return;
            string today = DateTime.Today.ToString("yyyy-MM-dd");
            List<string> dates = reports.Concat(new[] { today }).Distinct().OrderByDescending(d => d).ToList();
            loading = true;
            list.BeginUpdate();
            list.Items.Clear();
            foreach (string d in dates) list.Items.Add(new DateItem { Date = d, HasReport = reports.Contains(d) });
            list.SelectedIndex = current != null ? dates.IndexOf(current) : -1;
            list.EndUpdate();
            loading = false;
            listSig = Sig(reports);
        }

        public void RefreshList(string select)
        {
            SaveIfDirty();
            List<string> reports = ReportDates();
            listSig = Sig(reports);
            // 今日は日報がまだなくても出す（離席メモを書けるように）
            string today = DateTime.Today.ToString("yyyy-MM-dd");
            List<string> dates = reports.Concat(new[] { today }).Distinct().OrderByDescending(d => d).ToList();
            string want = select ?? current ?? reports.OrderByDescending(d => d).FirstOrDefault() ?? today;
            if (!dates.Contains(want)) want = today;

            loading = true;
            list.BeginUpdate();
            list.Items.Clear();
            foreach (string d in dates) list.Items.Add(new DateItem { Date = d, HasReport = reports.Contains(d) });
            list.SelectedIndex = want == null ? -1 : dates.IndexOf(want);
            list.EndUpdate();
            loading = false;

            if (want != null) ShowDate(want);
            else
            {
                current = null;
                sample = false;
                SetText("まだ日報がありません。\r\n\r\n日報は毎日 " + Settings.Get().ReportTime.ToString(@"hh\:mm") +
                        " を過ぎると自動で作られます。\r\nすぐ作るときは「今日の分を今すぐ作る」を押してください。\r\n" +
                        "できあがりの見本は「サンプル」で見られます。");
                dirty = false;
                LoadNotes();
                LoadChat();
                Text = "日報" + (Core.Test ? "（テスト）" : "");
                UpdateButtons();
            }
        }

        void ShowDate(string date)
        {
            SaveIfDirty();
            current = date;
            timeline.SetDate(date);
            sample = false;
            string path = Core.ReportPath(date);
            hasReport = File.Exists(path);
            try
            {
                SetText(hasReport ? Reporter.Number(Core.ReadText(path)) :
                    "この日の日報はまだありません。\r\n\r\n今日の分は毎日 " + Settings.Get().ReportTime.ToString(@"hh\:mm") +
                    " を過ぎると自動で作られます。すぐ作るときは「作り直す」か「今日の分を今すぐ作る」を押してください。\r\n" +
                    "PC から離れていたあいだのことは、右の「離席メモ」に書いておくと日報に入ります。");
                loadedStamp = hasReport ? File.GetLastWriteTime(path) : DateTime.MinValue;
            }
            catch (Exception ex) { SetText(""); message.Text = "読めませんでした: " + ex.Message; }
            dirty = false;
            LoadNotes();
            LoadChat();
            LoadAfk();
            LoadBreak();
            Text = "日報 " + date + (Core.Test ? "（テスト）" : "");
            if (!busy) message.Text = "";
            if (!aiRunning) SetAiStatus("", false);
            UpdateButtons();
        }

        // 本文を差し替える（色は黒に戻す）
        void SetText(string s)
        {
            loading = true;
            text.Text = s;
            text.SelectAll();
            text.SelectionColor = text.ForeColor;
            text.Select(0, 0);
            loading = false;
        }

        void LoadNotes()
        {
            loading = true;
            try { notes.Text = current != null ? Reporter.ReadNotes(current) : ""; }
            catch (Exception ex) { notes.Text = ""; message.Text = "訂正の記録を読めませんでした: " + ex.Message; }
            loading = false;
            notesDirty = false;
        }

        void SaveNotesIfDirty()
        {
            if (!notesDirty || current == null) return;
            try { Reporter.SaveNotes(current, notes.Text); notesDirty = false; }
            catch (Exception ex) { message.Text = "訂正の記録を保存できませんでした: " + ex.Message; }
        }

        // 手で直した本文を保存する。AI の案を出しているときは、確定するか元に戻すかを聞く
        public void SaveIfDirty()
        {
            if (Previewing)
            {
                DialogResult r = MessageBox.Show(this, "AI の案を保存しますか？\n\nはい: 確定して保存\nいいえ: 元に戻す", "日報",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (r == DialogResult.Yes && Accept()) return;
                Revert();
            }
            SaveQuiet();
        }

        // 確認を出さずに保存する（自動の日報作成の前など）。AI の案は確定していないので保存しない
        public void SaveQuiet()
        {
            if (current == null) return;
            SaveFreeIfDirty();
            if (Previewing) return;
            SaveNotesIfDirty();
            if (!dirty) return;
            RenumberView();
            string path = Core.ReportPath(current);
            try
            {
                if (File.Exists(path) && File.GetLastWriteTime(path) != loadedStamp)
                {
                    // 開いたあとに作り直されていたら上書きしない。手で直した内容は old に残す
                    string dir = Path.Combine(Core.Dir, "old");
                    Directory.CreateDirectory(dir);
                    Core.WriteText(Path.Combine(dir, "日報_" + current + "_手直し_" + DateTime.Now.ToString("yyyyMMdd-HHmmss") + ".txt"), Body());
                    message.Text = "日報が別に更新されたため、手で直した内容は old フォルダに保存";
                }
                else
                {
                    string before = File.Exists(path) ? Core.ReadText(path) : "";
                    Core.WriteText(path, Body());
                    loadedStamp = File.GetLastWriteTime(path);
                    message.Text = "保存しました（" + DateTime.Now.ToString("HH:mm") + "）";
                    BreakRemoved(before, Body());
                }
                dirty = false;
            }
            catch (Exception ex) { message.Text = "保存できませんでした: " + ex.Message; }
            UpdateButtons();
        }

        // 手で行を足したり消したりしたときに、番号を振り直す（カーソルの位置はそのまま）
        void RenumberView()
        {
            string numbered = Reporter.Number(Body());
            if (Norm(numbered) == Norm(Body())) return;
            int caret = text.SelectionStart;
            SetText(numbered);
            text.Select(Math.Min(caret, text.TextLength), 0);
        }

        void CopyBullets()
        {
            string[] lines = Reporter.Unnumber(Body()).Replace("\r\n", "\n").Split('\n');
            // 「■工程」「■軽作業」の見出しと「・」の行
            List<string> bullets = lines.Select(l => l.Trim()).Where(l => l.StartsWith("・") || l.StartsWith("■")).ToList();
            int count = bullets.Count(l => l.StartsWith("・"));
            if (count > 0) { Copy(string.Join("\r\n", bullets), count + "行の箇条書きをコピーしました"); return; }
            // 箇条書きがない（AI 失敗など）ときは見出し行を除いた本文
            string body = string.Join("\r\n", lines.Skip(1).SkipWhile(l => l.Trim().Length == 0));
            Copy(body, "本文をコピーしました");
        }

        void Copy(string s, string done)
        {
            if (string.IsNullOrWhiteSpace(s)) { message.Text = "コピーするものがありません"; return; }
            try
            {
                Clipboard.SetDataObject(s, true, 5, 100);
                message.Text = done + "（" + DateTime.Now.ToString("HH:mm:ss") + "）";
            }
            catch (Exception ex) { message.Text = "コピーできませんでした: " + ex.Message; }
        }

        void SetAiStatus(string s, bool error)
        {
            aiStatus.ForeColor = error ? Color.Firebrick : SystemColors.ControlText;
            aiStatus.Text = s;
        }

        // 右の欄の会話で、AI に今の本文（真ん中の日報）を直してもらう。案が出ているときは案をもとに直す
        void AskAi()
        {
            string inst = instruction.Text.Trim();
            if (current == null || !hasReport || aiRunning || inst.Length == 0) return;
            if (!Previewing) SaveQuiet();   // 手で直した分は先に保存しておく
            string date = current, sent = Body(), noteText = notes.Text;
            var history = new List<KeyValuePair<bool, string>>(chat);
            chat.Add(new KeyValuePair<bool, string>(true, inst));
            Reporter.AddChat(date, true, inst);
            AppendChat(true, inst);
            instruction.Text = "";
            aiRunning = true;
            UpdateButtons();
            SetAiStatus("AI が応答中…", false);
            var t = new Thread(() =>
            {
                string result = null, reply = null, error = null;
                try { result = Reporter.Chat(date, sent, history, inst, noteText, Settings.Get(), out reply); }
                catch (NeedLoginException ex) { error = ex.Message + "（ツール → AIの接続先）"; }
                catch (Exception ex) { error = "AI エラー: " + ex.Message; Core.Error("AIで直す " + date, ex.Message); }
                if (IsDisposed) return;
                BeginInvoke((Action)(() => AiDone(date, sent, inst, result, reply, error)));
            });
            t.IsBackground = true;
            t.Start();
        }

        void AiDone(string date, string sent, string inst, string result, string reply, string error)
        {
            aiRunning = false;
            if (error != null) { SetAiStatus(error, true); UpdateButtons(); return; }
            Reporter.AddChat(date, false, reply);
            if (date != current) { SetAiStatus(date + " の案（日付が変わったため使いません）", true); UpdateButtons(); return; }
            chat.Add(new KeyValuePair<bool, string>(false, reply));
            AppendChat(false, reply);
            if (result == null) { SetAiStatus("", false); UpdateButtons(); return; }   // 返事・質問だけ
            if (Norm(Body()) != Norm(sent) &&
                MessageBox.Show(this, "本文が変わっています。AI の案に置き換えますか？", "日報",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
            {
                SetAiStatus("案は使いません", false);
                UpdateButtons();
                return;
            }
            if (!Previewing)
            {
                previewBase = Body();
                previewNotes = notes.Text;
                previewDirty = dirty;
            }
            SetText(result);
            ColorChanges();
            dirty = true;
            // 頼んだことを訂正の記録に足す（確定すると保存。元に戻すと消える）
            string line = "・" + Reporter.ResolveNumbers(inst, sent).Replace("\r\n", " ").Replace("\n", " ");
            if (!notes.Text.Contains(line)) notes.Text = (notes.Text.TrimEnd() + "\r\n" + line).Trim();
            SetAiStatus(Norm(result) == Norm(previewBase) ? "変更なし" : "青字: AI が直した行", false);
            UpdateButtons();
        }

        // AI の案を出す前の本文にない行を青字にする
        void ColorChanges()
        {
            // 番号は振り直されるので、番号を外して比べる
            var old = new HashSet<string>(Reporter.Unnumber(previewBase).Replace("\r\n", "\n").Split('\n').Select(l => l.Trim()));
            loading = true;
            int i = 0;
            foreach (string l in text.Lines)
            {
                if (l.Trim().Length > 0 && !old.Contains(Reporter.Unnumber(l).Trim()))
                {
                    text.Select(text.GetFirstCharIndexFromLine(i), l.Length);
                    text.SelectionColor = AiBlue;
                }
                i++;
            }
            text.Select(0, 0);
            loading = false;
        }

        bool Accept()
        {
            if (!Previewing || current == null) return false;
            string path = Core.ReportPath(current);
            if (File.Exists(path) && File.GetLastWriteTime(path) != loadedStamp &&
                MessageBox.Show(this, "日報が別に更新されています。置き換えますか？", "日報",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes) return false;
            try
            {
                RenumberView();
                Reporter.SaveRevision(current, Body(), notes.Text);
                loadedStamp = File.GetLastWriteTime(path);
                string before = previewBase;
                previewBase = null;
                dirty = notesDirty = false;
                SetText(Body());
                SetAiStatus("保存（前の版は old フォルダ）", false);
                message.Text = "保存（" + DateTime.Now.ToString("HH:mm") + "）";
                BreakRemoved(before, Body());
                UpdateButtons();
                return true;
            }
            catch (Exception ex) { SetAiStatus("保存できません: " + ex.Message, true); return false; }
        }

        void Revert()
        {
            if (!Previewing) return;
            SetText(previewBase);
            loading = true;
            notes.Text = previewNotes;
            loading = false;
            dirty = previewDirty;
            previewBase = null;
            SetAiStatus("元に戻しました", false);
            UpdateButtons();
        }

        void Redo(string date)
        {
            if (date == null || app.Busy) return;
            // まだ日報の形になっていない日（AI に失敗した日など）は、合成せずに作り直すだけなので聞かない
            if (File.Exists(Core.ReportPath(date)) && Core.ReadText(Core.ReportPath(date)).Contains("■") &&
                MessageBox.Show(this, date + " の日報を作り直しますか？", "日報",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return;
            SaveIfDirty();
            app.Generate(date);
        }
    }

    // 離席から戻ったときに右下に出す小さな画面。前には出るが、入力中の作業からフォーカスは奪わない
    class AfkAskForm : Form
    {
        static readonly string[] Quick = { "会議・打合せ", "現場", "電話", "来客対応" };

        readonly Action<AfkEntry> answered;
        readonly Label head = new Label { AutoSize = true, MaximumSize = new Size(390, 0), Margin = new Padding(3, 4, 3, 6) };
        readonly TextBox memo = new TextBox { Dock = DockStyle.Fill, Margin = new Padding(3, 4, 3, 2) };
        readonly Label note = new Label { AutoSize = true, MaximumSize = new Size(390, 0), Margin = new Padding(3, 2, 3, 4) };
        AfkEntry entry;

        protected override bool ShowWithoutActivation { get { return true; } }

        protected override CreateParams CreateParams
        {
            get { CreateParams p = base.CreateParams; p.ExStyle |= 0x8; return p; }   // WS_EX_TOPMOST（TopMost だと前に出るときにフォーカスを取る）
        }

        public AfkAskForm(Icon icon, Action<AfkEntry> answered)
        {
            this.answered = answered;
            Text = "離席のあいだ" + (Core.Test ? "（テスト）" : "");
            Icon = icon;
            AutoScaleDimensions = new SizeF(96f, 96f);
            AutoScaleMode = AutoScaleMode.Dpi;
            Font = new Font("Meiryo UI", 10f);
            FormBorderStyle = FormBorderStyle.FixedToolWindow;
            MaximizeBox = MinimizeBox = false;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.Manual;
            ClientSize = new Size(410, 228);

            var quick = new FlowLayoutPanel { AutoSize = true, Dock = DockStyle.Fill, Margin = new Padding(0) };
            foreach (string q in Quick)
            {
                string word = q;
                var b = new Button { Text = word, AutoSize = true, Margin = new Padding(3, 2, 3, 2) };
                b.Click += (o, e) =>
                {
                    memo.Text = memo.Text.Trim().Length == 0 ? word : memo.Text.Trim() + "・" + word;
                    memo.Focus();
                    memo.SelectionStart = memo.Text.Length;
                };
                quick.Controls.Add(b);
            }

            // 右から並ぶ: [保存] [休憩・私用] [あとで書く]
            var bar = new FlowLayoutPanel { AutoSize = true, Dock = DockStyle.Fill, FlowDirection = FlowDirection.RightToLeft, WrapContents = false, Margin = new Padding(0, 4, 0, 0) };
            Button later = BarButton(bar, "あとで書く", (o, e) => Hide());
            BarButton(bar, "休憩・私用（日報に入れない）", (o, e) => Save(true));
            Button save = BarButton(bar, "保存", (o, e) => Save(false));
            save.Font = new Font(Font, FontStyle.Bold);
            AcceptButton = save;
            CancelButton = later;

            var t = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, Padding = new Padding(8, 4, 8, 4) };
            foreach (Control c in new Control[] { head, quick, memo, note, bar })
            {
                t.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                t.Controls.Add(c, 0, t.RowCount++);
            }
            Controls.Add(t);

            // ×は隠すだけ（記録はそのまま。日報画面の「離席メモ」で書ける）
            FormClosing += (o, e) => { if (e.CloseReason == CloseReason.UserClosing) { e.Cancel = true; Hide(); } };
        }

        static Button BarButton(FlowLayoutPanel bar, string text, EventHandler click)
        {
            var b = new Button { Text = text, AutoSize = true, Padding = new Padding(6, 1, 6, 1), Margin = new Padding(3) };
            b.Click += click;
            bar.Controls.Add(b);
            return b;
        }

        public void SetEntry(AfkEntry e)
        {
            entry = e;
            head.Text = e.Range + "（" + e.Duration + "）PC から離れていました。\nそのあいだ何をしていましたか？";
            memo.Text = e.Memo;
            note.ForeColor = Color.DimGray;
            note.Text = "この時間の作業として日報に入ります";
            Rectangle wa = Screen.PrimaryScreen.WorkingArea;
            Location = new Point(wa.Right - Width - 12, wa.Bottom - Height - 12);
        }

        void Save(bool skip)
        {
            if (entry == null) { Hide(); return; }
            if (!skip && memo.Text.Trim().Length == 0)
            {
                note.ForeColor = Color.Firebrick;
                note.Text = "何をしていたか書くか、上のボタンを押してください";
                return;
            }
            entry.Memo = skip ? "" : memo.Text.Trim();
            entry.Skip = skip;
            AfkEntry e = entry;
            entry = null;
            Hide();
            if (answered != null) answered(e);   // 保存して、日報があればその時刻の行として足す
        }
    }

    // 設定画面: settings.ini のすべての項目と、Windows 起動時の自動起動
    // バージョン情報: 名前と版だけを真ん中に出す。変わったところは GitHub のリリース（CHANGELOG.md）で見る
    class AboutForm : Form
    {
        public AboutForm(Icon icon)
        {
            Text = "バージョン情報";
            Icon = icon;
            AutoScaleDimensions = new SizeF(96f, 96f);
            AutoScaleMode = AutoScaleMode.Dpi;
            Font = new Font("Meiryo UI", 10f);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.CenterParent;
            BackColor = SystemColors.Window;
            AutoSize = true;
            AutoSizeMode = AutoSizeMode.GrowAndShrink;
            Padding = new Padding(0, 0, 40, 20);

            var t = new TableLayoutPanel { ColumnCount = 1, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Location = new Point(40, 24) };
            t.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            int px;
            using (Graphics g = CreateGraphics()) px = (int)Math.Round(64 * g.DpiX / 96f);
            var logo = new PictureBox { Image = TrayApp.Logo(px, Color.FromArgb(30, 111, 217)), SizeMode = PictureBoxSizeMode.AutoSize, Margin = new Padding(3, 3, 3, 12) };
            var name = new Label { Text = "作業ログ・日報", AutoSize = true, Font = new Font(Font.FontFamily, 15f, FontStyle.Bold) };
            var ver = new Label { Text = "バージョン " + Core.Version + (Core.Test ? "（テスト）" : ""), AutoSize = true, Margin = new Padding(3, 4, 3, 14) };
            var ai = new Label { Text = "AI（Anthropic の Claude）が作成したソフトウェアです。無保証です。", AutoSize = true, ForeColor = Color.DimGray };
            var link = new LinkLabel { Text = "更新履歴（GitHub のリリース）", AutoSize = true, Margin = new Padding(3, 8, 3, 18) };
            link.LinkClicked += (o, e) => { try { Process.Start(Updater.Releases); } catch (Exception ex) { Core.Error("リンク", ex.Message); } };
            var ok = new Button { Text = "OK", DialogResult = DialogResult.OK, AutoSize = true, MinimumSize = new Size(96, 0) };
            foreach (Control c in new Control[] { logo, name, ver, ai, link, ok })
            {
                c.Anchor = AnchorStyles.None;   // 列の真ん中に置く
                t.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                t.Controls.Add(c, 0, t.RowCount++);
            }
            Controls.Add(t);
            AcceptButton = ok;
            CancelButton = ok;
        }
    }

    class SettingsForm : Form
    {
        static readonly string[] Models = { "claude-haiku-5-5", "claude-sonnet-5-5", "claude-opus-5-5" };

        readonly Action saved;
        readonly DateTimePicker reportTime = new DateTimePicker { Format = DateTimePickerFormat.Custom, CustomFormat = "HH:mm", ShowUpDown = true, Width = 90 };
        readonly NumericUpDown idle = Num(30, 3600), minMinutes = Num(0, 60), minDay = Num(0, 600), redo = Num(1, 600), catchUp = Num(1, 60), keep = Num(7, 3650);
        readonly CheckBox openWindow = new CheckBox { Text = "日報ができたら日報画面を開く", AutoSize = true, Margin = new Padding(3, 10, 3, 3) };
        readonly CheckBox startup = new CheckBox { Text = "Windows 起動時に自動で起動する", AutoSize = true, Margin = new Padding(3, 10, 3, 3) };
        readonly CheckBox autoUpdate = new CheckBox { Text = "新しい版が出たら自動で更新する", AutoSize = true, Margin = new Padding(3, 10, 3, 3) };
        readonly CheckBox afkAsk = new CheckBox { Text = "離席から戻ったときに何をしていたか聞く", AutoSize = true, Margin = new Padding(3, 10, 3, 3) };
        readonly NumericUpDown afkMinutes = Num(0, 240), appendMinutes = Num(0, 480), offIdle = Num(0, 600);
        readonly DateTimePicker workStart = new DateTimePicker { Format = DateTimePickerFormat.Custom, CustomFormat = "HH:mm", ShowUpDown = true, Width = 90 };
        readonly DateTimePicker planFrom = new DateTimePicker { Format = DateTimePickerFormat.Custom, CustomFormat = "HH:mm", ShowUpDown = true, Width = 80 };
        readonly DateTimePicker planTo = new DateTimePicker { Format = DateTimePickerFormat.Custom, CustomFormat = "HH:mm", ShowUpDown = true, Width = 80 };
        readonly CheckBox[] planDays = Enumerable.Range(0, 7).Select(i => new CheckBox { Text = Settings.DayChars[i].ToString(), AutoSize = true, Margin = new Padding(3, 6, 3, 3) }).ToArray();
        readonly CheckBox breakOn = new CheckBox { Text = "休憩時間", AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(3, 6, 6, 4) };
        readonly DateTimePicker breakStart = new DateTimePicker { Format = DateTimePickerFormat.Custom, CustomFormat = "HH:mm", ShowUpDown = true, Width = 80 };
        readonly DateTimePicker breakEnd = new DateTimePicker { Format = DateTimePickerFormat.Custom, CustomFormat = "HH:mm", ShowUpDown = true, Width = 80 };
        readonly TextBox exclude = Lines(), words = Lines();
        readonly DataGridView mask = new DataGridView();
        readonly ComboBox model = new ComboBox { Width = 260 }, fixModel = new ComboBox { Width = 260 };
        readonly TextBox claudePath = new TextBox { Width = 360 };
        readonly TextBox linkName = new TextBox { Width = 200 }, linkUrl = new TextBox { Dock = DockStyle.Fill, MinimumSize = new Size(300, 0) };

        public SettingsForm(Icon icon, Action saved, Action showAi)
        {
            this.saved = saved;
            Text = "設定" + (Core.Test ? "（テスト）" : "");
            Icon = icon;
            AutoScaleDimensions = new SizeF(96f, 96f);
            AutoScaleMode = AutoScaleMode.Dpi;
            Font = new Font("Meiryo UI", 10f);
            ClientSize = new Size(720, 520);
            MinimumSize = new Size(600, 440);
            StartPosition = FormStartPosition.CenterScreen;

            var tabs = new TabControl { Dock = DockStyle.Fill, Padding = new Point(12, 4) };

            TableLayoutPanel t = Page(tabs, "日報", 3);
            Add(t, Auto(), L("今日の日報を"), appendMinutes, L("分ごとに追記する（0 で追記しない）"));
            Add(t, Auto(), L("今日の日報を仕上げて知らせる時刻"), reportTime, L("（追記しないときは、この時刻に作る）"));
            // 休憩時間は 1 行にまとめる（列の幅をそろえるため、表の 3 列にまたがせる）
            var breakRow = new FlowLayoutPanel { AutoSize = true, WrapContents = false, Margin = new Padding(0, 2, 0, 0) };
            breakRow.Controls.Add(breakOn);
            breakRow.Controls.Add(breakStart);
            breakRow.Controls.Add(new Label { Text = "〜", AutoSize = true, Margin = new Padding(2, 6, 2, 0) });
            breakRow.Controls.Add(breakEnd);
            breakOn.CheckedChanged += (o, e) => breakStart.Enabled = breakEnd.Enabled = breakOn.Checked;
            Add(t, Auto(), breakRow);
            Add(t, Auto(), L("自動で作るのは、操作時間が"), minDay, L("分以上の日だけ"));
            Add(t, Auto(), L("日報を作ったあとに"), redo, L("分以上作業したら追記する"));
            Add(t, Auto(), L("未作成の日報をさかのぼって探す"), catchUp, L("日"));
            Add(t, Auto(), openWindow);

            t = Page(tabs, "出勤・退勤", 3);
            var planRow = new FlowLayoutPanel { AutoSize = true, WrapContents = false, Margin = new Padding(0) };
            planRow.Controls.Add(planFrom);
            planRow.Controls.Add(new Label { Text = "〜", AutoSize = true, Margin = new Padding(3, 8, 3, 3) });
            planRow.Controls.Add(planTo);
            Add(t, Auto(), L("勤務予定"), planRow);
            var dayRow = new FlowLayoutPanel { AutoSize = true, WrapContents = false, Margin = new Padding(0) };
            foreach (int i in new[] { 1, 2, 3, 4, 5, 6, 0 }) dayRow.Controls.Add(planDays[i]);   // 月〜日の順に並べる
            Add(t, Auto(), L("勤務予定の曜日"), dayRow);
            Add(t, Auto(), L("自動で出勤"), workStart, L("以降に操作があれば"));
            Add(t, Auto(), L("自動で退勤"), offIdle, L("分操作がなければ（仕上げの時刻以降。0 で自動にしない）"));

            t = Page(tabs, "記録・集計", 3);
            Add(t, Auto(), L("無操作がこの時間続いたら離席扱い"), idle, L("秒"));
            Add(t, Auto(), L("集計しない短い作業"), minMinutes, L("分未満"));
            Add(t, Auto(), L("ログと古い日報を残す"), keep, L("日（これより古いものは削除）"));
            Add(t, Auto(), L("入力がこの時間なかったら離席として記録"), afkMinutes, L("分（0 で記録しない）"));
            Add(t, Auto(), afkAsk);

            t = Page(tabs, "除外・伏せ字", 1);
            Add(t, Auto(), L("集計しないプロセス（1行に1つ。例: explorer）"));
            Add(t, new RowStyle(SizeType.Percent, 28), exclude);
            Add(t, Auto(), L("タイトルを伏せるプロセス（ウィンドウのタイトルを、右の言葉に置き換えて AI に渡す）"));
            Add(t, new RowStyle(SizeType.Percent, 40), mask);
            Add(t, Auto(), L("タイトルの中で伏せる語（1行に1つ。例: 〇〇建設）"));
            Add(t, new RowStyle(SizeType.Percent, 32), words);

            t = Page(tabs, "AI", 3);
            Add(t, Auto(), L("日報を作るモデル（ふだん）"), model);
            Add(t, Auto(), L("「AIで直す」のモデル（高度）"), fixModel);
            var browse = new Button { Text = "参照…", AutoSize = true };
            browse.Click += (o, e) => Browse();
            Add(t, Auto(), L("Claude Code の場所"), claudePath, browse);
            var ai = new Button { Text = "AIの接続先を確かめる…", AutoSize = true, Margin = new Padding(3, 12, 3, 3) };
            ai.Click += (o, e) => showAi();
            Add(t, Auto(), ai);

            t = Page(tabs, "起動", 1);
            Add(t, Auto(), startup);
            Add(t, Auto(), autoUpdate);
            var folder = new Button { Text = "保存フォルダを開く", AutoSize = true, Margin = new Padding(3, 16, 3, 3) };
            folder.Click += (o, e) => { try { Process.Start("explorer.exe", "\"" + Core.Dir + "\""); } catch (Exception ex) { Fail("開けませんでした: " + ex.Message); } };
            Add(t, Auto(), folder);

            t = Page(tabs, "日報リンク", 3);
            Add(t, Auto(), L("ボタン名"), linkName);
            var pick = new Button { Text = "ファイルを選ぶ…", AutoSize = true };
            pick.Click += (o, e) => PickLinkFile();
            Add(t, Auto(), L("URL・ファイル"), linkUrl, pick);

            // 高さの余りは最後の空の行に回す（最後の項目の行が伸びて、見出しが下にずれないように）
            foreach (TabPage pg in tabs.TabPages)
            {
                var tl = pg.Controls.Count > 0 ? pg.Controls[0] as TableLayoutPanel : null;
                if (tl == null || tl.RowStyles.Cast<RowStyle>().Any(r => r.SizeType == SizeType.Percent)) continue;
                tl.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
                tl.RowCount++;
            }

            foreach (string m in Models) { model.Items.Add(m); fixModel.Items.Add(m); }
            mask.Columns.Add(GridSort.NoColumn());
            mask.Columns.Add("proc", "プロセス名（例: OUTLOOK）");
            mask.Columns.Add("word", "置き換える言葉（例: メール対応）");
            mask.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            GridSort.Enable(mask);
            mask.Dock = DockStyle.Fill;
            mask.BackgroundColor = SystemColors.Window;
            mask.RowHeadersWidth = 28;
            mask.AllowUserToResizeRows = false;
            startup.Enabled = !Core.Test;

            var bar = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = 50, Padding = new Padding(6, 6, 6, 4), WrapContents = false, FlowDirection = FlowDirection.RightToLeft };
            var cancel = AddButton(bar, "キャンセル", (o, e) => Close());
            AddButton(bar, "保存", (o, e) => { if (Save()) Close(); });
            AddButton(bar, "初期値に戻す", (o, e) => { bool st = startup.Checked; Fill(new Settings()); startup.Checked = st; });
            CancelButton = cancel;

            Controls.Add(tabs);
            Controls.Add(bar);
            Fill(Settings.Get());
        }

        static NumericUpDown Num(int min, int max)
        {
            return new NumericUpDown { Minimum = min, Maximum = max, Width = 90, TextAlign = HorizontalAlignment.Right };
        }

        static TextBox Lines()
        {
            return new TextBox { Multiline = true, AcceptsReturn = true, ScrollBars = ScrollBars.Vertical, Dock = DockStyle.Fill };
        }

        static Label L(string s) { return new Label { Text = s, AutoSize = true, Anchor = AnchorStyles.Left | AnchorStyles.Top, Margin = new Padding(3, 8, 6, 4) }; }

        static Label Hint(string s) { return new Label { Text = s, AutoSize = true, ForeColor = Color.DimGray, Margin = new Padding(3, 2, 3, 8) }; }

        static RowStyle Auto() { return new RowStyle(SizeType.AutoSize); }

        static TableLayoutPanel Page(TabControl tabs, string title, int columns)
        {
            var page = new TabPage(title) { Padding = new Padding(10), UseVisualStyleBackColor = true };
            var t = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = columns };
            for (int i = 0; i < columns; i++)
                t.ColumnStyles.Add(i == columns - 1 ? new ColumnStyle(SizeType.Percent, 100) : new ColumnStyle(SizeType.AutoSize));
            page.Controls.Add(t);
            tabs.TabPages.Add(page);
            return t;
        }

        // 1行足す。部品が列より少なければ、最後の部品を右端の列までまたがせる（ほかの行の幅で押し出されないように）
        static void Add(TableLayoutPanel t, RowStyle style, params Control[] cs)
        {
            int r = t.RowCount++;
            t.RowStyles.Add(style);
            for (int i = 0; i < cs.Length; i++) t.Controls.Add(cs[i], i, r);
            if (cs.Length < t.ColumnCount) t.SetColumnSpan(cs[cs.Length - 1], t.ColumnCount - cs.Length + 1);
        }

        static Button AddButton(FlowLayoutPanel bar, string text, EventHandler click)
        {
            var b = new Button { Text = text, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Padding = new Padding(8, 2, 8, 2), Margin = new Padding(3) };
            b.Click += click;
            bar.Controls.Add(b);
            return b;
        }

        static void SetNum(NumericUpDown n, int v) { n.Value = Math.Max(n.Minimum, Math.Min(n.Maximum, v)); }

        void Fill(Settings s)
        {
            reportTime.Value = DateTime.Today + s.ReportTime;
            SetNum(appendMinutes, s.AppendMinutes);
            breakOn.Checked = s.HasBreak;
            breakStart.Value = DateTime.Today + (s.HasBreak ? s.BreakStart : new TimeSpan(12, 0, 0));
            breakEnd.Value = DateTime.Today + (s.HasBreak ? s.BreakEnd : new TimeSpan(13, 0, 0));
            breakStart.Enabled = breakEnd.Enabled = breakOn.Checked;
            SetNum(idle, s.IdleLimit);
            SetNum(minMinutes, s.MinMinutes);
            SetNum(minDay, s.MinDayMinutes);
            SetNum(redo, s.RedoMinutes);
            SetNum(catchUp, s.CatchUpDays);
            SetNum(keep, s.KeepDays);
            SetNum(afkMinutes, s.AfkMinutes);
            SetNum(offIdle, s.OffIdleMinutes);
            workStart.Value = DateTime.Today + s.WorkStart;
            planFrom.Value = DateTime.Today + s.PlanFrom;
            planTo.Value = DateTime.Today + s.PlanTo;
            for (int i = 0; i < 7; i++) planDays[i].Checked = s.PlanDays.Contains((DayOfWeek)i);
            afkAsk.Checked = s.AfkAsk;
            openWindow.Checked = s.OpenWindowOnReport;
            exclude.Text = string.Join("\r\n", s.ExcludeProc);
            words.Text = string.Join("\r\n", s.MaskWords);
            mask.Rows.Clear();
            int no = 0;
            foreach (KeyValuePair<string, string> kv in s.MaskProc) GridSort.Set(mask.Rows[mask.Rows.Add(null, kv.Key, kv.Value)].Cells[0], ++no, no);
            model.Text = s.Model;
            fixModel.Text = s.FixModel;
            claudePath.Text = s.ClaudePath;
            linkName.Text = s.LinkName;
            linkUrl.Text = s.LinkUrl;
            startup.Checked = TrayApp.StartupPath() != null;
            autoUpdate.Checked = s.AutoUpdate;
        }

        static List<string> Split(TextBox box)
        {
            return box.Text.Split('\n').Select(x => x.Trim()).Where(x => x.Length > 0).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
        }

        // 日報リンクに開くファイル（Excel の日報など）を選ぶ
        void PickLinkFile()
        {
            using (var d = new OpenFileDialog { Title = "日報リンクで開くファイル", Filter = "Excel (*.xlsx;*.xlsm;*.xls)|*.xlsx;*.xlsm;*.xls|すべてのファイル (*.*)|*.*" })
            {
                string cur = linkUrl.Text.Trim().Trim('"');
                try { if (cur.Length > 0 && !cur.Contains("://") && Directory.Exists(Path.GetDirectoryName(cur))) d.InitialDirectory = Path.GetDirectoryName(cur); } catch { }
                if (d.ShowDialog(this) == DialogResult.OK) linkUrl.Text = d.FileName;
            }
        }

        void Browse()
        {
            using (var d = new OpenFileDialog { Filter = "claude.exe|claude.exe|実行ファイル (*.exe)|*.exe", Title = "Claude Code（claude.exe）を選ぶ" })
            {
                string cur = claudePath.Text.Trim().Trim('"');
                if (cur.Length > 0 && Directory.Exists(Path.GetDirectoryName(cur))) d.InitialDirectory = Path.GetDirectoryName(cur);
                if (d.ShowDialog(this) == DialogResult.OK) claudePath.Text = d.FileName;
            }
        }

        bool Fail(string msg)
        {
            MessageBox.Show(this, msg, "設定", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return false;
        }

        // settings.ini はカンマ区切りなので、区切りに使う文字は入れられない
        static bool Bad(string s, bool proc)
        {
            return s.IndexOf(',') >= 0 || s.IndexOf('、') >= 0 || (proc && s.IndexOf(':') >= 0);
        }

        bool Save()
        {
            var s = new Settings
            {
                ReportTime = new TimeSpan(reportTime.Value.Hour, reportTime.Value.Minute, 0),
                IdleLimit = (int)idle.Value, MinMinutes = (int)minMinutes.Value, MinDayMinutes = (int)minDay.Value,
                RedoMinutes = (int)redo.Value, CatchUpDays = (int)catchUp.Value, KeepDays = (int)keep.Value,
                Model = model.Text.Trim().Length > 0 ? model.Text.Trim() : new Settings().Model,
                FixModel = fixModel.Text.Trim().Length > 0 ? fixModel.Text.Trim() : new Settings().FixModel,
                ClaudePath = claudePath.Text.Trim().Trim('"'),
                LinkName = linkName.Text.Trim(),
                LinkUrl = linkUrl.Text.Trim(),
                ExcludeProc = new HashSet<string>(Split(exclude), StringComparer.OrdinalIgnoreCase),
                MaskProc = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase),
                MaskWords = Split(words),
                OpenWindowOnReport = openWindow.Checked,
                AfkMinutes = (int)afkMinutes.Value, AfkAsk = afkAsk.Checked,
                AppendMinutes = (int)appendMinutes.Value,
                WorkStart = new TimeSpan(workStart.Value.Hour, workStart.Value.Minute, 0),
                PlanFrom = new TimeSpan(planFrom.Value.Hour, planFrom.Value.Minute, 0),
                PlanTo = new TimeSpan(planTo.Value.Hour, planTo.Value.Minute, 0),
                PlanDays = new HashSet<DayOfWeek>(Enumerable.Range(0, 7).Where(i => planDays[i].Checked).Select(i => (DayOfWeek)i)),
                OffIdleMinutes = (int)offIdle.Value,
                AutoUpdate = autoUpdate.Checked,
                UserName = Settings.Get().UserName,
                BreakStart = breakOn.Checked ? breakStart.Value.TimeOfDay - TimeSpan.FromSeconds(breakStart.Value.Second) : TimeSpan.Zero,
                BreakEnd = breakOn.Checked ? breakEnd.Value.TimeOfDay - TimeSpan.FromSeconds(breakEnd.Value.Second) : TimeSpan.Zero
            };
            foreach (DataGridViewRow row in mask.Rows)
            {
                if (row.IsNewRow) continue;
                string p = Convert.ToString(row.Cells[1].Value).Trim(), w = Convert.ToString(row.Cells[2].Value).Trim();
                if (p.Length == 0 && w.Length == 0) continue;
                if (p.Length == 0 || w.Length == 0) return Fail("「タイトルを伏せるプロセス」は、プロセス名と置き換える言葉の両方を入れてください。");
                if (Bad(p, true) || Bad(w, false)) return Fail("「タイトルを伏せるプロセス」に「,」「、」「:」は使えません: " + p + " / " + w);
                s.MaskProc[p] = w;
            }
            if (breakOn.Checked && !s.HasBreak) return Fail("休憩時間は、終わりを始めより後にしてください。");
            if (s.PlanTo <= s.PlanFrom) return Fail("勤務予定は、終わりを始めより後にしてください。");
            string bad = s.ExcludeProc.Concat(s.MaskWords).FirstOrDefault(x => Bad(x, false));
            if (bad != null) return Fail("「,」「、」は使えません（1行に1つずつ書いてください）: " + bad);
            if (s.ClaudePath.Length > 0 && !File.Exists(s.ClaudePath) &&
                MessageBox.Show(this, "Claude Code が見つかりません:\n" + s.ClaudePath + "\n\nこのまま保存しますか？", "設定",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes) return false;
            foreach (string m in new[] { s.Model, s.FixModel }.Distinct())
                if (!Models.Contains(m) &&
                    MessageBox.Show(this, "モデル「" + m + "」は一覧にない名前です。このまま保存しますか？", "設定",
                        MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return false;

            try { Settings.Save(s); }
            catch (Exception ex) { return Fail("保存できませんでした: " + ex.Message); }
            if (!Core.Test && startup.Checked != (TrayApp.StartupPath() != null))
            {
                try { TrayApp.SetStartup(startup.Checked); }
                catch (Exception ex) { Core.Error("スタートアップ", ex.Message); Fail("設定は保存しましたが、自動起動を変えられませんでした: " + ex.Message); }
            }
            if (saved != null) saved();
            return true;
        }
    }

    // AIの接続先: 日報を作るときにどの Claude Code・どのログインを使うか、何をどこへ渡したかの履歴を見る
    class AiForm : Form
    {
        readonly TextBox info = new TextBox();
        readonly ListView history = new ListView();
        readonly Button recheck;
        Button deep;
        readonly Action ready;

        public AiForm(Icon icon, Action ready)
        {
            this.ready = ready;
            Text = "AIの接続先（Claude Code）" + (Core.Test ? "（テスト）" : "");
            Icon = icon;
            AutoScaleDimensions = new SizeF(96f, 96f);
            AutoScaleMode = AutoScaleMode.Dpi;
            Font = new Font("Meiryo UI", 10f);
            ClientSize = new Size(900, 580);
            MinimumSize = new Size(640, 420);
            StartPosition = FormStartPosition.CenterScreen;

            info.Multiline = true;
            info.ReadOnly = true;
            info.BackColor = SystemColors.Window;
            info.ScrollBars = ScrollBars.Vertical;
            info.Dock = DockStyle.Top;
            info.Height = 250;

            var label = new Label { Text = "■ 最近の呼び出し", Dock = DockStyle.Top, Height = 28, TextAlign = ContentAlignment.BottomLeft };

            history.View = View.Details;
            history.FullRowSelect = true;
            history.Dock = DockStyle.Fill;
            history.Columns.Add("No", 50, HorizontalAlignment.Right);
            foreach (string[] c in new[] {
                new[] { "日時", "155" }, new[] { "日報の日付", "100" }, new[] { "元のログ", "120" }, new[] { "ログイン", "290" },
                new[] { "モデル", "150" }, new[] { "トークン", "80" }, new[] { "結果", "400" } })
                history.Columns.Add(c[0], int.Parse(c[1]), c[0] == "トークン" ? HorizontalAlignment.Right : HorizontalAlignment.Left);
            ListSorter.Enable(history);

            var bar = new FlowLayoutPanel { Dock = DockStyle.Bottom, AutoSize = true, Padding = new Padding(6, 6, 6, 4), WrapContents = true };
            recheck = AddButton(bar, "再確認", (o, e) => Recheck(true));
            AddButton(bar, "ログイン（ブラウザが開きます）", (o, e) => Login());
            deep = AddButton(bar, "詳しく探す", (o, e) => Deep());
            AddButton(bar, "Claude Code の入れ方", (o, e) => Install());
            AddButton(bar, "履歴ファイルを開く", (o, e) =>
            {
                string p = Path.Combine(Core.Dir, "ai_calls.log");
                if (File.Exists(p)) Process.Start("notepad.exe", "\"" + p + "\"");
            });
            AddButton(bar, "閉じる", (o, e) => Close());

            // Dock は後から追加したものが外側になる
            Controls.Add(history);
            Controls.Add(label);
            Controls.Add(info);
            Controls.Add(bar);

            Shown += (o, e) => { RefreshHistory(); Recheck(false); };
        }

        static Button AddButton(FlowLayoutPanel bar, string text, EventHandler click)
        {
            var b = new Button { Text = text, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Padding = new Padding(8, 2, 8, 2), Margin = new Padding(3) };
            b.Click += click;
            bar.Controls.Add(b);
            return b;
        }

        // PC の中を広く探して、見つかった Claude Code を使うようにする（設定の ClaudePath に入れる）
        void Deep()
        {
            deep.Enabled = recheck.Enabled = false;
            info.ForeColor = SystemColors.WindowText;
            info.Text = "検索中…";
            var t = new Thread(() =>
            {
                List<KeyValuePair<string, string>> found = ClaudeCli.DeepSearch(m => BeginInvoke((Action)(() => info.Text = "検索中… " + m)));
                BeginInvoke((Action)(() =>
                {
                    deep.Enabled = recheck.Enabled = true;
                    if (found.Count == 0)
                    {
                        info.ForeColor = Color.Firebrick;
                        info.Text = "Claude Code は見つかりませんでした。「Claude Code の入れ方」を参照してください。";
                        return;
                    }
                    string pick = found[0].Key;
                    try { Settings.SetClaudePath(pick); }
                    catch (Exception ex) { info.Text = "設定に保存できません: " + ex.Message + "\r\n" + pick; return; }
                    Recheck(true);
                    MessageBox.Show(this, "使用する Claude Code:\r\n" + pick + "\r\n" + found[0].Value, "詳しく探す", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }));
            });
            t.IsBackground = true;
            t.Start();
        }

        // Claude Code の入れ方（公式のページ）をブラウザで開く。入れるのは利用者が公式の手順で行う
        void Install()
        {
            try
            {
                Process.Start("https://code.claude.com/docs/en/setup");
                info.Text = "入れ方のページを開きました。入れたあと「再確認」→「ログイン」。";
            }
            catch (Exception ex) { info.Text = "ブラウザを開けませんでした: " + ex.Message + "（https://code.claude.com/docs/en/setup）"; }
        }

        // Claude Code の場所・ログイン・環境変数を確かめる（数秒かかるので別スレッドで）
        void Recheck(bool byUser)
        {
            recheck.Enabled = false;
            info.ForeColor = SystemColors.WindowText;
            info.Text = "確かめています…";
            var t = new Thread(() =>
            {
                Settings s = Settings.Get();
                string exe;
                ClaudeCli.Auth a;
                string problem = ClaudeCli.Problem(s, out exe, out a);
                var sb = new StringBuilder();
                sb.AppendLine("■ 日報を作るときに使う AI");
                sb.AppendLine("Claude Code: " + (exe == null ? "見つかりません" : exe + "（" + ClaudeCli.Version(exe) + "）"));
                if (exe == null) sb.AppendLine("　" + ClaudeCli.NotFoundHelp().Replace("\r\n", "\r\n　"));
                if (exe != null) sb.AppendLine("　種類: " + ClaudeCli.Source(exe));
                sb.AppendLine("　探し方: " + (s.ClaudePath.Length > 0 ? "settings.ini の ClaudePath" : "PATH・%USERPROFILE%\\.local\\bin・Claude デスクトップアプリの中から自動"));
                if (a != null)
                {
                    sb.AppendLine("ログイン: " + (a.LoggedIn ? a.Method : "していません"));
                    if (a.Email.Length > 0) sb.AppendLine("アカウント: " + a.Email + (a.Org.Length > 0 ? "（" + a.Org + "）" : ""));
                    if (a.Plan.Length > 0) sb.AppendLine("プラン: " + a.Plan);
                }
                sb.AppendLine("料金: " + (problem == null ? "Claude のプランの利用枠から（API の従量課金は使いません）" : "—"));
                sb.AppendLine("モデル: 日報を作るとき " + s.Model + "／「AIで直す」 " + s.FixModel);
                foreach (string n in ClaudeCli.StripEnv)
                {
                    bool has = !string.IsNullOrEmpty(Environment.GetEnvironmentVariable(n)) ||
                               !string.IsNullOrEmpty(Environment.GetEnvironmentVariable(n, EnvironmentVariableTarget.User));
                    sb.AppendLine("環境変数 " + n + ": " + (has ? "あり（日報を作るときは外して呼ぶので使いません）" : "なし"));
                }
                if (problem != null) sb.AppendLine("→ 今は日報を作れません: " + problem);
                sb.AppendLine();
                sb.AppendLine("■ 日報の元にしているログ");
                sb.AppendLine(Core.Dir + "\\yyyy-MM-dd.csv（60秒ごとの前面ウィンドウの記録）");
                sb.Append("AI に渡すのは、このログを集計したもの（アプリ名・ウィンドウタイトル・時刻。伏せ字の設定を反映）だけです。");
                string text = sb.ToString();
                if (IsDisposed) return;
                BeginInvoke((Action)(() =>
                {
                    info.ForeColor = problem != null ? Color.Firebrick : SystemColors.WindowText;
                    info.Text = text;
                    recheck.Enabled = true;
                    if (byUser && problem == null && ready != null) ready();
                }));
            });
            t.IsBackground = true;
            t.Start();
        }

        void Login()
        {
            string exe = ClaudeCli.FindExe(Settings.Get());
            if (exe == null) { info.Text = "Claude Code（claude.exe）が見つかりません。\r\n" + ClaudeCli.NotFoundHelp(); return; }
            try
            {
                Process.Start(new ProcessStartInfo("cmd.exe", "/k \"\"" + exe + "\" auth login\"") { UseShellExecute = true });
                info.Text = "ログイン後、黒い画面を閉じて「再確認」。";
            }
            catch (Exception ex) { info.Text = "ログインを始められませんでした: " + ex.Message; }
        }

        public void RefreshHistory()
        {
            history.BeginUpdate();
            history.Items.Clear();
            foreach (string[] p in ClaudeCli.RecentLog(50))
            {
                long n;
                var it = new ListViewItem(new[] { (history.Items.Count + 1).ToString(), p[0], p[1], p[2], p[4], p[5], long.TryParse(p[6], out n) ? n.ToString("N0") : p[6], p[7] });
                it.ToolTipText = "Claude Code: " + p[3];
                if (!p[7].StartsWith("成功")) it.ForeColor = Color.Firebrick;
                history.Items.Add(it);
            }
            history.ShowItemToolTips = true;
            history.EndUpdate();
        }
    }

    // 動作確認（--selftest）: 作り直し・AIで直すの処理と、日報画面の帯を、決めた例で確かめる
    static class SelfTest
    {
        // テスト用の保存先に、2026-01-05 の記録・離席・勤務の履歴・AI の呼び出しを作る（何度呼んでも同じ）
        public static string MakeDay()
        {
            const string date = "2026-01-05";
            var csv = new StringBuilder("time,process,title,idle\r\n");
            for (DateTime t = new DateTime(2026, 1, 5, 8, 30, 0); t < new DateTime(2026, 1, 5, 18, 0, 0); t = t.AddMinutes(1))
                csv.Append("\"" + t.ToString("yyyy-MM-dd HH:mm:ss") + "\",\"EXCEL\",\"数量計算.xlsx\"," + (t.Hour == 10 && t.Minute < 10 ? 400 : 0) + "\r\n");
            Core.WriteText(Core.CsvPath(date), csv.ToString());
            string afk = Path.Combine(Core.Dir, "離席_" + date + ".txt");
            if (File.Exists(afk)) File.Delete(afk);
            Afk.Put(new AfkEntry { Start = new DateTime(2026, 1, 5, 15, 0, 0), End = new DateTime(2026, 1, 5, 15, 40, 0), Memo = "現場の立会い" });
            File.WriteAllText(Path.Combine(Core.Dir, "勤務履歴.log"),
                "off\t2026-01-05 17:00:00\r\non\t2026-01-06 08:00:00\r\n", new UTF8Encoding(false));
            File.WriteAllText(Path.Combine(Core.Dir, "ai_calls.log"),
                "2026-01-05 08:35:00\t2026-01-04\ta.csv\tclaude\tx\tm\t1\t成功\r\n" +
                "2026-01-05 09:40:00\t2026-01-05\ta.csv\tclaude\tx\tm\t1\t成功（追記）\r\n" +
                "2026-01-05 13:40:00\t2026-01-05\ta.csv\tclaude\tx\tm\t1\t失敗: 利用枠\r\n" +
                "2026-01-05 16:50:00\t2026-01-05\ta.csv\tclaude\tx\tm\t1\t成功（追記）\r\n", new UTF8Encoding(true));
            return date;
        }

        public static int Run()
        {
            var log = new StringBuilder();
            int ng = 0;
            Action<string, string, string> eq = (what, got, want) =>
            {
                bool ok = got == want;
                if (!ok) ng++;
                log.Append((ok ? "OK " : "NG ") + what + " => " + (got ?? "(null)") + (ok ? "" : "  期待: " + (want ?? "(null)")) + "\r\n");
            };
            // 作り直し: 真ん中の日報の書式の正しい行（AI が落としたら戻す）
            string center = "■工程\n・[1] 09:00〜10:30 A工区の数量計算\n資料作成\n■軽作業\n・[2] 10:40 メール対応\n13時 電話";
            eq("正しい行", string.Join(" | ", Reporter.CorrectLines(center).Select(k => k.Key + (k.Value ? "(軽)" : ""))),
               "・09:00〜10:30 A工区の数量計算 | ・10:40 メール対応(軽)");
            var aiOut = new List<string> { "■工程", "・09:00〜10:30 A工区の数量計算", "・11:00〜12:00 資料作成", "■軽作業", "・13:00 電話" };
            eq("AI が落とした行を戻す（数）", Reporter.KeepCorrect(center, aiOut).ToString(), "1");
            eq("AI が落とした行を戻す（行）", string.Join(" | ", aiOut), "■工程 | ・09:00〜10:30 A工区の数量計算 | ・11:00〜12:00 資料作成 | ■軽作業 | ・10:40 メール対応 | ・13:00 電話");

            // タスク: ウィンドウタイトル・アプリ名から当てはめる（上のタスクが先・完了は使わない・語が空なら名前）
            var tl = new List<TaskItem>
            {
                new TaskItem { Name = "〇〇橋 詳細設計", Keywords = "〇〇橋,ﾏﾙﾏﾙ橋" },
                new TaskItem { Name = "R8工程表", Keywords = "" },
                new TaskItem { Name = "済んだ案件", Keywords = "数量", Done = true },
                new TaskItem { Name = "メール", Keywords = "OUTLOOK" },
            };
            eq("タスク: 語", Tasks.Match(tl, "EXCEL", "〇〇橋_数量計算.xlsx - Excel"), "〇〇橋 詳細設計");
            eq("タスク: 名前", Tasks.Match(tl, "EXCEL", "全体R8工程表.xlsx"), "R8工程表");
            eq("タスク: 完了は使わない", Tasks.Match(tl, "EXCEL", "数量計算.xlsx") ?? "(なし)", "(なし)");
            eq("タスク: アプリ名", Tasks.Match(tl, "OUTLOOK", "受信トレイ"), "メール");

            eq("タイトルの見出し", Tasks.Key("EXCEL", " *数量計算.xlsx - Excel "), "excel\t数量計算.xlsx - Excel");
            var tmap = new Dictionary<string, Tasks.TitleTask>
            {
                { Tasks.Key("EXCEL", "〇〇橋_数量計算.xlsx - Excel"), new Tasks.TitleTask { Task = "別のタスク" } },
                { Tasks.Key("chrome", "YouTube"), new Tasks.TitleTask { Task = Tasks.NoTask } },
            };
            eq("タイトルの対応が語より先", Tasks.Resolve(tmap, tl, "EXCEL", "●〇〇橋_数量計算.xlsx - Excel"), "別のタスク");
            eq("「なし」はタスクではない", Tasks.Resolve(tmap, tl, "CHROME", "YouTube") ?? "(なし)", "(なし)");
            eq("対応がなければ語", Tasks.Resolve(tmap, tl, "OUTLOOK", "受信トレイ"), "メール");
            Dictionary<int, string> pp = Tasks.ParsePredict("1: 〇〇橋 詳細設計\n2：なし\nわかりません\n3. 「R8工程表」");
            eq("予測の読み取り", string.Join(" | ", pp.OrderBy(x => x.Key).Select(x => x.Key + "=" + x.Value)), "1=〇〇橋 詳細設計 | 2=なし | 3=R8工程表");

            // 工程をつなげる: 短いすき間は前を延ばす・15分以上は「軽作業」で埋める・休憩の後は次の始めを動かす・重なりは前の終わりを縮める
            var procs = new List<string> { "■工程", "・09:00〜10:00 A", "・10:05〜11:00 B", "・11:30〜12:00 C", "・12:00〜13:00 休憩", "・13:10〜15:00 D", "・14:50〜16:00 E", "", "■軽作業", "・11:10 メール" };
            int pc = Reporter.ConnectProcesses(procs, new HashSet<string> { "・12:00〜13:00 休憩" });
            eq("工程をつなげる（行）", string.Join(" | ", procs),
               "■工程 | ・09:00〜10:05 A | ・10:05〜11:00 B | ・11:00〜11:30 軽作業 | ・11:30〜12:00 C | ・12:00〜13:00 休憩 | ・13:00〜14:50 D | ・14:50〜16:00 E |  | ■軽作業 | ・11:10 メール");
            eq("工程をつなげる（数）", pc.ToString(), "3");

            // 日報リンク: 古い設定ツールの1行書きを直す・開く先
            Settings ls = Settings.FromText("LinkName=テスト LinkUrl=https://example.com/a\r\n");
            eq("日報リンク: 1行書きを直す", ls.LinkName + " | " + ls.LinkUrl, "テスト | https://example.com/a");
            eq("日報リンク: 既定は空", new Settings().LinkName + "|" + new Settings().LinkUrl, "|");
            DateTime ld = new DateTime(2026, 1, 5);
            eq("日報リンク: URL", Reporter.LinkTarget("https://example.com/{yyyy}/{MM}", ld), "https://example.com/2026/01");
            eq("日報リンク: ファイル", Reporter.LinkTarget(@"C:\日報\日報_{yyyy}{MM}.xlsx", ld), @"C:\日報\日報_202601.xlsx");
            eq("日報リンク: 共有フォルダ", Reporter.LinkTarget(@"\\server\share\日報.xlsx", ld), @"\\server\share\日報.xlsx");
            eq("日報リンク: ドメインだけ", Reporter.LinkTarget("example.com/x", ld), "https://example.com/x");

            // 勤務予定（8:30〜17:30）: 外の時刻を収め、始業・終業までのすき間を「現場（外勤）」に
            var plan1 = new List<string> { "■工程", "・08:00〜10:00 A", "・10:00〜15:00 B", "■軽作業", "・18:10 メール", "・08:10 電話" };
            Reporter.ConnectProcesses(plan1, new HashSet<string>(), new TimeSpan(8, 30, 0), new TimeSpan(17, 30, 0));
            eq("勤務予定: 収める・終業まで外勤", string.Join(" | ", plan1),
               "■工程 | ・08:30〜10:00 A | ・10:00〜15:00 B | ・15:00〜17:30 現場（外勤） | ■軽作業 | ・17:30 メール | ・08:30 電話");
            var plan2 = new List<string> { "■工程", "・09:40〜17:20 A" };
            Reporter.ConnectProcesses(plan2, new HashSet<string>(), new TimeSpan(8, 30, 0), new TimeSpan(17, 30, 0));
            eq("勤務予定: 始業まで外勤・短い終業前は延ばす", string.Join(" | ", plan2), "■工程 | ・08:30〜09:40 現場（外勤） | ・09:40〜17:30 A");
            Settings ps = Settings.FromText("PlanHours=09:00-18:00\r\nPlanDays=月水金\r\n");
            eq("勤務予定の設定", ps.PlanText + " " + ps.PlanDaysText + " " + ps.PlanOn(new DateTime(2026, 11, 11)) + " " + ps.PlanOn(new DateTime(2026, 11, 12)), "09:00〜18:00 月水金 True False");
            eq("勤務予定の既定", new Settings().PlanText + " " + new Settings().PlanDaysText, "08:30〜17:30 月火水木金");

            // AIで直す（会話）: 返事と日報の案に分ける
            string reply, rep;
            rep = Reporter.ParseChat("<返事>\n3番を消しました\n</返事>\n<日報>\n■工程\n・09:00〜10:00 A\n</日報>", out reply);
            eq("会話: 返事", reply, "3番を消しました");
            eq("会話: 案", rep, "■工程\n・09:00〜10:00 A");
            rep = Reporter.ParseChat("<返事>何時ごろの作業ですか？</返事>", out reply);
            eq("会話: 質問だけ", (rep == null) + " " + reply, "True 何時ごろの作業ですか？");
            rep = Reporter.ParseChat("■工程\n・09:00〜10:00 A", out reply);
            eq("会話: 形を守らず日報だけ", (rep != null) + " " + reply, "True 日報を直しました。青字のところを確かめてください。");

            // 保存先の直下の記録・日報をフォルダへ移す（0.22.0 より前の置き方）
            Core.WriteText(Path.Combine(Core.Dir, "2026-01-04.csv"), "x");
            Core.WriteText(Path.Combine(Core.Dir, "日報_2026-01-04.txt"), "y");
            Core.MoveToFolders();
            eq("フォルダへ移す", File.Exists(Core.CsvPath("2026-01-04")) + " " + File.Exists(Core.ReportPath("2026-01-04")) + " " +
               File.Exists(Path.Combine(Core.Dir, "2026-01-04.csv")) + " " + File.Exists(Path.Combine(Core.Dir, "日報_2026-01-04.txt")), "True True False False");

            // 開いていたウィンドウ: 開いた〜閉じた
            Core.WriteText(Core.WindowsPath("2026-01-05"), "\"time\",\"event\",\"process\",\"title\"\r\n" +
                "\"2026-01-05 09:00:00\",\"open\",\"EXCEL\",\"A.xlsx\"\r\n\"2026-01-05 09:05:00\",\"open\",\"WINWORD\",\"B.docx\"\r\n" +
                "\"2026-01-05 09:08:00\",\"close\",\"WINWORD\",\"B.docx\"\r\n\"2026-01-05 11:00:00\",\"close\",\"EXCEL\",\"A.xlsx\"\r\n" +
                "\"2026-01-05 16:00:00\",\"open\",\"EXCEL\",\"C.xlsx\"\r\n\"2026-01-05 17:00:00\",\"open\",\"EXCEL\",\"D.xlsx\"\r\n");
            eq("開いていたウィンドウ", string.Join(" | ", Reporter.WindowSpans("2026-01-05").Select(x => x.Title + " " + x.Start.ToString("HH:mm") + "-" + x.End.ToString("HH:mm"))),
               "A.xlsx 09:00-11:00 | B.docx 09:05-09:08 | C.xlsx 16:00-17:00");

            // 1日の帯: 記録 8:30〜17:59（10:00〜10:09 は無操作）、休憩 12:00〜13:00、離席 15:00〜15:40、17:00 に退勤
            Slot[] day = DayStatus.Build(MakeDay(), Settings.Get());
            Func<Slot, string> count = k => day.Count(x => x == k).ToString();
            eq("帯の作業（分）", count(Slot.Work), "400");
            eq("帯の無操作（分）", count(Slot.Idle), "10");
            eq("帯の休憩（分）", count(Slot.Break), "60");
            eq("帯の離席（分）", count(Slot.Afk), "40");
            eq("帯の退勤（分）", count(Slot.Off), "420");
            eq("AI の呼び出し", ClaudeCli.CallsOn(new DateTime(2026, 1, 5)).Count(c => c.Value) + "/" + ClaudeCli.CallsOn(new DateTime(2026, 1, 5)).Count, "3/4");
            // Claude Code を探す: デスクトップアプリ本体の別名は除く・版の読み取り・デスクトップアプリの中を探す
            eq("デスクトップアプリ本体（ストア版の別名）", ClaudeCli.IsDesktopApp(@"C:\Users\a\AppData\Local\Microsoft\WindowsApps\Claude.exe").ToString(), "True");
            eq("単体の Claude Code", ClaudeCli.IsDesktopApp(@"C:\Users\a\.local\bin\claude.exe").ToString(), "False");
            eq("パスの版", ClaudeCli.PathVersion(@"C:\Users\a\AppData\Roaming\Claude\claude-code\2.1.30\0a1b2c\claude.exe").ToString(), "2.1.30");
            string expect = Environment.GetEnvironmentVariable("WORKLOG_EXPECT_DESKTOP");
            if (!string.IsNullOrEmpty(expect)) eq("デスクトップアプリの中の Claude Code", ClaudeCli.FindDesktop(), expect);
            log.Append(ng == 0 ? "すべて期待どおり\r\n" : ng + " 件が期待と違う\r\n");
            Core.WriteText(Path.Combine(Core.Dir, "selftest.log"), log.ToString());
            return ng == 0 ? 0 : 1;
        }
    }

    static class Program
    {
        [STAThread]
        static int Main(string[] args)
        {
            try { Directory.CreateDirectory(Core.Dir); }
            catch (Exception ex) { MessageBox.Show("保存先を作れません: " + Core.Dir + "\n" + ex.Message, "作業ログ"); return 2; }
            Core.MoveToFolders();

            // 画面なしの実行
            if (args.Length >= 1 && (args[0] == "--auto" || args[0] == "--report"))
            {
                try
                {
                    List<Result> res;
                    if (args[0] == "--auto") res = Reporter.Locked(() => Reporter.Auto(Settings.Get()));
                    else
                    {
                        if (args.Length < 2 || !Reporter.IsDate(args[1])) return 2;
                        res = Reporter.Locked(() => new List<Result> { Reporter.Make(args[1], Settings.Get()) });
                    }
                    return res.All(r => r.Ok) ? 0 : 1;
                }
                catch (Exception ex) { Core.Error("日報", ex.ToString()); return 2; }
            }

            // テスト用: 日報の書式を直す処理を確かめる（0=すべて期待どおり 1=違う）。結果は保存先の selftest.log
            if (args.Length >= 1 && args[0] == "--selftest" && Core.Test) return SelfTest.Run();

            // テスト用: 日報画面の帯（テスト用の 2026-01-05）を画像にして終わる（見た目の確認）
            if (args.Length >= 2 && args[0] == "--screenshot-timeline" && Core.Test)
            {
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                string date = SelfTest.MakeDay();
                using (var f = new Form { Text = "帯", Font = new Font("Meiryo UI", 10f), ClientSize = new Size(900, 80), StartPosition = FormStartPosition.CenterScreen })
                {
                    var t = new DayTimeline(null);
                    f.Controls.Add(t);
                    f.Show();
                    t.SetDate(date);
                    Application.DoEvents();
                    using (var bmp = new Bitmap(t.Width, t.Height))
                    {
                        t.DrawToBitmap(bmp, new Rectangle(0, 0, t.Width, t.Height));
                        bmp.Save(args[1], System.Drawing.Imaging.ImageFormat.Png);
                    }
                }
                return 0;
            }

            // テスト用: 設定画面の各タブを画像にして終わる（見た目の確認）
            if (args.Length >= 2 && args[0] == "--screenshot-settings" && Core.Test)
            {
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Directory.CreateDirectory(args[1]);
                using (var f = new SettingsForm(SystemIcons.Application, null, null))
                {
                    f.StartPosition = FormStartPosition.CenterScreen;
                    f.Show();
                    TabControl tc = f.Controls.OfType<TabControl>().First();
                    for (int i = 0; i < tc.TabCount; i++)
                    {
                        tc.SelectedIndex = i;
                        Application.DoEvents();
                        using (var bmp = new Bitmap(f.Width, f.Height))
                        {
                            f.DrawToBitmap(bmp, new Rectangle(0, 0, f.Width, f.Height));
                            bmp.Save(Path.Combine(args[1], "settings_" + i + ".png"), System.Drawing.Imaging.ImageFormat.Png);
                        }
                    }
                }
                return 0;
            }

            // テスト用: バージョン情報の画面を画像にして終わる（見た目の確認）
            if (args.Length >= 2 && args[0] == "--screenshot-about" && Core.Test)
            {
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                using (var f = new AboutForm(SystemIcons.Application))
                {
                    f.StartPosition = FormStartPosition.CenterScreen;
                    f.Show();
                    Application.DoEvents();
                    using (var bmp = new Bitmap(f.Width, f.Height))
                    {
                        f.DrawToBitmap(bmp, new Rectangle(0, 0, f.Width, f.Height));
                        bmp.Save(args[1], System.Drawing.Imaging.ImageFormat.Png);
                    }
                }
                return 0;
            }

            // テスト用: 更新を確かめて落とす（--apply で入れ替えて新しい exe を起動する）
            if (args.Length >= 1 && args[0] == "--check-update" && Core.Test)
            {
                try
                {
                    Updater.Release rel = Updater.Latest();
                    Updater.Log("確認（--check-update）: 最新 " + rel.Version + " / 今 " + Updater.Current);
                    if (!Updater.IsNewer(rel.Version, Updater.Current)) return 0;
                    string exe = Updater.Download(rel);
                    if (args.Contains("--apply")) Updater.Apply(exe);
                    return 10;
                }
                catch (Exception ex) { Updater.Log("失敗（--check-update）: " + ex); return 2; }
            }

            // 自動アップデートで起動されたときは、前の版が終わるのを待つ
            int at = Array.IndexOf(args, "--updated"), oldPid = 0;
            bool updated = at >= 0 && at + 1 < args.Length && int.TryParse(args[at + 1], out oldPid);
            if (updated) Updater.AfterUpdate(oldPid); else Updater.Cleanup();
            bool firstRun = !File.Exists(Settings.FilePath);

            // 二重起動のときは、動いている方の日報画面を開いて終わる
            bool created;
            using (var mutex = new Mutex(true, "WorkLogTray" + Core.Suffix, out created))
            using (var showEvent = new EventWaitHandle(false, EventResetMode.AutoReset, "WorkLogTrayShow" + Core.Suffix))
            {
                if (!created) { showEvent.Set(); return 0; }

                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);
                Application.ThreadException += (o, e) => Core.Error("画面", e.Exception.ToString());
                AppDomain.CurrentDomain.UnhandledException += (o, e) => Core.Error("異常終了", e.ExceptionObject.ToString());

                var app = new TrayApp(args.Contains("--show"), updated, firstRun);
                if (args.Contains("--ai")) app.ShowAiForm();
                if (args.Contains("--settings")) app.ShowSettingsForm();
                if (args.Contains("--test-afk") && Core.Test) app.TestAfk();
                SynchronizationContext ui = SynchronizationContext.Current;
                var waiter = new Thread(() =>
                {
                    while (true)
                    {
                        showEvent.WaitOne();
                        ui.Post(_ => app.ShowForm(null), null);
                    }
                });
                waiter.IsBackground = true;
                waiter.Start();
                Application.Run(app);
            }
            return 0;
        }
    }
}
