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
        public const string Version = "0.10.0";   // 主.機能.修正
        public const string History =
            "0.10.0（2026-11-05）自動アップデート（GitHub のリリースに新しい版が出たら、日報画面を閉じているときに入れ替えて起動し直す。設定で「知らせるだけ」にもできる。メニュー「ヘルプ → 更新を確認」）。exe にアイコン。トレイのメニューと初めての起動でスタートアップに登録・解除\n" +
            "0.9.0（2026-10-30）作り直すときは、今の日報と作り直した案を Opus で比べて合成する。休憩（12:00〜13:00）の行を日報に残す（日報画面のチェックで外せる）。" +
            "本文の下の「手で追記」で時刻と内容を入れて追記。出勤・退勤（退勤中は記録と AI を止めて待機。17:30 を過ぎて操作がなければ退勤、8:00 を過ぎて操作があれば出勤）\n" +
            "0.8.1（2026-10-30）追記のたびに作業が1時間ごとに分かれないよう、続いている工程は終わりの時刻を延ばす（区切るのは休憩と離席だけ）\n" +
            "0.8.0（2026-10-30）日報を逐次追記する（今日の分は1時間ごと。手で直した行は変えない）。離席メモはすぐ日報に足す。休憩時間（12:00〜13:00）を日報に入れない\n" +
            "0.7.0（2026-10-30）作業の行に通し番号 [1] [2] … を付ける。「AIで直す」で「3番は…」と指示できる。コピーには番号を入れない\n" +
            "0.6.0（2026-10-30）離席メモ: 離席から戻ったときに何をしていたか聞く。日報画面の「離席メモ」。日報に時刻つきで入れる\n" +
            "0.5.0（2026-10-30）「AIで直す」を日報画面の右の欄へ。案は本文に青字で出す。直すときは Opus\n" +
            "0.4.0（2026-10-30）「AIで直す」を追加。訂正の記録を作り直しにも使う\n" +
            "0.3.0（2026-10-29）設定画面とメニューバー\n" +
            "0.2.0（2026-10-29）AI を Claude Code（プランのログイン）で呼ぶ。「AIの接続先」画面\n" +
            "0.1.0（2026-10-29）PowerShell 版からタスクトレイ常駐の exe に作り替え";        public static readonly bool Test = !string.IsNullOrEmpty(Environment.GetEnvironmentVariable("WORKLOG_TEST"));
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

        public static string ReportPath(string date) { return Path.Combine(Dir, "日報_" + date + ".txt"); }
        public static string CsvPath(string date) { return Path.Combine(Dir, date + ".csv"); }
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
        public int OffIdleMinutes = 30;  // ReportTime を過ぎて、操作がこの分数なかったら退勤とみなす（0 で自動で退勤しない）
        public bool AutoUpdate = true;   // 新しい版が出たら自動で入れ替える（false は知らせるだけ）

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
            add(@"Claude Code（claude.exe）の場所。空なら PATH と %USERPROFILE%\.local\bin から探す", "ClaudePath", ClaudePath);
            add("集計しないプロセス（カンマ区切り）", "ExcludeProc", string.Join(",", ExcludeProc));
            add("タイトルを伏せるプロセス（プロセス名:置き換える言葉、カンマ区切り）", "MaskProc", string.Join(",", MaskProc.Select(kv => kv.Key + ":" + kv.Value)));
            add("タイトル中で伏せる語（カンマ区切り。例: 〇〇建設,△△様）", "MaskWords", string.Join(",", MaskWords));
            add("日報ができたら画面を開く（1=開く 0=通知だけ）", "OpenWindowOnReport", OpenWindowOnReport ? "1" : "0");
            add("入力がこの分数なかったら離席として記録する（0 で記録しない）", "AfkMinutes", AfkMinutes.ToString());
            add("離席から戻ったときに何をしていたか聞く（1=聞く 0=日報画面の離席メモに書くだけ）", "AfkAsk", AfkAsk ? "1" : "0");
            add("退勤中（待機）でも、この時刻を過ぎて操作があれば出勤とみなす", "WorkStart", WorkStart.ToString(@"hh\:mm"));
            add("ReportTime を過ぎて、操作がこの分数なかったら退勤とみなす（0 で自動で退勤しない）", "OffIdleMinutes", OffIdleMinutes.ToString());
            add("新しい版が出たら自動で更新する（1=自動で更新 0=知らせるだけ）", "AutoUpdate", AutoUpdate ? "1" : "0");
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
                        case "OffIdleMinutes": s.OffIdleMinutes = int.Parse(v); break;
                        case "AutoUpdate": s.AutoUpdate = v == "1" || v.ToLower() == "true"; break;
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

        public static string FindExe(Settings s)
        {
            if (s.ClaudePath.Length > 0) return File.Exists(s.ClaudePath) ? s.ClaudePath : null;
            string path = Environment.GetEnvironmentVariable("PATH") + ";" + Environment.GetEnvironmentVariable("PATH", EnvironmentVariableTarget.User);
            foreach (string d in path.Split(';'))
            {
                try
                {
                    string p = Path.Combine(Environment.ExpandEnvironmentVariables(d.Trim().Trim('"')), "claude.exe");
                    if (d.Trim().Length > 0 && File.Exists(p)) return p;
                }
                catch { }
            }
            string home = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ".local", "bin", "claude.exe");
            return File.Exists(home) ? home : null;
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
        static string LogText(List<Row> active, Settings s)
        {
            var groups = new Dictionary<string, int>();
            var segs = new List<Seg>();
            foreach (Row a in active)
            {
                string t;
                if (!s.MaskProc.TryGetValue(a.Proc, out t)) t = a.Title;
                foreach (string w in s.MaskWords) t = t.Replace(w, "＊＊");
                string k = a.Proc + "\t" + t;
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
            return "合計時間:\n" + summary + "\n\n時刻順:\n" + timeline;
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

        // 「AIで直す」: 今の日報を訂正の指示どおりに直した案を返す（ファイルには書かない）。1行目の見出しはそのまま残す
        public static string Revise(string date, string current, string instruction, string notes, Settings s)
        {
            string[] lines = current.Replace("\r\n", "\n").Split('\n');
            bool hasHeader = lines.Length > 0 && lines[0].StartsWith("日報");
            string header = hasHeader ? lines[0] : null;
            string body = string.Join("\n", hasHeader ? lines.Skip(1) : lines).Trim();
            List<Row> active = ReadActive(date, s);

            string prompt =
                "次の日報を、利用者の訂正の指示どおりに直してください。\n\n" +
                "条件:\n" +
                "- 指示に関係する行だけを直し、ほかの行は一字も変えずに残す\n" +
                "- 形は変えない（■工程は「・[番号] HH:MM〜HH:MM 作業内容」、■軽作業は「・[番号] HH:MM 作業内容」、それぞれ時刻の早い順）\n" +
                "- 行頭の [番号] は作業の通し番号。指示の「3番」などはこの番号の行のこと。番号は付けたままでよい（あとで振り直す）\n" +
                "- 時刻を決めるときは、下の作業ログを参考にしてよい\n" +
                "- 「休憩」の行は、指示がない限りそのまま残す\n" +
                "- 出力は直した日報の本文だけ（前置き・説明・見出しの日付行は不要）\n\n" +
                "訂正の指示:\n" + instruction.Trim() + "\n\n" +
                (notes.Trim().Length > 0 ? "これまでの訂正（これも守る）:\n" + notes.Trim().Replace("\r\n", "\n") + "\n\n" : "") +
                Afk.Prompt(date) +
                "今の日報:\n" + body + "\n\n" +
                (active != null && active.Count > 0 ? "参考: この日の作業ログ（離席時間は除いています）\n" + LogText(active, s) : "（この日の作業ログはありません）");

            string revised = ClaudeCli.Ask(prompt, s, date, "（AIで直す）", s.FixModel).Trim();
            // 日報の形で返ってこなかった（質問や説明だった）ときは案にしない
            if (!revised.Replace("\r\n", "\n").Split('\n').Any(l => l.Trim().StartsWith("■")))
                throw new Exception("AI が日報の形で返しませんでした。指示を具体的に書き直してください（AI の返事: " + (revised.Length > 120 ? revised.Substring(0, 120) + "…" : revised) + "）");
            List<string> result = SplitLines((header != null ? header + "\n\n" : "") + revised);
            if (!instruction.Contains("休憩")) EnsureBreak(result, date, s);   // AI がうっかり休憩の行を落としたとき
            return Number(string.Join("\n", result).TrimEnd() + "\n");
        }

        // 直した日報と訂正の記録を保存する（前の日報は old フォルダへ）
        public static void SaveRevision(string date, string text, string notes)
        {
            string path = Core.ReportPath(date);
            Backup(path);
            Core.WriteText(path, text);
            SaveNotes(date, notes);
        }

        public static void SaveNotes(string date, string notes)
        {
            string p = NotesPath(date);
            if (notes.Trim().Length > 0) Core.WriteText(p, notes.Trim() + "\r\n");
            else if (File.Exists(p)) File.Delete(p);
        }

        // 1日分の日報を作る。日報がもうあるとき（手で直した行・離席メモ・手で足した作業）は、作り直した案と Opus で比べて合成する
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
                "- 工程・軽作業とも時刻の早い順。時刻は作業ログに出てくる時刻を使う\n" +
                "- 作業内容は20文字程度。ファイル名やページ名から業務内容を具体的に推定する（例:「〇〇工区の数量計算」）\n" +
                "- 推定できないもの、私用と思われるものは省く\n" +
                "- 「時刻のない本人のメモ」があれば、作業ログになくても業務として必ず入れる\n" +
                BreakRule(s, date) +
                "- 該当がなければ「・なし」\n\n" +
                NotesPrompt(date) + Afk.Prompt(date) + ManualPrompt(date) + LogText(active, s);

            string header = Header(date, active);
            string path = Core.ReportPath(date);
            string existing = File.Exists(path) ? Core.ReadText(path) : "";
            bool merge = existing.Contains("■");

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
                r.Message = ex.Message + "。トレイのメニュー「AIの接続先」で確かめてください" + (merge ? "（今の日報はそのままです）" : "");
                if (merge) return r;
                body = "（" + r.Message + "。直すと自動で作り直します。急ぐときは下の文を Claude に貼ってください）\n\n" + prompt;
            }
            catch (Exception ex)
            {
                r.Ok = false;
                r.Message = ex.Message + (merge ? "（今の日報はそのままです）" : "");
                Core.Error("日報 " + date, "AI失敗: " + ex.Message);
                if (merge) return r;
                body = "（AIで日報を作れませんでした: " + ex.Message + "）\n" +
                       "（あとで自動でもう一度試します。急ぐときは下の文を Claude に貼ってください）\n\n" + prompt;
            }

            if (merge) return Combine(date, s, r, existing, header, body, active);

            lock (reportLock)
            {
                Backup(path);
                List<string> lines = SplitLines(header + "\n\n" + body.Trim());
                if (r.Ok) { ApplyManual(lines, date); EnsureBreak(lines, date, s); }
                string text = string.Join("\n", lines).TrimEnd() + "\n";
                Core.WriteText(path, r.Ok ? Number(text) : Regex.Replace(text, "\r?\n", "\r\n"));
            }
            if (r.Ok) ApplyAllAfk(date);   // メモのある離席は、プログラムが時刻どおりに足す
            SaveState(date, active[active.Count - 1].Time, r.Ok);
            return r;
        }

        // 今の日報（A）と作り直した案（B）を Opus で合成する。B は old フォルダに残す（A は書き換えるときに Backup で残る）。合成できなければ今の日報はそのまま
        static Result Combine(string date, Settings s, Result r, string existing, string header, string body, List<Row> active)
        {
            string path = Core.ReportPath(date);
            SaveOld(date, "作り直し", header + "\r\n\r\n" + body.Trim() + "\r\n");
            // 一字も変えずに残してほしい行: 離席メモ・手で足した作業・休憩（今の日報にあるものだけ）
            var keep = new List<string>(Afk.Load(date).Select(e => e.Line).Where(x => x.Length > 0));
            keep.AddRange(LoadManual(date).Select(x => x.Value));
            if (BreakOn(date, s)) keep.Add(BreakLine(s));
            List<string> now = SplitLines(existing).Select(l => Unnumber(l).Trim()).ToList();
            keep = keep.Where(k => now.Contains(k)).Distinct().ToList();

            string prompt =
                "同じ日の日報が2つあります。2つを比べて、1つの日報に合成してください。\n" +
                "A: 今の日報。利用者が手で直した行、PC から離れていたときの作業、手で足した作業が入っている\n" +
                "B: 作業ログから今あらためて作った日報\n\n" +
                "条件:\n" +
                "- A の行は利用者が確かめた内容なので、文言を変えずに残す。B にない作業も消さない（PC の外の作業のことがある）\n" +
                "- B にだけある作業は足す\n" +
                "- 同じ作業が A と B の両方にあるときは1行にまとめ、文言は A のまま。時刻は、A の時刻が作業ログと明らかに食い違うときだけ作業ログに合わせる\n" +
                "- 工程の時間が重ならないようにする（重なるときは A を優先し、B の行を省くか時刻を縮める）\n" +
                (keep.Count > 0 ? "- 次の行は一字も変えずに残す:\n" + string.Concat(keep.Select(k => "  " + k + "\n")) : "") +
                BreakRule(s, date) +
                "- 形は「■工程」の下に「・HH:MM〜HH:MM 作業内容」、「■軽作業」の下に「・HH:MM 作業内容」、それぞれ時刻の早い順。行頭の [番号] は付けない\n" +
                "- 出力は合成した日報の本文だけ（前置き・説明・見出しの日付行は不要）\n\n" +
                NotesPrompt(date) +
                "A（今の日報）:\n" + Unnumber(StripHeader(existing)) + "\n\n" +
                "B（作り直した日報）:\n" + Unnumber(StripHeader(body)) + "\n\n" +
                "参考: この日の作業ログ（離席時間は除いています）\n" + LogText(active, s);

            string merged;
            try
            {
                merged = StripHeader(ClaudeCli.Ask(prompt, s, date, "（合成）", s.FixModel));
                if (!SplitLines(merged).Any(l => l.Trim().StartsWith("■"))) throw new Exception("AI が日報の形で返しませんでした");
            }
            catch (NeedLoginException ex)
            {
                r.Ok = false;
                r.NeedLogin = true;
                r.Message = ex.Message + "。今の日報はそのままです（作り直した案は old フォルダ）";
                return r;
            }
            catch (Exception ex)
            {
                r.Ok = false;
                r.Message = "合成できませんでした: " + ex.Message + "。今の日報はそのままです（作り直した案は old フォルダ）";
                Core.Error("合成 " + date, ex.Message);
                return r;
            }

            lock (reportLock)
            {
                // AI を待つあいだに手で直された・離席メモが足されたときは、上書きせずに案を old に残す
                if (!File.Exists(path) || Norm(Core.ReadText(path)) != Norm(existing))
                {
                    SaveOld(date, "合成", header + "\r\n\r\n" + merged + "\r\n");
                    r.Ok = false;
                    r.Message = "合成しているあいだに日報が直されたので、今の日報はそのままです（合成した案は old フォルダ）";
                    return r;
                }
                Backup(path);
                List<string> lines = SplitLines(header + "\n\n" + merged);
                EnsureBreak(lines, date, s);
                Core.WriteText(path, Number(string.Join("\n", lines).TrimEnd() + "\n"));
            }
            ApplyAllAfk(date);
            SaveState(date, active[active.Count - 1].Time, true);
            r.Ok = true;
            r.Message = "今の日報と作り直した案を合成しました（合成前の2つは old フォルダ）";
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
                "- 時刻は作業ログに出てくる時刻を使う。作業内容は20文字程度で、日報の書き方に合わせる\n" +
                "- 推定できないもの、私用と思われるものは省く\n" +
                "- 「時刻のない本人のメモ」のうち日報にまだないものがあれば、それも入れる\n" +
                BreakRule(s, date) +
                "- 新しい作業がなければ「・なし」とだけ書く\n\n" +
                NotesPrompt(date) + Afk.Prompt(date) +
                "今の日報:\n" + current.Replace("\r\n", "\n") + "\n\n" +
                since.ToString("HH:mm") + " より後の作業ログ:\n" + LogText(fresh, s);

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
                if (lines.Count > 0 && lines[0].StartsWith("日報 ")) lines[0] = Header(date, all);
                Core.WriteText(path, Number(string.Join("\n", lines).TrimEnd() + "\n"));
            }
            SaveState(date, all[all.Count - 1].Time, true);
            r.Ok = true;
            r.Message = n > 0 ? n + "件を追記しました" : "追記する作業はありませんでした";
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
                foreach (string f in Directory.GetFiles(Core.Dir, "*.csv"))
                    if (IsDate(Path.GetFileNameWithoutExtension(f)) && File.GetLastWriteTime(f) < limit) File.Delete(f);
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
            try { Core.WriteText(PathOf, (off ? "off" : "on") + "\t" + at.ToString("yyyy-MM-dd HH:mm:ss") + "\r\n"); }
            catch (Exception ex) { Core.Error("勤務", ex.Message); }
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
        public static Release Latest()
        {
            System.Net.HttpWebRequest req = Request("https://api.github.com/repos/" + Repo + "/releases/latest");
            req.Accept = "application/vnd.github+json";
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
            CheckAi();
            if (updated) Balloon("作業ログを " + Core.Version + " に更新しました", "変わったところは日報画面のメニュー「ヘルプ → バージョン情報」で見られます", ToolTipIcon.Info, () => ShowForm(null));
            if (firstRun && !Core.Test) AskStartup();
        }

        // 初めての起動: スタートアップに登録するか聞く（あとからトレイのメニューか設定のタブ「起動」で変えられる）
        void AskStartup()
        {
            try
            {
                if (StartupPath() != null) return;
                if (MessageBox.Show("作業ログを Windows 起動時に自動で起動しますか？\n（作業の記録は、起動しているあいだだけ取れます）\n\n" +
                        "あとからトレイのアイコンを右クリック →「Windows 起動時に自動で起動する」で変えられます。",
                        "作業ログ", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return;
                SetStartup(true);
                Balloon("自動で起動するように登録しました", "登録した場所: " + Application.ExecutablePath, ToolTipIcon.Info, null);
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
            if (checking) { if (manual) Balloon("更新", "更新を確かめている最中です", ToolTipIcon.Info, null); return; }
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
            Balloon("作業ログの新しい版 " + rel.Version + " があります", "クリックすると更新します", ToolTipIcon.Info, () => ConfirmUpdate(rel));
        }

        void ConfirmUpdate(Updater.Release rel)
        {
            string notes = (rel.Notes ?? "").Trim();
            if (notes.Length > 600) notes = notes.Substring(0, 600) + "…";
            if (MessageBox.Show("新しい版 " + rel.Version + " があります（今は " + Updater.Current + "）。\n\n" + (notes.Length > 0 ? notes + "\n\n" : "") +
                    "今すぐ更新しますか？\n作業ログが数秒止まって起動し直します。記録・日報・設定はそのまま引き継がれます。",
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
                if (applyNow) Balloon("更新", "日報を作り終えたら更新します", ToolTipIcon.Info, null);
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
                Balloon("作業ログを更新できませんでした", ex.Message + "\nクリックするとリリースのページを開きます（zip から手で置き換えられます）",
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
            catch (Exception ex) { Core.Error("離席メモ", ex.Message); Balloon("離席メモ", "日報に足せませんでした: " + ex.Message, ToolTipIcon.Warning, null); }
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
            Balloon("出勤しました（" + now.ToString("HH:mm") + "）", "作業の記録を始めました" + (manual ? "" : "（操作があったので出勤とみなしました）"), ToolTipIcon.Info, null);
            CheckAi();
            if (!busy && now >= started.AddMinutes(2)) { lastAuto = now; Run(() => Reporter.Auto(Settings.Get()), false, true); }
        }

        // 退勤: 今日の日報を仕上げて（残りの作業を追記）、記録を止めて待つ
        void ClockOut(DateTime at, bool manual)
        {
            off = true;
            offAt = at;
            Duty.Save(true, at);
            UpdateDuty();
            Balloon("退勤しました（" + at.ToString("HH:mm") + "）", (manual ? "" : "操作がないので退勤とみなしました。") +
                "記録を止めて待機します。次に出勤するときはトレイのメニューか日報画面の「出勤」", ToolTipIcon.Info, null);
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
                    ui.Post(_ => Balloon("AI（Claude Code）を使えません", problem + "\nクリックすると接続先を開きます", ToolTipIcon.Warning, ShowAiForm), null);
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
            if (newDay) CheckAi();   // 1日1回
            lastAuto = now;
            Run(() => Reporter.Auto(s), false, newDay || finish);
        }

        public bool Busy { get { return busy; } }

        public void Generate(string date)
        {
            if (busy) { Balloon("日報", "日報を作っている最中です。少し待ってください", ToolTipIcon.Info, null); return; }
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
                else if (!r.Ok) Balloon("日報 " + d + " をAIで作れませんでした", r.Message + "\nあとで自動でもう一度試します", ToolTipIcon.Warning, () => ShowForm(d));
                else Balloon("日報 " + d + (r.Appended ? " に追記しました" : " ができました"), r.Message ?? "クリックすると開きます", ToolTipIcon.Info, () => ShowForm(d));
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
            using (var bmp = new Bitmap(32, 32))
            {
                using (Graphics g = Graphics.FromImage(bmp))
                using (var path = new GraphicsPath())
                using (var bg = new SolidBrush(back))
                using (var f = new Font("Meiryo UI", 20f, FontStyle.Bold, GraphicsUnit.Pixel))
                using (var sf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center })
                {
                    g.SmoothingMode = SmoothingMode.AntiAlias;
                    g.TextRenderingHint = TextRenderingHint.AntiAlias;
                    int r = 14;
                    path.AddArc(1, 1, r, r, 180, 90);
                    path.AddArc(31 - r, 1, r, r, 270, 90);
                    path.AddArc(31 - r, 31 - r, r, r, 0, 90);
                    path.AddArc(1, 31 - r, r, r, 90, 90);
                    path.CloseFigure();
                    g.FillPath(bg, path);
                    g.DrawString("報", f, Brushes.White, new RectangleF(0, 1, 32, 32), sf);
                }
                return Icon.FromHandle(bmp.GetHicon());
            }
        }
    }

    // 時刻の入力欄。4桁（1430）で 14:30。ホイールか ↑↓ で、左半分（時）なら1時間、右半分（分）なら1分ずつ変わる。空なら時刻なし
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

        protected override void OnMouseWheel(MouseEventArgs e)
        {
            Step(e.Delta > 0 ? 1 : -1, e.X < ClientSize.Width / 2);
            var h = e as HandledMouseEventArgs;
            if (h != null) h.Handled = true;
        }

        protected override void OnKeyDown(KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Up || e.KeyCode == Keys.Down)
            {
                Step(e.KeyCode == Keys.Up ? 1 : -1, Text.Length == 5 && SelectionStart <= 2);
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
        readonly Button copyBullets, copyAll, saveBtn, redoBtn, todayBtn;
        readonly ToolStripMenuItem mSave, mCopyBullets, mCopyAll, mAiFix, mRedo, mToday, mError;

        // 「AIで直す」欄
        readonly Label aiTitle = new Label { AutoSize = true, Margin = new Padding(3, 6, 3, 2) };
        readonly TextBox instruction = new TextBox { Multiline = true, AcceptsReturn = true, ScrollBars = ScrollBars.Vertical, Dock = DockStyle.Fill };
        readonly Button askBtn = new Button { Text = "AIに直してもらう", Dock = DockStyle.Fill, Height = 34, Margin = new Padding(3, 4, 3, 2) };
        readonly Label aiStatus = new Label { AutoSize = true, Margin = new Padding(3, 4, 3, 4) };
        readonly Button acceptBtn = new Button { Text = "確定して保存", AutoSize = true, Margin = new Padding(3, 2, 3, 2) };
        readonly Button revertBtn = new Button { Text = "元に戻す", AutoSize = true, Margin = new Padding(3, 2, 3, 2) };
        readonly FlowLayoutPanel previewBar = new FlowLayoutPanel { AutoSize = true, WrapContents = false, Margin = new Padding(0) };
        readonly TextBox notes = new TextBox { Multiline = true, AcceptsReturn = true, ScrollBars = ScrollBars.Vertical, Dock = DockStyle.Fill };

        // 「離席メモ」欄
        readonly TabPage aiPage = new TabPage("AIで直す"), afkPage = new TabPage("離席メモ");
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
            Item(help, "更新を確認(&U)…", () => app.CheckUpdate(true));
            Item(help, "リリースのページを開く(&R)", () => Process.Start(Updater.Page));
            help.DropDownItems.Add(new ToolStripSeparator());
            Item(help, "バージョン情報(&A)", ShowAbout);
            MainMenuStrip = menu;
            // Dock は後から追加したものが外側になる
            var center = new Panel { Dock = DockStyle.Fill };
            center.Controls.Add(text);
            center.Controls.Add(BuildAddBar());
            Controls.Add(center);
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
            aiPage.BackColor = afkPage.BackColor = PanelBack;
            tabs.TabPages.Add(aiPage);
            tabs.TabPages.Add(afkPage);
            return tabs;
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
            Label hint = lbl("4桁かホイールで時刻を変更。開始だけなら軽作業");
            hint.ForeColor = Color.DimGray;
            row.Controls.Add(hint);
            string boxTip = "4桁で入力（1430 → 14:30）。ホイールか ↑↓ で、左側なら時、右側なら分が変わります";
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
            message.Text = "休憩の行を消したので、この日は休憩なしにしました（下の「休憩」に印を付けると戻ります）";
        }

        public void SetDuty(bool off)
        {
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
            afkGrid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "時間", ReadOnly = true, AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells });
            afkGrid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "何をしていたか", AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill });
            afkGrid.Columns.Add(new DataGridViewCheckBoxColumn { HeaderText = "除く", AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader, ToolTipText = "休憩・私用など、日報に入れないとき" });
            afkGrid.CurrentCellDirtyStateChanged += (o, e) =>
            {
                if (afkGrid.IsCurrentCellDirty && afkGrid.CurrentCell is DataGridViewCheckBoxCell) afkGrid.CommitEdit(DataGridViewDataErrorContexts.Commit);
            };
            afkGrid.CellValueChanged += (o, e) => { if (!loading && e.RowIndex >= 0 && e.RowIndex < afkList.Count) AfkEdited(e.RowIndex); };
            var freeHead = new Label
            {
                AutoSize = true, MaximumSize = new Size(W - 12, 0), Margin = new Padding(3, 10, 3, 2),
                Text = "時刻のわからないメモ（1行に1つ。離席時間と照らして、AI が時刻を決めます）"
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
            foreach (AfkEntry e in afkList) PaintAfkRow(afkGrid.Rows.Add(e.Range, e.Memo, e.Skip));
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
            afkGrid.Rows[i].DefaultCellStyle.BackColor = afkList[i].Answered ? SystemColors.Window : Color.FromArgb(255, 246, 214);
            afkGrid.Rows[i].Cells[0].ToolTipText = afkList[i].Duration;
        }

        void UpdateAfkTab()
        {
            int n = afkList.Count(e => !e.Answered);
            afkPage.Text = n > 0 ? "離席メモ（" + n + "）" : "離席メモ";
        }

        void AfkEdited(int i)
        {
            AfkEntry e = afkList[i];
            e.Memo = Convert.ToString(afkGrid.Rows[i].Cells[1].Value).Trim();
            e.Skip = Convert.ToBoolean(afkGrid.Rows[i].Cells[2].Value);
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
                if (date == current) message.Text = "日報に追記がありました。AI の案を確定するときに確認が出ます";
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

        Control BuildAiPanel()
        {
            var p = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, Padding = new Padding(4, 2, 2, 2), BackColor = PanelBack };
            p.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            aiTitle.Font = new Font(Font, FontStyle.Bold);
            aiStatus.MaximumSize = new Size(W - 16, 0);
            var hint = new Label
            {
                AutoSize = true, MaximumSize = new Size(W - 16, 0), ForeColor = Color.DimGray, Margin = new Padding(3, 0, 3, 4),
                Text = "違っているところと直し方を書いてください。行頭の番号で指せます。\n例: 3番は△△橋の図面確認"
            };
            var notesHead = new Label
            {
                AutoSize = true, MaximumSize = new Size(W - 16, 0), Margin = new Padding(3, 10, 3, 2),
                Text = "この日の訂正の記録（作り直すときにも AI に渡します。要らない行は消してください）"
            };
            askBtn.Font = new Font(Font, FontStyle.Bold);
            askBtn.Click += (o, e) => AskAi();
            acceptBtn.Click += (o, e) => Accept();
            revertBtn.Click += (o, e) => Revert();
            previewBar.Controls.Add(acceptBtn);
            previewBar.Controls.Add(revertBtn);
            instruction.TextChanged += (o, e) => UpdateButtons();
            notes.TextChanged += (o, e) => { if (!loading) notesDirty = true; };

            Control[] rows = { aiTitle, hint, instruction, askBtn, aiStatus, previewBar, notesHead, notes };
            foreach (Control c in rows)
            {
                RowStyle st = c == instruction ? new RowStyle(SizeType.Absolute, 110)
                            : c == notes ? new RowStyle(SizeType.Percent, 100) : new RowStyle(SizeType.AutoSize);
                p.RowStyles.Add(st);
                p.Controls.Add(c, 0, p.RowCount++);
            }
            return p;
        }

        static Button AddButton(FlowLayoutPanel bar, string label, EventHandler click)
        {
            var b = new Button { Text = label, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Padding = new Padding(8, 2, 8, 2), Margin = new Padding(3) };
            b.Click += click;
            bar.Controls.Add(b);
            return b;
        }

        void ShowAbout()
        {
            MessageBox.Show(this,
                "作業ログ・日報　バージョン " + Core.Version + (Core.Test ? "（テスト）" : "") + "\n" +
                "このソフトは AI（Anthropic の Claude）が作成しました。無保証です。\n\n" +
                "実行ファイル: " + Application.ExecutablePath + "\n" +
                "保存先: " + Core.Dir + "\n\n" +
                "更新履歴:\n" + Core.History,
                "バージョン情報", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
            todayBtn.Enabled = !busy;
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
            LoadAfk();
            LoadBreak();
            Text = "日報 サンプル" + (Core.Test ? "（テスト）" : "");
            if (!busy) message.Text = "見本を表示しています（左の日付を選ぶと戻ります）";
            SetAiStatus("", false);
            UpdateButtons();
        }

        const string BusyText = "日報を作っています…（1分ほどかかることがあります）";

        public void SetBusy(bool b)
        {
            busy = b;
            if (b) message.Text = BusyText;
            else if (message.Text == BusyText) message.Text = "";
            UpdateButtons();
        }

        public void SetRecordStatus(string s) { record.Text = s; }

        public void RefreshList(string select)
        {
            SaveIfDirty();
            List<string> reports = Directory.Exists(Core.Dir)
                ? Directory.GetFiles(Core.Dir, "日報_*.txt").Select(f => Path.GetFileNameWithoutExtension(f).Substring(3))
                    .Where(Reporter.IsDate).ToList()
                : new List<string>();
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
                Text = "日報" + (Core.Test ? "（テスト）" : "");
                UpdateButtons();
            }
        }

        void ShowDate(string date)
        {
            SaveIfDirty();
            current = date;
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
                DialogResult r = MessageBox.Show(this, "AI の案をまだ確定していません。\n\nはい: 確定して保存する\nいいえ: 元に戻す", "日報",
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
                    message.Text = "日報がほかで書き換えられていたので、手で直した内容は old フォルダに残しました";
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

        // 右の欄の指示で、AI に今の本文を直してもらう。2回目からは案をもとに直す
        void AskAi()
        {
            string inst = instruction.Text.Trim();
            if (current == null || !hasReport || aiRunning || inst.Length == 0) return;
            if (!Previewing) SaveQuiet();   // 手で直した分は先に保存しておく
            string date = current, sent = Body(), noteText = notes.Text;
            aiRunning = true;
            UpdateButtons();
            SetAiStatus("AI に頼んでいます…（1分ほどかかることがあります。そのあいだも本文は直せます）", false);
            var t = new Thread(() =>
            {
                string result = null, error = null;
                try { result = Reporter.Revise(date, sent, inst, noteText, Settings.Get()); }
                catch (NeedLoginException ex) { error = ex.Message + "。メニュー「ツール → AIの接続先」で確かめてください"; }
                catch (Exception ex) { error = "AI で直せませんでした: " + ex.Message; Core.Error("AIで直す " + date, ex.Message); }
                if (IsDisposed) return;
                BeginInvoke((Action)(() => AiDone(date, sent, inst, result, error)));
            });
            t.IsBackground = true;
            t.Start();
        }

        void AiDone(string date, string sent, string inst, string result, string error)
        {
            aiRunning = false;
            if (error != null) { SetAiStatus(error, true); UpdateButtons(); return; }
            if (date != current) { SetAiStatus(date + " の案ができましたが、日付を切り替えたので使いませんでした", true); UpdateButtons(); return; }
            if (Norm(Body()) != Norm(sent) &&
                MessageBox.Show(this, "AI に頼んだあとに本文を手で直しています。\nAI の案に置き換えますか？（いいえ: 案を使わない）", "日報",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
            {
                SetAiStatus("AI の案は使いませんでした", false);
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
            // 指示を訂正の記録に足す（確定すると保存。元に戻すと消える）
            string line = "・" + Reporter.ResolveNumbers(inst, sent).Replace("\r\n", " ").Replace("\n", " ");
            if (!notes.Text.Contains(line)) notes.Text = (notes.Text.TrimEnd() + "\r\n" + line).Trim();
            instruction.Text = "";
            SetAiStatus(Norm(result) == Norm(previewBase)
                ? "AI の案は元の日報と同じでした。書き方を変えてもう一度頼んでください"
                : "青字が AI の直した行です。本文をそのまま直してもかまいません。よければ「確定して保存」", false);
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
                MessageBox.Show(this, "AI に頼んだあとに、日報が別に書き換えられています（自動の作り直しなど）。\nそれでも置き換えますか？", "日報",
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
                SetAiStatus("確定して保存しました（前の版は old フォルダ）", false);
                message.Text = "AI で直した日報を保存しました（" + DateTime.Now.ToString("HH:mm") + "）";
                BreakRemoved(before, Body());
                UpdateButtons();
                return true;
            }
            catch (Exception ex) { SetAiStatus("保存できませんでした: " + ex.Message, true); return false; }
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
            SetAiStatus("AI の案を取り消して、元に戻しました", false);
            UpdateButtons();
        }

        void Redo(string date)
        {
            if (date == null || app.Busy) return;
            // まだ日報の形になっていない日（AI に失敗した日など）は、合成せずに作り直すだけなので聞かない
            if (File.Exists(Core.ReportPath(date)) && Core.ReadText(Core.ReportPath(date)).Contains("■") &&
                MessageBox.Show(this, date + " の日報を作り直しますか？\n\n作業ログから作り直した案と今の日報を、" + Settings.Get().FixModel + " で比べて合成します" +
                    "（手で直した行・離席メモ・手で足した作業は残します）。\n合成前の2つは old フォルダに残します。", "日報",
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
            note.Text = "書いた内容はこの時間の作業として、すぐ日報に足します。あとで日報画面の「離席メモ」でも書けます";
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
    class SettingsForm : Form
    {
        static readonly string[] Models = { "claude-haiku-5-5", "claude-sonnet-5-5", "claude-opus-5-5" };

        readonly Action saved;
        readonly DateTimePicker reportTime = new DateTimePicker { Format = DateTimePickerFormat.Custom, CustomFormat = "HH:mm", ShowUpDown = true, Width = 90 };
        readonly NumericUpDown idle = Num(30, 3600), minMinutes = Num(0, 60), minDay = Num(0, 600), redo = Num(1, 600), catchUp = Num(1, 60), keep = Num(7, 3650);
        readonly CheckBox openWindow = new CheckBox { Text = "日報ができたら日報画面を開く（外すと通知だけ）", AutoSize = true, Margin = new Padding(3, 10, 3, 3) };
        readonly CheckBox startup = new CheckBox { Text = "Windows 起動時に自動で起動する", AutoSize = true, Margin = new Padding(3, 10, 3, 3) };
        readonly CheckBox autoUpdate = new CheckBox { Text = "新しい版が出たら自動で更新する（外すと知らせるだけ）", AutoSize = true, Margin = new Padding(3, 10, 3, 3) };
        readonly CheckBox afkAsk = new CheckBox { Text = "離席から戻ったときに、何をしていたか聞く（外すと日報画面の「離席メモ」に書くだけ）", AutoSize = true, Margin = new Padding(3, 10, 3, 3) };
        readonly NumericUpDown afkMinutes = Num(0, 240), appendMinutes = Num(0, 480), offIdle = Num(0, 600);
        readonly DateTimePicker workStart = new DateTimePicker { Format = DateTimePickerFormat.Custom, CustomFormat = "HH:mm", ShowUpDown = true, Width = 90 };
        readonly CheckBox breakOn = new CheckBox { Text = "休憩時間", AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(3, 6, 6, 4) };
        readonly DateTimePicker breakStart = new DateTimePicker { Format = DateTimePickerFormat.Custom, CustomFormat = "HH:mm", ShowUpDown = true, Width = 80 };
        readonly DateTimePicker breakEnd = new DateTimePicker { Format = DateTimePickerFormat.Custom, CustomFormat = "HH:mm", ShowUpDown = true, Width = 80 };
        readonly TextBox exclude = Lines(), words = Lines();
        readonly DataGridView mask = new DataGridView();
        readonly ComboBox model = new ComboBox { Width = 260 }, fixModel = new ComboBox { Width = 260 };
        readonly TextBox claudePath = new TextBox { Width = 360 };

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
            Add(t, Auto(), Hint("日報に休憩の行を入れ、休憩時間の PC 操作は日報に入れません。この時間の離席は休憩として記録し、何をしていたか聞きません。\r\n日によって休憩なしにするときは、日報画面の下の「休憩」の印を外します"));
            Add(t, Auto(), L("自動で作るのは、操作時間が"), minDay, L("分以上の日だけ"));
            Add(t, Auto(), L("日報を作ったあとに"), redo, L("分以上作業したら追記する"));
            Add(t, Auto(), L("未作成の日報をさかのぼって探す"), catchUp, L("日"));
            Add(t, Auto(), openWindow);
            Add(t, Auto(), Hint("追記では、今ある行（手で直した行も）は変えません。「作り直す」は、作り直した案と今の日報を「AIで直す」のモデルで合成します（前の2つは old フォルダ）"));

            t = Page(tabs, "出勤・退勤", 3);
            Add(t, Auto(), L("退勤中でも、この時刻を過ぎて操作があれば出勤"), workStart, L("（記録と AI を始める）"));
            Add(t, Auto(), L("仕上げの時刻を過ぎて、操作が"), offIdle, L("分なかったら退勤（0 で自動で退勤しない）"));
            Add(t, Auto(), Hint("仕上げの時刻は「日報」の欄の「今日の日報を仕上げて知らせる時刻」です。\r\n" +
                                "退勤すると今日の日報を仕上げ（残りの作業を追記）、記録と AI を止めて待機します。\r\n" +
                                "トレイのメニューか日報画面の「出勤」「退勤」でも切り替えられます。\r\n" +
                                "退勤を押したあと、その日のうちは自動では出勤しません（次の日の朝に操作があれば出勤）"));

            t = Page(tabs, "記録・集計", 3);
            Add(t, Auto(), L("無操作がこの時間続いたら離席扱い"), idle, L("秒"));
            Add(t, Auto(), L("集計しない短い作業"), minMinutes, L("分未満"));
            Add(t, Auto(), L("ログと古い日報を残す"), keep, L("日（これより古いものは削除）"));
            Add(t, Auto(), L("入力がこの時間なかったら離席として記録"), afkMinutes, L("分（0 で記録しない）"));
            Add(t, Auto(), afkAsk);
            Add(t, Auto(), Hint("記録は " + Core.IntervalSec + " 秒ごとです。保存先: " + Core.Dir));

            t = Page(tabs, "除外・伏せ字", 1);
            Add(t, Auto(), L("集計しないプロセス（1行に1つ。例: explorer）"));
            Add(t, new RowStyle(SizeType.Percent, 28), exclude);
            Add(t, Auto(), L("タイトルを伏せるプロセス（ウィンドウのタイトルを、右の言葉に置き換えて AI に渡す）"));
            Add(t, new RowStyle(SizeType.Percent, 40), mask);
            Add(t, Auto(), L("タイトルの中で伏せる語（1行に1つ。例: 〇〇建設）"));
            Add(t, new RowStyle(SizeType.Percent, 32), words);

            t = Page(tabs, "AI", 3);
            Add(t, Auto(), L("日報を作るモデル（ふだん）"), model);
            Add(t, Auto(), Hint("毎日の自動作成に使います。claude-haiku-5-5 がいちばん利用枠を使いません"));
            Add(t, Auto(), L("「AIで直す」のモデル（高度）"), fixModel);
            Add(t, Auto(), Hint("違いを直すときだけ使います。claude-opus-5-5 は賢いぶん利用枠を多く使います。略称ではなく正式な名前で"));
            var browse = new Button { Text = "参照…", AutoSize = true };
            browse.Click += (o, e) => Browse();
            Add(t, Auto(), L("Claude Code の場所"), claudePath, browse);
            Add(t, Auto(), Hint(@"空のときは PATH と %USERPROFILE%\.local\bin から自動で探します"));
            var ai = new Button { Text = "AIの接続先を確かめる…", AutoSize = true, Margin = new Padding(3, 12, 3, 3) };
            ai.Click += (o, e) => showAi();
            Add(t, Auto(), ai);

            t = Page(tabs, "起動", 1);
            Add(t, Auto(), startup);
            Add(t, Auto(), Hint(Core.Test ? "テスト中は登録できません" : "登録する場所: " + Application.ExecutablePath +
                                             "\r\nexe を別の場所へ動かしたら、いったん外してから付け直してください" +
                                             "\r\nトレイのアイコンを右クリックしたメニューでも登録・解除できます"));
            Add(t, Auto(), autoUpdate);
            Add(t, Auto(), Hint("12時間ごとに GitHub のリリースを確かめます。新しい版は、日報画面を閉じているときに入れ替えて起動し直します" +
                                "\r\n（記録・日報・設定はそのまま）。いつでも日報画面のメニュー「ヘルプ → 更新を確認」で確かめられます"));

            foreach (string m in Models) { model.Items.Add(m); fixModel.Items.Add(m); }
            mask.Columns.Add("proc", "プロセス名（例: OUTLOOK）");
            mask.Columns.Add("word", "置き換える言葉（例: メール対応）");
            mask.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
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

        static Label L(string s) { return new Label { Text = s, AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(3, 8, 6, 4) }; }

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

        // 1行足す。部品が1つなら全部の列にまたがる
        static void Add(TableLayoutPanel t, RowStyle style, params Control[] cs)
        {
            int r = t.RowCount++;
            t.RowStyles.Add(style);
            for (int i = 0; i < cs.Length; i++) t.Controls.Add(cs[i], i, r);
            if (cs.Length == 1) t.SetColumnSpan(cs[0], t.ColumnCount);
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
            afkAsk.Checked = s.AfkAsk;
            openWindow.Checked = s.OpenWindowOnReport;
            exclude.Text = string.Join("\r\n", s.ExcludeProc);
            words.Text = string.Join("\r\n", s.MaskWords);
            mask.Rows.Clear();
            foreach (KeyValuePair<string, string> kv in s.MaskProc) mask.Rows.Add(kv.Key, kv.Value);
            model.Text = s.Model;
            fixModel.Text = s.FixModel;
            claudePath.Text = s.ClaudePath;
            startup.Checked = TrayApp.StartupPath() != null;
            autoUpdate.Checked = s.AutoUpdate;
        }

        static List<string> Split(TextBox box)
        {
            return box.Text.Split('\n').Select(x => x.Trim()).Where(x => x.Length > 0).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
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
                ExcludeProc = new HashSet<string>(Split(exclude), StringComparer.OrdinalIgnoreCase),
                MaskProc = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase),
                MaskWords = Split(words),
                OpenWindowOnReport = openWindow.Checked,
                AfkMinutes = (int)afkMinutes.Value, AfkAsk = afkAsk.Checked,
                AppendMinutes = (int)appendMinutes.Value,
                WorkStart = new TimeSpan(workStart.Value.Hour, workStart.Value.Minute, 0),
                OffIdleMinutes = (int)offIdle.Value,
                AutoUpdate = autoUpdate.Checked,
                BreakStart = breakOn.Checked ? breakStart.Value.TimeOfDay - TimeSpan.FromSeconds(breakStart.Value.Second) : TimeSpan.Zero,
                BreakEnd = breakOn.Checked ? breakEnd.Value.TimeOfDay - TimeSpan.FromSeconds(breakEnd.Value.Second) : TimeSpan.Zero
            };
            foreach (DataGridViewRow row in mask.Rows)
            {
                if (row.IsNewRow) continue;
                string p = Convert.ToString(row.Cells[0].Value).Trim(), w = Convert.ToString(row.Cells[1].Value).Trim();
                if (p.Length == 0 && w.Length == 0) continue;
                if (p.Length == 0 || w.Length == 0) return Fail("「タイトルを伏せるプロセス」は、プロセス名と置き換える言葉の両方を入れてください。");
                if (Bad(p, true) || Bad(w, false)) return Fail("「タイトルを伏せるプロセス」に「,」「、」「:」は使えません: " + p + " / " + w);
                s.MaskProc[p] = w;
            }
            if (breakOn.Checked && !s.HasBreak) return Fail("休憩時間は、終わりを始めより後にしてください。");
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

            var label = new Label { Text = "■ 最近の呼び出し（新しい順。日報1日分ごとに、元のログ・使ったログイン・モデル・トークン数）", Dock = DockStyle.Top, Height = 28, TextAlign = ContentAlignment.BottomLeft };

            history.View = View.Details;
            history.FullRowSelect = true;
            history.Dock = DockStyle.Fill;
            foreach (string[] c in new[] {
                new[] { "日時", "155" }, new[] { "日報の日付", "100" }, new[] { "元のログ", "120" }, new[] { "ログイン", "290" },
                new[] { "モデル", "150" }, new[] { "トークン", "80" }, new[] { "結果", "400" } })
                history.Columns.Add(c[0], int.Parse(c[1]), c[0] == "トークン" ? HorizontalAlignment.Right : HorizontalAlignment.Left);

            var bar = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = 50, Padding = new Padding(6, 6, 6, 4), WrapContents = false };
            recheck = AddButton(bar, "再確認", (o, e) => Recheck(true));
            AddButton(bar, "ログイン（ブラウザが開きます）", (o, e) => Login());
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
                sb.AppendLine("　探し方: " + (s.ClaudePath.Length > 0 ? "settings.ini の ClaudePath" : "PATH と %USERPROFILE%\\.local\\bin から自動"));
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
            if (exe == null) { info.Text = "Claude Code（claude.exe）が見つかりません。"; return; }
            try
            {
                Process.Start(new ProcessStartInfo("cmd.exe", "/k \"\"" + exe + "\" auth login\"") { UseShellExecute = true });
                info.Text = "黒い画面とブラウザでログインしてください。終わったら黒い画面を閉じて「再確認」を押してください。";
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
                var it = new ListViewItem(new[] { p[0], p[1], p[2], p[4], p[5], long.TryParse(p[6], out n) ? n.ToString("N0") : p[6], p[7] });
                it.ToolTipText = "Claude Code: " + p[3];
                if (!p[7].StartsWith("成功")) it.ForeColor = Color.Firebrick;
                history.Items.Add(it);
            }
            history.ShowItemToolTips = true;
            history.EndUpdate();
        }
    }

    static class Program
    {
        [STAThread]
        static int Main(string[] args)
        {
            try { Directory.CreateDirectory(Core.Dir); }
            catch (Exception ex) { MessageBox.Show("保存先を作れません: " + Core.Dir + "\n" + ex.Message, "作業ログ"); return 2; }

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
