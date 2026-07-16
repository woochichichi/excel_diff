using System;
using System.IO;
using System.Reflection;
using System.Text;

namespace ExcelDiffMerge
{
    /// <summary>
    /// 파일 로거 (폐쇄망 진단 필수). 디버거 없이 현장에서 동작/오류를 추적한다.
    ///  - exe 옆 ExcelDiffMerge.log 에 append. 쓰기 불가면 %TEMP% 로 폴백.
    ///  - COM 오류는 HRESULT 까지 기록(§11 DRM 예외 매핑에 유용).
    ///  - 2MB 초과 시 .1 로 1회 롤오버.
    /// 스레드 안전(lock) — STA 워커/UI 양쪽에서 호출 가능.
    /// </summary>
    public static class Logger
    {
        private static readonly object _lock = new object();
        private static string _path;
        private static bool _init;
        private const long MaxBytes = 2L * 1024 * 1024;

        public static string LogPath { get { return _path; } }

        public static void Init()
        {
            if (_init) return;
            _init = true;
            try
            {
                string dir = AppDomain.CurrentDomain.BaseDirectory;
                _path = Path.Combine(dir, "ExcelDiffMerge.log");
                File.AppendAllText(_path, "", Encoding.UTF8); // 쓰기 가능 여부 확인
            }
            catch
            {
                try { _path = Path.Combine(Path.GetTempPath(), "ExcelDiffMerge.log"); }
                catch { _path = null; }
            }

            string ver = "?";
            try { ver = Assembly.GetExecutingAssembly().GetName().Version.ToString(); }
            catch { }
            Info("================ ExcelDiffMerge v" + ver + " 시작 ================");
            try
            {
                Info("exe=" + AppDomain.CurrentDomain.BaseDirectory
                   + " | OS=" + Environment.OSVersion
                   + " | CLR=" + Environment.Version
                   + " | 64bit=" + Environment.Is64BitProcess
                   + " | log=" + _path);
            }
            catch { }
        }

        public static void Info(string msg) { Write("INFO", msg); }
        public static void Warn(string msg) { Write("WARN", msg); }
        public static void Error(string msg) { Write("ERROR", msg); }

        public static void Error(string msg, Exception ex)
        {
            Write("ERROR", msg + " :: " + Describe(ex));
        }

        public static string Describe(Exception ex)
        {
            if (ex == null) return "(null exception)";
            StringBuilder sb = new StringBuilder();
            Exception e = ex;
            int depth = 0;
            while (e != null && depth < 5)
            {
                if (depth > 0) sb.Append(" <- ");
                sb.Append(e.GetType().Name).Append(": ").Append(e.Message);
                try { sb.Append(" [HRESULT=0x").Append(e.HResult.ToString("X8")).Append("]"); }
                catch { }
                e = e.InnerException;
                depth++;
            }
            if (ex.StackTrace != null)
                sb.Append("\r\n    ").Append(ex.StackTrace.Replace("\n", "\n    "));
            return sb.ToString();
        }

        private static void Write(string level, string msg)
        {
            if (_path == null) return;
            lock (_lock)
            {
                try
                {
                    Rotate();
                    string line = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")
                                + " [" + level + "] " + msg + "\r\n";
                    File.AppendAllText(_path, line, Encoding.UTF8);
                }
                catch { /* 로깅 실패는 앱을 막지 않는다 */ }
            }
        }

        private static void Rotate()
        {
            try
            {
                FileInfo fi = new FileInfo(_path);
                if (fi.Exists && fi.Length > MaxBytes)
                {
                    string bak = _path + ".1";
                    if (File.Exists(bak)) File.Delete(bak);
                    File.Move(_path, bak);
                }
            }
            catch { }
        }
    }
}
