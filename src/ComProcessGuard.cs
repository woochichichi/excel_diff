using System;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace ExcelDiffMerge
{
    /// <summary>
    /// 좀비 프로세스(EXCEL.EXE/WINWORD.EXE) 방지 최후 안전망 (Windows COM 하드닝).
    /// ReleaseComObject + Quit + GC 로도 드물게 프로세스가 남을 수 있어, 우리가 만든
    /// 인스턴스의 PID 를 창 핸들(Hwnd)로 확보해 두고, 정리 후에도 살아있으면 강제 종료한다.
    ///  - 반드시 '우리가 생성한' 인스턴스의 PID 만 종료(사용자가 열어둔 Office 는 건드리지 않음).
    ///  - 프로세스 이름이 기대값(EXCEL/WINWORD)과 일치할 때만 종료(오종료 방지).
    /// </summary>
    public static class ComProcessGuard
    {
        [DllImport("user32.dll")]
        private static extern int GetWindowThreadProcessId(IntPtr hWnd, out int processId);

        /// <summary>COM 애플리케이션의 창 핸들(Application.Hwnd)로 PID 조회. 실패 시 0.</summary>
        public static int TryGetPid(dynamic app)
        {
            try
            {
                object hwndObj = (object)app.Hwnd;   // Excel.Application.Hwnd (int). Word 는 없음 → 예외.
                if (hwndObj == null) return 0;
                int hwnd = Convert.ToInt32(hwndObj);
                if (hwnd == 0) return 0;
                int pid;
                GetWindowThreadProcessId(new IntPtr(hwnd), out pid);
                return pid;
            }
            catch { return 0; }
        }

        /// <summary>pid 프로세스가 아직 살아있고 이름이 expectedUpper 를 포함하면 강제 종료.</summary>
        public static void KillIfAlive(int pid, string expectedUpper)
        {
            if (pid <= 0) return;
            try
            {
                Process p = Process.GetProcessById(pid); // 없으면 ArgumentException → catch
                string name = "";
                try { name = p.ProcessName.ToUpperInvariant(); }
                catch { }
                if (name.IndexOf(expectedUpper, StringComparison.Ordinal) >= 0)
                {
                    Logger.Warn("좀비 프로세스 강제 종료: " + name + " (pid=" + pid + ")");
                    p.Kill();
                }
            }
            catch { /* 이미 정상 종료됨 → 정상 */ }
        }
    }
}
