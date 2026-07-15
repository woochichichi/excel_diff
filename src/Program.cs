using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace ExcelDiffMerge
{
    internal static class Program
    {
        // winexe 로 빌드해도(§10) 콘솔 모드(--poc/--diff)에서 부모 콘솔에 출력이 보이도록 attach.
        private const int ATTACH_PARENT_PROCESS = -1;
        [DllImport("kernel32.dll")]
        private static extern bool AttachConsole(int dwProcessId);

        private static void EnsureConsole()
        {
            try { AttachConsole(ATTACH_PARENT_PROCESS); } catch { }
        }

        /// <summary>
        /// 진입점.
        ///  - 인자 없음 → WinForms GUI (기본).
        ///  - "--poc &lt;file&gt;" → COM 열기 PoC (spec §9-1): 1개 파일 로드 후 콘솔 요약.
        ///  - "--diff &lt;left&gt; &lt;right&gt;" → 2-way diff 콘솔 출력 (spec §9-2).
        /// COM 은 STA 필요 → [STAThread].
        /// </summary>
        [STAThread]
        private static int Main(string[] args)
        {
            if (args.Length >= 1 && args[0] == "--poc")
            {
                EnsureConsole();
                return RunPoc(args);
            }
            if (args.Length >= 1 && args[0] == "--diff")
            {
                EnsureConsole();
                return RunDiffConsole(args);
            }

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            try
            {
                Application.Run(new MainForm());
            }
            catch (Exception ex)
            {
                MessageBox.Show("치명적 오류: " + ex.Message, "ExcelDiffMerge",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return 1;
            }
            return 0;
        }

        // --------------------------------------------- 콘솔 PoC (§9-1)
        private static int RunPoc(string[] args)
        {
            if (args.Length < 2)
            {
                Console.WriteLine("사용법: ExcelDiffMerge.exe --poc <파일경로>");
                return 2;
            }
            string path = args[1];
            Console.WriteLine("[PoC] COM 열기: " + path);
            try
            {
                using (ExcelComReader reader = new ExcelComReader())
                {
                    WorkbookData wb = reader.LoadWorkbook(path);
                    Console.WriteLine("시트 수: " + wb.Sheets.Count);
                    foreach (SheetData s in wb.Sheets)
                    {
                        Console.WriteLine(string.Format(
                            " - {0}: {1}행 x {2}열 (시작 R{3}C{4})",
                            s.Name, s.RowCount, s.ColCount, s.FirstRow, s.FirstCol));
                        // 앞쪽 몇 셀만 미리보기.
                        int pr = Math.Min(3, s.RowCount);
                        int pc = Math.Min(5, s.ColCount);
                        for (int r = 0; r < pr; r++)
                        {
                            List<string> cells = new List<string>();
                            for (int c = 0; c < pc; c++)
                                cells.Add(DiffEngine.ToText(s.Values[r, c]));
                            Console.WriteLine("     " + string.Join(" | ", cells.ToArray()));
                        }
                    }
                }
                Console.WriteLine("[PoC] 정상 종료. 작업관리자에서 좀비 EXCEL.EXE 없는지 확인하세요.");
                return 0;
            }
            catch (Exception ex)
            {
                Console.WriteLine("[PoC] 오류: " + ex.Message);
                return 1;
            }
        }

        // --------------------------------------------- 콘솔 2-way diff (§9-2)
        private static int RunDiffConsole(string[] args)
        {
            if (args.Length < 3)
            {
                Console.WriteLine("사용법: ExcelDiffMerge.exe --diff <좌파일> <우파일>");
                return 2;
            }
            try
            {
                DiffResult diff;
                using (ExcelComReader reader = new ExcelComReader())
                {
                    WorkbookData l = reader.LoadWorkbook(args[1]);
                    WorkbookData r = reader.LoadWorkbook(args[2]);
                    diff = DiffEngine.Compare(l, r);
                }
                Console.WriteLine("=== DIFF 요약 ===");
                Console.WriteLine(diff.Summary());
                foreach (SheetDiff sd in diff.Sheets)
                {
                    if (!sd.HasChanges) continue;
                    Console.WriteLine(string.Format("[{0}] 변경 {1} / 추가 {2} / 삭제 {3}",
                        sd.Name, sd.ChangedCount, sd.AddedCount, sd.DeletedCount));
                    int shown = 0;
                    foreach (DiffCell dc in sd.Cells)
                    {
                        if (shown++ >= 20) { Console.WriteLine("   … (이하 생략)"); break; }
                        Console.WriteLine(string.Format("   {0}{1} [{2}]  '{3}' -> '{4}'",
                            MainForm.ColLetter(dc.Col), dc.Row, dc.Status,
                            DiffEngine.ToText(dc.OldValue), DiffEngine.ToText(dc.NewValue)));
                    }
                }
                return 0;
            }
            catch (Exception ex)
            {
                Console.WriteLine("[diff] 오류: " + ex.Message);
                return 1;
            }
        }
    }
}
