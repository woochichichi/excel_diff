using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace ExcelDiffMerge
{
    /// <summary>
    /// Word(.doc/.docx) COM late-binding 리더.
    ///  - DRM 은 Word 가 복호화 → 우리는 COM 으로 본문 텍스트만 읽는다(직접 파싱 없음).
    ///  - 문서를 "1열짜리 시트"로 표현: 각 문단 = 한 행(A열). 시트명은 고정 "문서".
    ///    → 기존 DiffEngine/RowAligner/그리드를 그대로 재사용해 문단 단위 diff 가 된다.
    ///  - 본문 전체를 Content.Text 로 '한 번에' 읽어 COM 왕복 최소화.
    /// 1차: 비교/보기 전용(병합 저장은 Excel 만 지원).
    /// </summary>
    public sealed class WordComReader : IWorkbookReader
    {
        private const int msoAutomationSecurityLow = 1;
        private const int msoAutomationSecurityForceDisable = 3;
        private const int wdDoNotSaveChanges = 0;
        private const int wdAlertsNone = 0;

        private dynamic _wd;
        private int _savedAutomationSecurity = msoAutomationSecurityLow;
        private int _pid;

        public WordComReader()
        {
            Type t = Type.GetTypeFromProgID("Word.Application");
            if (t == null)
                throw new InvalidOperationException(
                    "Word COM(ProgID 'Word.Application') 을 찾을 수 없습니다. Word 가 설치되어 있는지 확인하세요.");
            _wd = Activator.CreateInstance(t);
            _wd.Visible = false;
            try { _wd.DisplayAlerts = wdAlertsNone; } catch { }
            try { _savedAutomationSecurity = (int)_wd.AutomationSecurity; } catch { }
            try { _wd.AutomationSecurity = msoAutomationSecurityForceDisable; } catch { }
            _pid = ComProcessGuard.TryGetPid(_wd); // Word 는 Application.Hwnd 미제공 → 0(무해)
            try { Logger.Info("Word COM 인스턴스 생성. Version=" + (string)_wd.Version); }
            catch { Logger.Info("Word COM 인스턴스 생성(버전 조회 실패)"); }
        }

        public WorkbookData LoadWorkbook(string path)
        {
            WorkbookData wb = new WorkbookData(path);
            Logger.Info("Word Open(ReadOnly) 시도: " + path);
            dynamic docs = null;
            dynamic doc = null;
            try
            {
                docs = _wd.Documents;
                // Open(FileName, ConfirmConversions=false, ReadOnly=true, AddToRecentFiles=false)
                try
                {
                    doc = docs.Open(path, false, true, false);
                }
                catch (Exception exOpen)
                {
                    Logger.Error("Word Open 실패: " + path, exOpen);
                    throw;
                }

                string text = "";
                try
                {
                    dynamic content = doc.Content;
                    try { text = (string)content.Text; }
                    finally { Release(ref content); }
                }
                catch { text = ""; }

                SheetData sd = BuildSheet("문서", text);
                wb.Sheets.Add(sd);
                Logger.Info("Word 로드 완료: " + path + " (문단 " + sd.RowCount + "개)");
            }
            finally
            {
                if (doc != null)
                {
                    try { doc.Close(wdDoNotSaveChanges); } catch { }
                    Release(ref doc);
                }
                Release(ref docs);
            }
            return wb;
        }

        /// <summary>본문 텍스트를 문단 단위로 쪼개 1열 시트로 만든다.</summary>
        private static SheetData BuildSheet(string name, string text)
        {
            SheetData sd = new SheetData(name);
            sd.FirstRow = 1;
            sd.FirstCol = 1;

            List<string> lines = SplitParagraphs(text);
            sd.RowCount = lines.Count;
            sd.ColCount = lines.Count > 0 ? 1 : 0;
            sd.Values = new object[sd.RowCount, sd.ColCount == 0 ? 1 : sd.ColCount];
            sd.Formulas = new object[sd.RowCount, sd.ColCount == 0 ? 1 : sd.ColCount];
            for (int i = 0; i < lines.Count; i++)
            {
                string s = lines[i];
                sd.Values[i, 0] = (s.Length == 0) ? null : (object)s;
            }
            return sd;
        }

        private static List<string> SplitParagraphs(string text)
        {
            List<string> outp = new List<string>();
            if (text == null) return outp;
            // Word 문단/셀/줄 구분자 정규화: 표셀(U+0007), 수직탭(U+000B), 페이지(U+000C) → \r
            System.Text.StringBuilder norm = new System.Text.StringBuilder(text.Length);
            foreach (char ch in text)
            {
                if (ch == '\u0007' || ch == '\u000B' || ch == '\u000C') norm.Append('\r');
                else norm.Append(ch);
            }
            string[] parts = norm.ToString().Replace("\r\n", "\r").Replace('\n', '\r').Split('\r');
            // 마지막에 흔히 붙는 빈 문단들 제거.
            int end = parts.Length;
            while (end > 0 && parts[end - 1].Length == 0) end--;
            for (int i = 0; i < end; i++) outp.Add(parts[i]);
            return outp;
        }

        private static void Release(ref dynamic o)
        {
            object obj = (object)o;
            o = null;
            if (obj == null) return;
            try { if (Marshal.IsComObject(obj)) Marshal.ReleaseComObject(obj); }
            catch { }
        }

        public void Dispose()
        {
            object app = (object)_wd;
            _wd = null;
            try { if (app != null) ((dynamic)app).AutomationSecurity = _savedAutomationSecurity; } catch { }
            if (app != null)
            {
                try { ((dynamic)app).Quit(wdDoNotSaveChanges); } catch { }
                try { if (Marshal.IsComObject(app)) Marshal.ReleaseComObject(app); } catch { }
            }
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
            ComProcessGuard.KillIfAlive(_pid, "WINWORD");
            Logger.Info("Word COM 인스턴스 정리 완료(Quit+Release+GC)");
        }
    }
}
