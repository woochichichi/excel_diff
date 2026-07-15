using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace ExcelDiffMerge
{
    /// <summary>
    /// Excel COM late-binding 리더.
    ///
    /// 설계 원칙(spec §1, §4, §7):
    ///  - DRM 파일은 Excel 이 복호화하므로 반드시 Excel COM 경유로만 값을 읽는다(직접 파싱 금지).
    ///  - ProgID 로 late binding → PIA/Interop DLL 참조 0 (폐쇄망 DLL 반입 문제 소거).
    ///  - UsedRange.Value / .Formula 를 2D 배열로 "한 번에" 로드 (셀 단위 COM 왕복 금지).
    ///  - 우리가 띄운 Excel 인스턴스만 정리. 사용자가 이미 열어둔 Excel 은 건드리지 않음(별도 인스턴스).
    ///  - COM 객체는 역순으로 Marshal.ReleaseComObject, 예외 시에도 finally 로 Quit 보장.
    /// </summary>
    public sealed class ExcelComReader : IDisposable
    {
        // xlCalculation
        private const int xlCalculationManual = -4135;
        private const int xlCalculationAutomatic = -4105;
        // msoAutomationSecurity
        private const int msoAutomationSecurityForceDisable = 3;

        private dynamic _xl;              // Excel.Application (우리 전용 인스턴스)
        private int _savedCalculation = xlCalculationAutomatic;
        private bool _optionsApplied;

        public ExcelComReader()
        {
            Type xlType = Type.GetTypeFromProgID("Excel.Application");
            if (xlType == null)
                throw new InvalidOperationException(
                    "Excel COM(ProgID 'Excel.Application') 을 찾을 수 없습니다. Excel 이 설치되어 있는지 확인하세요.");

            _xl = Activator.CreateInstance(xlType);
            _xl.Visible = false;
            _xl.DisplayAlerts = false;
            _xl.AskToUpdateLinks = false;
            // 매크로 자동실행 차단 (spec §7).
            try { _xl.AutomationSecurity = msoAutomationSecurityForceDisable; } catch { }
        }

        /// <summary>열기 직후 성능 최적화 옵션 적용 (spec §4-3). 작업 끝나면 RestoreExcelOptions 로 원복.</summary>
        private void ApplyFastOptions()
        {
            if (_optionsApplied) return;
            try { _savedCalculation = (int)_xl.Calculation; } catch { _savedCalculation = xlCalculationAutomatic; }
            try { _xl.ScreenUpdating = false; } catch { }
            try { _xl.Calculation = xlCalculationManual; } catch { }
            try { _xl.EnableEvents = false; } catch { }
            _optionsApplied = true;
        }

        private void RestoreFastOptions()
        {
            if (!_optionsApplied) return;
            try { _xl.Calculation = _savedCalculation; } catch { }
            try { _xl.ScreenUpdating = true; } catch { }
            try { _xl.EnableEvents = true; } catch { }
            _optionsApplied = false;
        }

        /// <summary>
        /// 파일 1개를 ReadOnly 로 열어 전체 시트를 일괄 로드한다.
        /// DRM 권한 없음/이미 열림 등은 COMException 으로 상위에 전달 → 사용자 안내(§7).
        /// </summary>
        public WorkbookData LoadWorkbook(string path)
        {
            ApplyFastOptions();

            WorkbookData wbData = new WorkbookData(path);
            dynamic workbooks = null;
            dynamic wb = null;
            try
            {
                workbooks = _xl.Workbooks;
                // Open(Filename, UpdateLinks=0, ReadOnly=true) — dynamic 이므로 optional 인자는 런타임 처리.
                // named argument 미지원 csc 환경 대비: positional 로만 호출.
                wb = workbooks.Open(path, 0, true);

                dynamic sheets = wb.Worksheets;
                int sheetCount = (int)sheets.Count;
                for (int i = 1; i <= sheetCount; i++)
                {
                    dynamic ws = sheets[i];
                    try
                    {
                        SheetData sd = LoadSheet(ws);
                        if (sd != null) wbData.Sheets.Add(sd);
                    }
                    finally
                    {
                        Release(ref ws);
                    }
                }
                Release(ref sheets);
            }
            finally
            {
                if (wb != null)
                {
                    try { wb.Close(false); } catch { }
                    Release(ref wb);
                }
                Release(ref workbooks);
            }
            return wbData;
        }

        /// <summary>
        /// 한 워크시트를 UsedRange.Value / .Formula 로 일괄 로드.
        /// UsedRange 부풀림 함정(§4-4) 대비: 반환된 배열 기준 실제 경계를 재산정한다.
        /// </summary>
        private SheetData LoadSheet(dynamic ws)
        {
            string name = (string)ws.Name;
            SheetData sd = new SheetData(name);

            dynamic used = null;
            try
            {
                used = ws.UsedRange;
                if (used == null) return sd;

                int firstRow = (int)used.Row;
                int firstCol = (int)used.Column;
                int rowCount = (int)used.Rows.Count;
                int colCount = (int)used.Columns.Count;

                if (rowCount <= 0 || colCount <= 0) return sd;

                object rawValues = used.Value;   // 1개 셀이면 스칼라, 그 외엔 object[,] (1-based)
                object rawFormula;
                try { rawFormula = used.Formula; } catch { rawFormula = null; }

                object[,] values = Normalize(rawValues, rowCount, colCount);
                object[,] formulas = Normalize(rawFormula, rowCount, colCount);

                // 실제 경계 재산정: 완전히 빈 뒤쪽 행/열을 잘라 UsedRange 부풀림을 보정.
                int realRows, realCols;
                ComputeRealBounds(values, formulas, rowCount, colCount, out realRows, out realCols);

                sd.FirstRow = firstRow;
                sd.FirstCol = firstCol;
                sd.RowCount = realRows;
                sd.ColCount = realCols;

                if (realRows == rowCount && realCols == colCount)
                {
                    sd.Values = values;
                    sd.Formulas = formulas;
                }
                else
                {
                    sd.Values = Crop(values, realRows, realCols);
                    sd.Formulas = Crop(formulas, realRows, realCols);
                }
            }
            finally
            {
                Release(ref used);
            }
            return sd;
        }

        /// <summary>
        /// COM 이 돌려준 Value/Formula(스칼라 또는 1-based object[,])를 0-based object[rowCount,colCount] 로 정규화.
        /// </summary>
        private static object[,] Normalize(object raw, int rowCount, int colCount)
        {
            object[,] result = new object[rowCount, colCount];
            if (raw == null) return result;

            object[,] arr = raw as object[,];
            if (arr == null)
            {
                // 단일 셀 → 스칼라. [0,0] 에 배치.
                result[0, 0] = raw;
                return result;
            }

            int lr = arr.GetLowerBound(0);
            int lc = arr.GetLowerBound(1);
            int hr = arr.GetUpperBound(0);
            int hc = arr.GetUpperBound(1);

            int rows = Math.Min(rowCount, hr - lr + 1);
            int cols = Math.Min(colCount, hc - lc + 1);
            for (int r = 0; r < rows; r++)
                for (int c = 0; c < cols; c++)
                    result[r, c] = arr[lr + r, lc + c];
            return result;
        }

        /// <summary>값/수식 모두 빈 마지막 행·열을 잘라 실제 경계를 구한다(§4-4).</summary>
        private static void ComputeRealBounds(object[,] values, object[,] formulas,
            int rowCount, int colCount, out int realRows, out int realCols)
        {
            realRows = 0;
            realCols = 0;
            for (int r = 0; r < rowCount; r++)
            {
                for (int c = 0; c < colCount; c++)
                {
                    if (!IsEmpty(values[r, c]) || !IsEmpty(formulas[r, c]))
                    {
                        if (r + 1 > realRows) realRows = r + 1;
                        if (c + 1 > realCols) realCols = c + 1;
                    }
                }
            }
            // 완전히 빈 시트라도 최소 원래 크기를 유지하지 않고 0 으로 둔다(진짜 빈 시트).
        }

        private static object[,] Crop(object[,] src, int rows, int cols)
        {
            object[,] dst = new object[rows, cols];
            for (int r = 0; r < rows; r++)
                for (int c = 0; c < cols; c++)
                    dst[r, c] = src[r, c];
            return dst;
        }

        internal static bool IsEmpty(object v)
        {
            if (v == null) return true;
            string s = v as string;
            if (s != null) return s.Length == 0;
            return false;
        }

        private static void Release(ref dynamic o)
        {
            if (o == null) return;
            object obj = o; // dynamic 디스패치 회피용 캐스팅
            try
            {
                if (obj != null && Marshal.IsComObject(obj))
                    Marshal.ReleaseComObject(obj);
            }
            catch { }
            o = null;
        }

        public void Dispose()
        {
            RestoreFastOptions();
            if (_xl != null)
            {
                try { _xl.Quit(); } catch { }
                object app = _xl;
                try
                {
                    if (app != null && Marshal.IsComObject(app))
                        Marshal.ReleaseComObject(app);
                }
                catch { }
                _xl = null;
            }
            // 남은 RCW 강제 수거 → 좀비 EXCEL.EXE 방지 보강.
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }
}
