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
    public sealed class ExcelComReader : IWorkbookReader
    {
        // xlCalculation
        private const int xlCalculationManual = -4135;
        private const int xlCalculationAutomatic = -4105;
        // msoAutomationSecurity
        private const int msoAutomationSecurityLow = 1;
        private const int msoAutomationSecurityForceDisable = 3;

        private dynamic _xl;              // Excel.Application (우리 전용 인스턴스)
        private int _savedCalculation = xlCalculationAutomatic;
        private int _savedAutomationSecurity = msoAutomationSecurityLow;
        private bool _optionsApplied;
        private int _pid;                 // 우리 Excel 인스턴스 PID(좀비 방지용)

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
            try { _xl.Interactive = false; } catch { }
            // 매크로 자동실행 차단 (spec §7). AutomationSecurity 는 프로세스 전역이라 원복 대비 저장.
            try { _savedAutomationSecurity = (int)_xl.AutomationSecurity; } catch { }
            try { _xl.AutomationSecurity = msoAutomationSecurityForceDisable; } catch { }
            _pid = ComProcessGuard.TryGetPid(_xl);
            try { Logger.Info("Excel COM 인스턴스 생성. Version=" + (string)_xl.Version + " pid=" + _pid); }
            catch { Logger.Info("Excel COM 인스턴스 생성(버전 조회 실패) pid=" + _pid); }
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
            Logger.Info("COM Open(ReadOnly) 시도: " + path);
            try
            {
                workbooks = _xl.Workbooks;
                // Open positional (named arg 미지원 csc 대비):
                //  1 Filename, 2 UpdateLinks=0, 3 ReadOnly=true, 7 IgnoreReadOnlyRecommended=true, 11 Notify=false
                // → 이미 열린 파일/읽기전용 권장 프롬프트를 선제 차단(조사 반영).
                try
                {
                    wb = workbooks.Open(path, 0, true,
                        Type.Missing, Type.Missing, Type.Missing, true,
                        Type.Missing, Type.Missing, Type.Missing, false);
                }
                catch (Exception exOpen)
                {
                    // DRM 권한없음/잠김 등 → HRESULT 포함 로깅 후 상위로(§7, §11).
                    Logger.Error("COM Open 실패: " + path, exOpen);
                    // 대표 HRESULT 를 한국어 안내로 매핑해 예외 메시지 앞부분에 실어
                    // 기존 UI(MainForm.DescribeError)에도 그대로 표시되게 한다. HRESULT 로깅은 위에서 유지.
                    string ko = DescribeComError(exOpen);
                    if (ko != null)
                        throw new InvalidOperationException(ko + "\r\n(원본 오류: " + exOpen.Message + ")", exOpen);
                    throw;
                }

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
                Logger.Info("로드 완료: " + path + " (시트 " + wbData.Sheets.Count + "개)");
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
                // two-dot 체인 회피: Rows/Columns 를 지역변수로 받아 해제(COM 누수 최소화).
                dynamic usedRows = used.Rows;
                dynamic usedCols = used.Columns;
                int rowCount = (int)usedRows.Count;
                int colCount = (int)usedCols.Count;
                Release(ref usedRows);
                Release(ref usedCols);

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

        /// <summary>
        /// COM Open 실패 예외의 HRESULT 를 대표 케이스별 한국어 안내로 매핑한다.
        /// 매칭되는 케이스가 없으면 null 을 돌려 호출부가 원본 예외를 그대로 던지게 한다.
        /// </summary>
        private static string DescribeComError(Exception ex)
        {
            COMException ce = ex as COMException;
            if (ce == null) return null;
            int hr = ce.ErrorCode;
            // E_ACCESSDENIED
            if (hr == unchecked((int)0x80070005))
                return "DRM 열람 권한이 없거나 접근이 거부된 파일입니다.";
            // Excel 자동화 일반 오류(파일 열기 실패/잠김 등)
            if (hr == unchecked((int)0x800A03EC))
                return "파일이 다른 곳(Excel)에서 열려 있거나 열 수 없는 파일입니다.";
            // RPC_E_INVALID_OBJECT / RPC_E_DISCONNECTED
            if (hr == unchecked((int)0x80010114) || hr == unchecked((int)0x80010108))
                return "Excel COM 연결이 끊어졌습니다(Excel 재시작 후 재시도).";
            return null;
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
            // 중요: dynamic 에 '== null' 같은 동적 연산을 하면, 이미 끊긴 COM 객체(RPC_E_INVALID_OBJECT)
            // 에서 DLR 바인더가 QueryInterface 를 시도하다 예외가 난다.
            // → 먼저 명시적 (object) 참조변환으로 dynamic 을 벗겨낸 뒤 정적 코드로만 처리한다.
            object obj = (object)o;
            o = null;
            if (obj == null) return;
            try
            {
                if (Marshal.IsComObject(obj))
                    Marshal.ReleaseComObject(obj);
            }
            catch { }
        }

        public void Dispose()
        {
            RestoreFastOptions();
            object app = (object)_xl;   // dynamic 벗겨내기(동적 == 회피)
            _xl = null;
            // AutomationSecurity 는 프로세스 전역 → 원복(조사 반영).
            try { if (app != null) ((dynamic)app).AutomationSecurity = _savedAutomationSecurity; } catch { }
            if (app != null)
            {
                try { ((dynamic)app).Quit(); } catch { }
                try
                {
                    if (Marshal.IsComObject(app))
                        Marshal.ReleaseComObject(app);
                }
                catch { }
            }
            // 남은 RCW 강제 수거 → 좀비 EXCEL.EXE 방지 보강.
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
            // 최후 안전망: Quit 후에도 남아있으면 우리 PID 만 강제 종료.
            ComProcessGuard.KillIfAlive(_pid, "EXCEL");
            Logger.Info("Excel COM 인스턴스 정리 완료(Quit+Release+GC)");
        }
    }
}
