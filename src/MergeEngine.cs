using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;

namespace ExcelDiffMerge
{
    /// <summary>
    /// 병합 대상 셀 하나(절대좌표 1-based + 채택할 값).
    /// </summary>
    public sealed class MergeItem
    {
        public string SheetName;
        public int Row;
        public int Col;
        public object Value;

        public MergeItem(string sheet, int row, int col, object value)
        {
            SheetName = sheet;
            Row = row;
            Col = col;
            Value = value;
        }
    }

    /// <summary>
    /// COM 으로 target 워크북에 선택 셀을 써서 저장한다 (spec §4-2, §5-3, §7).
    ///  - 쓰기는 셀 단위 금지 → 연속 범위로 묶어 batch 쓰기(Range.Value = object[,]).
    ///  - 저장 전 백업 복사본 생성(옵션). 저장 실패 시 원본 보존.
    ///  - ReadOnly 로 열지 않는다(편집·저장 대상). DRM 편집권한 없으면 저장 예외 → 상위 안내.
    /// </summary>
    public sealed class MergeEngine : IDisposable
    {
        private const int xlCalculationManual = -4135;
        private const int msoAutomationSecurityForceDisable = 3;

        private dynamic _xl;
        private int _pid;

        public MergeEngine()
        {
            Type xlType = Type.GetTypeFromProgID("Excel.Application");
            if (xlType == null)
                throw new InvalidOperationException("Excel COM 을 찾을 수 없습니다.");
            _xl = Activator.CreateInstance(xlType);
            _xl.Visible = false;
            _xl.DisplayAlerts = false;
            try { _xl.AutomationSecurity = msoAutomationSecurityForceDisable; } catch { }
            try { _xl.ScreenUpdating = false; } catch { }
            try { _xl.EnableEvents = false; } catch { }
            try { _xl.Calculation = xlCalculationManual; } catch { }
            _pid = ComProcessGuard.TryGetPid(_xl);
        }

        /// <summary>
        /// items 를 targetPath 워크북에 적용하고 저장.
        /// createBackup=true 면 저장 전 "파일명.yyyyMMdd_HHmmss.bak.xlsx" 백업 생성.
        /// </summary>
        public void ApplyAndSave(string targetPath, List<MergeItem> items, bool createBackup, out string backupPath)
        {
            backupPath = null;
            Logger.Info("병합 저장 시작: " + targetPath + " (셀 " + (items != null ? items.Count : 0)
                        + "개, 백업=" + createBackup + ")");
            if (createBackup)
            {
                backupPath = MakeBackup(targetPath);
                Logger.Info("백업 생성: " + backupPath);
            }

            dynamic workbooks = null;
            dynamic wb = null;
            try
            {
                workbooks = _xl.Workbooks;
                wb = workbooks.Open(targetPath); // 편집용 → ReadOnly 아님

                if ((bool)wb.ReadOnly)
                    throw new InvalidOperationException(
                        "대상 파일이 읽기 전용으로 열렸습니다. 저장할 수 없습니다.\r\n"
                        + "→ 원본이 Excel에서 열려 있으면 닫고 다시 시도하세요. "
                        + "(또는 DRM 편집권한이 없는 파일일 수 있습니다.)");

                // 시트별로 묶어서 셀 쓰기.
                Dictionary<string, List<MergeItem>> bySheet = GroupBySheet(items);
                dynamic sheets = wb.Worksheets;
                foreach (KeyValuePair<string, List<MergeItem>> kv in bySheet)
                {
                    dynamic ws = null;
                    try
                    {
                        ws = sheets[kv.Key];
                        WriteCells(ws, kv.Value);
                    }
                    finally
                    {
                        Release(ref ws);
                    }
                }
                Release(ref sheets);

                wb.Save(); // 같은 경로/형식 저장 → Excel 세션이므로 DRM 재적용 유지
                Logger.Info("병합 저장 완료: " + targetPath);
            }
            catch (Exception exSave)
            {
                Logger.Error("병합 저장 실패: " + targetPath, exSave);
                throw;
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
        }

        private void WriteCells(dynamic ws, List<MergeItem> items)
        {
            // 1차: 셀 단위 쓰기(차이 셀만이라 수는 적음). 성능이 필요하면 연속 범위 묶기로 확장.
            // 단, Cells(r,c) 왕복을 최소화하기 위해 동일 행 연속 열은 한 번에 쓴다.
            items.Sort(delegate(MergeItem a, MergeItem b)
            {
                if (a.Row != b.Row) return a.Row.CompareTo(b.Row);
                return a.Col.CompareTo(b.Col);
            });

            int i = 0;
            while (i < items.Count)
            {
                int row = items[i].Row;
                int startCol = items[i].Col;
                int j = i;
                // 같은 행에서 연속된 열 묶기.
                while (j + 1 < items.Count &&
                       items[j + 1].Row == row &&
                       items[j + 1].Col == items[j].Col + 1)
                {
                    j++;
                }
                int len = j - i + 1;
                if (len == 1)
                {
                    dynamic cell = ws.Cells[row, startCol];
                    try { cell.Value = items[i].Value; }
                    finally { Release(ref cell); }
                }
                else
                {
                    object[,] block = new object[1, len];
                    for (int k = 0; k < len; k++) block[0, k] = items[i + k].Value;
                    dynamic topLeft = ws.Cells[row, startCol];
                    dynamic bottomRight = ws.Cells[row, startCol + len - 1];
                    dynamic range = ws.Range[topLeft, bottomRight];
                    try { range.Value = block; }
                    finally
                    {
                        Release(ref range);
                        Release(ref bottomRight);
                        Release(ref topLeft);
                    }
                }
                i = j + 1;
            }
        }

        private static Dictionary<string, List<MergeItem>> GroupBySheet(List<MergeItem> items)
        {
            Dictionary<string, List<MergeItem>> map = new Dictionary<string, List<MergeItem>>();
            foreach (MergeItem it in items)
            {
                List<MergeItem> list;
                if (!map.TryGetValue(it.SheetName, out list))
                {
                    list = new List<MergeItem>();
                    map[it.SheetName] = list;
                }
                list.Add(it);
            }
            return map;
        }

        private static string MakeBackup(string path)
        {
            string dir = Path.GetDirectoryName(path);
            string name = Path.GetFileNameWithoutExtension(path);
            string ext = Path.GetExtension(path);
            // Date.Now 사용 — 백업 파일명 유일화용.
            string stamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            string backup = Path.Combine(dir, name + "." + stamp + ".bak" + ext);
            File.Copy(path, backup, false);
            return backup;
        }

        private static void Release(ref dynamic o)
        {
            // dynamic '== null' 은 끊긴 COM 객체에서 바인더 예외(RPC_E_INVALID_OBJECT) 위험 → (object) 로 벗겨서 처리.
            object obj = (object)o;
            o = null;
            if (obj == null) return;
            try { if (Marshal.IsComObject(obj)) Marshal.ReleaseComObject(obj); }
            catch { }
        }

        public void Dispose()
        {
            object app = (object)_xl;   // dynamic 벗겨내기(동적 == 회피)
            _xl = null;
            if (app != null)
            {
                try { ((dynamic)app).Quit(); } catch { }
                try { if (Marshal.IsComObject(app)) Marshal.ReleaseComObject(app); } catch { }
            }
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
            ComProcessGuard.KillIfAlive(_pid, "EXCEL");
        }
    }
}
