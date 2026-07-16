using System;
using System.Collections.Generic;

namespace ExcelDiffMerge
{
    /// <summary>
    /// N-way 취합 결과의 셀 하나: base 대비 각 버전 값 + 충돌 여부 (spec §5-5).
    /// </summary>
    public sealed class NWayCell
    {
        public int Row;
        public int Col;
        public object BaseValue;
        // 버전 인덱스 → 값 (base 와 다른 버전만 담김).
        public Dictionary<int, object> Changes = new Dictionary<int, object>();
        public bool Conflict; // 2개 이상 버전이 서로 다른 값으로 변경

        public NWayCell(int row, int col)
        {
            Row = row;
            Col = col;
        }
    }

    public sealed class NWaySheetDiff
    {
        public string Name;
        public int MinRow = 1, MinCol = 1, MaxRow = 0, MaxCol = 0;
        public List<NWayCell> Cells = new List<NWayCell>();
        public int ChangedCount;
        public int ConflictCount;
        // 이 시트가 base 워크북에 존재하는지. false 면 버전에만 있는 시트라 base 에 병합/저장 불가(B2).
        public bool InBase;
    }

    public sealed class NWayResult
    {
        public string BasePath;
        public List<string> VersionPaths = new List<string>(); // base 제외 나머지
        public List<NWaySheetDiff> Sheets = new List<NWaySheetDiff>();
        public int TotalChanged;
        public int TotalConflict;

        public string Summary()
        {
            return string.Format("버전 {0}개 비교  |  변경 셀 {1}  충돌 {2}",
                VersionPaths.Count, TotalChanged, TotalConflict);
        }
    }

    /// <summary>
    /// N-way 엔진: base 1개 + 나머지 버전들을 base 대비 diff.
    /// 셀별로 "어느 버전이 뭘 바꿨나" 표시, 충돌은 별도 강조(§5-5).
    /// </summary>
    public static class NWayDiffEngine
    {
        public static NWayResult Compare(WorkbookData baseWb, List<WorkbookData> versions)
        {
            NWayResult result = new NWayResult();
            result.BasePath = baseWb.FilePath;
            foreach (WorkbookData v in versions) result.VersionPaths.Add(v.FilePath);

            // 시트 union (base 우선).
            List<string> names = new List<string>();
            HashSet<string> seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (SheetData s in baseWb.Sheets) if (seen.Add(s.Name)) names.Add(s.Name);
            foreach (WorkbookData v in versions)
                foreach (SheetData s in v.Sheets)
                    if (seen.Add(s.Name)) names.Add(s.Name);

            foreach (string name in names)
            {
                SheetData bs = baseWb.FindSheet(name);
                NWaySheetDiff nsd = new NWaySheetDiff();
                nsd.Name = name;
                nsd.InBase = (bs != null); // base 에 있는 시트만 저장 가능(B2)

                // union 경계 산정.
                int minRow = int.MaxValue, minCol = int.MaxValue, maxRow = 0, maxCol = 0;
                Include(bs, ref minRow, ref minCol, ref maxRow, ref maxCol);
                List<SheetData> vSheets = new List<SheetData>();
                foreach (WorkbookData v in versions)
                {
                    SheetData vs = v.FindSheet(name);
                    vSheets.Add(vs);
                    Include(vs, ref minRow, ref minCol, ref maxRow, ref maxCol);
                }
                if (maxRow == 0 || maxCol == 0) { result.Sheets.Add(nsd); continue; }
                if (minRow == int.MaxValue) minRow = 1;
                if (minCol == int.MaxValue) minCol = 1;

                nsd.MinRow = minRow; nsd.MinCol = minCol; nsd.MaxRow = maxRow; nsd.MaxCol = maxCol;

                for (int rr = minRow; rr <= maxRow; rr++)
                {
                    for (int cc = minCol; cc <= maxCol; cc++)
                    {
                        object bv = bs != null ? bs.GetValueAbs(rr, cc) : null;
                        NWayCell cell = null;
                        List<object> distinctChanges = new List<object>();

                        for (int vi = 0; vi < vSheets.Count; vi++)
                        {
                            SheetData vs = vSheets[vi];
                            object vv = vs != null ? vs.GetValueAbs(rr, cc) : null;

                            bool bEmpty = ExcelComReader.IsEmpty(bv);
                            bool vEmpty = ExcelComReader.IsEmpty(vv);
                            if (bEmpty && vEmpty) continue;
                            if (!DiffEngine.ValueEquals(bv, vv))
                            {
                                if (cell == null) cell = new NWayCell(rr, cc) { BaseValue = bv };
                                cell.Changes[vi] = vv;

                                bool dup = false;
                                foreach (object dv in distinctChanges)
                                    if (DiffEngine.ValueEquals(dv, vv)) { dup = true; break; }
                                if (!dup) distinctChanges.Add(vv);
                            }
                        }

                        if (cell != null)
                        {
                            // 서로 다른 변경값이 2종류 이상이면 충돌.
                            cell.Conflict = distinctChanges.Count >= 2;
                            nsd.Cells.Add(cell);
                            nsd.ChangedCount++;
                            if (cell.Conflict) nsd.ConflictCount++;
                        }
                    }
                }

                result.TotalChanged += nsd.ChangedCount;
                result.TotalConflict += nsd.ConflictCount;
                result.Sheets.Add(nsd);
            }
            return result;
        }

        private static void Include(SheetData s, ref int minRow, ref int minCol, ref int maxRow, ref int maxCol)
        {
            if (s == null || s.RowCount == 0 || s.ColCount == 0) return;
            if (s.FirstRow < minRow) minRow = s.FirstRow;
            if (s.FirstCol < minCol) minCol = s.FirstCol;
            if (s.LastRow > maxRow) maxRow = s.LastRow;
            if (s.LastCol > maxCol) maxCol = s.LastCol;
        }
    }
}
