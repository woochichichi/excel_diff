using System;
using System.Collections.Generic;
using System.Globalization;

namespace ExcelDiffMerge
{
    /// <summary>
    /// 2-way diff 엔진 (spec §5).
    ///  - 시트명 기준 매칭. 한쪽에만 있으면 추가/삭제.
    ///  - 셀은 절대좌표(1-based) union 으로 비교.
    ///  - 1차 비교 대상: 표시값(Value 를 문자열화) + 수식(Formula). 서식 diff 는 2차.
    ///  - 행 삽입/삭제는 1차에선 좌표 기준 단순 비교(경고). key/LCS 정렬은 2차.
    /// </summary>
    public static class DiffEngine
    {
        public static DiffResult Compare(WorkbookData left, WorkbookData right)
        {
            DiffResult result = new DiffResult();
            result.LeftPath = left.FilePath;
            result.RightPath = right.FilePath;

            // 시트 순서: left 순서 우선, 그 뒤 right 에만 있는 시트.
            List<string> sheetNames = new List<string>();
            HashSet<string> seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (SheetData s in left.Sheets)
                if (seen.Add(s.Name)) sheetNames.Add(s.Name);
            foreach (SheetData s in right.Sheets)
                if (seen.Add(s.Name)) sheetNames.Add(s.Name);

            foreach (string name in sheetNames)
            {
                SheetData l = left.FindSheet(name);
                SheetData r = right.FindSheet(name);
                SheetDiff sd = CompareSheet(name, l, r);
                result.Sheets.Add(sd);

                result.TotalChanged += sd.ChangedCount;
                result.TotalAdded += sd.AddedCount;
                result.TotalDeleted += sd.DeletedCount;
                if (sd.HasChanges) result.ChangedSheetCount++;
            }
            return result;
        }

        private static SheetDiff CompareSheet(string name, SheetData l, SheetData r)
        {
            SheetDiff sd = new SheetDiff();
            sd.Name = name;
            sd.Left = l;
            sd.Right = r;

            if (l == null && r != null)
            {
                sd.Status = SheetStatus.Added;
                SetBounds(sd, r);
                // 우측 전체를 Added 로 표기.
                for (int rr = r.FirstRow; rr <= r.LastRow; rr++)
                    for (int cc = r.FirstCol; cc <= r.LastCol; cc++)
                    {
                        object v = r.GetValueAbs(rr, cc);
                        if (ExcelComReader.IsEmpty(v)) continue;
                        DiffCell dc = new DiffCell(rr, cc, CellStatus.Added);
                        dc.NewValue = v;
                        dc.NewFormula = r.GetFormulaAbs(rr, cc);
                        sd.Cells.Add(dc);
                        sd.AddedCount++;
                    }
                sd.BuildIndex();
                return sd;
            }
            if (r == null && l != null)
            {
                sd.Status = SheetStatus.Deleted;
                SetBounds(sd, l);
                for (int rr = l.FirstRow; rr <= l.LastRow; rr++)
                    for (int cc = l.FirstCol; cc <= l.LastCol; cc++)
                    {
                        object v = l.GetValueAbs(rr, cc);
                        if (ExcelComReader.IsEmpty(v)) continue;
                        DiffCell dc = new DiffCell(rr, cc, CellStatus.Deleted);
                        dc.OldValue = v;
                        dc.OldFormula = l.GetFormulaAbs(rr, cc);
                        sd.Cells.Add(dc);
                        sd.DeletedCount++;
                    }
                sd.BuildIndex();
                return sd;
            }
            if (l == null && r == null)
            {
                sd.Status = SheetStatus.Same;
                return sd;
            }

            // 양쪽 존재 → 셀별 비교. 절대좌표 union.
            int minRow = Math.Min(l.FirstRow, r.FirstRow);
            int minCol = Math.Min(l.FirstCol, r.FirstCol);
            int maxRow = Math.Max(l.LastRow, r.LastRow);
            int maxCol = Math.Max(l.LastCol, r.LastCol);

            sd.MinRow = minRow;
            sd.MinCol = minCol;
            sd.MaxRow = maxRow;
            sd.MaxCol = maxCol;

            for (int rr = minRow; rr <= maxRow; rr++)
            {
                for (int cc = minCol; cc <= maxCol; cc++)
                {
                    object lv = l.GetValueAbs(rr, cc);
                    object rv = r.GetValueAbs(rr, cc);
                    object lf = l.GetFormulaAbs(rr, cc);
                    object rf = r.GetFormulaAbs(rr, cc);

                    bool lEmpty = ExcelComReader.IsEmpty(lv);
                    bool rEmpty = ExcelComReader.IsEmpty(rv);

                    if (lEmpty && rEmpty) continue; // 둘 다 빈 셀 → 스킵

                    CellStatus status;
                    if (lEmpty && !rEmpty) status = CellStatus.Added;
                    else if (!lEmpty && rEmpty) status = CellStatus.Deleted;
                    else
                    {
                        // 둘 다 값 있음 → 표시값 + 수식 비교
                        bool valSame = ValueEquals(lv, rv);
                        bool formulaSame = FormulaEquals(lf, rf);
                        if (valSame && formulaSame) continue; // 동일
                        status = CellStatus.Changed;
                    }

                    DiffCell dc = new DiffCell(rr, cc, status);
                    dc.OldValue = lv;
                    dc.NewValue = rv;
                    dc.OldFormula = lf;
                    dc.NewFormula = rf;
                    sd.Cells.Add(dc);

                    if (status == CellStatus.Changed) sd.ChangedCount++;
                    else if (status == CellStatus.Added) sd.AddedCount++;
                    else if (status == CellStatus.Deleted) sd.DeletedCount++;
                }
            }

            sd.Status = sd.HasChanges ? SheetStatus.Modified : SheetStatus.Same;
            sd.BuildIndex();
            return sd;
        }

        private static void SetBounds(SheetDiff sd, SheetData s)
        {
            sd.MinRow = s.FirstRow;
            sd.MinCol = s.FirstCol;
            sd.MaxRow = s.LastRow;
            sd.MaxCol = s.LastCol;
        }

        /// <summary>표시값 비교 — 숫자는 오차 없이, 그 외는 문자열 표현으로.</summary>
        public static bool ValueEquals(object a, object b)
        {
            if (a == null && b == null) return true;
            if (a == null || b == null) return false;

            // 둘 다 숫자면 double 비교.
            if (IsNumeric(a) && IsNumeric(b))
            {
                double da = Convert.ToDouble(a, CultureInfo.InvariantCulture);
                double db = Convert.ToDouble(b, CultureInfo.InvariantCulture);
                return da == db;
            }
            if (a is bool && b is bool)
                return (bool)a == (bool)b;
            if (a is DateTime && b is DateTime)
                return (DateTime)a == (DateTime)b;

            return string.Equals(ToText(a), ToText(b), StringComparison.Ordinal);
        }

        private static bool FormulaEquals(object a, object b)
        {
            string sa = a == null ? "" : a.ToString();
            string sb = b == null ? "" : b.ToString();
            return string.Equals(sa, sb, StringComparison.Ordinal);
        }

        private static bool IsNumeric(object o)
        {
            return o is double || o is int || o is long || o is float ||
                   o is decimal || o is short || o is byte;
        }

        /// <summary>셀 값을 표시 문자열로. UI/비교 공용.</summary>
        public static string ToText(object o)
        {
            if (o == null) return "";
            if (o is double)
            {
                double d = (double)o;
                // 정수형이면 소수점 제거.
                if (d == Math.Floor(d) && !double.IsInfinity(d) && Math.Abs(d) < 1e15)
                    return ((long)d).ToString(CultureInfo.InvariantCulture);
                return d.ToString(CultureInfo.InvariantCulture);
            }
            if (o is DateTime)
                return ((DateTime)o).ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
            if (o is bool)
                return ((bool)o) ? "TRUE" : "FALSE";
            return o.ToString();
        }
    }
}
