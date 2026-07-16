using System;
using System.Collections.Generic;
using System.Globalization;

namespace ExcelDiffMerge
{
    /// <summary>
    /// 2-way diff 엔진 (spec §5).
    ///  - 시트명 기준 매칭. 한쪽에만 있으면 추가/삭제 시트.
    ///  - 행은 RowAligner 로 정렬(좌표/자동LCS/키). 정렬 후 짝지어진 행끼리 셀 비교.
    ///  - 비교 대상: 표시값(Value) + 수식(Formula). 서식 diff 는 2차.
    /// </summary>
    public static class DiffEngine
    {
        public static DiffResult Compare(WorkbookData left, WorkbookData right)
        {
            return Compare(left, right, AlignMode.Auto, -1);
        }

        public static DiffResult Compare(WorkbookData left, WorkbookData right, AlignMode mode, int keyCol)
        {
            DiffResult result = new DiffResult();
            result.LeftPath = left.FilePath;
            result.RightPath = right.FilePath;

            List<string> sheetNames = new List<string>();
            HashSet<string> seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (SheetData s in left.Sheets) if (seen.Add(s.Name)) sheetNames.Add(s.Name);
            foreach (SheetData s in right.Sheets) if (seen.Add(s.Name)) sheetNames.Add(s.Name);

            foreach (string name in sheetNames)
            {
                SheetData l = left.FindSheet(name);
                SheetData r = right.FindSheet(name);
                SheetDiff sd = CompareSheet(name, l, r, mode, keyCol);
                result.Sheets.Add(sd);
                result.TotalChanged += sd.ChangedCount;
                result.TotalAdded += sd.AddedCount;
                result.TotalDeleted += sd.DeletedCount;
                if (sd.HasChanges) result.ChangedSheetCount++;
            }
            return result;
        }

        private static SheetDiff CompareSheet(string name, SheetData l, SheetData r, AlignMode mode, int keyCol)
        {
            SheetDiff sd = new SheetDiff();
            sd.Name = name;
            sd.Left = l;
            sd.Right = r;

            if (l == null && r == null) { sd.Status = SheetStatus.Same; return sd; }

            // 열 union.
            int minCol, maxCol;
            if (l != null && r != null)
            {
                minCol = Math.Min(l.FirstCol, r.FirstCol);
                maxCol = Math.Max(l.LastCol, r.LastCol);
            }
            else if (l != null) { minCol = l.FirstCol; maxCol = l.LastCol; }
            else { minCol = r.FirstCol; maxCol = r.LastCol; }
            if (maxCol < minCol) { minCol = 1; maxCol = 0; } // 빈 시트
            sd.MinCol = minCol;
            sd.MaxCol = maxCol;

            List<AlignPair> pairs = RowAligner.Align(l, r, mode, keyCol, minCol, maxCol);

            for (int p = 0; p < pairs.Count; p++)
            {
                AlignPair ap = pairs[p];
                DiffRow dr = new DiffRow(ap.LeftRow, ap.RightRow);
                Dictionary<int, CellStatus> changes = null;

                for (int c = minCol; c <= maxCol; c++)
                {
                    object lv = (l != null && ap.LeftRow >= 0) ? l.GetValueAbs(ap.LeftRow, c) : null;
                    object rv = (r != null && ap.RightRow >= 0) ? r.GetValueAbs(ap.RightRow, c) : null;
                    object lf = (l != null && ap.LeftRow >= 0) ? l.GetFormulaAbs(ap.LeftRow, c) : null;
                    object rf = (r != null && ap.RightRow >= 0) ? r.GetFormulaAbs(ap.RightRow, c) : null;

                    CellStatus st = CellStatusOf(lv, lf, rv, rf);
                    if (st == CellStatus.Same) continue;

                    if (changes == null) changes = new Dictionary<int, CellStatus>();
                    changes[c] = st;
                    sd.Nav.Add(new DiffNav(p, c));

                    if (st == CellStatus.Changed) sd.ChangedCount++;
                    else if (st == CellStatus.Added) sd.AddedCount++;
                    else if (st == CellStatus.Deleted) sd.DeletedCount++;
                }

                dr.Changes = changes;
                if (ap.LeftRow < 0) dr.Kind = RowKind.Added;
                else if (ap.RightRow < 0) dr.Kind = RowKind.Deleted;
                else if (changes != null) dr.Kind = RowKind.Changed;
                else dr.Kind = RowKind.Same;

                sd.Rows.Add(dr);
            }

            if (l == null) sd.Status = SheetStatus.Added;
            else if (r == null) sd.Status = SheetStatus.Deleted;
            else sd.Status = sd.HasChanges ? SheetStatus.Modified : SheetStatus.Same;
            return sd;
        }

        /// <summary>정렬된 셀 한 쌍의 상태.</summary>
        public static CellStatus CellStatusOf(object lv, object lf, object rv, object rf)
        {
            bool le = ExcelComReader.IsEmpty(lv);
            bool re = ExcelComReader.IsEmpty(rv);
            if (le && re) return CellStatus.Same;
            if (le && !re) return CellStatus.Added;
            if (!le && re) return CellStatus.Deleted;
            if (ValueEquals(lv, rv) && FormulaEquals(lf, rf)) return CellStatus.Same;
            return CellStatus.Changed;
        }

        public static bool ValueEquals(object a, object b)
        {
            if (a == null && b == null) return true;
            if (a == null || b == null) return false;
            bool na = IsNumeric(a);
            bool nb = IsNumeric(b);
            if (na && nb)
            {
                double da = Convert.ToDouble(a, CultureInfo.InvariantCulture);
                double db = Convert.ToDouble(b, CultureInfo.InvariantCulture);
                return da == db;
            }
            // 한쪽만 숫자 타입(double 등)이고 다른쪽이 비숫자(문자열 "100" 등)이면
            // 타입 변경이므로 '다름'으로 판정한다(B4). 문자열 비교로 넘겨 Same 처리하지 않는다.
            // (null/빈값은 위/호출부 IsEmpty 에서 이미 처리되므로 여기 도달하지 않는다.)
            if (na != nb) return false;
            if (a is bool && b is bool) return (bool)a == (bool)b;
            if (a is DateTime && b is DateTime) return (DateTime)a == (DateTime)b;
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

        /// <summary>셀 값을 표시 문자열로.</summary>
        public static string ToText(object o)
        {
            if (o == null) return "";
            if (o is double)
            {
                double d = (double)o;
                if (d == Math.Floor(d) && !double.IsInfinity(d) && Math.Abs(d) < 1e15)
                    return ((long)d).ToString(CultureInfo.InvariantCulture);
                return d.ToString(CultureInfo.InvariantCulture);
            }
            if (o is DateTime)
                return ((DateTime)o).ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
            if (o is bool) return ((bool)o) ? "TRUE" : "FALSE";
            return o.ToString();
        }
    }
}
