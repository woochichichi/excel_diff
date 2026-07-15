using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;

namespace ExcelDiffMerge
{
    /// <summary>
    /// 로컬 테스트용 CSV 로더 (Excel/COM/DRM 불필요).
    ///  - 경로가 폴더면 그 안의 *.csv 각각을 시트로 로드(시트명 = 파일명).
    ///  - 경로가 .csv 파일이면 단일 시트.
    ///  - 숫자는 double 로 파싱, '=' 로 시작하면 수식으로 간주(값=수식문자열).
    /// 운영 ExcelComReader 와 동일한 WorkbookData 를 만들어 같은 DiffEngine 을 태운다.
    /// </summary>
    public sealed class CsvWorkbookReader : IWorkbookReader
    {
        public WorkbookData LoadWorkbook(string path)
        {
            WorkbookData wb = new WorkbookData(path);
            if (Directory.Exists(path))
            {
                string[] files = Directory.GetFiles(path, "*.csv");
                Array.Sort(files, StringComparer.OrdinalIgnoreCase);
                foreach (string f in files)
                    wb.Sheets.Add(LoadSheet(f, Path.GetFileNameWithoutExtension(f)));
            }
            else if (File.Exists(path))
            {
                wb.Sheets.Add(LoadSheet(path, Path.GetFileNameWithoutExtension(path)));
            }
            else
            {
                throw new FileNotFoundException("CSV 경로를 찾을 수 없습니다: " + path);
            }
            return wb;
        }

        private static SheetData LoadSheet(string file, string name)
        {
            List<string[]> rows = ParseCsv(File.ReadAllText(file, Encoding.UTF8));
            int nrow = rows.Count;
            int ncol = 0;
            foreach (string[] r in rows) if (r.Length > ncol) ncol = r.Length;

            SheetData sd = new SheetData(name);
            sd.FirstRow = 1;
            sd.FirstCol = 1;
            sd.RowCount = nrow;
            sd.ColCount = ncol;
            sd.Values = new object[nrow, ncol];
            sd.Formulas = new object[nrow, ncol];

            for (int r = 0; r < nrow; r++)
            {
                string[] cells = rows[r];
                for (int c = 0; c < ncol; c++)
                {
                    string text = c < cells.Length ? cells[c] : "";
                    object val, formula;
                    ParseCell(text, out val, out formula);
                    sd.Values[r, c] = val;
                    sd.Formulas[r, c] = formula;
                }
            }
            return sd;
        }

        private static void ParseCell(string text, out object value, out object formula)
        {
            formula = null;
            if (string.IsNullOrEmpty(text)) { value = null; return; }
            if (text.Length > 0 && text[0] == '=')
            {
                formula = text;
                value = text; // 목업: 수식 셀은 값도 수식문자열로 취급
                return;
            }
            double d;
            if (double.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out d))
            {
                value = d;
                return;
            }
            value = text;
        }

        /// <summary>따옴표/쉼표/개행을 처리하는 최소 CSV 파서(외부 라이브러리 0).</summary>
        private static List<string[]> ParseCsv(string content)
        {
            List<string[]> rows = new List<string[]>();
            List<string> cur = new List<string>();
            StringBuilder field = new StringBuilder();
            bool inQuotes = false;
            int i = 0;
            int n = content.Length;

            while (i < n)
            {
                char ch = content[i];
                if (inQuotes)
                {
                    if (ch == '"')
                    {
                        if (i + 1 < n && content[i + 1] == '"') { field.Append('"'); i += 2; continue; }
                        inQuotes = false; i++; continue;
                    }
                    field.Append(ch); i++; continue;
                }
                if (ch == '"') { inQuotes = true; i++; continue; }
                if (ch == ',') { cur.Add(field.ToString()); field.Length = 0; i++; continue; }
                if (ch == '\r') { i++; continue; }
                if (ch == '\n')
                {
                    cur.Add(field.ToString()); field.Length = 0;
                    rows.Add(cur.ToArray()); cur = new List<string>();
                    i++; continue;
                }
                field.Append(ch); i++;
            }
            // 마지막 필드/행.
            if (field.Length > 0 || cur.Count > 0)
            {
                cur.Add(field.ToString());
                rows.Add(cur.ToArray());
            }
            return rows;
        }

        public void Dispose() { /* CSV 리더는 정리할 자원 없음 */ }
    }
}
