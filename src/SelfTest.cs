using System;
using System.Collections.Generic;
using System.Globalization;

namespace ExcelDiffMerge
{
    /// <summary>
    /// 내장 자기검증 (ExcelDiffMerge.exe --selftest).
    /// Excel/COM/CSV파일 없이 메모리 목업으로 DiffEngine + RowAligner 를 검증한다.
    /// tests/reference.py(파이썬 가상테스트)와 동일한 케이스/기대값을 담아,
    /// 현장 빌드 직후 대상 PC 에서도 로직 정상 여부를 즉시 확인할 수 있게 한다.
    /// </summary>
    public static class SelfTest
    {
        private static int _pass;
        private static int _fail;

        public static int Run()
        {
            _pass = 0; _fail = 0;
            Console.WriteLine("=== ExcelDiffMerge 자기검증 ===");

            WorkbookData oldWb = BuildOld();
            WorkbookData newWb = BuildNew();

            DiffResult auto = DiffEngine.Compare(oldWb, newWb, AlignMode.Auto, -1);
            DiffResult coord = DiffEngine.Compare(oldWb, newWb, AlignMode.Coordinate, -1);
            DiffResult keyed = DiffEngine.Compare(oldWb, newWb, AlignMode.KeyColumn, 1); // ID=A열=abs 1

            SheetDiff s1 = Find(auto, "Sheet1");
            int addedRows, deletedRows, changedRows;
            CountRowKinds(s1, out addedRows, out deletedRows, out changedRows);

            Check(addedRows >= 1, "Sheet1(auto): 삽입행(Blueberry) 을 '행추가'로 인식");
            Check(deletedRows >= 1, "Sheet1(auto): 삭제행(Date) 을 '행삭제'로 인식");
            Check(s1.ChangedCount <= 4, "Sheet1(auto): cascade 오판 없음 (changed=" + s1.ChangedCount + ")");

            SheetDiff s1c = Find(coord, "Sheet1");
            Check(s1c.ChangedCount > s1.ChangedCount,
                "좌표정렬은 cascade 로 변경이 더 많다 (coord=" + s1c.ChangedCount + " > auto=" + s1.ChangedCount + ")");

            SheetDiff s2 = Find(auto, "Sheet2");
            Check(s2.ChangedCount == 2, "Sheet2: 수식/값 변경 2건 (got " + s2.ChangedCount + ")");

            SheetDiff onlyNew = Find(auto, "OnlyNew");
            Check(onlyNew != null && onlyNew.Status == SheetStatus.Added, "OnlyNew: 추가 시트 인식");
            SheetDiff onlyOld = Find(auto, "OnlyOld");
            Check(onlyOld != null && onlyOld.Status == SheetStatus.Deleted, "OnlyOld: 삭제 시트 인식");

            SheetDiff s1k = Find(keyed, "Sheet1");
            int ka, kd, kc;
            CountRowKinds(s1k, out ka, out kd, out kc);
            Check(ka >= 1 && kd >= 1, "Sheet1(key): 키기준 삽입/삭제 인식");

            // B4 회귀: 숫자 타입(double 100) vs 문자열 "100" → 타입 변경으로 '다름' 감지.
            Check(!DiffEngine.ValueEquals(100.0, "100"),
                "B4: 숫자 100 vs 문자열 \"100\" → 값 다름(타입 변경) 감지");
            Check(DiffEngine.CellStatusOf(100.0, null, "100", null) == CellStatus.Changed,
                "B4: 숫자 100 vs 문자열 \"100\" → 셀 상태 Changed");
            // 회귀 방지: 같은 타입끼리는 기존 동작 유지.
            Check(DiffEngine.ValueEquals(100.0, 100.0), "B4: 숫자 100 vs 숫자 100 → 같음(회귀 방지)");
            Check(DiffEngine.ValueEquals("abc", "abc"), "B4: 문자 abc vs 문자 abc → 같음(회귀 방지)");
            Check(!DiffEngine.ValueEquals("100", "100.0"), "B4: 문자열끼리는 문자열 비교(\"100\"≠\"100.0\")");

            // B2 순수 로직: 버전에만 있는 시트는 InBase=false → 채택/저장 차단 대상.
            WorkbookData b2base = new WorkbookData("<b2-base>");
            b2base.Sheets.Add(Sheet("Common", new string[][] {
                new string[]{"A","B"}, new string[]{"1","2"},
            }));
            WorkbookData b2ver = new WorkbookData("<b2-ver>");
            b2ver.Sheets.Add(Sheet("Common", new string[][] {
                new string[]{"A","B"}, new string[]{"1","9"},
            }));
            b2ver.Sheets.Add(Sheet("VerOnly", new string[][] {
                new string[]{"X"}, new string[]{"7"},
            }));
            List<WorkbookData> b2versions = new List<WorkbookData>();
            b2versions.Add(b2ver);
            NWayResult b2res = NWayDiffEngine.Compare(b2base, b2versions);
            NWaySheetDiff nCommon = FindNWay(b2res, "Common");
            NWaySheetDiff nVerOnly = FindNWay(b2res, "VerOnly");
            Check(nCommon != null && nCommon.InBase, "B2: base 에 있는 시트 InBase=true");
            Check(nVerOnly != null && !nVerOnly.InBase, "B2: 버전 전용 시트 InBase=false(저장 차단 대상)");

            // ColLetter 왕복.
            Check(MainForm.ColLetter(1) == "A" && MainForm.ColLetter(27) == "AA", "ColLetter 변환");
            Check(MainForm.ParseColLetter("AA") == 27 && MainForm.ParseColLetter("A") == 1, "ParseColLetter 역변환");

            Console.WriteLine(string.Format("\n결과: {0} PASS / {1} FAIL  →  {2}",
                _pass, _fail, _fail == 0 ? "ALL PASS ✅" : "FAIL ❌"));
            return _fail == 0 ? 0 : 1;
        }

        private static void Check(bool cond, string msg)
        {
            Console.WriteLine((cond ? "  PASS " : "  FAIL ") + msg);
            if (cond) _pass++; else _fail++;
        }

        private static SheetDiff Find(DiffResult d, string name)
        {
            foreach (SheetDiff sd in d.Sheets)
                if (sd.Name == name) return sd;
            return null;
        }

        private static NWaySheetDiff FindNWay(NWayResult d, string name)
        {
            foreach (NWaySheetDiff sd in d.Sheets)
                if (sd.Name == name) return sd;
            return null;
        }

        private static void CountRowKinds(SheetDiff sd, out int added, out int deleted, out int changed)
        {
            added = deleted = changed = 0;
            if (sd == null) return;
            foreach (DiffRow dr in sd.Rows)
            {
                if (dr.Kind == RowKind.Added) added++;
                else if (dr.Kind == RowKind.Deleted) deleted++;
                else if (dr.Kind == RowKind.Changed) changed++;
            }
        }

        // ---- 목업 빌더 (tests/data 의 CSV 와 동일 내용) ----
        private static WorkbookData BuildOld()
        {
            WorkbookData wb = new WorkbookData("<mock-old>");
            wb.Sheets.Add(Sheet("Sheet1", new string[][] {
                new string[]{"ID","Name","Qty","Price","Note"},
                new string[]{"1","Apple","10","100","fresh"},
                new string[]{"2","Banana","5","200","ripe"},
                new string[]{"3","Cherry","7","300",""},
                new string[]{"4","Date","2","400","dry"},
                new string[]{"5","Elderberry","9","500","rare"},
            }));
            wb.Sheets.Add(Sheet("Sheet2", new string[][] {
                new string[]{"Item","Base","Tax","Total"},
                new string[]{"A","100","=B2*0.1","=B2+C2"},
                new string[]{"B","200","=B3*0.1","=B3+C3"},
                new string[]{"C","300","=B4*0.1","=B4+C4"},
            }));
            wb.Sheets.Add(Sheet("OnlyOld", new string[][] {
                new string[]{"P","Q"}, new string[]{"9","8"},
            }));
            return wb;
        }

        private static WorkbookData BuildNew()
        {
            WorkbookData wb = new WorkbookData("<mock-new>");
            wb.Sheets.Add(Sheet("Sheet1", new string[][] {
                new string[]{"ID","Name","Qty","Price","Note"},
                new string[]{"1","Apple","10","150","fresh"},
                new string[]{"2","Banana","5","200","ripe"},
                new string[]{"25","Blueberry","3","250","new"},
                new string[]{"3","Cherry","9","300","tart"},
                new string[]{"5","Elderberry","9","500",""},
            }));
            wb.Sheets.Add(Sheet("Sheet2", new string[][] {
                new string[]{"Item","Base","Tax","Total"},
                new string[]{"A","100","=B2*0.15","=B2+C2"},
                new string[]{"B","250","=B3*0.1","=B3+C3"},
                new string[]{"C","300","=B4*0.1","=B4+C4"},
            }));
            wb.Sheets.Add(Sheet("OnlyNew", new string[][] {
                new string[]{"X","Y"}, new string[]{"1","2"},
            }));
            return wb;
        }

        private static SheetData Sheet(string name, string[][] rows)
        {
            int nrow = rows.Length;
            int ncol = 0;
            for (int r = 0; r < nrow; r++) if (rows[r].Length > ncol) ncol = rows[r].Length;
            SheetData sd = new SheetData(name);
            sd.FirstRow = 1; sd.FirstCol = 1; sd.RowCount = nrow; sd.ColCount = ncol;
            sd.Values = new object[nrow, ncol];
            sd.Formulas = new object[nrow, ncol];
            for (int r = 0; r < nrow; r++)
            {
                for (int c = 0; c < ncol; c++)
                {
                    string t = c < rows[r].Length ? rows[r][c] : "";
                    object val, formula;
                    ParseCell(t, out val, out formula);
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
            if (text[0] == '=') { formula = text; value = text; return; }
            double d;
            if (double.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out d)) { value = d; return; }
            value = text;
        }
    }
}
