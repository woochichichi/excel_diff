using System;
using System.Collections.Generic;

namespace ExcelDiffMerge
{
    /// <summary>셀 비교 결과 상태.</summary>
    public enum CellStatus
    {
        Same = 0,     // 동일
        Changed = 1,  // 양쪽 값/수식 다름 (노랑)
        Added = 2,    // new(우)에만 존재 (초록)
        Deleted = 3,  // old(좌)에만 존재 (빨강)
        Conflict = 4  // N-way: 여러 버전이 같은 셀을 다르게 변경 (주황)
    }

    /// <summary>시트 단위 비교 상태.</summary>
    public enum SheetStatus { Same = 0, Modified = 1, Added = 2, Deleted = 3 }

    /// <summary>정렬된 표시행의 종류.</summary>
    public enum RowKind { Same = 0, Changed = 1, Added = 2, Deleted = 3 }

    /// <summary>행 정렬 방식 (spec §5-4 대응).</summary>
    public enum AlignMode
    {
        Coordinate = 0, // 좌표 기준(행 삽입/삭제 미인지) — 가장 빠름
        Auto = 1,       // 유사도 LCS 자동 정렬(행 삽입/삭제 인지)
        KeyColumn = 2   // 키 컬럼 기준 정렬(조인)
    }

    /// <summary>
    /// 한 시트를 UsedRange.Value / .Formula 로 일괄 로드한 결과.
    /// 좌표는 시트 절대좌표(1-based)를 기준으로 저장한다.
    /// </summary>
    public sealed class SheetData
    {
        public string Name;
        public int FirstRow = 1;
        public int FirstCol = 1;
        public int RowCount;
        public int ColCount;
        public object[,] Values;    // 0-based
        public object[,] Formulas;  // 0-based

        public SheetData(string name)
        {
            Name = name;
            Values = new object[0, 0];
            Formulas = new object[0, 0];
        }

        public int LastRow { get { return FirstRow + RowCount - 1; } }
        public int LastCol { get { return FirstCol + ColCount - 1; } }

        public object GetValueAbs(int absRow, int absCol)
        {
            int r = absRow - FirstRow;
            int c = absCol - FirstCol;
            if (r < 0 || c < 0 || r >= RowCount || c >= ColCount) return null;
            return Values[r, c];
        }

        public object GetFormulaAbs(int absRow, int absCol)
        {
            int r = absRow - FirstRow;
            int c = absCol - FirstCol;
            if (r < 0 || c < 0 || r >= RowCount || c >= ColCount) return null;
            return Formulas[r, c];
        }
    }

    /// <summary>한 워크북(파일) = 시트 목록.</summary>
    public sealed class WorkbookData
    {
        public string FilePath;
        public List<SheetData> Sheets = new List<SheetData>();

        public WorkbookData(string filePath) { FilePath = filePath; }

        public SheetData FindSheet(string name)
        {
            foreach (SheetData s in Sheets)
                if (string.Equals(s.Name, name, StringComparison.OrdinalIgnoreCase))
                    return s;
            return null;
        }
    }

    /// <summary>
    /// 정렬된 표시행 하나. LeftRow/RightRow 는 절대행(1-based), 없으면 -1.
    /// 변경 셀은 Changes(절대열 → 상태)로 보관.
    /// </summary>
    public sealed class DiffRow
    {
        public int LeftRow;
        public int RightRow;
        public RowKind Kind;
        public Dictionary<int, CellStatus> Changes; // null 이면 변경 없음

        public DiffRow(int leftRow, int rightRow)
        {
            LeftRow = leftRow;
            RightRow = rightRow;
            Kind = RowKind.Same;
        }

        public CellStatus StatusAt(int absCol)
        {
            if (Changes == null) return CellStatus.Same;
            CellStatus s;
            if (Changes.TryGetValue(absCol, out s)) return s;
            return CellStatus.Same;
        }
    }

    /// <summary>네비게이션 참조: 표시행 인덱스 + 절대열.</summary>
    public struct DiffNav
    {
        public int RowIndex;
        public int Col;
        public DiffNav(int rowIndex, int col) { RowIndex = rowIndex; Col = col; }
    }

    /// <summary>시트 한 쌍의 비교 결과(정렬 반영).</summary>
    public sealed class SheetDiff
    {
        public string Name;
        public SheetStatus Status;
        public SheetData Left;
        public SheetData Right;

        public int MinCol = 1;
        public int MaxCol = 0;

        public List<DiffRow> Rows = new List<DiffRow>();
        public List<DiffNav> Nav = new List<DiffNav>();

        public int ChangedCount;
        public int AddedCount;
        public int DeletedCount;

        public int ColCount { get { return Math.Max(0, MaxCol - MinCol + 1); } }

        public bool HasChanges
        {
            get { return ChangedCount > 0 || AddedCount > 0 || DeletedCount > 0 || Status != SheetStatus.Same; }
        }
    }

    /// <summary>전체 비교 결과.</summary>
    public sealed class DiffResult
    {
        public string LeftPath;
        public string RightPath;
        public List<SheetDiff> Sheets = new List<SheetDiff>();

        public int TotalChanged;
        public int TotalAdded;
        public int TotalDeleted;
        public int ChangedSheetCount;

        public string Summary()
        {
            return string.Format(
                "시트 {0}개 중 {1}개 변경  |  셀 변경 {2}  추가 {3}  삭제 {4}",
                Sheets.Count, ChangedSheetCount, TotalChanged, TotalAdded, TotalDeleted);
        }
    }
}
