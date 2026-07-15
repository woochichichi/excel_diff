using System;
using System.Collections.Generic;

namespace ExcelDiffMerge
{
    /// <summary>
    /// 셀 비교 결과 상태.
    /// </summary>
    public enum CellStatus
    {
        Same = 0,     // 동일
        Changed = 1,  // 양쪽 값 다름 (노랑)
        Added = 2,    // new(우)에만 존재 (초록)
        Deleted = 3,  // old(좌)에만 존재 (빨강)
        Conflict = 4  // N-way: 여러 버전이 같은 셀을 다르게 변경 (주황)
    }

    /// <summary>
    /// 시트 단위 비교 상태.
    /// </summary>
    public enum SheetStatus
    {
        Same = 0,
        Modified = 1,
        Added = 2,
        Deleted = 3
    }

    /// <summary>
    /// 한 시트를 UsedRange.Value / .Formula 로 일괄 로드한 결과.
    /// COM 왕복을 최소화하기 위해 값은 전부 2D 배열로 한 번에 담는다.
    /// 좌표는 시트 절대좌표(1-based)를 기준으로 저장한다.
    /// </summary>
    public sealed class SheetData
    {
        public string Name;

        // UsedRange 의 절대 시작 위치(1-based). 두 파일의 UsedRange 시작이 달라도
        // 절대좌표로 정렬해 비교할 수 있게 보관한다.
        public int FirstRow;
        public int FirstCol;

        public int RowCount;
        public int ColCount;

        // 0-based 로 정규화된 값/수식 배열. [r, c] = UsedRange 내부 상대좌표.
        public object[,] Values;
        public object[,] Formulas;

        public SheetData(string name)
        {
            Name = name;
            FirstRow = 1;
            FirstCol = 1;
            RowCount = 0;
            ColCount = 0;
            Values = new object[0, 0];
            Formulas = new object[0, 0];
        }

        /// <summary>절대좌표(1-based) 기준 마지막 행.</summary>
        public int LastRow { get { return FirstRow + RowCount - 1; } }

        /// <summary>절대좌표(1-based) 기준 마지막 열.</summary>
        public int LastCol { get { return FirstCol + ColCount - 1; } }

        /// <summary>
        /// 절대좌표(1-based)로 값을 조회. 범위 밖이면 null.
        /// </summary>
        public object GetValueAbs(int absRow, int absCol)
        {
            int r = absRow - FirstRow;
            int c = absCol - FirstCol;
            if (r < 0 || c < 0 || r >= RowCount || c >= ColCount) return null;
            return Values[r, c];
        }

        /// <summary>
        /// 절대좌표(1-based)로 수식을 조회. 범위 밖이면 null.
        /// </summary>
        public object GetFormulaAbs(int absRow, int absCol)
        {
            int r = absRow - FirstRow;
            int c = absCol - FirstCol;
            if (r < 0 || c < 0 || r >= RowCount || c >= ColCount) return null;
            return Formulas[r, c];
        }
    }

    /// <summary>
    /// 한 워크북(파일) 전체 = 시트 목록.
    /// </summary>
    public sealed class WorkbookData
    {
        public string FilePath;
        public List<SheetData> Sheets = new List<SheetData>();

        public WorkbookData(string filePath)
        {
            FilePath = filePath;
        }

        public SheetData FindSheet(string name)
        {
            foreach (SheetData s in Sheets)
            {
                if (string.Equals(s.Name, name, StringComparison.OrdinalIgnoreCase))
                    return s;
            }
            return null;
        }
    }

    /// <summary>
    /// 셀 하나의 diff 정보(절대좌표 기준).
    /// </summary>
    public sealed class DiffCell
    {
        public int Row;   // 절대좌표 1-based
        public int Col;   // 절대좌표 1-based
        public CellStatus Status;

        public object OldValue;
        public object NewValue;
        public object OldFormula;
        public object NewFormula;

        public DiffCell(int row, int col, CellStatus status)
        {
            Row = row;
            Col = col;
            Status = status;
        }
    }

    /// <summary>
    /// 시트 한 쌍의 비교 결과.
    /// </summary>
    public sealed class SheetDiff
    {
        public string Name;
        public SheetStatus Status;

        // 표시용 그리드 경계(절대좌표 union). 1-based.
        public int MinRow = 1;
        public int MinCol = 1;
        public int MaxRow = 0;
        public int MaxCol = 0;

        // 좌/우 원본 시트 데이터(그리드 채우기용). Added/Deleted 시 한쪽이 null.
        public SheetData Left;
        public SheetData Right;

        // 상태가 Same 이 아닌 셀만 담는다.
        public List<DiffCell> Cells = new List<DiffCell>();

        // 카운트
        public int ChangedCount;
        public int AddedCount;
        public int DeletedCount;

        public bool HasChanges
        {
            get { return ChangedCount > 0 || AddedCount > 0 || DeletedCount > 0 || Status != SheetStatus.Same; }
        }

        // (row,col) → DiffCell 빠른 조회.
        private Dictionary<long, DiffCell> _index;

        public void BuildIndex()
        {
            _index = new Dictionary<long, DiffCell>(Cells.Count);
            foreach (DiffCell c in Cells)
                _index[Key(c.Row, c.Col)] = c;
        }

        public DiffCell GetCell(int row, int col)
        {
            if (_index == null) BuildIndex();
            DiffCell dc;
            if (_index.TryGetValue(Key(row, col), out dc)) return dc;
            return null;
        }

        private static long Key(int row, int col)
        {
            return ((long)row << 20) ^ (long)col;
        }
    }

    /// <summary>
    /// 전체 비교 결과.
    /// </summary>
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
            int total = Sheets.Count;
            return string.Format(
                "시트 {0}개 중 {1}개 변경  |  셀 변경 {2}  추가 {3}  삭제 {4}",
                total, ChangedSheetCount, TotalChanged, TotalAdded, TotalDeleted);
        }
    }
}
