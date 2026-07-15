using System;
using System.Collections.Generic;

namespace ExcelDiffMerge
{
    /// <summary>정렬 결과의 행 쌍(절대행 1-based, 없으면 -1).</summary>
    public struct AlignPair
    {
        public int LeftRow;
        public int RightRow;
        public AlignPair(int l, int r) { LeftRow = l; RightRow = r; }
    }

    /// <summary>
    /// 행 정렬기 (spec §5-4). tests/reference.py 로 검증된 알고리즘의 C# 포팅.
    ///  - Coordinate: 좌표 기준(같은 인덱스끼리). 행 삽입/삭제 미인지.
    ///  - Auto: 유사도 LCS. "채워진 셀의 과반 일치" 를 앵커로 삼아 삽입/삭제를 인지하고,
    ///          동시에 수정된 행도 delete+insert 가 아닌 '변경'으로 짝지음(cascade 방지).
    ///  - KeyColumn: 키 컬럼 값 기준 조인(대용량/재정렬에 강함, O(n)).
    /// Auto 는 O(nL*nR*cols) 라 큰 시트에선 좌표로 폴백(LcsRowBudget).
    /// </summary>
    public static class RowAligner
    {
        public const int LcsRowBudget = 3000;

        public static List<AlignPair> Align(SheetData left, SheetData right,
            AlignMode mode, int keyCol, int minCol, int maxCol)
        {
            int lrows = left != null ? left.RowCount : 0;
            int rrows = right != null ? right.RowCount : 0;

            List<AlignPair> outp = new List<AlignPair>();

            if (left == null && right == null) return outp;
            if (left == null)
            {
                for (int r = 0; r < rrows; r++) outp.Add(new AlignPair(-1, right.FirstRow + r));
                return outp;
            }
            if (right == null)
            {
                for (int r = 0; r < lrows; r++) outp.Add(new AlignPair(left.FirstRow + r, -1));
                return outp;
            }

            if (mode == AlignMode.Coordinate)
                return AlignCoordinate(left, right);

            if (mode == AlignMode.KeyColumn && keyCol >= 0)
                return AlignKeyed(left, right, keyCol);

            // Auto (유사도 LCS) — 예산 초과 시 좌표 폴백.
            long cost = (long)lrows * rrows;
            if (cost > (long)LcsRowBudget * LcsRowBudget)
                return AlignCoordinate(left, right);

            return AlignSimilarityLcs(left, right, minCol, maxCol);
        }

        private static List<AlignPair> AlignCoordinate(SheetData left, SheetData right)
        {
            List<AlignPair> outp = new List<AlignPair>();
            int n = Math.Max(left.RowCount, right.RowCount);
            for (int i = 0; i < n; i++)
            {
                int lr = i < left.RowCount ? left.FirstRow + i : -1;
                int rr = i < right.RowCount ? right.FirstRow + i : -1;
                outp.Add(new AlignPair(lr, rr));
            }
            return outp;
        }

        private static List<AlignPair> AlignKeyed(SheetData left, SheetData right, int keyCol)
        {
            // 오른쪽 키 → 행 인덱스 큐(중복 키 순서 보존).
            Dictionary<string, Queue<int>> rightByKey = new Dictionary<string, Queue<int>>();
            for (int j = 0; j < right.RowCount; j++)
            {
                string k = DiffEngine.ToText(right.GetValueAbs(right.FirstRow + j, keyCol));
                Queue<int> q;
                if (!rightByKey.TryGetValue(k, out q)) { q = new Queue<int>(); rightByKey[k] = q; }
                q.Enqueue(j);
            }
            bool[] usedR = new bool[right.RowCount];
            List<AlignPair> outp = new List<AlignPair>();
            for (int i = 0; i < left.RowCount; i++)
            {
                string k = DiffEngine.ToText(left.GetValueAbs(left.FirstRow + i, keyCol));
                Queue<int> q;
                if (rightByKey.TryGetValue(k, out q) && q.Count > 0)
                {
                    int j = q.Dequeue();
                    usedR[j] = true;
                    outp.Add(new AlignPair(left.FirstRow + i, right.FirstRow + j));
                }
                else
                {
                    outp.Add(new AlignPair(left.FirstRow + i, -1));
                }
            }
            for (int j = 0; j < right.RowCount; j++)
                if (!usedR[j]) outp.Add(new AlignPair(-1, right.FirstRow + j));
            return outp;
        }

        private static List<AlignPair> AlignSimilarityLcs(SheetData left, SheetData right, int minCol, int maxCol)
        {
            int n = left.RowCount;
            int m = right.RowCount;

            string[] lsig = new string[n];
            string[] rsig = new string[m];
            for (int i = 0; i < n; i++) lsig[i] = Signature(left, left.FirstRow + i, minCol, maxCol);
            for (int j = 0; j < m; j++) rsig[j] = Signature(right, right.FirstRow + j, minCol, maxCol);

            // dp[i,j] = i..n / j..m 구간의 최대 매치 수.
            int[,] dp = new int[n + 1, m + 1];
            bool[,] match = new bool[n, m];
            for (int i = n - 1; i >= 0; i--)
            {
                for (int j = m - 1; j >= 0; j--)
                {
                    if (RowsMatch(left, right, left.FirstRow + i, right.FirstRow + j, minCol, maxCol, lsig[i], rsig[j]))
                    {
                        match[i, j] = true;
                        dp[i, j] = dp[i + 1, j + 1] + 1;
                    }
                    else
                    {
                        dp[i, j] = dp[i + 1, j] >= dp[i, j + 1] ? dp[i + 1, j] : dp[i, j + 1];
                    }
                }
            }

            // 백트래킹으로 순서 보존 매치 추출.
            List<int[]> matches = new List<int[]>();
            {
                int i = 0, j = 0;
                while (i < n && j < m)
                {
                    if (match[i, j] && dp[i, j] == dp[i + 1, j + 1] + 1)
                    {
                        matches.Add(new int[] { i, j });
                        i++; j++;
                    }
                    else if (dp[i + 1, j] >= dp[i, j + 1]) i++;
                    else j++;
                }
            }

            List<AlignPair> outp = new List<AlignPair>();
            int prevI = 0, prevJ = 0;
            // 센티넬 포함 순회.
            for (int idx = 0; idx <= matches.Count; idx++)
            {
                int mi = idx < matches.Count ? matches[idx][0] : n;
                int mj = idx < matches.Count ? matches[idx][1] : m;
                for (int k = prevI; k < mi; k++) outp.Add(new AlignPair(left.FirstRow + k, -1)); // 삭제
                for (int k = prevJ; k < mj; k++) outp.Add(new AlignPair(-1, right.FirstRow + k)); // 추가
                if (idx < matches.Count)
                {
                    outp.Add(new AlignPair(left.FirstRow + mi, right.FirstRow + mj)); // 동일/변경
                    prevI = mi + 1;
                    prevJ = mj + 1;
                }
            }
            return outp;
        }

        private static string Signature(SheetData s, int absRow, int minCol, int maxCol)
        {
            System.Text.StringBuilder sb = new System.Text.StringBuilder();
            for (int c = minCol; c <= maxCol; c++)
            {
                sb.Append(DiffEngine.ToText(s.GetValueAbs(absRow, c)));
                sb.Append('\u0001'); // 셀 구분자(제어문자)
            }
            return sb.ToString();
        }

        private static bool RowsMatch(SheetData left, SheetData right, int lAbs, int rAbs,
            int minCol, int maxCol, string lsig, string rsig)
        {
            if (lsig == rsig) return true;
            int eq = 0, nonEmpty = 0;
            for (int c = minCol; c <= maxCol; c++)
            {
                object lv = left.GetValueAbs(lAbs, c);
                object rv = right.GetValueAbs(rAbs, c);
                bool le = ExcelComReader.IsEmpty(lv);
                bool re = ExcelComReader.IsEmpty(rv);
                if (le && re) continue;
                nonEmpty++;
                if (DiffEngine.ValueEquals(lv, rv)) eq++;
            }
            if (eq == 0) return false;
            return eq * 2 >= nonEmpty; // 채워진 셀의 과반 일치 → 같은 행(수정)으로 간주
        }
    }
}
