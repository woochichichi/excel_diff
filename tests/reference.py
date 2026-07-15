#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
가상 테스트 / 알고리즘 레퍼런스.

C# DiffEngine + RowAligner 와 '동일한' 로직을 파이썬으로 미러링한다.
목적:
 - Excel/COM/WinForms 없이 로컬(리눅스 포함)에서 diff 알고리즘의 정확성을 검증.
 - 특히 행 삽입/삭제 정렬(LCS)이 하위 행을 오판(cascade)하지 않는지 확인.
C# 포팅이 이 결과와 일치해야 한다(--selftest 가 동일 케이스를 검증).

실행:  python3 tests/reference.py
"""
import csv
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(HERE, "data")

# ---------------------------------------------------------------- 값 정규화
def parse_cell(text):
    """CSV 텍스트 → (value, formula). 숫자는 float, '='로 시작하면 수식."""
    if text is None:
        return (None, None)
    t = text
    formula = t if t.startswith("=") else None
    if t == "":
        return (None, formula)
    if formula is not None:
        return (t, formula)  # 수식 셀: 값도 수식문자열로 취급(목업)
    # 숫자?
    try:
        f = float(t)
        return (f, None)
    except ValueError:
        return (t, None)

def to_text(v):
    if v is None:
        return ""
    if isinstance(v, float):
        if v == int(v) and abs(v) < 1e15:
            return str(int(v))
        return repr(v)
    return str(v)

def value_equals(a, b):
    if a is None and b is None:
        return True
    if a is None or b is None:
        return False
    if isinstance(a, float) and isinstance(b, float):
        return a == b
    return to_text(a) == to_text(b)

def formula_equals(a, b):
    return (a or "") == (b or "")

def is_empty(v):
    return v is None or (isinstance(v, str) and v == "")

# ---------------------------------------------------------------- 로드
class Sheet:
    def __init__(self, name):
        self.name = name
        self.values = []    # 2D [r][c]
        self.formulas = []
        self.rows = 0
        self.cols = 0

    def val(self, r, c):
        if 0 <= r < self.rows and 0 <= c < self.cols:
            return self.values[r][c]
        return None

    def frm(self, r, c):
        if 0 <= r < self.rows and 0 <= c < self.cols:
            return self.formulas[r][c]
        return None

def load_csv_sheet(path, name):
    with open(path, newline="", encoding="utf-8") as f:
        rows = list(csv.reader(f))
    ncol = max((len(r) for r in rows), default=0)
    s = Sheet(name)
    for r in rows:
        vrow, frow = [], []
        for c in range(ncol):
            cell = r[c] if c < len(r) else ""
            v, fo = parse_cell(cell)
            vrow.append(v)
            frow.append(fo)
        s.values.append(vrow)
        s.formulas.append(frow)
    s.rows = len(rows)
    s.cols = ncol
    return s

def load_workbook(folder):
    sheets = []
    for fn in sorted(os.listdir(folder)):
        if fn.lower().endswith(".csv"):
            name = os.path.splitext(fn)[0]
            sheets.append(load_csv_sheet(os.path.join(folder, fn), name))
    return sheets

# ---------------------------------------------------------------- 셀 상태
SAME, CHANGED, ADDED, DELETED = "Same", "Changed", "Added", "Deleted"

def cell_status(lv, lf, rv, rf):
    le, re_ = is_empty(lv), is_empty(rv)
    if le and re_:
        return SAME
    if le and not re_:
        return ADDED
    if not le and re_:
        return DELETED
    if value_equals(lv, rv) and formula_equals(lf, rf):
        return SAME
    return CHANGED

# ---------------------------------------------------------------- 행 정렬(LCS)
def row_signature(sheet, r, key_col):
    if sheet is None:
        return None
    if key_col is not None:
        return to_text(sheet.val(r, key_col))
    parts = [to_text(sheet.val(r, c)) for c in range(sheet.cols)]
    return "\x01".join(parts)  # 셀 구분자

LCS_ROW_BUDGET = 3000  # 이보다 크면 유사도 LCS 비용 과다 → 키/좌표로 폴백

def row_similarity(left, right, li, rj, ncol):
    """두 행에서 (둘 다 빈 셀 제외) 값이 같은 셀 수 / 채워진 셀 수."""
    eq = 0
    non_empty = 0
    for c in range(ncol):
        lv = left.val(li, c)
        rv = right.val(rj, c)
        le, re_ = is_empty(lv), is_empty(rv)
        if le and re_:
            continue
        non_empty += 1
        if value_equals(lv, rv):
            eq += 1
    return eq, non_empty

def rows_match(left, right, li, rj, ncol, lsig, rsig):
    """정렬 앵커 후보: 완전 동일하거나, 채워진 셀의 절반 이상이 일치."""
    if lsig[li] == rsig[rj]:
        return True
    eq, non_empty = row_similarity(left, right, li, rj, ncol)
    if eq == 0:
        return False
    return eq * 2 >= non_empty  # 채워진 셀의 과반 일치 → 같은 행(수정된 것)으로 간주

def similarity_lcs(left, right, ncol, lsig, rsig):
    """유사도 술어 기반 LCS. 순서 보존 정렬쌍 (i,j) 리스트 반환."""
    n = left.rows
    m = right.rows
    dp = [[0] * (m + 1) for _ in range(n + 1)]
    match = [[False] * m for _ in range(n)]
    for i in range(n - 1, -1, -1):
        for j in range(m - 1, -1, -1):
            if rows_match(left, right, i, j, ncol, lsig, rsig):
                match[i][j] = True
                dp[i][j] = dp[i + 1][j + 1] + 1
            else:
                dp[i][j] = dp[i + 1][j] if dp[i + 1][j] >= dp[i][j + 1] else dp[i][j + 1]
    res = []
    i = j = 0
    while i < n and j < m:
        if match[i][j] and dp[i][j] == dp[i + 1][j + 1] + 1:
            res.append((i, j))
            i += 1
            j += 1
        elif dp[i + 1][j] >= dp[i][j + 1]:
            i += 1
        else:
            j += 1
    return res

def align_rows(left, right, key_col=None, enabled=True):
    """
    반환: DiffRow 리스트. 각 원소 = dict(left_row, right_row, kind).
    left_row/right_row 는 0-based 시트 내부 인덱스, 없으면 -1.
    """
    lrows = left.rows if left else 0
    rrows = right.rows if right else 0

    if not enabled:
        # 좌표 기준: 같은 인덱스끼리. 행 삽입/삭제 인지 못함.
        out = []
        n = max(lrows, rrows)
        for i in range(n):
            lr = i if i < lrows else -1
            rr = i if i < rrows else -1
            out.append(dict(left_row=lr, right_row=rr, kind="?"))
        return out

    if left is None:
        return [dict(left_row=-1, right_row=r, kind=ADDED) for r in range(rrows)]
    if right is None:
        return [dict(left_row=r, right_row=-1, kind=DELETED) for r in range(lrows)]

    ncol = max(left.cols, right.cols)
    lsig = [row_signature(left, r, key_col) for r in range(lrows)]
    rsig = [row_signature(right, r, key_col) for r in range(rrows)]

    if key_col is not None:
        return _keyed_align(lsig, rsig)

    # 유사도 LCS는 O(n*m*cols) → 큰 시트는 좌표 폴백.
    if lrows * rrows > LCS_ROW_BUDGET * LCS_ROW_BUDGET:
        return align_rows(left, right, key_col=None, enabled=False)

    matches = similarity_lcs(left, right, ncol, lsig, rsig)
    out = []
    prev_i = prev_j = 0
    for (mi, mj) in matches + [(lrows, rrows)]:
        # 매치 사이 gap: 왼쪽=삭제, 오른쪽=추가
        for k in range(prev_i, mi):
            out.append(dict(left_row=k, right_row=-1, kind=DELETED))
        for k in range(prev_j, mj):
            out.append(dict(left_row=-1, right_row=k, kind=ADDED))
        if mi < lrows:  # 실제 매치(센티넬 제외) → 정렬쌍(동일 또는 변경)
            out.append(dict(left_row=mi, right_row=mj, kind="?"))
            prev_i = mi + 1
            prev_j = mj + 1
    return out

def _keyed_align(la, ra):
    """키 컬럼 조인 정렬."""
    from collections import defaultdict, deque
    right_by_key = defaultdict(deque)
    for j, k in enumerate(ra):
        right_by_key[k].append(j)
    used_r = set()
    out = []
    for i, k in enumerate(la):
        if right_by_key[k]:
            j = right_by_key[k].popleft()
            used_r.add(j)
            out.append(dict(left_row=i, right_row=j, kind="?"))
        else:
            out.append(dict(left_row=i, right_row=-1, kind=DELETED))
    for j in range(len(ra)):
        if j not in used_r:
            out.append(dict(left_row=-1, right_row=j, kind=ADDED))
    return out

# ---------------------------------------------------------------- 시트 diff
def diff_sheet(left, right, key_col=None, enabled=True):
    ncol = max(left.cols if left else 0, right.cols if right else 0)
    align = align_rows(left, right, key_col, enabled)

    changed = added = deleted = 0
    display = []
    for row in align:
        lr, rr = row["left_row"], row["right_row"]
        changed_cols = []
        for c in range(ncol):
            lv = left.val(lr, c) if (left and lr >= 0) else None
            lf = left.frm(lr, c) if (left and lr >= 0) else None
            rv = right.val(rr, c) if (right and rr >= 0) else None
            rf = right.frm(rr, c) if (right and rr >= 0) else None
            st = cell_status(lv, lf, rv, rf)
            if st == SAME:
                continue
            changed_cols.append((c, st))
            if st == CHANGED:
                changed += 1
            elif st == ADDED:
                added += 1
            elif st == DELETED:
                deleted += 1
        # 행 kind 확정
        if lr < 0:
            kind = ADDED
        elif rr < 0:
            kind = DELETED
        elif changed_cols:
            kind = CHANGED
        else:
            kind = SAME
        display.append((lr, rr, kind, changed_cols))
    return dict(changed=changed, added=added, deleted=deleted, rows=display)

# ---------------------------------------------------------------- 실행/검증
def run(enabled, key_col=None):
    old = {s.name: s for s in load_workbook(os.path.join(DATA, "old"))}
    new = {s.name: s for s in load_workbook(os.path.join(DATA, "new"))}
    names = list(old.keys()) + [n for n in new.keys() if n not in old]
    result = {}
    for name in names:
        result[name] = diff_sheet(old.get(name), new.get(name), key_col, enabled)
    return result

def print_sheet(name, d, old, new):
    print("\n=== [%s]  변경 %d 추가 %d 삭제 %d ===" % (name, d["changed"], d["added"], d["deleted"]))
    for (lr, rr, kind, cols) in d["rows"]:
        if kind == SAME:
            continue
        label = {ADDED: "＋추가", DELETED: "－삭제", CHANGED: "≠변경"}[kind]
        lrs = "L%d" % (lr + 1) if lr >= 0 else "L-"
        rrs = "R%d" % (rr + 1) if rr >= 0 else "R-"
        detail = ", ".join("col%d[%s]" % (c + 1, st) for (c, st) in cols)
        print("   %-5s %s/%s  %s" % (label, lrs, rrs, detail))

def main():
    print("################ 정렬 ON (LCS) ################")
    r_on = run(enabled=True)
    old = {s.name: s for s in load_workbook(os.path.join(DATA, "old"))}
    new = {s.name: s for s in load_workbook(os.path.join(DATA, "new"))}
    for name, d in r_on.items():
        print_sheet(name, d, old, new)

    print("\n\n################ 정렬 OFF (좌표) ################")
    r_off = run(enabled=False)
    for name, d in r_off.items():
        print_sheet(name, d, old, new)

    # ---- 기대값 검증(가상 테스트 assert) ----
    print("\n\n################ ASSERT ################")
    ok = True
    def check(cond, msg):
        nonlocal ok
        print(("  PASS " if cond else "  FAIL ") + msg)
        ok = ok and cond

    s1 = r_on["Sheet1"]
    # Sheet1 정렬 ON 기대:
    #  Apple: Price 변경(1) → changed
    #  Blueberry: 행 추가 → 5개 열 중 비지 않은 셀만 added
    #  Cherry: Qty 변경 + Note 추가 → changed 1, added 1
    #  Date: 행 삭제 → 비지 않은 셀 added? no, deleted
    #  Elderberry: Note 삭제 → deleted 1
    # 삽입/삭제로 인한 cascade 오판이 없어야 한다.
    kinds = [k for (_, _, k, _) in s1["rows"]]
    check(ADDED in kinds, "Sheet1: 삽입행(Blueberry) 을 추가로 인식")
    check(DELETED in kinds, "Sheet1: 삭제행(Date) 을 삭제로 인식")
    # cascade 방지: 변경 셀 수가 과하지 않아야(<= 5)
    check(s1["changed"] <= 5, "Sheet1: cascade 오판 없음 (changed=%d)" % s1["changed"])

    s1off = r_off["Sheet1"]
    check(s1off["changed"] > s1["changed"],
          "정렬 OFF 는 cascade 로 변경 수가 더 많다 (off=%d > on=%d)"
          % (s1off["changed"], s1["changed"]))

    s2 = r_on["Sheet2"]
    # Sheet2: A Tax 수식 변경, B Base 값 변경
    check(s2["changed"] == 2, "Sheet2: 수식/값 변경 2건 (got %d)" % s2["changed"])

    check("OnlyNew" in r_on and r_on["OnlyNew"]["added"] >= 1, "OnlyNew: 추가 시트 인식")
    check("OnlyOld" in r_on and r_on["OnlyOld"]["deleted"] >= 1, "OnlyOld: 삭제 시트 인식")

    # 키 컬럼(ID=col0) 정렬도 동일하게 삽입/삭제 인식
    r_key = run(enabled=True, key_col=0)
    s1k = r_key["Sheet1"]
    kk = [k for (_, _, k, _) in s1k["rows"]]
    check(ADDED in kk and DELETED in kk, "Sheet1(key): 키기준 삽입/삭제 인식")

    print("\n결과:", "ALL PASS ✅" if ok else "FAIL ❌")
    sys.exit(0 if ok else 1)

if __name__ == "__main__":
    main()
