# 변경 이력 (ExcelDiffMerge)

버전관리는 유의적 버전(Semantic Versioning)을 따른다: MAJOR.MINOR.PATCH.
빌드 산출물의 어셈블리 버전(`src/AssemblyInfo.cs`)과 `VERSION` 파일을 릴리스마다 함께 갱신한다.

---

## [1.1.0] - 2026-07-15

로컬 테스트 가능화 + 행 정렬(cascade 방지) + UX 강화 + 시장/오픈소스 조사 반영.

### 추가 (Added)
- **행 정렬 엔진(`RowAligner`)** — 행 삽입/삭제로 하위 행이 전부 "변경"으로 오판되는
  cascade 문제 해결(spec §5-4). 3가지 모드:
  - `자동정렬(LCS)`: 유사도 기반 LCS. "채워진 셀의 과반 일치"를 앵커로 삼아
    삽입·삭제를 인지하면서 수정된 행은 delete+insert 가 아닌 '변경'으로 짝지음.
  - `키 컬럼`: 키 값 기준 조인(O(n), 대용량/재정렬에 강함).
  - `좌표`: 기존 방식(가장 빠름, 삽입/삭제 미인지).
- **로컬 테스트 인프라** (Excel/COM/DRM 불필요):
  - `CsvWorkbookReader`: CSV 폴더(각 .csv=시트)를 동일 `WorkbookData` 로 로드.
  - GUI `[CSV폴더비교]`, CLI `--difcsv <좌폴더> <우폴더>`.
  - CLI `--selftest`: 메모리 목업으로 DiffEngine/RowAligner 자기검증(현장에서 즉시 확인).
  - `tests/reference.py`: 파이썬 가상테스트(알고리즘 레퍼런스, 리눅스에서도 실행 가능).
  - `tests/data/` 목업 CSV(변경/추가/삭제/행삽입/행삭제/수식/시트추가·삭제 케이스).
  - `tests/make_xlsx.py` + `tests/data/old.xlsx`,`new.xlsx`: 외부 라이브러리 0 으로
    생성한 실제 .xlsx 목업(진짜 수식 포함) → Excel COM 경로/DRM 흐름 로컬 테스트.
- **UX (시장조사 반영)**:
  - 셀 상세 패널(하단): 선택 셀의 좌/우 값·수식·상태 표시.
  - 색상 범례 바(변경/추가/삭제/병합됨/정렬빈칸).
  - 변경 마커바(우측 스크롤): 차이 밀도 시각화 + 클릭 점프.
  - 정렬 방식 선택 콤보 + 키 컬럼 입력.
  - 단축키: F7/F8(다음/이전 차이), Ctrl+S(저장), Ctrl+O(열기).
  - 시트 탭에 변경 건수/추가·삭제 표기.

### 변경 (Changed)
- **COM 을 STA 전용 스레드에서 실행**(`StaTask`) — 기존 `BackgroundWorker`(MTA)는
  간헐적 RPC/마샬링 오류 위험(조사 반영). 로드/저장을 STA 스레드로 이동, UI 는 Invoke 마샬링.
- **diff 모델을 정렬 기반으로 재설계** — `SheetDiff` 가 좌표 셀 리스트 대신
  정렬된 표시행(`DiffRow`, 좌/우 행 개별) 을 보유. 좌우 그리드가 논리적으로 같은
  레코드를 같은 화면행에 표시.
- `Workbooks.Open` 파라미터 강화: `IgnoreReadOnlyRecommended=true`, `Notify=false`
  로 '이미 열림/읽기전용 권장' 프롬프트 선제 차단.
- `AutomationSecurity` 는 프로세스 전역이므로 종료 시 원복.
- `Interactive=false` 추가.

### 문서 (Docs)
- README 에 테스트/정렬모드/버전 안내 추가.

---

## [1.0.0] - 2026-07-15

초기 구현(C# WinForms + Excel COM late binding).

### 추가
- Excel COM 리더(UsedRange 일괄 로드, 역순 ReleaseComObject, 좀비 방지).
- 2-way diff(시트 매칭, 표시값+수식 비교), N-way 취합(충돌 감지).
- 좌우 병렬 그리드(가상모드), 색상, 동기 스크롤, 시트 탭, F7/F8 네비, 변경만 보기.
- 방향 병합/저장(백업), settings.ini, 드래그앤드롭, 온보딩.
- 현장 csc 빌드(`build.bat`, winexe/x64), README.
