====================================================================
 ExcelDiffMerge — 엑셀 Diff/Merge 툴 (WinMerge 식)   v1.1.0
 폐쇄망 + 소프트캠프 DRM 환경 / C# WinForms + Excel COM late binding
 (버전 이력은 CHANGELOG.md 참고)
====================================================================

■ 개요
  여러 버전의 엑셀 파일을 Excel COM 을 통해 읽어(= DRM 은 Excel 이 복호화)
  셀 단위로 비교(diff)하고, 좌우 병렬 뷰로 변경점을 시각화한 뒤,
  원하는 값을 선택 병합(merge)해 저장하는 툴입니다.

  * DRM 우회/해제 없음. 파일 바이너리 직접 파싱 없음.
    Excel 프로세스가 복호화한 값을 COM 으로 읽기만 합니다.


■ 반입물 (텍스트만 — exe 반입 안 함)
  반입폴더\
  ├─ src\               ... C# 소스 (.cs) 전부
  │   ├─ Program.cs         진입점 (GUI / --poc / --diff / --difcsv / --selftest)
  │   ├─ Models.cs          데이터 모델 (정렬 기반 DiffRow)
  │   ├─ IWorkbookReader.cs 로더 추상화(운영=COM, 테스트=CSV)
  │   ├─ ExcelComReader.cs  Excel COM 리더 (UsedRange 일괄 로드 + COM 정리)
  │   ├─ CsvWorkbookReader.cs 로컬 테스트용 CSV 로더(Excel 불필요)
  │   ├─ StaTask.cs         COM 실행용 STA 전용 스레드 헬퍼
  │   ├─ DiffEngine.cs      2-way diff 엔진(정렬 반영)
  │   ├─ RowAligner.cs      행 정렬(좌표/유사도LCS/키컬럼) — 삽입/삭제 cascade 방지
  │   ├─ NWayDiffEngine.cs  N-way 취합 엔진
  │   ├─ MergeEngine.cs     COM 쓰기/저장 (배치 + 백업)
  │   ├─ Settings.cs        settings.ini 저장/복원
  │   ├─ MarkerBar.cs       변경 밀도 마커바(우측)
  │   ├─ MainForm.cs        메인 UI (좌우 그리드/동기 스크롤/색상/네비/필터/병합/상세)
  │   ├─ NWayForm.cs        N-way 취합 UI
  │   ├─ SelfTest.cs        내장 자기검증(--selftest)
  │   └─ AssemblyInfo.cs    버전 정보
  ├─ tests\             ... 로컬 테스트 자료(선택 반입)
  │   ├─ reference.py       파이썬 가상테스트(알고리즘 레퍼런스)
  │   └─ data\old, data\new 목업 CSV
  ├─ build.bat          csc 현장 빌드 스크립트
  ├─ VERSION / CHANGELOG.md 버전/이력
  └─ README.txt         본 파일


■ 빌드 (대상 PC 에서)
  1) 반입폴더에서 build.bat 더블클릭 (또는 cmd 에서 실행).
     - 정찰 실측 csc 경로 사용:
       C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe
     - /platform:x64 (AMD64), Microsoft.CSharp.dll 참조(dynamic 용).
  2) 성공 시 ExcelDiffMerge.exe 생성.
  3) exe 가 AhnLab V3 휴리스틱에 걸릴 수 있음 → 관리자에게 V3 예외(신뢰) 등록 요청.
     소스가 이미 반입·검토된 상태라 정당성 입증이 쉽습니다.


■ 실행
  · GUI:             ExcelDiffMerge.exe
  · COM 열기 PoC:     ExcelDiffMerge.exe --poc "C:\경로\파일.xlsx"
       파일 1개를 ReadOnly 로 열어 시트/크기/미리보기를 콘솔 출력.
       (winexe 지만 부모 콘솔에 attach 하여 cmd 에서 출력이 보입니다.)
  · 콘솔 2-way diff:   ExcelDiffMerge.exe --diff "좌.xlsx" "우.xlsx"


■ 로컬 테스트 (Excel/DRM 없이 — 폐쇄망 반입 부담 없이 로직 검증)
  실제 DRM 파일을 계속 폐쇄망으로 옮겨 테스트하기 어려우므로, Excel 없이
  로컬(개발 PC/리눅스 포함)에서 diff·정렬 로직을 검증할 수 있게 했습니다.

  1) 내장 자기검증 (권장, 어디서나):
       ExcelDiffMerge.exe --selftest
     → 메모리 목업으로 DiffEngine/RowAligner 검증. "ALL PASS" 확인.
       현장 빌드 직후 대상 PC 에서도 즉시 확인 가능.

  2) CSV 폴더 비교(GUI): [CSV폴더비교] → 좌/우 CSV 폴더 선택.
     각 .csv 파일이 하나의 시트로 로드됩니다. tests\data\old, tests\data\new 사용.

  3) CSV 폴더 비교(콘솔):
       ExcelDiffMerge.exe --difcsv tests\data\old tests\data\new

  4) 파이썬 가상테스트(빌드 불필요, 알고리즘 레퍼런스):
       python3 tests\reference.py
     → C# RowAligner 와 동일 알고리즘을 검증. 행 삽입/삭제 cascade 방지 확인.


■ 행 정렬 방식 (툴바 '정렬' 선택 — spec §5-4 해결)
  · 자동정렬(LCS)  : 유사도 기반. 행 삽입/삭제를 인지하고, 수정된 행은 '변경'으로
                     짝지어 이후 행이 전부 "변경"으로 오판되는 문제(cascade)를 방지.
                     (O(nL*nR*열) → 아주 큰 시트는 자동으로 좌표 폴백)
  · 키 컬럼        : '키열' 에 기준 열문자(예: A) 입력. 키 값 기준 조인(O(n),
                     재정렬/대용량에 강함). ID 같은 고유 열이 있으면 가장 정확.
  · 좌표          : 같은 위치끼리 단순 비교(가장 빠름, 행 삽입/삭제 미인지).


■ GUI 사용법
  1) [좌측 열기]/[우측 열기] 로 파일 2개 지정 (또는 창에 드래그앤드롭).
  2) [비교] → 좌우 병렬 그리드에 차이 표시.
       노랑=변경  초록=추가  빨강=삭제  파랑=병합 적용됨
  3) F7/F8 또는 [다음차이]/[이전차이] 로 차이 셀 이동.
  4) [변경만 보기] 로 동일 행 접기.
  5) 셀 선택 후 [좌→우 복사]/[우→좌 복사] 로 병합 대기 등록 → [저장].
       저장 전 백업 복사본 생성 여부를 묻습니다(권장: 예).
  6) [N-way…] 로 base 1개 + 다버전 취합 비교(충돌은 주황). 값 채택 후 base 에 저장.


■ 성능/안정성 설계 (요약)
  · UsedRange.Value/.Formula 를 2D 배열로 "한 번에" 로드 (셀 단위 COM 왕복 금지).
  · 열기 직후 ScreenUpdating=off, Calculation=manual, EnableEvents=off → 종료 시 원복.
  · 모든 COM 객체 Marshal.ReleaseComObject 역순 해제 + finally Quit + GC → 좀비 EXCEL.EXE 방지.
  · 우리 전용 Excel 인스턴스만 사용(사용자가 열어둔 Excel 오염 안 함).
  · 비교용 Open 은 ReadOnly, 매크로 자동실행 차단(AutomationSecurity=ForceDisable).
  · 로드/비교는 백그라운드 워커 → UI 블로킹 방지.


■ 현장 실측 검증 항목 (spec §11 — DRM 실파일로 확인)
  1) 소프트캠프 저장 동작: COM 수정·저장 시 DRM 재적용 정상 여부.
     열람권한만 있는 파일은 저장 실패 가능 → merge 대상은 편집권한 있는 파일만.
  2) csc 언어버전: 본 소스는 named argument 미사용(positional) 이라 하위호환 안전.
  3) AhnLab V3 반응: 빌드 exe 가 걸리면 예외 등록.
  4) DRM 권한없는 파일 Open 예외: 메시지/HRESULT 수집 → 안내 문구 매핑.


■ 알려진 1차 한계 (spec 로드맵)
  · 행 삽입/삭제 시 좌표 기준 비교라 이후 행이 "변경"으로 보일 수 있음
    → 2차: key 컬럼 매칭 / LCS 행 정렬 예정.
  · 서식(배경/폰트/숫자포맷) diff 는 2차. 현재는 표시값 + 수식 비교.
