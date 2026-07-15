@echo off
setlocal enabledelayedexpansion
rem ============================================================
rem  ExcelDiffMerge 현장 빌드 스크립트 (spec §10)
rem  exe 반입 없이 소스만 반입 -> 대상 PC 의 csc.exe 로 직접 빌드.
rem  AhnLab V3 의 exe 반입 심사를 회피(소스는 텍스트).
rem ============================================================

rem 정찰 실측 경로. 없으면 32비트판/기타 경로로 폴백 탐색.
set "CSC=C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
if not exist "%CSC%" set "CSC=C:\Windows\Microsoft.NET\Framework\v4.0.30319\csc.exe"
if not exist "%CSC%" (
  echo [오류] csc.exe 를 찾을 수 없습니다. .NET Framework 4.x 설치 경로를 확인하세요.
  echo   기대 경로: C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe
  pause
  exit /b 1
)

echo [정보] 사용 컴파일러: %CSC%
echo [정보] 빌드 시작...

"%CSC%" /nologo /target:winexe /platform:x64 /out:ExcelDiffMerge.exe ^
  /optimize+ ^
  /reference:System.dll ^
  /reference:System.Core.dll ^
  /reference:System.Data.dll ^
  /reference:System.Windows.Forms.dll ^
  /reference:System.Drawing.dll ^
  /reference:Microsoft.CSharp.dll ^
  src\*.cs

if errorlevel 1 (
  echo.
  echo [실패] 빌드 실패. 위 오류 메시지를 확인하세요.
  echo   - named argument 컴파일 실패 시: 소스는 positional 인자만 사용하므로 무관.
  echo   - dynamic 관련 실패 시: Microsoft.CSharp.dll 참조 확인.
  pause
  exit /b 1
)

echo.
echo [성공] 빌드 완료: ExcelDiffMerge.exe
echo.
echo  실행:            ExcelDiffMerge.exe
echo  COM 열기 PoC:    ExcelDiffMerge.exe --poc "C:\경로\파일.xlsx"
echo  콘솔 2-way diff:  ExcelDiffMerge.exe --diff "좌.xlsx" "우.xlsx"
echo.
echo  * exe 가 AhnLab V3 에 걸리면 관리자에게 V3 예외(신뢰) 등록 요청 (spec §10).
pause
endlocal
