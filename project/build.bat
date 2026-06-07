@echo off
chcp 65001 >nul
echo ============================================
echo  안전점검 모니터링 시스템 EXE 빌드 시작...
echo ============================================

REM 이전 빌드 정리
if exist dist rmdir /s /q dist
if exist build rmdir /s /q build

REM PyInstaller 빌드 (onedir)
pyinstaller 안전점검_모니터링.spec

if %errorlevel% neq 0 (
    echo.
    echo [실패] 빌드 실패! 위 오류 메시지를 확인하세요.
    pause
    exit /b 1
)

REM ----------------------------------------------------------------
REM 후처리: config/ data/ logs/ results/ 를 EXE 옆(외부)으로 복사
REM   frozen 시 BASE_DIR = sys.executable 의 폴더 = dist\안전점검_모니터링\
REM   이므로 이 폴더들이 EXE 와 같은 위치에 있어야 경로가 일치한다.
REM ----------------------------------------------------------------
set "DEST=dist\안전점검_모니터링"

echo.
echo config 폴더 복사 중...
xcopy /e /i /y "config" "%DEST%\config" >nul

echo 빈 작업 폴더 생성 중 (data, logs, results)...
if not exist "%DEST%\data"    mkdir "%DEST%\data"
if not exist "%DEST%\logs"    mkdir "%DEST%\logs"
if not exist "%DEST%\results" mkdir "%DEST%\results"

echo install.bat 복사 중...
if exist "install.bat" copy /y "install.bat" "%DEST%\install.bat" >nul

echo .env 템플릿 확인 중...
if exist ".env" (
    copy /y ".env" "%DEST%\.env" >nul
) else (
    echo .env 파일이 없습니다. 배포 후 사용자가 직접 생성해야 합니다.
)

echo.
echo ============================================
echo  빌드 성공!
echo  배포 패키지: %DEST%\
echo ============================================
echo  다음 단계(선택): python deploy\make_release.py
echo  -> 날짜 기반 릴리스 폴더로 정리합니다.
echo ============================================
pause
