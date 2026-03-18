@echo off
setlocal enabledelayedexpansion
title MessageManager Setup Wizard

echo ===================================================
echo   MessageManager 초기 설정 마법사
echo ===================================================
echo.

REM 1. 경로 입력 받기
set /p IN_DIR="1. MHT 파일이 있는 폴더 경로 (예: C:\MHT\Inputs): "
set /p OUT_DIR="2. 마크다운 저장 폴더 경로 (예: C:\MHT\Outputs): "
set /p ARC_DIR="3. 처리 완료 파일 백업 폴더 경로 (예: C:\MHT\Archive): "
set /p SCHED_TIME="4. 자동 실행 시간 (24시간 형식, 예: 19:00): "
echo.
echo [옵션] 파이썬 또는 실행 파일의 경로를 직접 지정하시겠습니까? 
echo (그냥 엔터를 치면 기본 'python' 명령어를 사용합니다.)
set /p PY_PATH="5. 실행 파일 경로 (필요 없으면 엔터): "
if "!PY_PATH!"=="" set "PY_PATH=python"

REM JSON용 경로 이스케이프 ( \ -> \\ )
set "JS_IN=!IN_DIR:\=\\!"
set "JS_OUT=!OUT_DIR:\=\\!"
set "JS_ARC=!ARC_DIR:\=\\!"

echo.
echo [1/3] config.json 생성 중...
(
echo {
echo   "input_dir": "!JS_IN!",
echo   "output_dir": "!JS_OUT!",
echo   "archive_dir": "!JS_ARC!",
echo   "data_dir": "data/json",
echo   "log_file": "manager.log",
echo   "max_retries": 3
echo }
) > config.json

echo [2/3] 실행용 배치 파일(run_task.bat) 생성 중...
echo @echo off > run_task.bat
echo cd /d "%%~dp0" >> run_task.bat
echo if exist "MessageManager.exe" ( >> run_task.bat
echo     start "" "MessageManager.exe" >> run_task.bat
echo ) else ( >> run_task.bat
echo     "!PY_PATH!" main.py >> run_task.bat
echo ) >> run_task.bat

echo [3/3] 스케줄러 등록 스크립트(register_task.bat) 생성 중...
echo @echo off > register_task.bat
echo echo 작업 스케줄러에 등록 중입니다... >> register_task.bat
echo schtasks /create /tn "MessageManagerAutoTask" /tr "%%~dp0run_task.bat" /sc daily /st !SCHED_TIME! /f >> register_task.bat
echo if %%ERRORLEVEL%% EQU 0 ( >> register_task.bat
echo     echo. >> register_task.bat
echo     echo [성공] 등록이 완료되었습니다. 매일 !SCHED_TIME!에 실행됩니다. >> register_task.bat
echo ) else ( >> register_task.bat
echo     echo. >> register_task.bat
echo     echo [실패] 권한 문제일 수 있습니다. '관리자 권한'으로 다시 실행해 보세요. >> register_task.bat
echo ) >> register_task.bat
echo pause >> register_task.bat

echo.
echo ===================================================
echo   설정이 완료되었습니다!
echo.
echo   - 사용된 실행 경로: !PY_PATH!
echo   - 자동 실행 시간: !SCHED_TIME!
echo.
echo   1. 'register_task.bat'을 실행하여 스케줄러에 등록하세요.
echo   2. 수동 실행은 'run_task.bat'을 사용하세요.
echo ===================================================
pause
