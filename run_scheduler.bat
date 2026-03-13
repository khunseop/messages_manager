@echo off
REM Windows 작업 스케줄러 등록용 실행 스크립트
REM 이 스크립트는 백그라운드 작업 시 경로 문제를 방지하기 위해 
REM 현재 배치 파일이 있는 경로로 이동한 뒤 파이썬 스크립트를 실행합니다.

cd /d "%~dp0"

REM 파이썬 가상환경을 사용하는 경우 아래의 주석을 풀고 경로를 맞춰주세요.
REM call venv\Scripts\activate.bat

echo [%DATE% %TIME%] Starting MessageManager...
python main.py
