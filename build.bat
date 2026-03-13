@echo off
REM MessageManager 빌드 스크립트 (PyInstaller 필요)

echo [1/3] 의존성 확인...
pip install -r requirements.txt
pip install pyinstaller

echo [2/3] 실행 파일 빌드 시작...
REM --onefile: 단일 파일 생성
REM --noconsole: 실행 시 콘솔 창 숨김 (GUI 환경이나 백그라운드 작업 시 유용)
REM --name: 출력 파일명 지정
pyinstaller --onefile --noconsole --name "MessageManager" main.py

echo [3/3] 빌드 완료!
echo 'dist/MessageManager.exe' 파일과 'config.json'을 같은 폴더에 두고 사용하세요.
pause
