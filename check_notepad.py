import os
import subprocess
import time
import win32gui
import win32con
import win32api

def get_text_from_notepad(file_path):
    """
    메모장을 최소화 상태로 열어 내부 텍스트를 직접 읽어옵니다.
    """
    abs_path = os.path.abspath(file_path)
    
    # 1. 메모장을 최소화 상태로 실행
    # SW_SHOWMINIMIZED = 2
    info = subprocess.STARTUPINFO()
    info.dwFlags |= subprocess.STARTF_USESHOWWINDOW
    info.wShowWindow = win32con.SW_SHOWMINIMIZED
    
    proc = subprocess.Popen(['notepad.exe', abs_path], startupinfo=info)
    
    # 2. 메모장 창이 뜰 때까지 대기
    content = ""
    timeout = 5
    start_time = time.time()
    
    while time.time() - start_time < timeout:
        # 메모장 클래스 이름은 "Notepad", 그 안의 텍스트 편집창은 "Edit" (또는 "RichEditD2DPT")
        hwnd = win32gui.FindWindow("Notepad", None)
        if hwnd:
            # 메모장 내부의 Edit 컨트롤 찾기
            edit_hwnd = win32gui.FindWindowEx(hwnd, None, "Edit", None)
            if not edit_hwnd:
                # 윈도우 11 최신 버전 메모장은 구조가 다를 수 있음
                edit_hwnd = win32gui.FindWindowEx(hwnd, None, "RichEditD2DPT", None)
            
            if edit_hwnd:
                # WM_GETTEXTLENGTH로 길이 파악
                text_len = win32gui.SendMessage(edit_hwnd, win32con.WM_GETTEXTLENGTH, 0, 0)
                if text_len > 0:
                    # 버퍼 생성 후 텍스트 가져오기
                    buffer = win32gui.PyMakeBuffer(text_len + 1)
                    win32gui.SendMessage(edit_hwnd, win32con.WM_GETTEXT, text_len + 1, buffer)
                    content = buffer[:text_len].tobytes().decode('utf-16', errors='ignore')
                    break
        time.sleep(0.2)
    
    # 3. 메모장 종료
    proc.terminate()
    
    return content

if __name__ == "__main__":
    import glob
    files = glob.glob("inputs/*.mht")
    if files:
        print(f"Testing Notepad memory read: {files[0]}")
        res = get_text_from_notepad(files[0])
        if res and ("html" in res.lower() or "<body" in res.lower()):
            print("✅ Success! Found HTML tags.")
            print(f"Sample: {res[:200]}...")
        else:
            print("❌ Failed to read or content empty.")
    else:
        print("No files found.")
