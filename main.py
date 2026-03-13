import os
import re
import time
import win32gui
import win32con
import ctypes
import subprocess
from bs4 import BeautifulSoup

def get_text_from_notepad_memory(file_path):
    """
    메모장을 실행하여 파일 내용을 메모리에서 직접 추출합니다.
    """
    abs_path = os.path.abspath(file_path)
    filename = os.path.basename(file_path)
    
    info = subprocess.STARTUPINFO()
    info.dwFlags |= subprocess.STARTF_USESHOWWINDOW
    info.wShowWindow = win32con.SW_SHOWMINIMIZED
    proc = subprocess.Popen(['notepad.exe', abs_path], startupinfo=info)
    
    hwnd = 0
    start_wait = time.time()
    target_title_part = filename
    
    while time.time() - start_wait < 5:
        def callback(h, extra):
            title = win32gui.GetWindowText(h)
            if target_title_part in title and "메모장" in title:
                extra.append(h)
        hwnds = []
        win32gui.EnumWindows(callback, hwnds)
        if hwnds:
            hwnd = hwnds[0]
            break
        time.sleep(0.2)
    
    if not hwnd:
        proc.terminate()
        return None

    content = ""
    try:
        edit_hwnd = win32gui.FindWindowEx(hwnd, None, "RichEditD2Dpt", None)
        if not edit_hwnd:
            edit_hwnd = win32gui.FindWindowEx(hwnd, None, "Edit", None)
            
        if edit_hwnd:
            length = win32gui.SendMessage(edit_hwnd, win32con.WM_GETTEXTLENGTH, 0, 0)
            buffer = ctypes.create_unicode_buffer(length + 1)
            win32gui.SendMessage(edit_hwnd, win32con.WM_GETTEXT, length + 1, buffer)
            content = buffer.value
    finally:
        win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
        
    return content

def parse_mht_html(html_source):
    """
    추출된 HTML 소스에서 제목, 참석자, 시간을 파싱합니다.
    접두어(제목 :, 참석자... :)를 제거합니다.
    """
    if not html_source:
        return None
        
    soup = BeautifulSoup(html_source, 'lxml')
    
    # 1. 제목 및 참석자 파싱 (chat_title 클래스의 dl)
    chat_title_dl = soup.find('dl', class_='chat_title')
    title = "N/A"
    participants = "N/A"
    
    if chat_title_dl:
        # 제목 추출 및 정제 (첫 번째 dt)
        dt_tag = chat_title_dl.find('dt')
        if dt_tag:
            raw_title = dt_tag.get_text(strip=True)
            # "제목 :" 패턴 제거
            title = re.sub(r'^제목\s*:\s*', '', raw_title).strip()
            
        # 참석자 추출 및 정제 (첫 번째 dd)
        dd_tag = chat_title_dl.find('dd')
        if dd_tag:
            raw_participants = dd_tag.get_text(strip=True)
            # "참석자(숫자) :" 또는 "참석자 :" 패턴 제거
            participants = re.sub(r'^참석자(\(\d+\))?\s*:\s*', '', raw_participants).strip()

    # 2. 대화 시작 시간 파싱
    start_time = "N/A"
    time_wrap = soup.find('div', class_='im_time_wrap')
    if time_wrap:
        im_time = time_wrap.find('div', class_='im_time')
        if im_time:
            corner_c = im_time.find('span', class_='corner_C')
            if corner_c:
                start_time = corner_c.get_text(strip=True)

    return {
        "title": title,
        "participants": participants,
        "start_time": start_time
    }

if __name__ == "__main__":
    import glob
    files = glob.glob("inputs/*.mht")
    if files:
        print(f"파일 처리 중: {files[0]}")
        raw_html = get_text_from_notepad_memory(files[0])
        if raw_html:
            result = parse_mht_html(raw_html)
            print("\n[파싱 결과 - 정제됨]")
            print(f"제목: {result['title']}")
            print(f"참석자: {result['participants']}")
            print(f"시작 시간: {result['start_time']}")
        else:
            print("데이터 추출 실패")
