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
    
    # 1. 메모장 실행 (최소화 상태)
    info = subprocess.STARTUPINFO()
    info.dwFlags |= subprocess.STARTF_USESHOWWINDOW
    info.wShowWindow = win32con.SW_SHOWMINIMIZED
    proc = subprocess.Popen(['notepad.exe', abs_path], startupinfo=info)
    
    # 2. 메모장 창 찾기 (최대 5초 대기)
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

    # 3. 텍스트 영역(RichEditD2Dpt 또는 Edit) 찾기 및 추출
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
        # 4. 메모장 종료
        win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
        
    return content

def parse_mht_html(html_source):
    """
    추출된 HTML 소스에서 제목, 참석자, 시간을 파싱합니다.
    """
    if not html_source:
        return None
        
    soup = BeautifulSoup(html_source, 'lxml')
    
    # 1. 제목 및 참석자 파싱 (html/body/dl[@class='chat_title'])
    chat_title_dl = soup.find('dl', class_='chat_title')
    title = "N/A"
    participants = "N/A"
    
    if chat_title_dl:
        # 첫 번째 dt: 제목
        dt_tags = chat_title_dl.find_all('dt')
        if dt_tags:
            title = dt_tags[0].get_text(strip=True)
            
        # 세 번째 dd: 참석자
        dd_tags = chat_title_dl.find_all('dd')
        if len(dd_tags) >= 3:
            participants = dd_tags[2].get_text(strip=True)

    # 2. 대화 시작 시간 파싱
    # im_time_wrap -> im_time -> span.corner_C
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
    # 테스트 실행
    import glob
    files = glob.glob("inputs/*.mht")
    if files:
        print(f"파일 처리 중: {files[0]}")
        raw_html = get_text_from_notepad_memory(files[0])
        if raw_html:
            result = parse_mht_html(raw_html)
            print("\n[파싱 결과]")
            print(f"제목: {result['title']}")
            print(f"참석자: {result['participants']}")
            print(f"시작 시간: {result['start_time']}")
        else:
            print("데이터 추출 실패")
