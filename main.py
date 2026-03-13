import os
import re
import time
import win32gui
import win32con
import ctypes
import subprocess
import html
import win32process
from bs4 import BeautifulSoup

def get_text_from_notepad_hidden(file_path):
    """
    PID를 기반으로 특정 메모장 프로세스의 텍스트를 정밀하게 추출합니다.
    """
    abs_path = os.path.abspath(file_path)
    filename = os.path.basename(file_path)
    
    # 1. 메모장 실행
    info = subprocess.STARTUPINFO()
    info.dwFlags |= subprocess.STARTF_USESHOWWINDOW
    info.wShowWindow = win32con.SW_HIDE
    try:
        proc = subprocess.Popen(['notepad.exe', abs_path], startupinfo=info)
        target_pid = proc.pid
    except Exception as e:
        print(f"  [Error] 메모장 실행 실패: {e}")
        return None

    content = ""
    hwnd = 0
    
    # 2. PID에 해당하는 창 핸들 찾기 (재시도 로직)
    for attempt in range(10): # 0.5초 간격으로 10번 시도 (총 5초)
        def callback(h, extra):
            if win32gui.IsWindowVisible(h) or True: # Hide 상태여도 찾아야 함
                _, pid = win32process.GetWindowThreadProcessId(h)
                if pid == target_pid:
                    # 클래스명이 Notepad인 창만 필터링
                    if win32gui.GetClassName(h) == "Notepad":
                        extra.append(h)
        
        hwnds = []
        win32gui.EnumWindows(callback, hwnds)
        if hwnds:
            hwnd = hwnds[0]
            break
        time.sleep(0.5)
    
    if not hwnd:
        print(f"  [Error] 메모장 핸들을 찾을 수 없음 (PID: {target_pid}): {filename}")
        proc.terminate()
        return None

    # 3. 텍스트 추출
    try:
        edit_hwnd = win32gui.FindWindowEx(hwnd, None, "RichEditD2Dpt", None)
        if not edit_hwnd:
            edit_hwnd = win32gui.FindWindowEx(hwnd, None, "Edit", None)
            
        if edit_hwnd:
            length = win32gui.SendMessage(edit_hwnd, win32con.WM_GETTEXTLENGTH, 0, 0)
            if length > 0:
                buffer = ctypes.create_unicode_buffer(length + 1)
                win32gui.SendMessage(edit_hwnd, win32con.WM_GETTEXT, length + 1, buffer)
                content = buffer.value
            else:
                print(f"  [Warning] 메모장 텍스트 길이가 0임: {filename}")
    except Exception as e:
        print(f"  [Error] 텍스트 추출 중 예외 발생: {e}")
    finally:
        # 4. 종료
        win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
        time.sleep(0.1)
        proc.terminate()
        
    return content

def parse_table_to_markdown(table_tag):
    rows = []
    for tr in table_tag.find_all('tr'):
        cells = []
        for td in tr.find_all(['td', 'th']):
            c_text = td.get_text('\n', strip=True).replace('\n', '<br>')
            cells.append(c_text.replace('|', r'\|'))
        if not any(cells): continue
        rows.append(f"| {' | '.join(cells)} |")
        if len(rows) == 1:
            rows.append(f"| {' | '.join(['---'] * len(cells))} |")
    return "\n".join(rows)

def parse_mht_html(html_source):
    if not html_source: return None
    try:
        soup = BeautifulSoup(html_source, 'lxml')
    except Exception as e:
        print(f"  [Error] BeautifulSoup 파싱 실패: {e}")
        return None
    
    metadata = {"title": "N/A", "participants": "N/A", "start_date": "N/A"}
    chat_title_dl = soup.find('dl', class_='chat_title')
    if chat_title_dl:
        dt_tag = chat_title_dl.find('dt')
        if dt_tag:
            metadata["title"] = re.sub(r'^제목\s*:\s*', '', dt_tag.get_text(strip=True)).strip()
        dd_tag = chat_title_dl.find('dd')
        if dd_tag:
            metadata["participants"] = re.sub(r'^참석자(\(\d+\))?\s*:\s*', '', dd_tag.get_text(strip=True)).strip()

    time_wrap = soup.find('div', class_='im_time_wrap')
    if time_wrap:
        corner_c = time_wrap.find('span', class_='corner_C')
        if corner_c:
            raw_time = corner_c.get_text(strip=True)
            date_match = re.search(r'^(\d{4}년 \d{1,2}월 \d{1,2}일 \w+요일)', raw_time)
            if date_match: metadata["start_date"] = date_match.group(1)

    messages = []
    chat_items = soup.find_all('li', class_=re.compile(r'user(You|Me)'))
    for item in chat_items:
        author_div = item.find('div', class_='author')
        sender, msg_time = "N/A", "N/A"
        if author_div:
            name_span = author_div.find('span', class_='name')
            if name_span: sender = name_span.get_text(strip=True).rstrip('/')
            date_span = author_div.find('span', class_='date')
            if date_span:
                raw_time = date_span.get_text(strip=True)
                time_digits = re.sub(r'[^0-9]', '', raw_time)
                msg_time = f"{time_digits[:2]}:{time_digits[2:]}" if len(time_digits) == 4 else raw_time.strip('[] :')

        message_div = item.find('div', class_='message')
        content = ""
        if message_div:
            table = message_div.find('table')
            content = parse_table_to_markdown(table) if table else message_div.get_text('\n', strip=True)

        if sender != "N/A" or content:
            messages.append({"date": metadata["start_date"], "sender": sender, "time": msg_time, "content": content})

    return {"metadata": metadata, "messages": messages}
