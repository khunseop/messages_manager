import os
import re
import time
import win32gui
import win32con
import ctypes
import subprocess
import html
from bs4 import BeautifulSoup

def get_text_from_notepad_hidden(file_path):
    """
    메모장을 완전히 숨겨진 상태로 실행하여 텍스트를 추출합니다.
    """
    abs_path = os.path.abspath(file_path)
    filename = os.path.basename(file_path)
    
    # 1. 메모장을 완전히 숨김 상태로 실행 (SW_HIDE = 0)
    info = subprocess.STARTUPINFO()
    info.dwFlags |= subprocess.STARTF_USESHOWWINDOW
    info.wShowWindow = win32con.SW_HIDE
    proc = subprocess.Popen(['notepad.exe', abs_path], startupinfo=info)
    
    hwnd = 0
    start_wait = time.time()
    
    # 2. 핸들 찾기 최적화 (창 제목이 아닌 프로세스 ID로 찾는 것이 정확함)
    # 하지만 여기서는 기존의 안정적인 타이틀 검색 방식을 유지하되 빠르게 스캔
    while time.time() - start_wait < 3:
        def callback(h, extra):
            title = win32gui.GetWindowText(h)
            if filename in title and "메모장" in title:
                extra.append(h)
        hwnds = []
        win32gui.EnumWindows(callback, hwnds)
        if hwnds:
            hwnd = hwnds[0]
            break
        time.sleep(0.1)
    
    if not hwnd:
        proc.terminate()
        return None

    content = ""
    try:
        # 3. 텍스트 영역 핸들 찾기 (윈도우 10/11 호환)
        edit_hwnd = win32gui.FindWindowEx(hwnd, None, "RichEditD2Dpt", None)
        if not edit_hwnd:
            edit_hwnd = win32gui.FindWindowEx(hwnd, None, "Edit", None)
            
        if edit_hwnd:
            length = win32gui.SendMessage(edit_hwnd, win32con.WM_GETTEXTLENGTH, 0, 0)
            buffer = ctypes.create_unicode_buffer(length + 1)
            win32gui.SendMessage(edit_hwnd, win32con.WM_GETTEXT, length + 1, buffer)
            content = buffer.value
    finally:
        # 4. 즉시 종료
        win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
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
    soup = BeautifulSoup(html_source, 'lxml')
    
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
