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
    메모장을 화면 밖으로 보내어 작업표시줄과 화면에 나타나지 않게 처리하며 텍스트를 추출합니다.
    """
    abs_path = os.path.abspath(file_path)
    filename = os.path.basename(file_path)
    
    # 1. 메모장을 숨김 상태로 시작 시도
    info = subprocess.STARTUPINFO()
    info.dwFlags |= subprocess.STARTF_USESHOWWINDOW
    info.wShowWindow = win32con.SW_HIDE # 일단 숨김으로 시작
    
    try:
        proc = subprocess.Popen(['notepad.exe', abs_path], startupinfo=info)
        target_pid = proc.pid
    except Exception as e:
        print(f"  [Error] 메모장 실행 실패: {e}")
        return None

    content = ""
    hwnd = 0
    
    # 2. 핸들 획득 (SW_HIDE 상태여도 PID로 찾기 가능)
    start_time = time.time()
    while time.time() - start_time < 8:
        def callback(h, extra):
            _, pid = win32process.GetWindowThreadProcessId(h)
            if pid == target_pid and win32gui.GetClassName(h) == "Notepad":
                extra.append(h)
        hwnds = []
        win32gui.EnumWindows(callback, hwnds)
        
        if hwnds:
            hwnd = hwnds[0]
            # 만약 HIDE 상태라 텍스트 로딩이 안 된다면 잠시 화면 밖으로 이동시켜 활성화
            # win32gui.ShowWindow(hwnd, win32con.SW_SHOWMINIMIZED) # 필요 시 활성화
            
            edit_hwnd = win32gui.FindWindowEx(hwnd, None, "RichEditD2Dpt", None)
            if not edit_hwnd:
                edit_hwnd = win32gui.FindWindowEx(hwnd, None, "Edit", None)
            
            if edit_hwnd:
                length = win32gui.SendMessage(edit_hwnd, win32con.WM_GETTEXTLENGTH, 0, 0)
                if length > 0:
                    buffer = ctypes.create_unicode_buffer(length + 1)
                    win32gui.SendMessage(edit_hwnd, win32con.WM_GETTEXT, length + 1, buffer)
                    content = buffer.value
                    if content.strip(): break
        time.sleep(0.5)
    
    # 3. 종료
    try:
        if hwnd: win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
        time.sleep(0.1)
        proc.terminate()
    except: pass
        
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

    # 대화방 상단 날짜 정보 추출 (가장 중요)
    time_wrap = soup.find('div', class_='im_time_wrap')
    if time_wrap:
        corner_c = time_wrap.find('span', class_='corner_C')
        if corner_c:
            raw_time = corner_c.get_text(strip=True)
            # 날짜 패턴 매칭 (예: 2026년 3월 13일 금요일)
            date_match = re.search(r'(\d{4}년 \d{1,2}월 \d{1,2}일 \w+요일)', raw_time)
            if date_match:
                metadata["start_date"] = date_match.group(1)

    messages = []
    # li 태그 순회하며 메시지 추출
    chat_items = soup.find_all('li', class_=re.compile(r'user(You|Me)'))
    
    # 만약 개별 메시지 사이사이에 날짜 구분선(im_time_wrap)이 더 있다면 업데이트하는 로직 추가 가능
    current_msg_date = metadata["start_date"]

    for item in chat_items:
        # 메시지 바로 직전에 날짜 구분선이 있는지 확인 (형제 태그 검색)
        prev_sibling = item.find_previous_sibling()
        if prev_sibling and 'im_time_wrap' in prev_sibling.get('class', []):
            new_date_span = prev_sibling.find('span', class_='corner_C')
            if new_date_span:
                date_match = re.search(r'(\d{4}년 \d{1,2}월 \d{1,2}일 \w+요일)', new_date_span.get_text(strip=True))
                if date_match: current_msg_date = date_match.group(1)

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
            messages.append({
                "date": current_msg_date,
                "sender": sender,
                "time": msg_time,
                "content": content
            })

    return {"metadata": metadata, "messages": messages}
