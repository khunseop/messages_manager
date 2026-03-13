import os
import re
import time
import win32gui
import win32con
import ctypes
import subprocess
import html
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

def parse_table_to_markdown(table_tag):
    """
    HTML table 태그를 Markdown 표 형식으로 변환합니다.
    """
    rows = []
    for tr in table_tag.find_all('tr'):
        cells = []
        for td in tr.find_all(['td', 'th']):
            # 셀 내부 개행은 <br>로 치환
            c_text = td.get_text('\n', strip=True).replace('\n', '<br>')
            cells.append(c_text.replace('|', r'\|'))
        
        if not any(cells): continue
        rows.append(f"| {' | '.join(cells)} |")
        if len(rows) == 1:
            rows.append(f"| {' | '.join(['---'] * len(cells))} |")
            
    return "\n".join(rows)

def parse_mht_html(html_source):
    """
    추출된 HTML 소스에서 대화방 정보 및 모든 메시지를 파싱합니다.
    """
    if not html_source:
        return None
        
    soup = BeautifulSoup(html_source, 'lxml')
    
    # 1. 메타데이터 파싱
    metadata = {"title": "N/A", "participants": "N/A", "start_date": "N/A"}
    chat_title_dl = soup.find('dl', class_='chat_title')
    if chat_title_dl:
        dt_tag = chat_title_dl.find('dt')
        if dt_tag:
            metadata["title"] = re.sub(r'^제목\s*:\s*', '', dt_tag.get_text(strip=True)).strip()
        dd_tag = chat_title_dl.find('dd')
        if dd_tag:
            metadata["participants"] = re.sub(r'^참석자(\(\d+\))?\s*:\s*', '', dd_tag.get_text(strip=True)).strip()

    # 시작 시간에서 날짜 부분만 추출 (예: 2026년 3월 13일 금요일)
    time_wrap = soup.find('div', class_='im_time_wrap')
    if time_wrap:
        corner_c = time_wrap.find('span', class_='corner_C')
        if corner_c:
            raw_time = corner_c.get_text(strip=True)
            date_match = re.search(r'^(\d{4}년 \d{1,2}월 \d{1,2}일 \w+요일)', raw_time)
            if date_match:
                metadata["start_date"] = date_match.group(1)

    # 2. 메시지 파싱
    messages = []
    # li 태그 중 대화 내용을 포함하는 것들을 찾음 (userYou, userMe 등)
    chat_items = soup.find_all('li', class_=re.compile(r'user(You|Me)'))
    
    for item in chat_items:
        # 발신자 정보 추출
        author_div = item.find('div', class_='author')
        sender = "N/A"
        msg_time = "N/A"
        
        if author_div:
            name_span = author_div.find('span', class_='name')
            if name_span:
                sender = name_span.get_text(strip=True).rstrip('/')
            
            date_span = author_div.find('span', class_='date')
            if date_span:
                # [10:57]: 에서 숫자만 추출하여 HH:MM 형식 보장
                raw_time = date_span.get_text(strip=True)
                time_digits = re.sub(r'[^0-9]', '', raw_time)
                if len(time_digits) == 4:
                    msg_time = f"{time_digits[:2]}:{time_digits[2:]}"
                else:
                    msg_time = raw_time.strip('[] :')

        # 메시지 내용 추출
        message_div = item.find('div', class_='message')
        content = ""
        if message_div:
            # 내부에 표가 있는지 확인
            table = message_div.find('table')
            if table:
                content = parse_table_to_markdown(table)
            else:
                # 일반 텍스트 (HTML 엔티티 변환 포함)
                content = message_div.get_text('\n', strip=True)

        if sender != "N/A" or content:
            messages.append({
                "date": metadata["start_date"],
                "sender": sender,
                "time": msg_time,
                "content": content
            })

    return {
        "metadata": metadata,
        "messages": messages
    }

if __name__ == "__main__":
    import glob
    files = glob.glob("inputs/*.mht")
    if files:
        raw_html = get_text_from_notepad_memory(files[0])
        if raw_html:
            data = parse_mht_html(raw_html)
            print(f"\n[파싱 결과] {data['metadata']['title']}")
            print(f"메시지 수: {len(data['messages'])}개")
            if data['messages']:
                print(f"마지막 메시지 예시: {data['messages'][-1]['sender']}: {data['messages'][-1]['content'][:30]}...")
