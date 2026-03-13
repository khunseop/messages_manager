import os
import re
import json
import time
import pyperclip
from pywinauto.application import Application
from bs4 import BeautifulSoup

def get_html_via_notepad(file_path):
    """
    Notepad를 자동화하여 DRM이 해제된 원시 HTML 텍스트를 클립보드를 통해 가져옵니다.
    """
    abs_path = os.path.abspath(file_path)
    
    # 1. 클립보드 초기화
    pyperclip.copy('')
    
    # 2. 메모장 실행 및 파일 열기
    # notepad.exe 파일경로 형태로 실행
    app = Application(backend="uia").start(f'notepad.exe "{abs_path}"')
    
    # 3. 창 대기 및 제어
    # 한국어 윈도우 환경을 고려하여 정규식으로 창 제목 매칭
    # 파일 이름에 상관없이 활성화된 메모장 창을 찾음
    try:
        dlg = app.window(class_name="Notepad")
        dlg.wait('ready', timeout=5)
        
        # 4. 전체 선택 (Ctrl+A) 및 복사 (Ctrl+C)
        dlg.type_keys('^a^c')
        
        # 클립보드에 데이터가 들어갈 때까지 약간 대기
        time.sleep(0.5)
        
        # 5. 메모장 닫기
        dlg.close()
    except Exception as e:
        print(f"Notepad 제어 중 오류 발생: {e}")
        try:
            app.kill() # 강제 종료
        except:
            pass
        return None
        
    # 6. 클립보드 내용 반환
    html_content = pyperclip.paste()
    return html_content

def parse_word_to_json(mht_path):
    """
    MHT 파일을 Notepad로 열어 HTML을 추출하고, BeautifulSoup으로 정밀 파싱합니다.
    (함수명은 호환성을 위해 유지)
    """
    print("  - 메모장 자동화를 통해 데이터 추출 중...")
    html_source = get_html_via_notepad(mht_path)
    
    if not html_source or len(html_source.strip()) == 0:
        raise ValueError("Notepad를 통해 HTML 소스를 추출하지 못했습니다. (클립보드 비어있음)")

    # 추출된 HTML 소스를 BeautifulSoup으로 파싱
    soup = BeautifulSoup(html_source, 'lxml')
    metadata = {"title": "N/A", "period": "N/A", "participants": "N/A"}
    messages = []
    
    # MHT 내의 <body> 태그 내부를 탐색
    body = soup.find('body')
    if not body:
        return {"metadata": metadata, "messages": []}

    # 요일까지만 추출하는 날짜 패턴
    date_pattern = re.compile(r'^(\d{4}년 \d{1,2}월 \d{1,2}일 \w+요일)')
    # 발신자 패턴: [00:00]: 기준
    sender_pattern = re.compile(r'^([^\n]+)\s*\[(\d{2}:\d{2})\]:')

    current_date, current_sender, current_time = "N/A", "N/A", "N/A"

    # 모든 요소(문단 p, div, 표 table 등)를 문서 순서대로 순회
    for element in body.find_all(['p', 'div', 'table'], recursive=False):
        tag_name = element.name
        
        if tag_name in ['p', 'div']:
            text = element.get_text(strip=True)
            if not text: continue
            
            # 메타데이터 추출
            if text.startswith("제목 :") and metadata["title"] == "N/A":
                metadata["title"] = text.replace("제목 :", "").strip()
            elif text.startswith("기간 :") and metadata["period"] == "N/A":
                metadata["period"] = text.replace("기간 :", "").strip()
            elif "참석자" in text and ":" in text and metadata["participants"] == "N/A":
                metadata["participants"] = text.split(":", 1)[1].strip()
            
            # 날짜 및 발신자 정보 업데이트
            date_m = date_pattern.match(text)
            sender_m = sender_pattern.match(text)
            
            if date_m:
                current_date = date_m.group(1)
            elif sender_m:
                current_sender = sender_m.group(1).strip()
                current_time = sender_m.group(2).strip()
            else:
                # 일반 대화 내용 (발신자가 있는 경우만)
                if current_sender != "N/A":
                    # 셀 내부 개행 등을 고려하여 텍스트 추출 (HTML 개행을 유지)
                    content = element.get_text('\n', strip=True)
                    messages.append({
                        "date": current_date,
                        "sender": current_sender,
                        "time": current_time,
                        "content": content
                    })
        
        elif tag_name == 'table':
            # 표 파싱 (HTML 태그 기반이므로 열이 합쳐지거나 구조가 깨지지 않음)
            table_rows = []
            for tr in element.find_all('tr'):
                cells = []
                for td in tr.find_all(['td', 'th']):
                    # 셀 내부의 개행은 마크다운 표 깨짐 방지를 위해 <br>로 치환
                    c_text = td.get_text('\n', strip=True).replace('\n', '<br>')
                    # 파이프 기호 이스케이프
                    cells.append(c_text.replace('|', r'\|'))
                
                if not any(cells): continue 
                
                table_rows.append(f"| {' | '.join(cells)} |")
                if len(table_rows) == 1: # 헤더 구분선 추가
                    table_rows.append(f"| {' | '.join(['---'] * len(cells))} |")
            
            if table_rows and current_sender != "N/A":
                messages.append({
                    "date": current_date,
                    "sender": current_sender,
                    "time": current_time,
                    "content": "\n".join(table_rows)
                })

    return {"metadata": metadata, "messages": messages}

if __name__ == "__main__":
    # 단독 실행 시 테스트용 코드
    import glob
    test_files = glob.glob("inputs/*.mht")
    if test_files:
        try:
            start_t = time.time()
            res = parse_word_to_json(test_files[0])
            print(f"파싱 성공: {len(res['messages'])}개 메시지 (소요시간: {time.time()-start_t:.2f}초)")
            print(f"제목: {res['metadata']['title']}")
        except Exception as e:
            print(f"파싱 실패: {e}")
