import win32com.client as win32
import os
import html
import re
import json
import tempfile
from bs4 import BeautifulSoup

def parse_mht_via_html(mht_path, word_app=None):
    """
    Word로 MHT를 열어 '필터링된 HTML'로 저장한 뒤, BeautifulSoup으로 정밀 파싱
    """
    is_internal_word = False
    if word_app is None:
        # Word 실행 (DRM 복호화 역할)
        word = win32.Dispatch('Word.Application')
        is_internal_word = True
    else:
        word = word_app

    word.Visible = False
    word.DisplayAlerts = 0
    
    # 1. 임시 HTML 파일 경로 생성
    temp_fd, temp_html_path = tempfile.mkstemp(suffix='.html')
    os.close(temp_fd)
    
    doc = None
    try:
        abs_mht_path = os.path.abspath(mht_path)
        doc = word.Documents.Open(abs_mht_path, ReadOnly=True, Visible=False)
        
        # wdFormatFilteredHTML = 10 (필터링된 HTML로 저장 - 가장 깔끔함)
        doc.SaveAs2(temp_html_path, FileFormat=10)
        doc.Close(False)
        doc = None
        
        # 2. 저장된 HTML 파일을 BeautifulSoup으로 읽기
        with open(temp_html_path, 'r', encoding='utf-8', errors='ignore') as f:
            soup = BeautifulSoup(f, 'lxml')
            
        # 3. 데이터 추출 로직
        metadata = {"title": "N/A", "period": "N/A", "participants": "N/A"}
        messages = []
        
        # Word HTML은 보통 <body> 내부에 <p>와 <table> 태그가 평면적으로 나열됨
        body = soup.find('body')
        if not body: return {"metadata": metadata, "messages": []}

        # 요일까지만 추출하는 날짜 패턴
        date_pattern = re.compile(r'^(\d{4}년 \d{1,2}월 \d{1,2}일 \w+요일)')
        # 발신자 패턴: [00:00]: 기준
        sender_pattern = re.compile(r'^([^\n]+)\s*\[(\d{2}:\d{2})\]:')

        current_date, current_sender, current_time = "N/A", "N/A", "N/A"

        # 모든 p와 table 태그를 문서 순서대로 순회
        for element in body.find_all(['p', 'table'], recursive=False):
            if element.name == 'p':
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
                        # <br> 태그 보존을 위해 get_text 시 개행 추가
                        content = element.get_text('\n', strip=True).replace('\x0b', '\n')
                        messages.append({
                            "date": current_date,
                            "sender": current_sender,
                            "time": current_time,
                            "content": content
                        })
            
            elif element.name == 'table':
                # 표 파싱 (BeautifulSoup으로 컬럼 완벽 분리)
                table_rows = []
                for tr in element.find_all('tr'):
                    # 셀 내부의 텍스트를 가져오되 마데카운 파이프 기호 이스케이프
                    cells = [td.get_text('\n', strip=True).replace('\x0b', '\n').replace('|', r'\|') for td in tr.find_all(['td', 'th'])]
                    # 셀 내부 개행은 마크다운 표 깨짐 방지를 위해 <br>로 변환
                    cells = [c.replace('\n', '<br>') for c in cells]
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

    except Exception as e:
        print(f"  - HTML 파싱 중 에러 발생: {e}")
        return {"metadata": metadata, "messages": []}
    finally:
        if doc:
            doc.Close(False)
        if is_internal_word:
            word.Quit()
        # 임시 HTML 파일 삭제
        if os.path.exists(temp_html_path):
            try: os.remove(temp_html_path)
            except: pass

if __name__ == "__main__":
    test_file = "inputs/your_file.mht"
    if os.path.exists(test_file):
        data = parse_mht_via_html(test_file)
        print(f"추출 완료: {len(data['messages'])}개 메시지")
        print(f"제목: {data['metadata']['title']}")
        print(f"기간: {data['metadata']['period']}")
