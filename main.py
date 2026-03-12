import win32com.client as win32
import os
import html
import re
import json

def parse_word_to_json(mht_path):
    word = win32.gencache.EnsureDispatch('Word.Application')
    word.Visible = False
    word.DisplayAlerts = 0
    word.ScreenUpdating = False # 성능 향상을 위해 화면 업데이트 중지
    
    abs_path = os.path.abspath(mht_path)
    doc = word.Documents.Open(abs_path, ReadOnly=True, Visible=False)
    
    # [최적화 1] 전체 텍스트를 한 번의 COM 호출로 가져옴
    full_text = doc.Content.Text
    
    elements = []
    table_ranges = []

    # [최적화 2] 테이블 벌크 파싱 (열 보존 로직 강화)
    for table in doc.Tables:
        start = table.Range.Start
        end = table.Range.End
        table_ranges.append((start, end))
        
        # Cell(r,c) 순회 대신 전체 텍스트 파싱
        raw_table_text = table.Range.Text
        # Word 테이블 구분자: 행(\r\x07), 셀(\x07)
        rows_raw = raw_table_text.strip('\r\x07').split('\r\x07')
        
        table_md = []
        for i, row_raw in enumerate(rows_raw):
            cells_raw = row_raw.split('\x07')
            if cells_raw and not cells_raw[-1]:
                cells_raw.pop() # 행 끝의 불필요한 빈 요소 제거
            
            clean_cells = []
            for cell in cells_raw:
                c = cell.replace('|', r'\|').replace('\x0b', '<br>').replace('\r', '<br>').strip()
                c = re.sub(r'(<br>)+$', '', c)
                clean_cells.append(c)
            
            if not clean_cells: continue
            table_md.append(f"| {' | '.join(clean_cells)} |")
            if i == 0:
                table_md.append(f"| {' | '.join(['---'] * len(clean_cells))} |")
        
        elements.append({
            "start": start, 
            "type": "content", 
            "content": "\n".join(table_md)
        })

    # [최적화 3] 정규표현식을 이용한 고속 메타데이터 및 패턴 추출
    metadata = {"title": "N/A", "period": "N/A", "participants": "N/A"}
    
    # 메타데이터 추출
    title_match = re.search(r'제목\s*:\s*(.*)', full_text)
    if title_match: metadata["title"] = title_match.group(1).strip()
    
    period_match = re.search(r'기간\s*:\s*(.*)', full_text)
    if period_match: metadata["period"] = period_match.group(1).strip()
    
    participants_match = re.search(r'참석자.*?\s*:\s*(.*)', full_text)
    if participants_match: metadata["participants"] = participants_match.group(1).strip()

    # 날짜 및 발신자 패턴 정의
    date_pattern = re.compile(r'^(\d{4}년 \d{1,2}월 \d{1,2}일 \w+요일)', re.MULTILINE)
    sender_pattern = re.compile(r'^([^\r\n]+)\s*\[(\d{2}:\d{2})\]:', re.MULTILINE)

    # 문단 단위 분리 및 분류 (COM 호출 없이 full_text 기반)
    current_pos = 0
    # \r 은 Word에서 문단 구분자임
    for p_text in full_text.split('\r'):
        p_len = len(p_text) + 1 # \r 길이 포함
        p_strip = p_text.strip()
        
        if p_strip:
            # 해당 위치가 테이블 내부인지 확인
            is_inside_table = any(s <= current_pos < e for s, e in table_ranges)
            
            if not is_inside_table:
                # 메타데이터 줄은 무시 (이미 추출함)
                is_meta = any(p_strip.startswith(x) for x in ["제목 :", "기간 :"]) or "참석자" in p_strip[:10]
                
                if not is_meta:
                    date_m = date_pattern.match(p_strip)
                    sender_m = sender_pattern.match(p_strip)
                    
                    if date_m:
                        elements.append({"start": current_pos, "type": "date", "content": date_m.group(1)})
                    elif sender_m:
                        elements.append({
                            "start": current_pos, 
                            "type": "sender_info", 
                            "sender": sender_m.group(1).strip(),
                            "time": sender_m.group(2).strip()
                        })
                    else:
                        elements.append({"start": current_pos, "type": "content", "content": p_strip.replace('\x0b', '\n')})
        
        current_pos += p_len

    # 4. 정렬 및 구조화
    elements.sort(key=lambda x: x["start"])
    
    structured_messages = []
    current_date = "N/A"
    current_sender = "N/A"
    current_time = "N/A"

    for e in elements:
        if e["type"] == "date":
            current_date = e["content"]
        elif e["type"] == "sender_info":
            current_sender = e["sender"]
            current_time = e["time"]
        elif e["type"] == "content":
            if current_sender != "N/A":
                structured_messages.append({
                    "date": current_date,
                    "sender": current_sender,
                    "time": current_time,
                    "content": html.unescape(e["content"])
                })

    doc.Close(False)
    word.ScreenUpdating = True
    word.Quit()
    
    return {
        "metadata": metadata,
        "messages": structured_messages
    }

if __name__ == "__main__":
    input_file = "your_file.mht"
    if os.path.exists(input_file):
        import time
        start_time = time.time()
        print(f"고속 파싱 시작: {input_file}")
        
        data = parse_word_to_json(input_file)
        
        with open("messenger_backup.json", "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
            
        elapsed = time.time() - start_time
        print("-" * 30)
        print(f"파싱 완료! (소요 시간: {elapsed:.2f}초)")
        print(f"메시지 수: {len(data['messages'])}")
    else:
        print(f"파일을 찾을 수 없습니다: {input_file}")
