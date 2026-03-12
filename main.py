import win32com.client as win32
import os
import html
import re
import json

def parse_word_to_json(mht_path):
    word = win32.gencache.EnsureDispatch('Word.Application')
    word.Visible = False
    word.DisplayAlerts = 0
    word.ScreenUpdating = False
    
    abs_path = os.path.abspath(mht_path)
    doc = word.Documents.Open(abs_path, ReadOnly=True, Visible=False)
    
    # 전체 텍스트 로드
    full_text = doc.Content.Text
    
    elements = []
    table_ranges = []

    # 1. 테이블 파싱 및 범위 기록
    for table in doc.Tables:
        start = table.Range.Start
        end = table.Range.End
        table_ranges.append((start, end))
        
        raw_table_text = table.Range.Text
        rows_raw = raw_table_text.strip('\r\x07').split('\r\x07')
        
        table_md = []
        for i, row_raw in enumerate(rows_raw):
            cells_raw = row_raw.split('\x07')
            if cells_raw and not cells_raw[-1]:
                cells_raw.pop()
            
            clean_cells = []
            for cell in cells_raw:
                # 표 내부 텍스트 정제
                c = cell.replace('\x07', '').replace('|', r'\|').replace('\x0b', '<br>').replace('\r', '<br>').strip()
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

    # 2. 메타데이터 추출 (줄바꿈 \r 전까지만 매칭하도록 수정)
    metadata = {"title": "N/A", "period": "N/A", "participants": "N/A"}
    
    title_match = re.search(r'제목\s*:\s*([^\r\n]*)', full_text)
    if title_match: metadata["title"] = title_match.group(1).strip()
    
    period_match = re.search(r'기간\s*:\s*([^\r\n]*)', full_text)
    if period_match: metadata["period"] = period_match.group(1).strip()
    
    participants_match = re.search(r'참석자.*?\s*:\s*([^\r\n]*)', full_text)
    if participants_match: metadata["participants"] = participants_match.group(1).strip()

    # 3. 문단 파싱 (날짜, 발신자, 본문)
    date_pattern = re.compile(r'^(\d{4}년 \d{1,2}월 \d{1,2}일 \w+요일)')
    sender_pattern = re.compile(r'^([^\r\n]+)\s*\[(\d{2}:\d{2})\]:')

    current_pos = 0
    # Word의 문단 구분자인 \r 로 분리
    for p_text in full_text.split('\r'):
        p_len = len(p_text) + 1 
        p_strip = p_text.replace('\x07', '').strip() # 제어 문자 제거
        
        if p_strip:
            # 표 범위 내 텍스트인지 체크 (미세 오차 방지를 위해 중앙값으로 체크)
            mid_pos = current_pos + (len(p_text) // 2)
            is_inside_table = any(s <= mid_pos < e for s, e in table_ranges)
            
            if not is_inside_table:
                # 메타데이터 라인 제외
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
                # 최종 내용에서 한 번 더 제어 문자 및 불필요한 공백 정리
                clean_content = e["content"].replace('\x07', '').strip()
                if clean_content:
                    structured_messages.append({
                        "date": current_date,
                        "sender": current_sender,
                        "time": current_time,
                        "content": html.unescape(clean_content)
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
        print(f"고속 정밀 파싱 시작: {input_file}")
        
        data = parse_word_to_json(input_file)
        
        with open("messenger_backup.json", "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
            
        elapsed = time.time() - start_time
        print("-" * 30)
        print(f"파싱 완료! (소요 시간: {elapsed:.2f}초)")
        print(f"제목: {data['metadata']['title']}")
        print(f"메시지 수: {len(data['messages'])}")
    else:
        print(f"파일을 찾을 수 없습니다: {input_file}")
