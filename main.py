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
    
    full_text = doc.Content.Text
    elements = []
    table_ranges = []

    # 1. 테이블 파싱 (구분자 치환 전략으로 내부 개행 보존)
    for table in doc.Tables:
        start = table.Range.Start
        end = table.Range.End
        table_ranges.append((start, end))
        
        raw_text = table.Range.Text
        
        # [핵심] 행/셀 구분자를 임시 문자열로 치환하여 내부 개행과 분리
        # Word 테이블 구분자 순서: 셀 종료(\x07), 행 종료(\r\x07)
        temp_text = raw_text.replace('\r\x07', '||ROW||').replace('\x07', '||CELL||')
        
        # 이제 셀 내부에 남은 모든 종류의 개행(\r, \x0b, \n)을 <br>로 변환
        temp_text = temp_text.replace('\r', '<br>').replace('\x0b', '<br>').replace('\n', '<br>')
        
        # 다시 행과 셀 단위로 분리
        rows_raw = temp_text.split('||ROW||')
        
        table_md = []
        for i, row_raw in enumerate(rows_raw):
            cells_raw = row_raw.split('||CELL||')
            # 마지막 빈 요소 제거
            if cells_raw and not cells_raw[-1]:
                cells_raw.pop()
            
            if not cells_raw: continue
            
            clean_cells = []
            for cell in cells_raw:
                # 마크다운 파이프 기호 이스케이프 및 정제
                c = cell.replace('|', r'\|').strip(' ')
                # 연속된 <br> 및 마지막 <br> 정리
                c = re.sub(r'(<br>)+$', '', c)
                clean_cells.append(c)
            
            table_md.append(f"| {' | '.join(clean_cells)} |")
            if i == 0:
                table_md.append(f"| {' | '.join(['---'] * len(clean_cells))} |")
        
        elements.append({
            "start": start, 
            "type": "content", 
            "content": "\n".join(table_md)
        })

    # 2. 메타데이터 추출 (초반 3000자 내외)
    metadata = {"title": "N/A", "period": "N/A", "participants": "N/A"}
    top_text = full_text[:3000]
    
    title_match = re.search(r'제목\s*:\s*([^\r\n]*)', top_text)
    if title_match: metadata["title"] = title_match.group(1).strip()
    
    period_match = re.search(r'기간\s*:\s*([^\r\n]*)', top_text)
    if period_match: metadata["period"] = period_match.group(1).strip()
    
    participants_match = re.search(r'참석자.*?\s*:\s*([^\r\n]*)', top_text)
    if participants_match: metadata["participants"] = participants_match.group(1).strip()

    # 3. 표 영역 제외 문단 파싱
    date_pattern = re.compile(r'^(\d{4}년 \d{1,2}월 \d{1,2}일 \w+요일)')
    sender_pattern = re.compile(r'^([^\r\n]+)\s*\[(\d{2}:\d{2})\]:')

    table_ranges.sort()
    last_pos = 0

    def process_segment(segment_text, start_offset):
        current_offset = start_offset
        for p_text in segment_text.split('\r'):
            p_len = len(p_text) + 1
            p_strip = p_text.replace('\x07', '').strip()
            
            if p_strip:
                is_meta = any(p_strip.startswith(x) for x in ["제목 :", "기간 :"]) or "참석자" in p_strip[:10]
                if not is_meta:
                    date_m = date_pattern.match(p_strip)
                    sender_m = sender_pattern.match(p_strip)
                    
                    if date_m:
                        elements.append({"start": current_offset, "type": "date", "content": date_m.group(1)})
                    elif sender_m:
                        elements.append({
                            "start": current_offset, 
                            "type": "sender_info", 
                            "sender": sender_m.group(1).strip(),
                            "time": sender_m.group(2).strip()
                        })
                    else:
                        elements.append({"start": current_offset, "type": "content", "content": p_strip.replace('\x0b', '\n')})
            current_offset += p_len

    for t_start, t_end in table_ranges:
        if t_start > last_pos:
            process_segment(full_text[last_pos:t_start], last_pos)
        last_pos = t_end
    
    if last_pos < len(full_text):
        process_segment(full_text[last_pos:], last_pos)

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
                clean_content = e["content"].replace('\x07', '').strip(' ')
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
        print(f"고속 정밀 파싱 시작 (개행 보존 모드): {input_file}")
        
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
