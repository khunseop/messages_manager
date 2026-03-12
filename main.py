import win32com.client as win32
import os
import html
import re
import json

def parse_word_to_json(mht_path):
    word = win32.gencache.EnsureDispatch('Word.Application')
    word.Visible = False
    word.DisplayAlerts = 0
    
    abs_path = os.path.abspath(mht_path)
    doc = word.Documents.Open(abs_path, ReadOnly=True, Visible=False)
    
    elements = []
    metadata = {"title": "N/A", "period": "N/A", "participants": "N/A"}

    # 1. 모든 표 가져오기
    table_ranges = []
    for table in doc.Tables:
        start = table.Range.Start
        end = table.Range.End
        table_ranges.append((start, end))
        
        table_md = []
        for r in range(1, table.Rows.Count + 1):
            row_data = []
            for c in range(1, table.Columns.Count + 1):
                try:
                    cell_text = table.Cell(r, c).Range.Text
                    clean_text = cell_text.replace('\x07', '').replace('\r', '<br>').replace('\x0b', '<br>')
                    clean_text = clean_text.replace('|', r'\|').strip()
                    clean_text = re.sub(r'(<br>)+$', '', clean_text)
                    row_data.append(clean_text)
                except:
                    row_data.append("") 
            table_md.append(f"| {' | '.join(row_data)} |")
            if r == 1:
                table_md.append(f"| {' | '.join(['---'] * len(row_data))} |")
        
        formatted_table = "\n".join(table_md)
        elements.append({"start": start, "type": "content", "content": formatted_table})

    # 2. 정규표현식 정의
    # 날짜 패턴: 2026년 3월 11일 수요일...
    date_pattern = re.compile(r'^\d{4}년 \d{1,2}월 \d{1,2}일 \w+요일')
    # 발신자 패턴: 이름/직책/그룹/회사 [09:16]:
    # 대괄호 안의 시간 형식을 기준으로 분리
    sender_pattern = re.compile(r'^(.*)\s*\[(\d{2}:\d{2})\]:$')

    # 3. 모든 문단 분류
    for para in doc.Paragraphs:
        p_start = para.Range.Start
        text = para.Range.Text.strip()
        if not text: continue
        
        # 메타데이터 추출
        if text.startswith("제목 :") and metadata["title"] == "N/A":
            metadata["title"] = text.replace("제목 :", "").strip()
            continue
        elif text.startswith("기간 :") and metadata["period"] == "N/A":
            metadata["period"] = text.replace("기간 :", "").strip()
            continue
        elif "참석자" in text and ":" in text and metadata["participants"] == "N/A":
            metadata["participants"] = text.split(":", 1)[1].strip()
            continue

        is_inside_table = any(start <= p_start < end for start, end in table_ranges)
        if not is_inside_table:
            if date_pattern.match(text):
                elements.append({"start": p_start, "type": "date", "content": text})
            elif sender_pattern.match(text):
                match = sender_pattern.match(text)
                elements.append({
                    "start": p_start, 
                    "type": "sender_info", 
                    "sender": match.group(1).strip(),
                    "time": match.group(2).strip()
                })
            else:
                elements.append({"start": p_start, "type": "content", "content": text.replace('\x0b', '\n')})

    # 4. 정렬 및 메시지 구조화
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
            # 이전 sender_info가 있는 경우에만 메시지로 추가 (메타데이터 문구 등 제외 방지)
            if current_sender != "N/A":
                structured_messages.append({
                    "date": current_date,
                    "sender": current_sender,
                    "time": current_time,
                    "content": html.unescape(e["content"])
                })

    doc.Close(False)
    word.Quit()
    
    return {
        "metadata": metadata,
        "messages": structured_messages
    }

# 실행
if __name__ == "__main__":
    input_file = "your_file.mht"
    if os.path.exists(input_file):
        print(f"JSON 구조화 파싱 시작: {input_file}")
        data = parse_word_to_json(input_file)
        
        output_filename = "messenger_backup.json"
        with open(output_filename, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
            
        print("-" * 30)
        print(f"파싱 완료! 메시지 수: {len(data['messages'])}")
        print(f"결과 저장: {output_filename}")
    else:
        print(f"파일을 찾을 수 없습니다: {input_file}")
