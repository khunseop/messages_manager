import win32com.client as win32
import os
import html
import re

def parse_word_keeping_order(mht_path):
    # Word 어플리케이션 실행
    word = win32.gencache.EnsureDispatch('Word.Application')
    word.Visible = False
    word.DisplayAlerts = 0
    
    abs_path = os.path.abspath(mht_path)
    doc = word.Documents.Open(abs_path, ReadOnly=True, Visible=False)
    
    elements = []
    metadata = {
        "title": "N/A",
        "period": "N/A",
        "participants": "N/A"
    }

    # 1. 모든 표 가져오기 (이미지나 특정 서식 대응)
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
        
        formatted_table = "\n" + "\n".join(table_md) + "\n"
        elements.append((start, "table", formatted_table))

    # 2. 정규표현식 정의
    # 날짜 패턴: 2026년 3월 11일 수요일...
    date_pattern = re.compile(r'^\d{4}년 \d{1,2}월 \d{1,2}일 \w+요일')
    # 발신자 패턴: 이름/그룹/회사 오전/오후 12:34:
    # 그룹명에 /가 섞여있을 수 있으므로 뒤에서부터 시간 패턴을 찾음
    sender_pattern = re.compile(r'^(.*)\s+((?:오전|오후)\s+\d{1,2}:\d{2}):$')

    # 3. 모든 문단 가져오기 및 분류
    for para in doc.Paragraphs:
        p_start = para.Range.Start
        text = para.Range.Text.strip()
        if not text: continue
        
        # 메타데이터 추출 (이미 완료된 로직)
        if text.startswith("제목 :") and metadata["title"] == "N/A":
            metadata["title"] = text.replace("제목 :", "").strip()
            continue
        elif text.startswith("기간 :") and metadata["period"] == "N/A":
            metadata["period"] = text.replace("기간 :", "").strip()
            continue
        elif "참석자" in text and ":" in text and metadata["participants"] == "N/A":
            metadata["participants"] = text.split(":", 1)[1].strip()
            continue

        # 표 외부 문단 분류 가공
        is_inside_table = any(start <= p_start < end for start, end in table_ranges)
        if not is_inside_table:
            # 날짜 구분선인 경우
            if date_pattern.match(text):
                elements.append((p_start, "date", f"\n---\n### 📅 {text}\n"))
            # 발신자 정보인 경우
            elif sender_pattern.match(text):
                match = sender_pattern.match(text)
                user_info = match.group(1).strip()
                time_str = match.group(2).strip()
                elements.append((p_start, "sender", f"\n**[{user_info}]** ({time_str})"))
            # 일반 대화 내용
            else:
                clean_p_text = text.replace('\x0b', '\n')
                elements.append((p_start, "text", clean_p_text))

    # 4. 정렬 및 최종 조립
    elements.sort(key=lambda x: x[0])
    
    header_info = f"""# Messenger Backup Report
- **제목**: {metadata['title']}
- **기간**: {metadata['period']}
- **참석자**: {metadata['participants']}

---
"""
    
    # 대화 내용 조립
    chat_body = []
    for _, e_type, e_content in elements:
        if e_type == "text":
            # 일반 텍스트는 바로 앞의 발신자나 날짜 뒤에 자연스럽게 붙도록 함
            chat_body.append(e_content)
        else:
            chat_body.append(e_content)
            
    final_content = header_info + "\n".join(chat_body)
    
    # 후처리
    final_content = html.unescape(final_content)
    final_content = re.sub(r'\n{3,}', '\n\n', final_content)

    doc.Close(False)
    word.Quit()
    
    return final_content, metadata

# 실행
if __name__ == "__main__":
    input_file = "your_file.mht"
    if os.path.exists(input_file):
        print(f"채팅 로그 재구성 시작: {input_file}")
        result_text, meta = parse_word_keeping_order(input_file)
        
        output_filename = "ordered_messenger_backup.md"
        with open(output_filename, "w", encoding="utf-8") as f:
            f.write(result_text)
            
        print("-" * 30)
        print(f"분석 완료: {meta['title']}")
        print(f"결과 저장: {output_filename}")
    else:
        print(f"파일을 찾을 수 없습니다: {input_file}")
