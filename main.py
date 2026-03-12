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

    # 1. 모든 표 가져오기 (Cell 순회 방식)
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
                    # 셀 끝 문자(\x07) 제거, 개행 및 vt를 <br>로 변환
                    clean_text = cell_text.replace('\x07', '').replace('\r', '<br>').replace('\x0b', '<br>')
                    # Markdown 파이프 기호 이스케이프 (raw string 사용으로 경고 해결)
                    clean_text = clean_text.replace('|', r'\|').strip()
                    clean_text = re.sub(r'(<br>)+$', '', clean_text)
                    row_data.append(clean_text)
                except:
                    row_data.append("") 
            
            table_md.append(f"| {' | '.join(row_data)} |")
            if r == 1:
                table_md.append(f"| {' | '.join(['---'] * len(row_data))} |")
        
        formatted_table = "\n" + "\n".join(table_md) + "\n"
        elements.append((table.Range.Start, "table", formatted_table))

    # 2. 모든 문단 가져오기 및 메타데이터 추출
    for para in doc.Paragraphs:
        p_start = para.Range.Start
        text = para.Range.Text.strip()
        
        # 메타데이터 추출 (제목, 기간, 참석자)
        if not text: continue
        
        if text.startswith("제목 :") and metadata["title"] == "N/A":
            metadata["title"] = text.replace("제목 :", "").strip()
        elif text.startswith("기간 :") and metadata["period"] == "N/A":
            metadata["period"] = text.replace("기간 :", "").strip()
        elif "참석자" in text and ":" in text and metadata["participants"] == "N/A":
            metadata["participants"] = text.split(":", 1)[1].strip()

        # 표 외부 문단만 저장
        is_inside_table = any(start <= p_start < end for start, end in table_ranges)
        if not is_inside_table:
            clean_p_text = text.replace('\x0b', '\n')
            elements.append((p_start, "text", clean_p_text))

    # 3. 정렬 및 병합
    elements.sort(key=lambda x: x[0])
    
    # 메타데이터를 상단에 배치
    header_info = f"""# Messenger Backup Report
- **제목**: {metadata['title']}
- **기간**: {metadata['period']}
- **참석자**: {metadata['participants']}

---
"""
    
    final_content = header_info + "\n\n".join([item[2] for item in elements])
    
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
        print(f"파싱 시작: {input_file}")
        result_text, meta = parse_word_keeping_order(input_file)
        
        # 파일명에 제목이나 날짜를 넣어서 저장할 수도 있습니다.
        output_filename = "ordered_messenger_backup.md"
        with open(output_filename, "w", encoding="utf-8") as f:
            f.write(result_text)
            
        print("-" * 30)
        print(f"추출된 정보:")
        print(f"  - 제목: {meta['title']}")
        print(f"  - 기간: {meta['period']}")
        print(f"  - 완료: {output_filename}")
    else:
        print(f"파일을 찾을 수 없습니다: {input_file}")
