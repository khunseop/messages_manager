import win32com.client as win32
import os
import html
import re

def parse_word_keeping_order(mht_path):
    word = win32.gencache.EnsureDispatch('Word.Application')
    word.Visible = False
    
    abs_path = os.path.abspath(mht_path)
    doc = word.Documents.Open(abs_path, ReadOnly=True)
    
    elements = []

    # 1. 모든 문단 가져오기 (표 안에 있는 문단은 일단 제외)
    for para in doc.Paragraphs:
        if not para.Range.Information(12):  # 12 = wdWithinTable (표 밖인 경우만)
            text = para.Range.Text.strip()
            if text:
                # (시작 위치, 유형, 내용) 저장
                elements.append((para.Range.Start, "text", text))

    # 2. 모든 표 가져오기
    for table in doc.Tables:
        table_md = []
        for r in range(1, table.Rows.Count + 1):
            row_data = []
            for c in range(1, table.Columns.Count + 1):
                try:
                    cell_text = table.Cell(r, c).Range.Text
                    clean_text = cell_text.replace('\r', '').replace('\x07', '').strip()
                    row_data.append(clean_text)
                except:
                    row_data.append("") 
            
            table_md.append(f"| {' | '.join(row_data)} |")
            if r == 1:
                table_md.append(f"| {' | '.join(['---'] * len(row_data))} |")
        
        formatted_table = "\n" + "\n".join(table_md) + "\n"
        # (시작 위치, 유형, 내용) 저장
        elements.append((table.Range.Start, "table", formatted_table))

    # 3. 문서 내 위치(Start 인덱스) 기준으로 정렬 (핵심!)
    elements.sort(key=lambda x: x[0])

    # 4. 정렬된 순서대로 합치기 및 정제
    final_list = [item[2] for item in elements]
    final_content = "\n\n".join(final_list)
    
    # HTML 엔티티 변환 (&apos; 등 처리)
    final_content = html.unescape(final_content)
    
    # 중복 개행 정리
    final_content = re.sub(r'\n{3,}', '\n\n', final_content)

    doc.Close(False)
    word.Quit()
    
    return final_content

# 실행
result = parse_word_keeping_order("your_file.mht")
with open("ordered_messenger_backup.md", "w", encoding="utf-8") as f:
    f.write(result)

