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

    # 1. 모든 표 가져오기 (사용자께서 검증하신 Cell 순회 방식 적용)
    table_ranges = []
    for table in doc.Tables:
        start = table.Range.Start
        end = table.Range.End
        table_ranges.append((start, end))
        
        table_md = []
        # 행과 열을 직접 순회하여 정확한 구조 파악
        for r in range(1, table.Rows.Count + 1):
            row_data = []
            for c in range(1, table.Columns.Count + 1):
                try:
                    cell_text = table.Cell(r, c).Range.Text
                    # 셀 끝 문자(\x07) 및 \r 제거, vt(\x0b)와 개행을 <br>로 변환
                    clean_text = cell_text.replace('\x07', '').replace('\r', '<br>').replace('\x0b', '<br>')
                    # Markdown 표 구분자 | 이스케이프 및 앞뒤 공백 정리
                    clean_text = clean_text.replace('|', '\|').strip()
                    # 연속된 <br> 정리
                    clean_text = re.sub(r'(<br>)+$', '', clean_text)
                    row_data.append(clean_text)
                except:
                    row_data.append("") 
            
            table_md.append(f"| {' | '.join(row_data)} |")
            if r == 1:
                table_md.append(f"| {' | '.join(['---'] * len(row_data))} |")
        
        formatted_table = "\n" + "\n".join(table_md) + "\n"
        elements.append((start, "table", formatted_table))

    # 2. 모든 문단 가져오기 (표 외부에 있는 것만)
    for para in doc.Paragraphs:
        p_start = para.Range.Start
        # 표 범위 내에 있는지 효율적으로 체크 (속도 향상 포인트)
        is_inside_table = any(start <= p_start < end for start, end in table_ranges)
        
        if not is_inside_table:
            # vt(\x0b) 문자를 \n으로 변환
            text = para.Range.Text.replace('\x0b', '\n').strip()
            if text:
                elements.append((p_start, "text", text))

    # 3. 문서 내 위치(Start 인덱스) 기준으로 정렬
    elements.sort(key=lambda x: x[0])

    # 4. 정렬된 순서대로 합치기 및 정제
    final_content = "\n\n".join([item[2] for item in elements])
    
    # HTML 엔티티 변환 (&apos; 등 처리)
    final_content = html.unescape(final_content)
    
    # 중복 개행 정리
    final_content = re.sub(r'\n{3,}', '\n\n', final_content)

    doc.Close(False)
    word.Quit()
    
    return final_content

# 실행
if __name__ == "__main__":
    input_file = "your_file.mht"
    if os.path.exists(input_file):
        print(f"파싱 시작 (정확도 우선 모드): {input_file}")
        result = parse_word_keeping_order(input_file)
        with open("ordered_messenger_backup.md", "w", encoding="utf-8") as f:
            f.write(result)
        print("파싱이 완료되었습니다.")
    else:
        print(f"파일을 찾을 수 없습니다: {input_file}")
