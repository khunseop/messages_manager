import win32com.client as win32
import os
import html
import re

def parse_word_keeping_order(mht_path):
    # Word 어플리케이션 실행
    word = win32.gencache.EnsureDispatch('Word.Application')
    word.Visible = False
    
    abs_path = os.path.abspath(mht_path)
    doc = word.Documents.Open(abs_path, ReadOnly=True)
    
    elements = []

    # 1. 모든 문단 가져오기 (표 안에 있는 문단은 제외)
    for para in doc.Paragraphs:
        # 12 = wdWithinTable (표 내부 여부 확인)
        if not para.Range.Information(12): 
            # \x0b (vt) 문자를 \n (개행)으로 변환
            text = para.Range.Text.replace('\x0b', '\n').strip()
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
                    # \r 제거, \x07 (셀 구분자) 제거, \x0b (vt) 를 \n 으로 변환
                    clean_text = cell_text.replace('\r', '').replace('\x07', '').replace('\x0b', '\n').strip()
                    row_data.append(clean_text)
                except:
                    row_data.append("") 
            
            table_md.append(f"| {' | '.join(row_data)} |")
            if r == 1:
                # 헤더 구분선 추가
                table_md.append(f"| {' | '.join(['---'] * len(row_data))} |")
        
        formatted_table = "\n" + "\n".join(table_md) + "\n"
        # (시작 위치, 유형, 내용) 저장
        elements.append((table.Range.Start, "table", formatted_table))

    # 3. 문서 내 위치(Start 인덱스) 기준으로 정렬
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
if __name__ == "__main__":
    # 파일명이 한글이거나 경로가 복잡할 수 있으므로 절대 경로 사용 권장
    input_file = "your_file.mht"
    if os.path.exists(input_file):
        result = parse_word_keeping_order(input_file)
        with open("ordered_messenger_backup.md", "w", encoding="utf-8") as f:
            f.write(result)
        print("파싱이 완료되었습니다.")
    else:
        print(f"파일을 찾을 수 없습니다: {input_file}")
