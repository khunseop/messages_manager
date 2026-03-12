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

    # 1. 표 정보 미리 추출 (속도 최적화를 위해 벌크 데이터 로드)
    table_ranges = []
    for table in doc.Tables:
        start = table.Range.Start
        end = table.Range.End
        table_ranges.append((start, end))
        
        # Word 표 텍스트 특징: 셀 끝은 \x07, 행 끝은 \r\x07
        raw_text = table.Range.Text
        
        # 행 단위로 분리 (\r\x07 기준)
        rows_raw = raw_text.split('\r\x07')
        if rows_raw and not rows_raw[-1].strip('\x07'):
            rows_raw.pop() # 마지막 빈 행 제거
            
        table_md = []
        for i, row_raw in enumerate(rows_raw):
            # 셀 단위로 분리 (\x07 기준)
            cells_raw = row_raw.split('\x07')
            if cells_raw and not cells_raw[-1]:
                cells_raw.pop() # 행 끝의 구분자로 인해 생기는 빈 셀 제거
            
            clean_cells = []
            for cell in cells_raw:
                # Markdown 표 문법 보호 및 개행 처리
                c = cell.replace('|', '\|') # 파이프 기호 이스케이프
                c = c.replace('\x0b', '<br>').replace('\r', '<br>') # 셀 내 개행을 <br>로 변환
                # 중복된 <br> 및 앞뒤 공백 정리
                c = re.sub(r'(<br>)+$', '', c.strip())
                clean_cells.append(c)
            
            if not clean_cells:
                continue
                
            table_md.append(f"| {' | '.join(clean_cells)} |")
            
            # 헤더 아래 구분선 (첫 행 기준)
            if i == 0:
                table_md.append(f"| {' | '.join(['---'] * len(clean_cells))} |")
        
        formatted_table = "\n" + "\n".join(table_md) + "\n"
        elements.append((start, "table", formatted_table))

    # 2. 문단 가져오기 (표 외부에 있는 것만)
    for para in doc.Paragraphs:
        p_start = para.Range.Start
        # 표 범위 내에 있는지 효율적으로 체크
        is_inside_table = any(start <= p_start < end for start, end in table_ranges)
        
        if not is_inside_table:
            text = para.Range.Text.replace('\x0b', '\n').strip()
            if text:
                elements.append((p_start, "text", text))

    # 3. 정렬 및 병합
    elements.sort(key=lambda x: x[0])
    final_content = "\n\n".join([item[2] for item in elements])
    
    # 마무리 정제
    final_content = html.unescape(final_content)
    final_content = re.sub(r'\n{3,}', '\n\n', final_content)

    doc.Close(False)
    word.Quit()
    
    return final_content

# 실행
if __name__ == "__main__":
    input_file = "your_file.mht"
    if os.path.exists(input_file):
        print(f"파싱 시작 (고속 모드): {input_file}")
        result = parse_word_keeping_order(input_file)
        with open("ordered_messenger_backup.md", "w", encoding="utf-8") as f:
            f.write(result)
        print("파싱이 완료되었습니다.")
    else:
        print(f"파일을 찾을 수 없습니다: {input_file}")
