import win32com.client as win32
import os
import html
import re

def parse_word_keeping_order(mht_path):
    # Word 어플리케이션 실행
    word = win32.gencache.EnsureDispatch('Word.Application')
    word.Visible = False
    word.DisplayAlerts = 0  # wdAlertsNone
    
    abs_path = os.path.abspath(mht_path)
    doc = word.Documents.Open(abs_path, ReadOnly=True, Visible=False)
    
    elements = []

    # 1. 표 정보 미리 추출 (범위 파악 및 데이터 벌크 로드 - 속도 최적화)
    table_ranges = []
    for table in doc.Tables:
        start = table.Range.Start
        end = table.Range.End
        table_ranges.append((start, end))
        
        # [속도 최적화] Cell 하나씩 접근하지 않고 전체 텍스트를 한 번에 가져옴
        # Word 표 텍스트: 셀 구분자는 \x07, 행 구분자는 \r\x07
        raw_text = table.Range.Text
        # 행 단위로 분리 (마지막 \r\x07 제거 후 분리)
        rows = raw_text.strip('\r\x07').split('\r\x07')
        
        table_md = []
        for i, row in enumerate(rows):
            # 셀 분리 (\x07 기준)
            raw_cells = [c for c in row.split('\x07') if c]
            if not raw_cells: continue
            
            clean_cells = []
            for cell in raw_cells:
                # [표 깨짐 방지] 셀 내부의 \x0b(vt), \r(개행) 등을 Markdown에서 허용하는 <br>로 변환
                # 또한 Markdown 표 구분자인 | 도 이스케이프 처리
                c = cell.replace('|', '\|')
                c = c.replace('\x0b', '<br>').replace('\r', '<br>')
                # 마지막에 남는 중복 <br> 및 공백 정리
                c = re.sub(r'(<br>)+$', '', c.strip())
                clean_cells.append(c)
            
            table_md.append(f"| {' | '.join(clean_cells)} |")
            
            # 헤더 아래 구분선 추가 (첫 번째 행 뒤에)
            if i == 0:
                table_md.append(f"| {' | '.join(['---'] * len(clean_cells))} |")
        
        formatted_table = "\n" + "\n".join(table_md) + "\n"
        elements.append((start, "table", formatted_table))

    # 2. 모든 문단 가져오기 (표 안에 있는 문단은 제외)
    for para in doc.Paragraphs:
        p_start = para.Range.Start
        
        # [속도 최적화] COM 호출 대신 미리 저장된 표 범위로 체크
        is_inside_table = any(start <= p_start < end for start, end in table_ranges)
        
        if not is_inside_table:
            # 일반 문단 내 소프트 개행(\x0b)은 \n으로 변환
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
    # Word 종료는 함수 외부에서 관리하거나 매번 종료
    word.Quit()
    
    return final_content

# 실행
if __name__ == "__main__":
    input_file = "your_file.mht"
    if os.path.exists(input_file):
        print(f"파싱 시작: {input_file}")
        result = parse_word_keeping_order(input_file)
        with open("ordered_messenger_backup.md", "w", encoding="utf-8") as f:
            f.write(result)
        print("파싱이 완료되었습니다.")
    else:
        print(f"파일을 찾을 수 없습니다: {input_file}")
