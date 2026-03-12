import win32com.client as win32
import os
import html
import re

class MessengerParser:
    def __init__(self):
        self.word = None

    def _ensure_word(self):
        if self.word is None:
            self.word = win32.gencache.EnsureDispatch('Word.Application')
            self.word.Visible = False
            self.word.DisplayAlerts = 0 # wdAlertsNone

    def parse_mht(self, mht_path):
        self._ensure_word()
        abs_path = os.path.abspath(mht_path)
        doc = self.word.Documents.Open(abs_path, ReadOnly=True, Visible=False)
        
        elements = []

        # 1. 표 정보 미리 추출 (범위 파악 및 데이터 벌크 로드)
        # table.Range.Text는 셀 구분자로 \x07을 사용함
        table_ranges = []
        for table in doc.Tables:
            start = table.Range.Start
            end = table.Range.End
            table_ranges.append((start, end))
            
            # 셀 단위 접근 대신 전체 텍스트를 가져와서 파싱 (속도 핵심)
            raw_text = table.Range.Text
            # Word 표 텍스트 특성: 셀 끝은 \x07, 행 끝은 \r\x07
            rows = raw_text.strip('\r\x07').split('\r\x07')
            
            table_md = []
            for i, row in enumerate(rows):
                # 셀 분리 및 특수문자 제거
                cells = [c.replace('\x07', '').replace('\r', '').strip() for c in row.split('\x07') if c]
                if not cells: continue
                
                table_md.append(f"| {' | '.join(cells)} |")
                if i == 0: # 헤더 구분선
                    table_md.append(f"| {' | '.join(['---'] * len(cells))} |")
            
            formatted_table = "\n" + "\n".join(table_md) + "\n"
            elements.append((start, "table", formatted_table))

        # 2. 문단 추출 (표 안에 없는 것만)
        # 모든 문단을 순회하되, 이미 표 범위에 포함된 Start 지점은 건너뜀
        for para in doc.Paragraphs:
            p_start = para.Range.Start
            
            # 표 내부에 있는지 확인 (COM 호출 대신 미리 계산한 범위로 체크)
            is_inside_table = any(start <= p_start < end for start, end in table_ranges)
            
            if not is_inside_table:
                text = para.Range.Text.strip()
                if text:
                    elements.append((p_start, "text", text))

        # 3. 정렬 및 병합
        elements.sort(key=lambda x: x[0])
        final_content = "\n\n".join([item[2] for item in elements])
        
        # 후처리
        final_content = html.unescape(final_content)
        final_content = re.sub(r'\n{3,}', '\n\n', final_content)

        doc.Close(False)
        return final_content

    def quit(self):
        if self.word:
            self.word.Quit()
            self.word = None

# 실행 예시
if __name__ == "__main__":
    parser = MessengerParser()
    try:
        # 단일 파일 처리 (나중에 loop로 확장 가능)
        result = parser.parse_mht("your_file.mht")
        with open("ordered_messenger_backup.md", "w", encoding="utf-8") as f:
            f.write(result)
    finally:
        parser.quit()

