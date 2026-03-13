import os
import json
import re
from datetime import datetime
from main import get_text_from_notepad_memory, parse_mht_html

# 경로 설정
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_DIR = os.path.join(BASE_DIR, 'inputs')
DATA_DIR = os.path.join(BASE_DIR, 'data', 'json')
OUTPUT_DIR = os.path.join(BASE_DIR, 'outputs')

# 폴더 생성 보장
for d in [INPUT_DIR, DATA_DIR, OUTPUT_DIR]:
    os.makedirs(d, exist_ok=True)

def get_unique_key(msg):
    """메시지의 고유 키 생성 (중복 제거용)"""
    return (msg.get('date', 'N/A'), 
            msg.get('sender', 'N/A'), 
            msg.get('time', 'N/A'), 
            msg.get('content', '').strip())

def merge_messages(existing_messages, new_messages):
    """기존 메시지와 신규 메시지 병합 및 중복 제거"""
    seen_keys = set()
    merged = []
    
    # 기존 데이터 먼저 추가
    for m in existing_messages:
        key = get_unique_key(m)
        if key not in seen_keys:
            seen_keys.add(key)
            merged.append(m)
            
    # 신규 데이터 중복 없이 추가
    new_count = 0
    for m in new_messages:
        key = get_unique_key(m)
        if key not in seen_keys:
            seen_keys.add(key)
            merged.append(m)
            new_count += 1
            
    return merged, new_count

def export_to_markdown(room_name, data):
    """JSON 데이터를 옵시디언용 마크다운으로 변환"""
    meta = data.get('metadata', {})
    messages = data.get('messages', [])
    
    md_content = f"# {room_name}\n\n"
    md_content += f"- **최종 업데이트**: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
    md_content += f"- **참석자**: {meta.get('participants', 'N/A')}\n\n---\n\n"
    
    current_date = ""
    for m in messages:
        if m['date'] != current_date:
            current_date = m['date']
            md_content += f"\n### 📅 {current_date}\n\n"
        
        md_content += f"**[{m['sender']}]** ({m['time']})\n{m['content']}\n\n"
        
    output_path = os.path.join(OUTPUT_DIR, f"{room_name}.md")
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(md_content)
    return output_path

def process_file(file_path):
    """단일 MHT 파일 처리 및 동기화"""
    file_name = os.path.basename(file_path)
    
    # 1. 메모장 메모리 추출 및 파싱
    raw_html = get_text_from_notepad_memory(file_path)
    if not raw_html:
        print(f"  - [실패] 메모리 추출 실패: {file_name}")
        return

    data = parse_mht_html(raw_html)
    if not data:
        print(f"  - [실패] HTML 파싱 실패: {file_name}")
        return

    # 대화방 이름 결정 (파일 이름에서 날짜 제거 혹은 메타데이터 사용)
    room_name = data['metadata']['title']
    if room_name == "N/A":
        # 파일명에서 추출 시도: 대화방명(YYYY-MM-DD).mht
        room_match = re.match(r'^(.*)\(\d{4}-\d{2}-\d{2}\)', file_name)
        room_name = room_match.group(1).strip() if room_match else os.path.splitext(file_name)[0]

    print(f"\n[Processing] {file_name} -> 대화방: {room_name}")

    # 2. 기존 데이터 로드 및 병합
    json_path = os.path.join(DATA_DIR, f"{room_name}.json")
    existing_data = {"metadata": data['metadata'], "messages": []}
    
    if os.path.exists(json_path):
        with open(json_path, 'r', encoding='utf-8') as f:
            try:
                existing_data = json.load(f)
            except:
                pass
            
    merged_messages, added_count = merge_messages(existing_data['messages'], data['messages'])
    
    # 3. JSON 저장 (누적 데이터)
    final_data = {
        "metadata": data['metadata'],
        "messages": merged_messages
    }
    
    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(final_data, f, ensure_ascii=False, indent=2)
    
    # 4. 마크다운 변환
    md_path = export_to_markdown(room_name, final_data)
    
    print(f"  - 결과 요약: 신규 {added_count}개, 누적 총 {len(merged_messages)}개 메시지")
    print(f"  - 파일 저장: {json_path}")
    print(f"  - 마크다운: {md_path}")

def run_sync():
    """inputs/ 폴더의 모든 MHT 파일을 동기화"""
    files = [f for f in os.listdir(INPUT_DIR) if f.lower().endswith('.mht')]
    if not files:
        print(f"'{INPUT_DIR}' 폴더에 처리할 MHT 파일이 없습니다.")
        return
        
    for f in files:
        file_path = os.path.join(INPUT_DIR, f)
        try:
            process_file(file_path)
        except Exception as e:
            print(f"  - 에러 발생 ({f}): {e}")

if __name__ == "__main__":
    start_time = datetime.now()
    run_sync()
    print(f"\n[완료] 전체 처리 소요 시간: {datetime.now() - start_time}")
