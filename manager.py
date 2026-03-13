import os
import json
import re
from datetime import datetime
from main import parse_word_to_json

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
    
    # 기존 데이터 먼저 추가 (순서 유지)
    for m in existing_messages:
        key = get_unique_key(m)
        if key not in seen_keys:
            seen_keys.add(key)
            merged.append(m)
            
    # 신규 데이터 중복 없이 추가
    for m in new_messages:
        key = get_unique_key(m)
        if key not in seen_keys:
            seen_keys.add(key)
            merged.append(m)
            
    return merged

def export_to_markdown(room_name, data):
    """JSON 데이터를 옵시디언용 마크다운으로 변환"""
    meta = data.get('metadata', {})
    messages = data.get('messages', [])
    
    md_content = f"# {room_name}\n\n"
    md_content += f"- **최종 업데이트**: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
    md_content += f"- **대화방**: {meta.get('title', room_name)}\n"
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

def process_file(file_path):
    """단일 MHT 파일 처리 프로세스"""
    file_name = os.path.basename(file_path)
    room_match = re.match(r'^(.*)\(\d{4}-\d{2}-\d{2}\)', file_name)
    room_name = room_match.group(1).strip() if room_match else os.path.splitext(file_name)[0]
    
    print(f"\n[Processing] {file_name}")
    
    # 1. 새 데이터 파싱 (Notepad 자동화 + BS4)
    try:
        new_data = parse_word_to_json(file_path)
    except Exception as e:
        print(f"  - 파싱 실패: {e}")
        return

    new_messages = new_data.get('messages', [])
    
    # 2. 기존 데이터 로드
    json_path = os.path.join(DATA_DIR, f"{room_name}.json")
    existing_data = {"metadata": new_data['metadata'], "messages": []}
    
    if os.path.exists(json_path):
        with open(json_path, 'r', encoding='utf-8') as f:
            existing_data = json.load(f)
            
    # 3. 데이터 병합 및 중복 제거
    merged_messages = merge_messages(existing_data['messages'], new_messages)
    
    # 4. JSON 저장
    final_data = {
        "metadata": new_data['metadata'],
        "messages": merged_messages
    }
    
    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(final_data, f, ensure_ascii=False, indent=2)
    print(f"  - 누적 데이터 저장 완료 ({len(merged_messages)}개 메시지)")
    
    # 5. 마크다운 내보내기
    export_to_markdown(room_name, final_data)
    print(f"  - 마크다운 업데이트 완료: {room_name}.md")

def run_sync():
    """inputs/ 폴더의 모든 파일을 스캔하여 동기화"""
    files = [f for f in os.listdir(INPUT_DIR) if f.lower().endswith('.mht')]
    if not files:
        print("처리할 MHT 파일이 없습니다.")
        return
        
    for f in files:
        file_path = os.path.join(INPUT_DIR, f)
        process_file(file_path)

if __name__ == "__main__":
    start_t = datetime.now()
    run_sync()
    print(f"\n[완료] 총 소요시간: {datetime.now() - start_t}")
