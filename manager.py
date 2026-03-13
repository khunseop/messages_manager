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

def clean_date_string(date_str):
    """'2026년 3월 13일 금요일' -> '2026-03-13' 형식으로 변환"""
    try:
        match = re.search(r'(\d{4})년\s*(\d{1,2})월\s*(\d{1,2})일', date_str)
        if match:
            year, month, day = match.groups()
            return f"{year}-{int(month):02d}-{int(day):02d}"
    except:
        pass
    return date_str.replace(' ', '_')

def merge_messages(existing_messages, new_messages):
    """기존 메시지와 신규 메시지 병합 및 중복 제거"""
    seen_keys = set()
    merged = []
    
    for m in existing_messages:
        key = get_unique_key(m)
        if key not in seen_keys:
            seen_keys.add(key)
            merged.append(m)
            
    added_count = 0
    for m in new_messages:
        key = get_unique_key(m)
        if key not in seen_keys:
            seen_keys.add(key)
            merged.append(m)
            added_count += 1
            
    return merged, added_count

def export_to_split_markdown(room_name, data):
    """JSON 데이터를 대화방 폴더 내 날짜별 마크다운 파일로 분리 저장"""
    meta = data.get('metadata', {})
    messages = data.get('messages', [])
    
    # 대화방별 전용 폴더 생성
    room_output_dir = os.path.join(OUTPUT_DIR, room_name)
    os.makedirs(room_output_dir, exist_ok=True)
    
    # 날짜별로 메시지 그룹화
    date_groups = {}
    for m in messages:
        date_key = m['date']
        if date_key not in date_groups:
            date_groups[date_key] = []
        date_groups[date_key].append(m)
        
    for date_key, msg_list in date_groups.items():
        file_date = clean_date_string(date_key)
        file_name = f"{file_date}_{room_name}.md"
        output_path = os.path.join(room_output_dir, file_name)
        
        md_content = f"# {room_name} ({date_key})\n\n"
        md_content += f"- **참석자**: {meta.get('participants', 'N/A')}\n"
        md_content += f"- **파일 생성일**: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n---\n\n"
        
        for m in msg_list:
            content = m['content']
            if content.strip().startswith('|'):
                content = "\n" + content
            md_content += f"**[{m['sender']}]** ({m['time']})\n{content}\n\n"
            
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(md_content)
            
    return room_output_dir

def process_file(file_path):
    """단일 MHT 파일 처리 및 동기화"""
    file_name = os.path.basename(file_path)
    
    raw_html = get_text_from_notepad_memory(file_path)
    if not raw_html:
        print(f"  - [실패] 메모리 추출 실패: {file_name}")
        return

    data = parse_mht_html(raw_html)
    if not data:
        print(f"  - [실패] HTML 파싱 실패: {file_name}")
        return

    room_name = data['metadata']['title']
    if room_name == "N/A":
        room_match = re.match(r'^(.*)\(\d{4}-\d{2}-\d{2}\)', file_name)
        room_name = room_match.group(1).strip() if room_match else os.path.splitext(file_name)[0]

    # 특수문자 제거 (폴더명용)
    room_name = re.sub(r'[\/:*?"<>|]', '_', room_name)

    print(f"\n[Processing] {file_name} -> 대화방: {room_name}")

    json_path = os.path.join(DATA_DIR, f"{room_name}.json")
    existing_data = {"metadata": data['metadata'], "messages": []}
    
    if os.path.exists(json_path):
        with open(json_path, 'r', encoding='utf-8') as f:
            try:
                existing_data = json.load(f)
            except: pass
            
    merged_messages, added_count = merge_messages(existing_data['messages'], data['messages'])
    
    final_data = {
        "metadata": data['metadata'],
        "messages": merged_messages
    }
    
    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(final_data, f, ensure_ascii=False, indent=2)
    
    # 날짜별 분리 내보내기 호출
    room_dir = export_to_split_markdown(room_name, final_data)
    
    print(f"  - 결과 요약: 신규 {added_count}개, 누적 총 {len(merged_messages)}개 메시지")
    print(f"  - 저장 위치: {room_dir}")

def run_sync():
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
