import os
import json
import re
import concurrent.futures
from datetime import datetime
from main import get_text_from_notepad_hidden, parse_mht_html

# 경로 설정
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_DIR = os.path.join(BASE_DIR, 'inputs')
DATA_DIR = os.path.join(BASE_DIR, 'data', 'json')
OUTPUT_DIR = os.path.join(BASE_DIR, 'outputs')

# 폴더 생성 보장
for d in [INPUT_DIR, DATA_DIR, OUTPUT_DIR]:
    os.makedirs(d, exist_ok=True)

def get_unique_key(msg):
    return (msg.get('date', 'N/A'), msg.get('sender', 'N/A'), msg.get('time', 'N/A'), msg.get('content', '').strip())

def clean_date_string(date_str):
    try:
        match = re.search(r'(\d{4})년\s*(\d{1,2})월\s*(\d{1,2})일', date_str)
        if match:
            year, month, day = match.groups()
            return f"{year}-{int(month):02d}-{int(day):02d}"
    except: pass
    return date_str.replace(' ', '_')

def merge_messages(existing_messages, new_messages):
    seen_keys = set(get_unique_key(m) for m in existing_messages)
    merged = list(existing_messages)
    added_count = 0
    for m in new_messages:
        key = get_unique_key(m)
        if key not in seen_keys:
            seen_keys.add(key)
            merged.append(m)
            added_count += 1
    return merged, added_count

def export_to_split_markdown(room_name, data):
    meta, messages = data.get('metadata', {}), data.get('messages', [])
    room_output_dir = os.path.join(OUTPUT_DIR, room_name)
    os.makedirs(room_output_dir, exist_ok=True)
    
    date_groups = {}
    for m in messages:
        date_groups.setdefault(m['date'], []).append(m)
        
    for date_key, msg_list in date_groups.items():
        file_date = clean_date_string(date_key)
        output_path = os.path.join(room_output_dir, f"{file_date}_{room_name}.md")
        
        md_content = f"# {room_name} ({date_key})\n\n- **참석자**: {meta.get('participants', 'N/A')}\n- **업데이트**: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n---\n\n"
        for m in msg_list:
            content = m['content']
            if content.strip().startswith('|'): content = "\n" + content
            md_content += f"**[{m['sender']}]** ({m['time']})\n{content}\n\n"
            
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(md_content)
    return room_output_dir

def process_file(file_path):
    """단일 파일을 백그라운드에서 파싱하고 데이터를 반환"""
    file_name = os.path.basename(file_path)
    # 메모장 숨김 실행 및 텍스트 추출
    raw_html = get_text_from_notepad_hidden(file_path)
    if not raw_html: return None
    
    data = parse_mht_html(raw_html)
    if not data: return None
    
    # 방 이름 정리
    room_name = data['metadata']['title']
    if room_name == "N/A":
        room_match = re.match(r'^(.*)\(\d{4}-\d{2}-\d{2}\)', file_name)
        room_name = room_match.group(1).strip() if room_match else os.path.splitext(file_name)[0]
    room_name = re.sub(r'[\/:*?"<>|]', '_', room_name)
    
    return {"room_name": room_name, "data": data, "file_name": file_name}

def run_sync_parallel():
    files = [os.path.join(INPUT_DIR, f) for f in os.listdir(INPUT_DIR) if f.lower().endswith('.mht')]
    if not files:
        print("처리할 MHT 파일이 없습니다.")
        return

    print(f"총 {len(files)}개 파일 백그라운드 병렬 처리 시작...")
    
    # 1. 병렬 추출 및 파싱 (최대 4개 동시 처리 - 시스템 부하 고려)
    results = []
    with concurrent.futures.ThreadPoolExecutor(max_workers=4) as executor:
        future_to_file = {executor.submit(process_file, f): f for f in files}
        for future in concurrent.futures.as_completed(future_to_file):
            res = future.result()
            if res: results.append(res)
            print(f"  - 완료: {os.path.basename(future_to_file[future])}")

    # 2. 결과 순차 병합 (데이터 무결성을 위해 병합은 순차 진행)
    for res in results:
        room_name, data = res['room_name'], res['data']
        json_path = os.path.join(DATA_DIR, f"{room_name}.json")
        
        existing_data = {"metadata": data['metadata'], "messages": []}
        if os.path.exists(json_path):
            with open(json_path, 'r', encoding='utf-8') as f:
                try: existing_data = json.load(f)
                except: pass
        
        merged_messages, added = merge_messages(existing_data['messages'], data['messages'])
        final_data = {"metadata": data['metadata'], "messages": merged_messages}
        
        with open(json_path, 'w', encoding='utf-8') as f:
            json.dump(final_data, f, ensure_ascii=False, indent=2)
        
        export_to_split_markdown(room_name, final_data)
        print(f"[병합 완료] {room_name}: 신규 {added}개 메시지 추가됨.")

if __name__ == "__main__":
    start_time = datetime.now()
    run_sync_parallel()
    print(f"\n[최종 완료] 소요 시간: {datetime.now() - start_time}")
