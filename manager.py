import os
import json
import re
import time
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
    """'2026년 3월 13일 금요일' -> '2026-03-13' 형식으로 변환"""
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
    """JSON 데이터를 대화방 폴더 내 날짜별 마크다운 파일로 분리 저장"""
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
    """단일 파일을 순차적으로 파싱하고 결과를 즉시 저장/병합"""
    file_name = os.path.basename(file_path)
    
    max_retries = 3
    for attempt in range(1, max_retries + 1):
        try:
            if attempt > 1:
                print(f"  - [재시도 {attempt}/{max_retries}] {file_name}")
                time.sleep(1)

            # 1. 텍스트 추출 (HIDE 모드)
            raw_html = get_text_from_notepad_hidden(file_path)
            if not raw_html: continue
            
            # 2. 파싱
            data = parse_mht_html(raw_html)
            if not data: continue
            
            # 3. 방 이름 결정: 메타데이터 우선 사용 (파일명 날짜 무시)
            room_name = data['metadata']['title']
            if room_name == "N/A" or not room_name:
                # 파일명에서 날짜 부분 제거하고 순수 방 이름만 추출
                room_name = re.sub(r'\(\d{4}-\d{2}-\d{2}\)', '', file_name)
                room_name = os.path.splitext(room_name)[0].strip()
            
            # 메타데이터 제목 뒤에 날짜가 붙어있는 경우도 제거
            room_name = re.sub(r'\(\d{4}-\d{2}-\d{2}\)', '', room_name).strip()
            room_name = re.sub(r'[\/:*?"<>|]', '_', room_name)
            
            # 4. 데이터 병합 (JSON)
            json_path = os.path.join(DATA_DIR, f"{room_name}.json")
            existing_data = {"metadata": data['metadata'], "messages": []}
            if os.path.exists(json_path):
                with open(json_path, 'r', encoding='utf-8') as f:
                    try: existing_data = json.load(f)
                    except: pass
            
            merged_messages, added = merge_messages(existing_data['messages'], data['messages'])
            # 메타데이터도 최신 정보(특히 start_date)로 갱신 가능하나, 
            # 중복 제거된 전체 메시지 리스트를 보존하는 것이 핵심
            final_data = {"metadata": data['metadata'], "messages": merged_messages}
            
            with open(json_path, 'w', encoding='utf-8') as f:
                json.dump(final_data, f, ensure_ascii=False, indent=2)
            
            # 5. 마크다운 생성 (내부 날짜별로 분리됨)
            export_to_split_markdown(room_name, final_data)
            print(f"  [성공] {room_name}: 신규 {added}개 추가 (총 {len(merged_messages)}개)")
            return True
            
        except Exception as e:
            print(f"  [에러] 시도 {attempt}: {e}")
            
    print(f"  [최종 실패] {file_name}")
    return False

def run_sync_sequential():
    files = [os.path.join(INPUT_DIR, f) for f in os.listdir(INPUT_DIR) if f.lower().endswith('.mht')]
    if not files:
        print("처리할 MHT 파일이 없습니다.")
        return

    print(f"총 {len(files)}개 파일 처리 시작 (비가시적 모드)...")
    success_count = 0
    for i, f in enumerate(files):
        print(f"[{i+1}/{len(files)}] {os.path.basename(f)}")
        if process_file(f): success_count += 1
        time.sleep(0.3)

    print(f"\n[최종 요약] 전체 {len(files)}개 중 {success_count}개 성공 완료.")

if __name__ == "__main__":
    start_time = datetime.now()
    run_sync_sequential()
    print(f"[완료] 전체 소요 시간: {datetime.now() - start_time}")
