"""
기존 날짜별 분리 마크다운(YYYY-MM-DD_방명.md)을 대화방별 단일 파일(방명.md)로 통합하는 마이그레이션 스크립트.
data/json/ 의 JSON이 source of truth이므로 JSON → 새 마크다운으로 재생성 후 구형 파일 삭제.
"""

import os
import sys

# main.py 와 동일 경로에서 실행 가정
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from main import DATA_DIR, OUTPUT_DIR, export_to_merged_markdown, cleanup_legacy_split_files
import json


def migrate():
    json_files = [f for f in os.listdir(DATA_DIR) if f.endswith('.json')]
    if not json_files:
        print("data/json/ 에 JSON 파일이 없습니다.")
        return

    print(f"총 {len(json_files)}개 대화방 마이그레이션 시작...")
    for fname in sorted(json_files):
        room_name = os.path.splitext(fname)[0]
        json_path = os.path.join(DATA_DIR, fname)
        try:
            with open(json_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            export_to_merged_markdown(room_name, data)
            print(f"  [완료] {room_name}.md ({len(data.get('messages', []))}개 메시지)")
        except Exception as e:
            print(f"  [실패] {room_name}: {e}")

    print("\n마이그레이션 완료.")


if __name__ == "__main__":
    migrate()
