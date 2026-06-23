"""
기존 날짜별 분리 마크다운(YYYY-MM-DD_방명.md)을 대화방별 단일 파일(방명.md)로 통합하는 마이그레이션 스크립트.
data/json/ 의 JSON이 source of truth이므로 JSON → 새 마크다운으로 재생성 후 구형 파일 삭제.

모든 경로는 스크립트 실행 위치(CWD) 기준으로 결정된다.
config.json 이 있으면 읽고, 없으면 기본값(data/json, outputs) 사용.
"""

import os
import re
import json
from datetime import datetime


def clean_date_string(date_str):
    try:
        match = re.search(r'(\d{4})년\s*(\d{1,2})월\s*(\d{1,2})일', date_str)
        if match:
            year, month, day = match.groups()
            return f"{year}-{int(month):02d}-{int(day):02d}"
    except Exception:
        pass
    return date_str.replace(' ', '_')


CWD = os.getcwd()


def load_config():
    cfg = {"data_dir": "data/json", "output_dir": "outputs"}
    config_path = os.path.join(CWD, "config.json")
    if os.path.exists(config_path):
        for enc in ["utf-8", "cp949"]:
            try:
                with open(config_path, "r", encoding=enc) as f:
                    loaded = json.load(f)
                    cfg.update({k: v for k, v in loaded.items() if k in cfg})
                break
            except Exception:
                continue
    return cfg


def resolve(path_str):
    if os.path.isabs(path_str):
        return path_str
    return os.path.abspath(os.path.join(CWD, path_str))


def cleanup_legacy_split_files(output_dir, room_name):
    pattern = re.compile(r"^\d{4}-\d{2}-\d{2}_" + re.escape(room_name) + r"\.md$")
    for fname in os.listdir(output_dir):
        if pattern.match(fname):
            try:
                os.remove(os.path.join(output_dir, fname))
            except Exception as e:
                print(f"    [경고] {fname} 삭제 실패: {e}")


def build_frontmatter(date_order, participants_str):
    tag_set = []
    seen_year = set()
    seen_month = set()
    for d in date_order:
        iso = clean_date_string(d)
        year, month = iso[:4], iso[:7]
        if year not in seen_year:
            tag_set.append(f"message/{year}")
            seen_year.add(year)
        if month not in seen_month:
            tag_set.append(f"message/{month}")
            seen_month.add(month)
        tag_set.append(f"message/{iso}")

    for name in (p.strip() for p in participants_str.split(",") if p.strip() and p.strip() != "N/A"):
        tag_set.append(f"sender/{name}")

    tag_lines = "\n".join(f"  - {t}" for t in tag_set)
    return f"---\ntags:\n{tag_lines}\nparticipants: {participants_str}\n---\n\n"


def export_to_merged_markdown(output_dir, room_name, data):
    meta = data.get("metadata", {})
    messages = data.get("messages", [])
    output_path = os.path.join(output_dir, f"{room_name}.md")

    date_groups = {}
    date_order = []
    for m in messages:
        d = m["date"]
        if d not in date_groups:
            date_groups[d] = []
            date_order.append(d)
        date_groups[d].append(m)

    participants = meta.get("participants", "N/A")
    frontmatter = build_frontmatter(date_order, participants)
    md = frontmatter + f"# {room_name}\n\n- **참석자**: {participants}\n- **업데이트**: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n---\n\n"
    for date_key in date_order:
        iso_date = clean_date_string(date_key)
        md += f"## {date_key}\n\n"
        for m in date_groups[date_key]:
            content = m["content"]
            if content.strip().startswith("|"):
                content = "\n" + content
            md += f"**{m['sender']}** ({iso_date} {m['time']})\n{content}\n\n"

    with open(output_path, "w", encoding="utf-8") as f:
        f.write(md)


def migrate():
    cfg = load_config()
    data_dir = resolve(cfg["data_dir"])
    output_dir = resolve(cfg["output_dir"])

    if not os.path.isdir(data_dir):
        print(f"JSON 폴더를 찾을 수 없습니다: {data_dir}")
        return

    os.makedirs(output_dir, exist_ok=True)

    json_files = [f for f in os.listdir(data_dir) if f.endswith(".json")]
    if not json_files:
        print(f"JSON 파일이 없습니다: {data_dir}")
        return

    print(f"실행 경로: {CWD}")
    print(f"JSON 경로: {data_dir}")
    print(f"출력 경로: {output_dir}")
    print(f"총 {len(json_files)}개 대화방 마이그레이션 시작...\n")

    for fname in sorted(json_files):
        room_name = os.path.splitext(fname)[0]
        json_path = os.path.join(data_dir, fname)
        try:
            with open(json_path, "r", encoding="utf-8") as f:
                data = json.load(f)
            export_to_merged_markdown(output_dir, room_name, data)
            cleanup_legacy_split_files(output_dir, room_name)
            print(f"  [완료] {room_name}.md ({len(data.get('messages', []))}개 메시지)")
        except Exception as e:
            print(f"  [실패] {room_name}: {e}")

    print("\n마이그레이션 완료.")


if __name__ == "__main__":
    migrate()
