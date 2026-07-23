"""
기존 대화방별 통합 마크다운(방명.md)을 대화방 폴더 + 날짜별 마크다운(방명/YYYY-MM-DD.md)으로
재생성하는 마이그레이션 스크립트.
data/json/ 의 JSON이 source of truth이므로 JSON -> 새 구조로 재생성 후 구형 통합 파일을 삭제하고,
outputs/dashboard.md 현황판도 함께 갱신한다.

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


def build_frontmatter(room_name, iso_date, participants_str):
    names = [p.strip() for p in participants_str.split(",") if p.strip() and p.strip() != "N/A"]
    participant_lines = "\n".join(f"  - {n}" for n in names)
    return (
        f"---\n"
        f"tags:\n  - message\n"
        f"room: {room_name}\n"
        f"date: {iso_date}\n"
        f"participants:\n{participant_lines}\n"
        f"---\n\n"
    )


def export_to_daily_markdown(output_dir, room_name, data):
    meta, messages = data.get("metadata", {}), data.get("messages", [])
    if not messages:
        return

    room_dir = os.path.join(output_dir, room_name)
    os.makedirs(room_dir, exist_ok=True)

    date_groups = {}
    date_order = []
    for m in messages:
        d = m["date"]
        if d not in date_groups:
            date_groups[d] = []
            date_order.append(d)
        date_groups[d].append(m)

    participants = meta.get("participants", "N/A")
    for date_key in date_order:
        iso_date = clean_date_string(date_key)
        frontmatter = build_frontmatter(room_name, iso_date, participants)
        md = frontmatter + f"# {room_name} — {iso_date}\n\n- **참석자**: {participants}\n- **업데이트**: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n---\n\n"
        for m in date_groups[date_key]:
            content = m["content"]
            if content.strip().startswith("|"):
                content = "\n" + content
            md += f"**{m['sender']}** ({iso_date} {m['time']})\n{content}\n\n"

        with open(os.path.join(room_dir, f"{iso_date}.md"), "w", encoding="utf-8") as f:
            f.write(md)


def cleanup_old_merged_file(output_dir, room_name):
    """구 구조(방명.md 단일 파일)를 새 폴더 구조로 옮긴 뒤 남은 잔재 삭제"""
    old_path = os.path.join(output_dir, f"{room_name}.md")
    if os.path.exists(old_path):
        try:
            os.remove(old_path)
        except Exception as e:
            print(f"    [경고] {room_name}.md 삭제 실패: {e}")


def generate_dashboard(data_dir, output_dir):
    json_files = [f for f in os.listdir(data_dir) if f.endswith(".json")]

    rooms = []
    participant_rooms = {}
    for fname in json_files:
        json_path = os.path.join(data_dir, fname)
        try:
            with open(json_path, "r", encoding="utf-8") as f:
                data = json.load(f)
        except Exception:
            continue

        meta, messages = data.get("metadata", {}), data.get("messages", [])
        if not messages:
            continue

        room_name = os.path.splitext(fname)[0]
        dates = sorted({clean_date_string(m["date"]) for m in messages})
        names = [p.strip() for p in meta.get("participants", "N/A").split(",") if p.strip() and p.strip() != "N/A"]

        rooms.append({
            "room": room_name,
            "participants": names,
            "message_count": len(messages),
            "date_count": len(dates),
            "last_date": dates[-1] if dates else "N/A",
            "mtime": os.path.getmtime(json_path),
        })
        for n in names:
            participant_rooms.setdefault(n, set()).add(room_name)

    if not rooms:
        return

    def room_link(r):
        return f"[[{r['room']}/{r['last_date']}|{r['room']}]]"

    rooms_by_mtime = sorted(rooms, key=lambda r: r["mtime"], reverse=True)
    last_room = rooms_by_mtime[0]
    total_messages = sum(r["message_count"] for r in rooms)

    lines = [
        "---", "tags:", "  - dashboard", "---", "",
        "# 메시지 대시보드", "",
        f"- **마지막 갱신**: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
        f"- **최근 업데이트 대화방**: {room_link(last_room)} ({datetime.fromtimestamp(last_room['mtime']).strftime('%Y-%m-%d %H:%M:%S')})",
        f"- **전체 대화방 수**: {len(rooms)}개",
        f"- **전체 메시지 수**: {total_messages}개",
        "",
        "## 최근 업데이트된 대화방 (상위 10개)", "",
        "| 대화방 | 최근 날짜 | 메시지 수 | 갱신 시각 |",
        "|---|---|---|---|",
    ]
    for r in rooms_by_mtime[:10]:
        lines.append(f"| {room_link(r)} | {r['last_date']} | {r['message_count']} | {datetime.fromtimestamp(r['mtime']).strftime('%Y-%m-%d %H:%M:%S')} |")

    lines += ["", "## 대화방별 목록", "", "| 대화방 | 참석자 | 대화일수 | 메시지 수 | 최근 날짜 |", "|---|---|---|---|---|"]
    for r in sorted(rooms, key=lambda r: r["room"]):
        lines.append(f"| {room_link(r)} | {', '.join(r['participants']) or 'N/A'} | {r['date_count']}일 | {r['message_count']} | {r['last_date']} |")

    lines += ["", "## 참여자별 대화방", ""]
    for name in sorted(participant_rooms.keys()):
        links = ", ".join(room_link(r) for r in sorted(rooms, key=lambda r: r["room"]) if r["room"] in participant_rooms[name])
        lines.append(f"- **{name}** ({len(participant_rooms[name])}개 대화방): {links}")

    with open(os.path.join(output_dir, "dashboard.md"), "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")


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
            export_to_daily_markdown(output_dir, room_name, data)
            cleanup_old_merged_file(output_dir, room_name)
            print(f"  [완료] {room_name}/ ({len(data.get('messages', []))}개 메시지)")
        except Exception as e:
            print(f"  [실패] {room_name}: {e}")

    generate_dashboard(data_dir, output_dir)
    print("\n대시보드 갱신 완료 (outputs/dashboard.md)")
    print("마이그레이션 완료.")


if __name__ == "__main__":
    migrate()
