"""
기존 대화방별 통합 마크다운(방명.md) 또는 대화방 폴더 구조(방명/YYYY-MM-DD.md)를
날짜_방명.md 형식의 평면 구조로 재생성하는 마이그레이션 스크립트.
data/json/ 의 JSON이 source of truth이므로 JSON -> 새 구조로 재생성 후 구형 파일/폴더를 삭제하고,
outputs/dashboard.md 현황판도 함께 갱신한다.

모든 경로는 스크립트 실행 위치(CWD) 기준으로 결정된다.
config.json 이 있으면 읽고, 없으면 기본값(data/json, outputs) 사용.
"""

import os
import re
import json
import shutil
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


def load_room_aliases():
    """room_aliases.json: {"옛 제목": "canonical 방이름"} — 참여자 변동 등으로 조각난 이력 통합용"""
    path = os.path.join(CWD, "room_aliases.json")
    if os.path.exists(path):
        for enc in ["utf-8", "cp949"]:
            try:
                with open(path, "r", encoding=enc) as f:
                    return json.load(f)
            except Exception:
                continue
    return {}


def get_unique_key(msg):
    return (msg.get("date", "N/A"), msg.get("sender", "N/A"), msg.get("time", "N/A"), msg.get("content", "").strip())


def merge_messages(existing_messages, new_messages):
    seen = set(get_unique_key(m) for m in existing_messages)
    merged = list(existing_messages)
    for m in new_messages:
        key = get_unique_key(m)
        if key not in seen:
            seen.add(key)
            merged.append(m)
    return merged


def consolidate_aliased_rooms(data_dir, output_dir, aliases):
    """room_aliases.json에 등록된 옛 이름의 JSON들을 canonical 이름 하나로 병합"""
    if not aliases:
        return

    json_files = [f for f in os.listdir(data_dir) if f.endswith(".json")]
    groups = {}
    for fname in json_files:
        room_name = os.path.splitext(fname)[0]
        canonical = aliases.get(room_name, room_name)
        groups.setdefault(canonical, []).append(room_name)

    for canonical, members in groups.items():
        if len(members) == 1 and members[0] == canonical:
            continue

        merged_messages = []
        participants = []
        seen_names = set()
        latest_meta, latest_mtime = {}, -1
        for m in sorted(members):
            path = os.path.join(data_dir, f"{m}.json")
            try:
                with open(path, "r", encoding="utf-8") as f:
                    d = json.load(f)
            except Exception:
                continue
            merged_messages = merge_messages(merged_messages, d.get("messages", []))
            names = [p.strip() for p in d.get("metadata", {}).get("participants", "").split(",") if p.strip() and p.strip() != "N/A"]
            for n in names:
                if n not in seen_names:
                    seen_names.add(n)
                    participants.append(n)
            mtime = os.path.getmtime(path)
            if mtime > latest_mtime:
                latest_mtime, latest_meta = mtime, d.get("metadata", {})

        merged_metadata = dict(latest_meta)
        merged_metadata["title"] = canonical
        merged_metadata["participants"] = ", ".join(participants) if participants else "N/A"

        with open(os.path.join(data_dir, f"{canonical}.json"), "w", encoding="utf-8") as f:
            json.dump({"metadata": merged_metadata, "messages": merged_messages}, f, ensure_ascii=False, indent=2)

        for m in members:
            if m == canonical:
                continue
            try:
                os.remove(os.path.join(data_dir, f"{m}.json"))
            except Exception as e:
                print(f"    [경고] {m}.json 삭제 실패: {e}")
            cleanup_old_structures(output_dir, m)
            pattern = re.compile(r'^\d{4}-\d{2}-\d{2}_' + re.escape(m) + r'\.md$')
            for out_fname in os.listdir(output_dir):
                if pattern.match(out_fname):
                    try:
                        os.remove(os.path.join(output_dir, out_fname))
                    except Exception as e:
                        print(f"    [경고] {out_fname} 삭제 실패: {e}")

        print(f"  [병합] {', '.join(sorted(members))} -> {canonical}.json ({len(merged_messages)}개 메시지)")


def build_frontmatter(room_name, iso_date):
    return (
        f"---\n"
        f"tags:\n  - message\n"
        f"room: {room_name}\n"
        f"date: {iso_date}\n"
        f"---\n\n"
    )


def export_to_daily_markdown(output_dir, room_name, data):
    meta, messages = data.get("metadata", {}), data.get("messages", [])
    if not messages:
        return

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
        frontmatter = build_frontmatter(room_name, iso_date)
        md = frontmatter + f"# {room_name} — {iso_date}\n\n- **참석자**: {participants}\n- **업데이트**: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n---\n\n"
        for m in date_groups[date_key]:
            content = m["content"]
            if content.strip().startswith("|"):
                content = "\n" + content
            md += f"**{m['sender']}** ({iso_date} {m['time']})\n{content}\n\n"

        with open(os.path.join(output_dir, f"{iso_date}_{room_name}.md"), "w", encoding="utf-8") as f:
            f.write(md)


def cleanup_old_structures(output_dir, room_name):
    """구 구조(방명.md 단일 파일, 방명/ 폴더)를 새 평면 구조로 옮긴 뒤 남은 잔재 삭제"""
    old_merged_path = os.path.join(output_dir, f"{room_name}.md")
    if os.path.exists(old_merged_path):
        try:
            os.remove(old_merged_path)
        except Exception as e:
            print(f"    [경고] {room_name}.md 삭제 실패: {e}")

    old_room_dir = os.path.join(output_dir, room_name)
    if os.path.isdir(old_room_dir):
        try:
            shutil.rmtree(old_room_dir)
        except Exception as e:
            print(f"    [경고] {room_name}/ 폴더 삭제 실패: {e}")


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
            "dates": dates,
            "date_count": len(dates),
            "first_date": dates[0] if dates else "N/A",
            "last_date": dates[-1] if dates else "N/A",
            "mtime": os.path.getmtime(json_path),
        })
        for n in names:
            participant_rooms.setdefault(n, set()).add(room_name)

    if not rooms:
        return

    def room_link(r):
        """방 상세 섹션(헤딩)으로 이동하는 링크"""
        return f"[[dashboard#{r['room']}|{r['room']}]]"

    def date_link(r, d):
        return f"[[{d}_{r['room']}|{d}]]"

    rooms_sorted = sorted(rooms, key=lambda r: r["room"])
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
        "## 바로가기", "",
    ]
    lines.append(", ".join(room_link(r) for r in rooms_sorted))

    lines += ["", "## 최근 업데이트된 대화방 (상위 10개)", "",
              "| 대화방 | 최근 날짜 | 메시지 수 | 갱신 시각 |", "|---|---|---|---|"]
    for r in rooms_by_mtime[:10]:
        lines.append(f"| {room_link(r)} | {r['last_date']} | {r['message_count']} | {datetime.fromtimestamp(r['mtime']).strftime('%Y-%m-%d %H:%M:%S')} |")

    lines += ["", "## 대화방 목록", "", "| 대화방 | 참석자 | 최초 날짜 | 최근 날짜 | 대화일수 | 메시지 수 |", "|---|---|---|---|---|---|"]
    for r in rooms_sorted:
        lines.append(f"| {room_link(r)} | {', '.join(r['participants']) or 'N/A'} | {r['first_date']} | {r['last_date']} | {r['date_count']}일 | {r['message_count']} |")

    lines += ["", "## 대화방별 상세", ""]
    for r in rooms_sorted:
        lines.append(f"### {r['room']}")
        lines.append(f"- **참석자**: {', '.join(r['participants']) or 'N/A'}")
        lines.append(f"- **기간**: {r['first_date']} ~ {r['last_date']} ({r['date_count']}일, {r['message_count']}개 메시지)")
        lines.append("- **날짜별 이력**: " + ", ".join(date_link(r, d) for d in r["dates"]))
        lines.append("")

    lines += ["## 참여자별 대화방", ""]
    for name in sorted(participant_rooms.keys()):
        lines.append(f"- **{name}** ({len(participant_rooms[name])}개 대화방)")
        for r in rooms_sorted:
            if r["room"] in participant_rooms[name]:
                lines.append(f"  - {room_link(r)}")

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

    aliases = load_room_aliases()
    if aliases:
        print(f"room_aliases.json 발견: {len(aliases)}개 별칭 규칙 적용, 조각난 이력 통합 시작...")
        consolidate_aliased_rooms(data_dir, output_dir, aliases)
        json_files = [f for f in os.listdir(data_dir) if f.endswith(".json")]
        print()

    print(f"총 {len(json_files)}개 대화방 마이그레이션 시작...\n")

    for fname in sorted(json_files):
        room_name = os.path.splitext(fname)[0]
        json_path = os.path.join(data_dir, fname)
        try:
            with open(json_path, "r", encoding="utf-8") as f:
                data = json.load(f)
            export_to_daily_markdown(output_dir, room_name, data)
            cleanup_old_structures(output_dir, room_name)
            print(f"  [완료] {room_name} ({len(data.get('messages', []))}개 메시지)")
        except Exception as e:
            print(f"  [실패] {room_name}: {e}")

    generate_dashboard(data_dir, output_dir)
    print("\n대시보드 갱신 완료 (outputs/dashboard.md)")
    print("마이그레이션 완료.")


if __name__ == "__main__":
    migrate()
