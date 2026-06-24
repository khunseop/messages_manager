"""
기존 JSON의 sender 필드에 남아있는 대괄호를 제거하는 스크립트.
실행 후 migrate_to_merged.py 를 이어서 실행하면 MD 파일까지 반영된다.

모든 경로는 스크립트 실행 위치(CWD) 기준으로 결정된다.
"""

import os
import json

CWD = os.getcwd()


def load_config():
    cfg = {"data_dir": "data/json"}
    config_path = os.path.join(CWD, "config.json")
    if os.path.exists(config_path):
        for enc in ["utf-8", "cp949"]:
            try:
                with open(config_path, "r", encoding=enc) as f:
                    loaded = json.load(f)
                    if "data_dir" in loaded:
                        cfg["data_dir"] = loaded["data_dir"]
                break
            except Exception:
                continue
    return cfg


def resolve(path_str):
    if os.path.isabs(path_str):
        return path_str
    return os.path.abspath(os.path.join(CWD, path_str))


def fix_senders(data_dir):
    json_files = [f for f in os.listdir(data_dir) if f.endswith(".json")]
    if not json_files:
        print(f"JSON 파일이 없습니다: {data_dir}")
        return

    print(f"실행 경로: {CWD}")
    print(f"JSON 경로: {data_dir}")
    print(f"총 {len(json_files)}개 파일 처리 시작...\n")

    for fname in sorted(json_files):
        json_path = os.path.join(data_dir, fname)
        try:
            with open(json_path, "r", encoding="utf-8") as f:
                data = json.load(f)

            fixed = 0
            for msg in data.get("messages", []):
                original = msg.get("sender", "")
                cleaned = original.strip("[]")
                if cleaned != original:
                    msg["sender"] = cleaned
                    fixed += 1

            with open(json_path, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)

            status = f"{fixed}개 수정" if fixed else "변경 없음"
            print(f"  [완료] {fname} ({status})")
        except Exception as e:
            print(f"  [실패] {fname}: {e}")

    print("\nJSON 정리 완료. migrate_to_merged.py 를 실행해 MD 파일을 재생성하세요.")


if __name__ == "__main__":
    cfg = load_config()
    fix_senders(resolve(cfg["data_dir"]))
