---
tags:
  - dashboard
---

# 메시지 대시보드

## 최근 업데이트된 대화방

```dataview
TABLE participants, length(dates) + "일" AS 대화일수
FROM #message
SORT file.mtime DESC
LIMIT 20
```

---

## 전체 대화방 목록

```dataview
TABLE participants, length(dates) + "일" AS 대화일수
FROM #message
SORT file.name ASC
```

---

## 참여자별 검색

> `"이름"` 부분을 원하는 이름으로 바꿔서 사용하세요.

```dataview
TABLE dates, participants
FROM #message
WHERE contains(participants, "이름")
SORT file.name ASC
```

---

## 날짜별 검색

> `"YYYY-MM-DD"` 부분을 원하는 날짜로 바꿔서 사용하세요.

```dataview
TABLE participants
FROM #message
WHERE contains(dates, "2026-01-01")
SORT file.name ASC
```

---

## 월별 검색

> `"YYYY-MM"` 부분을 원하는 연-월로 바꿔서 사용하세요.

```dataview
TABLE participants, length(dates) + "일" AS 대화일수
FROM #message
WHERE any(dates, (d) => startswith(d, "2026-01"))
SORT file.name ASC
```

---

## 참여자 통계 (대화방 수 기준)

```dataviewjs
const pages = dv.pages("#message");
const counter = {};
for (const p of pages) {
    for (const name of (p.participants || [])) {
        counter[name] = (counter[name] || 0) + 1;
    }
}
const rows = Object.entries(counter)
    .sort((a, b) => b[1] - a[1])
    .map(([name, count]) => [name, count + "개 대화방"]);
dv.table(["참여자", "참여 대화방 수"], rows);
```
