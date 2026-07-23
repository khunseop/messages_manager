---
tags:
  - dashboard
---

# 메시지 대시보드

## 최근 업데이트된 대화방

```dataview
TABLE first(rows.participants) AS participants, max(rows.date) AS 최근날짜, length(rows) + "일" AS 대화일수
FROM #message
GROUP BY room
SORT max(rows.file.mtime) DESC
LIMIT 20
```

---

## 전체 대화방 목록

```dataview
TABLE first(rows.participants) AS participants, length(rows) + "일" AS 대화일수
FROM #message
GROUP BY room
SORT room ASC
```

---

## 참여자별 검색

> `"이름"` 부분을 원하는 이름으로 바꿔서 사용하세요.

```dataview
TABLE room, date
FROM #message
WHERE contains(participants, "이름")
SORT room ASC, date ASC
```

---

## 날짜별 검색

> `"YYYY-MM-DD"` 부분을 원하는 날짜로 바꿔서 사용하세요.

```dataview
TABLE room, participants
FROM #message
WHERE date = "2026-01-01"
SORT room ASC
```

---

## 월별 검색

> `"YYYY-MM"` 부분을 원하는 연-월로 바꿔서 사용하세요.

```dataview
TABLE first(rows.participants) AS participants, length(rows) + "일" AS 대화일수
FROM #message
WHERE startswith(date, "2026-01")
GROUP BY room
SORT room ASC
```

---

## 참여자 통계 (대화방 수 기준)

```dataviewjs
const pages = dv.pages("#message");
const roomParticipants = {};
for (const p of pages) {
    if (!p.room) continue;
    roomParticipants[p.room] = p.participants || [];
}
const counter = {};
for (const names of Object.values(roomParticipants)) {
    for (const name of names) {
        counter[name] = (counter[name] || 0) + 1;
    }
}
const rows = Object.entries(counter)
    .sort((a, b) => b[1] - a[1])
    .map(([name, count]) => [name, count + "개 대화방"]);
dv.table(["참여자", "참여 대화방 수"], rows);
```
