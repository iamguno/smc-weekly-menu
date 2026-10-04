#!/usr/bin/env python3
"""Build index.html from weekly cafeteria spreadsheets."""
from __future__ import annotations

import json
import re
import subprocess
import unicodedata
import xml.etree.ElementTree as ET
import zipfile
from datetime import datetime, timedelta
from pathlib import Path

NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
EXCEL_EPOCH = datetime(1899, 12, 30)
ROOT = Path(__file__).resolve().parent
WEEKDAYS_KO = ["월", "화", "수", "목", "금", "토", "일"]


def nfc(s: str) -> str:
    return unicodedata.normalize("NFC", s)


def find_xlsx(keyword: str) -> Path:
    for f in ROOT.glob("*.xlsx"):
        if keyword in nfc(f.name):
            return f
    raise FileNotFoundError(keyword)


def col_row(ref: str) -> tuple[int, int]:
    col = ""
    row = ""
    for ch in ref:
        if ch.isalpha():
            col += ch
        else:
            row += ch
    n = 0
    for ch in col:
        n = n * 26 + (ord(ch) - 64)
    return n, int(row)


def parse_merge_ref(ref: str) -> tuple[int, int, int, int]:
    a, b = ref.split(":")
    c1, r1 = col_row(a)
    c2, r2 = col_row(b)
    return r1, c1, r2, c2


def load_sheet(path: Path):
    with zipfile.ZipFile(path) as z:
        root = ET.fromstring(z.read("xl/sharedStrings.xml"))
        ss = [
            "".join(t.text or "" for t in si.findall(".//m:t", NS))
            for si in root.findall("m:si", NS)
        ]
        ws = ET.fromstring(z.read("xl/worksheets/sheet1.xml"))
        cells = {}
        for c in ws.findall(".//m:c", NS):
            ref = c.attrib.get("r")
            if not ref:
                continue
            col, row = col_row(ref)
            t = c.attrib.get("t")
            v = c.find("m:v", NS)
            is_el = c.find("m:is", NS)
            val = None
            if t == "s" and v is not None and v.text:
                val = ss[int(v.text)]
            elif t == "inlineStr" and is_el is not None:
                val = "".join(x.text or "" for x in is_el.findall(".//m:t", NS))
            elif v is not None:
                val = v.text
            if val is not None:
                cells[(row, col)] = str(val)
        merges = [
            parse_merge_ref(m.attrib["ref"]) for m in ws.findall(".//m:mergeCell", NS)
        ]
    return cells, merges


def in_merge(merges, r, c):
    for r1, c1, r2, c2 in merges:
        if r1 <= r <= r2 and c1 <= c <= c2:
            return r1, c1, r2, c2
    return None


def visible_cell(cells, merges, r, c):
    m = in_merge(merges, r, c)
    if m and (r, c) != (m[0], m[1]):
        return None
    v = cells.get((r, c))
    if v is None:
        return None
    v = v.replace("\r", "").strip()
    return v or None


def excel_date(v: str):
    try:
        n = float(v)
        if n > 40000:
            return EXCEL_EPOCH + timedelta(days=int(n))
    except ValueError:
        pass
    m = re.search(r"(\d{1,2})\s*/\s*(\d{1,2})", v.replace("\n", " "))
    if m:
        return datetime(2026, int(m.group(1)), int(m.group(2)))
    return None


def clean_item(s: str) -> str:
    s = s.replace("\n", " ")
    s = re.sub(r"[ \t]+", " ", s).strip()
    return s


def collect_items(cells, merges, rows, cols):
    items = []
    seen = set()
    for r in rows:
        for c in cols:
            v = visible_cell(cells, merges, r, c)
            if not v:
                continue
            v = clean_item(v)
            if v in {"♡", "♥", "❤"}:
                continue
            if v in seen:
                continue
            seen.add(v)
            items.append(v)
    return items


def date_blocks(cells, merges, max_row, last_end):
    starts = []
    for r in range(1, max_row + 1):
        v = visible_cell(cells, merges, r, 1) or cells.get((r, 1))
        if not v:
            continue
        d = excel_date(v)
        if d:
            starts.append((r, d))
    if starts:
        last_start = starts[-1][0]
        for r in range(last_start + 1, max_row + 1):
            if cells.get((r, 1), "").strip():
                last_end = r - 1
                break
    blocks = []
    for i, (r, d) in enumerate(starts):
        end = starts[i + 1][0] - 1 if i + 1 < len(starts) else last_end
        blocks.append((r, end, d))
    return blocks


def drop_empty_courses(meal: dict) -> dict:
    out = {}
    for name, items in meal.items():
        if items:
            out[name] = items
    return out


def parse_bongwan():
    cells, merges = load_sheet(find_xlsx("본관"))
    data = {}
    for r0, r1, d in date_blocks(cells, merges, 70, 63):
        rows = range(r0, r1 + 1)
        data[d.strftime("%Y-%m-%d")] = {
            "아침": drop_empty_courses({"메뉴": collect_items(cells, merges, rows, [2, 3])}),
            "점심": drop_empty_courses(
                {
                    "A코스": collect_items(cells, merges, rows, [4]),
                    "B코스": collect_items(cells, merges, rows, [5]),
                }
            ),
            "저녁": drop_empty_courses(
                {
                    "일반식": collect_items(cells, merges, rows, [6]),
                    "건강다이어트식": collect_items(cells, merges, rows, [7]),
                }
            ),
            "야간": drop_empty_courses(
                {
                    "A코스": collect_items(cells, merges, rows, [8]),
                    "B코스": collect_items(cells, merges, rows, [9]),
                }
            ),
        }
    return data


def parse_am():
    cells, merges = load_sheet(find_xlsx("암병원"))
    data = {}
    for r0, r1, d in date_blocks(cells, merges, 70, 62):
        rows = range(r0, r1 + 1)
        data[d.strftime("%Y-%m-%d")] = {
            "아침": drop_empty_courses({"메뉴": collect_items(cells, merges, rows, [2, 3])}),
            "점심": drop_empty_courses(
                {
                    "A코너": collect_items(cells, merges, rows, [4]),
                    "그린테이블": collect_items(cells, merges, rows, [5]),
                }
            ),
            "저녁": drop_empty_courses(
                {
                    "A코너": collect_items(cells, merges, rows, [6]),
                    "아삭아삭샐러드": collect_items(cells, merges, rows, [7]),
                }
            ),
            "야간": drop_empty_courses(
                {
                    "A코너": collect_items(cells, merges, rows, [8]),
                    "B코너": collect_items(cells, merges, rows, [9]),
                }
            ),
        }
        if d.weekday() >= 5:
            data[d.strftime("%Y-%m-%d")] = fix_am_weekend_shift(data[d.strftime("%Y-%m-%d")])
    return data


def fix_am_weekend_shift(meals: dict) -> dict:
    """Some weeks put weekend dinner/night one column left (dinner under 그린테이블)."""
    lunch = meals.get("점심", {})
    dinner = meals.get("저녁", {})
    night = meals.get("야간", {})
    if "그린테이블" not in lunch or dinner.get("A코너"):
        return meals
    return {
        "아침": meals.get("아침", {}),
        "점심": drop_empty_courses({"A코너": lunch.get("A코너", [])}),
        "저녁": drop_empty_courses({"A코너": lunch["그린테이블"]}),
        "야간": drop_empty_courses(
            {
                "A코너": dinner.get("아삭아삭샐러드", []),
                "B코너": night.get("A코너", []),
            }
        ),
    }


def parse_ilwon():
    cells, merges = load_sheet(find_xlsx("일원"))
    data = {}
    for r0, r1, d in date_blocks(cells, merges, 50, 45):
        rows = range(r0, r1 + 1)
        data[d.strftime("%Y-%m-%d")] = {
            "점심": drop_empty_courses(
                {
                    "A코스": collect_items(cells, merges, rows, [2]),
                    "B코스": collect_items(cells, merges, rows, [3]),
                    "TOGO샐러드": collect_items(cells, merges, rows, [4]),
                    "Take-out": collect_items(cells, merges, rows, [5]),
                }
            ),
            "저녁": drop_empty_courses(
                {"일반식": collect_items(cells, merges, rows, [6])}
            ),
        }
    return data


# 본관 밀카페 — 주간 메뉴 이미지에서 정리
CAFE_BONGWAN = {
    "2026-10-05": {"안내": ["대체 휴일 휴무"]},
    "2026-10-06": {
        "빵": ["카스테라", "빠네디까사런치롤"],
        "샐러드": ["푸실리샐러드", "콥샐러드"],
        "랩·샌드위치": ["너비아니토마토랩", "게맛살샌드위치", "치킨&튜나샌드위치"],
        "컵밥": ["추억의도시락컵밥"],
    },
    "2026-10-07": {
        "빵": ["크로와상", "모카번"],
        "샐러드": ["시리얼샐러드", "수제요거트볼"],
        "랩·샌드위치": ["게맛살에그랩", "칠리맛살모닝롤샌드위치", "아이돌에그샌드위치"],
        "컵밥": ["나시고랭"],
    },
    "2026-10-08": {
        "빵": ["블루베리머핀", "바질베이글"],
        "샐러드": ["컵샐러드", "카프레제샐러드"],
        "랩·샌드위치": ["단호박리코타치즈랩", "데리야끼치킨샌드위치", "햄에그샌드위치"],
        "컵밥": ["양념치킨컵밥"],
    },
    "2026-10-09": {"안내": ["한글날 휴무"]},
}

# 암병원 밀카페 — 주간 메뉴 이미지에서 정리
CAFE_AM = {
    "2026-10-05": {"안내": ["대체 휴일 휴무"]},
    "2026-10-06": {
        "빵": ["올리브 포카치아", "치즈케이크", "크림치즈프레즐"],
        "샐러드": ["닭가슴살샐러드", "구운감자샐러드", "수제요거트"],
        "랩·샌드위치": ["랜치소시지랩", "블랙번샌드위치"],
        "컵밥": ["참치생야채컵밥"],
    },
    "2026-10-07": {
        "빵": ["통밀베이글 + 크림치즈", "얼그레이스콘", "소보로빵"],
        "샐러드": ["견과샐러드", "푸실리샐러드", "수제요거트"],
        "랩·샌드위치": ["멕시칸치킨랩", "불고기치즈버거"],
        "컵밥": ["닭갈비컵밥"],
    },
    "2026-10-08": {
        "빵": ["무화과로프", "크림치즈프레즐", "플레인카스테라"],
        "샐러드": ["구운버섯샐러드", "구운감자샐러드", "수제요거트"],
        "랩·샌드위치": ["케이준치킨시저랩", "대만식햄치즈샌드위치"],
        "컵밥": ["버터장조림컵밥"],
    },
    "2026-10-09": {"안내": ["한글날 휴무"]},
}

HOURS = {
    "본관 직원식당": {
        "아침": "06:30–08:00",
        "점심": "평일 11:00–15:00 · 주말 11:30–13:30",
        "저녁": "평일 17:30–20:00 · 주말 17:30–19:30",
        "야간": "23:30–02:30",
    },
    "암병원 직원식당": {
        "아침": "06:30–08:00",
        "점심": "평일 11:00–14:00 · 주말 11:30–13:30",
        "저녁": "평일 17:30–20:00 · 주말 17:30–19:30",
        "야간": "23:30–02:30",
    },
    "일원역캠퍼스 식당": {
        "점심": "11:00–13:30",
        "저녁": "17:30–19:00",
    },
    "본관 밀카페": {
        "운영": "07:30–15:30 (평일)",
        "빵": "오픈부터",
        "샐러드": "08:00~ · 소진 시 다음 메뉴",
        "랩·샌드위치": "1차 07:30 · 2차 11:00",
        "컵밥": "11:00~",
    },
    "암병원 밀카페": {
        "운영": "07:30–15:30 (평일)",
        "빵": "08:00~",
        "샐러드": "08:00~ · 선택 가능 · 품절·수급 시 변경",
        "랩·샌드위치": "07:30~ · 샌드위치 07:30–08:00, 12:00~ 소진 시까지",
        "컵밥": "11:00~",
    },
}


def cafe_to_meals(day_map: dict) -> dict:
    return {k: {"카페": v} for k, v in day_map.items()}


def align_ilwon(ilwon: dict, week_isos: list[str]) -> tuple[dict, bool]:
    """Keep 일원 only when it matches this week; otherwise mark weekdays pending."""
    week = set(week_isos)
    if set(ilwon) & week:
        return {k: v for k, v in ilwon.items() if k in week}, False
    pending = {}
    for iso in week_isos:
        if datetime.strptime(iso, "%Y-%m-%d").weekday() < 5:
            pending[iso] = {"점심": {"안내": ["업데이트 예정입니다."]}}
    return pending, True


def build_days(bongwan, am, ilwon):
    dates = sorted(set(bongwan) | set(am) | set(CAFE_BONGWAN) | set(CAFE_AM))
    days = []
    for iso in dates:
        dt = datetime.strptime(iso, "%Y-%m-%d")
        wd = WEEKDAYS_KO[dt.weekday()]
        days.append(
            {
                "date": iso,
                "md": f"{dt.month}/{dt.day}",
                "weekday": wd,
                "label": f"{dt.month}/{dt.day} ({wd})",
                "locations": {
                    "본관 직원식당": bongwan.get(iso, {}),
                    "본관 밀카페": cafe_to_meals(CAFE_BONGWAN).get(iso, {}),
                    "암병원 직원식당": am.get(iso, {}),
                    "암병원 밀카페": cafe_to_meals(CAFE_AM).get(iso, {}),
                    "일원역캠퍼스 식당": ilwon.get(iso, {}),
                },
            }
        )
    return days


def published_days() -> list[dict]:
    """Days from the last committed index.html (the local copy is usually deleted)."""
    try:
        html = subprocess.run(
            ["git", "show", "HEAD:index.html"],
            cwd=ROOT, capture_output=True, text=True, check=True,
        ).stdout
    except (OSError, subprocess.CalledProcessError):
        return []
    m = re.search(r'<script id="meal-data" type="application/json">(.*?)</script>', html, re.S)
    if not m:
        return []
    return json.loads(m.group(1)).get("days", [])


def week_label(monday: datetime) -> str:
    thursday = monday + timedelta(days=3)
    return f"{thursday.year}년 {thursday.month}월 {(thursday.day - 1) // 7 + 1}주"


def main():
    bongwan = parse_bongwan()
    am = parse_am()
    week_isos = sorted(set(bongwan) | set(am) | set(CAFE_BONGWAN) | set(CAFE_AM))
    ilwon, ilwon_pending = align_ilwon(parse_ilwon(), week_isos)
    days = build_days(bongwan, am, ilwon)
    monday = datetime.strptime(days[0]["date"], "%Y-%m-%d")
    # Sunday night 야간 is served past midnight, so keep the previous Sunday visible on rollover.
    prev_sunday = (monday - timedelta(days=1)).strftime("%Y-%m-%d")
    carried = next((d for d in published_days() if d["date"] == prev_sunday), None)
    if carried:
        locs = carried["locations"]
        locs["암병원 직원식당"] = fix_am_weekend_shift(locs.get("암병원 직원식당", {}))
        days.insert(0, carried)
    first = datetime.strptime(days[0]["date"], "%Y-%m-%d")
    last = datetime.strptime(days[-1]["date"], "%Y-%m-%d")
    cafe_open = sorted(
        d for d, meals in CAFE_BONGWAN.items() if list(meals.keys()) != ["안내"]
    )
    cafe_first = datetime.strptime(cafe_open[0], "%Y-%m-%d")
    cafe_last = datetime.strptime(cafe_open[-1], "%Y-%m-%d")
    holidays = sorted(
        d for d, meals in CAFE_BONGWAN.items() if list(meals.keys()) == ["안내"]
    )
    cafe_note = (
        f"본관·암병원 밀카페는 평일({cafe_first.month}/{cafe_first.day}"
        f"–{cafe_last.month}/{cafe_last.day})만 운영합니다."
    )
    if holidays:
        parts = []
        for iso in holidays:
            h = datetime.strptime(iso, "%Y-%m-%d")
            parts.append(f"{h.month}/{h.day} {CAFE_BONGWAN[iso]['안내'][0]}")
        cafe_note = cafe_note[:-1] + "(" + ", ".join(parts) + ")."
    if ilwon_pending:
        cafe_note += " 일원역캠퍼스 식단은 업데이트 예정입니다."
    payload = {
        "week": week_label(monday),
        "range": (
            f"{first.month}/{first.day} ({WEEKDAYS_KO[first.weekday()]}) – "
            f"{last.month}/{last.day} ({WEEKDAYS_KO[last.weekday()]})"
        ),
        "hours": HOURS,
        "days": days,
        "cafeNote": cafe_note,
    }
    html = render_html(payload)
    out = ROOT / "index.html"
    out.write_text(html, encoding="utf-8")
    print(f"wrote {out} ({out.stat().st_size} bytes), days={len(days)}")
    print(json.dumps({d["date"]: list(d["locations"].keys()) for d in days}, ensure_ascii=False))


def render_html(payload: dict) -> str:
    data_json = json.dumps(payload, ensure_ascii=False)
    # prevent </script> breakout
    data_json = data_json.replace("<", "\\u003c")
    return HTML_TEMPLATE.replace("__MEAL_DATA__", data_json)


HTML_TEMPLATE = r"""<!DOCTYPE html>
<html lang="ko">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover" />
  <meta http-equiv="Cache-Control" content="no-cache, no-store, must-revalidate" />
  <meta http-equiv="Pragma" content="no-cache" />
  <meta name="theme-color" content="#eaf3fc" />
  <meta name="mobile-web-app-capable" content="yes" />
  <meta name="apple-mobile-web-app-capable" content="yes" />
  <meta name="apple-mobile-web-app-title" content="SMC 식단" />
  <meta name="apple-mobile-web-app-status-bar-style" content="default" />
  <link rel="manifest" href="manifest.webmanifest" />
  <link rel="icon" href="icons/icon-192.png" />
  <link rel="apple-touch-icon" href="icons/apple-touch-icon.png" />
  <title>SMC 주간 식단</title>
  <script>
    (function () {
      var pref = null;
      try { pref = localStorage.getItem("theme"); } catch (e) {}
      var dark = pref ? pref === "dark" : window.matchMedia("(prefers-color-scheme: dark)").matches;
      document.documentElement.dataset.theme = dark ? "dark" : "light";
    })();
  </script>
  <style>
    :root {
      --bg: #eaf3fc;
      --paper: #f6faff;
      --card: #ffffff;
      --ink: #1b2a4a;
      --muted: #5f6f8c;
      --line: #d3e2f3;
      --accent: #2f5fc4;
      --green: #2c7a6c;
      --blue: #4f6a9a;
      --header: rgba(234, 243, 252, 0.93);
      --on-ink: #ffffff;
      --sel: #4c8be0;
      --on-sel: #ffffff;
      --open: #1f6b3a;
      --open-bg: #dcefe1;
      --next: #1d5fa8;
      --next-bg: #dcebfb;
      --like: #e11d48;
      --like-bg: #fde8ee;
      color-scheme: light;
    }
    :root[data-theme="dark"] {
      --bg: #0d0d0f;
      --paper: #161618;
      --card: #1d1d20;
      --ink: #ececee;
      --muted: #9a9aa2;
      --line: #2e2e33;
      --accent: #8cb8f5;
      --green: #8fd3c3;
      --blue: #aab9d6;
      --header: rgba(13, 13, 15, 0.93);
      --on-ink: #0d0d0f;
      --sel: #6fa3ec;
      --on-sel: #0d0d0f;
      --open: #9ee0b1;
      --open-bg: #1a2e22;
      --next: #9cc6ff;
      --next-bg: #1a2638;
      --like: #fb7193;
      --like-bg: #33161f;
      color-scheme: dark;
    }
    * { box-sizing: border-box; }
    html, body { margin: 0; padding: 0; background: var(--bg); color: var(--ink);
      font-family: "Apple SD Gothic Neo", "Pretendard", "Noto Sans KR", sans-serif;
      line-height: 1.45; }
    header {
      position: sticky; top: 0; z-index: 20;
      background: var(--header); backdrop-filter: blur(10px);
      border-bottom: 1px solid var(--line);
    }
    .wrap { max-width: 1180px; margin: 0 auto; padding: 18px 20px 40px; }
    header .wrap { padding-bottom: 14px; }
    .title-row { display: flex; align-items: center; justify-content: space-between; gap: 12px; }
    h1 { margin: 0 0 4px; font-size: 22px; letter-spacing: -0.03em; }
    .head-actions { display: flex; gap: 6px; flex-shrink: 0; }
    .pill-btn {
      appearance: none; border: 1px solid var(--line); background: var(--paper);
      border-radius: 999px; padding: 6px 12px; cursor: pointer; color: var(--muted);
      font: inherit; font-size: 12px; white-space: nowrap;
    }
    .pill-btn:hover { color: var(--ink); border-color: var(--ink); }
    .pill-btn[hidden] { display: none; }
    .sub { color: var(--muted); font-size: 13px; }
    .dates { display: flex; gap: 8px; flex-wrap: wrap; margin-top: 14px; }
    .dates button {
      appearance: none; border: 1px solid var(--line); background: var(--paper);
      border-radius: 999px; padding: 8px 14px; cursor: pointer; color: var(--ink);
      font: inherit; font-size: 14px;
    }
    .dates button[aria-selected="true"] {
      background: var(--sel); color: var(--on-sel); border-color: var(--sel);
    }
    .dates button.is-today:not([aria-selected="true"]) {
      border-color: var(--accent); color: var(--accent);
    }
    .views {
      margin-top: 12px; font-size: 12px; color: var(--muted);
      display: flex; flex-wrap: wrap; gap: 8px 14px; align-items: baseline;
    }
    .views strong { color: var(--ink); font-weight: 700; }
    .views .v-day { font-variant-numeric: tabular-nums; }
    .loc-nav { display: flex; gap: 8px; flex-wrap: wrap; margin: 16px 0 8px; }
    .loc-nav a {
      color: var(--muted); text-decoration: none; font-size: 13px;
      border-bottom: 1px solid transparent;
    }
    .loc-nav a:hover { color: var(--ink); border-color: var(--ink); }
    .loc-nav a.is-live { color: var(--open); font-weight: 700; }
    section.place {
      background: var(--paper); border: 1px solid var(--line);
      border-radius: 16px; padding: 20px 20px 16px; margin: 16px 0 0;
    }
    .slot-head { display: flex; align-items: center; gap: 8px; margin-bottom: 4px; }
    section.place h2 { margin: 0; font-size: 18px; letter-spacing: -0.02em; }
    .share-btn { margin-left: auto; padding: 4px 10px; }
    .live {
      display: inline-block; font-size: 11px; font-weight: 700; line-height: 1;
      padding: 4px 8px; border-radius: 999px; white-space: nowrap;
    }
    .live-open { background: var(--open-bg); color: var(--open); }
    .live-next { background: var(--next-bg); color: var(--next); }
    .when .live { margin-left: 6px; padding: 2px 6px; font-size: 10px; }
    .hours-line { color: var(--muted); font-size: 12px; margin-bottom: 14px; }
    .slot-head + .meals, .slot-head + .cafe-grid, .slot-head + .empty { margin-top: 10px; }
    .meals {
      display: grid; gap: 10px;
      grid-template-columns: repeat(auto-fit, minmax(150px, 1fr));
    }
    .meal {
      border: 1px solid var(--line); border-radius: 12px; padding: 12px 12px 8px;
      background: var(--card); display: flex; flex-direction: column;
    }
    .meal .like-row { margin-top: auto; }
    .meal h3 { margin: 0 0 2px; font-size: 13px; color: var(--accent); }
    .meal .when { font-size: 11px; color: var(--muted); margin-bottom: 8px; }
    .course { margin-bottom: 10px; }
    .course h4 { margin: 0 0 4px; font-size: 12px; color: var(--blue); font-weight: 700; }
    .course ul, .cafe-cell ul { margin: 0; padding: 0 0 0 16px; }
    .course li, .cafe-cell li { font-size: 13.5px; margin: 2px 0; }
    li.theme { list-style: none; margin-left: -16px; font-weight: 700; color: var(--green); }
    .tag {
      display: inline-block; font-size: 10px; font-weight: 700; line-height: 15px;
      padding: 0 5px; border-radius: 5px; margin-right: 4px; vertical-align: 1px; color: #fff;
    }
    .tag-new { background: #e11d48; }
    .tag-hot { background: #ea580c; }
    .tag-encore { background: #7c3aed; }
    .tag-season { background: #16a34a; }
    .tag-collab { background: #2563eb; }
    .cafe-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 6px 10px; }
    .cafe-head {
      margin: 0; font-size: 13px; color: var(--accent);
      border-bottom: 1px solid var(--line); padding-bottom: 6px;
    }
    .cafe-cat { grid-column: 1 / -1; margin: 8px 0 0; font-size: 12px; color: var(--blue); font-weight: 700; }
    .cafe-cell { border: 1px solid var(--line); border-radius: 10px; padding: 8px 10px; background: var(--card); }
    .cafe-cell .when { font-size: 11px; color: var(--muted); margin-bottom: 4px; }
    .cafe-cell .empty { padding: 0; }
    .empty { color: var(--muted); font-size: 13px; padding: 8px 0; }
    .like-row {
      display: flex; justify-content: flex-end;
      border-top: 1px dashed var(--line); padding-top: 8px; margin: 2px 0 4px;
    }
    .cafe-like { border-top: 0; padding-top: 2px; }
    .section-like { border-top: 0; padding-top: 0; margin-top: 10px; }
    .like-btn {
      appearance: none; display: inline-flex; align-items: center; gap: 4px;
      border: 1px solid var(--line); background: transparent; color: var(--muted);
      border-radius: 999px; padding: 3px 10px; font: inherit; font-size: 12px; cursor: pointer;
    }
    .like-btn:hover { color: var(--like); border-color: var(--like); }
    .like-btn.liked { color: var(--like); border-color: var(--like); background: var(--like-bg); font-weight: 700; }
    .like-btn .heart { font-size: 13px; line-height: 1; }
    footer { color: var(--muted); font-size: 12px; margin-top: 28px; }
    .toast {
      position: fixed; left: 50%; bottom: calc(24px + env(safe-area-inset-bottom));
      transform: translate(-50%, 20px); opacity: 0; pointer-events: none;
      background: var(--ink); color: var(--on-ink); font-size: 13px;
      padding: 10px 16px; border-radius: 999px; max-width: calc(100% - 32px);
      transition: opacity 0.2s, transform 0.2s; z-index: 30; text-align: center;
    }
    .toast.show { opacity: 1; transform: translate(-50%, 0); }
    @media (max-width: 600px) {
      h1 { font-size: 20px; }
      .dates {
        flex-wrap: nowrap; overflow-x: auto; scrollbar-width: none;
        margin-left: -20px; margin-right: -20px; padding: 0 20px;
      }
      .dates::-webkit-scrollbar { display: none; }
      .dates button { flex: 0 0 auto; }
      section.place { padding: 14px 12px 12px; }
      .meal { padding: 10px 10px 6px; }
      .course li, .cafe-cell li { font-size: 13px; }
      .course ul, .cafe-cell ul { padding-left: 14px; }
    }
    @media print {
      header { position: static; background: #fff; }
      .dates, .loc-nav, .views, .head-actions, .share-btn, .live, .like-row { display: none; }
      section.place { break-inside: avoid; }
    }
  </style>
</head>
<body>
  <header>
    <div class="wrap">
      <div class="title-row">
        <h1>삼성서울병원 주간 식단</h1>
        <div class="head-actions">
          <button type="button" class="pill-btn" id="install-btn" hidden>앱으로 추가</button>
          <button type="button" class="pill-btn" id="theme-btn">다크 모드</button>
        </div>
      </div>
      <div class="sub" id="weekline"></div>
      <div class="dates" id="dates" role="tablist" aria-label="날짜 선택"></div>
      <div class="views" id="views" hidden>조회수를 불러오는 중</div>
      <nav class="loc-nav" id="locnav"></nav>
    </div>
  </header>
  <main class="wrap" id="main"></main>
  <div class="toast" id="toast" role="status" aria-live="polite"></div>
  <script id="meal-data" type="application/json">__MEAL_DATA__</script>
  <script>
    const DATA = JSON.parse(document.getElementById("meal-data").textContent);
    const SITE_URL = "https://iamguno.github.io/smc-weekly-menu/";
    const STAFF = ["본관 직원식당", "암병원 직원식당"];
    const CAFES = ["본관 밀카페", "암병원 밀카페"];
    const ILWON = "일원역캠퍼스 식당";
    const SHORT = {
      "본관 직원식당": "본관", "암병원 직원식당": "암병원",
      "본관 밀카페": "본관", "암병원 밀카페": "암병원",
    };
    const MEAL_ORDER = ["아침", "점심", "저녁", "야간"];
    const CAFE_CATS = ["빵", "샐러드", "랩·샌드위치", "컵밥"];
    const SLOTS = [...MEAL_ORDER, "밀카페", "일원역캠퍼스"];
    const LOC_ID = {
      "본관 직원식당": "bg", "암병원 직원식당": "am",
      "본관 밀카페": "bc", "암병원 밀카페": "ac", "일원역캠퍼스 식당": "iw",
    };
    const MEAL_ID = { "아침": "b", "점심": "l", "저녁": "d", "야간": "n", "카페": "c", "하루": "all" };
    const ON_SITE = location.hostname === "iamguno.github.io";
    const TAGS = [
      [/\[인기메뉴\]/, "인기", "hot"],
      [/\[앵콜메뉴\]/, "앵콜", "encore"],
      [/\[시즌메뉴\]/, "시즌", "season"],
      [/\[브랜드콜라보\]|^브랜드콜라보\s*-\s*/, "콜라보", "collab"],
      [/\(NEW\)|^NEW(?=[가-힣\s(])/, "NEW", "new"],
    ];

    function isoOf(d) {
      return [d.getFullYear(), String(d.getMonth() + 1).padStart(2, "0"), String(d.getDate()).padStart(2, "0")].join("-");
    }
    const loadedAt = new Date();
    const todayIso = isoOf(loadedAt);
    const yesterdayIso = isoOf(new Date(loadedAt.getFullYear(), loadedAt.getMonth(), loadedAt.getDate() - 1));

    const weekline = document.getElementById("weekline");
    weekline.textContent = DATA.week + " · " + DATA.range;

    const datesEl = document.getElementById("dates");
    const viewsEl = document.getElementById("views");
    const viewsByDay = {};
    const COUNT_NS = "iamguno.github.io";
    const COUNT_BASE = "https://abacus.jsn.cam";
    let selected = DATA.days.some(d => d.date === todayIso) ? todayIso : DATA.days[0].date;

    function dayOf(iso) {
      return DATA.days.find(d => d.date === iso);
    }

    // ---- 운영 시간 ----
    function hasMenu(day, loc, meal) {
      const m = ((day.locations[loc] || {})[meal]) || {};
      return Object.keys(m).some(k => k !== "안내" && m[k] && m[k].length);
    }

    function isRestDay(day) {
      const wd = new Date(day.date + "T00:00:00").getDay();
      if (wd === 0 || wd === 6) return true;
      return CAFES.some(loc => {
        const c = (day.locations[loc] || {})["카페"];
        return c && c["안내"];
      });
    }

    function ranges(day, text) {
      if (!text) return [];
      const rest = isRestDay(day);
      if (/\(평일\)/.test(text) && rest) return [];
      const base = new Date(day.date + "T00:00:00").getTime();
      const out = [];
      text.split("·").forEach(part => {
        part = part.trim();
        if (part.startsWith("평일") && rest) return;
        if (part.startsWith("주말") && !rest) return;
        const m = part.match(/(\d{1,2}):(\d{2})\s*[–-]\s*(\d{1,2}):(\d{2})/);
        if (!m) return;
        const s = Number(m[1]) * 60 + Number(m[2]);
        let e = Number(m[3]) * 60 + Number(m[4]);
        if (e <= s) e += 1440;
        out.push([base + s * 60000, base + e * 60000]);
      });
      return out;
    }

    const openAt = (rs, t) => rs.some(([s, e]) => s <= t && t < e);

    function slotRanges(day, slot) {
      if (MEAL_ORDER.includes(slot)) {
        return STAFF.flatMap(loc => hasMenu(day, loc, slot) ? ranges(day, DATA.hours[loc][slot]) : []);
      }
      if (slot === "밀카페") {
        return CAFES.some(loc => hasMenu(day, loc, "카페")) ? ranges(day, DATA.hours[CAFES[0]]["운영"]) : [];
      }
      return Object.keys(DATA.hours[ILWON])
        .flatMap(meal => hasMenu(day, ILWON, meal) ? ranges(day, DATA.hours[ILWON][meal]) : []);
    }

    function liveStatus(day, t) {
      const open = new Set();
      let next = null;
      let nextStart = Infinity;
      SLOTS.forEach(slot => {
        const rs = slotRanges(day, slot);
        if (openAt(rs, t)) open.add(slot);
        if (!MEAL_ORDER.includes(slot)) return;
        rs.forEach(([s]) => {
          if (s > t && s < nextStart) { nextStart = s; next = slot; }
        });
      });
      const mealOpen = MEAL_ORDER.some(s => open.has(s));
      return { open, next: mealOpen ? null : next };
    }

    function initialFocus() {
      const t = Date.now();
      const yesterday = dayOf(yesterdayIso);
      if (yesterday && openAt(slotRanges(yesterday, "야간"), t)) return [yesterdayIso, "야간"];
      const today = dayOf(todayIso);
      if (!today) return [DATA.days[0].date, null];
      const st = liveStatus(today, t);
      return [todayIso, MEAL_ORDER.find(s => st.open.has(s)) || st.next];
    }

    function liveBadge(text, kind) {
      const el = document.createElement("span");
      el.className = "live live-" + kind;
      el.textContent = text;
      return el;
    }

    function refreshLive() {
      document.querySelectorAll(".live").forEach(el => el.remove());
      document.querySelectorAll("#locnav a").forEach(a => a.classList.remove("is-live"));
      if (selected !== todayIso && selected !== yesterdayIso) return;
      const day = dayOf(selected);
      const t = Date.now();
      const st = liveStatus(day, t);
      SLOTS.forEach(slot => {
        const head = document.querySelector('[data-slot="' + slot + '"] .slot-head h2');
        if (!head) return;
        if (st.open.has(slot)) {
          head.after(liveBadge("운영 중", "open"));
          const a = document.querySelector('#locnav a[data-slot="' + slot + '"]');
          if (a) a.classList.add("is-live");
        } else if (st.next === slot) {
          head.after(liveBadge("다음 식사", "next"));
        }
      });
      document.querySelectorAll(".meal[data-loc]").forEach(card => {
        const loc = card.dataset.loc;
        const meal = card.dataset.meal;
        if (hasMenu(day, loc, meal) && openAt(ranges(day, DATA.hours[loc][meal]), t)) {
          card.querySelector(".when").appendChild(liveBadge("운영 중", "open"));
        }
      });
    }

    // ---- 화면 ----
    function renderDateButtons() {
      datesEl.innerHTML = "";
      DATA.days.forEach(d => {
        const b = document.createElement("button");
        b.type = "button";
        b.textContent = d.label;
        b.setAttribute("aria-selected", d.date === selected ? "true" : "false");
        if (d.date === todayIso) {
          b.classList.add("is-today");
          if (d.date === selected) b.textContent = d.label + " · 오늘";
        }
        b.addEventListener("click", () => {
          selected = d.date;
          render();
        });
        datesEl.appendChild(b);
      });
      const sel = datesEl.querySelector('[aria-selected="true"]');
      if (sel && datesEl.scrollWidth > datesEl.clientWidth) {
        datesEl.scrollLeft = sel.offsetLeft - datesEl.offsetLeft - (datesEl.clientWidth - sel.offsetWidth) / 2;
      }
    }

    function itemList(items) {
      const ul = document.createElement("ul");
      items.forEach(it => {
        const li = document.createElement("li");
        const isTheme = /^<<.+>>$/.test(it) || (/^\*.+\*$/.test(it) && it.length < 40);
        if (isTheme) {
          li.className = "theme";
          li.textContent = it.replace(/^<<|>>$/g, "").replace(/^\*|\*$/g, "");
        } else {
          let text = it;
          TAGS.forEach(([re, label, cls]) => {
            if (!re.test(text)) return;
            text = text.replace(re, "");
            const tag = document.createElement("span");
            tag.className = "tag tag-" + cls;
            tag.textContent = label;
            li.appendChild(tag);
          });
          text = text.trim();
          if (!text && ul.lastChild && !ul.lastChild.classList.contains("theme")) {
            ul.lastChild.prepend(...li.childNodes);
            return;
          }
          li.appendChild(document.createTextNode(text));
        }
        ul.appendChild(li);
      });
      return ul;
    }

    function emptyNote(text) {
      const el = document.createElement("div");
      el.className = "empty";
      el.textContent = text;
      return el;
    }

    function slotSection(slot, hoursText) {
      const sec = document.createElement("section");
      sec.className = "place";
      sec.id = "slot-" + slot;
      sec.dataset.slot = slot;
      const head = document.createElement("div");
      head.className = "slot-head";
      const h2 = document.createElement("h2");
      h2.textContent = slot;
      head.appendChild(h2);
      const share = document.createElement("button");
      share.type = "button";
      share.className = "pill-btn share-btn";
      share.textContent = "공유";
      share.addEventListener("click", () => shareSlot(dayOf(selected), slot));
      head.appendChild(share);
      sec.appendChild(head);
      if (hoursText) {
        const p = document.createElement("div");
        p.className = "hours-line";
        p.textContent = hoursText;
        sec.appendChild(p);
      }
      return sec;
    }

    function courseCard(title, hours, courses, loc, meal, key) {
      const card = document.createElement("div");
      card.className = "meal";
      if (loc) {
        card.dataset.loc = loc;
        card.dataset.meal = meal;
      }
      const h3 = document.createElement("h3");
      h3.textContent = title;
      card.appendChild(h3);
      const when = document.createElement("div");
      when.className = "when";
      when.textContent = hours;
      card.appendChild(when);
      const names = Object.keys(courses).filter(n => courses[n] && courses[n].length);
      if (!names.length) card.appendChild(emptyNote("이 날은 메뉴가 없습니다."));
      names.forEach(name => {
        const c = document.createElement("div");
        c.className = "course";
        if (!(names.length === 1 && (name === "메뉴" || name === "안내"))) {
          const h4 = document.createElement("h4");
          h4.textContent = name;
          c.appendChild(h4);
        }
        c.appendChild(itemList(courses[name]));
        card.appendChild(c);
      });
      if (key && names.length) card.appendChild(likeRow(key, "like-row"));
      return card;
    }

    function renderMealSlot(day, meal) {
      const sec = slotSection(meal, "");
      const grid = document.createElement("div");
      grid.className = "meals";
      STAFF.filter(loc => DATA.hours[loc][meal]).forEach(loc => {
        const courses = (day.locations[loc] || {})[meal] || {};
        const key = hasMenu(day, loc, meal) ? likeKey(day.date, loc, meal) : null;
        grid.appendChild(courseCard(SHORT[loc], DATA.hours[loc][meal], courses, loc, meal, key));
      });
      sec.appendChild(grid);
      return sec;
    }

    function renderIlwonSlot(day) {
      const sec = slotSection("일원역캠퍼스", "");
      const meals = day.locations[ILWON] || {};
      if (!Object.values(meals).some(m => Object.keys(m).length)) {
        sec.appendChild(emptyNote("이 날은 운영하지 않습니다."));
        return sec;
      }
      const grid = document.createElement("div");
      grid.className = "meals";
      const mealNames = Object.keys(DATA.hours[ILWON]);
      mealNames.forEach(meal => {
        grid.appendChild(courseCard(meal, DATA.hours[ILWON][meal], meals[meal] || {}, ILWON, meal));
      });
      sec.appendChild(grid);
      if (mealNames.some(meal => hasMenu(day, ILWON, meal))) {
        sec.appendChild(likeRow(likeKey(day.date, ILWON, "하루"), "like-row section-like"));
      }
      return sec;
    }

    function renderCafeSlot(day) {
      const sec = slotSection("밀카페", DATA.hours[CAFES[0]]["운영"]);
      const menus = CAFES.map(loc => ((day.locations[loc] || {})["카페"]) || {});
      const cats = ["안내", ...CAFE_CATS].filter(cat => menus.some(m => m[cat] && m[cat].length));
      if (!cats.length) {
        sec.appendChild(emptyNote("이 날은 밀카페를 운영하지 않습니다."));
        return sec;
      }
      const grid = document.createElement("div");
      grid.className = "cafe-grid";
      CAFES.forEach(loc => {
        const h3 = document.createElement("h3");
        h3.className = "cafe-head";
        h3.textContent = SHORT[loc];
        grid.appendChild(h3);
      });
      cats.forEach(cat => {
        if (cat !== "안내") {
          const label = document.createElement("h4");
          label.className = "cafe-cat";
          label.textContent = cat;
          grid.appendChild(label);
        }
        CAFES.forEach((loc, i) => {
          const cell = document.createElement("div");
          cell.className = "cafe-cell";
          const hours = cat !== "안내" && DATA.hours[loc][cat];
          if (hours) {
            const when = document.createElement("div");
            when.className = "when";
            when.textContent = hours;
            cell.appendChild(when);
          }
          const items = menus[i][cat];
          cell.appendChild(items && items.length ? itemList(items) : emptyNote("—"));
          grid.appendChild(cell);
        });
      });
      const open = CAFES.filter(loc => hasMenu(day, loc, "카페"));
      if (open.length) {
        CAFES.forEach(loc => {
          grid.appendChild(open.includes(loc)
            ? likeRow(likeKey(day.date, loc, "카페"), "like-row cafe-like")
            : document.createElement("div"));
        });
      }
      sec.appendChild(grid);
      return sec;
    }

    function scrollToSlot(slot, behavior) {
      const target = document.getElementById("slot-" + slot);
      if (!target) return;
      const offset = document.querySelector("header").offsetHeight + 8;
      window.scrollTo({ top: target.getBoundingClientRect().top + window.scrollY - offset, behavior });
    }

    function render() {
      renderDateButtons();
      const day = dayOf(selected);
      const main = document.getElementById("main");
      main.innerHTML = "";
      const locnav = document.getElementById("locnav");
      locnav.innerHTML = "";

      SLOTS.forEach(slot => {
        const a = document.createElement("a");
        a.href = "#slot-" + slot;
        a.dataset.slot = slot;
        a.textContent = slot;
        a.addEventListener("click", e => {
          e.preventDefault();
          scrollToSlot(slot, "smooth");
        });
        locnav.appendChild(a);
        main.appendChild(
          slot === "밀카페" ? renderCafeSlot(day)
          : slot === "일원역캠퍼스" ? renderIlwonSlot(day)
          : renderMealSlot(day, slot)
        );
      });

      const foot = document.createElement("footer");
      foot.textContent = DATA.cafeNote + " 메뉴·원산지는 당일 사정에 따라 바뀔 수 있습니다.";
      main.appendChild(foot);
      refreshLive();
      watchLikes();
    }

    // ---- 공유 ----
    const toastEl = document.getElementById("toast");
    let toastTimer = null;
    function toast(text, ms) {
      toastEl.textContent = text;
      toastEl.classList.add("show");
      clearTimeout(toastTimer);
      toastTimer = setTimeout(() => toastEl.classList.remove("show"), ms || 2200);
    }

    function cleanItem(it) {
      return it
        .replace(/^<<|>>$/g, "").replace(/^\*+|\*+$/g, "")
        .replace(/\[[^\]]*\]/g, "").replace(/\(NEW\)/g, "").replace(/^NEW(?=[가-힣\s(])/, "")
        .replace(/\([^()]*:[^()]*\)/g, "")
        .replace(/\s+/g, " ").trim();
    }

    function joinItems(items) {
      return items.map(cleanItem).filter(Boolean).join(", ");
    }

    function courseLines(courses) {
      const names = Object.keys(courses).filter(n => courses[n] && courses[n].length);
      return names.map(n => (names.length === 1 && (n === "메뉴" || n === "안내") ? "" : n + ": ") + joinItems(courses[n]));
    }

    function shareText(day, slot) {
      const lines = ["[SMC 식단] " + day.label + " " + slot];
      if (MEAL_ORDER.includes(slot)) {
        STAFF.forEach(loc => {
          const body = courseLines((day.locations[loc] || {})[slot] || {});
          if (body.length) lines.push("", "■ " + SHORT[loc], ...body);
        });
      } else if (slot === "밀카페") {
        const menus = CAFES.map(loc => ((day.locations[loc] || {})["카페"]) || {});
        ["안내", ...CAFE_CATS].forEach(cat => {
          const parts = CAFES
            .map((loc, i) => (menus[i][cat] && menus[i][cat].length ? SHORT[loc] + ": " + joinItems(menus[i][cat]) : null))
            .filter(Boolean);
          if (parts.length) lines.push("■ " + cat + " — " + parts.join(" / "));
        });
      } else {
        const meals = day.locations[ILWON] || {};
        Object.keys(DATA.hours[ILWON]).forEach(meal => {
          const body = courseLines(meals[meal] || {});
          if (body.length) lines.push("", "■ " + meal, ...body);
        });
      }
      if (lines.length === 1) lines.push("", "이 날은 메뉴가 없습니다.");
      lines.push("", SITE_URL);
      return lines.join("\n");
    }

    async function copyText(text) {
      try {
        await navigator.clipboard.writeText(text);
      } catch (err) {
        const ta = document.createElement("textarea");
        ta.value = text;
        ta.setAttribute("readonly", "");
        ta.style.position = "fixed";
        ta.style.opacity = "0";
        document.body.appendChild(ta);
        ta.select();
        document.execCommand("copy");
        ta.remove();
      }
    }

    async function shareSlot(day, slot) {
      const text = shareText(day, slot);
      if (navigator.share && window.matchMedia("(pointer: coarse)").matches) {
        try {
          await navigator.share({ text });
          return;
        } catch (err) {
          if (err && err.name === "AbortError") return;
        }
      }
      await copyText(text);
      toast("복사했어요. 카톡에 붙여넣기 하세요.");
    }

    // ---- 다크 모드 ----
    const themeBtn = document.getElementById("theme-btn");
    const themeMeta = document.querySelector('meta[name="theme-color"]');
    function applyTheme(theme, save) {
      document.documentElement.dataset.theme = theme;
      themeBtn.textContent = theme === "dark" ? "라이트 모드" : "다크 모드";
      themeMeta.content = theme === "dark" ? "#0d0d0f" : "#eaf3fc";
      if (save) {
        try { localStorage.setItem("theme", theme); } catch (err) {}
      }
    }
    applyTheme(document.documentElement.dataset.theme, false);
    themeBtn.addEventListener("click", () => {
      applyTheme(document.documentElement.dataset.theme === "dark" ? "light" : "dark", true);
    });
    window.matchMedia("(prefers-color-scheme: dark)").addEventListener("change", e => {
      let pref = null;
      try { pref = localStorage.getItem("theme"); } catch (err) {}
      if (!pref) applyTheme(e.matches ? "dark" : "light", false);
    });

    // ---- 홈 화면 앱 ----
    const installBtn = document.getElementById("install-btn");
    const standalone = window.matchMedia("(display-mode: standalone)").matches || navigator.standalone === true;
    const isIOS = /iPhone|iPad|iPod/.test(navigator.userAgent);
    let installEvent = null;
    if (isIOS && !standalone) installBtn.hidden = false;
    window.addEventListener("beforeinstallprompt", e => {
      e.preventDefault();
      installEvent = e;
      installBtn.hidden = false;
    });
    installBtn.addEventListener("click", async () => {
      if (installEvent) {
        installEvent.prompt();
        await installEvent.userChoice;
        installEvent = null;
        installBtn.hidden = true;
      } else if (isIOS) {
        toast("Safari 아래쪽 공유 버튼 → '홈 화면에 추가'를 누르세요.", 4500);
      }
    });
    window.addEventListener("appinstalled", () => { installBtn.hidden = true; });
    if ("serviceWorker" in navigator && location.protocol === "https:") {
      navigator.serviceWorker.register("sw.js").catch(() => {});
    }

    // ---- 조회수 ----
    async function countGet(key) {
      const res = await fetch(COUNT_BASE + "/get/" + COUNT_NS + "/" + encodeURIComponent(key));
      if (res.status === 404) return 0;
      if (!res.ok) throw new Error("count get");
      const data = await res.json();
      return Number(data.value) || 0;
    }

    async function countHit(key) {
      const res = await fetch(COUNT_BASE + "/hit/" + COUNT_NS + "/" + encodeURIComponent(key));
      if (!res.ok) throw new Error("count hit");
      const data = await res.json();
      return Number(data.value) || 0;
    }

    async function loadViews() {
      const onSite = location.hostname === "iamguno.github.io";
      try {
        for (const d of DATA.days) {
          viewsByDay[d.date] = d.date === todayIso && onSite
            ? await countHit("menu-" + d.date)
            : await countGet("menu-" + d.date);
        }
        const todayCount = viewsByDay[todayIso] || 0;
        const weekCount = DATA.days.reduce((s, day) => s + (viewsByDay[day.date] || 0), 0);
        viewsEl.hidden = false;
        viewsEl.innerHTML = "";
        const summary = document.createElement("div");
        summary.innerHTML = "조회수 · 오늘 <strong>" + todayCount + "</strong>회 · 이번 주 <strong>" + weekCount + "</strong>회";
        viewsEl.appendChild(summary);
        DATA.days.forEach(d => {
          const span = document.createElement("span");
          span.className = "v-day";
          span.textContent = d.md + " " + (viewsByDay[d.date] || 0) + "회";
          viewsEl.appendChild(span);
        });
      } catch (err) {
        viewsEl.hidden = true;
      }
    }

    // ---- 좋아요 ----
    // Abacus allows ~30 requests per 10s per IP (shared on hospital Wi-Fi), so counts load per visible section.
    const likeCounts = {};
    const likeLoading = new Set();
    let likeObserver = null;

    function likeKey(date, loc, meal) {
      return (ON_SITE ? "" : "dev-") + "like-" + date + "-" + LOC_ID[loc] + "-" + MEAL_ID[meal];
    }

    function isLiked(key) {
      try { return localStorage.getItem("liked:" + key) === "1"; } catch (err) { return false; }
    }

    function likeRow(key, cls) {
      const row = document.createElement("div");
      row.className = cls;
      const btn = document.createElement("button");
      btn.type = "button";
      btn.className = "like-btn";
      btn.dataset.likeKey = key;
      btn.addEventListener("click", () => toggleLike(key));
      row.appendChild(btn);
      paintLikeButton(btn);
      return row;
    }

    function paintLikeButton(btn) {
      const key = btn.dataset.likeKey;
      const liked = isLiked(key);
      const n = likeCounts[key];
      btn.classList.toggle("liked", liked);
      btn.setAttribute("aria-pressed", liked ? "true" : "false");
      btn.innerHTML = "";
      const heart = document.createElement("span");
      heart.className = "heart";
      heart.textContent = liked ? "♥" : "♡";
      btn.appendChild(heart);
      btn.appendChild(document.createTextNode(typeof n === "number" && n > 0 ? String(n) : "좋아요"));
    }

    function paintLikes() {
      document.querySelectorAll(".like-btn").forEach(paintLikeButton);
    }

    async function loadLikes(sec) {
      const keys = [...sec.querySelectorAll(".like-btn")]
        .map(b => b.dataset.likeKey)
        .filter(k => !(k in likeCounts) && !likeLoading.has(k));
      if (!keys.length) return;
      keys.forEach(k => likeLoading.add(k));
      try {
        for (const k of keys) likeCounts[k] = await countGet(k);
      } catch (err) {
      } finally {
        keys.forEach(k => likeLoading.delete(k));
      }
      paintLikes();
    }

    function watchLikes() {
      if (likeObserver) likeObserver.disconnect();
      const secs = [...document.querySelectorAll("section.place")].filter(s => s.querySelector(".like-btn"));
      if (!("IntersectionObserver" in window)) {
        secs.forEach(loadLikes);
        return;
      }
      likeObserver = new IntersectionObserver(entries => {
        entries.forEach(e => {
          if (!e.isIntersecting) return;
          likeObserver.unobserve(e.target);
          loadLikes(e.target);
        });
      }, { rootMargin: "100px 0px" });
      secs.forEach(s => likeObserver.observe(s));
    }

    async function toggleLike(key) {
      if (isLiked(key)) {
        toast("이미 좋아요를 눌렀어요.");
        return;
      }
      const before = likeCounts[key];
      try { localStorage.setItem("liked:" + key, "1"); } catch (err) {}
      likeCounts[key] = (typeof before === "number" ? before : 0) + 1;
      paintLikes();
      try {
        likeCounts[key] = await countHit(key);
      } catch (err) {
        try { localStorage.removeItem("liked:" + key); } catch (e) {}
        if (typeof before === "number") likeCounts[key] = before;
        else delete likeCounts[key];
        toast("잠시 후 다시 눌러 주세요.");
      }
      paintLikes();
    }

    const [focusDate, focusSlot] = initialFocus();
    selected = focusDate;
    render();
    if (focusSlot && !location.hash) scrollToSlot(focusSlot, "auto");
    setInterval(refreshLive, 60000);
    document.addEventListener("visibilitychange", () => {
      if (document.visibilityState === "visible") refreshLive();
    });
    loadViews();
  </script>
</body>
</html>
"""


if __name__ == "__main__":
    main()
