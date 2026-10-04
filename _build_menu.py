#!/usr/bin/env python3
"""Build index.html from weekly cafeteria spreadsheets."""
from __future__ import annotations

import json
import re
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
    return data


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
    "2026-09-28": {
        "빵": ["초코칩머핀", "크림치즈프레즐", "블루베리베이글 + 크림치즈"],
        "샐러드": ["만다린샐러드", "컵샐러드"],
        "랩·샌드위치": ["햄치즈베이글샌드위치", "파프리카불고기랩"],
        "컵밥": ["치킨가라아게덮밥"],
    },
    "2026-09-29": {
        "빵": ["크로와상", "모카번", "빠네디까사런치롤"],
        "샐러드": ["시리얼샐러드", "카프레제샐러드"],
        "랩·샌드위치": ["크랜베리리코타샌드위치", "너비아니토마토랩"],
        "컵밥": ["제육덮밥"],
    },
    "2026-09-30": {
        "빵": ["단백쿠키(그랜드아몬드)", "아몬드머핀", "먹물치아바타"],
        "샐러드": ["견과샐러드", "수제요거트볼"],
        "랩·샌드위치": ["치킨리코타모닝롤샌드위치", "하와이안포케랩"],
        "컵밥": ["스팸김치볶음밥"],
    },
    "2026-10-01": {
        "빵": ["블루베리머핀", "롤롤페스츄리", "바질베이글"],
        "샐러드": ["푸실리샐러드", "컵샐러드"],
        "랩·샌드위치": ["할라피뇨볼로냐샌드위치", "단호박리코타치즈랩"],
        "컵밥": ["불닭곤약오븐밥"],
    },
    "2026-10-02": {
        "빵": ["하드소금빵", "소보로빵", "어니언베이글 + 크림치즈"],
        "샐러드": ["치즈볼샐러드", "수제요거트볼"],
        "랩·샌드위치": ["햄치즈샌드위치", "스파이시치킨랩"],
        "컵밥": ["새우락소스오므라이스"],
    },
}

# 암병원 밀카페 — 주간 메뉴 이미지에서 정리
CAFE_AM = {
    "2026-09-28": {
        "빵": ["바게트감자뉴", "크림치즈프레즐", "녹차카스테라"],
        "샐러드": ["보코치니샐러드", "푸실리샐러드", "수제요거트"],
        "랩·샌드위치": ["랜치소시지랩", "폴드포크샌드위치"],
        "컵밥": ["스팸김치컵밥"],
    },
    "2026-09-29": {
        "빵": ["멀티그레인치아바타", "피칸파이", "고구마케이크"],
        "샐러드": ["닭가슴살샐러드", "구운감자샐러드", "수제요거트"],
        "랩·샌드위치": ["멕시칸치킨랩", "햄치즈크라상샌드위치"],
        "컵밥": ["치킨마요컵밥"],
    },
    "2026-09-30": {
        "빵": ["호밀빵", "갈릭치즈버거볼", "플레인카스테라"],
        "샐러드": ["견과샐러드", "푸실리샐러드", "수제요거트"],
        "랩·샌드위치": ["비프치폴레샐러드랩", "대만식햄치즈샌드위치"],
        "컵밥": ["떡갈비컵밥"],
    },
    "2026-10-01": {
        "빵": ["무화과로프", "크림치즈프레즐", "크로와상"],
        "샐러드": ["구운버섯샐러드", "구운감자샐러드", "수제요거트"],
        "랩·샌드위치": ["케이준치킨랩", "로제치킨샌드위치"],
        "컵밥": ["소시지오므라이스컵밥"],
    },
    "2026-10-02": {
        "빵": ["바질베이글", "얼그레이스콘", "소보로빵"],
        "샐러드": ["닭가슴살샐러드", "보코치니샐러드", "수제요거트"],
        "랩·샌드위치": ["간장불고기샐러드랩", "BELT샌드위치"],
        "컵밥": ["불고기컵밥"],
    },
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


def main():
    bongwan = parse_bongwan()
    am = parse_am()
    week_isos = sorted(set(bongwan) | set(am) | set(CAFE_BONGWAN) | set(CAFE_AM))
    ilwon, ilwon_pending = align_ilwon(parse_ilwon(), week_isos)
    days = build_days(bongwan, am, ilwon)
    first = datetime.strptime(days[0]["date"], "%Y-%m-%d")
    last = datetime.strptime(days[-1]["date"], "%Y-%m-%d")
    cafe_open = sorted(
        d for d, meals in CAFE_BONGWAN.items() if list(meals.keys()) != ["안내"]
    )
    cafe_first = datetime.strptime(cafe_open[0], "%Y-%m-%d")
    cafe_last = datetime.strptime(cafe_open[-1], "%Y-%m-%d")
    cafe_note = (
        f"본관·암병원 밀카페는 평일({cafe_first.month}/{cafe_first.day}"
        f"–{cafe_last.month}/{cafe_last.day})만 운영합니다."
    )
    holidays = sorted(
        d for d, meals in CAFE_BONGWAN.items() if list(meals.keys()) == ["안내"]
    )
    if holidays:
        h0 = datetime.strptime(holidays[0], "%Y-%m-%d")
        h1 = datetime.strptime(holidays[-1], "%Y-%m-%d")
        cafe_note += f" {h0.month}/{h0.day}–{h1.month}/{h1.day}는 추석 연휴입니다."
    if ilwon_pending:
        cafe_note += " 일원역캠퍼스 식단은 업데이트 예정입니다."
    payload = {
        "week": f"{first.year}년 {first.month}월 {first.day // 7 + 1}주",
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
    return f"""<!DOCTYPE html>
<html lang="ko">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <meta http-equiv="Cache-Control" content="no-cache, no-store, must-revalidate" />
  <meta http-equiv="Pragma" content="no-cache" />
  <title>SMC 주간 식단</title>
  <style>
    :root {{
      --bg: #f4f1ea;
      --paper: #fffdf8;
      --ink: #1c1916;
      --muted: #6b645b;
      --line: #e4ddd2;
      --accent: #b45309;
      --accent-soft: #f3e4d0;
      --green: #3f6b4a;
      --blue: #2f4f73;
      --night: #4b3f6b;
    }}
    * {{ box-sizing: border-box; }}
    html, body {{ margin: 0; padding: 0; background: var(--bg); color: var(--ink);
      font-family: "Apple SD Gothic Neo", "Pretendard", "Noto Sans KR", sans-serif;
      line-height: 1.45; }}
    header {{
      position: sticky; top: 0; z-index: 20;
      background: #f4f1eaee; backdrop-filter: blur(10px);
      border-bottom: 1px solid var(--line);
    }}
    .wrap {{ max-width: 1180px; margin: 0 auto; padding: 18px 20px 40px; }}
    header .wrap {{ padding-bottom: 14px; }}
    h1 {{ margin: 0 0 4px; font-size: 22px; letter-spacing: -0.03em; }}
    .sub {{ color: var(--muted); font-size: 13px; }}
    .dates {{ display: flex; gap: 8px; flex-wrap: wrap; margin-top: 14px; }}
    .dates button {{
      appearance: none; border: 1px solid var(--line); background: var(--paper);
      border-radius: 999px; padding: 8px 14px; cursor: pointer; color: var(--ink);
      font: inherit; font-size: 14px;
    }}
    .dates button[aria-selected="true"] {{
      background: var(--ink); color: #fff; border-color: var(--ink);
    }}
    .dates button.is-today:not([aria-selected="true"]) {{
      border-color: var(--accent); color: var(--accent);
    }}
    .views {{
      margin-top: 12px; font-size: 12px; color: var(--muted);
      display: flex; flex-wrap: wrap; gap: 8px 14px; align-items: baseline;
    }}
    .views strong {{ color: var(--ink); font-weight: 700; }}
    .views .v-day {{ font-variant-numeric: tabular-nums; }}
    .loc-nav {{ display: flex; gap: 8px; flex-wrap: wrap; margin: 16px 0 8px; }}
    .loc-nav a {{
      color: var(--muted); text-decoration: none; font-size: 13px;
      border-bottom: 1px solid transparent;
    }}
    .loc-nav a:hover {{ color: var(--ink); border-color: var(--ink); }}
    section.place {{
      background: var(--paper); border: 1px solid var(--line);
      border-radius: 16px; padding: 20px 20px 16px; margin: 16px 0 0;
    }}
    section.place h2 {{
      margin: 0 0 4px; font-size: 18px; letter-spacing: -0.02em;
    }}
    .hours-line {{ color: var(--muted); font-size: 12px; margin-bottom: 14px; }}
    .meals {{
      display: grid; gap: 12px;
      grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
    }}
    .meal {{
      border: 1px solid var(--line); border-radius: 12px; padding: 12px 12px 8px;
      background: #fff;
    }}
    .meal h3 {{
      margin: 0 0 2px; font-size: 13px; letter-spacing: 0.04em;
      text-transform: none; color: var(--accent);
    }}
    .meal .when {{ font-size: 11px; color: var(--muted); margin-bottom: 8px; }}
    .course {{ margin-bottom: 10px; }}
    .course h4 {{
      margin: 0 0 4px; font-size: 12px; color: var(--blue); font-weight: 700;
    }}
    .course ul {{ margin: 0; padding: 0 0 0 16px; }}
    .course li {{ font-size: 13.5px; margin: 2px 0; }}
    .course li.theme {{ list-style: none; margin-left: -16px; font-weight: 700; color: var(--green); }}
    .empty {{ color: var(--muted); font-size: 13px; padding: 8px 0; }}
    footer {{ color: var(--muted); font-size: 12px; margin-top: 28px; }}
    .title-row {{ display: flex; align-items: center; justify-content: space-between; gap: 12px; }}
    .pals-toggle {{
      appearance: none; border: 1px solid var(--line); background: var(--paper);
      border-radius: 999px; padding: 6px 12px; cursor: pointer; color: var(--muted);
      font: inherit; font-size: 12px; white-space: nowrap;
    }}
    .pals-toggle:hover {{ color: var(--ink); border-color: var(--ink); }}
    .pals {{ position: fixed; inset: 0; z-index: 15; pointer-events: none; overflow: hidden; }}
    .pals[hidden] {{ display: none; }}
    .pal {{
      position: absolute; left: 0; top: 0; pointer-events: auto; cursor: pointer;
      will-change: transform; -webkit-tap-highlight-color: transparent; user-select: none;
    }}
    .pal-body {{ transform-origin: 50% 100%; }}
    .pal-body.boing {{ animation: boing 0.5s ease-out; }}
    .pal img {{
      display: block; height: var(--pal-size, 68px); width: auto;
      filter: drop-shadow(0 3px 3px rgba(60, 40, 20, 0.18));
      -webkit-user-drag: none; pointer-events: none;
    }}
    .pal-say {{
      position: absolute; left: 50%; bottom: 100%; transform: translate(-50%, -4px) scale(0.6);
      background: #fff; border: 1.5px solid #4a3a35; border-radius: 12px;
      padding: 3px 9px; font-size: 12px; font-weight: 700; color: #4a3a35;
      white-space: nowrap; opacity: 0; transition: opacity 0.15s, transform 0.15s;
    }}
    .pal-say.show {{ opacity: 1; transform: translate(-50%, -4px) scale(1); }}
    @keyframes boing {{
      0% {{ transform: scale(1, 1); }}
      20% {{ transform: scale(1.25, 0.75); }}
      45% {{ transform: scale(0.85, 1.2); }}
      70% {{ transform: scale(1.08, 0.94); }}
      100% {{ transform: scale(1, 1); }}
    }}
    @media print {{
      header {{ position: static; background: #fff; }}
      .dates, .loc-nav, .views, .pals, .pals-toggle {{ display: none; }}
      section.place {{ break-inside: avoid; }}
    }}
  </style>
</head>
<body>
  <header>
    <div class="wrap">
      <div class="title-row">
        <h1>삼성서울병원 주간 식단</h1>
        <button type="button" class="pals-toggle" id="pals-toggle">친구들 숨기기</button>
      </div>
      <div class="sub" id="weekline"></div>
      <div class="dates" id="dates" role="tablist" aria-label="날짜 선택"></div>
      <div class="views" id="views" hidden>조회수를 불러오는 중</div>
      <nav class="loc-nav" id="locnav"></nav>
    </div>
  </header>
  <main class="wrap" id="main"></main>
  <div class="pals" id="pals" aria-hidden="true"></div>
  <script id="meal-data" type="application/json">{data_json}</script>
  <script>
    const DATA = JSON.parse(document.getElementById("meal-data").textContent);
    const LOC_ORDER = ["본관 직원식당", "본관 밀카페", "암병원 직원식당", "암병원 밀카페", "일원역캠퍼스 식당"];
    const MEAL_ORDER = ["아침", "점심", "저녁", "야간", "카페"];
    const today = new Date();
    const todayIso = [
      today.getFullYear(),
      String(today.getMonth() + 1).padStart(2, "0"),
      String(today.getDate()).padStart(2, "0"),
    ].join("-");

    const weekline = document.getElementById("weekline");
    weekline.textContent = DATA.week + " · " + DATA.range;

    const datesEl = document.getElementById("dates");
    const viewsEl = document.getElementById("views");
    const viewsByDay = {{}};
    const COUNT_NS = "iamguno.github.io";
    const COUNT_BASE = "https://abacus.jsn.cam";
    let selected = DATA.days.some(d => d.date === todayIso) ? todayIso : DATA.days[0].date;

    function dayOf(iso) {{
      return DATA.days.find(d => d.date === iso);
    }}

    function renderDateButtons() {{
      datesEl.innerHTML = "";
      DATA.days.forEach(d => {{
        const b = document.createElement("button");
        b.type = "button";
        b.textContent = d.label;
        b.setAttribute("aria-selected", d.date === selected ? "true" : "false");
        if (d.date === todayIso) {{
          b.classList.add("is-today");
          if (d.date === selected) b.textContent = d.label + " · 오늘";
        }}
        b.addEventListener("click", () => {{
          selected = d.date;
          render();
        }});
        datesEl.appendChild(b);
      }});
    }}

    function hoursFor(loc, meal) {{
      const h = DATA.hours[loc];
      if (!h) return "";
      return h[meal] || h["운영"] || "";
    }}

    function renderCourses(courses, loc, mealName) {{
      const wrap = document.createElement("div");
      wrap.className = "meal";
      const h3 = document.createElement("h3");
      h3.textContent = mealName === "카페" ? "오늘의 메뉴" : mealName;
      wrap.appendChild(h3);
      const when = document.createElement("div");
      when.className = "when";
      when.textContent = hoursFor(loc, mealName);
      if (when.textContent) wrap.appendChild(when);
      const names = Object.keys(courses);
      names.forEach(name => {{
        const items = courses[name];
        if (!items || !items.length) return;
        const c = document.createElement("div");
        c.className = "course";
        if (!(names.length === 1 && (name === "메뉴" || name === "카페"))) {{
          const h4 = document.createElement("h4");
          const extra = (DATA.hours[loc] && DATA.hours[loc][name]) ? " · " + DATA.hours[loc][name] : "";
          h4.textContent = name + extra;
          c.appendChild(h4);
        }}
        const ul = document.createElement("ul");
        items.forEach(it => {{
          const li = document.createElement("li");
          const isTheme = /^<<.+>>$/.test(it) || (/^\\*.+\\*$/.test(it) && it.length < 40);
          li.textContent = isTheme ? it.replace(/^<<|>>$/g, "").replace(/^\\*|\\*$/g, "") : it;
          if (isTheme) li.className = "theme";
          ul.appendChild(li);
        }});
        c.appendChild(ul);
        wrap.appendChild(c);
      }});
      return wrap;
    }}

    function renderPlace(title, meals) {{
      const sec = document.createElement("section");
      sec.className = "place";
      sec.id = "loc-" + title;
      const h2 = document.createElement("h2");
      h2.textContent = title;
      sec.appendChild(h2);
      const hours = DATA.hours[title];
      if (hours && hours["운영"]) {{
        const p = document.createElement("div");
        p.className = "hours-line";
        p.textContent = hours["운영"];
        sec.appendChild(p);
      }}
      const grid = document.createElement("div");
      grid.className = "meals";
      const keys = MEAL_ORDER.filter(k => meals[k] && Object.keys(meals[k]).length);
      if (!keys.length) {{
        const empty = document.createElement("div");
        empty.className = "empty";
        empty.textContent = title.includes("밀카페") || title.includes("일원")
          ? "이 요일에는 운영/식단이 없습니다."
          : "이 날짜에 등록된 메뉴가 없습니다.";
        sec.appendChild(empty);
        return sec;
      }}
      keys.forEach(k => grid.appendChild(renderCourses(meals[k], title, k)));
      sec.appendChild(grid);
      return sec;
    }}

    function render() {{
      renderDateButtons();
      const day = dayOf(selected);
      const main = document.getElementById("main");
      main.innerHTML = "";
      const locnav = document.getElementById("locnav");
      locnav.innerHTML = "";

      LOC_ORDER.forEach(name => {{
        const a = document.createElement("a");
        a.href = "#loc-" + name;
        a.textContent = name;
        locnav.appendChild(a);

        const meals = day.locations[name] || {{}};
        main.appendChild(renderPlace(name, meals));
      }});

      const foot = document.createElement("footer");
      foot.textContent = DATA.cafeNote + " 메뉴·원산지는 당일 사정에 따라 바뀔 수 있습니다.";
      main.appendChild(foot);
    }}

    async function countGet(key) {{
      const res = await fetch(COUNT_BASE + "/get/" + COUNT_NS + "/" + encodeURIComponent(key));
      if (res.status === 404) return 0;
      if (!res.ok) throw new Error("count get");
      const data = await res.json();
      return Number(data.value) || 0;
    }}

    async function countHit(key) {{
      const res = await fetch(COUNT_BASE + "/hit/" + COUNT_NS + "/" + encodeURIComponent(key));
      if (!res.ok) throw new Error("count hit");
      const data = await res.json();
      return Number(data.value) || 0;
    }}

    async function loadViews() {{
      const onSite = location.hostname === "iamguno.github.io";
      try {{
        for (const d of DATA.days) {{
          viewsByDay[d.date] = d.date === todayIso && onSite
            ? await countHit("menu-" + d.date)
            : await countGet("menu-" + d.date);
        }}
        const todayCount = viewsByDay[todayIso] || 0;
        const weekCount = DATA.days.reduce((s, day) => s + (viewsByDay[day.date] || 0), 0);
        viewsEl.hidden = false;
        viewsEl.innerHTML = "";
        const summary = document.createElement("div");
        summary.innerHTML = "조회수 · 오늘 <strong>" + todayCount + "</strong>회 · 이번 주 <strong>" + weekCount + "</strong>회";
        viewsEl.appendChild(summary);
        DATA.days.forEach(d => {{
          const span = document.createElement("span");
          span.className = "v-day";
          span.textContent = d.md + " " + (viewsByDay[d.date] || 0) + "회";
          viewsEl.appendChild(span);
        }});
      }} catch (err) {{
        viewsEl.hidden = true;
      }}
    }}

    const PALS = [
      {{ name: "chiikawa", speed: 0.9, jump: 9, lines: ["우…", "야!", "울먹…", "맛있겠다…"] }},
      {{ name: "hachiware", speed: 1.2, jump: 10, lines: ["어떻게든 될 거야!", "오늘 점심 뭐야?", "같이 먹자~"] }},
      {{ name: "usagi", speed: 2.2, jump: 15, lines: ["우라!", "야하!", "하아?", "푸루루루"] }},
      {{ name: "momonga", speed: 1.5, jump: 11, lines: ["나 귀엽지?", "칭찬해 줘!", "그거 내 거야!"] }},
      {{ name: "kurimanju", speed: 0.6, jump: 7, lines: ["하아~", "한 잔 더…", "안주 뭐 있어?"] }},
      {{ name: "rakko", speed: 1.0, jump: 13, lines: ["…", "수련 중", "흠."] }},
      {{ name: "shisa", speed: 1.1, jump: 10, lines: ["열심히 할게요!", "~입니다!", "어서 오세요!"] }},
    ];
    const palsEl = document.getElementById("pals");
    const palsToggle = document.getElementById("pals-toggle");
    const reduceMotion = window.matchMedia("(prefers-reduced-motion: reduce)").matches;
    const pals = [];
    let palsFrame = null;

    function makePals() {{
      const size = window.innerWidth < 600 ? 52 : 68;
      palsEl.style.setProperty("--pal-size", size + "px");
      PALS.forEach((def, i) => {{
        const el = document.createElement("div");
        el.className = "pal";
        const body = document.createElement("div");
        body.className = "pal-body";
        const img = document.createElement("img");
        img.src = "assets/characters/" + def.name + ".png";
        img.alt = "";
        const say = document.createElement("span");
        say.className = "pal-say";
        body.appendChild(img);
        el.appendChild(say);
        el.appendChild(body);
        palsEl.appendChild(el);
        const angle = Math.random() * Math.PI * 2;
        const p = {{
          ...def, el, body, img, say,
          x: Math.random() * Math.max(1, window.innerWidth - size),
          y: 120 + Math.random() * Math.max(1, window.innerHeight - size - 120),
          vx: Math.cos(angle) * def.speed,
          vy: Math.sin(angle) * def.speed * 0.6,
          hop: 0, hopV: 0, hopping: false,
          phase: i * 1.7, sayTimer: null,
        }};
        el.addEventListener("click", () => poke(p));
        pals.push(p);
      }});
    }}

    function poke(p) {{
      p.hopping = true;
      p.hopV = -p.jump;
      p.vx = (Math.random() < 0.5 ? -1 : 1) * p.speed * 1.8;
      p.body.classList.remove("boing");
      void p.body.offsetWidth;
      p.body.classList.add("boing");
      p.say.textContent = p.lines[Math.floor(Math.random() * p.lines.length)];
      p.say.classList.add("show");
      clearTimeout(p.sayTimer);
      p.sayTimer = setTimeout(() => p.say.classList.remove("show"), 1600);
    }}

    function stepPals(t) {{
      const W = window.innerWidth;
      const H = window.innerHeight;
      pals.forEach(p => {{
        const w = p.el.offsetWidth || 60;
        const h = p.el.offsetHeight || 60;
        const maxV = p.hopping ? p.speed * 1.8 : p.speed;
        p.vx += (Math.random() - 0.5) * 0.08 * p.speed;
        p.vy += (Math.random() - 0.5) * 0.06 * p.speed;
        p.vx = Math.max(-maxV, Math.min(maxV, p.vx));
        p.vy = Math.max(-maxV * 0.6, Math.min(maxV * 0.6, p.vy));
        p.x += p.vx;
        p.y += p.vy;
        if (p.x < 0) {{ p.x = 0; p.vx = Math.abs(p.vx); }}
        if (p.x > W - w) {{ p.x = W - w; p.vx = -Math.abs(p.vx); }}
        if (p.y < 0) {{ p.y = 0; p.vy = Math.abs(p.vy); }}
        if (p.y > H - h) {{ p.y = H - h; p.vy = -Math.abs(p.vy); }}
        if (p.hopping) {{
          p.hop += p.hopV;
          p.hopV += 0.7;
          if (p.hop >= 0) {{ p.hop = 0; p.hopV = 0; p.hopping = false; }}
        }}
        const bob = Math.sin(t / 500 * p.speed + p.phase) * 5;
        const tilt = Math.sin(t / 350 + p.phase) * 7;
        p.el.style.transform = "translate(" + p.x.toFixed(1) + "px," + (p.y + bob + p.hop).toFixed(1) + "px)";
        p.img.style.transform = "scaleX(" + (p.vx < 0 ? -1 : 1) + ") rotate(" + tilt.toFixed(1) + "deg)";
      }});
      palsFrame = requestAnimationFrame(stepPals);
    }}

    function setPals(on) {{
      palsEl.hidden = !on;
      palsToggle.textContent = on ? "친구들 숨기기" : "친구들 부르기";
      try {{ localStorage.setItem("pals", on ? "on" : "off"); }} catch (err) {{}}
      if (on) {{
        if (!pals.length) makePals();
        if (palsFrame === null) palsFrame = requestAnimationFrame(stepPals);
      }} else if (palsFrame !== null) {{
        cancelAnimationFrame(palsFrame);
        palsFrame = null;
      }}
    }}

    let palsPref = null;
    try {{ palsPref = localStorage.getItem("pals"); }} catch (err) {{}}
    palsToggle.addEventListener("click", () => setPals(palsEl.hidden));
    setPals(palsPref ? palsPref === "on" : !reduceMotion);

    render();
    loadViews();
  </script>
</body>
</html>
"""


if __name__ == "__main__":
    main()
