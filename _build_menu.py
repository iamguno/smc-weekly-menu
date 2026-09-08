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


# 본관 밀카페 — 주간 메뉴 이미지(9/7–9/11)에서 정리
CAFE_BONGWAN = {
    "2026-09-07": {
        "빵": ["크로와상", "롤롤페스츄리", "시나몬베이글 + 크림치즈"],
        "샐러드": ["견과샐러드", "컵샐러드"],
        "랩·샌드위치": ["햄치즈베이글샌드위치", "에그햄치즈랩"],
        "컵밥": ["아비꼬돈가스카레라이스"],
    },
    "2026-09-08": {
        "빵": ["소보로빵", "크랜베리스콘 + 딸기잼", "빠네디까사런치롤"],
        "샐러드": ["푸실리샐러드", "콥샐러드"],
        "랩·샌드위치": ["햄치즈샌드위치", "너비아니토마토랩"],
        "컵밥": ["나시고랭"],
    },
    "2026-09-09": {
        "빵": ["단백쿠키(초코)", "크림치즈프레즐", "어니언베이글 + 크림치즈"],
        "샐러드": ["시리얼샐러드", "수제요거트볼"],
        "랩·샌드위치": ["닭가슴살햄샌드위치", "게맛살에그랩"],
        "컵밥": ["양념치킨컵밥"],
    },
    "2026-09-10": {
        "빵": ["블루베리머핀", "치즈번", "다크치아바타"],
        "샐러드": ["컵샐러드", "카프레제샐러드"],
        "랩·샌드위치": ["데리야끼치킨샌드위치", "단호박리코타치즈랩"],
        "컵밥": ["불고기버섯덮밥"],
    },
    "2026-09-11": {
        "빵": ["카스테라", "초코칩머핀", "모짜렐라베이글 + 크림치즈"],
        "샐러드": ["닭가슴살샐러드", "수제요거트볼"],
        "랩·샌드위치": ["매콤참치샌드위치", "스파이시치킨랩"],
        "컵밥": ["추억의도시락컵밥"],
    },
}

# 암병원 밀카페 — 주간 메뉴 이미지(9/7–9/11)에서 정리
CAFE_AM = {
    "2026-09-07": {
        "빵": ["먹물치아바타", "고구마빵", "플레인카스테라"],
        "샐러드": ["보코치니샐러드", "푸실리샐러드", "수제요거트"],
        "랩·샌드위치": ["어니언치킨랩", "폴드포크샌드위치"],
        "컵밥": ["베이컨김치볶음밥"],
    },
    "2026-09-08": {
        "빵": ["올리브 포카치아", "치즈케이크", "크림치즈프레즐"],
        "샐러드": ["닭가슴살샐러드", "구운감자샐러드", "수제요거트"],
        "랩·샌드위치": ["랜치소시지랩", "블랙번샌드위치"],
        "컵밥": ["참치생야채컵밥"],
    },
    "2026-09-09": {
        "빵": ["통밀베이글 + 크림치즈", "플레인스콘", "소보로빵"],
        "샐러드": ["견과샐러드", "푸실리샐러드", "수제요거트"],
        "랩·샌드위치": ["멕시칸치킨랩", "불고기치즈버거"],
        "컵밥": ["닭갈비컵밥"],
    },
    "2026-09-10": {
        "빵": ["무화과로프", "크림치즈프레즐", "플레인카스테라"],
        "샐러드": ["구운버섯샐러드", "구운감자샐러드", "수제요거트"],
        "랩·샌드위치": ["케이준치킨시저랩", "대만식햄치즈샌드위치"],
        "컵밥": ["버터장조림컵밥"],
    },
    "2026-09-11": {
        "빵": ["브라운브레드", "치즈방앗간", "크로와상"],
        "샐러드": ["닭가슴살샐러드", "보코치니샐러드", "수제요거트"],
        "랩·샌드위치": ["비프치폴레랩", "햄치즈크라상샌드위치"],
        "컵밥": ["제육불고기컵밥"],
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


def build_days(bongwan, am, ilwon):
    dates = sorted(set(bongwan) | set(am) | set(ilwon) | set(CAFE_BONGWAN) | set(CAFE_AM))
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
    ilwon = parse_ilwon()
    days = build_days(bongwan, am, ilwon)
    payload = {
        "week": "2026년 9월 2주",
        "range": "9/7 (월) – 9/13 (일)",
        "hours": HOURS,
        "days": days,
        "cafeNote": "본관·암병원 밀카페는 평일(9/7–9/11)만 운영합니다.",
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
    @media print {{
      header {{ position: static; background: #fff; }}
      .dates, .loc-nav {{ display: none; }}
      section.place {{ break-inside: avoid; }}
    }}
  </style>
</head>
<body>
  <header>
    <div class="wrap">
      <h1>삼성서울병원 주간 식단</h1>
      <div class="sub" id="weekline"></div>
      <div class="dates" id="dates" role="tablist" aria-label="날짜 선택"></div>
      <nav class="loc-nav" id="locnav"></nav>
    </div>
  </header>
  <main class="wrap" id="main"></main>
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

    render();
  </script>
</body>
</html>
"""


if __name__ == "__main__":
    main()
