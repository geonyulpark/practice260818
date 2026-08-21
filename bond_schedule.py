# -*- coding: utf-8 -*-
"""회사채 수요예측 일정 — Compass 부가 모듈.

한국투자증권 / 하나증권 / NH투자증권이 매일 보내는 발행 예정 리스트를 읽어
딜 단위로 합치고, 캘린더로 보여준다.

이 파일은 Streamlit Cloud 에서도 돌아가야 하므로 Windows 전용 의존성(pywin32 등)을
쓰지 않는다. Outlook 수집은 로컬 push_to_sheets.py 가 맡는다.

app.py 는 Compass 의 Google Sheets 헬퍼를 넘겨주는 방식으로 이 모듈을 쓴다
(순환 import 를 피하려고 함수를 주입받는다).
"""
from __future__ import annotations

import calendar
import datetime as dt
import difflib
import html
import io
import re
import unicodedata

import openpyxl
import pandas as pd
import pdfplumber

SHEET_NAME = "수요예측일정_이력"

# 별칭 시트에 추가로 쓸 소스 열 이름 (app.py 의 ALIAS_SOURCE_COLUMNS 에도 넣어둘 것)
ALIAS_SOURCE_COL = "증권사(발행리스트)"

COLUMNS = [
    "수요예측일", "발행일", "종목명", "신용등급", "만기", "물량",
    "신고발행금액", "최대발행가능액", "금리밴드", "대표주관",
    "상태", "출처", "자료일자", "비고",
]
DATE_COLS = ("수요예측일", "발행일", "자료일자")
NUM_COLS = ("신고발행금액", "최대발행가능액")

PRIORITY = {"한국투자": 0, "하나": 1, "NH": 2}
DATE_TOL = 7      # 같은 딜로 볼 발행일 오차(일). 자료마다 며칠씩 다르게 적는다.
NAME_SIM = 0.75   # 발행사명 유사도 하한

RATING_RE = re.compile(
    r"^(AAA|AA\+|AA0|AA-|A\+|A0|A-|BBB\+|BBB0|BBB-|BB\+|BB0|BB-|B\+|B0|B-|CCC|CC|C|D)"
    r"(\(P\)|\(S\)|\(N\))?$")


# ==================================================================== 정규화
def squash(s) -> str:
    if s is None:
        return ""
    s = str(s)
    if s.lower() in ("nan", "nat", "none"):
        return ""
    s = unicodedata.normalize("NFKC", s).replace(" ", " ")
    return re.sub(r"\s+", " ", s).strip()


def lines(s) -> list:
    if s is None:
        return []
    s = unicodedata.normalize("NFKC", str(s)).replace(" ", " ")
    out = [re.sub(r"\s+", " ", x).strip() for x in s.split("\n")]
    return [x for x in out if x]


def norm_header(s) -> str:
    return re.sub(r"\s+", "", squash(s))


_MD = re.compile(r"(\d{1,2})\s*[/월.\-]\s*(\d{1,2})")


def _safe_date(y, m, d):
    try:
        return dt.date(y, m, d)
    except ValueError:
        return None


def parse_md(text, as_of, prefer: str | None = None):
    """'8/21(금)' -> date. 연도는 자료일자 기준으로 추정한다.

    1월/8월처럼 6개월 이상 떨어진 날짜는 작년/내년이 모호하므로 딜 상태를 힌트로 받는다.
      prefer="past"   : 완료 딜 -> 자료일자 이전 중 가장 최근 연도
      prefer="future" : 예정 딜 -> 자료일자 이후 중 가장 이른 연도
    """
    t = squash(text)
    if not t or "미정" in t or t in ("-", "TBD"):
        return None
    m = _MD.search(t)
    if not m:
        return None
    mm, dd = int(m.group(1)), int(m.group(2))
    if not (1 <= mm <= 12 and 1 <= dd <= 31):
        return None
    base = as_of or dt.date.today()
    cands = [c for c in (_safe_date(y, mm, dd)
                         for y in (base.year - 1, base.year, base.year + 1)) if c]
    if not cands:
        return None
    SLACK = 45
    if prefer == "past":
        past = [c for c in cands if (c - base).days <= SLACK]
        if past:
            return max(past)
    elif prefer == "future":
        fut = [c for c in cands if (c - base).days >= -SLACK]
        if fut:
            return min(fut)
    return min(cands, key=lambda c: abs((c - base).days))


def to_date(v):
    if v is None:
        return None
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    t = squash(v)
    if not t or "미정" in t:
        return None
    for fmt in ("%Y-%m-%d", "%Y.%m.%d", "%Y/%m/%d", "%y-%m-%d", "%y.%m.%d"):
        try:
            return dt.datetime.strptime(t[:10], fmt).date()
        except ValueError:
            pass
    return None


def date_from_filename(name: str):
    """파일명에서 자료일자 추출: _260820 / _2026.08.20 / 20260820"""
    n = squash(name)
    m = re.search(r"(20\d{2})[.\-_]?(\d{2})[.\-_]?(\d{2})", n)
    if m:
        return _safe_date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    m = re.search(r"[_\-](\d{2})(\d{2})(\d{2})(?!\d)", n)
    if m:
        return _safe_date(2000 + int(m.group(1)), int(m.group(2)), int(m.group(3)))
    return None


def to_num(v):
    if v is None:
        return None
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return float(v)
    t = squash(v).replace(",", "")
    if not t or "미정" in t or t in ("-", "/"):
        return None
    m = re.search(r"-?\d+(?:\.\d+)?", t)
    return float(m.group()) if m else None


def join_tranches(vals) -> str:
    return "/".join(squash(v) for v in vals if squash(v))


def sum_tranches(vals):
    nums = [n for n in (to_num(v) for v in vals) if n is not None]
    return sum(nums) if nums else None


def clean_volume(v) -> str:
    t = squash(v)
    return "" if ("미정" in t or t in ("-", "/")) else t


def clean_issuer(v) -> str:
    """표시용 종목명. 줄바꿈으로 생긴 괄호 앞 공백을 제거한다."""
    return re.sub(r"\s+\(", "(", squash(v))


def clean_rating(v) -> str:
    return squash(v).replace(" ", "").replace("（", "(").replace("）", ")")


_HANGUL_LETTER = [
    ("더블유", "W"), ("에이치", "H"), ("에스", "S"), ("에프", "F"), ("에이", "A"),
    ("제이", "J"), ("케이", "K"), ("아이", "I"), ("브이", "V"), ("제트", "Z"),
    ("엑스", "X"), ("와이", "Y"), ("엘", "L"), ("엠", "M"), ("엔", "N"),
    ("오", "O"), ("피", "P"), ("큐", "Q"), ("알", "R"), ("티", "T"),
    ("유", "U"), ("비", "B"), ("씨", "C"), ("디", "D"), ("이", "E"), ("지", "G"),
]


def latinize(name: str) -> str:
    """한글 이니셜 표기를 영문으로. 비교 키 전용(표시용 아님).

    '케이씨씨글라스' -> 'KCC글라스', '보령엘엔지터미널' -> '보령LNG터미널'
    과하게 걸려도('이마트' -> 'E마트') 3사 모두 같은 규칙을 타므로 비교에는 지장이 없다.
    """
    out, i = [], 0
    while i < len(name):
        for syl, ch in _HANGUL_LETTER:
            if name.startswith(syl, i):
                out.append(ch)
                i += len(syl)
                break
        else:
            out.append(name[i])
            i += 1
    return "".join(out)


_SUFFIX = re.compile(r"(주식회사|\(주\)|㈜|\(유\)|유한회사|Co\.,?Ltd\.?|Inc\.?)", re.I)

# 3사 표기가 다른 발행사 (Compass 별칭 시트로도 덮어쓸 수 있다)
ALIAS = {
    "한화위탁관리부동산투자회사": "한화리츠",
    "교보생명": "교보생명보험",
}


def norm_issuer(name: str, alias_map: dict | None = None) -> str:
    """비교용 키. Compass 별칭 시트를 먼저 적용하고, 없으면 자체 규칙."""
    t = squash(name)
    if alias_map:
        t = alias_map.get(t, t)
    t = _SUFFIX.sub("", t).replace(" ", "")
    # 신종자본증권/후순위는 같은 발행사의 다른 종류이므로 구분을 남긴다
    kind = ""
    if re.search(r"신종", t):
        kind = "(신종)"
    elif re.search(r"후순위|\(후\)", t):
        kind = "(후순위)"
    base = re.sub(r"\([^)]*\)", "", t)
    if alias_map:
        base = alias_map.get(base, base)
    base = ALIAS.get(base, base)
    return (latinize(base) + kind).lower()


_BAND_KIND = [
    (("개평", "개별"), "개별민평"),
    (("등평", "등급"), "등급민평"),
    (("절대", "고정"), "고정"),
]


def norm_band(v) -> str:
    """3사 금리밴드 표기를 통일.
    '개평 ± 30bp' / '개별 -30~+30'         -> '개별민평 ±30bp'
    '절대 4.30% ~ 4.90%' / '고정 4.3%~4.9%' -> '고정 4.30~4.90%'
    """
    t = squash(v)
    if not t or "미정" in t:
        return ""
    kind = next((label for keys, label in _BAND_KIND if any(k in t for k in keys)), "")
    pct = re.findall(r"(\d+(?:\.\d+)?)\s*%", t)
    if len(pct) >= 2:
        return f"{kind or '고정'} {float(pct[0]):.2f}~{float(pct[1]):.2f}%".strip()
    bps = [int(b) for b in re.findall(r"([+-]?\d+)\s*(?:bp)?", t.replace("±", " "))
           if b.lstrip("+-").isdigit()]
    if bps:
        return f"{kind or '민평'} ±{max(abs(b) for b in bps)}bp".strip()
    return kind or t


# ==================================================================== 파서
def _blank_row(**kw) -> dict:
    row = {c: None for c in COLUMNS}
    row.update(kw)
    return row


# ---------------------------------------------------- 한국투자증권 PDF
KIS_HEADER_MAP = {
    "종목명": "종목명", "신용등급": "신용등급", "만기": "만기",
    "최초모집물량(억)": "물량", "신고(억)": "신고발행금액", "증액(억)": "최대발행가능액",
    "금리밴드": "금리밴드", "수요예측일": "수요예측일", "발행일": "발행일",
    "대표주관(인수단)": "대표주관", "비고": "비고",
}


def _kis_parse(pdf, filename: str) -> list:
    head = (pdf.pages[0].extract_text() or "")[:200]
    m = re.search(r"(20\d{2})[-.](\d{1,2})[-.](\d{1,2})", head)
    as_of = _safe_date(*map(int, m.groups())) if m else date_from_filename(filename)

    out = []
    for page in pdf.pages:
        for table in page.extract_tables():
            idx, status = {}, "예정"
            for raw in table:
                cells = [c if c is not None else "" for c in raw]
                joined = norm_header(" ".join(cells))
                if not joined:
                    continue
                # 섹션 제목행 ('진행 중인 Deal' / '발행 결정 완료 Deal')
                if joined.endswith("Deal") and sum(1 for c in cells if squash(c)) <= 2:
                    status = "완료" if "완료" in joined else "예정"
                    continue
                if "종목명" in [norm_header(c) for c in cells]:
                    idx = {}
                    for i, c in enumerate(cells):
                        key = KIS_HEADER_MAP.get(norm_header(c))
                        if key and key not in idx:
                            idx[key] = i
                    continue
                if "종목명" not in idx:
                    continue

                def g(k):
                    i = idx.get(k)
                    return cells[i] if i is not None and i < len(cells) else ""

                name = squash(g("종목명"))
                rating = clean_rating(g("신용등급"))
                if not name or not rating:
                    continue
                hint = "past" if status == "완료" else "future"
                vol = squash(g("물량"))
                out.append(_blank_row(
                    수요예측일=parse_md(g("수요예측일"), as_of, hint),
                    발행일=parse_md(g("발행일"), as_of, hint),
                    종목명=clean_issuer(name),
                    신용등급=rating,
                    만기=squash(g("만기")).replace(" ", ""),
                    물량=join_tranches(vol.split("/")) if vol else "",
                    신고발행금액=to_num(g("신고발행금액")),
                    최대발행가능액=to_num(g("최대발행가능액")),
                    금리밴드=squash(g("금리밴드")),
                    대표주관=squash(g("대표주관")),
                    상태=status, 출처="한국투자", 자료일자=as_of,
                    비고=squash(g("비고")),
                ))
    return out


# ---------------------------------------------------- 하나증권 PDF
# 표가 병합셀 때문에 헤더가 한 칸씩 밀린다. '등급' 열을 기준점으로 상대 인덱싱한다.
HANA_OFF = {
    "상태": -3, "종목명": -2, "신용등급": 0, "전망": 1, "발행목적": 2,
    "신고서제출일": 3, "수요예측일": 4, "수요예측시간": 5, "발행일": 6,
    "만기": 7, "ESG": 8, "신고발행금액": 9, "최대발행가능액": 10, "금리밴드": 11,
}


def _hana_parse(pdf, filename: str) -> list:
    head = (pdf.pages[0].extract_text() or "")[:300]
    m = re.search(r"(20\d{2})-(\d{2})-(\d{2})", head)
    as_of = _safe_date(*map(int, m.groups())) if m else date_from_filename(filename)

    out = []
    for page in pdf.pages:
        for table in page.extract_tables():
            g_idx, lead_idx = None, None
            for raw in table:
                cells = [c if c is not None else "" for c in raw]
                hdr = [norm_header(c) for c in cells]
                if "등급" in hdr and "만기(Y)" in hdr and "납입일" in hdr:
                    g_idx = hdr.index("등급")
                    lead_idx = next((i for i, h in enumerate(hdr) if "대표주관사" in h), None)
                    continue
                if g_idx is None:
                    continue

                def at(off):
                    i = g_idx + off
                    return cells[i] if 0 <= i < len(cells) else ""

                name = squash(at(HANA_OFF["종목명"]))
                rating = clean_rating(at(HANA_OFF["신용등급"]))
                if not name or not RATING_RE.match(rating):
                    continue

                status = "완료" if "완료" in squash(at(HANA_OFF["상태"])) else "예정"
                hint = "past" if status == "완료" else "future"

                vol = lines(at(HANA_OFF["신고발행금액"]))
                filed = sum_tranches(vol)
                maxraw = squash(at(HANA_OFF["최대발행가능액"]))
                maxamt = filed if "증액없음" in maxraw else to_num(maxraw)

                lead = ""
                if lead_idx is not None and lead_idx < len(cells):
                    parts = [re.sub(r"^-\s*", "", x) for x in lines(cells[lead_idx])]
                    parts = [x for x in parts if not re.fullmatch(r"[^:]+:\s*", x)]
                    lead = " / ".join(parts)

                note = " / ".join(x for x in (
                    ("ESG:" + squash(at(HANA_OFF["ESG"]))) if squash(at(HANA_OFF["ESG"])) else "",
                    squash(at(HANA_OFF["발행목적"])),
                    squash(at(HANA_OFF["수요예측시간"])),
                ) if x)

                out.append(_blank_row(
                    수요예측일=parse_md(at(HANA_OFF["수요예측일"]), as_of, hint),
                    발행일=parse_md(at(HANA_OFF["발행일"]), as_of, hint),
                    종목명=clean_issuer(name),
                    신용등급=rating,
                    만기=join_tranches(_yfmt(x) for x in lines(at(HANA_OFF["만기"]))),
                    물량=join_tranches(vol),
                    신고발행금액=filed,
                    최대발행가능액=maxamt,
                    금리밴드=" ".join(lines(at(HANA_OFF["금리밴드"]))),
                    대표주관=lead,
                    상태=status, 출처="하나", 자료일자=as_of, 비고=note,
                ))
    return out


def _yfmt(m: str) -> str:
    t = squash(m)
    return f"{t}y" if re.fullmatch(r"\d+(\.\d+)?", t) else t


# ---------------------------------------------------- NH투자증권 XLSX
NH_HEAD_KEYS = {
    "발행회사": "종목명", "신용등급": "신용등급", "만기": "만기",
    "예측": "물량", "증액가능": "최대발행가능액", "대표주관": "대표주관",
    "수요예측": "수요예측일", "납입": "발행일", "ESG": "ESG",
}


def _nh_parse(fileobj, filename: str) -> list:
    wb = openpyxl.load_workbook(fileobj, data_only=True, read_only=True)
    ws = wb.worksheets[0]
    rows = [list(r) for r in ws.iter_rows(values_only=True)]
    wb.close()
    if not rows:
        return []

    as_of = to_date(rows[0][0]) if rows[0] else None
    as_of = as_of or date_from_filename(filename)

    hdr_i, idx, band_idx = None, {}, None
    for i, r in enumerate(rows[:15]):
        hdr = [norm_header(c) for c in r]
        if "발행회사" in hdr and "신용등급" in hdr:
            hdr_i = i
            # '수요예측' 과 '수요예측전일Spr' 을 구분해야 하므로 정확 일치를 먼저 본다
            for j, h in enumerate(hdr):
                for k, field in NH_HEAD_KEYS.items():
                    if h == norm_header(k) and field not in idx:
                        idx[field] = j
            for j, h in enumerate(hdr):
                for k, field in NH_HEAD_KEYS.items():
                    if field not in idx and h.startswith(norm_header(k)):
                        idx[field] = j
            idx["예측결과"] = next(
                (j for j, h in enumerate(hdr) if h.startswith("예측결과")), None)
            band_idx = next((j for j, h in enumerate(hdr) if h.startswith("금리밴드")), None)
            break
    if hdr_i is None:
        return []

    def at(row, key_or_i):
        i = idx.get(key_or_i) if isinstance(key_or_i, str) else key_or_i
        if i is None or row is None or i >= len(row):
            return ""
        return row[i]

    # 발행사 셀은 딜의 첫 트랜치 행에만 있고, 이후 행은 만기별 트랜치다
    deals = []
    for r in rows[hdr_i + 2:]:
        if squash(at(r, "종목명")):
            deals.append([r])
        elif deals and any(squash(c) for c in r[:20]):
            deals[-1].append(r)

    out = []
    for grp in deals:
        head = grp[0]
        name = clean_issuer(at(head, "종목명"))
        rating = clean_rating(at(head, "신용등급"))
        if not name or (rating and not RATING_RE.match(rating)):
            continue

        vols = [at(r, "물량") for r in grp]
        filed = sum_tranches(vols)
        maxraw = squash(at(head, "최대발행가능액"))
        maxamt = filed if "증액없음" in maxraw else to_num(maxraw)

        band = ""
        if band_idx is not None:
            band = " ".join(x for x in (squash(at(head, band_idx)),
                                        squash(at(head, band_idx + 1))) if x)

        esg = sorted({squash(at(r, "ESG")) for r in grp if squash(at(r, "ESG"))})
        done = bool(re.search(r"\d", squash(at(head, "예측결과"))))

        out.append(_blank_row(
            수요예측일=to_date(at(head, "수요예측일")),
            발행일=to_date(at(head, "발행일")),
            종목명=name, 신용등급=rating,
            만기=join_tranches(_yfmt(squash(at(r, "만기"))) for r in grp),
            물량=join_tranches(_vfmt(v) for v in vols),
            신고발행금액=filed, 최대발행가능액=maxamt, 금리밴드=band,
            대표주관=squash(at(head, "대표주관")).replace("/", ", "),
            상태="완료" if done else "예정", 출처="NH", 자료일자=as_of,
            비고=("ESG:" + " / ".join(esg)) if esg else "",
        ))
    return out


def _vfmt(v) -> str:
    n = to_num(v)
    return squash(v).replace("\n", " ") if n is None else f"{int(round(n)):,}"


# ---------------------------------------------------- 진입점
def parse_file(fileobj, filename: str) -> tuple:
    """파일 하나를 알맞은 파서로 처리. (증권사, rows) 반환. 인식 실패 시 ("", [])."""
    name = squash(filename)
    data = fileobj.read() if hasattr(fileobj, "read") else fileobj
    if isinstance(data, bytes):
        buf = io.BytesIO(data)
    else:
        buf = data

    if name.lower().endswith((".xlsx", ".xlsm")):
        if "발행예정" in name.replace(" ", "") or "발행예정리스트" in name:
            return "NH", _nh_parse(buf, name)
        return "", []

    if not name.lower().endswith(".pdf"):
        return "", []

    with pdfplumber.open(buf) as pdf:
        head = (pdf.pages[0].extract_text() or "")[:800]
        if "한투" in name or "한국투자증권" in head:
            return "한국투자", _kis_parse(pdf, name)
        if "하나증권" in name or "하나증권 공모회사채" in head:
            return "하나", _hana_parse(pdf, name)
    return "", []


# ==================================================================== 병합
def _anchor(row):
    """딜 식별 기준일: 발행일이 가장 안정적, 없으면 수요예측일."""
    return row["발행일"] or row["수요예측일"]


def _same_deal(a, b) -> bool:
    na, da = a
    nb, db = b
    if da is None or db is None or abs((da - db).days) > DATE_TOL:
        return False
    if na == nb or (len(na) >= 2 and (na in nb or nb in na)):
        return True
    return difflib.SequenceMatcher(None, na, nb).ratio() >= NAME_SIM


def _cluster(pairs: list) -> list:
    """(발행사명, 기준일) 리스트 -> 각 원소의 클러스터 id (union-find)."""
    n = len(pairs)
    parent = list(range(n))

    def find(x):
        while parent[x] != x:
            parent[x] = parent[parent[x]]
            x = parent[x]
        return x

    order = sorted((i for i in range(n) if pairs[i][1] is not None),
                   key=lambda i: pairs[i][1])
    for oi, i in enumerate(order):
        for j in order[oi + 1:]:
            if (pairs[j][1] - pairs[i][1]).days > DATE_TOL:
                break
            if _same_deal(pairs[i], pairs[j]):
                rx, ry = find(i), find(j)
                if rx != ry:
                    parent[max(rx, ry)] = min(rx, ry)
    return [find(i) for i in range(n)]


def _pick(series):
    """우선순위 순으로 첫 유효값. '미정'/빈값은 건너뛴다."""
    for v in series:
        if v is None:
            continue
        t = squash(v)
        if t and "미정" not in t and t != "-":
            return v
    return None


def _row_prio(src) -> int:
    parts = [x.strip() for x in squash(src).split(",") if x.strip()]
    return min((PRIORITY.get(x, 9) for x in parts), default=9)


def _fold(grp: pd.DataFrame) -> dict:
    """한 딜의 여러 기록을 필드 단위로 합친다.
    최신 자료일자 우선, 같은 날이면 증권사 우선순위. 빈 값은 과거 기록으로 보충."""
    grp = grp.sort_values(["_asof", "_prio"], ascending=[False, True])
    rec = {c: _pick(grp[c]) for c in COLUMNS
           if c not in ("출처", "상태", "자료일자")}
    srcs = set()
    for v in grp["출처"]:
        srcs |= {x.strip() for x in squash(v).split(",") if x.strip()}
    rec["출처"] = ", ".join(sorted(srcs, key=lambda x: PRIORITY.get(x, 9)))
    rec["상태"] = "완료" if (grp["상태"] == "완료").any() else "예정"
    rec["자료일자"] = max([d for d in grp["자료일자"] if d], default=None)
    return rec


def _drop_junk(df: pd.DataFrame) -> pd.DataFrame:
    """각주만 들어간 행('(GS글로벌 원리금 지급보증)') 등 딜이 아닌 행 제거."""
    name = df["종목명"].map(squash)
    ok = (name != "") & ~name.str.startswith("(")
    ok &= df["발행일"].notna() | df["수요예측일"].notna()
    return df[ok].reset_index(drop=True)


def merge_deals(df: pd.DataFrame, alias_map: dict | None = None) -> pd.DataFrame:
    """여러 증권사/여러 날짜의 행을 딜 단위 1행으로 합친다."""
    if df is None or df.empty:
        return pd.DataFrame(columns=COLUMNS)
    df = _drop_junk(df.copy().reset_index(drop=True))
    if df.empty:
        return pd.DataFrame(columns=COLUMNS)

    pairs = [(norm_issuer(r["종목명"], alias_map), _anchor(r)) for _, r in df.iterrows()]
    df["_cid"] = _cluster(pairs)
    df["_asof"] = df["자료일자"].map(lambda d: d or dt.date(1900, 1, 1))
    df["_prio"] = df["출처"].map(_row_prio)

    out = pd.DataFrame([_fold(g) for _, g in df.groupby("_cid", sort=False)])[COLUMNS]
    out["물량"] = out["물량"].map(clean_volume)
    out["금리밴드"] = out["금리밴드"].map(norm_band)
    return out.sort_values(["수요예측일", "발행일", "종목명"],
                           na_position="last").reset_index(drop=True)


def parse_uploads(files, alias_map: dict | None = None) -> tuple:
    """업로드된 파일들 -> (병합된 df, 파일별 로그). files 는 (fileobj, name) 또는 UploadedFile."""
    rows, log = [], []
    for f in files:
        if isinstance(f, tuple):
            obj, name = f
        else:
            obj, name = f, getattr(f, "name", "")
        try:
            src, rs = parse_file(obj, name)
        except Exception as e:
            log.append((name, "오류", str(e)[:120]))
            continue
        if not src:
            log.append((name, "인식 실패", "3사 형식이 아닙니다"))
            continue
        rows += rs
        log.append((name, src, f"{len(rs)}건"))
    if not rows:
        return pd.DataFrame(columns=COLUMNS), log
    return merge_deals(pd.DataFrame(rows)[COLUMNS], alias_map), log


# ==================================================================== 저장/조회
def _to_sheet_str(v, col) -> str:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    if col in DATE_COLS:
        d = to_date(v)
        return d.isoformat() if d else ""
    if col in NUM_COLS:
        n = to_num(v)
        return "" if n is None else str(int(round(n)))
    return squash(v)


def save_deals(get_client, sheet_id: str, df: pd.DataFrame):
    """워크시트를 통째로 교체한다.
    딜 병합이 유사도 기반이라 행 단위 append 로는 중복을 못 막기 때문."""
    import gspread

    client = get_client()
    sh = client.open_by_key(sheet_id)
    try:
        ws = sh.worksheet(SHEET_NAME)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=SHEET_NAME,
                              rows=max(1000, len(df) + 100), cols=len(COLUMNS) + 2)
    body = [COLUMNS] + [[_to_sheet_str(r[c], c) for c in COLUMNS]
                        for _, r in df.iterrows()]
    ws.clear()
    ws.update(body, "A1")


def read_deals(read_history) -> pd.DataFrame:
    """Compass 의 read_history 로 시트를 읽어 타입을 되살린다."""
    df = read_history(SHEET_NAME)
    if df is None or df.empty:
        return pd.DataFrame(columns=COLUMNS)
    for c in COLUMNS:
        if c not in df.columns:
            df[c] = ""
    df = df[COLUMNS].copy()
    for c in DATE_COLS:
        df[c] = df[c].map(to_date)
    for c in NUM_COLS:
        df[c] = df[c].map(to_num)
    for c in df.columns:
        if c not in DATE_COLS and c not in NUM_COLS:
            df[c] = df[c].map(squash)
    return df


def merge_into_sheet(new_df: pd.DataFrame, read_history, get_client, sheet_id: str,
                     alias_map: dict | None = None) -> pd.DataFrame:
    """기존 시트 + 새 데이터를 합쳐서 다시 저장하고, 최종 결과를 반환."""
    old = read_deals(read_history)
    both = pd.concat([old, new_df], ignore_index=True) if not old.empty else new_df
    merged = merge_deals(both, alias_map)
    save_deals(get_client, sheet_id, merged)
    return merged


# ==================================================================== 화면
WEEKDAYS = ["월", "화", "수", "목", "금", "토", "일"]
RATING_COLORS = [
    ("AAA", "#1e40af"), ("AA", "#2563eb"), ("A", "#0d9488"),
    ("BBB", "#d97706"), ("BB", "#dc2626"), ("B", "#b91c1c"),
]
FALLBACK = "#64748b"

CSS = """
<style>
.bsch-wrap { border:1px solid #e2e8f0; border-radius:10px; overflow:hidden; }
.bsch-head { display:grid; grid-template-columns:repeat(7,1fr); background:#f8fafc; }
.bsch-head div { padding:8px 6px; text-align:center; font-weight:600; font-size:.82rem;
                 color:#475569; border-right:1px solid #e2e8f0; }
.bsch-head div:last-child { border-right:none; }
.bsch-head .sat { color:#2563eb; } .bsch-head .sun { color:#dc2626; }
.bsch-grid { display:grid; grid-template-columns:repeat(7,1fr); }
.bsch-cell { min-height:104px; padding:5px 6px; border-top:1px solid #e2e8f0;
             border-right:1px solid #e2e8f0; }
.bsch-cell:nth-child(7n) { border-right:none; }
.bsch-cell.out { background:#fafafa; }
.bsch-cell.today { background:#fef9c3; }
.bsch-day { font-size:.78rem; font-weight:600; color:#334155; margin-bottom:3px; }
.bsch-cell.out .bsch-day { color:#94a3b8; }
.bsch-chip { display:block; font-size:.72rem; line-height:1.25; padding:2px 5px;
             margin:2px 0; border-radius:4px; color:#fff; white-space:nowrap;
             overflow:hidden; text-overflow:ellipsis; }
.bsch-chip.issue { background:#fff; border:1px dashed currentColor; }
.bsch-chip .amt { opacity:.85; font-weight:400; }
.bsch-legend { display:flex; gap:14px; flex-wrap:wrap; font-size:.78rem; color:#475569;
               margin:8px 2px 0; align-items:center; }
.bsch-legend .sw { display:inline-block; width:11px; height:11px; border-radius:3px;
                   margin-right:5px; vertical-align:-1px; }
</style>
"""


def rating_color(rating) -> str:
    r = str(rating or "").upper().replace(" ", "")
    return next((c for p, c in RATING_COLORS if r.startswith(p)), FALLBACK)


RATING_EMOJI = [
    ("AAA", "🔵"), ("AA", "🔵"), ("A", "🟢"),
    ("BBB", "🟠"), ("BB", "🔴"), ("B", "⚫"),
]
FALLBACK_EMOJI = "⚪"


def rating_emoji(rating) -> str:
    r = str(rating or "").upper().replace(" ", "")
    return next((e for p, e in RATING_EMOJI if r.startswith(p)), FALLBACK_EMOJI)


def fmt_amt(v) -> str:
    n = to_num(v)
    if n is None or n <= 0:
        return "-"
    return f"{n / 10000:.2f}조" if n >= 10000 else f"{n:,.0f}억"


def has_date(d) -> bool:
    # pd.NaT 는 datetime 서브클래스라 isinstance 만으로는 걸러지지 않는다
    return isinstance(d, dt.date) and not pd.isna(d)


def month_grid(df: pd.DataFrame, year: int, month: int, show_issue: bool) -> str:
    first = dt.date(year, month, 1)
    start = first - dt.timedelta(days=first.weekday())
    last = dt.date(year, month, calendar.monthrange(year, month)[1])
    end = last + dt.timedelta(days=6 - last.weekday())
    today = dt.date.today()

    by_day = {}
    for _, r in df.iterrows():
        if has_date(r["수요예측일"]):
            by_day.setdefault(r["수요예측일"], []).append(("demand", r))
        if show_issue and has_date(r["발행일"]):
            by_day.setdefault(r["발행일"], []).append(("issue", r))

    head = "".join(
        '<div class="{}">{}</div>'.format(
            "sat" if i == 5 else "sun" if i == 6 else "", w)
        for i, w in enumerate(WEEKDAYS))

    cells, d = [], start
    while d <= end:
        klass = "bsch-cell"
        if d.month != month:
            klass += " out"
        if d == today:
            klass += " today"
        chips = []
        for kind, r in sorted(by_day.get(d, []),
                              key=lambda x: (x[0] != "demand", str(x[1]["종목명"]))):
            color = rating_color(r["신용등급"])
            amt = fmt_amt(r["최대발행가능액"])
            tip = html.escape("{} ({}) · {}\n만기 {} · 신고 {} / 최대 {}\n{}\n주관 {}".format(
                r["종목명"], r["신용등급"], "수요예측" if kind == "demand" else "발행",
                r["만기"] or "-", fmt_amt(r["신고발행금액"]), amt,
                r["금리밴드"] or "밴드 미정", r["대표주관"] or "-"))
            nm = html.escape(str(r["종목명"])[:11])
            if kind == "demand":
                cls, style, label = "bsch-chip", "background:" + color, nm
            else:
                cls, style, label = "bsch-chip issue", "color:" + color, "발행 " + nm
            chips.append('<span class="{}" style="{}" title="{}">{} '
                         '<span class="amt">{}</span></span>'.format(
                             cls, style, tip, label, amt))
        cells.append('<div class="{}"><div class="bsch-day">{}</div>{}</div>'.format(
            klass, d.day, "".join(chips)))
        d += dt.timedelta(days=1)

    legend = "".join(
        '<span><span class="sw" style="background:{}"></span>{}</span>'.format(c, p)
        for p, c in RATING_COLORS)
    return (CSS
            + '<div class="bsch-wrap"><div class="bsch-head">' + head + "</div>"
            + '<div class="bsch-grid">' + "".join(cells) + "</div></div>"
            + '<div class="bsch-legend">' + legend
            + '<span style="margin-left:auto">■ 채움 = 수요예측일 &nbsp;/&nbsp; '
            + "⬚ 점선 = 발행일</span></div>")


def _noop_popover_change():
    """st.popover가 key로 열림/닫힘 상태를 추적하려면 on_change가 반드시 있어야 한다
    (Streamlit 공식 문서 명시 사항). 별도로 할 일은 없어서 빈 함수로 둔다."""
    pass


def _render_deal_detail_body(st, r, btn_key, popover_key=None):
    """수요예측 상세 내용(팝업/팝오버 공용). btn_key는 '이동' 버튼의 고유 key.
    popover_key를 주면 발행사명 오른쪽 가장자리에 닫기(X) 버튼이 뜨고, 누르면 그 팝오버를 닫는다.
    st.session_state[popover_key]는 위젯이 이미 그려진 뒤에는 직접 대입할 수 없어서
    (StreamlitAPIException), on_click 콜백 안에서 처리한다 — 콜백은 다음 스크립트
    실행 전에 미리 처리되므로 이 제약을 받지 않는다."""
    color = rating_color(r["신용등급"])

    if popover_key:
        def _close_popover():
            st.session_state[popover_key] = False

        name_col, xcol = st.columns([5, 1])
        with name_col:
            st.markdown(
                "<div style='font-size:1.15rem;font-weight:700;margin:0;line-height:1.3'>{}</div>".format(r["종목명"]),
                unsafe_allow_html=True)
        with xcol:
            st.button("✕", key="closepop_" + btn_key, width="stretch", on_click=_close_popover)
    else:
        st.markdown(
            "<div style='font-size:1.15rem;font-weight:700;margin:0;line-height:1.3'>{}</div>".format(r["종목명"]),
            unsafe_allow_html=True)

    st.markdown(
        "<span style='background:{};color:#fff;padding:2px 9px;border-radius:5px;"
        "font-size:.85rem'>{}</span>".format(color, r["신용등급"] or "-"),
        unsafe_allow_html=True)

    def _small_field(label, value):
        st.markdown(
            "<div style='font-size:.75rem;color:#64748b;margin-top:4px'>{}</div>"
            "<div style='font-size:1.05rem;font-weight:700'>{}</div>".format(
                label, value),
            unsafe_allow_html=True)

    c1, c2 = st.columns(2)
    with c1:
        _small_field("수요예측일", "{:%Y-%m-%d}".format(r["수요예측일"]) if has_date(r["수요예측일"]) else "미정")
        _small_field("신고발행금액", fmt_amt(r["신고발행금액"]))
    with c2:
        _small_field("발행일", "{:%Y-%m-%d}".format(r["발행일"]) if has_date(r["발행일"]) else "미정")
        _small_field("최대발행가능액", fmt_amt(r["최대발행가능액"]))

    fields = [("만기", r["만기"] or "-"), ("물량", r["물량"] or "-"), ("금리밴드", r["금리밴드"] or "밴드 미정"),
              ("대표주관", r["대표주관"] or "-"), ("상태", r["상태"] or "-"), ("출처", r["출처"] or "-")]
    if r.get("비고"):
        fields.append(("비고", r["비고"]))
    st.markdown(
        "<div style='font-size:.85rem;margin-top:6px'>"
        + "".join("<div style='margin-bottom:5px'><b>{}</b>: {}</div>".format(k, v) for k, v in fields)
        + "</div>",
        unsafe_allow_html=True)

    _, _btn_col, _ = st.columns([1, 2, 1])
    with _btn_col:
        if st.button("발행사별 상세보기로 이동", key=btn_key, width="stretch"):
            # 종목명의 (신종)/(후순위) 같은 괄호 표기는 떼고 순수 회사명만 넘긴다
            issuer_base = re.sub(r"\([^)]*\)", "", str(r["종목명"])).strip()
            st.session_state["nav_page"] = "발행사별 상세보기"
            st.session_state["bsch_target_issuer"] = issuer_base
            st.rerun()


WEEKDAYS_SUN_FIRST = ["일", "월", "화", "수", "목", "금", "토"]

_CAL_CSS = """
<style>
/* 날짜 칸: 테두리 없음, 정사각형에 가깝도록 최소 높이 부여 */
div[class*="st-key-calday_"] {
    border-radius: 0 !important;
    border: none !important;
    padding: 6px 8px !important;
    min-height: 100px !important;
}
/* 요일/날짜 칸이 나열되는 가로줄 자체의 칸 사이 간격 제거 */
[data-testid="stHorizontalBlock"] {
    gap: 0 !important;
}
/* 발행사 칩(팝오버 트리거) 버튼: 테두리 제거, 왼쪽정렬, 밀도 있게 */
div[class*="st-key-calpop_"] button {
    border: none !important;
    box-shadow: none !important;
    background: transparent !important;
    text-align: left !important;
    justify-content: flex-start !important;
    padding: 1px 4px !important;
    min-height: 1.5rem !important;
    font-size: .82rem !important;
}
div[class*="st-key-calpop_"] button p {
    text-align: left !important;
}
/* 팝오버 트리거들 사이 위아래 간격을 최대한 좁힘 (칸 내부 세로 블록 자체의 gap까지 제거) */
div[class*="st-key-calday_"] [data-testid="stVerticalBlock"] {
    gap: 0 !important;
}
div[class*="st-key-calday_"] div[data-testid="stPopover"],
div[class*="st-key-calday_"] div[data-testid="stElementContainer"] {
    margin-bottom: 0 !important;
}
/* 발행사 클릭 시 뜨는 상세 팝업 패널: 더 작게, 상단 여백도 최소화해서 밀도 있게 */
div[class*="st-key-calpop_"] div[data-testid="stPopoverBody"] {
    width: 200px !important;
    max-width: 200px !important;
    padding: 6px 10px 10px !important;
}
div[class*="st-key-calpop_"] div[data-testid="stPopoverBody"] div[data-testid="stVerticalBlock"] {
    gap: 2px !important;
}
div[class*="st-key-calpop_"] div[data-testid="stPopoverBody"] div[data-testid="stElementContainer"] {
    margin-bottom: 0 !important;
}
div[class*="st-key-calpop_"] div[data-testid="stPopoverBody"] h1,
div[class*="st-key-calpop_"] div[data-testid="stPopoverBody"] h2,
div[class*="st-key-calpop_"] div[data-testid="stPopoverBody"] h3,
div[class*="st-key-calpop_"] div[data-testid="stPopoverBody"] h4,
div[class*="st-key-calpop_"] div[data-testid="stPopoverBody"] p {
    margin: 0 !important;
    padding: 0 !important;
}
/* 팝업 안 닫기(X) 버튼: 테두리 제거 */
div[class*="st-key-closepop_"] button {
    border: none !important;
    box-shadow: none !important;
    background: transparent !important;
}
div[class*="st-key-closepop_"] button:hover {
    background: rgba(0,0,0,0.06) !important;
}
/* 이전/오늘/다음 버튼: 테두리 제거, 밀도 있게 */
div[class*="st-key-bsch_prev"] button,
div[class*="st-key-bsch_today"] button,
div[class*="st-key-bsch_next"] button {
    border: none !important;
    box-shadow: none !important;
    background: transparent !important;
}
div[class*="st-key-bsch_prev"] button:hover,
div[class*="st-key-bsch_today"] button:hover,
div[class*="st-key-bsch_next"] button:hover {
    background: rgba(49,51,63,0.06) !important;
}
/* 다가오는 일정 카드: 창이 커져도 일정 폭 이상 늘어나지 않도록 상한 지정 */
div[class*="st-key-updeal_"] {
    max-width: 720px !important;
}
</style>
"""


def render_month_calendar(st, df: pd.DataFrame, year: int, month: int, show_issue: bool):
    """st.columns + st.popover로 그리는 네이티브 캘린더.
    칩(버튼)을 클릭하면 새로고침 없이 그 옆에 상세 팝업이 뜬다.
    HTML 그리드(month_grid)보다는 성기지만, 클릭 상호작용이 필요해 이 방식을 쓴다.
    일요일이 첫 열이다. 날짜 칸은 고정 높이가 없어서, 일정이 많은 주는 그 줄 전체가
    자연스럽게(같은 줄의 7칸이 함께) 늘어난다 — Streamlit의 flex 레이아웃이 같은 줄
    칸들의 높이를 자동으로 맞춰주기 때문에 스크롤바 없이 해결된다.
    캘린더 전체를 key가 있는 컨테이너로 감싸서, 창이 아무리 넓어져도 가로 폭이
    무한정 늘어나지 않도록 CSS로 상한을 둔다."""
    st.markdown(_CAL_CSS, unsafe_allow_html=True)

    first = dt.date(year, month, 1)
    # 파이썬 weekday()는 월=0..일=6 이라, 일요일을 기준으로 한 주의 시작을 다시 계산한다
    start = first - dt.timedelta(days=(first.weekday() + 1) % 7)
    last = dt.date(year, month, calendar.monthrange(year, month)[1])
    end = last + dt.timedelta(days=(5 - last.weekday()) % 7)
    today = dt.date.today()

    by_day = {}
    for _, r in df.iterrows():
        if has_date(r["수요예측일"]):
            by_day.setdefault(r["수요예측일"], []).append(("demand", r))
        if show_issue and has_date(r["발행일"]):
            by_day.setdefault(r["발행일"], []).append(("issue", r))

    with st.container(key="calgrid_wrap"):
        head_cols = st.columns(7)
        for i, w in enumerate(WEEKDAYS_SUN_FIRST):
            color = "#dc2626" if i == 0 else "#2563eb" if i == 6 else "#475569"
            head_cols[i].markdown(
                "<div style='text-align:center;font-weight:600;font-size:.82rem;color:{}'>{}</div>".format(color, w),
                unsafe_allow_html=True)

        d = start
        while d <= end:
            cols = st.columns(7)
            for i in range(7):
                day = d + dt.timedelta(days=i)
                with cols[i]:
                    with st.container(border=True, key="calday_{}".format(day.isoformat())):
                        if day == today:
                            st.markdown(":orange[**{}**]".format(day.day))
                        elif day.month != month:
                            st.caption(str(day.day))
                        else:
                            st.markdown("**{}**".format(day.day))

                        events = sorted(
                            by_day.get(day, []),
                            key=lambda x: (x[0] != "demand", str(x[1]["종목명"]))
                        )
                        for kind, r in events:
                            emoji = rating_emoji(r["신용등급"])
                            amt = fmt_amt(r["최대발행가능액"])
                            nm = str(r["종목명"])[:8]
                            label = "{}{} {}".format("📤" if kind == "issue" else "", emoji, nm)
                            pop_key = "cal_{}_{}_{}".format(day.isoformat(), kind, r.name)
                            pop_key_full = "calpop_{}".format(pop_key)
                            with st.popover(label, width="stretch", key=pop_key_full, on_change=_noop_popover_change):
                                _render_deal_detail_body(st, r, btn_key="goto_" + pop_key, popover_key=pop_key_full)
            d += dt.timedelta(days=7)

    st.markdown(
        "<div style='display:flex;gap:14px;flex-wrap:wrap;font-size:.78rem;"
        "color:#475569;margin-top:8px'>"
        + "".join("<span>{} {}</span>".format(e, p) for p, e in RATING_EMOJI)
        + "<span style='margin-left:auto'>📤 = 발행일 (수요예측일과 구분)</span></div>",
        unsafe_allow_html=True)


def render_page(st, read_history):
    """Compass 뷰 페이지: 캘린더 + 목록 + 다가오는 일정."""
    st.header("수요예측 일정")

    df = read_deals(read_history)
    if df.empty:
        st.info("아직 데이터가 없습니다. 관리자 계정으로 **데이터 업로드 → 회사채 발행 리스트** "
                "에서 3사 파일을 올리거나, 로컬 수집 스크립트를 돌려주세요.")
        return

    today = dt.date.today()
    c1, c2, c3, c4 = st.columns([2, 2, 3, 3])
    status = c1.radio("상태", ["예정", "완료", "전체"], horizontal=True,
                      key="bsch_status")
    grades = sorted({str(g)[:3].rstrip("+-0") for g in df["신용등급"] if g}, reverse=True)
    pick_g = c2.multiselect("신용등급", grades, key="bsch_grade")
    srcs = sorted({s.strip() for v in df["출처"] for s in str(v).split(",") if s.strip()})
    pick_s = c3.multiselect("출처 증권사", srcs, key="bsch_src")
    query = c4.text_input("종목명 검색", key="bsch_q")

    view = df.copy()
    if status != "전체":
        view = view[view["상태"] == status]
    if pick_g:
        view = view[view["신용등급"].map(
            lambda x: any(str(x).upper().startswith(g) for g in pick_g))]
    if pick_s:
        view = view[view["출처"].map(lambda x: any(s in str(x) for s in pick_s))]
    if query:
        view = view[view["종목명"].str.contains(query, case=False, na=False)]

    wk_end = today + dt.timedelta(days=7)
    soon = view[view["수요예측일"].map(lambda d: has_date(d) and today <= d <= wk_end)]
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("표시 중인 딜", "{:,}건".format(len(view)))
    m2.metric("향후 7일 수요예측", "{}건".format(len(soon)))
    m3.metric("향후 7일 신고금액", fmt_amt(soon["신고발행금액"].sum()))
    m4.metric("향후 7일 최대발행", fmt_amt(soon["최대발행가능액"].sum()))

    tab_cal, tab_up = st.tabs(["🗓 캘린더", "⏭ 다가오는 일정"])

    with tab_cal:
        if "bsch_ym" not in st.session_state:
            st.session_state.bsch_ym = (today.year, today.month)
        y, m = st.session_state.bsch_ym
        n1, n2, n3, n4 = st.columns([1, 1, 1, 9])
        if n1.button("◀ 이전", key="bsch_prev", width="stretch"):
            st.session_state.bsch_ym = (y - 1, 12) if m == 1 else (y, m - 1)
            st.rerun()
        if n2.button("오늘", key="bsch_today", width="stretch"):
            st.session_state.bsch_ym = (today.year, today.month)
            st.rerun()
        if n3.button("다음 ▶", key="bsch_next", width="stretch"):
            st.session_state.bsch_ym = (y + 1, 1) if m == 12 else (y, m + 1)
            st.rerun()
        n4.subheader("{}년 {}월".format(y, m))
        cb1, cb2 = st.columns([1, 1])
        show_issue = cb1.checkbox("발행일도 표시", value=True, key="bsch_showissue")
        show_hybrid = cb2.checkbox("신종증권 표시", value=True, key="bsch_showhybrid")

        cal_view = view
        if not show_hybrid:
            cal_view = cal_view[~cal_view["종목명"].astype(str).str.contains("신종", na=False)]

        render_month_calendar(st, cal_view, y, m, show_issue)
        st.caption("칩을 클릭하면 새로고침 없이 옆에 상세정보가 뜹니다.")

    with tab_up:
        up = view[view["수요예측일"].map(lambda d: has_date(d) and d >= today)]
        up = up.sort_values("수요예측일")
        if up.empty:
            st.info("다가오는 수요예측 일정이 없습니다.")
        for d, grp in up.groupby("수요예측일"):
            dday = (d - today).days
            tag = "오늘" if dday == 0 else ("내일" if dday == 1 else "D-{}".format(dday))
            st.markdown("#### {:%m/%d} ({}) · {}".format(d, WEEKDAYS[d.weekday()], tag))
            for _, r in grp.iterrows():
                color = rating_color(r["신용등급"])
                with st.container(border=True, key="updeal_{}".format(r.name)):
                    a, b = st.columns([3, 5])
                    a.markdown(
                        "**{}** &nbsp;<span style='background:{};color:#fff;padding:1px 7px;"
                        "border-radius:4px;font-size:.78rem'>{}</span><br>"
                        "<span style='color:#64748b;font-size:.85rem'>만기 {} · 발행일 {}"
                        "</span>".format(
                            r["종목명"], color, r["신용등급"], r["만기"] or "-",
                            r["발행일"] if has_date(r["발행일"]) else "미정"),
                        unsafe_allow_html=True)
                    b.markdown(
                        "신고 **{}** → 최대 **{}** · 물량 {}<br>"
                        "<span style='color:#64748b;font-size:.85rem'>{} · 주관 {}</span>".format(
                            fmt_amt(r["신고발행금액"]), fmt_amt(r["최대발행가능액"]),
                            r["물량"] or "-", r["금리밴드"] or "밴드 미정",
                            r["대표주관"] or "-"),
                        unsafe_allow_html=True)


def render_upload_tab(st, read_history, get_client, sheet_id, alias_map=None):
    """Compass '데이터 업로드' 페이지의 탭 하나로 들어가는 업로더."""
    st.caption("한국투자증권 · 하나증권 PDF, NH투자증권 XLSX 를 한 번에 올리면 "
               "같은 딜끼리 합쳐서 저장합니다. 매일 아침 로컬 수집 스크립트가 "
               "자동으로 넣어주므로, 여기서는 놓친 날을 메울 때만 쓰면 됩니다.")

    files = st.file_uploader(
        "3사 발행 리스트 파일", type=["pdf", "xlsx", "xlsm"],
        accept_multiple_files=True, key="bsch_upload")
    if not files:
        return

    parsed, log = parse_uploads(files, alias_map)
    st.write("**파일 인식 결과**")
    st.dataframe(pd.DataFrame(log, columns=["파일", "증권사", "결과"]),
                 width="stretch", hide_index=True)

    if parsed.empty:
        st.error("읽어낸 딜이 없습니다. 파일 형식을 확인해주세요.")
        return

    st.success("{}건의 딜을 인식했습니다 (예정 {}건).".format(
        len(parsed), int((parsed["상태"] == "예정").sum())))
    st.dataframe(parsed, width="stretch", hide_index=True, height=320)

    if st.button("Google Sheets 에 반영", type="primary", key="bsch_save"):
        with st.spinner("기존 데이터와 합쳐 저장하는 중..."):
            merged = merge_into_sheet(parsed, read_history, get_client,
                                      sheet_id, alias_map)
        st.success("저장 완료 — 전체 {}건".format(len(merged)))
        st.cache_data.clear()
