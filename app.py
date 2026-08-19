import streamlit as st
import pandas as pd
import altair as alt
import requests
import zipfile
import io
import re
import datetime
import xml.etree.ElementTree as ET
import gspread
from google.oauth2.service_account import Credentials

st.set_page_config(page_title="Compass", layout="wide")
st.title("🧭 Compass")

DART_API_KEY = st.secrets.get("DART_API_KEY", "")


NAV_VIEW_PAGES = ["채권 스프레드", "신용등급 트리거", "익스포저", "발행사별 상세보기"]
NAV_ADMIN_PAGES = ["데이터 업로드", "관리자 설정"]

if "nav_page" not in st.session_state:
    st.session_state.nav_page = NAV_VIEW_PAGES[0]
if "is_admin" not in st.session_state:
    st.session_state.is_admin = False

with st.sidebar:
    for _p in NAV_VIEW_PAGES:
        if st.button(_p, key=f"nav_btn_{_p}", use_container_width=True,
                     type="primary" if st.session_state.nav_page == _p else "secondary"):
            st.session_state.nav_page = _p
            st.rerun()

    st.markdown("---")

    if st.session_state.is_admin:
        for _p in NAV_ADMIN_PAGES:
            if st.button(_p, key=f"nav_btn_{_p}", use_container_width=True,
                         type="primary" if st.session_state.nav_page == _p else "secondary"):
                st.session_state.nav_page = _p
                st.rerun()
        st.markdown("---")
        if st.button("로그아웃", key="admin_logout_btn", use_container_width=True):
            st.session_state.is_admin = False
            if st.session_state.nav_page in NAV_ADMIN_PAGES:
                st.session_state.nav_page = NAV_VIEW_PAGES[0]
            st.rerun()
    else:
        with st.expander("관리자 로그인"):
            _admin_pw_secret = st.secrets.get("ADMIN_PASSWORD", "")
            _pw_input = st.text_input("비밀번호", type="password", key="sidebar_admin_pw")
            if st.button("로그인", key="admin_login_btn", use_container_width=True):
                if not _admin_pw_secret:
                    st.error("관리자 비밀번호가 설정되어 있지 않습니다.")
                elif _pw_input == _admin_pw_secret:
                    st.session_state.is_admin = True
                    st.rerun()
                else:
                    st.error("비밀번호가 올바르지 않습니다.")

page = st.session_state.nav_page


# ==============================================================
# 공통 유틸
# ==============================================================

GOOGLE_SHEETS_SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

@st.cache_resource
def get_gsheet_client():
    """구글 서비스 계정으로 인증된 gspread 클라이언트 반환 (세션 내 재사용)."""
    creds_dict = dict(st.secrets["gcp_service_account"])
    creds = Credentials.from_service_account_info(creds_dict, scopes=GOOGLE_SHEETS_SCOPES)
    return gspread.authorize(creds)


def append_history(df: pd.DataFrame, worksheet_name: str, date_str: str, date_col: str = "업데이트일자"):
    """df를 지정한 워크시트에 날짜열과 함께 이력으로 쌓는다.
    같은 날짜(date_str) 데이터가 이미 있으면 먼저 지우고 새로 씀 (재업로드시 중복 방지)."""
    client = get_gsheet_client()
    sh = client.open_by_key(st.secrets["GOOGLE_SHEET_ID"])
    try:
        ws = sh.worksheet(worksheet_name)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=worksheet_name, rows=2000, cols=max(30, len(df.columns) + 1))
        ws.append_row([date_col] + [str(c) for c in df.columns])

    existing = ws.get_all_values()
    if existing:
        header = existing[0]
        data_rows = existing[1:]
        keep_rows = [row for row in data_rows if row and row[0] != date_str]
        if len(keep_rows) != len(data_rows):
            ws.clear()
            ws.append_row(header)
            if keep_rows:
                ws.append_rows(keep_rows)
    else:
        ws.append_row([date_col] + [str(c) for c in df.columns])

    new_rows = [[date_str] + [("" if pd.isna(v) else str(v)) for v in row] for row in df.values.tolist()]
    if new_rows:
        ws.append_rows(new_rows)


def extract_date_from_filename(filename: str):
    """파일명에서 8자리 날짜(YYYYMMDD)를 찾아 반환, 없으면 오늘 날짜."""
    m = re.search(r'(20\d{6})', filename or "")
    if m:
        return m.group(1)
    return datetime.date.today().strftime("%Y%m%d")


def append_history_multi(df: pd.DataFrame, worksheet_name: str, dedup_cols: list):
    """df를 워크시트에 추가. dedup_cols 조합이 새 데이터에 이미 존재하는 기존 행은 먼저 지우고 새로 씀.
    시트의 기존 열 구성과 df의 열 구성이 다르면(예: 새 열이 추가된 경우), 열을 합쳐서
    기존 행 전체를 새 구조로 재정렬(마이그레이션)한 뒤 다시 쓴다 — 열이 밀려서 어긋나는 것을 방지."""
    client = get_gsheet_client()
    sh = client.open_by_key(st.secrets["GOOGLE_SHEET_ID"])
    new_cols = [str(c) for c in df.columns]

    try:
        ws = sh.worksheet(worksheet_name)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=worksheet_name, rows=3000, cols=max(30, len(df.columns)))
        ws.append_row(new_cols)
        new_rows = [[("" if pd.isna(v) else str(v)) for v in row] for row in df.values.tolist()]
        if new_rows:
            ws.append_rows(new_rows)
        return

    existing = ws.get_all_values()
    if not existing:
        ws.append_row(new_cols)
        existing_header, existing_rows = new_cols, []
    else:
        existing_header, existing_rows = existing[0], existing[1:]

    # 기존 헤더에 없는 새 열이 있으면 합쳐서 통합 헤더를 만든다 (기존 열 순서는 유지, 새 열은 뒤에 추가)
    union_header = existing_header + [c for c in new_cols if c not in existing_header]

    if union_header != existing_header:
        migrated_rows = []
        for row in existing_rows:
            row_map = {existing_header[i]: (row[i] if i < len(row) else "") for i in range(len(existing_header))}
            migrated_rows.append([row_map.get(c, "") for c in union_header])
        existing_rows = migrated_rows

    header = union_header

    if all(c in header for c in dedup_cols):
        col_idx = [header.index(c) for c in dedup_cols]
        new_keys = set(
            tuple(("" if pd.isna(v) else str(v)) for v in row)
            for row in df[dedup_cols].values.tolist()
        )
        existing_rows = [
            row for row in existing_rows
            if tuple(row[i] if i < len(row) else "" for i in col_idx) not in new_keys
        ]

    df_reordered = df.reindex(columns=header)
    new_rows = [
        [("" if pd.isna(v) else str(v)) for v in row]
        for row in df_reordered.values.tolist()
    ]

    ws.clear()
    ws.append_row(header)
    if existing_rows:
        ws.append_rows(existing_rows)
    if new_rows:
        ws.append_rows(new_rows)


@st.cache_data(ttl=60)
def read_history(worksheet_name: str) -> pd.DataFrame:
    """이력 워크시트를 읽어 DataFrame으로 반환. 시트가 없으면 빈 DataFrame.
    시트가 실제 열 수보다 넓게 만들어져 뒤쪽에 제목 없는 빈 열이 남아있어도
    (gspread의 get_all_records는 이 경우 '중복 헤더' 오류를 내므로) 직접 파싱해서 우회한다."""
    client = get_gsheet_client()
    sh = client.open_by_key(st.secrets["GOOGLE_SHEET_ID"])
    try:
        ws = sh.worksheet(worksheet_name)
    except gspread.WorksheetNotFound:
        return pd.DataFrame()

    values = ws.get_all_values()
    if not values:
        return pd.DataFrame()

    header = values[0]
    while header and header[-1] == "":
        header = header[:-1]
    if not header:
        return pd.DataFrame()

    n = len(header)
    data_rows = [row[:n] + [""] * (n - len(row)) for row in values[1:]]
    return pd.DataFrame(data_rows, columns=header)


def extract_effective_date(file_obj, sheet_name=None, max_rows=30):
    """엑셀 파일 상단 텍스트에서 '기준일/기준시점/유효시점' 뒤에 오는 날짜를 찾아 YYYYMMDD로 반환."""
    try:
        file_obj.seek(0)
    except Exception:
        pass
    try:
        raw = pd.read_excel(file_obj, sheet_name=sheet_name, header=None, nrows=max_rows)
    except Exception:
        return None
    finally:
        try:
            file_obj.seek(0)
        except Exception:
            pass
    keywords = ["기준일", "기준시점", "유효시점"]
    for i in range(len(raw)):
        for v in raw.iloc[i].dropna():
            text = str(v)
            if any(k in text for k in keywords):
                m = re.search(r'(\d{4})[.\-](\d{1,2})[.\-](\d{1,2})', text)
                if m:
                    y, mo, d = m.groups()
                    return f"{y}{int(mo):02d}{int(d):02d}"
                m = re.search(r'(\d{4})년\s*(\d{1,2})월\s*(\d{1,2})일', text)
                if m:
                    y, mo, d = m.groups()
                    return f"{y}{int(mo):02d}{int(d):02d}"
    return None


# 발행사명 표기 차이 매핑 — 이제 코드가 아니라 Google Sheets의 '발행사별칭' 탭에서 관리합니다.
# 추가/수정/삭제는 대시보드의 '관리자 설정' 탭(비밀번호 필요)에서만 가능합니다.
# 구조: 회사 하나가 한 행, 소스(인포맥스/한신평/나신평/한기평/위너스/DART)가 각각 열.
ALIAS_SHEET_NAME = "발행사별칭"

ALIAS_SOURCE_COLUMNS = ["한국신용평가", "나이스신용평가", "한국기업평가", "인포맥스", "위너스(WINUS)", "DART", "기타"]
ALIAS_EXTRA_COLUMNS = ["법인고유번호"]  # 이름이 아닌 코드값 — 발행사명 별칭 로직에는 포함하지 않는다
ALIAS_HEADER = ["표준명"] + ALIAS_SOURCE_COLUMNS + ALIAS_EXTRA_COLUMNS


def _migrate_alias_rows(old_header: list, old_rows: list) -> list:
    """예전 형식(별칭/표준명 2열, 또는 별칭/표준명/출처 3열 - 세로형)이나
    열이 일부 빠진 예전 가로형을 지금의 가로형(표준명 + 소스별 열 + 법인고유번호)으로 변환.
    이미 완전히 같은 형식이면 그대로 통과."""
    if old_header == ALIAS_HEADER:
        return old_rows

    if "별칭" in old_header and "표준명" in old_header:
        # 세로형(별칭,표준명[,출처]) -> 가로형 변환
        alias_idx = old_header.index("별칭")
        canon_idx = old_header.index("표준명")
        source_idx = old_header.index("출처") if "출처" in old_header else None

        companies = {}  # 표준명 -> {source_col: alias}
        for row in old_rows:
            alias = row[alias_idx].strip() if alias_idx < len(row) else ""
            canon = row[canon_idx].strip() if canon_idx < len(row) else ""
            if not alias or not canon:
                continue
            source = (row[source_idx].strip() if source_idx is not None and source_idx < len(row) else "")
            col = source if source in ALIAS_SOURCE_COLUMNS else "기타"
            companies.setdefault(canon, {})[col] = alias

        new_rows = []
        for canon, src_map in companies.items():
            new_rows.append([canon] + [src_map.get(c, "") for c in ALIAS_SOURCE_COLUMNS] + [""] * len(ALIAS_EXTRA_COLUMNS))
        return new_rows

    if "표준명" in old_header:
        # 이미 가로형인데 열 구성만 다른 경우(예: 법인고유번호 열이 새로 추가된 경우)
        # 열 이름 기준으로 재정렬하고, 없는 열은 빈 값으로 채운다 — 기존 데이터는 보존.
        migrated = []
        for row in old_rows:
            row_map = {old_header[i]: (row[i] if i < len(row) else "") for i in range(len(old_header))}
            migrated.append([row_map.get(c, "") for c in ALIAS_HEADER])
        return migrated

    # 표준명 열조차 없는 완전히 알 수 없는 형식
    return []


def _get_or_migrate_alias_ws():
    """별칭 워크시트를 열고, 예전 형식이면 가로형으로 자동 마이그레이션."""
    client = get_gsheet_client()
    sh = client.open_by_key(st.secrets["GOOGLE_SHEET_ID"])
    try:
        ws = sh.worksheet(ALIAS_SHEET_NAME)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=ALIAS_SHEET_NAME, rows=1000, cols=len(ALIAS_HEADER))
        ws.append_row(ALIAS_HEADER)
        return ws

    values = ws.get_all_values()
    if not values:
        ws.append_row(ALIAS_HEADER)
        return ws

    header = values[0]
    while header and header[-1] == "":
        header = header[:-1]

    if header != ALIAS_HEADER:
        migrated = _migrate_alias_rows(header, values[1:])
        ws.clear()
        ws.append_row(ALIAS_HEADER)
        if migrated:
            ws.append_rows(migrated)

    return ws


@st.cache_data(ttl=300)
def load_issuer_aliases_full() -> pd.DataFrame:
    """별칭 시트를 가로형 표(표준명 + 소스별 열) 그대로 반환."""
    try:
        client = get_gsheet_client()
        sh = client.open_by_key(st.secrets["GOOGLE_SHEET_ID"])
        try:
            ws = sh.worksheet(ALIAS_SHEET_NAME)
        except gspread.WorksheetNotFound:
            return pd.DataFrame(columns=ALIAS_HEADER)
        values = ws.get_all_values()
        if not values:
            return pd.DataFrame(columns=ALIAS_HEADER)
        header = values[0]
        while header and header[-1] == "":
            header = header[:-1]
        if not header:
            return pd.DataFrame(columns=ALIAS_HEADER)
        n = len(header)
        data_rows = [row[:n] + [""] * (n - len(row)) for row in values[1:]]
        df = pd.DataFrame(data_rows, columns=header)
        if header != ALIAS_HEADER:
            migrated = _migrate_alias_rows(header, data_rows)
            df = pd.DataFrame(migrated, columns=ALIAS_HEADER) if migrated else pd.DataFrame(columns=ALIAS_HEADER)
        return df
    except Exception:
        return pd.DataFrame(columns=ALIAS_HEADER)


@st.cache_data(ttl=300)
def load_issuer_aliases() -> dict:
    """가로형 별칭 표를 {별칭: 표준명} 딕셔너리로 펼쳐서 반환 (매칭 로직에서 사용)."""
    df = load_issuer_aliases_full()
    result = {}
    if df.empty or "표준명" not in df.columns:
        return result
    source_cols = [c for c in df.columns if c != "표준명" and c not in ALIAS_EXTRA_COLUMNS]
    for _, row in df.iterrows():
        canon = str(row.get("표준명", "")).strip()
        if not canon:
            continue
        for col in source_cols:
            val = str(row.get(col, "")).strip()
            if val and val != canon:
                result[val] = canon
    return result


def upsert_issuer_alias(canonical: str, source_col: str, alias: str):
    """표준명 행이 있으면 해당 소스 칸만 갱신, 없으면 새 행 생성."""
    ws = _get_or_migrate_alias_ws()
    values = ws.get_all_values()
    header = values[0] if values else ALIAS_HEADER
    canon_idx = header.index("표준명")
    src_idx = header.index(source_col) if source_col in header else None

    for row_num, row in enumerate(values[1:], start=2):
        if canon_idx < len(row) and row[canon_idx].strip() == canonical.strip():
            if src_idx is not None:
                ws.update_cell(row_num, src_idx + 1, alias)
            load_issuer_aliases.clear()
            load_issuer_aliases_full.clear()
            return

    new_row = [canonical] + ["" for _ in ALIAS_SOURCE_COLUMNS] + ["" for _ in ALIAS_EXTRA_COLUMNS]
    if source_col in ALIAS_SOURCE_COLUMNS:
        new_row[1 + ALIAS_SOURCE_COLUMNS.index(source_col)] = alias
    ws.append_row(new_row)
    load_issuer_aliases.clear()
    load_issuer_aliases_full.clear()


def upsert_issuer_aliases_bulk(pairs: list, source_col: str):
    """[(별칭, 표준명), ...] 목록을 한 번에 반영 (있으면 갱신, 없으면 새 행)."""
    if not pairs:
        return
    ws = _get_or_migrate_alias_ws()
    values = ws.get_all_values()
    header = values[0] if values else ALIAS_HEADER
    canon_idx = header.index("표준명")
    src_idx = header.index(source_col) if source_col in header else None

    canon_to_row = {}
    for row_num, row in enumerate(values[1:], start=2):
        if canon_idx < len(row):
            canon_to_row[row[canon_idx].strip()] = row_num

    new_rows = []
    for alias, canonical in pairs:
        canonical = canonical.strip()
        if canonical in canon_to_row and src_idx is not None:
            ws.update_cell(canon_to_row[canonical], src_idx + 1, alias)
        else:
            new_row = [canonical] + ["" for _ in ALIAS_SOURCE_COLUMNS] + ["" for _ in ALIAS_EXTRA_COLUMNS]
            if source_col in ALIAS_SOURCE_COLUMNS:
                new_row[1 + ALIAS_SOURCE_COLUMNS.index(source_col)] = alias
            new_rows.append(new_row)
            canon_to_row[canonical] = None  # 중복 등록 방지(같은 배치 내)

    if new_rows:
        ws.append_rows(new_rows)
    load_issuer_aliases.clear()
    load_issuer_aliases_full.clear()


def delete_issuer_company(canonical: str):
    """표준명 행 전체를 삭제."""
    ws = _get_or_migrate_alias_ws()
    values = ws.get_all_values()
    if not values:
        return
    header = values[0]
    if "표준명" not in header:
        return
    canon_idx = header.index("표준명")
    for row_num, row in enumerate(values[1:], start=2):
        if canon_idx < len(row) and row[canon_idx].strip() == canonical.strip():
            ws.delete_rows(row_num)
            break
    load_issuer_aliases.clear()
    load_issuer_aliases_full.clear()


def build_alias_excel_bytes(df: pd.DataFrame) -> bytes:
    """별칭 표를 엑셀 파일 바이트로 변환 (다운로드용)."""
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="발행사별칭")
    return buf.getvalue()


def save_alias_excel_upload(df: pd.DataFrame):
    """업로드된 엑셀(표준명 + 소스별 열)로 별칭 시트를 통째로 덮어쓴다."""
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    if "표준명" not in df.columns:
        raise ValueError("업로드한 엑셀에 '표준명' 열이 없습니다.")
    for col in ALIAS_SOURCE_COLUMNS + ALIAS_EXTRA_COLUMNS:
        if col not in df.columns:
            df[col] = ""
    df = df[ALIAS_HEADER]
    df = df.fillna("")
    df = df[df["표준명"].astype(str).str.strip() != ""]

    client = get_gsheet_client()
    sh = client.open_by_key(st.secrets["GOOGLE_SHEET_ID"])
    try:
        ws = sh.worksheet(ALIAS_SHEET_NAME)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=ALIAS_SHEET_NAME, rows=max(1000, len(df) + 10), cols=len(ALIAS_HEADER))

    ws.clear()
    ws.append_row(ALIAS_HEADER)
    rows = [[str(v) for v in row] for row in df.values.tolist()]
    if rows:
        ws.append_rows(rows)
    load_issuer_aliases.clear()
    load_issuer_aliases_full.clear()


def validate_alias_table(df: pd.DataFrame) -> dict:
    """별칭 표에서 흔한 문제 3가지를 점검: 표준명 중복행 / 별칭 충돌 / 표준명에 남은 법인접미어."""
    result = {"dup_rows": pd.DataFrame(), "conflicts": [], "suffix_rows": pd.DataFrame()}
    if df.empty or "표준명" not in df.columns:
        return result

    source_cols = [c for c in df.columns if c != "표준명" and c not in ALIAS_EXTRA_COLUMNS]

    # 1) 표준명 완전 중복행
    dup_mask = df["표준명"].astype(str).str.strip().duplicated(keep=False) & (df["표준명"].astype(str).str.strip() != "")
    result["dup_rows"] = df[dup_mask].sort_values("표준명")

    # 2) 같은 표기가 서로 다른 표준명에 매핑된 충돌
    alias_to_canons = {}
    for _, row in df.iterrows():
        canon = str(row.get("표준명", "")).strip()
        if not canon:
            continue
        for col in source_cols:
            val = row.get(col)
            if pd.isna(val):
                continue
            val = str(val).strip()
            if not val:
                continue
            alias_to_canons.setdefault(val, set()).add(canon)
    result["conflicts"] = [(a, sorted(c)) for a, c in alias_to_canons.items() if len(c) > 1]

    # 3) 표준명 자체에 법인접미어(주식회사/(주)/㈜)가 남아있는 행
    suffix_mask = df["표준명"].astype(str).str.contains("주식회사|\\(주\\)|㈜", na=False, regex=True)
    result["suffix_rows"] = df[suffix_mask]

    return result


def build_cleaned_alias_table(df: pd.DataFrame) -> pd.DataFrame:
    """표준명 법인접미어 제거 + 중복행 병합(각 열의 첫 non-null 값을 채택)한 정리안을 만든다."""
    if df.empty or "표준명" not in df.columns:
        return df

    cleaned = df.copy()
    for suf in ["주식회사", "(주)", "㈜"]:
        cleaned["표준명"] = cleaned["표준명"].astype(str).str.replace(suf, "", regex=False)
    cleaned["표준명"] = cleaned["표준명"].str.strip()
    cleaned = cleaned[cleaned["표준명"] != ""]

    def first_non_blank(series):
        vals = [str(v).strip() for v in series if pd.notna(v) and str(v).strip() != ""]
        return vals[0] if vals else ""

    other_cols = [c for c in cleaned.columns if c != "표준명"]
    merged = cleaned.groupby("표준명", as_index=False).agg({c: first_non_blank for c in other_cols})
    return merged[ALIAS_HEADER] if all(c in merged.columns for c in ALIAS_HEADER) else merged



@st.cache_data(ttl=60 * 60 * 24)
def load_corp_code_map(api_key: str) -> dict:
    """DART Open API에서 전체 기업 고유번호 목록을 받아 {회사명: corp_code} 딕셔너리로 반환. 24시간 캐시."""
    url = f"https://opendart.fss.or.kr/api/corpCode.xml?crtfc_key={api_key}"
    resp = requests.get(url, timeout=30)
    resp.raise_for_status()
    zf = zipfile.ZipFile(io.BytesIO(resp.content))
    xml_bytes = zf.read(zf.namelist()[0])
    root = ET.fromstring(xml_bytes)
    mapping = {}
    for item in root.findall("list"):
        corp_name = (item.findtext("corp_name") or "").strip()
        corp_code = (item.findtext("corp_code") or "").strip()
        if corp_name and corp_code:
            mapping[corp_name] = corp_code
    return mapping


def find_corp_code(issuer_name: str, corp_map: dict):
    """발행사명으로 DART corp_code를 찾는다 (정확일치 → 접미어 제거 → 부분일치 순).
    부분일치 단계에서는 너무 짧은(2글자 미만) DART 등록명은 후보에서 제외한다
    ('은' 한 글자짜리 등록명이 '~은행'을 전부 잘못 잡아채는 것 같은 오탐을 막기 위함)."""
    if issuer_name in corp_map:
        return corp_map[issuer_name]
    for suffix in ["지주", "홀딩스", "㈜", "(주)"]:
        candidate = issuer_name.replace(suffix, "").strip()
        if candidate in corp_map:
            return corp_map[candidate]
    candidates = [
        name for name in corp_map
        if len(name) >= 3 and (issuer_name in name or name in issuer_name)
    ]
    if len(candidates) == 1:
        return corp_map[candidates[0]]
    return None


def find_dart_name(issuer_name: str, corp_map: dict):
    """발행사명으로 DART상의 정식 회사명을 찾는다 (find_corp_code와 같은 매칭 규칙, 결과는 이름)."""
    if pd.isna(issuer_name):
        return None
    issuer_name = str(issuer_name).strip()
    if issuer_name in corp_map:
        return issuer_name
    for suffix in ["지주", "홀딩스", "㈜", "(주)"]:
        candidate = issuer_name.replace(suffix, "").strip()
        if candidate in corp_map:
            return candidate
    candidates = [
        name for name in corp_map
        if len(name) >= 3 and (issuer_name in name or name in issuer_name)
    ]
    if len(candidates) == 1:
        return candidates[0]
    return None


@st.cache_data(ttl=60 * 60 * 24)
def fetch_financials(corp_code: str, api_key: str):
    """corp_code로 가장 최근 재무제표(매출액/영업이익/당기순이익)를 조회. 최신 분기→반기→1분기→사업보고서 순으로 시도."""
    if not corp_code:
        return {"매출액": None, "영업이익": None, "당기순이익": None, "재무제표기준일": "매칭 실패"}
    this_year = datetime.date.today().year
    attempts = []
    for year in [this_year, this_year - 1]:
        attempts += [
            (year, "11014", f"{year} 3분기보고서"),
            (year, "11012", f"{year} 반기보고서"),
            (year, "11013", f"{year} 1분기보고서"),
            (year, "11011", f"{year} 사업보고서"),
        ]
    url = "https://opendart.fss.or.kr/api/fnlttSinglAcntAll.json"
    for year, reprt_code, label in attempts:
        for fs_div in ["CFS", "OFS"]:
            params = {"crtfc_key": api_key, "corp_code": corp_code,
                      "bsns_year": str(year), "reprt_code": reprt_code, "fs_div": fs_div}
            try:
                resp = requests.get(url, params=params, timeout=15)
                data = resp.json()
            except Exception:
                continue
            if data.get("status") != "000":
                continue
            revenue = op_income = net_income = None
            for row in data.get("list", []):
                if row.get("sj_div") not in ("IS", "CIS"):
                    continue
                name = row.get("account_nm", "")
                raw = (row.get("thstrm_amount") or "").replace(",", "")
                try:
                    amt = int(raw)
                except ValueError:
                    continue
                if name in ("매출액", "수익(매출액)", "영업수익"):
                    revenue = amt
                elif name == "영업이익":
                    op_income = amt
                elif name in ("당기순이익", "당기순이익(손실)"):
                    net_income = amt
            if revenue is not None or op_income is not None or net_income is not None:
                return {"매출액": revenue, "영업이익": op_income, "당기순이익": net_income, "재무제표기준일": label}
    return {"매출액": None, "영업이익": None, "당기순이익": None, "재무제표기준일": "조회 실패"}


@st.cache_data(ttl=60 * 60 * 24)
def fetch_company_overview(corp_code: str, api_key: str):
    """DART 기업개황(company.json) API로 회사 기본 정보를 조회."""
    if not corp_code:
        return None
    url = "https://opendart.fss.or.kr/api/company.json"
    params = {"crtfc_key": api_key, "corp_code": corp_code}
    try:
        resp = requests.get(url, params=params, timeout=15)
        data = resp.json()
    except Exception:
        return None
    if data.get("status") != "000":
        return None
    return data


@st.cache_data(ttl=60 * 60 * 24)
def fetch_financials_detail(corp_code: str, api_key: str):
    """'발행사별 상세보기' 전용 상세 재무제표 조회.
    매출액/영업이익/당기순이익 각각에 대해 당기(해당 분/반기)·당기누적·전년동기·전년동기누적을 함께 담아 반환.
    DART API 응답의 thstrm_amount(당기)/thstrm_add_amount(당기누적)/
    frmtrm_q_amount(전년동기)/frmtrm_add_amount(전년동기누적) 필드를 그대로 활용한다."""
    if not corp_code:
        return None

    target_accounts = {
        "매출액": ["매출액", "수익(매출액)", "영업수익"],
        "영업이익": ["영업이익", "영업이익(손실)"],
        "당기순이익": ["당기순이익", "당기순이익(손실)"],
    }

    def to_num(x):
        try:
            return int(str(x).replace(",", ""))
        except (ValueError, TypeError):
            return None

    this_year = datetime.date.today().year
    attempts = []
    for year in [this_year, this_year - 1]:
        attempts += [
            (year, "11014", f"{year} 3분기보고서"),
            (year, "11012", f"{year} 반기보고서"),
            (year, "11013", f"{year} 1분기보고서"),
            (year, "11011", f"{year} 사업보고서"),
        ]

    url = "https://opendart.fss.or.kr/api/fnlttSinglAcntAll.json"
    for year, reprt_code, label in attempts:
        for fs_div in ["CFS", "OFS"]:
            params = {"crtfc_key": api_key, "corp_code": corp_code,
                      "bsns_year": str(year), "reprt_code": reprt_code, "fs_div": fs_div}
            try:
                resp = requests.get(url, params=params, timeout=15)
                data = resp.json()
            except Exception:
                continue
            if data.get("status") != "000":
                continue

            result = {}
            for row in data.get("list", []):
                if row.get("sj_div") not in ("IS", "CIS"):
                    continue
                name = row.get("account_nm", "")
                for metric, names in target_accounts.items():
                    if metric in result or name not in names:
                        continue
                    result[metric] = {
                        "당기": to_num(row.get("thstrm_amount")),
                        "당기누적": to_num(row.get("thstrm_add_amount")),
                        "전년동기": to_num(row.get("frmtrm_q_amount")) or to_num(row.get("frmtrm_amount")),
                        "전년동기누적": to_num(row.get("frmtrm_add_amount")),
                    }
            if result:
                return {"기준": label, "fs_div": fs_div, "항목": result}
    return None


# ------------------------------------------------------------
# 재무상태표(BS) + 연도별/분기별 손익계산서(IS) 종합 조회
# ------------------------------------------------------------
BS_ACCOUNTS = {"총자산": ["자산총계"], "총부채": ["부채총계"], "자기자본": ["자본총계"]}
IS_ACCOUNTS = {
    "매출액": ["매출액", "수익(매출액)", "영업수익"],
    "영업이익": ["영업이익", "영업이익(손실)"],
    "당기순이익": ["당기순이익", "당기순이익(손실)"],
}

UNIT_DIVISORS = {"원": 1, "천원": 1_000, "백만원": 1_000_000, "억원": 100_000_000,
                  "십억원": 1_000_000_000, "조원": 1_000_000_000_000}


def fmt_unit(x, unit):
    """금액을 선택한 단위로 환산해서 콤마 포함 문자열로 변환."""
    if x is None:
        return "-"
    divisor = UNIT_DIVISORS.get(unit, 1)
    val = x / divisor
    return f"{val:,.0f}" if divisor == 1 else f"{val:,.1f}"


def fmt_unit_change(cur, prev, unit):
    """증감(선택 단위)과 증감률(%)을 함께 반환."""
    if cur is None or prev is None or prev == 0:
        return "-", "-"
    diff = cur - prev
    pct = diff / abs(prev) * 100
    sign = "+" if diff >= 0 else ""
    divisor = UNIT_DIVISORS.get(unit, 1)
    diff_val = diff / divisor
    diff_str = f"{sign}{diff_val:,.0f}" if divisor == 1 else f"{sign}{diff_val:,.1f}"
    return diff_str, f"{sign}{pct:.1f}%"


def _to_num(x):
    try:
        return int(str(x).replace(",", ""))
    except (ValueError, TypeError):
        return None


@st.cache_data(ttl=60 * 60 * 24)
def _fetch_dart_report(corp_code: str, api_key: str, year: int, reprt_code: str):
    """단일 (연도, 보고서코드) 조합의 재무제표 원본 목록을 가져온다.
    연결재무제표(CFS) 우선, 없으면 개별재무제표(OFS). 실패 시 (None, None)."""
    url = "https://opendart.fss.or.kr/api/fnlttSinglAcntAll.json"
    for fs_div in ["CFS", "OFS"]:
        params = {"crtfc_key": api_key, "corp_code": corp_code,
                  "bsns_year": str(year), "reprt_code": reprt_code, "fs_div": fs_div}
        try:
            resp = requests.get(url, params=params, timeout=15)
            data = resp.json()
        except Exception:
            continue
        if data.get("status") == "000" and data.get("list"):
            return data["list"], fs_div
    return None, None


def _extract_bs_values(rows):
    result = {}
    for row in rows or []:
        if row.get("sj_div") != "BS":
            continue
        name = row.get("account_nm", "")
        for metric, names in BS_ACCOUNTS.items():
            if metric in result or name not in names:
                continue
            result[metric] = {
                "당기말": _to_num(row.get("thstrm_amount")),
                "전기말": _to_num(row.get("frmtrm_amount")),
                "전전기말": _to_num(row.get("bfefrmtrm_amount")),
            }
    return result


def _extract_is_values(rows):
    result = {}
    for row in rows or []:
        if row.get("sj_div") not in ("IS", "CIS"):
            continue
        name = row.get("account_nm", "")
        for metric, names in IS_ACCOUNTS.items():
            if metric in result or name not in names:
                continue
            result[metric] = {
                "당기": _to_num(row.get("thstrm_amount")),
                "당기누적": _to_num(row.get("thstrm_add_amount")),
                "전기": _to_num(row.get("frmtrm_amount")),
                "전전기": _to_num(row.get("bfefrmtrm_amount")),
                "전기누적": _to_num(row.get("frmtrm_add_amount")),
            }
    return result


@st.cache_data(ttl=60 * 60 * 24)
def fetch_full_financial_report(corp_code: str, api_key: str):
    """재무상태표(총자산/총부채/자기자본 추이)와 연도별·분기별 손익계산서를 종합 조회.
    여러 번의 DART 조회가 필요해 다소 시간이 걸릴 수 있다 (각 조회는 24시간 캐시)."""
    if not corp_code:
        return None

    this_year = datetime.date.today().year
    last_year = this_year - 1

    result = {"BS": None, "연간손익": None, "분기손익": None, "meta": {}}

    QUARTER_REPRT = [("11013", "1분기"), ("11012", "2분기"), ("11014", "3분기")]

    # --- 0) 필요한 원본들을 미리 확보 ---
    rows_last_annual, _ = _fetch_dart_report(corp_code, api_key, last_year, "11011")   # 작년 사업보고서
    annual_is = _extract_is_values(rows_last_annual)
    annual_bs = _extract_bs_values(rows_last_annual)

    # 작년 분기별 원본 전부 확보 (연간손익 비교용 + 분기손익 둘 다 재사용)
    last_year_quarter_rows = {}
    for reprt_code, q_label in QUARTER_REPRT:
        rows, _ = _fetch_dart_report(corp_code, api_key, last_year, reprt_code)
        if rows:
            last_year_quarter_rows[q_label] = rows
    last_year_quarter_is = {q: _extract_is_values(rows) for q, rows in last_year_quarter_rows.items()}

    # 올해 존재하는 분기 원본들을 전부 확보 (1분기→2분기→3분기 순, 있는 만큼)
    this_year_quarters = {}
    for reprt_code, q_label in QUARTER_REPRT:
        rows, _ = _fetch_dart_report(corp_code, api_key, this_year, reprt_code)
        if rows:
            this_year_quarters[q_label] = rows
    this_year_quarter_is = {q: _extract_is_values(rows) for q, rows in this_year_quarters.items()}

    # --- 1) 재무상태표: 2024년말/2025년말은 작년 사업보고서에서, 올해 최신은 올해 최신 분기에서 ---
    latest_q_label_for_bs, latest_q_rows = None, None
    for q_label in ["3분기", "2분기", "1분기"]:
        if q_label in this_year_quarters:
            latest_q_label_for_bs = q_label
            latest_q_rows = this_year_quarters[q_label]
            break

    bs_result = {}
    for metric in BS_ACCOUNTS:
        d_annual = annual_bs.get(metric, {})
        v_2024 = d_annual.get("전기말")     # 작년 사업보고서의 전기말 = 전전년(2024)말
        v_2025 = d_annual.get("당기말")     # 작년 사업보고서의 당기말 = 작년(2025)말
        v_latest = None
        if latest_q_rows:
            d_latest = _extract_bs_values(latest_q_rows).get(metric, {})
            v_latest = d_latest.get("당기말")
        bs_result[metric] = {f"{last_year - 1}년": v_2024, f"{last_year}년": v_2025,
                              "최신": v_latest}
    result["BS"] = bs_result
    result["meta"]["BS_라벨"] = [f"{last_year - 1}년", f"{last_year}년",
                                f"{this_year}년 {latest_q_label_for_bs}" if latest_q_label_for_bs else f"{this_year}년(자료없음)"]

    # --- 2) 연도별 손익계산서: 작년 사업보고서(전년+전전년) + 올해 최신 누적 vs 전년동기누적 ---
    latest_ytd, latest_ytd_label, latest_ytd_q = {}, None, None
    for q_label, cum_label in [("3분기", f"{this_year} 3분기 누적(9개월)"),
                                ("2분기", f"{this_year} 반기 누적(6개월)"),
                                ("1분기", f"{this_year} 1분기 누적(3개월)")]:
        if q_label in this_year_quarter_is:
            is_vals = this_year_quarter_is[q_label]
            if is_vals:
                latest_ytd, latest_ytd_label, latest_ytd_q = is_vals, cum_label, q_label
                break

    prior_ytd = last_year_quarter_is.get(latest_ytd_q, {}) if latest_ytd_q else {}
    prior_ytd_cum_label = {
        "1분기": f"{last_year} 1분기 누적(3개월)", "2분기": f"{last_year} 반기 누적(6개월)",
        "3분기": f"{last_year} 3분기 누적(9개월)",
    }.get(latest_ytd_q, "-")

    result["연간손익"] = {
        "전전년도": {"라벨": f"{last_year - 1}년(연간)",
                   "값": {m: annual_is.get(m, {}).get("전기") for m in IS_ACCOUNTS}},
        "전년도": {"라벨": f"{last_year}년(연간)",
                 "값": {m: annual_is.get(m, {}).get("당기") for m in IS_ACCOUNTS}},
        "올해누적": {"라벨": latest_ytd_label or "조회 실패",
                  "값": {m: latest_ytd.get(m, {}).get("당기누적") for m in IS_ACCOUNTS}},
        "전년동기누적": {"라벨": prior_ytd_cum_label,
                    "값": {m: prior_ytd.get(m, {}).get("당기누적") for m in IS_ACCOUNTS}},
    }

    # --- 3) 분기별 손익계산서: 작년 1~4분기(단독) + 올해 존재하는 분기 전부(단독) ---
    q4_last = {}
    for m in IS_ACCOUNTS:
        annual_val = annual_is.get(m, {}).get("당기")                          # 사업보고서 당기 = 연간 전체
        q3_cum = last_year_quarter_is.get("3분기", {}).get(m, {}).get("당기누적")  # 3분기보고서 누적 = 9개월
        q4_last[m] = (annual_val - q3_cum) if (annual_val is not None and q3_cum is not None) else None

    last_year_q_data = {
        "1분기": {m: last_year_quarter_is.get("1분기", {}).get(m, {}).get("당기") for m in IS_ACCOUNTS},
        "2분기": {m: last_year_quarter_is.get("2분기", {}).get(m, {}).get("당기") for m in IS_ACCOUNTS},
        "3분기": {m: last_year_quarter_is.get("3분기", {}).get(m, {}).get("당기") for m in IS_ACCOUNTS},
        "4분기": q4_last,
    }

    this_year_q_data = {}
    for q_label in ["1분기", "2분기", "3분기"]:
        if q_label in this_year_quarter_is:
            this_year_q_data[q_label] = {m: this_year_quarter_is[q_label].get(m, {}).get("당기") for m in IS_ACCOUNTS}

    result["분기손익"] = {
        "직전년도": {"연도": last_year, "분기": last_year_q_data},
        "올해": {"연도": this_year, "분기": this_year_q_data},
    }

    return result


def normalize_issuer_name(name):
    """회사명 표기 차이를 최대한 흡수 (지주/홀딩스/괄호 표기 + Google Sheets 별칭)."""
    if pd.isna(name):
        return None
    s = str(name).strip()
    for suf in ["주식회사", "(주)", "㈜"]:
        s = s.replace(suf, "")
    s = s.strip()
    aliases = load_issuer_aliases()
    return aliases.get(s, s)


NICE_OUTLOOK_MAP = {"S": "안정적", "P": "긍정적", "N": "부정적"}

def split_rating_outlook(raw, source):
    """'A+/안정적', 'AA-/S' 같은 표기를 (등급, 등급전망)으로 분리.
    NICE는 알파벳 코드(S/P/N)를 한글로 변환, 한신평은 이미 한글이라 그대로 사용."""
    if pd.isna(raw):
        return None, None
    text = str(raw).strip()
    if "/" not in text:
        return text, None
    grade, outlook_code = text.split("/", 1)
    grade, outlook_code = grade.strip(), outlook_code.strip()
    if source == "나이스신용평가":
        outlook = NICE_OUTLOOK_MAP.get(outlook_code, outlook_code)
    else:
        outlook = outlook_code
    return grade, outlook


RATER_ABBREV = {"한국신용평가": "한신평", "나이스신용평가": "나신평", "한국기업평가": "한기평"}

def parse_period_label(col):
    """열 이름(2026.03 같은 숫자 또는 '26.03 같은 문자열)을 'YY.MM' 라벨로 변환."""
    if isinstance(col, (int, float)):
        year = int(col)
        frac = round((col - year) * 100)
        month = frac if frac != 0 else 12
        return f"{year % 100:02d}.{month:02d}"
    s = str(col).strip()
    m = re.match(r"^'?(\d{2,4})\.(\d{1,2})", s)
    if m:
        yy, mm = m.groups()
        yy = int(yy)
        year = yy if yy > 100 else (2000 + yy)
        return f"{year % 100:02d}.{int(mm):02d}"
    return s


def check_trigger_met(actual, operator, threshold, direction):
    """실제 수치가 상향/하향 트리거 조건을 충족하는지 판정.
    판정 불가(숫자 아님 등)면 None, 충족하면 True, 아니면 False.
    Google Sheets에서 불러온 값은 전부 문자열일 수 있으므로 둘 다 안전하게 float 변환한다."""
    if pd.isna(actual) or pd.isna(operator) or pd.isna(threshold) or operator == "":
        return None
    try:
        actual_num = float(str(actual).replace(",", ""))
        threshold_num = float(str(threshold).replace(",", ""))
    except (ValueError, TypeError):
        return None
    ops = {
        ">=": actual_num >= threshold_num, "<=": actual_num <= threshold_num,
        ">": actual_num > threshold_num, "<": actual_num < threshold_num,
    }
    return ops.get(operator)


def trigger_signal(row):
    """방향(상향/하향)과 충족 여부를 신호등 이모지로 변환."""
    met = check_trigger_met(row.get("실제수치"), row.get("연산자"), row.get("값"), row.get("방향"))
    if met is None:
        return "-"
    if not met:
        return ""
    return "🟢" if row.get("방향") == "상향" else "🔴"


def _norm_col(c):
    """엑셀 헤더에 섞인 줄바꿈/공백을 제거해서 비교하기 쉽게 만든다."""
    return re.sub(r"\s+", "", str(c))


WINUS_COL_MAP = {
    "투자회사명": "발행사명",
    "검토의견": "검토의견",
    "내부등급코드": "내부등급",
    "(투자한도)회사채등": "투자한도",
    "(잔여한도)회사채등": "잔여한도",
    "(투자현황)회사채": "회사채잔액",
    "(투자현황)CP": "CP잔액",
    "(투자현황)CD": "CD잔액",
    "(투자현황)RP": "RP잔액",
    "(투자현황)정기예금": "정기예금잔액",
}


def parse_winus_file(file_obj):
    """위너스(WINUS) '유니버스조회' 엑셀에서 익스포저 관련 열만 뽑아 정제.
    반환: (정제된 DataFrame, 오류메시지). 성공 시 오류메시지는 None."""
    try:
        df = pd.read_excel(file_obj, sheet_name=0, header=0)
    except Exception as e:
        return None, f"파일을 읽는 중 오류가 발생했습니다: {e}"

    norm_to_orig = {_norm_col(c): c for c in df.columns}
    selected, missing = {}, []
    for key, friendly in WINUS_COL_MAP.items():
        if key in norm_to_orig:
            selected[friendly] = df[norm_to_orig[key]]
        else:
            missing.append(key)
    if missing:
        return None, f"예상한 열을 찾을 수 없습니다: {missing}"

    result = pd.DataFrame(selected)
    result = result[result["발행사명"].notna()].reset_index(drop=True)
    return result, None



# ==============================================================
# 페이지: 데이터 업로드
# ==============================================================
if page == "데이터 업로드":
    if not st.session_state.get("is_admin"):
        st.warning("관리자 로그인이 필요합니다. 왼쪽 사이드바에서 로그인해주세요.")
        st.stop()
    st.header("데이터 업로드")
    upload_tab1, upload_tab2, upload_tab3 = st.tabs(
        ["인포맥스 채권 스프레드", "신용평가사 트리거", "위너스 익스포저"]
    )
    with upload_tab1:
        st.caption("인포맥스 채권 수익률 파일을 업로드하면 '공모/무보증' 발행사만 자동으로 필터링해 보여줍니다.")

        KSIC_LARGE = {
            "A": "농업, 임업 및 어업", "B": "광업", "C": "제조업",
            "D": "전기,가스,증기 및 공기조절 공급업", "E": "수도,하수 및 폐기물 처리업",
            "F": "건설업", "G": "도매 및 소매업", "H": "운수 및 창고업",
            "I": "숙박 및 음식점업", "J": "정보통신업", "K": "금융 및 보험업",
            "L": "부동산업", "M": "전문,과학 및 기술 서비스업",
            "N": "사업시설관리 및 사업지원 서비스업", "O": "공공행정,국방 및 사회보장행정",
            "P": "교육 서비스업", "Q": "보건업 및 사회복지 서비스업",
            "R": "예술,스포츠 및 여가관련 서비스업", "S": "협회·단체,수리 및 기타 개인서비스업",
            "T": "가구내 고용활동 및 자가소비 생산활동", "U": "국제 및 외국기관",
        }
        MATURITY_COLS_WANTED = ["3M", "6M", "9M", "1Y", "3Y", "5Y"]

        BOND_CATEGORIES = ["공모/무보증", "은행채", "카드채", "기타금융채"]

        def _split_category_grade(raw):
            """'은행채 AA+' -> ('은행채', 'AA+') 처럼 앞의 카테고리 표기와 등급을 분리."""
            text = str(raw).strip()
            for cat in BOND_CATEGORIES:
                if text.startswith(cat):
                    return cat, text[len(cat):].strip()
            return None, text

        def parse_infomax_bond_file(file_obj):
            """인포맥스 채권 수익률 엑셀 → (정제된 result df, 사용된 만기 열 목록). 실패 시 (None, 오류메시지).
            공모/무보증, 은행채, 카드채, 기타금융채 네 종류만 추출한다."""
            try:
                df = pd.read_excel(file_obj, header=2)
            except Exception as e:
                return None, f"파일을 읽는 중 오류가 발생했습니다: {e}"

            rating_col = df.columns[4]
            issuer_col = '발행사'
            if rating_col not in df.columns or issuer_col not in df.columns:
                return None, "예상한 열 구조와 다릅니다. 파일 형식을 확인해주세요."

            mask_rating = df[rating_col].astype(str).str.strip().apply(
                lambda x: any(str(x).startswith(cat) for cat in BOND_CATEGORIES)
            )
            filtered = df[mask_rating].copy()
            filtered = filtered[filtered[issuer_col] != filtered[rating_col]]

            maturity_cols = [c for c in MATURITY_COLS_WANTED if c in filtered.columns]
            industry_col = "업종분류(소)"
            industry_code_col = "업종코드"

            display_cols = [rating_col, industry_col, issuer_col] + maturity_cols
            if industry_code_col in filtered.columns:
                display_cols = [industry_code_col] + display_cols
            display_cols = [c for c in display_cols if c in filtered.columns]

            result = filtered[display_cols].rename(
                columns={rating_col: "신용등급그룹", industry_col: "업종구분", industry_code_col: "업종코드"}
            )

            cat_grade = result["신용등급그룹"].apply(_split_category_grade)
            result.insert(0, "채권종류", [c for c, g in cat_grade])
            result["신용등급그룹"] = [g for c, g in cat_grade]
            # 'AAA(산금-이표)', 'AAA(중금-할인)' 같은 괄호 부가설명은 제거하고 순수 등급만 남긴다
            result["신용등급그룹"] = (
                result["신용등급그룹"].astype(str).str.replace(r"\(.*\)", "", regex=True).str.strip()
            )

            result["업종구분"] = result["업종구분"].fillna("미분류")

            if "업종코드" in result.columns:
                result["산업대분류"] = result["업종코드"].astype(str).str[0].map(KSIC_LARGE).fillna("미분류")
                result = result.drop(columns=["업종코드"])
                cols = ["산업대분류"] + [c for c in result.columns if c != "산업대분류"]
                result = result[cols]

            return result, maturity_cols

        # ------------------------------------------------------------
        # 과거 데이터 일괄 업로드 (여러 파일을 한 번에 이력에 저장)
        # ------------------------------------------------------------
        with st.expander("과거 데이터 일괄 업로드 (여러 날짜 파일을 한 번에 이력 저장)"):
            st.caption(
                "파일명에 8자리 날짜(YYYYMMDD)가 포함되어 있으면 자동으로 그 날짜로 저장됩니다 "
                "(예: 4788-20260714.xlsx → 2026-07-14). 날짜가 없는 파일은 오늘 날짜로 저장됩니다."
            )
            bulk_files = st.file_uploader(
                "과거 인포맥스 파일 여러 개 선택", type=["xlsx", "xls"],
                accept_multiple_files=True, key="bulk_infomax"
            )

            if bulk_files:
                st.write(f"{len(bulk_files)}개 파일이 선택되었습니다.")
                if st.button("전체 이력에 일괄 저장", key="bulk_save_btn"):
                    if "GOOGLE_SHEET_ID" not in st.secrets or "gcp_service_account" not in st.secrets:
                        st.warning(
                            "Google Sheets 연동 정보가 설정되어 있지 않습니다. "
                            "Streamlit Cloud Secrets에 GOOGLE_SHEET_ID와 gcp_service_account를 등록해주세요."
                        )
                    else:
                        progress = st.progress(0.0, text="일괄 저장 중...")
                        success_count, fail_list = 0, []
                        for i, f in enumerate(bulk_files):
                            res, info = parse_infomax_bond_file(f)
                            if res is None:
                                fail_list.append(f"{f.name}: {info}")
                            else:
                                date_str = extract_date_from_filename(f.name)
                                try:
                                    append_history(res, "채권스프레드_이력", date_str)
                                    success_count += 1
                                except Exception as e:
                                    fail_list.append(f"{f.name}: 저장 오류 - {e}")
                            progress.progress((i + 1) / len(bulk_files),
                                               text=f"처리 중... ({i+1}/{len(bulk_files)})")
                        progress.empty()

                        st.success(f"{success_count}/{len(bulk_files)}개 파일 저장 완료")
                        if fail_list:
                            st.error("다음 파일은 처리하지 못했습니다:\n" + "\n".join(fail_list))

        uploaded_file = st.file_uploader(
            "인포맥스 엑셀 파일을 업로드하세요 (예: 4788-YYYYMMDD.xlsx)", type=["xlsx", "xls"], key="infomax"
        )

        if uploaded_file is not None:
            result, info = parse_infomax_bond_file(uploaded_file)
            if result is None:
                st.error(info)
                st.stop()
            maturity_cols = info
            issuer_col = '발행사'

            st.subheader("필터")
            col_bt, col0, col1, col2 = st.columns(4)

            with col_bt:
                bond_type_options = sorted(result["채권종류"].dropna().unique())
                selected_bond_types = st.multiselect(
                    "채권종류 선택", options=bond_type_options, default=bond_type_options
                )

            with col0:
                industry_options = sorted(
                    result[result["채권종류"].isin(selected_bond_types)]["산업대분류"].unique()
                )
                selected_industries = st.multiselect(
                    "산업 대분류 선택", options=industry_options, default=industry_options
                )

            with col1:
                rating_options = sorted(
                    result[
                        result["채권종류"].isin(selected_bond_types)
                        & result["산업대분류"].isin(selected_industries)
                    ]["신용등급그룹"].unique()
                )
                selected_ratings = st.multiselect(
                    "신용등급 선택", options=rating_options, default=rating_options
                )

            issuer_pool = result[
                result["채권종류"].isin(selected_bond_types)
                & result["산업대분류"].isin(selected_industries)
                & result["신용등급그룹"].isin(selected_ratings)
            ][issuer_col].sort_values().unique()

            with col2:
                selected_issuers = st.multiselect(
                    "발행사 선택 (비워두면 전체 표시)", options=issuer_pool
                )

            view = result[
                result["채권종류"].isin(selected_bond_types)
                & result["산업대분류"].isin(selected_industries)
                & result["신용등급그룹"].isin(selected_ratings)
            ]
            if selected_issuers:
                view = view[view[issuer_col].isin(selected_issuers)]

            use_dart = st.checkbox(
                "DART 재무제표 함께 조회 (매출액·영업이익·당기순이익 — 최초 조회 시 다소 시간 소요, 이후 캐시됨)",
                value=False
            )

            financial_cols = ["매출액", "영업이익", "당기순이익", "재무제표기준일"]

            if use_dart:
                if not DART_API_KEY:
                    st.warning(
                        "DART API 키가 설정되어 있지 않습니다. Streamlit Cloud의 App settings → Secrets에 "
                        "DART_API_KEY = \"발급받은키\" 형식으로 등록해주세요."
                    )
                else:
                    with st.spinner("DART에서 재무제표를 조회하는 중입니다..."):
                        corp_map = load_corp_code_map(DART_API_KEY)
                        unique_issuers = view[issuer_col].unique().tolist()
                        progress = st.progress(0.0, text="재무제표 조회 중...")
                        fin_rows = []
                        for i, name in enumerate(unique_issuers):
                            corp_code = find_corp_code(name, corp_map)
                            fin = fetch_financials(corp_code, DART_API_KEY)
                            fin[issuer_col] = name
                            fin_rows.append(fin)
                            progress.progress((i + 1) / max(len(unique_issuers), 1),
                                               text=f"재무제표 조회 중... ({i+1}/{len(unique_issuers)})")
                        progress.empty()

                    fin_df = pd.DataFrame(fin_rows)
                    view = view.merge(fin_df, on=issuer_col, how="left")
                    other_cols = [c for c in view.columns if c not in financial_cols]
                    view = view[other_cols + ["매출액", "영업이익", "당기순이익", "재무제표기준일"]]

            st.subheader(f"공모/무보증 발행사 목록 ({len(view)}개)")
            st.dataframe(view, use_container_width=True, hide_index=True)

            st.subheader("신용등급별 발행사 수")
            counts_by_type_rating = view.groupby(["채권종류", "신용등급그룹"]).size()
            counts_by_type_rating.index = [f"{t} {r}" for t, r in counts_by_type_rating.index]
            st.bar_chart(counts_by_type_rating.rename("발행사 수"))

            if maturity_cols:
                st.subheader("등급별 만기별 평균 수익률 (채권종류별)")
                st.caption("은행채·카드채·기타금융채·공모무보증은 같은 등급이라도 스프레드 수준이 달라 채권종류별로 따로 계산합니다.")
                rating_order = ["AAA", "AA+", "AA0", "AA-", "A+", "A0", "A-",
                                 "BBB+", "BBB0", "BBB-", "BB+", "BB0", "BB-"]
                for bond_type in [c for c in bond_type_options if c in view["채권종류"].unique()]:
                    subset = view[view["채권종류"] == bond_type]
                    if subset.empty:
                        continue
                    avg_by_rating = subset.groupby("신용등급그룹")[maturity_cols].mean().round(3)
                    ordered = [r for r in rating_order if r in avg_by_rating.index]
                    remaining = [r for r in avg_by_rating.index if r not in rating_order]
                    avg_by_rating = avg_by_rating.loc[ordered + remaining]
                    st.markdown(f"**{bond_type}**")
                    st.dataframe(avg_by_rating, use_container_width=True)
                    st.line_chart(avg_by_rating.T)

            if maturity_cols:
                st.subheader("발행사별 만기 수익률 비교")
                pick = st.multiselect(
                    "비교할 발행사 선택 (최대 10개 추천)",
                    options=view[issuer_col].tolist(), default=view[issuer_col].tolist()[:5]
                )
                if pick:
                    chart_data = view[view[issuer_col].isin(pick)].set_index(issuer_col)[maturity_cols]
                    st.line_chart(chart_data.T)

            st.download_button(
                "필터링된 결과 CSV로 다운로드",
                data=view.to_csv(index=False).encode("utf-8-sig"),
                file_name="공모무보증_발행사_필터결과.csv", mime="text/csv"
            )

            # ------------------------------------------------------------
            # 이력 저장 (Google Sheets 누적) — 선택사항
            # ------------------------------------------------------------
            st.subheader("이력 저장 (선택사항)")
            st.caption(
                "현재 화면에 표시된 데이터(위 필터가 적용된 상태)를 날짜와 함께 Google Sheets에 누적 저장합니다. "
                "같은 날짜로 다시 저장하면 그 날짜 데이터만 덮어씁니다."
            )
            default_date = extract_date_from_filename(uploaded_file.name)
            date_str = st.text_input("기준일자 (YYYYMMDD)", value=default_date, key="history_date")

            if st.button("이력에 저장", key="save_history_btn"):
                if "GOOGLE_SHEET_ID" not in st.secrets or "gcp_service_account" not in st.secrets:
                    st.warning(
                        "Google Sheets 연동 정보가 설정되어 있지 않습니다. "
                        "Streamlit Cloud Secrets에 GOOGLE_SHEET_ID와 gcp_service_account를 등록해주세요."
                    )
                elif not re.fullmatch(r"20\d{6}", date_str or ""):
                    st.error("기준일자는 YYYYMMDD 8자리 숫자로 입력해주세요 (예: 20260814).")
                else:
                    try:
                        with st.spinner("Google Sheets에 저장하는 중입니다..."):
                            append_history(view, "채권스프레드_이력", date_str)
                        st.success(f"{date_str} 기준 {len(view)}건을 이력에 저장했습니다.")
                    except Exception as e:
                        st.error(f"저장 중 오류가 발생했습니다: {e}")
        else:
            st.info("왼쪽 상단에서 인포맥스 엑셀 파일을 업로드하면 대시보드가 표시됩니다.")


    with upload_tab2:
        st.caption(
            "한국신용평가(IS)·나이스신용평가(NICE)·한국기업평가(KR) 3개사의 등급변동 트리거 파일을 "
            "업로드하면, 표현 방식이 달라도 하나의 표로 통합해서 보여줍니다."
        )

        DIRECTION_MAP = {"이상": ">=", "이하": "<=", "초과": ">", "미만": "<", "상회": ">", "하회": "<"}
        SYMBOLIC_RE = re.compile(r'^(?P<op>>=|<=|>|<)\s*(?P<val>-?\.?\d[\d,\.]*)\s*(?P<unit>%|배)?$')
        PHRASE_RE = re.compile(r'^(?P<val>-?[\d,\.]+)\s*(?P<unit>%|배|억원|조원|만원)?\s*(?P<word>이상|이하|초과|미만|상회|하회)$')

        def parse_threshold(raw):
            if pd.isna(raw):
                return None
            text = str(raw).strip()
            if text == "" or text.lower() == "n.a.":
                return None
            m = SYMBOLIC_RE.match(text)
            if m:
                val = float(m.group("val").replace(",", ""))
                return {"연산자": m.group("op"), "값": val, "단위": m.group("unit"), "원문": text, "특이조건": None}
            m = PHRASE_RE.match(text)
            if m:
                val = float(m.group("val").replace(",", ""))
                op = DIRECTION_MAP[m.group("word")]
                return {"연산자": op, "값": val, "단위": m.group("unit"), "원문": text, "특이조건": None}
            return {"연산자": None, "값": None, "단위": None, "원문": text, "특이조건": text}

        def infer_unit(indicator_name, parsed_unit):
            if parsed_unit:
                return parsed_unit
            name = str(indicator_name)
            if "배" in name:
                return "배"
            if any(k in name for k in ["비율", "율", "Margin", "M/S", "ROA", "ROE"]):
                return "%"
            if "억원" in name:
                return "억원"
            return None

        st.caption(
            "💡 회사명 표기 차이(별칭)는 이제 **'관리자 설정'** 탭에서 관리합니다. "
            "새로운 표기 차이를 발견하시면 그 탭에서 추가해주세요."
        )

        st.markdown("**5개 파일 업로드**")
        c1, c2 = st.columns(2)
        with c1:
            f_is_corp = st.file_uploader("한국신용평가(IS) — 기업부문", type=["xlsx"], key="is_corp")
            f_is_fin = st.file_uploader("한국신용평가(IS) — 금융부문", type=["xlsx"], key="is_fin")
            f_kr = st.file_uploader("한국기업평가(KR)", type=["xls", "xlsx"], key="kr")
        with c2:
            f_nice_corp = st.file_uploader("나이스신용평가(NICE) — 기업", type=["xlsx"], key="nice_corp")
            f_nice_fin = st.file_uploader("나이스신용평가(NICE) — 금융", type=["xlsx"], key="nice_fin")

        rows = []

        # --- 한신평 기업부문 ---
        if f_is_corp is not None:
            try:
                eff_date = extract_effective_date(f_is_corp, sheet_name='업체별 KMI 지표 및 실적')
                df = pd.read_excel(f_is_corp, sheet_name='업체별 KMI 지표 및 실적', header=13)
                up_idx = df.columns.get_loc('상향가능성')
                latest_col = df.columns[up_idx - 1]
                latest_label = parse_period_label(latest_col)
                for _, r in df.iterrows():
                    issuer, indicator = r.get('업체명'), r.get('KMI 지표')
                    if pd.isna(issuer) or pd.isna(indicator):
                        continue
                    for direction, col in [('상향', '상향가능성'), ('하향', '하향가능성')]:
                        parsed = parse_threshold(r.get(col))
                        if parsed is None:
                            continue
                        grade, outlook = split_rating_outlook(r.get('등급'), "한국신용평가")
                        rows.append({"신평사": "한국신용평가", "기준일": eff_date, "원본발행사명": issuer,
                                      "발행사(정규화)": normalize_issuer_name(issuer), "등급": grade, "등급전망": outlook,
                                      "지표명": indicator, "실제수치": r.get(latest_col), "기준월": latest_label,
                                      "방향": direction, **parsed,
                                      "단위": infer_unit(indicator, parsed["단위"])})
            except Exception as e:
                st.error(f"한신평 기업부문 파일 처리 오류: {e}")

        # --- 한신평 금융부문 ---
        if f_is_fin is not None:
            try:
                eff_date = extract_effective_date(f_is_fin, sheet_name='업체별 KMI 지표 및 실적')
                df = pd.read_excel(f_is_fin, sheet_name='업체별 KMI 지표 및 실적', header=13)
                up_idx = df.columns.get_loc('상향가능성')
                latest_col = df.columns[up_idx - 1]
                latest_label = parse_period_label(latest_col)
                for _, r in df.iterrows():
                    issuer, indicator = r.get('업체명'), r.get('KMI')
                    if pd.isna(issuer) or pd.isna(indicator) or r.get('정성/정량') != '정량':
                        continue
                    for direction, col in [('상향', '상향가능성'), ('하향', '하향가능성')]:
                        parsed = parse_threshold(r.get(col))
                        if parsed is None:
                            continue
                        grade, outlook = split_rating_outlook(r.get('등급'), "한국신용평가")
                        rows.append({"신평사": "한국신용평가", "기준일": eff_date, "원본발행사명": issuer,
                                      "발행사(정규화)": normalize_issuer_name(issuer), "등급": grade, "등급전망": outlook,
                                      "지표명": indicator, "실제수치": r.get(latest_col), "기준월": latest_label,
                                      "방향": direction, **parsed,
                                      "단위": infer_unit(indicator, parsed["단위"])})
            except Exception as e:
                st.error(f"한신평 금융부문 파일 처리 오류: {e}")

        # --- NICE 기업/금융 (동일 구조) ---
        for f, label in [(f_nice_corp, "NICE(기업)"), (f_nice_fin, "NICE(금융)")]:
            if f is not None:
                try:
                    eff_date = extract_effective_date(f, sheet_name=0)
                    df = pd.read_excel(f, header=12)
                    df.columns = [str(c).strip() for c in df.columns]
                    issuer_col = df.columns[1]
                    rating_col = [c for c in df.columns if c == "'26.06"][0]
                    indicator_col = [c for c in df.columns if '트리거지표' in c][0]
                    up_col = [c for c in df.columns if '상향' in c][0]
                    down_col = [c for c in df.columns if '하향' in c][0]
                    up_idx = df.columns.get_loc(up_col)
                    latest_col = df.columns[up_idx - 1]
                    latest_label = parse_period_label(latest_col)
                    for _, r in df.iterrows():
                        issuer, indicator = r.get(issuer_col), r.get(indicator_col)
                        if pd.isna(issuer) or pd.isna(indicator):
                            continue
                        for direction, col in [('상향', up_col), ('하향', down_col)]:
                            parsed = parse_threshold(r.get(col))
                            if parsed is None:
                                continue
                            grade, outlook = split_rating_outlook(r.get(rating_col), "나이스신용평가")
                            rows.append({"신평사": "나이스신용평가", "기준일": eff_date, "원본발행사명": issuer,
                                          "발행사(정규화)": normalize_issuer_name(issuer), "등급": grade, "등급전망": outlook,
                                          "지표명": indicator, "실제수치": r.get(latest_col), "기준월": latest_label,
                                          "방향": direction, **parsed,
                                          "단위": infer_unit(indicator, parsed["단위"])})
                except Exception as e:
                    st.error(f"{label} 파일 처리 오류: {e}")

        # --- 한기평(KR) ---
        if f_kr is not None:
            try:
                eff_date = extract_effective_date(f_kr, sheet_name='Sheet1')
                df = pd.read_excel(f_kr, sheet_name='Sheet1', header=25)
                # KR은 시계열 실적 열이 상향/하향보다 뒤에 위치 (예: 2023.12, 2024.12, 2025.12, 2026.03)
                period_candidates = [c for c in df.columns if isinstance(c, (int, float)) and c > 2000]
                latest_col = max(period_candidates) if period_candidates else None
                latest_label = parse_period_label(latest_col) if latest_col is not None else None
                for _, r in df.iterrows():
                    issuer, indicator = r.get('업체명'), r.get('등급변동요인')
                    if pd.isna(issuer) or pd.isna(indicator):
                        continue
                    for direction, col in [('상향', '상향'), ('하향', '하향')]:
                        parsed = parse_threshold(r.get(col))
                        if parsed is None:
                            continue
                        rows.append({"신평사": "한국기업평가", "기준일": eff_date, "원본발행사명": issuer,
                                      "발행사(정규화)": normalize_issuer_name(issuer), "등급": r.get('등급'),
                                      "등급전망": r.get('등급전망'),
                                      "지표명": indicator,
                                      "실제수치": r.get(latest_col) if latest_col is not None else None,
                                      "기준월": latest_label,
                                      "방향": direction, **parsed,
                                      "단위": infer_unit(indicator, parsed["단위"])})
            except Exception as e:
                st.error(f"한기평(KR) 파일 처리 오류: {e}")

        if rows:
            unified = pd.DataFrame(rows)
            unified["충족여부"] = unified.apply(trigger_signal, axis=1)
            st.success(f"총 {len(unified)}개 트리거 항목을 통합했습니다 (신평사 {unified['신평사'].nunique()}개사).")

            date_summary = unified.groupby("신평사")["기준일"].first()
            st.caption("신평사별 인식된 기준일: " + ", ".join(f"{k} {v}" for k, v in date_summary.items()))

            period_labels = unified["기준월"].dropna().unique()
            period_label_text = period_labels[0] if len(period_labels) > 0 else "YY.MM"
            actual_col_name = f"실제 수치({period_label_text})"

            st.subheader("발행사 검색")
            search = st.text_input("발행사명 입력 (부분 검색 가능)", key="trigger_search")

            if search:
                hit = unified[unified["발행사(정규화)"].str.contains(search, na=False)]
            else:
                hit = unified

            display_df = hit.copy()
            display_df["신평사"] = display_df["신평사"].map(RATER_ABBREV).fillna(display_df["신평사"])
            display_df = display_df.rename(columns={"실제수치": actual_col_name})

            st.dataframe(
                display_df[[
                    "원본발행사명", "발행사(정규화)", "신평사", "기준일", "등급", "등급전망",
                    "지표명", actual_col_name, "충족여부",
                    "방향", "원문", "연산자", "값", "단위", "특이조건"
                ]],
                use_container_width=True, hide_index=True
            )

            if search:
                matched_names = sorted(hit["발행사(정규화)"].dropna().unique())
                rater_count = hit.groupby("발행사(정규화)")["신평사"].nunique()
                st.caption(
                    f"검색된 발행사: {', '.join(matched_names[:10])}"
                    + (" ..." if len(matched_names) > 10 else "")
                )
                for name in matched_names:
                    n_raters = rater_count.get(name, 0)
                    if n_raters < 3:
                        st.caption(f"⚠️ '{name}': {n_raters}개 신평사에서만 발견됨 — 회사명 표기 차이로 일부가 안 잡혔을 수 있습니다.")

            st.download_button(
                "통합 트리거 결과 CSV 다운로드",
                data=unified.to_csv(index=False).encode("utf-8-sig"),
                file_name="신용등급_변동트리거_통합.csv", mime="text/csv"
            )

            # ------------------------------------------------------------
            # 이력 저장 (Google Sheets 누적) — 분기 단위 기준일로 저장
            # ------------------------------------------------------------
            st.subheader("이력 저장 (선택사항)")
            st.caption(
                "각 신평사 파일 안에서 자동으로 인식한 기준일 기준으로 저장됩니다. "
                "같은 (신평사, 기준일) 조합이 이미 있으면 그 부분만 덮어씁니다."
            )
            if st.button("이력에 저장", key="save_trigger_history_btn"):
                if "GOOGLE_SHEET_ID" not in st.secrets or "gcp_service_account" not in st.secrets:
                    st.warning(
                        "Google Sheets 연동 정보가 설정되어 있지 않습니다. "
                        "Streamlit Cloud Secrets에 GOOGLE_SHEET_ID와 gcp_service_account를 등록해주세요."
                    )
                elif unified["기준일"].isna().any():
                    missing = unified[unified["기준일"].isna()]["신평사"].unique()
                    st.error(
                        f"다음 신평사 파일에서 기준일을 자동으로 못 찾았습니다: {', '.join(missing)}. "
                        "파일 형식을 확인해주세요 (기준일 저장은 건너뛰었습니다)."
                    )
                else:
                    try:
                        with st.spinner("Google Sheets에 저장하는 중입니다..."):
                            append_history_multi(unified, "신용등급트리거_이력", dedup_cols=["신평사", "기준일"])
                        read_history.clear()
                        st.success(f"{len(unified)}건을 이력에 저장했습니다.")

                        # 저장 직후 즉시 재조회해서 실제로 반영됐는지 앱 안에서 바로 확인
                        verify_df = read_history("신용등급트리거_이력")
                        if verify_df.empty:
                            st.error(
                                "⚠️ 저장은 오류 없이 끝났지만, 방금 다시 읽어보니 시트가 비어 있습니다. "
                                "GOOGLE_SHEET_ID가 지금 보고 계신 시트와 다른 시트를 가리키고 있을 수 있습니다. "
                                "Secrets의 GOOGLE_SHEET_ID 값을 다시 확인해주세요."
                            )
                        else:
                            st.info(f"✅ 저장 직후 재확인: 시트에서 {len(verify_df)}행을 정상적으로 읽어왔습니다.")

                        sheet_id = st.secrets.get("GOOGLE_SHEET_ID", "")
                        if sheet_id:
                            st.caption(
                                f"현재 연결된 시트: https://docs.google.com/spreadsheets/d/{sheet_id}/edit "
                                "(이 링크가 지금 보고 계신 시트와 같은지 확인해주세요)"
                            )
                    except Exception as e:
                        st.error(f"저장 중 오류가 발생했습니다: {e}")
        else:
            st.info("왼쪽에서 신평사 파일을 하나 이상 업로드하면 통합 결과가 표시됩니다.")

    with upload_tab3:
        with st.expander("위너스(WINUS) 익스포저 업로드"):
            st.caption("'유니버스조회' 엑셀을 업로드하면 발행사별 익스포저(투자한도·잔여한도·잔액)를 이력에 저장합니다.")
            winus_file = st.file_uploader("위너스 유니버스조회 엑셀 업로드", type=["xlsx"], key="winus_upload")
            if winus_file is not None:
                winus_parsed, winus_err = parse_winus_file(winus_file)
                if winus_err:
                    st.error(winus_err)
                else:
                    st.success(f"{len(winus_parsed)}개 발행사 익스포저를 확인했습니다.")
                    st.dataframe(winus_parsed, use_container_width=True, hide_index=True, height=400)
                    winus_date = st.text_input(
                        "기준일자 (YYYYMMDD)",
                        value=extract_date_from_filename(winus_file.name),
                        key="winus_date_input"
                    )
                    if st.button("위너스 이력에 저장", key="save_winus_history_btn"):
                        if not re.fullmatch(r"20\d{6}", winus_date or ""):
                            st.error("기준일자는 YYYYMMDD 8자리 숫자로 입력해주세요.")
                        else:
                            try:
                                with st.spinner("Google Sheets에 저장하는 중입니다..."):
                                    append_history(winus_parsed, "위너스_이력", winus_date)
                                read_history.clear()
                                st.success(f"{winus_date} 기준 {len(winus_parsed)}건을 이력에 저장했습니다.")
                            except Exception as e:
                                st.error(f"저장 중 오류가 발생했습니다: {e}")

# ==============================================================
# 페이지: 채권 스프레드 (조회)
# ==============================================================
elif page == "채권 스프레드":
    st.header("채권 스프레드")
    st.caption("저장된 채권 스프레드 이력을 날짜 기준으로 조회합니다. 새 데이터는 '데이터 업로드' 메뉴에서 올려주세요.")

    if "GOOGLE_SHEET_ID" not in st.secrets or "gcp_service_account" not in st.secrets:
        st.warning(
            "Google Sheets 연동 정보가 설정되어 있지 않습니다. "
            "Streamlit Cloud Secrets에 GOOGLE_SHEET_ID와 gcp_service_account를 등록해주세요."
        )
    else:
        col_a, col_b = st.columns([3, 1])
        with col_a:
            q_date_obj = st.date_input(
                "조회할 날짜", value=datetime.date.today(), format="YYYY-MM-DD", key="spread_page_date"
            )
        with col_b:
            st.write("")
            st.write("")
            if st.button("새로고침", key="spread_page_refresh"):
                read_history.clear()
        q_date = q_date_obj.strftime("%Y%m%d")

        with st.spinner("이력을 불러오는 중입니다..."):
            spread_hist = read_history("채권스프레드_이력")

        if spread_hist.empty or "업데이트일자" not in spread_hist.columns:
            st.info("저장된 채권 스프레드 이력이 없습니다. '데이터 업로드' 메뉴에서 인포맥스 파일을 저장해주세요.")
        else:
            sh_all = spread_hist.copy()
            sh_all["업데이트일자"] = sh_all["업데이트일자"].astype(str)
            sh_valid = sh_all[sh_all["업데이트일자"] <= q_date]
            if sh_valid.empty:
                available_dates = sorted(sh_all["업데이트일자"].unique())
                st.info(
                    f"{q_date} 이전에 저장된 데이터가 없습니다. "
                    f"저장된 날짜: {', '.join(available_dates[-10:])}"
                    + (" ..." if len(available_dates) > 10 else "")
                )
            else:
                latest_date = sh_valid["업데이트일자"].max()
                day_spread = sh_valid[sh_valid["업데이트일자"] == latest_date]
                if latest_date == q_date:
                    st.caption(f"{q_date} 당일 데이터입니다. ({len(day_spread)}개 발행사)")
                else:
                    st.info(
                        f"{q_date} 날짜의 데이터가 없어, 가장 가까운 이전 날짜인 "
                        f"{latest_date} 데이터를 대신 보여드립니다. ({len(day_spread)}개 발행사)"
                    )

                st.subheader("필터")
                col_bt, col0, col1, col2 = st.columns(4)

                with col_bt:
                    bt_opts = sorted(day_spread["채권종류"].dropna().unique()) if "채권종류" in day_spread.columns else []
                    sel_bt = st.multiselect("채권종류 선택", options=bt_opts, default=bt_opts, key="spr_bt")

                with col0:
                    if "산업대분류" in day_spread.columns:
                        ind_opts = sorted(day_spread[day_spread["채권종류"].isin(sel_bt)]["산업대분류"].unique())
                    else:
                        ind_opts = []
                    sel_ind = st.multiselect("산업 대분류 선택", options=ind_opts, default=ind_opts, key="spr_ind")

                with col1:
                    if "신용등급그룹" in day_spread.columns:
                        rt_opts = sorted(
                            day_spread[
                                day_spread["채권종류"].isin(sel_bt) & day_spread["산업대분류"].isin(sel_ind)
                            ]["신용등급그룹"].unique()
                        )
                    else:
                        rt_opts = []
                    sel_rt = st.multiselect("신용등급 선택", options=rt_opts, default=rt_opts, key="spr_rt")

                mask = (
                    day_spread["채권종류"].isin(sel_bt)
                    & day_spread["산업대분류"].isin(sel_ind)
                    & day_spread["신용등급그룹"].isin(sel_rt)
                )
                issuer_pool = sorted(day_spread[mask]["발행사"].unique()) if "발행사" in day_spread.columns else []

                with col2:
                    sel_issuer = st.multiselect("발행사 선택 (비워두면 전체 표시)", options=issuer_pool, key="spr_issuer")

                view = day_spread[mask]
                if sel_issuer:
                    view = view[view["발행사"].isin(sel_issuer)]

                st.subheader(f"발행사 목록 ({len(view)}개)")
                st.dataframe(view.drop(columns=["업데이트일자"]), use_container_width=True, hide_index=True, height=400)

                maturity_cols = [c for c in ["3M", "6M", "9M", "1Y", "3Y", "5Y"] if c in view.columns]
                if maturity_cols and not view.empty:
                    # Google Sheets 이력에서 불러온 값은 전부 문자열이라, 평균 계산 전에 숫자로 변환한다
                    view = view.copy()
                    for c in maturity_cols:
                        view[c] = pd.to_numeric(view[c], errors="coerce")

                    st.subheader("등급별 만기별 평균 수익률 (채권종류별)")
                    st.caption("은행채·카드채·기타금융채·공모무보증은 같은 등급이라도 스프레드 수준이 달라 채권종류별로 따로 계산합니다.")
                    rating_order = ["AAA", "AA+", "AA0", "AA-", "A+", "A0", "A-",
                                     "BBB+", "BBB0", "BBB-", "BB+", "BB0", "BB-"]
                    for bond_type in sel_bt:
                        subset = view[view["채권종류"] == bond_type]
                        if subset.empty:
                            continue
                        avg_by_rating = subset.groupby("신용등급그룹")[maturity_cols].mean().round(3)
                        ordered = [r for r in rating_order if r in avg_by_rating.index]
                        remaining = [r for r in avg_by_rating.index if r not in rating_order]
                        avg_by_rating = avg_by_rating.loc[ordered + remaining]
                        st.markdown(f"**{bond_type}**")
                        st.dataframe(avg_by_rating, use_container_width=True)
                        st.line_chart(avg_by_rating.T)

                if not view.empty:
                    st.download_button(
                        "필터링된 결과 CSV로 다운로드",
                        data=view.to_csv(index=False).encode("utf-8-sig"),
                        file_name=f"채권스프레드_{q_date}.csv", mime="text/csv"
                    )

# ==============================================================
# 페이지: 신용등급 트리거 (조회)
# ==============================================================
elif page == "신용등급 트리거":
    st.header("신용등급 트리거")
    st.caption("저장된 신용등급 변동 트리거 이력을 날짜 기준으로 조회합니다. 새 데이터는 '데이터 업로드' 메뉴에서 올려주세요.")

    if "GOOGLE_SHEET_ID" not in st.secrets or "gcp_service_account" not in st.secrets:
        st.warning(
            "Google Sheets 연동 정보가 설정되어 있지 않습니다. "
            "Streamlit Cloud Secrets에 GOOGLE_SHEET_ID와 gcp_service_account를 등록해주세요."
        )
    else:
        col_a, col_b = st.columns([3, 1])
        with col_a:
            q_date_obj = st.date_input(
                "조회할 날짜", value=datetime.date.today(), format="YYYY-MM-DD", key="trigger_page_date"
            )
        with col_b:
            st.write("")
            st.write("")
            if st.button("새로고침", key="trigger_page_refresh"):
                read_history.clear()
        q_date = q_date_obj.strftime("%Y%m%d")

        with st.spinner("이력을 불러오는 중입니다..."):
            trigger_hist = read_history("신용등급트리거_이력")

        if trigger_hist.empty or "기준일" not in trigger_hist.columns:
            st.info("저장된 트리거 이력이 없습니다. '데이터 업로드' 메뉴에서 신평사 파일을 저장해주세요.")
        else:
            th = trigger_hist.copy()
            th["기준일"] = th["기준일"].astype(str)
            th_valid = th[th["기준일"] <= q_date]
            if th_valid.empty:
                st.info(f"{q_date} 이전에 저장된 트리거 이력이 없습니다.")
            else:
                latest_by_rater = th_valid.groupby("신평사")["기준일"].max()
                st.caption(
                    "신평사별로 적용된 기준일: "
                    + ", ".join(f"{k} ({v})" for k, v in latest_by_rater.items())
                )
                keep_mask = th_valid.apply(
                    lambda row: row["기준일"] == latest_by_rater[row["신평사"]], axis=1
                )
                as_of_trigger = th_valid[keep_mask]

                st.subheader("발행사 검색")
                search = st.text_input("발행사명 입력 (부분 검색 가능)", key="trigger_page_search")

                if search and "발행사(정규화)" in as_of_trigger.columns:
                    hit = as_of_trigger[as_of_trigger["발행사(정규화)"].str.contains(search, na=False)]
                else:
                    hit = as_of_trigger

                display_df = hit.copy()
                if "신평사" in display_df.columns:
                    display_df["신평사"] = display_df["신평사"].map(RATER_ABBREV).fillna(display_df["신평사"])

                actual_col_name = "실제수치"
                if "기준월" in display_df.columns:
                    labels = display_df["기준월"].dropna().unique()
                    if len(labels) > 0:
                        actual_col_name = f"실제 수치({labels[0]})"

                if "충족여부" not in display_df.columns and {"실제수치", "연산자", "값", "방향"} <= set(display_df.columns):
                    display_df["충족여부"] = display_df.apply(trigger_signal, axis=1)

                if "실제수치" in display_df.columns:
                    display_df = display_df.rename(columns={"실제수치": actual_col_name})

                show_cols = [c for c in
                    ["원본발행사명", "발행사(정규화)", "신평사", "기준일", "등급", "등급전망",
                     "지표명", actual_col_name, "충족여부",
                     "방향", "원문", "연산자", "값", "단위", "특이조건"]
                    if c in display_df.columns]

                st.subheader(f"트리거 목록 ({len(display_df)}건)")
                st.dataframe(display_df[show_cols], use_container_width=True, hide_index=True, height=400)

                if search and "발행사(정규화)" in hit.columns:
                    matched_names = sorted(hit["발행사(정규화)"].dropna().unique())
                    rater_count = hit.groupby("발행사(정규화)")["신평사"].nunique()
                    st.caption(
                        f"검색된 발행사: {', '.join(matched_names[:10])}"
                        + (" ..." if len(matched_names) > 10 else "")
                    )
                    for name in matched_names:
                        n_raters = rater_count.get(name, 0)
                        if n_raters < 3:
                            st.caption(f"⚠️ '{name}': {n_raters}개 신평사에서만 발견됨 — 회사명 표기 차이일 수 있습니다.")

                if not display_df.empty:
                    st.download_button(
                        "결과 CSV 다운로드",
                        data=display_df.to_csv(index=False).encode("utf-8-sig"),
                        file_name=f"신용등급트리거_{q_date}.csv", mime="text/csv"
                    )

# ==============================================================
# 페이지: 익스포저 (조회)
# ==============================================================
elif page == "익스포저":
    st.header("익스포저")
    st.caption("저장된 위너스(WINUS) 익스포저 이력을 날짜 기준으로 조회합니다. 새 데이터는 '데이터 업로드' 메뉴에서 올려주세요.")

    if "GOOGLE_SHEET_ID" not in st.secrets or "gcp_service_account" not in st.secrets:
        st.warning(
            "Google Sheets 연동 정보가 설정되어 있지 않습니다. "
            "Streamlit Cloud Secrets에 GOOGLE_SHEET_ID와 gcp_service_account를 등록해주세요."
        )
    else:
        col_a, col_b = st.columns([3, 1])
        with col_a:
            q_date_obj = st.date_input(
                "조회할 날짜", value=datetime.date.today(), format="YYYY-MM-DD", key="exposure_page_date"
            )
        with col_b:
            st.write("")
            st.write("")
            if st.button("새로고침", key="exposure_page_refresh"):
                read_history.clear()
        q_date = q_date_obj.strftime("%Y%m%d")

        with st.spinner("이력을 불러오는 중입니다..."):
            winus_hist = read_history("위너스_이력")

        if winus_hist.empty or "업데이트일자" not in winus_hist.columns:
            st.info("저장된 위너스 이력이 없습니다. '데이터 업로드' 메뉴에서 위너스 파일을 저장해주세요.")
        else:
            wh = winus_hist.copy()
            wh["업데이트일자"] = wh["업데이트일자"].astype(str)
            wh_valid = wh[wh["업데이트일자"] <= q_date]
            if wh_valid.empty:
                available_winus_dates = sorted(wh["업데이트일자"].unique())
                st.info(
                    f"{q_date} 이전에 저장된 위너스 이력이 없습니다. "
                    f"저장된 날짜: {', '.join(available_winus_dates[-10:])}"
                    + (" ..." if len(available_winus_dates) > 10 else "")
                )
            else:
                latest_winus_date = wh_valid["업데이트일자"].max()
                as_of_winus = wh_valid[wh_valid["업데이트일자"] == latest_winus_date]
                st.caption(f"적용된 기준일: {latest_winus_date} ({len(as_of_winus)}개 발행사)")

                col_f1, col_f2 = st.columns(2)
                with col_f1:
                    review_opts = sorted(as_of_winus["검토의견"].dropna().unique()) if "검토의견" in as_of_winus.columns else []
                    sel_review = st.multiselect("검토의견 필터 (비워두면 전체)", options=review_opts, key="exp_review")
                with col_f2:
                    search_issuer = st.text_input("발행사명 검색", key="exp_search")

                view = as_of_winus.drop(columns=["업데이트일자"])
                if sel_review and "검토의견" in view.columns:
                    view = view[view["검토의견"].isin(sel_review)]
                if search_issuer and "발행사명" in view.columns:
                    view = view[view["발행사명"].str.contains(search_issuer, na=False)]

                st.subheader(f"익스포저 목록 ({len(view)}개)")
                st.dataframe(view, use_container_width=True, hide_index=True, height=400)

                if not view.empty:
                    st.download_button(
                        "결과 CSV 다운로드",
                        data=view.to_csv(index=False).encode("utf-8-sig"),
                        file_name=f"위너스익스포저_{q_date}.csv", mime="text/csv"
                    )

# ==============================================================
# 페이지: 발행사별 상세보기 (조회)
# ==============================================================
elif page == "발행사별 상세보기":
    st.header("발행사별 상세 보기")
    st.caption(
        "특정 날짜를 기준으로, 발행사 하나를 골라 채권 스프레드·신용등급 트리거·위너스 익스포저를 한 화면에 모아 보여줍니다."
    )

    if "GOOGLE_SHEET_ID" not in st.secrets or "gcp_service_account" not in st.secrets:
        st.warning(
            "Google Sheets 연동 정보가 설정되어 있지 않습니다. "
            "Streamlit Cloud Secrets에 GOOGLE_SHEET_ID와 gcp_service_account를 등록해주세요."
        )
    else:
        col_a, col_b = st.columns([3, 1])
        with col_a:
            query_date_obj = st.date_input(
                "조회할 날짜", value=datetime.date.today(), format="YYYY-MM-DD", key="detail_page_date"
            )
        with col_b:
            st.write("")
            st.write("")
            if st.button("새로고침 (최신 이력 다시 불러오기)", key="detail_page_refresh"):
                read_history.clear()

        query_date = query_date_obj.strftime("%Y%m%d")

        with st.spinner("이력을 불러오는 중입니다..."):
            spread_hist = read_history("채권스프레드_이력")
            trigger_hist = read_history("신용등급트리거_이력")
            winus_hist = read_history("위너스_이력")

        # --- 그날 기준 유효한(가장 최근 저장된) 채권 스프레드 ---
        if spread_hist.empty or "업데이트일자" not in spread_hist.columns:
            day_spread = pd.DataFrame()
        else:
            sh_all = spread_hist.copy()
            sh_all["업데이트일자"] = sh_all["업데이트일자"].astype(str)
            sh_valid = sh_all[sh_all["업데이트일자"] <= query_date]
            if sh_valid.empty:
                day_spread = pd.DataFrame()
            else:
                latest_spread_date = sh_valid["업데이트일자"].max()
                day_spread = sh_valid[sh_valid["업데이트일자"] == latest_spread_date]

        # --- 해당 시점 기준 최신 트리거 (신평사별로 기준일 <= query_date 중 최댓값) ---
        if trigger_hist.empty or "기준일" not in trigger_hist.columns:
            as_of_trigger = pd.DataFrame()
        else:
            th = trigger_hist.copy()
            th["기준일"] = th["기준일"].astype(str)
            th_valid = th[th["기준일"] <= query_date]
            if th_valid.empty:
                as_of_trigger = pd.DataFrame()
            else:
                latest_by_rater = th_valid.groupby("신평사")["기준일"].max()
                keep_mask = th_valid.apply(
                    lambda row: row["기준일"] == latest_by_rater[row["신평사"]], axis=1
                )
                as_of_trigger = th_valid[keep_mask]

        # --- 해당 시점 기준 최신 위너스 익스포저 ---
        if winus_hist.empty or "업데이트일자" not in winus_hist.columns:
            as_of_winus = pd.DataFrame()
        else:
            wh = winus_hist.copy()
            wh["업데이트일자"] = wh["업데이트일자"].astype(str)
            wh_valid = wh[wh["업데이트일자"] <= query_date]
            if wh_valid.empty:
                as_of_winus = pd.DataFrame()
            else:
                latest_winus_date = wh_valid["업데이트일자"].max()
                as_of_winus = wh_valid[wh_valid["업데이트일자"] == latest_winus_date]

        if day_spread.empty and as_of_trigger.empty and as_of_winus.empty:
            st.info(
                "저장된 이력이 없습니다. '데이터 업로드' 메뉴에서 인포맥스·신평사·위너스 파일을 먼저 저장해주세요."
            )
        else:
            # --- 발행사 선택 목록: 채권스프레드·트리거·위너스 3개 소스를 합쳐서 구성 ---
            issuer_col_name = "발행사" if "발행사" in day_spread.columns else None

            options_map = {}  # 정규화명 -> 화면에 보여줄 대표 표시명
            if issuer_col_name:
                for raw in day_spread[issuer_col_name].dropna().unique():
                    norm = normalize_issuer_name(raw)
                    options_map.setdefault(norm, raw)
            if not as_of_trigger.empty and "원본발행사명" in as_of_trigger.columns:
                for raw in as_of_trigger["원본발행사명"].dropna().unique():
                    norm = normalize_issuer_name(raw)
                    options_map.setdefault(norm, norm)
            if not as_of_winus.empty and "발행사명" in as_of_winus.columns:
                for raw in as_of_winus["발행사명"].dropna().unique():
                    norm = normalize_issuer_name(raw)
                    options_map.setdefault(norm, norm)

            if options_map:
                st.caption(
                    "채권 스프레드·신용등급 트리거·위너스 익스포저 중 하나라도 정보가 있는 발행사는 모두 목록에 나옵니다. "
                    "일부 소스에 정보가 없으면 해당 항목에서만 '찾지 못했습니다'로 표시됩니다."
                )
                pick_issuer = st.selectbox(
                    "발행사 선택", options=sorted(options_map.values()), key="detail_page_issuer"
                )
                if pick_issuer:
                    norm_pick = normalize_issuer_name(pick_issuer)

                    if issuer_col_name and not day_spread.empty:
                        live_normalized_spread = day_spread[issuer_col_name].apply(normalize_issuer_name)
                        issuer_spread = day_spread[live_normalized_spread == norm_pick]
                        spread_date_used = (
                            day_spread["업데이트일자"].iloc[0] if "업데이트일자" in day_spread.columns and not day_spread.empty
                            else query_date
                        )
                    else:
                        issuer_spread = pd.DataFrame()
                        spread_date_used = query_date

                    date_note = "" if spread_date_used == query_date else f" — {spread_date_used} 기준(가장 가까운 이전 데이터)"
                    st.markdown(f"**{pick_issuer} — 채권 스프레드{date_note}**")
                    if issuer_spread.empty:
                        st.caption("이 발행사에 대한 채권 스프레드 정보를 찾지 못했습니다 (그날 해당 채권이 없었을 수 있습니다).")
                    else:
                        st.dataframe(
                            issuer_spread.drop(columns=["업데이트일자"]),
                            use_container_width=True, hide_index=True
                        )

                    if not as_of_trigger.empty and "원본발행사명" in as_of_trigger.columns:
                        # 저장 당시 얼어붙은 '발행사(정규화)' 값 대신, 원본발행사명을 지금 시점의
                        # 최신 별칭(관리자 설정에서 추가된 것 포함)으로 다시 계산해서 매칭한다.
                        live_normalized = as_of_trigger["원본발행사명"].apply(normalize_issuer_name)
                        issuer_trigger = as_of_trigger[live_normalized == norm_pick]
                        st.markdown(f"**{pick_issuer} — 신용등급 변동 트리거 (해당 시점 기준)**")
                        if issuer_trigger.empty:
                            st.caption("이 발행사에 대한 트리거 정보를 찾지 못했습니다 (회사명 표기 차이일 수 있습니다).")
                        else:
                            it = issuer_trigger.copy()
                            if "신평사" in it.columns:
                                it["신평사"] = it["신평사"].map(RATER_ABBREV).fillna(it["신평사"])

                            actual_header = "실제수치"
                            if "기준월" in it.columns:
                                labels = it["기준월"].dropna().unique()
                                if len(labels) > 0:
                                    actual_header = f"실제 수치({labels[0]})"

                            if "충족여부" not in it.columns and {"실제수치", "연산자", "값", "방향"} <= set(it.columns):
                                it["충족여부"] = it.apply(trigger_signal, axis=1)

                            if "실제수치" in it.columns:
                                it = it.rename(columns={"실제수치": actual_header})

                            display_cols = [c for c in
                                ["원본발행사명", "신평사", "기준일", "등급", "등급전망",
                                 "지표명", actual_header, "충족여부",
                                 "방향", "원문", "특이조건"]
                                if c in it.columns]
                            st.dataframe(
                                it[display_cols],
                                use_container_width=True, hide_index=True
                            )

                    if not as_of_winus.empty and "발행사명" in as_of_winus.columns:
                        live_normalized_winus = as_of_winus["발행사명"].apply(normalize_issuer_name)
                        issuer_winus = as_of_winus[live_normalized_winus == norm_pick]
                        st.markdown(f"**{pick_issuer} — 익스포저 현황 (해당 시점 기준)**")
                        if issuer_winus.empty:
                            st.caption("이 발행사에 대한 위너스 익스포저 정보를 찾지 못했습니다 (회사명 표기 차이일 수 있습니다).")
                        else:
                            winus_cols = [c for c in
                                ["검토의견", "내부등급", "투자한도", "잔여한도",
                                 "회사채잔액", "CP잔액", "CD잔액", "RP잔액", "정기예금잔액"]
                                if c in issuer_winus.columns]
                            st.dataframe(
                                issuer_winus[winus_cols],
                                use_container_width=True, hide_index=True
                            )

                    st.markdown(f"**{pick_issuer} — DART 기업개황 및 재무제표**")
                    if not DART_API_KEY:
                        st.caption("DART API 키가 설정되어 있지 않아 이 정보를 조회할 수 없습니다.")
                    else:
                        # 1) corp_code 확보: 별칭 표에 저장된 법인고유번호를 우선 사용 (빠르고 정확함)
                        corp_code = None
                        alias_full_for_detail = load_issuer_aliases_full()
                        if (
                            not alias_full_for_detail.empty
                            and "표준명" in alias_full_for_detail.columns
                            and "법인고유번호" in alias_full_for_detail.columns
                        ):
                            match_row = alias_full_for_detail[alias_full_for_detail["표준명"] == norm_pick]
                            if not match_row.empty:
                                stored_code = str(match_row.iloc[0]["법인고유번호"]).strip()
                                if stored_code:
                                    corp_code = stored_code

                        # 2) 별칭 표에 없으면 DART 전체 목록에서 실시간 매칭
                        if not corp_code:
                            with st.spinner("DART에서 기업 코드를 찾는 중입니다..."):
                                corp_map = load_corp_code_map(DART_API_KEY)
                                corp_code = find_corp_code(pick_issuer, corp_map) or find_corp_code(norm_pick, corp_map)

                        if not corp_code:
                            st.caption(
                                "DART에서 이 발행사의 기업 코드를 찾지 못했습니다. "
                                "관리자 탭에서 법인고유번호를 직접 등록해두시면 더 정확하게 찾을 수 있습니다."
                            )
                        else:
                            with st.spinner("DART에서 기업개황을 조회하는 중입니다..."):
                                overview = fetch_company_overview(corp_code, DART_API_KEY)

                            st.caption("기업개황")
                            if overview:
                                overview_rows = [
                                    ("대표자명", overview.get("ceo_nm")),
                                    ("법인구분", overview.get("corp_cls")),
                                    ("법인등록번호", overview.get("jurir_no")),
                                    ("사업자등록번호", overview.get("bizr_no")),
                                    ("주소", overview.get("adres")),
                                    ("홈페이지", overview.get("hm_url")),
                                    ("설립일", overview.get("est_dt")),
                                    ("결산월", overview.get("acc_mt")),
                                ]
                                st.dataframe(
                                    pd.DataFrame(overview_rows, columns=["항목", "값"]),
                                    use_container_width=True, hide_index=True
                                )
                            else:
                                st.caption("기업개황 정보를 가져오지 못했습니다.")

                            st.caption("재무제표")
                            unit = st.selectbox(
                                "금액 단위", options=["원", "천원", "백만원", "억원", "십억원", "조원"],
                                index=3, key="fin_unit_select"
                            )

                            with st.spinner("DART에서 재무제표를 조회하는 중입니다 (여러 보고서를 확인하느라 다소 걸릴 수 있습니다)..."):
                                full_report = fetch_full_financial_report(corp_code, DART_API_KEY)

                            if not full_report:
                                st.caption("재무제표 정보를 가져오지 못했습니다.")
                            else:
                                # --- 1) 재무상태표: 총자산/총부채/자기자본 ---
                                st.markdown(f"**재무상태표** ({unit})")
                                bs = full_report.get("BS")
                                bs_labels = full_report["meta"].get("BS_라벨", ["전전기", "전기", "최신"])
                                if bs:
                                    bs_rows = []
                                    for metric in ["총자산", "총부채", "자기자본"]:
                                        d = bs.get(metric, {})
                                        prev2 = d.get(bs_labels[0])
                                        prev1 = d.get(bs_labels[1])
                                        cur = d.get("최신")
                                        diff, pct = fmt_unit_change(cur, prev1, unit)
                                        bs_rows.append({
                                            "항목": metric,
                                            bs_labels[0]: fmt_unit(prev2, unit),
                                            bs_labels[1]: fmt_unit(prev1, unit),
                                            bs_labels[2]: fmt_unit(cur, unit),
                                            f"증감({bs_labels[1]} 대비)": diff,
                                            "증감률": pct,
                                        })
                                    st.dataframe(pd.DataFrame(bs_rows), use_container_width=True, hide_index=True)
                                else:
                                    st.caption("재무상태표 정보를 가져오지 못했습니다.")

                                # --- 2) 연도별 손익계산서 (올해누적 vs 전년동기누적 증감 포함) ---
                                st.markdown(f"**연도별 손익계산서** ({unit})")
                                annual = full_report.get("연간손익")
                                if annual:
                                    labels = [annual["전전년도"]["라벨"], annual["전년도"]["라벨"], annual["올해누적"]["라벨"]]
                                    prior_ytd_label = annual["전년동기누적"]["라벨"]
                                    st.caption("열: " + " / ".join(labels) + f" (비교기준: {prior_ytd_label})")
                                    annual_rows = []
                                    for metric in ["매출액", "영업이익", "당기순이익"]:
                                        cur_ytd = annual["올해누적"]["값"].get(metric)
                                        prior_ytd_val = annual["전년동기누적"]["값"].get(metric)
                                        diff, pct = fmt_unit_change(cur_ytd, prior_ytd_val, unit)
                                        annual_rows.append({
                                            "항목": metric,
                                            labels[0]: fmt_unit(annual["전전년도"]["값"].get(metric), unit),
                                            labels[1]: fmt_unit(annual["전년도"]["값"].get(metric), unit),
                                            labels[2]: fmt_unit(cur_ytd, unit),
                                            f"증감({prior_ytd_label} 대비)": diff,
                                            "증감률": pct,
                                        })
                                    st.dataframe(pd.DataFrame(annual_rows), use_container_width=True, hide_index=True)
                                else:
                                    st.caption("연도별 손익계산서 정보를 가져오지 못했습니다.")

                                # --- 3) 분기별 손익계산서: 작년 1~4분기 + 올해 존재하는 분기 전부 (+ 최신 vs 직전분기 증감) ---
                                st.markdown(f"**분기별 손익계산서** ({unit})")
                                quarterly = full_report.get("분기손익")
                                q_series = []  # 차트용: [(라벨, 매출액, 영업이익), ...] 시간순
                                if quarterly:
                                    last_yr = quarterly["직전년도"]["연도"]
                                    q_map = quarterly["직전년도"]["분기"]
                                    this_yr = quarterly["올해"]["연도"]
                                    this_q_map = quarterly["올해"]["분기"]

                                    this_year_qs = [q for q in ["1분기", "2분기", "3분기"] if q in this_q_map]
                                    col_labels = [f"{last_yr} {q}" for q in ["1분기", "2분기", "3분기", "4분기"]] \
                                                 + [f"{this_yr} {q}" for q in this_year_qs]
                                    st.caption("열: " + " / ".join(col_labels))

                                    # 시간순 정렬된 (라벨, 값dict) 시퀀스 — 최신분기 vs 직전분기 비교 및 차트에 사용
                                    ordered_seq = [(f"{last_yr} {q}", q_map.get(q, {})) for q in ["1분기", "2분기", "3분기", "4분기"]]
                                    ordered_seq += [(f"{this_yr} {q}", this_q_map.get(q, {})) for q in this_year_qs]

                                    latest_label, latest_vals = ordered_seq[-1]
                                    prev_label, prev_vals = ordered_seq[-2] if len(ordered_seq) >= 2 else (None, {})

                                    q_rows = []
                                    for metric in ["매출액", "영업이익", "당기순이익"]:
                                        row = {"항목": metric}
                                        for q in ["1분기", "2분기", "3분기", "4분기"]:
                                            row[f"{last_yr} {q}"] = fmt_unit(q_map.get(q, {}).get(metric), unit)
                                        for q in this_year_qs:
                                            row[f"{this_yr} {q}"] = fmt_unit(this_q_map.get(q, {}).get(metric), unit)
                                        diff, pct = fmt_unit_change(latest_vals.get(metric), prev_vals.get(metric), unit)
                                        row[f"증감({prev_label} 대비)" if prev_label else "증감(직전분기 대비)"] = diff
                                        row["증감률"] = pct
                                        q_rows.append(row)
                                    st.dataframe(pd.DataFrame(q_rows), use_container_width=True, hide_index=True)
                                    st.caption(
                                        f"'{latest_label}'을 직전분기인 '{prev_label}'과 비교했습니다. "
                                        "4분기는 사업보고서(연간) 수치에서 3분기 누적을 뺀 값이라, 결산 조정 등으로 실제 공시치와 약간 다를 수 있습니다."
                                    )

                                    q_series = [(label, v.get("매출액"), v.get("영업이익")) for label, v in ordered_seq]
                                else:
                                    st.caption("분기별 손익계산서 정보를 가져오지 못했습니다.")

                                # --- 4) 매출액·영업이익(막대, 나란히) + 영업이익률(오른쪽 축 꺾은선) ---
                                if q_series:
                                    st.markdown(f"**매출액 · 영업이익 · 영업이익률 추이** ({unit})")
                                    divisor = UNIT_DIVISORS.get(unit, 1)
                                    quarter_order = [label for label, _, _ in q_series]

                                    bar_records = []
                                    margin_records = []
                                    for label, rev, op in q_series:
                                        rev_scaled = (rev / divisor) if rev is not None else None
                                        op_scaled = (op / divisor) if op is not None else None
                                        bar_records.append({"분기": label, "구분": "매출액", "금액": rev_scaled})
                                        bar_records.append({"분기": label, "구분": "영업이익", "금액": op_scaled})
                                        margin = (op / rev * 100) if (rev not in (None, 0) and op is not None) else None
                                        margin_records.append({"분기": label, "영업이익률": margin})

                                    bar_df = pd.DataFrame(bar_records)
                                    margin_df = pd.DataFrame(margin_records)

                                    bars = alt.Chart(bar_df).mark_bar().encode(
                                        x=alt.X("분기:N", sort=quarter_order, axis=alt.Axis(labelAngle=-45, title=None)),
                                        xOffset=alt.XOffset("구분:N", sort=["매출액", "영업이익"]),
                                        y=alt.Y("금액:Q", title=f"금액({unit})"),
                                        color=alt.Color("구분:N", scale=alt.Scale(
                                            domain=["매출액", "영업이익"], range=["#1f77b4", "#ff7f0e"]
                                        ), legend=alt.Legend(title=None)),
                                    )
                                    line = alt.Chart(margin_df).mark_line(
                                        color="#2ca02c", point=True, strokeWidth=2
                                    ).encode(
                                        x=alt.X("분기:N", sort=quarter_order),
                                        y=alt.Y("영업이익률:Q", title="영업이익률(%)"),
                                    )
                                    combo = alt.layer(bars, line).resolve_scale(y="independent").properties(height=400)
                                    st.altair_chart(combo, use_container_width=True)

# ==============================================================
# 페이지: 관리자 설정
# ==============================================================
elif page == "관리자 설정":
    if not st.session_state.get("is_admin"):
        st.warning("관리자 로그인이 필요합니다. 왼쪽 사이드바에서 로그인해주세요.")
        st.stop()
    st.header("관리자 설정")
    st.caption("발행사명 별칭 등 대시보드의 공용 설정을 관리하는 곳입니다.")

    if "GOOGLE_SHEET_ID" not in st.secrets or "gcp_service_account" not in st.secrets:
        st.warning("Google Sheets 연동 정보가 설정되어 있지 않아 별칭 관리를 사용할 수 없습니다.")
    else:
        st.subheader("발행사명 별칭 관리")
        st.caption(
            "회사 하나가 한 행, 각 소스(인포맥스·한신평·나신평·한기평·위너스·DART)가 각각 열입니다. "
            "여기서 수정한 내용은 Google Sheets에 즉시 저장되며, 모든 사용자의 대시보드에 바로 반영됩니다."
        )

        alias_full_df = load_issuer_aliases_full()

        search_canon = st.text_input(
            "표준명으로 검색", key="alias_search_canon"
        )
        view_df = alias_full_df
        if search_canon and not alias_full_df.empty and "표준명" in alias_full_df.columns:
            view_df = alias_full_df[alias_full_df["표준명"].str.contains(search_canon, na=False)]

        if not view_df.empty:
            st.dataframe(view_df, use_container_width=True, hide_index=True)
        else:
            st.caption("등록된 별칭이 아직 없습니다." if alias_full_df.empty else "검색 결과가 없습니다.")

        # ------------------------------------------------------------
        # 엑셀로 다운로드 / 엑셀 업로드로 일괄 저장
        # ------------------------------------------------------------
        st.markdown("---")
        st.markdown("**엑셀로 내려받아 수정 후 일괄 반영**")
        col_dl, col_ul = st.columns(2)

        with col_dl:
            st.caption("현재 별칭 표를 엑셀로 내려받습니다.")
            excel_bytes = build_alias_excel_bytes(
                alias_full_df if not alias_full_df.empty else pd.DataFrame(columns=ALIAS_HEADER)
            )
            st.download_button(
                "발행사별칭.xlsx 다운로드",
                data=excel_bytes,
                file_name="발행사별칭.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="alias_excel_download"
            )

        with col_ul:
            st.caption("엑셀에서 수정한 파일을 업로드하면, 시트 전체를 이 내용으로 덮어씁니다.")
            alias_upload = st.file_uploader(
                "수정한 발행사별칭.xlsx 업로드", type=["xlsx"], key="alias_excel_upload"
            )
            if alias_upload is not None:
                if st.button("업로드한 내용으로 전체 저장", key="alias_excel_save_btn"):
                    try:
                        uploaded_df = pd.read_excel(alias_upload)
                        save_alias_excel_upload(uploaded_df)
                        st.success(f"{len(uploaded_df)}개 회사 정보로 별칭 표를 갱신했습니다.")

                        verify_df = load_issuer_aliases_full()
                        if len(verify_df) == len(uploaded_df):
                            st.info(f"✅ 저장 직후 재확인: 클라우드 시트에서 {len(verify_df)}행을 정상적으로 읽어왔습니다.")
                        else:
                            st.error(
                                f"⚠️ 재확인해보니 {len(verify_df)}행이 조회됩니다 (기대값 {len(uploaded_df)}행). "
                                "GOOGLE_SHEET_ID 설정을 확인해주세요."
                            )
                    except Exception as e:
                        st.error(f"저장 중 오류가 발생했습니다: {e}")

        st.warning(
            "⚠️ 엑셀 업로드는 시트 **전체를 덮어씁니다**. 다운로드한 최신 파일을 기준으로 수정해서 "
            "올려주세요 (다른 사람이 그 사이에 추가한 내용이 있다면 먼저 새로 다운로드하고 반영해서 병합하세요)."
        )

        # ------------------------------------------------------------
        # 별칭 표 검증
        # ------------------------------------------------------------
        st.markdown("---")
        st.subheader("별칭 표 검증")
        st.caption(
            "표준명이 중복된 행, 같은 표기가 서로 다른 표준명에 잘못 매핑된 충돌, "
            "표준명에 법인접미어(주식회사/(주)/㈜)가 그대로 남아있는 경우를 찾아줍니다."
        )
        if st.button("검증 실행", key="validate_alias_btn"):
            st.session_state["alias_validation_df"] = load_issuer_aliases_full()

        if "alias_validation_df" in st.session_state:
            check_df = st.session_state["alias_validation_df"]
            validation = validate_alias_table(check_df)

            n_dup = check_df.loc[
                check_df["표준명"].astype(str).str.strip().duplicated(keep=False)
                & (check_df["표준명"].astype(str).str.strip() != "")
            ]["표준명"].nunique() if not check_df.empty else 0

            if validation["dup_rows"].empty and not validation["conflicts"] and validation["suffix_rows"].empty:
                st.success("문제가 발견되지 않았습니다. 👍")
            else:
                if not validation["dup_rows"].empty:
                    st.warning(f"표준명이 중복된 회사 {n_dup}개 ({len(validation['dup_rows'])}행)")
                    st.dataframe(validation["dup_rows"], use_container_width=True, hide_index=True)

                if validation["conflicts"]:
                    st.warning(f"같은 표기가 서로 다른 표준명에 매핑된 충돌 {len(validation['conflicts'])}건")
                    conflict_df = pd.DataFrame(
                        [{"표기": a, "충돌하는 표준명들": " / ".join(c)} for a, c in validation["conflicts"]]
                    )
                    st.dataframe(conflict_df, use_container_width=True, hide_index=True)

                if not validation["suffix_rows"].empty:
                    st.warning(f"표준명에 법인접미어가 남아있는 행 {len(validation['suffix_rows'])}건")
                    st.dataframe(validation["suffix_rows"], use_container_width=True, hide_index=True)

                st.markdown("**자동 정리안 미리보기**")
                st.caption(
                    "표준명에서 법인접미어를 제거하고, 중복행은 각 열의 값을 하나로 합칩니다 "
                    "(충돌 나는 값은 먼저 나온 값을 채택합니다). 아래 미리보기를 확인 후 적용하세요."
                )
                cleaned_preview = build_cleaned_alias_table(check_df)
                st.caption(f"정리 전 {len(check_df)}행 → 정리 후 {len(cleaned_preview)}행")
                st.dataframe(cleaned_preview, use_container_width=True, hide_index=True)

                if st.button("이 정리안으로 저장", key="apply_cleaned_alias_btn"):
                    try:
                        save_alias_excel_upload(cleaned_preview)
                        del st.session_state["alias_validation_df"]
                        st.success(f"정리된 내용({len(cleaned_preview)}행)을 클라우드 시트에 저장했습니다.")

                        verify_df = load_issuer_aliases_full()
                        if len(verify_df) == len(cleaned_preview):
                            st.info(f"✅ 저장 직후 재확인: 클라우드 시트에서 {len(verify_df)}행을 정상적으로 읽어왔습니다.")
                        else:
                            st.error(
                                f"⚠️ 저장은 오류 없이 끝났지만, 재확인해보니 {len(verify_df)}행이 조회됩니다 "
                                f"(기대값 {len(cleaned_preview)}행). GOOGLE_SHEET_ID가 다른 시트를 "
                                "가리키고 있을 수 있으니 Secrets 설정을 확인해주세요."
                            )
                        sheet_id = st.secrets.get("GOOGLE_SHEET_ID", "")
                        if sheet_id:
                            st.caption(
                                f"현재 연결된 시트: https://docs.google.com/spreadsheets/d/{sheet_id}/edit"
                            )
                    except Exception as e:
                        st.error(f"저장 중 오류가 발생했습니다: {e}")

        # ------------------------------------------------------------
        # 화면에서 직접 추가/삭제 (엑셀 없이 빠르게)
        # ------------------------------------------------------------
        st.markdown("---")
        st.markdown("**화면에서 바로 추가 (한 개씩)**")
        col_x, col_y, col_s, col_z = st.columns([2, 2, 2, 1])
        with col_x:
            new_alias = st.text_input("이 소스에서 쓰는 표기", key="new_alias_input")
        with col_y:
            new_canonical = st.text_input("표준명 (통일해서 쓸 이름)", key="new_canonical_input")
        with col_s:
            new_source = st.selectbox("소스(열)", options=ALIAS_SOURCE_COLUMNS, key="new_alias_source")
        with col_z:
            st.write("")
            st.write("")
            if st.button("추가/갱신", key="add_alias_btn"):
                if not new_alias.strip() or not new_canonical.strip():
                    st.error("표기와 표준명을 모두 입력해주세요.")
                else:
                    try:
                        upsert_issuer_alias(new_canonical.strip(), new_source, new_alias.strip())
                        st.success(f"'{new_canonical}' 행의 '{new_source}' 칸에 '{new_alias}' 저장되었습니다.")
                        st.rerun()
                    except Exception as e:
                        st.error(f"저장 중 오류가 발생했습니다: {e}")

        st.markdown("**화면에서 바로 추가 (여러 개 한 번에, 같은 소스 기준)**")
        st.caption("한 줄에 하나씩 '표기 = 표준명' 형식으로 입력하세요. 표준명 행이 이미 있으면 해당 소스 칸만 채웁니다.")
        col_bulk1, col_bulk2 = st.columns([3, 1])
        with col_bulk1:
            bulk_text = st.text_area(
                "일괄 입력", height=120,
                placeholder="디비증권 = DB증권\n엘지전자 = LG전자\n씨제이씨지브이 = CJ CGV",
                key="bulk_alias_text"
            )
        with col_bulk2:
            bulk_source = st.selectbox("소스(열, 전체 적용)", options=ALIAS_SOURCE_COLUMNS, key="bulk_alias_source")
            st.write("")
            if st.button("일괄 추가/갱신", key="bulk_add_alias_btn"):
                pairs = []
                for line in (bulk_text or "").splitlines():
                    if "=" in line:
                        a, c = line.split("=", 1)
                        a, c = a.strip(), c.strip()
                        if a and c:
                            pairs.append((a, c))
                if not pairs:
                    st.error("형식에 맞는 줄이 없습니다 ('표기 = 표준명' 형식으로 입력해주세요).")
                else:
                    try:
                        upsert_issuer_aliases_bulk(pairs, bulk_source)
                        st.success(f"{len(pairs)}건 반영되었습니다.")
                        st.rerun()
                    except Exception as e:
                        st.error(f"일괄 저장 중 오류가 발생했습니다: {e}")

        if not alias_full_df.empty:
            st.markdown("**회사 행 전체 삭제**")
            col_p, col_q = st.columns([3, 1])
            with col_p:
                del_target = st.selectbox(
                    "삭제할 표준명 선택",
                    options=sorted(alias_full_df["표준명"].dropna().unique()),
                    key="del_company_select"
                )
            with col_q:
                st.write("")
                st.write("")
                if st.button("행 삭제", key="del_company_btn"):
                    try:
                        delete_issuer_company(del_target)
                        st.success(f"'{del_target}' 행이 삭제되었습니다.")
                        st.rerun()
                    except Exception as e:
                        st.error(f"삭제 중 오류가 발생했습니다: {e}")

        st.markdown("---")
        st.subheader("DART 기업명·법인고유번호 일괄 매칭")
        st.caption(
            "전자공시시스템(DART)의 전체 기업 목록(회사명 + 고유번호)을 한 번에 받아와서, "
            "별칭 표의 표준명과 자동으로 매칭해 'DART' 열과 '법인고유번호' 열을 함께 채웁니다. "
            "DART 쪽 기업 목록은 하루 단위로 캐시되니, 최신 목록이 필요하면 버튼을 다시 누르면 됩니다."
        )
        if not DART_API_KEY:
            st.warning(
                "DART API 키가 설정되어 있지 않습니다. Streamlit Cloud Secrets에 "
                "DART_API_KEY = \"발급받은키\" 를 등록해주세요."
            )
        else:
            if st.button("DART 기업 목록 불러와서 매칭", key="dart_match_btn"):
                with st.spinner("DART 기업 목록을 받아오는 중입니다 (처음 조회 시 몇 초 걸릴 수 있습니다)..."):
                    corp_map = load_corp_code_map(DART_API_KEY)
                st.caption(f"DART에 등록된 전체 기업 수: {len(corp_map):,}개")

                alias_df_for_dart = load_issuer_aliases_full()
                if alias_df_for_dart.empty or "표준명" not in alias_df_for_dart.columns:
                    st.info("별칭 표에 등록된 회사가 아직 없습니다. 먼저 위쪽에서 회사를 추가해주세요.")
                else:
                    updated = alias_df_for_dart.copy()
                    if "법인고유번호" not in updated.columns:
                        updated["법인고유번호"] = ""
                    matched_count, filled_rows = 0, []
                    for idx, row in updated.iterrows():
                        canon = str(row.get("표준명", "")).strip()
                        existing_dart = str(row.get("DART", "")).strip() if "DART" in updated.columns else ""
                        existing_code = str(row.get("법인고유번호", "")).strip()
                        if not canon or (existing_dart and existing_code):
                            continue  # 이름·코드 둘 다 이미 있으면 건드리지 않음
                        matched_name = find_dart_name(canon, corp_map)
                        if matched_name:
                            if not existing_dart:
                                updated.at[idx, "DART"] = matched_name
                            if not existing_code:
                                updated.at[idx, "법인고유번호"] = corp_map.get(matched_name, "")
                            matched_count += 1
                            filled_rows.append({
                                "표준명": canon, "DART 매칭명": matched_name,
                                "법인고유번호": corp_map.get(matched_name, "")
                            })

                    st.session_state["dart_match_preview"] = updated
                    st.success(f"{matched_count}개 회사에 새로 DART 회사명/법인고유번호를 채웠습니다 (기존에 값이 있던 칸은 건드리지 않았습니다).")
                    if filled_rows:
                        st.dataframe(pd.DataFrame(filled_rows), use_container_width=True, hide_index=True)

            if "dart_match_preview" in st.session_state:
                st.markdown("**매칭 결과 미리보기 (DART·법인고유번호 열 반영됨)**")
                st.dataframe(st.session_state["dart_match_preview"], use_container_width=True, hide_index=True, height=300)
                if st.button("이 매칭 결과 저장", key="save_dart_match_btn"):
                    try:
                        save_alias_excel_upload(st.session_state["dart_match_preview"])
                        del st.session_state["dart_match_preview"]
                        st.success("DART 매칭 결과를 저장했습니다.")
                        st.rerun()
                    except Exception as e:
                        st.error(f"저장 중 오류가 발생했습니다: {e}")

        st.markdown("---")
        st.subheader("미매칭 발행사 자동 탐지")
        st.caption(
            "저장된 채권 스프레드 이력(인포맥스 발행사 목록)을 기준으로, "
            "신용등급 트리거·위너스 이력의 발행사명 중 지금 별칭으로도 매칭이 안 되는 이름들을 찾아줍니다. "
            "이 목록에 뜨는 회사는 '발행사별 상세 보기' 드롭다운에도 같은 회사가 두 번 나타날 수 있습니다."
        )
        if st.button("미매칭 발행사 찾기", key="find_unmatched_btn"):
            with st.spinner("비교하는 중입니다..."):
                spread_hist_admin = read_history("채권스프레드_이력")
                trigger_hist_admin = read_history("신용등급트리거_이력")
                winus_hist_admin = read_history("위너스_이력")

            if spread_hist_admin.empty or "발행사" not in spread_hist_admin.columns:
                st.warning("채권 스프레드 이력이 없습니다. 먼저 인포맥스 파일을 저장해주세요.")
            else:
                infomax_set = set(spread_hist_admin["발행사"].dropna().unique())
                infomax_set_normalized = {normalize_issuer_name(n) for n in infomax_set}

                unmatched_parts = []

                if not trigger_hist_admin.empty and "원본발행사명" in trigger_hist_admin.columns:
                    th = trigger_hist_admin.copy()
                    th["정규화결과"] = th["원본발행사명"].apply(normalize_issuer_name)
                    th_unmatched = th[~th["정규화결과"].isin(infomax_set_normalized)]
                    unmatched_parts.append(
                        th_unmatched[["신평사", "원본발행사명", "정규화결과"]].rename(
                            columns={"신평사": "소스", "원본발행사명": "원본표기"}
                        )
                    )

                if not winus_hist_admin.empty and "발행사명" in winus_hist_admin.columns:
                    wh = winus_hist_admin.copy()
                    wh["정규화결과"] = wh["발행사명"].apply(normalize_issuer_name)
                    wh_unmatched = wh[~wh["정규화결과"].isin(infomax_set_normalized)]
                    wh_unmatched = wh_unmatched[["발행사명", "정규화결과"]].rename(columns={"발행사명": "원본표기"})
                    wh_unmatched.insert(0, "소스", "위너스(WINUS)")
                    unmatched_parts.append(wh_unmatched)

                if not unmatched_parts:
                    st.warning("비교할 트리거·위너스 이력이 없습니다. 먼저 각 탭에서 저장해주세요.")
                else:
                    unmatched = pd.concat(unmatched_parts, ignore_index=True)
                    if unmatched.empty:
                        st.success("모든 발행사명이 인포맥스 목록과 매칭됩니다. 👍")
                    else:
                        summary = (
                            unmatched.groupby(["소스", "원본표기", "정규화결과"])
                            .size().reset_index(name="건수")
                            .sort_values(["소스", "원본표기"])
                        )
                        st.warning(f"인포맥스 목록과 매칭되지 않는 발행사명 {len(summary)}건을 찾았습니다.")
                        st.dataframe(summary, use_container_width=True, hide_index=True)
                        st.caption(
                            "위 '정규화결과' 값이 실제로는 인포맥스의 어떤 발행사명과 같은 회사인지 확인해서, "
                            "위쪽 '새 별칭 추가'에 '정규화결과 = 인포맥스표기명' 형태로 등록해주세요."
                        )
