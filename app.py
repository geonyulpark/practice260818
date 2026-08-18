import streamlit as st
import pandas as pd
import requests
import zipfile
import io
import datetime
import xml.etree.ElementTree as ET

st.set_page_config(page_title="크레딧 모니터링 대시보드", layout="wide")

st.title("📊 크레딧 모니터링 대시보드")
st.caption("인포맥스 채권 수익률 파일을 업로드하면 '공모/무보증' 발행사만 자동으로 필터링해 보여줍니다.")

# ==============================================================
# DART API 관련 함수
# ==============================================================
DART_API_KEY = st.secrets.get("DART_API_KEY", "")

@st.cache_data(ttl=60 * 60 * 24)  # 24시간 캐시 (고유번호 목록은 하루 한 번이면 충분)
def load_corp_code_map(api_key: str):
    """DART 고유번호(corp_code) 전체 목록을 받아 {회사명: corp_code} 딕셔너리로 반환."""
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
    """발행사명을 DART 고유번호와 매칭. 정확일치 → 접미사 제거 재시도 → 부분일치 순."""
    if issuer_name in corp_map:
        return corp_map[issuer_name]

    for suffix in ["지주", "홀딩스", "㈜", "(주)"]:
        candidate = issuer_name.replace(suffix, "").strip()
        if candidate in corp_map:
            return corp_map[candidate]

    # 부분일치 (발행사명이 짧아 여러 후보가 잡힐 수 있어 마지막 수단으로만 사용)
    candidates = [name for name in corp_map if issuer_name in name or name in issuer_name]
    if len(candidates) == 1:
        return corp_map[candidates[0]]
    return None


@st.cache_data(ttl=60 * 60 * 24)  # 24시간 캐시 (하루 한 번 갱신되면 충분)
def fetch_financials(corp_code: str, api_key: str):
    """가장 최근 분기 → 반기 → 사업보고서 순으로 매출액/영업이익/당기순이익 조회."""
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
        for fs_div in ["CFS", "OFS"]:  # 연결재무제표 우선, 없으면 개별재무제표
            params = {
                "crtfc_key": api_key,
                "corp_code": corp_code,
                "bsns_year": str(year),
                "reprt_code": reprt_code,
                "fs_div": fs_div,
            }
            try:
                resp = requests.get(url, params=params, timeout=15)
                data = resp.json()
            except Exception:
                continue

            if data.get("status") != "000":
                continue

            revenue = op_income = net_income = None
            for row in data.get("list", []):
                if row.get("sj_div") != "IS":
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
                return {
                    "매출액": revenue, "영업이익": op_income, "당기순이익": net_income,
                    "재무제표기준일": label
                }

    return {"매출액": None, "영업이익": None, "당기순이익": None, "재무제표기준일": "조회 실패"}


# ==============================================================
# 1. 파일 업로드
# ==============================================================
uploaded_file = st.file_uploader(
    "인포맥스 엑셀 파일을 업로드하세요 (예: 4788-YYYYMMDD.xlsx)",
    type=["xlsx", "xls"]
)

if uploaded_file is not None:
    # ------------------------------------------------------------
    # 2. 데이터 읽기
    #    인포맥스 파일은 3번째 줄(0-indexed: 2)이 실제 헤더입니다.
    # ------------------------------------------------------------
    try:
        df = pd.read_excel(uploaded_file, header=2)
    except Exception as e:
        st.error(f"파일을 읽는 중 오류가 발생했습니다: {e}")
        st.stop()

    rating_col = df.columns[4]      # E열: 채권그룹(신용등급 포함)
    issuer_col = '발행사'

    if rating_col not in df.columns or issuer_col not in df.columns:
        st.error("예상한 열 구조와 다릅니다. 파일 형식을 확인해주세요.")
        st.stop()

    # ------------------------------------------------------------
    # 3. "공모/무보증" 필터링
    # ------------------------------------------------------------
    mask_rating = df[rating_col].astype(str).str.contains("공모/무보증", na=False)
    filtered = df[mask_rating].copy()
    filtered = filtered[filtered[issuer_col] != filtered[rating_col]]

    # ------------------------------------------------------------
    # 4. 보여줄 열 선택 + 산업대분류 매핑
    # ------------------------------------------------------------
    maturity_cols = ["3M", "6M", "9M", "1Y", "3Y", "5Y"]
    maturity_cols = [c for c in maturity_cols if c in filtered.columns]

    industry_col = "업종분류(소)"
    industry_code_col = "업종코드"

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

    display_cols = [rating_col, industry_col, issuer_col] + maturity_cols
    if industry_code_col in filtered.columns:
        display_cols = [industry_code_col] + display_cols
    display_cols = [c for c in display_cols if c in filtered.columns]

    result = filtered[display_cols].rename(
        columns={rating_col: "신용등급그룹", industry_col: "업종구분", industry_code_col: "업종코드"}
    )
    result["신용등급그룹"] = (
        result["신용등급그룹"].astype(str).str.replace("공모/무보증", "", regex=False).str.strip()
    )
    result["업종구분"] = result["업종구분"].fillna("미분류")

    if "업종코드" in result.columns:
        result["산업대분류"] = result["업종코드"].astype(str).str[0].map(KSIC_LARGE).fillna("미분류")
        result = result.drop(columns=["업종코드"])
        cols = ["산업대분류"] + [c for c in result.columns if c != "산업대분류"]
        result = result[cols]

    # ------------------------------------------------------------
    # 5. 필터 (산업대분류 → 신용등급 → 발행사)
    # ------------------------------------------------------------
    st.subheader("필터")
    col0, col1, col2 = st.columns(3)

    with col0:
        industry_options = sorted(result["산업대분류"].unique())
        selected_industries = st.multiselect(
            "산업 대분류 선택", options=industry_options, default=industry_options
        )

    with col1:
        rating_options = sorted(
            result[result["산업대분류"].isin(selected_industries)]["신용등급그룹"].unique()
        )
        selected_ratings = st.multiselect(
            "신용등급 선택", options=rating_options, default=rating_options
        )

    issuer_pool = result[
        result["산업대분류"].isin(selected_industries)
        & result["신용등급그룹"].isin(selected_ratings)
    ][issuer_col].sort_values().unique()

    with col2:
        selected_issuers = st.multiselect(
            "발행사 선택 (비워두면 전체 표시)", options=issuer_pool
        )

    view = result[
        result["산업대분류"].isin(selected_industries)
        & result["신용등급그룹"].isin(selected_ratings)
    ]
    if selected_issuers:
        view = view[view[issuer_col].isin(selected_issuers)]

    # ------------------------------------------------------------
    # 5-1. DART 재무제표 연동 (선택적 — API 호출이라 다소 시간이 걸립니다)
    # ------------------------------------------------------------
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

            # 열 순서 재배치: 만기 수익률 다음에 매출액/영업이익/당기순이익, 재무제표기준일은 맨 마지막
            other_cols = [c for c in view.columns if c not in financial_cols]
            ordered_cols = other_cols + ["매출액", "영업이익", "당기순이익", "재무제표기준일"]
            view = view[ordered_cols]

    # ------------------------------------------------------------
    # 6. 결과 표시
    # ------------------------------------------------------------
    st.subheader(f"공모/무보증 발행사 목록 ({len(view)}개)")
    st.dataframe(view, use_container_width=True, hide_index=True)

    st.subheader("신용등급별 발행사 수")
    st.bar_chart(view["신용등급그룹"].value_counts())

    # ------------------------------------------------------------
    # 등급별 만기별 평균 수익률
    # ------------------------------------------------------------
    if maturity_cols:
        st.subheader("등급별 만기별 평균 수익률")
        rating_order = ["AAA", "AA+", "AA0", "AA-", "A+", "A0", "A-",
                         "BBB+", "BBB0", "BBB-", "BB+", "BB0", "BB-"]
        avg_by_rating = result.groupby("신용등급그룹")[maturity_cols].mean().round(3)
        ordered = [r for r in rating_order if r in avg_by_rating.index]
        remaining = [r for r in avg_by_rating.index if r not in rating_order]
        avg_by_rating = avg_by_rating.loc[ordered + remaining]

        st.dataframe(avg_by_rating, use_container_width=True)
        st.caption("등급별 만기 수익률 곡선 (평균)")
        st.line_chart(avg_by_rating.T)

    # 만기별 수익률 비교 (선택한 발행사 기준)
    if maturity_cols:
        st.subheader("발행사별 만기 수익률 비교")
        pick = st.multiselect(
            "비교할 발행사 선택 (최대 10개 추천)",
            options=view[issuer_col].tolist(),
            default=view[issuer_col].tolist()[:5]
        )
        if pick:
            chart_data = view[view[issuer_col].isin(pick)].set_index(issuer_col)[maturity_cols]
            st.line_chart(chart_data.T)

    # CSV 다운로드
    st.download_button(
        "필터링된 결과 CSV로 다운로드",
        data=view.to_csv(index=False).encode("utf-8-sig"),
        file_name="공모무보증_발행사_필터결과.csv",
        mime="text/csv"
    )

else:
    st.info("왼쪽 상단에서 인포맥스 엑셀 파일을 업로드하면 대시보드가 표시됩니다.")
