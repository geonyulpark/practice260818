import streamlit as st
import pandas as pd

st.set_page_config(page_title="크레딧 모니터링 대시보드", layout="wide")

st.title("📊 크레딧 모니터링 대시보드")
st.caption("인포맥스 채권 수익률 파일을 업로드하면 '공모/무보증' 발행사만 자동으로 필터링해 보여줍니다.")

# ------------------------------------------------------------
# 1. 파일 업로드
# ------------------------------------------------------------
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

    # E열 = 5번째 열(인덱스 4) = "채권그룹" 관련 열 (신용등급 정보 포함)
    rating_col = df.columns[4]      # 예: '채권그룹.1'
    issuer_col = '발행사'

    if rating_col not in df.columns or issuer_col not in df.columns:
        st.error("예상한 열 구조와 다릅니다. 파일 형식을 확인해주세요.")
        st.stop()

    # ------------------------------------------------------------
    # 3. "공모/무보증" 필터링
    # ------------------------------------------------------------
    mask_rating = df[rating_col].astype(str).str.contains("공모/무보증", na=False)
    filtered = df[mask_rating].copy()

    # 그룹 평균행(발행사명 == 등급그룹명) 제외 → 실제 발행사만 남김
    filtered = filtered[filtered[issuer_col] != filtered[rating_col]]

    # ------------------------------------------------------------
    # 4. 보여줄 열 선택
    #    - 15Y/20Y/30Y 제외
    #    - "평가사" 열 제외
    #    - 신용등급그룹은 "공모/무보증 " 접두어를 떼고 등급만 표시 (예: "A-")
    # ------------------------------------------------------------
    maturity_cols = ["3M", "6M", "9M", "1Y", "1.5Y", "2Y", "2.5Y",
                      "3Y", "4Y", "5Y", "7Y", "10Y"]
    maturity_cols = [c for c in maturity_cols if c in filtered.columns]

    display_cols = [rating_col, issuer_col] + maturity_cols
    display_cols = [c for c in display_cols if c in filtered.columns]

    result = filtered[display_cols].rename(columns={rating_col: "신용등급그룹"})
    result["신용등급그룹"] = (
        result["신용등급그룹"].astype(str).str.replace("공모/무보증", "", regex=False).str.strip()
    )

    # ------------------------------------------------------------
    # 5. 필터 (표 바로 위에 배치 — 등급 선택 + 발행사 선택)
    # ------------------------------------------------------------
    st.subheader("필터")
    col1, col2 = st.columns(2)

    with col1:
        rating_options = sorted(result["신용등급그룹"].unique())
        selected_ratings = st.multiselect(
            "신용등급 선택", options=rating_options, default=rating_options
        )

    # 선택된 등급에 해당하는 발행사만 옵션으로 제공
    issuer_pool = result[result["신용등급그룹"].isin(selected_ratings)][issuer_col].sort_values().unique()

    with col2:
        selected_issuers = st.multiselect(
            "발행사 선택 (비워두면 전체 표시)", options=issuer_pool
        )

    view = result[result["신용등급그룹"].isin(selected_ratings)]
    if selected_issuers:
        view = view[view[issuer_col].isin(selected_issuers)]

    # ------------------------------------------------------------
    # 6. 결과 표시
    # ------------------------------------------------------------
    st.subheader(f"공모/무보증 발행사 목록 ({len(view)}개)")
    st.dataframe(view, use_container_width=True, hide_index=True)

    # 등급별 개수 요약
    st.subheader("신용등급별 발행사 수")
    st.bar_chart(view["신용등급그룹"].value_counts())

    # ------------------------------------------------------------
    # 등급별 만기별 평균 수익률
    # ------------------------------------------------------------
    if maturity_cols:
        st.subheader("등급별 만기별 평균 수익률")

        # 등급 정렬 순서 (신용등급 표준 순서대로)
        rating_order = ["AAA", "AA+", "AA0", "AA-", "A+", "A0", "A-",
                         "BBB+", "BBB0", "BBB-", "BB+", "BB0", "BB-"]

        avg_by_rating = result.groupby("신용등급그룹")[maturity_cols].mean().round(3)
        # 정의된 순서대로 정렬하되, 목록에 없는 등급은 뒤에 붙임
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
