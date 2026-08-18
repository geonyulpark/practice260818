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
    # ------------------------------------------------------------
    maturity_cols = ["3M", "6M", "9M", "1Y", "1.5Y", "2Y", "2.5Y",
                      "3Y", "4Y", "5Y", "7Y", "10Y", "15Y", "20Y", "30Y"]
    maturity_cols = [c for c in maturity_cols if c in filtered.columns]

    display_cols = [rating_col, issuer_col, "평가사"] + maturity_cols
    display_cols = [c for c in display_cols if c in filtered.columns]

    result = filtered[display_cols].rename(columns={rating_col: "신용등급그룹"})

    # ------------------------------------------------------------
    # 5. 사이드바 필터 (등급/발행사 검색)
    # ------------------------------------------------------------
    st.sidebar.header("추가 필터")
    rating_options = sorted(result["신용등급그룹"].unique())
    selected_ratings = st.sidebar.multiselect(
        "신용등급 선택", options=rating_options, default=rating_options
    )
    search_issuer = st.sidebar.text_input("발행사명 검색")

    view = result[result["신용등급그룹"].isin(selected_ratings)]
    if search_issuer:
        view = view[view[issuer_col].str.contains(search_issuer, na=False)]

    # ------------------------------------------------------------
    # 6. 결과 표시
    # ------------------------------------------------------------
    st.subheader(f"공모/무보증 발행사 목록 ({len(view)}개)")
    st.dataframe(view, use_container_width=True, hide_index=True)

    # 등급별 개수 요약
    st.subheader("신용등급별 발행사 수")
    st.bar_chart(view["신용등급그룹"].value_counts())

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
