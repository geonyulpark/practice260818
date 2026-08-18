import streamlit as st
import pandas as pd
import numpy as np

# ------------------------------------------------------------
# 연습용 샘플 크레딧 모니터링 대시보드
# 실제 데이터 대신 무작위 예시 데이터를 사용합니다.
# ------------------------------------------------------------

st.set_page_config(page_title="크레딧 모니터링 (연습용)", layout="wide")

st.title("📊 크레딧 모니터링 대시보드 (연습용)")
st.caption("Streamlit 배포 연습을 위한 샘플 앱입니다. 데이터는 실제가 아닌 예시입니다.")

# ------------------------------------------------------------
# 1. 샘플 데이터 생성
# ------------------------------------------------------------
np.random.seed(0)
companies = ["A건설", "B화학", "C항공", "D쇼핑", "E전자"]
ratings = ["AAA", "AA+", "AA", "AA-", "A+", "A", "A-", "BBB+"]

data = pd.DataFrame({
    "기업명": companies,
    "신용등급": np.random.choice(ratings, size=len(companies)),
    "부채비율(%)": np.random.randint(80, 250, size=len(companies)),
    "EBITDA마진(%)": np.random.randint(5, 30, size=len(companies)),
    "이자보상배율": np.round(np.random.uniform(0.5, 8.0, size=len(companies)), 2),
})

# ------------------------------------------------------------
# 2. 사이드바 필터
# ------------------------------------------------------------
st.sidebar.header("필터")
selected_companies = st.sidebar.multiselect(
    "기업 선택", options=companies, default=companies
)

filtered = data[data["기업명"].isin(selected_companies)]

# ------------------------------------------------------------
# 3. 테이블 표시
# ------------------------------------------------------------
st.subheader("기업별 주요 지표")
st.dataframe(filtered, use_container_width=True)

# ------------------------------------------------------------
# 4. 차트
# ------------------------------------------------------------
col1, col2 = st.columns(2)

with col1:
    st.subheader("부채비율 비교")
    st.bar_chart(filtered.set_index("기업명")["부채비율(%)"])

with col2:
    st.subheader("이자보상배율 비교")
    st.bar_chart(filtered.set_index("기업명")["이자보상배율"])

st.divider()
st.markdown("이 앱은 **연습용 예시**이며, 실제 신용평가 데이터와 무관합니다.")
