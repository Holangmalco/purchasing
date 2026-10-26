import streamlit as st

# 1. 전체 페이지 기본 설정
st.set_page_config(page_title="구매부서 통합시스템", page_icon="🏫", layout="wide")

# 2. 현재 활발히 사용 중인 핵심 페이지만 연결 (홈 예시 화면 및 재산증감은 숨김 처리)
p_main = st.Page("pages/1_구매추출기.py", title="구매부서 통합시스템", icon="🏫", default=True)

# -------------------------------------------------------------------
# [보관] 향후 추가 예정 화면 (추후 기능 완성 시 다시 활성화 가능)
# -------------------------------------------------------------------
# def main_dashboard(): ...
# p_home = st.Page(main_dashboard, title="홈", icon="🏠")
# p_asset = st.Page("pages/3_재산증감보고.py", title="재산증감", icon="📊")

# 3. 상단 불필요한 네비게이션 숨김 (내부 메뉴로 직관적 전환)
pg = st.navigation(
    [p_main],
    position="hidden"
)
pg.run()