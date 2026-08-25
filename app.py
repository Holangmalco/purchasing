import streamlit as st

# 1. 전체 페이지 기본 설정 (가장 먼저 와야 함)
st.set_page_config(page_title="구매부서 통합 대시보드", page_icon="🏫", layout="wide")

# 2. 메인 대시보드 화면 만들기
def main_dashboard():
    st.markdown("<h2 style='text-align: center; color: #004797;'>🏫 구매부서 통합 업무 대시보드</h2>", unsafe_allow_html=True)
    st.markdown("---")
    
    # 기획하신 3가지 핵심 대시보드 영역으로 분할
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.subheader("🛒 일반구매 현황")
        st.info("💡 향후 노션 DB API 및 스크래퍼 연동 예정")
        # 가상 데이터 (나중에 노션 DB에서 긁어와서 연결)
        st.metric(label="이번 주 신규 구매요청", value="12 건", delta="3 건 증가")
        
    with col2:
        st.subheader("📜 입찰 캘린더 (예정)")
        st.success("💡 담당자 데이터 연동 및 일정표 구현 예정")
        # 캘린더 기능 뼈대 (현재는 날짜 선택기 수준이지만 추후 고도화 가능)
        st.date_input("입찰 주요 일정 확인", value="today")
        
    with col3:
        st.subheader("🏢 자산관리 현황")
        st.warning("💡 담당자와 세부 지표 협의 후 추가 예정")
        # 가상 데이터
        st.metric(label="이번 달 신규 자산 등재", value="158 건", delta="지난달 대비 20% 상승")

    st.write("")
    st.write("")
    st.markdown("#### 📌 툴 사용 가이드")
    st.write("1. **구매/계약 툴:** 상단 메뉴에서 툴 선택 후 ERP 텍스트를 복사해 넣으면 AI가 서류를 교차 검증합니다.")
    st.write("2. **ERP 스크래퍼:** 버튼 한 번으로 결재 대기함의 구매요청 문서를 솎아냅니다.")
    st.write("3. **재산증감 보고:** rERP 엑셀 데이터를 업로드하면 월별 증감통계가 자동 생성됩니다.")

# 3. 각 페이지(파일)를 깔끔하게 정의하고 아이콘 달기
p_main = st.Page(main_dashboard, title="메인 대시보드", icon="🏠", default=True)
p1 = st.Page("pages/1_구매추출기_3.py", title="구매/계약 추출기", icon="💰")
p2 = st.Page("pages/2_스크래퍼.py", title="ERP 스크래퍼", icon="🤖")
p3 = st.Page("pages/3_재산증감보고.py", title="재산증감 보고", icon="📊")

# 4. ★핵심★ 상단 네비게이션으로 묶어서 실행
pg = st.navigation(
    {"": [p_main], "업무 툴": [p1, p2, p3]},
    position="top" # 여기서 왼쪽 사이드바 대신 '상단'으로 설정됩니다!
)
pg.run()