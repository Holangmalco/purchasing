import streamlit as st
import streamlit.components.v1 as components

# 1. 전체 페이지 기본 설정
st.set_page_config(page_title="구매부서 통합 대시보드", page_icon="🏫", layout="wide")

# 2. 메인 대시보드 화면 만들기
def main_dashboard():
    st.markdown("<h2 style='text-align: center; color: #004797;'>🏫 구매부서 통합 업무 대시보드</h2>", unsafe_allow_html=True)
    st.markdown("---")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.subheader("🛒 일반구매 현황 (예시)")
        st.info("💡 향후 노션 DB API 및 스크래퍼 연동 예정")
        st.metric(label="이번 주 신규 구매요청", value="N 건", delta="계약 중 N'건")
        
    with col2:
        st.subheader("📜 입찰 캘린더")
        calendar_url = "https://calendar.google.com/calendar/embed?src=hagrid15303%40gmail.com&ctz=Asia%2FSeoul"
        components.iframe(calendar_url, width=450, height=400)
        
    with col3:
        st.subheader("🏢 자산관리 현황 (예시)")
        st.warning("💡 담당자와 협의 후 추가 예정")
        st.metric(label="필요한 내용", value="N 건", delta="특이사항")

    st.write("")
    st.write("")
    st.markdown("#### 📌 툴 사용 가이드")
    st.write("1. **구매/계약:** ERP 결재 대기 텍스트와 서류 PDF를 첨부하면 Gemini AI가 품명, 금액, 견적서, 비교견적서를 교차 검증하고 결과 보고서를 생성합니다.")
    st.write("2. **재산증감 보고:** rERP 자산내역 엑셀 파일을 업로드하면 월별 증감통계와 엑셀 서식을 자동 생성합니다.")

# 3. 각 페이지 정의 
p_main = st.Page(main_dashboard, title="홈", icon="🏠", default=True)
p1 = st.Page("pages/1_구매추출기.py", title="구매/계약", icon="💰")
p2 = st.Page("pages/3_재산증감보고.py", title="재산증감", icon="📊")

# 4. 상단 네비게이션
pg = st.navigation(
    [p_main, p1, p2],
    position="top"
)
pg.run()