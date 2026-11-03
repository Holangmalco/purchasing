import streamlit as st
import pandas as pd
import datetime

# -----------------------------------------------------------------------------
# 1. 페이지 기본 설정 (와이드 모드, 간결한 타이틀)
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="구매부서 디자인 미리보기",
    page_icon="🎨",
    layout="wide",
    initial_sidebar_state="expanded"
)

# -----------------------------------------------------------------------------
# 2. 토스 & 애플 스타일 전역 CSS 인젝션
# -----------------------------------------------------------------------------
st.markdown("""
<style>
/* [1] Pretendard 웹폰트 로드 & 전역 적용 */
@import url("https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.9/dist/web/static/pretendard.css");

* {
    font-family: 'Pretendard', -apple-system, BlinkMacSystemFont, "Apple SD Gothic Neo", sans-serif !important;
    letter-spacing: -0.025em;
}

/* [2] 전체 배경색 (소프트 미색 그레이) */
.stApp {
    background-color: #F8F9FA !important;
    color: #191F28;
}

/* [3] 스트림릿 기본 상단 무지개 바 및 불필요한 헤더 숨김 */
div[data-testid="stDecoration"] {
    display: none !important;
}
header[data-testid="stHeader"] {
    background: rgba(248, 249, 250, 0.85) !important;
    backdrop-filter: blur(12px) !important;
}

/* [4] 메인 콘텐츠 여백 최적화 (넓고 편안한 좌우 여백) */
.block-container {
    padding-top: 2.2rem !important;
    padding-bottom: 5rem !important;
    padding-left: 5rem !important;
    padding-right: 5rem !important;
    max-width: 100% !important;
    width: 100% !important;
    box-sizing: border-box !important;
}

@media (max-width: 1400px) {
    .block-container {
        padding-left: 3.5rem !important;
        padding-right: 3.5rem !important;
    }
}

@media (max-width: 992px) {
    .block-container {
        padding-left: 1.5rem !important;
        padding-right: 1.5rem !important;
    }
}

/* [5] 사이드바 토스 스타일 */
section[data-testid="stSidebar"] {
    background-color: #FFFFFF !important;
    border-right: 1px solid #ECEFF2 !important;
}
section[data-testid="stSidebar"] .block-container {
    padding-top: 2.5rem !important;
    padding-left: 1.5rem !important;
    padding-right: 1.5rem !important;
}

/* 사이드바 메뉴 캡슐 버튼 */
section[data-testid="stSidebar"] div[role="radiogroup"] {
    gap: 6px;
}
section[data-testid="stSidebar"] div[role="radiogroup"] label {
    background-color: transparent;
    padding: 12px 16px;
    border-radius: 12px;
    font-size: 15px;
    font-weight: 500;
    color: #4E5968;
    transition: all 0.18s cubic-bezier(0.4, 0, 0.2, 1);
    margin-bottom: 4px;
    border: 1px solid transparent;
}
section[data-testid="stSidebar"] div[role="radiogroup"] label:hover {
    background-color: #F2F4F6;
    color: #191F28;
}
section[data-testid="stSidebar"] div[role="radiogroup"] label[data-checked="true"] {
    background-color: #E8F3FF !important;
    color: #1B64DA !important;
    font-weight: 700 !important;
    border: 1px solid #D0E4FF !important;
}

/* [6] 플로팅 카드 컨테이너 (애플/토스 시그니처 카드) */
div[data-testid="stVerticalBlock"] > div[data-testid="stVerticalBlockBorderWrapper"] {
    background-color: #FFFFFF !important;
    border-radius: 20px !important;
    border: 1px solid #EAECEF !important;
    box-shadow: 0 4px 20px rgba(0, 0, 0, 0.03) !important;
    padding: 24px 28px !important;
    margin-bottom: 1.2rem !important;
    transition: transform 0.2s ease, box-shadow 0.2s ease;
}
div[data-testid="stVerticalBlock"] > div[data-testid="stVerticalBlockBorderWrapper"]:hover {
    box-shadow: 0 8px 28px rgba(0, 0, 0, 0.05) !important;
}

/* [7] 버튼 (토스 시그니처 라운드 버튼) */
div.stButton > button {
    border-radius: 14px !important;
    padding: 12px 22px !important;
    font-size: 15px !important;
    font-weight: 600 !important;
    border: 1px solid #E5E8EB !important;
    background-color: #FFFFFF !important;
    color: #333D4B !important;
    box-shadow: 0 1px 2px rgba(0,0,0,0.03) !important;
    transition: all 0.18s cubic-bezier(0.4, 0, 0.2, 1) !important;
}
div.stButton > button:hover {
    background-color: #F2F4F6 !important;
    border-color: #D1D6DB !important;
    color: #191F28 !important;
    transform: translateY(-1px);
}
/* Primary 버튼 (토스 블루) */
div.stButton > button[kind="primary"] {
    background-color: #3182F6 !important;
    color: #FFFFFF !important;
    border: none !important;
    box-shadow: 0 4px 14px rgba(49, 130, 246, 0.3) !important;
}
div.stButton > button[kind="primary"]:hover {
    background-color: #1B64DA !important;
    box-shadow: 0 6px 18px rgba(49, 130, 246, 0.4) !important;
    transform: translateY(-2px);
}

/* [8] 텍스트 입력창 & 셀렉트박스 (배경 대비 극대화) */
div[data-baseweb="input"], 
div[data-baseweb="select"] > div, 
div[data-baseweb="textarea"] {
    border-radius: 12px !important;
    border: 1.5px solid #D1D6DB !important;
    background-color: #F8F9FA !important;
    transition: all 0.2s cubic-bezier(0.4, 0, 0.2, 1);
}

div[data-baseweb="input"]:hover, 
div[data-baseweb="select"] > div:hover, 
div[data-baseweb="textarea"]:hover {
    border-color: #B0B8C1 !important;
    background-color: #F2F4F6 !important;
}

div[data-baseweb="input"]:focus-within, 
div[data-baseweb="select"] > div:focus-within, 
div[data-baseweb="textarea"]:focus-within {
    border-color: #3182F6 !important;
    background-color: #FFFFFF !important;
    box-shadow: 0 0 0 3.5px rgba(49, 130, 246, 0.15) !important;
}

div[data-baseweb="input"] input, 
div[data-baseweb="textarea"] textarea {
    background-color: transparent !important;
    color: #191F28 !important;
    font-weight: 500 !important;
}

/* [9] 탭 (Apple 세그먼트 컨트롤 스타일) */
div[data-testid="stTabs"] div[role="tablist"] {
    background-color: #EFEFEF;
    padding: 4px;
    border-radius: 14px;
    gap: 4px;
}
div[data-testid="stTabs"] button[role="tab"] {
    border-radius: 10px;
    padding: 8px 18px;
    font-size: 14px;
    font-weight: 500;
    color: #6B7684;
    border: none !important;
    transition: all 0.18s ease;
}
div[data-testid="stTabs"] button[role="tab"][aria-selected="true"] {
    background-color: #FFFFFF !important;
    color: #191F28 !important;
    font-weight: 700 !important;
    box-shadow: 0 2px 8px rgba(0, 0, 0, 0.08) !important;
}

/* [10] 배지 & 칩 스타일 */
.toss-badge {
    display: inline-flex;
    align-items: center;
    padding: 4px 10px;
    border-radius: 8px;
    font-size: 12px;
    font-weight: 600;
    letter-spacing: -0.01em;
}
.badge-blue { background: #E8F3FF; color: #1B64DA; }
.badge-green { background: #E6F8F0; color: #059669; }
.badge-gray { background: #F2F4F6; color: #4E5968; }
.badge-orange { background: #FFF4E5; color: #D97706; }

/* 요약 통계 카드 스타일 */
.metric-card {
    background: #FFFFFF;
    border-radius: 18px;
    padding: 22px;
    border: 1px solid #ECEFF2;
    box-shadow: 0 2px 10px rgba(0,0,0,0.02);
}
.metric-label {
    font-size: 13px;
    font-weight: 600;
    color: #8B95A1;
    margin-bottom: 6px;
}
.metric-value {
    font-size: 26px;
    font-weight: 700;
    color: #191F28;
    margin-bottom: 4px;
}
.metric-sub {
    font-size: 12px;
    color: #4E5968;
}
</style>
""", unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# 3. 사이드바 구성
# -----------------------------------------------------------------------------
with st.sidebar:
    st.markdown("<div style='padding: 8px 4px 18px 4px;'>", unsafe_allow_html=True)
    st.markdown("<h2 style='font-size: 20px; font-weight: 700; margin: 0; color: #191F28;'>구매부서 유틸리티</h2>", unsafe_allow_html=True)
    st.markdown("<span class='toss-badge badge-blue' style='margin-top: 6px;'>🎨 디자인 프리뷰</span>", unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)

    selected_menu = st.radio(
        "메뉴 선택",
        ["💰 구매추출기", "🤝 구매계약분석", "🕒 분석내역 불러오기", "⚙️ 시스템 설정"],
        index=0,
        label_visibility="collapsed"
    )

    st.markdown("<div style='height: 30px;'></div>", unsafe_allow_html=True)
    with st.container(border=True):
        st.markdown("<span style='font-size: 13px; font-weight: 600; color: #8B95A1;'>안내사항</span>", unsafe_allow_html=True)
        st.markdown(
            "<p style='font-size: 12px; color: #4E5968; line-height: 1.5; margin: 6px 0 0 0;'>"
            "본 화면은 토스 & 애플 스타일 UI 데모입니다.<br>기존 업무용 앱과 독립적으로 동작하며 데이터에 영향을 주지 않습니다."
            "</p>",
            unsafe_allow_html=True
        )

# -----------------------------------------------------------------------------
# 4. 헤더 영역 (시원한 여백 & 정돈된 타이포그래피)
# -----------------------------------------------------------------------------
col_header_left, col_header_right = st.columns([8, 4])
with col_header_left:
    st.markdown("<h1 style='font-size: 28px; font-weight: 700; color: #191F28; margin: 0 0 6px 0;'>구매추출기</h1>", unsafe_allow_html=True)
    st.markdown("<p style='font-size: 15px; color: #8B95A1; margin: 0;'>견적서 및 세금계산서 문서를 분석하여 발주서와 계약 정보를 자동으로 생성합니다.</p>", unsafe_allow_html=True)

with col_header_right:
    st.markdown("<div style='display: flex; justify-content: flex-end; align-items: center; gap: 8px; height: 100%;'>", unsafe_allow_html=True)
    st.markdown("<span class='toss-badge badge-green'>● AI 엔진 정상 (gemini-2.5-flash)</span>", unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)

st.markdown("<div style='height: 20px;'></div>", unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# 5. 토스 대시보드형 상단 지표 카드 (KPI Summary)
# -----------------------------------------------------------------------------
col_m1, col_m2, col_m3, col_m4 = st.columns(4)

with col_m1:
    st.markdown("""
    <div class='metric-card'>
        <div class='metric-label'>이번 달 총 집행 건수</div>
        <div class='metric-value'>28 <span style='font-size: 16px; font-weight: 500;'>건</span></div>
        <div class='metric-sub'><span style='color: #059669; font-weight: 600;'>▲ 4건</span> 지난달 대비 증가</div>
    </div>
    """, unsafe_allow_html=True)

with col_m2:
    st.markdown("""
    <div class='metric-card'>
        <div class='metric-label'>누적 계약 체결액</div>
        <div class='metric-value'>4,820 <span style='font-size: 16px; font-weight: 500;'>만원</span></div>
        <div class='metric-sub'>수의계약 82% · 일반계약 18%</div>
    </div>
    """, unsafe_allow_html=True)

with col_m3:
    st.markdown("""
    <div class='metric-card'>
        <div class='metric-label'>서명 대기 문서</div>
        <div class='metric-value' style='color: #1B64DA;'>3 <span style='font-size: 16px; font-weight: 500; color: #191F28;'>건</span></div>
        <div class='metric-sub'><span class='toss-badge badge-blue'>eformsign 연동</span></div>
    </div>
    """, unsafe_allow_html=True)

with col_m4:
    st.markdown("""
    <div class='metric-card'>
        <div class='metric-label'>국세청 사업자 검증률</div>
        <div class='metric-value' style='color: #059669;'>100 <span style='font-size: 16px; font-weight: 500; color: #191F28;'>%</span></div>
        <div class='metric-sub'>최근 24시간 휴폐업 의심 없음</div>
    </div>
    """, unsafe_allow_html=True)

st.markdown("<div style='height: 24px;'></div>", unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# 6. 메인 작업 영역 (Apple 탭 세그먼트 + 토스 플로팅 카드)
# -----------------------------------------------------------------------------
tab_extract, tab_preview, tab_history = st.tabs(["📄 문서 업로드 & 분석", "📋 결과 검토 & 다운로드", "📊 과거 분석 데이터"])

with tab_extract:
    st.markdown("<div style='height: 8px;'></div>", unsafe_allow_html=True)
    
    # [Card 1] 파일 업로드 & 기본 설정 카드
    with st.container(border=True):
        st.markdown("<h3 style='font-size: 18px; font-weight: 700; color: #191F28; margin: 0 0 16px 0;'>문서 첨부 및 기본 설정</h3>", unsafe_allow_html=True)
        
        c_up1, c_up2 = st.columns([7, 5])
        with c_up1:
            uploaded_files = st.file_uploader(
                "견적서 / 사업자등록증 / 관련 서류",
                type=["pdf", "png", "jpg", "jpeg", "docx"],
                accept_multiple_files=True,
                help="PDF 문서 및 영수증 이미지 파일을 최대 10개까지 업로드할 수 있습니다."
            )
        with c_up2:
            st.markdown("<div style='padding-left: 8px;'>", unsafe_allow_html=True)
            doc_category = st.selectbox(
                "계약 및 문서 유형 선택",
                ["일반 교비/산단 (1천만원 미만)", "일반 교비/산단 (1천만원 이상)", "신임교원 정착연구비용 (단독 템플릿)"]
            )
            tax_included = st.checkbox("부가세(VAT) 포함 금액으로 계산", value=True)
            auto_nts = st.checkbox("국세청 실시간 사업자 상태 자동 조회", value=True)
            st.markdown("</div>", unsafe_allow_html=True)
            
        st.markdown("<div style='height: 12px;'></div>", unsafe_allow_html=True)
        
        # 액션 버튼
        c_btn1, c_btn2, c_sp = st.columns([3, 2, 5])
        with c_btn1:
            start_btn = st.button("🚀 AI 분석 시작하기", type="primary", use_container_width=True)
        with c_btn2:
            reset_btn = st.button("초기화", use_container_width=True)

    # [Card 2] 인터랙티브 파싱 결과 카드 (시뮬레이션 예시)
    with st.container(border=True):
        st.markdown("<h3 style='font-size: 18px; font-weight: 700; color: #191F28; margin: 0 0 16px 0;'>추출 데이터 즉시 검토</h3>", unsafe_allow_html=True)
        
        c_f1, c_f2, c_f3 = st.columns(3)
        with c_f1:
            st.text_input("업체명", value="(주)알파테크놀로지")
            st.text_input("사업자등록번호", value="123-45-67890")
            st.markdown("<span class='toss-badge badge-green'>✅ 계속사업자 (국세청 확인완료)</span>", unsafe_allow_html=True)
        with c_f2:
            st.text_input("대표자명", value="홍길동")
            st.text_input("업체 담당자 이메일", value="sales@alphatech.co.kr")
            st.text_input("연락처", value="02-1234-5678")
        with c_f3:
            st.text_input("총 공급가액", value="₩ 8,800,000")
            st.text_input("실사용자(학내 교직원)", value="김교수 (sejong.ac.kr)")
            st.text_input("납품 기한", value="2026-11-20")

        st.markdown("<div style='height: 16px;'></div>", unsafe_allow_html=True)
        st.divider()
        
        c_down1, c_down2, c_empty = st.columns([3, 3, 4])
        with c_down1:
            st.button("📥 발주서(eformsign) DOCX 다운로드", use_container_width=True)
        with c_down2:
            st.button("⚡ eformsign 즉시 서명 요청 전송", type="primary", use_container_width=True)

with tab_preview:
    with st.container(border=True):
        st.markdown("<h3 style='font-size: 18px; font-weight: 700; color: #191F28; margin: 0 0 12px 0;'>품목 및 세부 내역</h3>", unsafe_allow_html=True)
        sample_df = pd.DataFrame([
            {"품번": 1, "품목명": "고성능 GPU 딥러닝 워크스테이션", "규격": "RTX 4090 24GB / 64GB RAM", "수량": 1, "단가": "8,000,000", "금액": "8,000,000"},
            {"품번": 2, "품목명": "고해상도 4K 모니터 32인치", "규격": "IPS 패널 144Hz", "수량": 1, "단가": "800,000", "금액": "800,000"}
        ])
        st.dataframe(sample_df, use_container_width=True, hide_index=True)

with tab_history:
    with st.container(border=True):
        st.markdown("<h3 style='font-size: 18px; font-weight: 700; color: #191F28; margin: 0 0 12px 0;'>최근 분석 내역 히스토리</h3>", unsafe_allow_html=True)
        history_df = pd.DataFrame([
            {"분석일시": "2026-11-03 10:45", "업체명": "(주)알파테크놀로지", "금액": "8,800,000원", "상태": "문서 생성 완료", "담당자": "김교수"},
            {"분석일시": "2026-11-02 16:20", "업체명": "(주)바이오사이언스", "금액": "14,500,000원", "상태": "계약 검토 완료", "담당자": "이교수"},
            {"분석일시": "2026-11-01 11:10", "업체명": "한국시약상사", "금액": "2,400,000원", "상태": "서명 완료", "담당자": "박연구원"}
        ])
        st.dataframe(history_df, use_container_width=True, hide_index=True)

