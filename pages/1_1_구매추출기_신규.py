import streamlit as st
import re
import requests
import os
from dotenv import load_dotenv

# 💡 [신규 추가] 로컬에서 텍스트를 바로 긁어오기 위한 라이브러리들
import pdfplumber 
from PIL import Image
import pytesseract

st.set_page_config(page_title="구매추출기 (단순추출)", page_icon="⚡", layout="wide")

# 1. 환경변수 및 국세청 API 세팅
load_dotenv()
NTS_API_KEY = os.getenv("NTS_API_KEY")

def check_nts_business(biz_no_str):
    clean_b = re.sub(r'[^\d]', '', str(biz_no_str))
    if len(clean_b) != 10: 
        return "⚠️ 번호오류"
    
    url = f"https://api.odcloud.kr/api/nts-businessman/v1/status?serviceKey={NTS_API_KEY}"
    headers = {"Content-Type": "application/json"}
    data = {"b_no": [clean_b]}
    
    try:
        res = requests.post(url, headers=headers, json=data, timeout=5)
        if res.status_code == 200:
            r_data = res.json().get('data', [])
            if r_data:
                stt_cd = r_data[0].get('b_stt_cd', '')
                if stt_cd == '01': return "✅ 계속사업자"
                elif stt_cd == '02': return "🚨 휴업자"
                elif stt_cd == '03': return "🚨 폐업자"
    except Exception as e:
        return f"⚠️ 조회실패(서버오류: {e})"
    return "⚠️ 조회실패"

# 2. 화면 UI
st.title("⚡ 구매추출기 (선택형 추출 모드)")
st.info("견적서를 로컬에서 빠르게 텍스트만 뽑아보거나, 필요시 AI를 호출해 정밀 분석할 수 있습니다.")

# 국세청 수동 조회부
with st.container(border=True):
    col1, col2 = st.columns([1, 2])
    with col1:
        raw_biz_no = st.text_input("🔍 사업자등록번호 입력 (숫자만 입력 시 자동 조회)", placeholder="예: 1234567890")
    with col2:
        if raw_biz_no:
            with st.spinner("국세청 상태 실시간 조회 중..."):
                status = check_nts_business(raw_biz_no)
                if "계속사업자" in status:
                    st.success(f"결과: [{raw_biz_no}] {status}")
                else:
                    st.error(f"결과: [{raw_biz_no}] {status}")

st.divider()

# 파일 업로드 및 선택적 추출부
st.write("### 📎 견적서 업로드 및 정보 추출")
uploaded_files = st.file_uploader("PDF 또는 이미지 형태의 견적서를 업로드하세요.", accept_multiple_files=True, type=['pdf', 'png', 'jpg', 'jpeg'])

if uploaded_files:
    # 💡 [신규 추가] 사용자가 원하는 추출 방식을 고를 수 있도록 버튼을 2개로 분리
    c_btn1, c_btn2 = st.columns(2)
    
    with c_btn1:
        btn_local = st.button("⚡ 빠르고 단순하게 텍스트만 추출 (로컬)", use_container_width=True)
    with c_btn2:
        btn_ai = st.button("🧠 제미나이 AI 정밀 분석 (API 호출/느림)", type="primary", use_container_width=True)

    # 1️⃣ 로컬 추출 버튼을 눌렀을 때 (API 미사용)
    if btn_local:
        with st.spinner("로컬 환경에서 텍스트를 긁어오고 있습니다... (약 1~2초 소요)"):
            for f in uploaded_files:
                st.markdown(f"#### 📄 파일명: {f.name}")
                extracted_text = ""
                
                try:
                    # PDF 파일일 경우 pdfplumber로 텍스트 추출
                    if f.name.lower().endswith('.pdf'):
                        with pdfplumber.open(f) as pdf:
                            for page in pdf.pages:
                                text = page.extract_text()
                                if text: extracted_text += text + "\n"
                    
                    # 이미지 파일일 경우 pytesseract(로컬 OCR)로 텍스트 추출
                    elif f.name.lower().endswith(('.png', '.jpg', '.jpeg')):
                        img = Image.open(f)
                        extracted_text = pytesseract.image_to_string(img, lang='kor+eng')
                    
                    if extracted_text.strip():
                        st.text_area("추출된 텍스트 원본", value=extracted_text, height=200, key=f.name)
                    else:
                        st.warning("텍스트를 추출하지 못했습니다. (이미지로 된 PDF일 수 있습니다.)")
                
                except Exception as e:
                    st.error(f"추출 중 오류 발생: {e}")

    # 2️⃣ AI 분석 버튼을 눌렀을 때 (기존 1_구매추출기.py의 로직을 간소화하여 연결할 예정)
    if btn_ai:
        st.info("이곳에 기존처럼 제미나이 API를 호출하여 JSON 형태로 뽑아주는 로직이 실행됩니다. (현재 뼈대만 구성됨)")