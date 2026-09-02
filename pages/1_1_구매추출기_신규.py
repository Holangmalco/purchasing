import streamlit as st
import re
import requests
import os
import io
import pandas as pd
from dotenv import load_dotenv

import pdfplumber 
from PIL import Image
import pytesseract
from pdf2image import convert_from_bytes

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

# 💡 [핵심] 정규표현식으로 찾은 단어의 '앞뒤 문맥'까지 같이 뽑아주는 함수
def find_with_context(text, pattern, pad=15):
    results = []
    clean_text = text.replace('\n', ' ') 
    
    for match in re.finditer(pattern, clean_text):
        start = max(0, match.start() - pad)
        end = min(len(clean_text), match.end() + pad)
        context_str = clean_text[start:end].strip()
        highlighted = context_str.replace(match.group(), f" 🎯[{match.group()}] ")
        results.append(highlighted)
        
    return list(set(results))

st.title("⚡ 구매추출기 (선택형 추출 모드)")
st.info("다양한 형태의 견적서(PDF, 이미지, 엑셀)를 로컬에서 즉시 텍스트로 변환하고 핵심 정보를 뽑아냅니다.")

# 1. 구매/계약 기본 정보 선택
st.markdown("#### 📂 1. 구매/계약 기본 정보 선택")
c_type1, c_type2 = st.columns(2)
with c_type1:
    req_type = st.radio("구매 유형", ["📦 물품", "🛠️ 용역", "🏗️ 공사"], horizontal=True, key="new_req_type")
with c_type2:
    con_type = st.radio("계약 방식", ["비교견적", "수의계약"], horizontal=True, key="new_con_type")

st.divider()

# 2. ERP 텍스트 복사 및 업로드란 (2단 분할)
col_in1, col_in2 = st.columns(2)

with col_in1:
    with st.container(border=True):
        st.markdown("#### 📝 2. ERP 텍스트 복사")
        if "p_req" not in st.session_state: st.session_state["p_req"] = ""
        if "p_fin" not in st.session_state: st.session_state["p_fin"] = ""
        if "p_itm" not in st.session_state: st.session_state["p_itm"] = ""
        
        raw_req = st.text_area("1️⃣ 구매요청서 전체", height=120, value=st.session_state["p_req"])
        st.session_state["p_req"] = raw_req
        
        raw_fin = st.text_area("2️⃣ 재원내역 행", height=60, value=st.session_state["p_fin"])
        st.session_state["p_fin"] = raw_fin
        
        raw_itm = st.text_area("3️⃣ 물품/용역내역 행", height=80, value=st.session_state["p_itm"])
        st.session_state["p_itm"] = raw_itm

with col_in2:
    with st.container(border=True):
        st.markdown("#### 📎 3. 만능 견적서 업로드 & 국세청 조회")
        
        raw_biz_no = st.text_input("🔍 사업자등록번호 수동 입력 (실시간 조회)", placeholder="예: 1234567890")
        if raw_biz_no:
            with st.spinner("국세청 상태 조회 중..."):
                status = check_nts_business(raw_biz_no)
                if "계속사업자" in status:
                    st.success(f"결과: [{raw_biz_no}] {status}")
                else:
                    st.error(f"결과: [{raw_biz_no}] {status}")
                    
        st.markdown("---")
        # 엑셀, PDF, 이미지 모두 지원하는 업로더
        uploaded_files = st.file_uploader("PDF, 이미지, 엑셀(.xlsx) 모두 지원", accept_multiple_files=True, type=['pdf', 'png', 'jpg', 'jpeg', 'xlsx'])

st.divider()

# 3. 추출 실행 버튼부
if uploaded_files:
    c_btn1, c_btn2 = st.columns(2)
    with c_btn1:
        btn_local = st.button("⚡ 빠르고 단순하게 텍스트만 추출 (로컬)", use_container_width=True)
    with c_btn2:
        btn_ai = st.button("🧠 제미나이 AI 정밀 분석 (추후 연결)", type="primary", use_container_width=True)

    if btn_local:
        with st.spinner("로컬 환경에서 파일을 분석 중입니다..."):
            for f in uploaded_files:
                st.markdown(f"#### 📄 파일명: {f.name}")
                extracted_text = ""
                
                try:
                    # [케이스 1] 엑셀 파일 (.xlsx)
                    if f.name.lower().endswith('.xlsx'):
                        df = pd.read_excel(f)
                        extracted_text = df.to_string(index=False)
                        st.success("✅ 엑셀 데이터 추출 완료")

                    # [케이스 2] PDF 및 스캔본 방어
                    elif f.name.lower().endswith('.pdf'):
                        with pdfplumber.open(f) as pdf:
                            for page in pdf.pages:
                                text = page.extract_text()
                                if text: extracted_text += text + "\n"
                        
                        # 텍스트가 너무 적으면 스캔본으로 판단하여 이미지로 쪼개기
                        if len(extracted_text.strip()) < 50:
                            st.warning("⚠️ 스캔본 PDF 감지됨 ➔ 이미지 분석 모드로 우회합니다.")
                            f.seek(0) 
                            images = convert_from_bytes(f.read())
                            extracted_text = ""
                            for img in images:
                                extracted_text += pytesseract.image_to_string(img, lang='kor+eng') + "\n"
                        else:
                            st.success("✅ 일반 PDF 추출 완료")

                    # [케이스 3] 이미지 파일 및 PDF 변환기
                    elif f.name.lower().endswith(('.png', '.jpg', '.jpeg')):
                        img = Image.open(f)
                        extracted_text = pytesseract.image_to_string(img, lang='kor+eng')
                        st.success("✅ 이미지 OCR 추출 완료")
                        
                        pdf_buffer = io.BytesIO()
                        img.convert('RGB').save(pdf_buffer, format="PDF")
                        st.download_button(
                            label="📥 이 이미지를 PDF로 변환하여 다운로드",
                            data=pdf_buffer.getvalue(),
                            file_name=f"{f.name.split('.')[0]}_변환.pdf",
                            mime="application/pdf",
                            key=f"dl_{f.name}"
                        )
                    
                    # 💡 파일 형태와 무관하게 텍스트가 뽑혔다면 문맥 추출 로직 가동
                    if extracted_text.strip():
                        biz_pattern = r'\d{3}-\d{2}-\d{5}'
                        biz_results = find_with_context(extracted_text, biz_pattern, pad=10)
                        
                        date_pattern = r'\d{4}[-./년]\s?\d{1,2}[-./월]\s?\d{1,2}[일]?'
                        date_results = find_with_context(extracted_text, date_pattern, pad=15)
                        
                        col_a, col_b = st.columns(2)
                        with col_a:
                            st.markdown("**🏢 감지된 사업자번호 주변 텍스트**")
                            if biz_results:
                                for r in biz_results: st.info(f"...{r}...")
                            else:
                                st.write("감지 안 됨")
                                
                        with col_b:
                            st.markdown("**📅 감지된 날짜 주변 텍스트**")
                            if date_results:
                                for r in date_results: st.info(f"...{r}...")
                            else:
                                st.write("감지 안 됨")
                        
                        st.markdown("---")
                        st.text_area("📄 추출된 전체 텍스트 원본", value=extracted_text, height=150, key=f"{f.name}_txt")
                        
                    else:
                        st.error("텍스트를 추출하지 못했습니다.")
                
                except Exception as e:
                    st.error(f"추출 중 오류 발생: {e}")

    if btn_ai:
        st.info("이곳에 기존처럼 제미나이 API를 호출하여 JSON 형태로 뽑아주는 로직이 들어갈 예정입니다.")