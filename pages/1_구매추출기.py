import streamlit as st
import os

# 파이썬에게 "지금 이 Purchasing.py 파일이 들어있는 폴더를 모든 작업의 '본진'으로 삼아라!"라고 명령합니다.
# os.chdir(os.path.dirname(os.path.abspath(__file__)))

import re
import datetime
import pandas as pd
import time
import json
import difflib
from google import genai
from google.genai import types
from io import BytesIO
from docxtpl import DocxTemplate
from dotenv import load_dotenv
import requests

# ---------------------------------------------------------
# 💡 [신규] 국세청 사업자 휴폐업 조회 엔진
# ---------------------------------------------------------
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
        log(f"국세청 API 호출 에러: {e}")
        pass
    
    return "⚠️ 조회실패(서버오류)"

# ==========================================
# 1. 시스템 설정 및 환경 변수
# ==========================================
DEBUG = False 
MAX_UPLOAD_FILES = 10
MAX_FILE_SIZE_MB = 10
AI_MODEL = "gemini-2.5-flash"
SESSION_AI_KEY = "ai_res"
SESSION_REQ_KEY = "req_analysis_done"
REQUIRED_FILES = ["template_under10mil.docx", "template_over10mil.docx"]

st.set_page_config(page_title="구매추출기 V2", page_icon="💰", layout="wide")

load_dotenv()
API_KEY = os.getenv("GEMINI_API_KEY")
NTS_API_KEY = os.getenv("NTS_API_KEY")

missing_files = [f for f in REQUIRED_FILES if not os.path.exists(f)]
if missing_files:
    st.error(f"🚨 필수 파일 누락으로 실행 불가: {', '.join(missing_files)}")
    st.stop()

if not API_KEY or not NTS_API_KEY:
    st.error("🚨 환경변수(.env)에 API 키(구글 또는 국세청)가 누락되었습니다.")
    st.stop()

ai_client = genai.Client(api_key=API_KEY)

INSPECTOR_INFO = {
    "name": "경제용",
    "phone": "02-3408-3031",
    "email": "rudwpdyd@sejong.ac.kr"
}

# ==========================================
# 2. 공통 유틸리티
# ==========================================
def log(msg):
    if DEBUG: print(f"[LOG] {msg}")

def safe_int(val):
    try: return int(re.sub(r'[^\d]', '', str(val)))
    except: return 0

def safe_json_load(text):
    try: return json.loads(text)
    except:
        match = re.search(r'\{.*\}', text, re.DOTALL)
        if match:
            try: return json.loads(match.group())
            except: return None
    return None

def is_missing(val):
    return not val or str(val).strip() in ["미검출", "없음", "확인불가", "N/A", "확인 불가", ""]

def parse_amount(text):
    if is_missing(text): return 0
    text = re.sub(r'[^\d]', '', str(text))
    return int(text) if text else 0

def format_phone(p_str):
    p = re.sub(r'[^\d]', '', str(p_str))
    if len(p) == 11: return f"{p[:3]}-{p[3:7]}-{p[7:]}"
    elif len(p) == 10:
        if p.startswith('02'): return f"{p[:2]}-{p[2:6]}-{p[6:]}"
        return f"{p[:3]}-{p[3:6]}-{p[6:]}"
    return p_str

# 💡 [코드 단축] 웹에서는 클립보드 복사 모듈(pyperclip) 작동이 불가능하므로 불필요한 복사 버튼 로직을 제거함
def edit_box(lbl, key, val, width_ratio=[4, 1]):
    st.markdown(f"**{lbl}**")
    e_val = st.text_input(lbl, value=val, placeholder="직접 입력", label_visibility="collapsed", key=f"in_{key}")
    return e_val

def display_status_value(label, val):
    st.write(f"**{label}**")
    if not val or val == "미검출": st.info("🚨 입력 필요")
    elif "확인 불가" in str(val): st.warning("⚠️ 자동 추출 실패 (직접 입력)")
    else: st.code(val, language=None)

def clean_name(s): 
    s = str(s)
    s = re.sub(r'\(.*?\)', '', s) 
    s = re.sub(r'[^\w가-힣]', '', s).lower() 
    for kw in ['주식회사', '유한회사', '주', '유', '등']:
        s = s.replace(kw, '')
    return s.strip()

@st.cache_data(ttl=60) 
def load_user_db():
    db = {}
    
    if not os.path.exists("user_list.xlsx"):
        log("user_list.xlsx 파일이 없어 빈 DB로 시작합니다.")
        return db
        
    try:
        df = pd.read_excel("user_list.xlsx").fillna("")
        for _, row in df.iterrows():
            name = str(row.get('성명 및 직위', '')).strip()
            dept = str(row.get('소속', '')).strip()
            if name:
                key = f"{name} ({dept})" if dept else name
                db[key] = {"phone": str(row.get('연락처', '')).strip(), "email": str(row.get('이메일', '')).strip()}
    except Exception as e: 
        log(f"DB 로드 실패: {e}")
    return db

def handle_error(err_type, msg, e=None):
    log_msg = f"[{err_type}] {msg}"
    if e: log_msg += f" (상세: {str(e)})"
    log(log_msg)
    
    if err_type == "ERP":
        st.error(f"🚨 **[ERP 양식 오류]** {msg}")
    elif err_type == "AI":
        st.warning(f"⚠️ **[AI 분석 지연]** {msg}")
    elif err_type == "AMOUNT":
        st.error(f"💰 **[금액 이상 감지]** {msg}")
    elif err_type == "FILE":
        st.error(f"📁 **[파일/폴더 오류]** {msg}")
    else:
        st.error(f"❗ **[시스템 오류]** {msg}")

USER_DB = load_user_db()
CACHE_FILE = "local_cache.json"

def save_to_cache(req_no, meta_data, raw_data):
    cache = {}
    if os.path.exists(CACHE_FILE):
        try:
            with open(CACHE_FILE, "r", encoding="utf-8") as f: cache = json.load(f)
        except: pass
    
    existing_record = cache.get(req_no, {"meta": {}, "raw": {}})
    existing_record["meta"].update(meta_data)
    existing_record["raw"].update(raw_data)
    cache[req_no] = existing_record
    
    with open(CACHE_FILE, "w", encoding="utf-8") as f:
        json.dump(cache, f, ensure_ascii=False, indent=2)

def get_cache_list():
    if not os.path.exists(CACHE_FILE): return []
    try:
        with open(CACHE_FILE, "r", encoding="utf-8") as f: cache = json.load(f)
        result = []
        for idx, (k, v) in enumerate(cache.items(), 1):
            meta = v.get("meta", {})
            result.append({
                "구매요구번호": meta.get("구매요구번호", k),
                "과제번호": meta.get("과제번호", ""),
                "구매요청액": meta.get("구매요청액", ""),
                "연구책임자정보": meta.get("연구책임자정보", ""),
                "물품담당자": meta.get("물품담당자", ""),
                "업체명": meta.get("업체명", ""),
                "업체담당자정보": meta.get("업체담당자정보", "")
            })
        return result[::-1]
    except: return []

def load_from_cache(req_no):
    try:
        with open(CACHE_FILE, "r", encoding="utf-8") as f: 
            return json.load(f).get(req_no, {}).get("raw", None)
    except: return None

# ==========================================
# 3. 데이터 파싱 엔진
# ==========================================
def extract_base_info(text):
    info = {
        'p_no': "⚠️ 확인 불가", 'req_no': "⚠️ 확인 불가", 'pi_name': "", 
        'pi_dept': "", 'pi_formatted': "", 'prj_end': "", 'pi_email': ""
    }
    
    p_no_match = re.search(r'(\d{4}-\d{4}-\d{2})', text)
    if p_no_match: info['p_no'] = p_no_match.group(1)
    
    req_no_match = re.search(r'구매요구번호[^\d]*(\d{12})', text)
    if req_no_match: info['req_no'] = req_no_match.group(1)
    
    prj_match = re.search(r'연차연구기간\s*(\d{4}-\d{2}-\d{2})\s*~\s*(\d{4}-\d{2}-\d{2})', text)
    if prj_match: info['prj_end'] = prj_match.group(2)
    
    pi_match = re.search(r'연구책임자정보\s*(.*?)\s*연차연구기간', text, re.DOTALL)
    if pi_match:
        parts = [p.strip() for p in pi_match.group(1).strip().split('/')]
        if len(parts) >= 4:
            info['pi_name'] = parts[0]
            info['pi_dept'] = parts[2] if len(parts) >= 5 else parts[1]
            info['pi_email'] = parts[-1] 
            info['pi_formatted'] = f"{parts[0]} / {format_phone(parts[-2])} / {parts[-1]}"
            
            clean_text = text.replace(" ", "")
            pi_name = info['pi_name']
            prj_context_match = re.search(r'과제성격(.*?)(연구책임자정보|연차연구기간)', clean_text, flags=re.DOTALL)
            prj_context = prj_context_match.group(1) if prj_context_match else ""
            
            if pi_name == "박재우" and ("지역혁신중심" in prj_context or "RISE" in prj_context.upper()):
                info['pi_dept'] += "(RISE사업단)"
            elif pi_name == "김재호" and "사물인터넷" in prj_context:
                info['pi_dept'] += "(사물인터넷혁신융합대학사업단)"
            elif pi_name == "송오영" and "SW중심대학" in prj_context:
                info['pi_dept'] += "(SW중심대학사업단)"
        else:
            info['pi_formatted'] = pi_match.group(1).strip()
            
    return info

def get_accurate_total(text, keyword=""):
    search_area = text.split(keyword)[-1] if keyword else text
    search_area = re.sub(r'\b\d+(?:-\d+)+\b', '', search_area)
    search_area = re.sub(r'\b\d[\d,]*\s*(mm|cm|m|kg|g|mg|ml|l|v|hz|w|ea|u|%)\b', '', search_area, flags=re.IGNORECASE)
    search_area = re.sub(r'\b[a-zA-Z]+\d+|\d+[a-zA-Z]+\b', '', search_area)
    all_numbers = re.findall(r'\b\d[\d,]*\b', search_area)
    
    clean_numbers = []
    for n in all_numbers:
        num_str = n.replace(',', '')
        val = safe_int(num_str)
        if num_str.startswith('0') and len(num_str) > 1: continue
        if len(num_str) >= 11: continue 
        if len(num_str) == 8 and (num_str.startswith('19') or num_str.startswith('20')):
            month = safe_int(num_str[4:6])
            day = safe_int(num_str[6:8])
            if 1 <= month <= 12 and 1 <= day <= 31: continue
        if val >= 1000 or val == 0: clean_numbers.append(val)
    return max(clean_numbers) if clean_numbers else 0

def guess_category_by_keyword(name, price):
    clean_name = name.replace(" ", "").lower()
    consumable_kws = ['시약', '용액', '항체', '키트', 'kit', '튜브', '팁', '케이블', '부품', '커버', '케이스', '필터', '토너', '잉크', '용지', '리필', '건전지', '배터리', '소모품']
    if any(kw in clean_name for kw in consumable_kws): return "소모품"
        
    fixture_kws = ['책상', '의자', '파티션', '캐비닛', '수납장', '테이블', '책장', '칠판', '소파', '침대', '블라인드']
    if any(kw in clean_name for kw in fixture_kws): return "집기"
        
    machine_kws = ['서버', 'pc', '데스크탑', '노트북', '모니터', '분석기', '현미경', '측정기', '펌프', '레이저', '냉장고', '카메라', '프린터', '라우터', '스위치', '오실로스코프']
    if any(kw in clean_name for kw in machine_kws): return "기계기구"
        
    if price >= 1000000: return "비품(분류확인요망)"
    return "분류확인요망"

def is_tax_exempt_project(raw_req):
    if not raw_req: return False
    clean_req = raw_req.replace(' ', '')
    if re.search(r'(과제성격|회계단위).*?면세', clean_req) or "면세" in clean_req.split("과제정보")[0]:
        return True
    return False

def is_vat_zero_in_finance(raw_req, raw_fin):
    combined_text = f"{raw_req}\n{raw_fin}"
    for line in combined_text.split('\n'):
        tokens = [t.strip() for t in line.strip().split('\t') if t.strip()]
        if len(tokens) == 9 and all(t.replace(',', '').lstrip('-').isdigit() for t in tokens):
            nums = [int(t.replace(',', '')) for t in tokens]
            return nums[1] == 0 
            
    if raw_fin:
        lines = raw_fin.strip().split('\n')
        total_vat = 0
        valid_row_count = 0
        for line in lines:
            if '합계' in line or '공급가' in line or '부가세' in line: continue
            tokens = [t.strip().replace(',', '') for t in line.split('\t')]
            nums = [int(t) for t in tokens if t.lstrip('-').isdigit()]
            if len(nums) >= 3:
                total_vat += nums[-1]
                valid_row_count += 1
        if valid_row_count > 0:
            return total_vat == 0
    return False

def parse_contact(token_str):
    if not token_str or not token_str.strip() or token_str in ["미검출", "없음", "확인불가", "N/A", "확인 불가"]:
        return "⚠️ 연락처 없음"
    raw = token_str.strip()
    
    email_match = re.search(r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}', raw)
    email = email_match.group() if email_match else None
    
    pure_digits = re.sub(r'[^\d]', '', raw)
    phone = None
    
    internal_match = re.search(r'(3408|6935)\d{4}', pure_digits)
    
    if internal_match:
        raw_phone = internal_match.group()
        phone = f"02-{raw_phone[:4]}-{raw_phone[4:]}"
    elif len(pure_digits) in [9, 10, 11]:  
        phone = format_phone(pure_digits)
    elif len(pure_digits) == 4:  
        if pure_digits.startswith(('34', '69')):
            phone = f"02-3408-{pure_digits}" 
        else:
            phone = f"내선 {pure_digits}"

    clean_name = raw
    if email: clean_name = clean_name.replace(email, '')
    
    phone_digits_match = re.search(r'[\d\-\s]{4,}', clean_name)
    if phone_digits_match: clean_name = clean_name.replace(phone_digits_match.group(), '')
        
    clean_name = re.sub(r'[\(\)\[\]\-\:\/\_\=\+]', '', clean_name).strip()

    if clean_name and phone: return f"{clean_name} / {phone}"
    elif clean_name and email: return f"{clean_name} / {email}"
    elif email: return f"담당자 / {email}"
    elif phone: return f"담당자 / {phone}"
    elif clean_name and len(clean_name) >= 2: return f"{clean_name} / ⚠️ 연락처 미기재"
        
    final_fallback = raw.strip()
    return final_fallback if final_fallback else "⚠️ 연락처 없음"

def parse_goods_erp(text):
    items = []
    valid_format = False
    if not text.strip(): return items, valid_format
    for line in text.strip().split('\n'):
        if not line.strip(): continue
        tokens = line.split('\t')
        try:
            anchor_idx = next((i for i, t in enumerate(tokens) if t.strip() in ['예', '아니오']), -1)
            if anchor_idx != -1:
                valid_format = True
                reg_status = tokens[anchor_idx].strip()
                cat_idx, con_idx = anchor_idx - 5, anchor_idx + 1
                
                contact_raw = tokens[con_idx].strip() if con_idx < len(tokens) else ""
                price = safe_int(tokens[anchor_idx - 6]) if anchor_idx >= 6 else 0
                qty = safe_int(tokens[anchor_idx - 11]) if anchor_idx >= 11 else 1
                unit = tokens[anchor_idx - 10].strip() if anchor_idx >= 10 else ""
                unit_price = safe_int(tokens[anchor_idx - 9]) if anchor_idx >= 9 else 0
                
                if unit_price > 0 and unit_price <= 150000 and reg_status == "예": reg_status = "예/확인필요"
                elif unit_price > 150000 and reg_status == "아니오": reg_status = "아니오/확인필요"
                
                raw_cat = tokens[cat_idx].strip() if cat_idx >= 0 and len(tokens) > cat_idx else ""
                name_str = tokens[1].strip() if len(tokens) > 1 else "무명"
                if len(tokens) > 2 and tokens[2].strip(): name_str = f"{name_str} ({tokens[2].strip()})"
                
                items.append({
                    'name': name_str, 'category': raw_cat if raw_cat else "공란", 
                    'reg_status': reg_status, 'contact': parse_contact(contact_raw), 
                    'qty': qty, 'unit': unit, 'price': price, 'unit_price': unit_price
                })
        except Exception as e: 
            log(f"물품 파싱 에러: {e}")
            continue 
    return items, valid_format

def parse_service_erp(text, force_category):
    items = []
    valid_format = False
    if not text.strip(): return items, valid_format
    for line in text.strip().split('\n'):
        if not line.strip(): continue
        tokens = line.split('\t')
        try:
            if len(tokens) >= 15:
                valid_format = True
                name_str = tokens[1].strip() if len(tokens) > 1 else "무명"
                if len(tokens) > 2 and tokens[2].strip(): name_str = f"{name_str} ({tokens[2].strip()})"
                
                items.append({
                    'name': name_str, 
                    'category': tokens[9].strip() if tokens[9].strip() else force_category, 
                    'reg_status': "N/A(용역)", 
                    'contact': parse_contact(tokens[14].strip()), 
                    'qty': safe_int(tokens[3]), 'unit': tokens[4].strip(), 
                    'price': safe_int(tokens[8]), 'unit_price': safe_int(tokens[5])
                })
        except Exception as e:
            log(f"용역/공사 파싱 에러: {e}")
            continue
    return items, valid_format

def generate_opinion(req_type, contract_type, amount, item_list, raw_finance_text=""):
    rules_triggered = []
    if "간접비" in raw_finance_text: rules_triggered.append("간접비건 확인")
    
    is_goods = "물품" in req_type
    is_service = "용역" in req_type or "공사" in req_type
    
    if is_goods and any(i['unit'].lower() not in ['', 'u'] for i in item_list): rules_triggered.append("단위 'u'로 수정 필요")
    if is_service and any(i['unit'] not in ['', '식'] for i in item_list): rules_triggered.append("단위 '식'으로 수정 필요")

    expected_cats = []
    for itm in item_list:
        erp_cat = itm.get('category', '공란').strip()
        reg_status = itm.get('reg_status', '')
        price = itm.get('price', 0)
        unit_price = itm.get('unit_price', 0)
        name = itm.get('name', '').upper()

        is_sw = any(kw in name for kw in ['소프트웨어', 'SW', 'S/W', '라이선스', '구독']) or '소프트웨어' in erp_cat

        if is_service:
            exp_cat = "공사, 용역" if "공사" in req_type else "용역"
        elif unit_price > 0 and unit_price <= 150000:
            exp_cat = "소모품"
            if "예" in reg_status: rules_triggered.append("물품등록여부 변경필요")
        elif is_sw and (price <= 2000000 or any(kw in name for kw in ['1년', '12개월', '1YR', '1 YR'])):
            exp_cat = "소모품"
        elif "예" in reg_status:
            exp_cat = "비품"
        elif erp_cat in ['', '공란'] or erp_cat == '비품':
            exp_cat = guess_category_by_keyword(name, price)
        else:
            exp_cat = erp_cat if erp_cat else "소모품"

        expected_cats.append(exp_cat)

        if erp_cat in ['', '공란']:
            rules_triggered.append("물품분류구분 입력필요")
        else:
            is_erp_bipum = "비품" in erp_cat or "기계기구" in erp_cat or "집기" in erp_cat
            is_exp_bipum = "비품" in exp_cat or "기계기구" in exp_cat or "집기" in exp_cat
            
            if (is_erp_bipum and exp_cat == "소모품") or ("소모품" in erp_cat and is_exp_bipum):
                rules_triggered.append("물품분류구분 변경필요")
            elif exp_cat == "분류확인요망" or "확인요망" in exp_cat:
                rules_triggered.append("물품분류구분 확인필요")

    clean_cats = [c.replace("(분류확인요망)", "").replace("분류확인요망", "확인필요") for c in expected_cats]
    final_category = " / ".join(list(dict.fromkeys(clean_cats))) if clean_cats else "소모품"

    unique_rules = list(dict.fromkeys(rules_triggered))
    issues_text = " - 특이사항 없음" if not unique_rules else "\n".join([f" - {r}" for r in unique_rules])
    
    return f"1. {contract_type}에 의한 구매건\n - 물품분류구분 : {final_category}\n\n2. 확인사항\n{issues_text}"

# ==========================================
# 4. AI 서류 검증 엔진
# ==========================================
def run_ai_analysis(files, req_type):
    today_str = datetime.datetime.today().strftime('%Y-%m-%d')
    
    extra_instruction = ""
    if req_type == "🏗️ 공사":
        extra_instruction = "🚨 이 건은 '공사' 건입니다. 첨부 서류 중 '시설공사 확인서' 또는 '장비설치 확인서'가 반드시 포함되어야 하므로 이를 꼼꼼히 확인하세요.\n"
    
    contents = [
        f"🚨 [기준일] 오늘 날짜는 {today_str}입니다. 만료일자 계산 시 참고하세요.\n"
        f"{extra_instruction}"
        "분석 후 반드시 아래 형태의 JSON 형식으로만 응답해. 절대 마크다운(```json)이나 다른 텍스트를 넣지 마.\n"
        "{\n"
        '  "비교견적리스트": [ {"업체명": "A업체", "사업자번호": "123-45-67890", "제출금액": 10000}, {"업체명": "B업체", "사업자번호": "987-65-43210", "제출금액": 12000} ],\n'
        '  "최저가판단과정": "위 비교견적리스트를 바탕으로 가장 금액이 낮은 업체를 최종 선정하는 과정을 짧게 작성",\n'
        '  "수신자명확인": "첨부된 모든 견적서(비교견적서 포함)의 수신자가 모두 \'세종대학교 산학협력단\' 명의가 맞으면 \'예\', 하나라도 누락되거나 다르면 \'아니오(사유)\'",\n'
        '  "만료일자": "유효기간 만료일자를 YYYY-MM-DD 형식으로 기재",\n'
        '  "동일규격확인": "복수 견적서 간 동일 물품이면 \'예\', 다르면 \'아니오(사유)\', 1장뿐이면 \'단일견적\'",\n'
        '  "사업자일치여부": "최저가(최종 선정) 견적서와 사업자등록증 상의 정보가 일치하면 \'예\', 다르면 \'아니오(사유)\'",\n'
        '  "업체명": "첨부된 사업자등록증 상의 \'법인명(단체명)\' 또는 \'상호\'를 최우선으로 추출 (사업자등록증이 없는 경우 견적서의 상호명 추출)",\n'
        '  "사업자번호": "최저가 업체의 사업자등록번호 (반드시 000-00-00000 형식)",\n'
        '  "대표자명": "사업자등록증 상의 대표자 성명",\n'
        '  "담당자": "최저가 업체의 견적 담당자명",\n'
        '  "연락처": "최저가 업체 담당자의 010 전화번호 (없으면 대표번호)",\n'
        '  "이메일": "최저가 업체의 이메일",\n'
        '  "예금주": "통장사본의 예금주명",\n'
        '  "계좌번호": "통장사본의 계좌번호",\n'
        '  "총금액": "최종 선정된 업체의 견적 총액 (반드시 부가가치세/VAT 포함된 최종 합계 금액을 숫자로만 기재)",\n'
        '  "공사확인서유무": "첨부 파일 중 \'시설공사\' 또는 \'장비설치 확인서\' 문서가 있으면 \'예\', 없으면 \'아니오\', 비슷한 문서가 있으면 \'확인요망\'",\n'
        '  "품목리스트": [ {"품명": "품명 텍스트", "수량": 1, "단가": 1000, "금액": 1000} ]\n'
        "}"
    ]
    
    for f in files:
        if "excel" in f.type or "spreadsheet" in f.type:
            try: contents.append(f"엑셀 문서 내용:\n{pd.read_excel(f).to_csv(index=False)}")
            except: pass
        else: contents.append(types.Part.from_bytes(data=f.getvalue(), mime_type=f.type))
            
    gen_config = types.GenerateContentConfig(response_mime_type="application/json")
            
    fallback_chain = [
        (AI_MODEL, 2),
        ("gemini-2.5-pro", 1),
        ("gemini-1.5-flash", 1)
    ]
    
    for model_name, max_retries in fallback_chain:
        for attempt in range(max_retries):
            try:
                response = ai_client.models.generate_content(model=model_name, contents=contents, config=gen_config)
                if response and response.text:
                    parsed = safe_json_load(response.text.strip())
                    if parsed: 
                        log(f"✅ AI 분석 성공: {model_name} (시도 {attempt+1})")
                        return parsed
            except Exception as e: 
                log(f"⚠️ AI 호출 실패 ({model_name}, 시도 {attempt+1}): {e}")
                time.sleep(2)

    log("🚨 모든 Fallback 모델 호출 실패. 수동 입력 모드로 전환합니다.")
    return None

def extract_contract_info(text):
    clean_text = text.replace('\n', ' ').replace('\r', '').replace('\t', ' ')
    info = {
        'req_no': re.search(r'구매요구번호\s*\*?\s*(\d+)', clean_text),
        'con_name': re.search(r'계약건명\s*(.*?)\s*납품기간', clean_text),
        'con_amount': re.search(r'계약금액\s*([\d,]+)', clean_text),
        'vendor_name': re.search(r'계약업체명\s*(.*?)\s*사업자번호', clean_text),
        'email': re.search(r'e-mail주소\s*([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})', clean_text)
    }
    result = {k: v.group(1).strip() if v else "" for k, v in info.items()}
    result['con_amount'] = str(safe_int(result.get('con_amount', 0)))
    
    loc_match = re.search(r'납품장소\s*(.*?)\s*계약금액', clean_text)
    result['location'] = ""
    if loc_match and '/' in loc_match.group(1):
        parts = [p.strip() for p in loc_match.group(1).split('/')]
        if len(parts) >= 3:
            room_match = re.search(r'(\d+)', parts[-1])
            if room_match: result['location'] = f"{parts[0]} {int(room_match.group(1))}호"
    return result

# ==========================================
# 5. 화면 UI 구성
# ==========================================
with st.sidebar:
    st.title("💰 구매추출기")
    menu = st.radio("메뉴", ["📋 구매요청 분석", "🤝 구매계약 분석", "🕒 분석내용 불러오기", "🗂️ 유틸리티"], label_visibility="collapsed", key="main_menu")
    st.divider()
    st.success("✅ 시스템 정상 가동 중")

# ----------------- [1. 구매요청 분석] -----------------
if menu == "📋 구매요청 분석":
    st.title("📋 구매요청 분석")
    
    def on_critical_condition_change():
        if SESSION_REQ_KEY in st.session_state:
            del st.session_state[SESSION_REQ_KEY]
        st.session_state['api_failed'] = False
        st.session_state['is_saved'] = False
    
    st.write("### 📥 [STEP 1] 정보 입력 및 서류 업로드")
    
    st.markdown("#### 📂 1. 구매/계약 기본 정보 선택")
    c_type1, c_type2 = st.columns(2)

    with c_type1:
        req_type = st.radio("구매 유형", ["📦 물품", "🛠️ 용역", "🏗️ 공사"], horizontal=True, key="saved_req_type", on_change=on_critical_condition_change)
    with c_type2:
        con_type = st.radio("계약 방식", ["비교견적", "수의계약"], horizontal=True, key="saved_con_type", on_change=on_critical_condition_change)

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
            st.markdown("#### 📎 3. AI 교차검증용 서류")
            st.caption("견적서 및 사업자등록증 (드래그 앤 드롭)")
            
            if "file_key" not in st.session_state: st.session_state["file_key"] = 0
            
            up_files = st.file_uploader(
                f"최대 {MAX_UPLOAD_FILES}개 / {MAX_FILE_SIZE_MB}MB", 
                accept_multiple_files=True, 
                label_visibility="collapsed", 
                key=f"file_input_{st.session_state['file_key']}"
            )

            st.info(
                "**🤖 AI 서류 검증 체크리스트**\n\n"
                "- ✅ **유효 기간:** 견적서 유효기간 7일 이상 잔여 여부\n"
                "- ✅ **명의 일치:** 세종대학교 산학협력단 및 예금주/대표자 교차 검증\n"
                "- ✅ **공사 확인:** 시설공사/장비설치 확인서 첨부 여부 (해당 시)\n"
            )
        
    col_btn1, col_btn2 = st.columns(2)
    with col_btn1: start_all = st.button("🚀 전체 데이터 통합 분석 시작", width="stretch", type="primary")
        
    def reset_all_state():
        st.session_state["p_req"] = ""  
        st.session_state["p_fin"] = ""
        st.session_state["p_itm"] = ""
        st.session_state["file_key"] = st.session_state.get("file_key", 0) + 1
        st.session_state["is_saved"] = False
        if 'show_reset_warning' in st.session_state: del st.session_state['show_reset_warning']
        for k in [SESSION_REQ_KEY, SESSION_AI_KEY]:
            if k in st.session_state: del st.session_state[k]

    @st.dialog("🚨 데이터 초기화 경고")
    def reset_warning_popup():
        st.error("아직 분석 내역을 시스템(DB)에 저장하지 않았습니다!\n지금 초기화하면 작업 중인 데이터가 **모두 날아갑니다.**")
        st.info("정말 초기화하시겠습니까?")
        c_warn1, c_warn2 = st.columns(2)
        
        with c_warn1:
            if st.button("그래도 초기화 진행 🗑️", width="stretch"):
                reset_all_state()
                st.rerun() 
                
        with c_warn2:
            if st.button("취소 (저장하러 가기) 🔙", width="stretch", type="primary"):
                st.rerun() 

    with col_btn2: 
        if st.button("🔄 전체 초기화", width="stretch"):
            if st.session_state.get(SESSION_REQ_KEY) and not st.session_state.get('is_saved', False):
                reset_warning_popup()
            else:
                reset_all_state()
                st.rerun()

    if start_all:
        if not raw_req:
            st.warning("📌 구매요청서 텍스트를 먼저 입력해주세요.")
            st.stop()

        st.session_state[SESSION_REQ_KEY] = True
        st.session_state['is_saved'] = False
            
        if up_files:
            with st.spinner(f"🤖 AI가 '{req_type}' 기준에 맞춰 서류를 검증 중입니다..."):
                rt = run_ai_analysis(up_files, req_type)
                if rt: 
                    st.session_state[SESSION_AI_KEY] = rt
                    st.session_state['api_failed'] = False
                else: 
                    handle_error("AI", "구글 AI 서버 응답 실패. 수동 입력 모드로 전환합니다.")
                    st.session_state['api_failed'] = True
        else:
            st.session_state[SESSION_AI_KEY] = {}
            st.session_state['api_failed'] = False

    if st.session_state.get(SESSION_REQ_KEY):
        st.divider()
        st.write("### 📤 [STEP 2] 분석 결과")

        if st.session_state.get('api_failed', False):
            st.error("🚨 **AI 자동 검증 실패 (API 서버 과부하 또는 토큰 초과)**\n\n시스템이 멈추지 않도록 우회(수동) 모드를 가동합니다. 아래 가이드에 따라 제미나이 웹에서 직접 데이터를 추출하여 붙여넣으세요.")
            
            today_str = datetime.datetime.today().strftime('%Y-%m-%d')
            extra_instruction = "🚨 이 건은 '공사' 건입니다. 첨부 서류 중 '시설공사 확인서' 또는 '장비설치 확인서'가 반드시 포함되어야 하므로 이를 꼼꼼히 확인하세요.\n" if req_type == "🏗️ 공사" else ""
            
            fallback_prompt = (
                f"🚨 [기준일] 오늘 날짜는 {today_str}입니다. 만료일자 계산 시 참고하세요.\n"
                f"{extra_instruction}"
                "분석 후 반드시 아래 형태의 JSON 형식으로만 응답해. 절대 마크다운(```json)이나 다른 텍스트를 넣지 마.\n"
                "{\n"
                '  "비교견적리스트": [ {"업체명": "A업체", "사업자번호": "123-45-67890", "제출금액": 10000}, {"업체명": "B업체", "사업자번호": "987-65-43210", "제출금액": 12000} ],\n'
                '  "최저가판단과정": "위 비교견적리스트를 바탕으로 가장 금액이 낮은 업체를 최종 선정하는 과정을 짧게 작성",\n'
                '  "수신자명확인": "첨부된 모든 견적서(비교견적서 포함)의 수신자가 모두 \'세종대학교 산학협력단\' 명의가 맞으면 \'예\', 하나라도 누락되거나 다르면 \'아니오(사유)\'",\n'
                '  "만료일자": "유효기간 만료일자를 YYYY-MM-DD 형식으로 기재",\n'
                '  "동일규격확인": "복수 견적서 간 동일 물품이면 \'예\', 다르면 \'아니오(사유)\', 1장뿐이면 \'단일견적\'",\n'
                '  "사업자일치여부": "최저가(최종 선정) 견적서와 사업자등록증 상의 정보가 일치하면 \'예\', 다르면 \'아니오(사유)\'",\n'
                '  "업체명": "첨부된 사업자등록증 상의 \'법인명(단체명)\' 또는 \'상호\'를 최우선으로 추출 (사업자등록증이 없는 경우 견적서의 상호명 추출)",\n'
                '  "사업자번호": "최저가 업체의 사업자등록번호 (반드시 000-00-00000 형식)",\n'
                '  "대표자명": "사업자등록증 상의 대표자 성명",\n'
                '  "담당자": "최저가 업체의 견적 담당자명",\n'
                '  "연락처": "최저가 업체 담당자의 010 전화번호 (없으면 대표번호)",\n'
                '  "이메일": "최저가 업체의 이메일",\n'
                '  "예금주": "통장사본의 예금주명",\n'
                '  "계좌번호": "통장사본의 계좌번호",\n'
                '  "총금액": "최종 선정된 업체의 견적 총액 (반드시 부가가치세/VAT 포함된 최종 합계 금액을 숫자로만 기재)",\n'
                '  "공사확인서유무": "첨부 파일 중 \'시설공사\' 또는 \'장비설치 확인서\' 문서가 있으면 \'예\', 없으면 \'아니오\', 비슷한 문서가 있으면 \'확인요망\'",\n'
                '  "품목리스트": [ {"품명": "품명 텍스트", "수량": 1, "단가": 1000, "금액": 1000} ]\n'
                "}"
            )

            st.info("💡 **[수동 분석 가이드]**\n1. 제미나이(Gemini) 공식 웹페이지에 견적서 등 첨부파일을 업로드하세요.\n2. 아래 📋 아이콘을 눌러 프롬프트를 복사한 뒤 제미나이에 붙여넣고 전송합니다.\n3. 결과물로 나온 JSON 데이터( `{` 부터 `}` 까지 )를 복사해 아래 입력칸에 붙여넣습니다.")
            
            st.code(fallback_prompt, language="text")

            with st.container(border=True):
                st.markdown("#### ✍️ 제미나이 분석 결과(JSON) 붙여넣기")
                manual_json = st.text_area("결과값", height=150, label_visibility="collapsed", placeholder="여기에 결과값 { JSON } 형태를 그대로 붙여넣으세요.")
                
                c_btn1, c_btn2 = st.columns(2)
                with c_btn1:
                    if st.button("🚀 수동 데이터 시스템 적용", type="primary", width="stretch"):
                        parsed_manual = safe_json_load(manual_json)
                        if parsed_manual:
                            st.session_state[SESSION_AI_KEY] = parsed_manual
                            st.session_state['api_failed'] = False
                            st.success("✅ 수동 데이터가 완벽하게 적용되었습니다!")
                            time.sleep(1)
                            st.rerun() 
                        else:
                            st.error("🚨 형식 오류: 중괄호 { } 가 포함된 JSON 결과값을 정확히 붙여넣어주세요.")
                with c_btn2:
                    if st.button("⏭️ AI 검증 포기하고 그냥 넘어가기", width="stretch"):
                        st.session_state[SESSION_AI_KEY] = {}
                        st.session_state['api_failed'] = False
                        st.rerun()

            st.stop() 
        
        base_data = extract_base_info(raw_req)
        total_amt = get_accurate_total(raw_req, "합계액")
        
        if req_type == "📦 물품":
            item_list, is_valid = parse_goods_erp(raw_itm)
        elif req_type == "🛠️ 용역":
            item_list, is_valid = parse_service_erp(raw_itm, force_category="용역")
        elif req_type == "🏗️ 공사":
            item_list, is_valid = parse_service_erp(raw_itm, force_category="공사, 용역")
            
        if not is_valid and raw_itm.strip():
            handle_error("ERP", "물품/용역내역 행 텍스트가 정상적인 ERP 양식이 아닙니다. 표 전체를 정확히 드래그했는지 확인해주세요.")
            st.stop()
        
        st.success(f"✅ 텍스트 분석 완료 | 금액: {total_amt:,}원 | 품목: {len(item_list)}건 | 분류: {req_type}")

        prj_end_str = base_data.get('prj_end', '')
        if prj_end_str:
            try:
                prj_end_date = datetime.datetime.strptime(prj_end_str, '%Y-%m-%d').date()
                today = datetime.date.today()
                days_left = (prj_end_date - today).days
                
                if 0 <= days_left <= 30:
                    st.error(f"⏰ **[과제 종료 임박]** 과제 연차 종료일({prj_end_str})이 **{days_left}일** 남았습니다. 기한 내 대금 청구가 완료되도록 주의하세요!")
                elif days_left < 0:
                    st.error(f"🚨 **[과제 종료됨]** 과제 연차 종료일({prj_end_str})이 이미 **{abs(days_left)}일 지났습니다!** 과제 이월 또는 취소 여부를 즉시 확인하세요.")
            except Exception as e:
                pass
        
        if total_amt > 0 and total_amt < 1000:
            handle_error("AMOUNT", f"총금액이 {total_amt:,}원입니다. 원본 입력 시 '0'이 누락되었는지 확인하세요.")

        is_vat_zero = is_vat_zero_in_finance(raw_req, raw_fin)
        if is_vat_zero:
            handle_error("AMOUNT", "재원내역에 부가세가 0원으로 설정되어 있습니다. (면세 등 특수건 여부 확인 요망)")

        ai = st.session_state.get(SESSION_AI_KEY, {})
        if isinstance(ai, list): ai = ai[0] if len(ai) > 0 else {} 

        if up_files or ai:
            def format_bizno(bizno):
                clean_b = re.sub(r'[^\d]', '', str(bizno))
                return f"{clean_b[:3]}-{clean_b[3:5]}-{clean_b[5:]}" if len(clean_b) == 10 else bizno
                
            auto_reject_reasons = []
            nts_status_res = "⚠️ 사업자번호 미검출 (조회 불가)"

            if ai:
                biz_check_list = []
                
                main_biz = ai.get('사업자번호', '')
                if not is_missing(main_biz):
                    ai['사업자번호'] = format_bizno(main_biz)
                    biz_check_list.append({"name": ai.get('업체명', '최종선정업체'), "no": ai['사업자번호']})
                else:
                    st.warning("⚠️ 최저가 업체의 사업자번호가 감지되지 않았습니다.")

                for q in ai.get('비교견적리스트', []):
                    q_biz = q.get('사업자번호', '')
                    if not is_missing(q_biz):
                        q_formatted = format_bizno(q_biz)
                        if not any(b['no'] == q_formatted for b in biz_check_list):
                            biz_check_list.append({"name": q.get('업체명', '비교업체'), "no": q_formatted})

                nts_reports = []
                if biz_check_list:
                    with st.spinner(f"🔍 국세청에 사업자 휴·폐업 상태를 실시간 다중 조회 중입니다... (총 {len(biz_check_list)}건)"):
                        for biz in biz_check_list:
                            status = check_nts_business(biz['no'])
                            nts_reports.append(f"{biz['name']}: {status}")
                            
                            if "휴업자" in status:
                                auto_reject_reasons.append(f"국세청 조회 결과: '휴업' 상태인 사업자 포함 - {biz['name']} ({biz['no']})")
                            elif "폐업자" in status:
                                auto_reject_reasons.append(f"국세청 조회 결과: '폐업' 상태인 사업자 포함 - {biz['name']} ({biz['no']})")
                
                for q in ai.get('비교견적리스트', []):
                    if is_missing(q.get('사업자번호', '')):
                        nts_reports.append(f"{q.get('업체명', '비교업체')}: ⚠️ 번호 미기재(조회생략)")
                                
                nts_status_res = " / ".join(nts_reports) if nts_reports else "⚠️ 사업자번호 미검출 (전체 조회 불가)"
                
                exp_date_str = str(ai.get('만료일자', ''))
                v_exp_result = "⚠️ 유효기간 확인불가 (직접 확인 필요)"
                
                if not is_missing(exp_date_str):
                    match = re.search(r'\d{4}-\d{2}-\d{2}', exp_date_str)
                    if match:
                        try:
                            exp_date = datetime.datetime.strptime(match.group(), '%Y-%m-%d').date()
                            today = datetime.date.today()
                            days_left = (exp_date - today).days
                            
                            if days_left >= 7: v_exp_result = f"✅ 통과 (만료일: {match.group()}, {days_left}일 남음)"
                            else:
                                v_exp_result = f"🚨 아니오 (만료일: {match.group()}, {days_left}일 남음 - 7일 미만)"
                                auto_reject_reasons.append("견적서 유효기간 7일 미만")
                        except: pass

                biz_match = str(ai.get('사업자일치여부', '미검출'))
                if "아니오" in biz_match: auto_reject_reasons.append("견적서와 사업자등록증 상의 업체 정보 불일치")

                v_name = ai.get('업체명', '')
                r_name = ai.get('대표자명', '')
                b_holder = ai.get('예금주', '')

                cv, cr, cb = clean_name(v_name), clean_name(r_name), clean_name(b_holder)
                is_corp = any(kw in str(v_name) for kw in ["주식회사", "(주)", "유한회사", "(유)", "Inc", "Ltd"])
                
                if cb and cb != "미검출":
                    sim_v = difflib.SequenceMatcher(None, cv, cb).ratio()
                    sim_r = difflib.SequenceMatcher(None, cr, cb).ratio()
                    
                    if is_corp:
                        if cb in cv or cv in cb or sim_v >= 0.7: acc_match_res = "✅ 일치 (법인명 예금주 정상)"
                        else:
                            acc_match_res = f"🚨 불일치 의심 (예금주: {b_holder} / 법인명: {v_name})"
                            auto_reject_reasons.append(f"통장 예금주({b_holder})와 법인명({v_name}) 불일치")
                    else:
                        if cb in cv or cv in cb or cb in cr or cr in cb or sim_v >= 0.7 or sim_r >= 0.7:
                            acc_match_res = "✅ 일치 (개인사업자 예금주/대표자 정상)"
                        else:
                            acc_match_res = f"🚨 불일치 의심 (예금주: {b_holder} / 업체명: {v_name} / 대표자: {r_name})"
                            auto_reject_reasons.append(f"통장 예금주({b_holder})와 사업자 명의 불일치")
                else:
                    acc_match_res = "⚠️ 통장사본 미검출"

                is_const_doc = (req_type == "🏗️ 공사")
                has_const_doc_file = any("시설공사" in f.name.replace(" ", "") or "장비설치" in f.name.replace(" ", "") for f in up_files) if up_files else False
                const_res = ""
                
                if is_const_doc:
                    if ai.get('공사확인서유무', '') == '예' or has_const_doc_file:
                        const_res = "\n- **공사서류:** ✅ 시설공사/장비설치 확인서 첨부됨"
                    else:
                        const_res = "\n- **공사서류:** 🚨 **[서류 누락] '시설공사, 장비설치 확인서' 미첨부**"
                        auto_reject_reasons.append("시설공사 또는 장비설치 확인서 미첨부")

                if "아니오" in str(ai.get('수신자명확인', '')):
                    reason_detail = str(ai.get('수신자명확인', '')).replace('아니오', '').replace('(', '').replace(')', '').strip()
                    reject_msg = f"견적서 수신자 명의 누락/오기입 ({reason_detail})" if reason_detail else "견적서 수신자 명의(산학협력단) 누락 또는 오기입"
                    auto_reject_reasons.append(reject_msg)

                quotes_list = ai.get('비교견적리스트', [])
                if quotes_list:
                    q_texts = [f"{q.get('업체명', '미상')}({safe_int(q.get('제출금액', 0)):,}원)" for q in quotes_list]
                    quotes_res = f"\n- **단가비교:** ⚖️ {', '.join(q_texts)}"
                else:
                    quotes_res = "\n- **단가비교:** ⚠️ 단일견적 또는 비교견적서 미검출"

                report_text = (
                    f"- **산단명의:** {ai.get('수신자명확인', '미검출')}\n"
                    f"- **국세청조회:** {nts_status_res}\n"  
                    f"- **유효기간:** {v_exp_result}\n"
                    f"- **동일규격:** {ai.get('동일규격확인', '미검출')}"
                    f"{quotes_res}\n"
                    f"- **사업자일치:** {biz_match}\n"
                    f"- **예금주검증:** {acc_match_res}"
                    f"{const_res}"
                )
                
                if auto_reject_reasons:
                    st.warning(f"#### ⚠️ AI 서류 검증 경고\n{report_text}")
                    st.error("### 🚫 [반려 사유 자동 생성]")
                    reject_str = "\n".join([f"- {r}" for r in auto_reject_reasons])
                    
                    st.code(f"아래 사유로 반려합니다.\n\n{reject_str}", language="text")
                else:
                    st.info(f"#### 🤖 AI 서류 검증 통과\n{report_text}")

                ai_amt_raw = ai.get('총금액', '미검출')
                if is_missing(ai_amt_raw):
                    st.warning("⚠️ 견적서 금액 미검출 (AI 결과 없음)")
                else:
                    ai_amt = parse_amount(ai_amt_raw)
                    if ai_amt == 0: st.warning(f"⚠️ 금액 파싱 실패 (원본: {ai_amt_raw}) - 직접 확인하세요.")
                    elif ai_amt == total_amt: st.success(f"⚖️ **[금액 교차검증 일치]** ERP 금액과 견적서 총액이 일치합니다. ({total_amt:,}원)")
                    else:
                        ai_items = ai.get('품목리스트', [])
                        calc_total = 0
                        for itm in ai_items:
                            item_price = safe_int(itm.get('금액', 0))
                            if item_price == 0: item_price = safe_int(itm.get('단가', 0)) * safe_int(itm.get('수량', 1))
                            calc_total += item_price

                        vat_inc_ai = round(ai_amt * 1.1)
                        vat_inc_calc = round(calc_total * 1.1)
                        allow_vat_correction = ("부가세" in str(ai.values()).upper() or "VAT" in str(ai.values()).upper() or safe_int(ai.get('부가세', 0)) > 0)
                        margin = 10 
                        
                        if abs(calc_total - total_amt) <= margin and calc_total > 0:
                            st.success(f"⚖️ **[금액 검증]** 품목 합산액이 일치합니다. ({total_amt:,}원) ➔ `[ITEM_SUM_MATCH]`")
                        elif allow_vat_correction and abs(vat_inc_ai - total_amt) <= margin:
                            st.success(f"⚖️ **[금액 검증]** 총액+VAT(10%) 보정 시 일치합니다. ({total_amt:,}원) ➔ `[VAT_CORRECTED]`")
                        elif allow_vat_correction and abs(vat_inc_calc - total_amt) <= margin and calc_total > 0:
                            st.success(f"⚖️ **[금액 검증]** 합산액+VAT(10%) 보정 시 일치합니다. ({total_amt:,}원) ➔ `[VAT_CORRECTED_SUM]`")
                        else:
                            best_guess = calc_total if (calc_total > 0 and abs(calc_total - total_amt) < abs(ai_amt - total_amt)) else ai_amt
                            st.error(f"🚨 **[금액 불일치]** ERP: {total_amt:,}원 / 추출액: {ai_amt:,}원 / 차액: {best_guess - total_amt:+,}원")

                st.divider()
                st.write("#### 🔍 품목 교차 검증 (AI 유사도 분석)")
                ai_items = ai.get('품목리스트', [])
                
                if not ai_items:
                    st.warning("⚠️ 견적서 세부 품목 미검출 (총액 견적서 등) ➔ 유사도 검사를 건너뜁니다.")
                elif not item_list:
                    st.warning("⚠️ ERP 품목 내역이 없어 비교가 불가능합니다.")
                else:
                    c_itm1, c_itm2 = st.columns(2)
                    with c_itm1:
                        st.markdown(f"**[🏢 ERP 입력 품목: {len(item_list)}건]**")
                        for i, itm in enumerate(item_list):
                            pr_str = f" / {itm.get('price', 0):,}원" if itm.get('price', 0) > 0 else ""
                            st.text(f"{i+1}. {itm['name']} ({itm.get('qty', 1)} {itm.get('unit', '')}){pr_str}")
                    with c_itm2:
                        st.markdown(f"**[📄 견적서 추출 품목: {len(ai_items)}건]**")
                        for i, itm in enumerate(ai_items):
                            st.text(f"{i+1}. {itm.get('품명', '미검출')} | 수량: {itm.get('수량', '')} | 단가: {safe_int(itm.get('단가', 0)):,}")

                    if len(ai_items) == 1 and any(word in str(ai_items[0].get('품명', '')) for word in ['식', '세트', '외', '통합']):
                        st.info(f"⚠️ '1식(세트)' 통합견적입니다. (품명: {ai_items[0].get('품명', '')})")
                    
                    else:
                        st.markdown("**📊 품명 및 금액/수량 교차검증 결과 (스마트 매칭)**")
                        match_results = []
                        
                        for erp_itm in item_list:
                            erp_name = erp_itm['name']
                            erp_qty = erp_itm.get('qty', 1) 
                            erp_price = safe_int(erp_itm.get('price', 0))
                            
                            best_match, best_score, best_ai_qty = None, 0, 0
                            
                            for ai_itm in ai_items:
                                ai_name = ai_itm.get('품명', '')
                                ai_qty = safe_int(ai_itm.get('수량', 1)) or 1
                                ai_unit_price = safe_int(ai_itm.get('단가', 0))
                                ai_total_price = safe_int(ai_itm.get('금액', 0))
                                if ai_total_price == 0: ai_total_price = ai_unit_price * ai_qty
                                
                                text_score = difflib.SequenceMatcher(None, re.sub(r'[^\w가-힣]', '', str(erp_name)).lower(), re.sub(r'[^\w가-힣]', '', str(ai_name)).lower()).ratio() * 100
                                
                                price_score = 0
                                if erp_price > 0 and ai_total_price > 0:
                                    if erp_price == ai_total_price: 
                                        price_score = 500  
                                    elif erp_price == round(ai_total_price * 1.1): 
                                        price_score = 500  
                                    elif abs(erp_price - round(ai_total_price * 1.1)) <= 10:
                                        price_score = 400  
                                
                                total_score = text_score + price_score
                                
                                if total_score > best_score:
                                    best_score = total_score
                                    best_match = ai_name
                                    best_ai_qty = ai_qty
                                    
                            is_matched = best_score >= 400 or best_score >= 50
                            
                            name_status = "✅ 매칭 성공" if is_matched else "🚨 매칭 실패"
                            qty_status = "✅ 일치" if is_matched and erp_qty == best_ai_qty else (f"🚨 다름 (ERP: {erp_qty} / 견적: {best_ai_qty})" if is_matched else "➖ 비교 불가")
                            
                            match_results.append({
                                "ERP 입력 품명": erp_name, "견적서 매칭 품명": best_match if is_matched else "매칭 실패",
                                "매칭 기준": "💰 금액 매칭" if best_score >= 400 else ("📝 이름 매칭" if best_score >= 50 else "실패"), 
                                "수량 검증": qty_status, "최종 판정": name_status
                            })
                        st.dataframe(pd.DataFrame(match_results), width="stretch", hide_index=True)
        else:
            st.info("📄 첨부된 서류가 없어 AI 교차 검증(금액/품목 대조)은 생략하고 ERP 텍스트 분석 결과만 출력합니다.")

        st.divider()
        st.write("#### 📝 ERP 기본 정보")
        r1_c1, r1_c2, r1_c3 = st.columns(3)
        with r1_c1: display_status_value("🔢 구매요구번호", base_data['req_no'])
        with r1_c2: display_status_value("📌 과제번호", base_data['p_no'])
        with r1_c3: display_status_value("👤 연구책임자 성명", base_data['pi_name'])
        
        r2_c1, r2_c2, r2_c3 = st.columns(3)
        with r2_c1: display_status_value("🏫 연구책임자 학과", base_data['pi_dept'])
        with r2_c2: display_status_value("💰 구매요청액", f"{total_amt:,}원")
        with r2_c3: display_status_value("📧 연구책임자 정보", base_data['pi_formatted'])
        
        st.divider()
        st.write("#### 📦 ERP 품목 내역")
        if item_list:
            t_data = []
            for i, itm in enumerate(item_list):
                stt = itm['reg_status']
                s_disp = f"🚨 {stt}" if "확인필요" in stt else (f"✅ 예" if stt=="예" else ("➖ 해당없음" if "N/A" in stt else f"⚠️ 아니오"))

                t_data.append({"No.": i+1, "품명": itm['name'], "단위": itm['unit'], "등록여부": s_disp, "물품담당자": itm['contact']})
            st.dataframe(pd.DataFrame(t_data), width="stretch", hide_index=True)
        else:
            st.warning("품목 내역이 없습니다.")
        
        st.divider()
        st.write("#### ✍️ 결재 의견 (자동 생성)")
        st.code(generate_opinion(req_type, con_type, total_amt, item_list, raw_fin), language="text")
        
        st.divider()
        
        # 💡 [코드 단축] 데스크탑의 탐색기 기능이 불필요하여 깔끔하게 폴더 이름 텍스트 복사 뷰어만 남김
        st.write("#### 📁 업체 정보 통합 및 추천 폴더명")
        
        c_e1, c_e2 = st.columns(2)
        with c_e1: t_dept = st.text_input("학과명", value=base_data['pi_dept'])
        with c_e2: t_name = st.text_input("성명", value=base_data['pi_name'])
        
        ai_phone = ai.get('연락처', '') if ai else ''
        if ai_phone and ai_phone != "미검출":
            ai_ph_clean = re.sub(r'[^\d]', '', ai_phone)
            internal_phones = [re.sub(r'[^\d]', '', base_data.get('pi_formatted', ''))]
            for itm in item_list: internal_phones.append(re.sub(r'[^\d]', '', itm.get('contact', '')))
            for int_ph in [p for p in internal_phones if len(p) >= 9]:
                if int_ph in ai_ph_clean or ai_ph_clean in int_ph:
                    ai_phone = ""; break

        c_v1, c_v2, c_v3, c_v4, c_v5 = st.columns(5)
        with c_v1: t_ven = st.text_input("업체명", value=ai.get('업체명', '') if ai else '')
        with c_v2: t_biz = st.text_input("사업자번호", value=ai.get('사업자번호', '') if ai else '')
        with c_v3: t_rep = st.text_input("업체담당자명", value=ai.get('담당자', '') if ai else '')
        with c_v4: t_phn = st.text_input("연락처", value=ai_phone) 
        with c_v5: t_eml = st.text_input("이메일", value=ai.get('이메일', '') if ai else '')
        
        st.caption("✨ **원클릭 복사 보드**")
        c_cp1, c_cp2, c_cp3 = st.columns([1, 1, 2])
        with c_cp1: st.code(t_ven, language=None)
        with c_cp2: st.code(t_biz, language=None)
        with c_cp3: st.code(f"{t_rep} / {t_phn} / {t_eml}", language=None) 
        
        now = datetime.datetime.today()
        fold_nm = f"({now.strftime('%Y.%m.%d')}) {t_dept} {t_name} - {t_ven}"
        
        st.info("💡 오른쪽 아이콘을 눌러서 폴더 이름을 복사하고 바탕화면에 새 폴더를 만드세요.")
        st.code(fold_nm, language="text")

        st.divider()
        st.write("#### 🏦 통장사본 계좌정보")
        c_b1, c_b2 = st.columns(2)
        with c_b1:
            t_acc_holder = st.text_input("예금주 (수정 가능)", value=ai.get('예금주', '') if ai else '', key="req_acc_holder")
        with c_b2:
            t_acc_num = st.text_input("계좌번호 (수정 가능)", value=ai.get('계좌번호', '') if ai else '', key="req_acc_num")

        st.divider()
        if st.button("💾 현재 분석 내역을 시스템(DB)에 저장하기", type="primary", width="stretch"):
            raw_to_save = {
                "req_input": raw_req, "fin_input": raw_fin, "itm_input": raw_itm, "ai_res": ai, "req_type": req_type, "con_type": con_type
            }
            
            item_mgrs = []
            for itm in item_list:
                if itm.get('contact') and "없음" not in itm['contact']:
                    item_mgrs.append(itm['contact'])
            unique_item_mgrs = list(dict.fromkeys(item_mgrs))
            final_item_mgr = " / ".join(unique_item_mgrs) if unique_item_mgrs else "⚠️ 미기재"

            vendor_mgr_info = f"{t_rep} / {t_phn} / {t_eml}"
            
            meta_to_save = {
                "요청분석일자": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "구매요구번호": base_data['req_no'],
                "과제번호": base_data['p_no'],
                "구매요청액": f"{total_amt:,}원",
                "연구책임자정보": base_data['pi_formatted'], 
                "물품담당자": final_item_mgr,               
                "업체명": t_ven,
                "업체담당자정보": vendor_mgr_info,          
                "사업자번호": t_biz,
                "계좌번호": t_acc_num,
                "예금주": t_acc_holder,
                "연구책임자_성명": base_data['pi_name'],
                "연구책임자_학과": base_data['pi_dept'],
                "연구책임자_이메일": base_data['pi_email']
            }
            
            save_to_cache(base_data['req_no'], meta_to_save, raw_to_save)
            st.session_state['is_saved'] = True
            st.success(f"✅ 구매요구번호 [{base_data['req_no']}] 건의 핵심 데이터가 완벽하게 저장되었습니다! '분석내용 불러오기' 메뉴에서 확인 가능합니다.")

# ----------------- [2. 구매계약 분석 및 발주서 자동 생성] -----------------
elif menu == "🤝 구매계약 분석":
    st.title("🤝 구매계약 다중 분석 및 발주서 자동 생성")
    st.info("여러 개의 계약 건을 한 번에 입력하고 분석하여, 각각의 워드 발주서를 빠르게 생성할 수 있습니다.")
    
    if 'con_input_count' not in st.session_state:
        st.session_state['con_input_count'] = 2

    raw_cons = []
    for i in range(st.session_state['con_input_count']):
        txt = st.text_area(f"{i+1}️⃣ 구매계약 화면 본문 붙여넣기", height=100, key=f"raw_con_input_{i}")
        raw_cons.append(txt)
    
    if st.button("➕ 입력창 1개 더 추가하기", type="secondary"):
        st.session_state['con_input_count'] += 1
        st.rerun()

    st.markdown("<br>", unsafe_allow_html=True) 

    col_btn1, col_btn2 = st.columns(2)
    with col_btn1: start_con = st.button("🚀 다중 데이터 분석 및 폼 렌더링", width="stretch", type="primary")
    with col_btn2: 
        if st.button("🔄 전체 초기화", width="stretch"): 
            st.session_state.pop('multi_con_data', None)
            st.session_state['con_input_count'] = 2 
            for k in list(st.session_state.keys()):
                if k.startswith("raw_con_input_"):
                    del st.session_state[k]
            st.rerun()

    if start_con:
        valid_data = []
        for txt in raw_cons:
            if txt.strip(): 
                valid_data.append(extract_contract_info(txt))
        
        if valid_data:
            st.session_state['multi_con_data'] = valid_data
        else:
            st.warning("입력된 텍스트가 없습니다. 먼저 텍스트를 붙여넣어주세요.")

    if 'multi_con_data' in st.session_state:
        st.divider()
        st.write(f"### 📝 2. 발주요청서 입력 폼 (총 {len(st.session_state['multi_con_data'])}건 추출됨)")
        
        for i, cd in enumerate(st.session_state['multi_con_data']):
            with st.expander(f"📌 [계약 {i+1}] {cd.get('con_name', '건명 미상')}", expanded=True):
                
                c1, c2, c3 = st.columns(3)
                with c1: e_name = st.text_input("계약건명 (발주개요)", value=cd.get('con_name', ''), key=f"e_nm_{i}")
                with c2: e_amt = st.text_input("계약금액 (숫자만)", value=cd.get('con_amount', '0'), key=f"e_amt_{i}")
                with c3: e_ven = st.text_input("협력사 (업체명)", value=cd.get('vendor_name', ''), key=f"e_ven_{i}")

                c4, c5 = st.columns(2)
                with c4:
                    delivery_options = ["납품장소도", "현장설치도", "용역", "공사", "기타"]
                    e_deliv = st.selectbox("납품방법", delivery_options, key=f"e_del_{i}")
                with c5:
                    e_loc = st.text_input("납품장소", value=cd.get('location', ''), key=f"e_loc_{i}")

                st.caption("1) 실사용자 (직접 입력)")
                c_u1, c_u2, c_u3 = st.columns(3)
                with c_u1: u_nm = st.text_input("실사용자 성명", key=f"u_nm_{i}")
                with c_u2: u_ph = st.text_input("실사용자 연락처", key=f"u_ph_{i}")
                with c_u3: u_em = st.text_input("실사용자 이메일", key=f"u_em_{i}")

                st.caption("2) 구매/검수 담당자")
                buyer_name = cd.get('buyer_name', '')
                b_db = {"phone": "", "email": ""}
                for key, val in USER_DB.items():
                    if buyer_name in key: b_db = val; break
                    
                c_b1, c_b2, c_b3 = st.columns(3)
                with c_b1: b_nm = st.text_input("구매담당자", value=buyer_name, key=f"b_nm_{i}")
                with c_b2: i_nm = st.text_input("검수담당자", value="경제용", key=f"i_nm_{i}")
                
                amt_val = safe_int(e_amt)
                with c_b3: e_deposit = st.text_input("계약보증금액", value=cd.get('con_deposit', str(amt_val // 10)), key=f"e_dep_{i}")

                if st.button(f"📄 [{i+1}번] 워드파일 생성 준비", key=f"btn_gen_{i}"):
                    if not e_name or amt_val <= 0 or not e_ven or not u_nm:
                        st.error("🚨 필수 정보(계약건명, 계약금액, 협력사, 실사용자 성명)를 모두 입력해주세요.")
                    else:
                        try:
                            tpl_name = "template_over10mil.docx" if amt_val >= 10000000 else "template_under10mil.docx"
                            doc = DocxTemplate(tpl_name)
                            context = {
                                "구매담당자": b_nm, "구매담당자_연락처": b_db.get('phone', ''), "구매담당자_메일": b_db.get('email', ''),
                                "발주_개요": e_name, "계약금액": f"{amt_val:,}", 
                                "납품방법": e_deliv, "납품장소": e_loc,
                                "실사용자": u_nm, "실사용자_연락처": u_ph, "실사용자_메일": u_em,
                                "협력사": e_ven,
                                "검수담당자": i_nm, "검수담당자_연락처": "02-3408-4058", "검수담당자_메일": "rudwpdyd@sejong.ac.kr",
                                "계약보증금액": f"{safe_int(e_deposit):,}" if amt_val >= 10000000 else ""
                            }
                            doc.render(context)
                            ts = BytesIO()
                            doc.save(ts)
                            ts.seek(0)

                            st.success(f"✅ {i+1}번 발주서 생성 완료!")
                            st.download_button(
                                label=f"📥 [{i+1}번] 완성된 발주요청서 다운로드 (.docx)",
                                data=ts,
                                file_name=f"발주요청서_{e_ven}_{e_name}.docx",
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                key=f"btn_dl_{i}"
                            )
                        except Exception as e:
                            st.error(f"🚨 워드 파일 생성 중 오류 발생: {e}")

# ----------------- [3. 분석내용 불러오기] -----------------
elif menu == "🕒 분석내용 불러오기":
    st.title("🕒 과거 작업 내역 대시보드")
    st.info("이전에 작업하고 저장해둔 분석 내역을 엑셀 대장처럼 확인하고, 언제든 다시 불러와서 이어서 작업할 수 있습니다.")
    
    cache_list = get_cache_list()
    
    if not cache_list:
        st.warning("아직 저장된 작업 내역이 없습니다. '구매요청 분석' 탭에서 작업을 마치고 [저장] 버튼을 눌러주세요.")
    else:
        st.write("#### 📊 내장 데이터베이스 (목록 직접 수정 가능 ✏️)")
        st.info("💡 표 안의 글씨를 **더블클릭**해서 엑셀처럼 바로 수정할 수 있습니다. (고유번호인 구매요구번호는 수정 불가)\n\n수정을 마친 후 반드시 아래의 **[저장]** 버튼을 눌러주세요!")
        
        edited_df = st.data_editor(
            pd.DataFrame(cache_list), 
            width="stretch", 
            hide_index=True,
            disabled=["구매요구번호"] 
        )
        
        if st.button("💾 대시보드 표 수정사항을 DB에 즉시 반영하기", type="secondary"):
            try:
                with open(CACHE_FILE, "r", encoding="utf-8") as f:
                    cache_db = json.load(f)
                
                for _, row in edited_df.iterrows():
                    req_no = str(row["구매요구번호"]) 
                    if req_no in cache_db:
                        cache_db[req_no]["meta"]["과제번호"] = row["과제번호"]
                        cache_db[req_no]["meta"]["구매요청액"] = row["구매요청액"]
                        cache_db[req_no]["meta"]["연구책임자정보"] = row["연구책임자정보"]
                        cache_db[req_no]["meta"]["물품담당자"] = row["물품담당자"]
                        cache_db[req_no]["meta"]["업체명"] = row["업체명"]
                        cache_db[req_no]["meta"]["업체담당자정보"] = row["업체담당자정보"]
                
                with open(CACHE_FILE, "w", encoding="utf-8") as f:
                    json.dump(cache_db, f, ensure_ascii=False, indent=2)
                    
                st.success("✅ 대시보드 수정사항이 DB에 완벽하게 덮어쓰기 되었습니다!")
                time.sleep(1)
                st.rerun()
            except Exception as e:
                st.error(f"🚨 DB 업데이트 중 오류 발생: {e}")
        
        st.divider()
        st.write("#### 📂 데이터 복구하기")
        
        def execute_load_data():
            req_val = st.session_state.get("target_req_no", "").strip()
            if not req_val:
                st.session_state["load_msg"] = "🚨 구매요구번호를 입력해주세요."
                return
            
            loaded_data = load_from_cache(req_val)
            if loaded_data:
                raw_r = loaded_data.get("req_input", "")
                raw_f = loaded_data.get("fin_input", "")
                raw_i = loaded_data.get("itm_input", "")
                
                st.session_state["p_req"] = raw_r
                st.session_state["p_fin"] = raw_f
                st.session_state["p_itm"] = raw_i
                st.session_state[SESSION_AI_KEY] = loaded_data.get("ai_res", {})
                
                saved_req = loaded_data.get("req_type")
                if not saved_req:
                    _, is_goods = parse_goods_erp(raw_i)
                    _, is_svc = parse_service_erp(raw_i, "용역")
                    
                    if is_svc and not is_goods:
                        saved_req = "🛠️ 용역"
                    elif "공사" in raw_r:
                        saved_req = "🏗️ 공사"
                    else:
                        saved_req = "📦 물품"
                        
                saved_con = loaded_data.get("con_type")
                if not saved_con:
                    if "수의계약" in raw_r or "수의" in raw_r:
                        saved_con = "수의계약"
                    else:
                        saved_con = "비교견적"

                st.session_state["saved_req_type"] = saved_req
                st.session_state["saved_con_type"] = saved_con
                
                st.session_state[SESSION_REQ_KEY] = True
                st.session_state['is_saved'] = True  
                st.session_state["main_menu"] = "📋 구매요청 분석"
            else:
                st.session_state["load_msg"] = f"🚨 입력하신 번호 [{req_val}]에 해당하는 저장 데이터가 없습니다."

        c_load1, c_load2 = st.columns([3, 1])
        with c_load1:
            st.text_input("불러올 구매요구번호를 입력(또는 붙여넣기)하세요", placeholder="예: 202606170004", key="target_req_no")
            
        with c_load2:
            st.text(" ")
            st.button("🚀 데이터 불러오기", width="stretch", type="primary", on_click=execute_load_data)

        if "load_msg" in st.session_state:
            st.error(st.session_state["load_msg"])
            del st.session_state["load_msg"]