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
AI_MODEL = "gemini-3.5-flash-lite"
SESSION_AI_KEY = "ai_res"
SESSION_REQ_KEY = "req_analysis_done"
REQUIRED_FILES = ["template_under10mil.docx", "template_over10mil.docx"]

def get_config_path():
    p1 = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "app_config.json")
    if os.path.exists(p1): return p1
    p2 = os.path.join(os.path.dirname(os.path.abspath(__file__)), "app_config.json")
    if os.path.exists(p2): return p2
    if os.path.exists("app_config.json"): return os.path.abspath("app_config.json")
    return p1

DEFAULT_APP_CONFIG = {
    "app_title": "구매 업무 보조",
    "sidebar_title": "💼 구매 업무 보조",
    "menu_extractor": "💰 구매추출기",
    "menu_contract": "🤝 구매계약관리",
    "sidebar_status": "✅ 구매 보조 툴 가동 중",

    "headline_extractor": "💰 구매추출기",
    "extractor_step1": "📥 서류 및 정보 입력",
    "extractor_sec1": "📂 1. 기본 정보 선택",
    "extractor_sec2": "📝 2. ERP 텍스트 복사",
    "extractor_sec3": "📎 3. 파일 직접 업로드",
    "extractor_btn_analyze": "🚀 전체 데이터 통합 분석 시작",
    "extractor_step2": "📤 데이터 분석 및 검증 결과",
    "extractor_res_sec_erp": "📝 ERP 기본 정보",
    "extractor_res_sec_item": "📦 ERP 품목 내역",
    "extractor_res_sec_opinion": "✍️ 결재 의견 (자동 생성)",
    "extractor_res_sec_vendor": "📁 업체 정보 통합 및 추천 폴더명",
    "extractor_res_sec_bank": "🏦 통장사본 계좌정보",

    "headline_contract": "🤝 구매계약관리",
    "contract_desc": "여러 개의 계약 건을 한 번에 입력하고 분석하여, 각각의 워드 발주서를 빠르게 생성할 수 있습니다.",
    "btn_load_contract": "계약 데이터 불러오기",
    "contract_sec_user": "1) 실사용자 정보 (스마트 자동 추출)",
    "contract_sec_vendor": "2) 협력사(업체) 담당자 정보 (이메일 미기재 시 직접 입력 가능)",
    "contract_sec_staff": "3) 구매/검수 담당자 및 계약보증금",

    "headline_history": "🕒 과거 작업 내역 대시보드",
    "headline_admin": "⚙️ 관리자 설정",

    "admin_pin": "1234",
    "buyer_staff": {
        "권순철 사원": {"name": "권순철", "phone": "02-6935-2624", "email": "hagrid0924@sejong.ac.kr"},
        "김민석 사원": {"name": "김민석", "phone": "02-3408-3032", "email": "sejongtwice@sejong.ac.kr"}
    },
    "inspector_info": {
        "name": "경제용",
        "phone": "02-3408-4058",
        "email": "rudwpdyd@sejong.ac.kr"
    }
}

def load_app_config():
    c_path = get_config_path()
    if os.path.exists(c_path):
        try:
            with open(c_path, "r", encoding="utf-8") as f:
                data = json.load(f)
                merged = DEFAULT_APP_CONFIG.copy()
                merged.update(data)
                return merged
        except Exception:
            pass
    return DEFAULT_APP_CONFIG.copy()

def save_app_config(new_cfg):
    global APP_CONFIG
    c_path = get_config_path()
    try:
        with open(c_path, "w", encoding="utf-8") as f:
            json.dump(new_cfg, f, ensure_ascii=False, indent=2)
        if "APP_CONFIG" in globals():
            APP_CONFIG.update(new_cfg)
        return True
    except Exception:
        return False

APP_CONFIG = load_app_config()

try:
    st.set_page_config(page_title=APP_CONFIG.get("app_title", "구매 업무 보조"), page_icon="💼", layout="wide")
except Exception:
    pass

load_dotenv()

def get_secret(primary_key, aliases=None):
    if aliases is None: aliases = []
    all_keys = [primary_key] + aliases
    
    # 1. os.getenv 환경변수 검색
    for k in all_keys:
        val = os.getenv(k) or os.getenv(k.lower()) or os.getenv(k.upper())
        if val and str(val).strip():
            return str(val).strip()
            
    # 2. st.secrets 검색
    try:
        if hasattr(st, "secrets"):
            for k in all_keys:
                if k in st.secrets and str(st.secrets[k]).strip():
                    return str(st.secrets[k]).strip()
                if k.lower() in st.secrets and str(st.secrets[k.lower()]).strip():
                    return str(st.secrets[k.lower()]).strip()
                if k.upper() in st.secrets and str(st.secrets[k.upper()]).strip():
                    return str(st.secrets[k.upper()]).strip()
            # secrets 딕셔너리 순회 (대소문자 무시)
            for s_key in list(st.secrets.keys()):
                for k in all_keys:
                    if s_key.lower() == k.lower() and str(st.secrets[s_key]).strip():
                        return str(st.secrets[s_key]).strip()
    except Exception:
        pass
        
    return None

API_KEY = get_secret("GEMINI_API_KEY", ["GEMINI_KEY", "GOOGLE_API_KEY", "API_KEY"])
NTS_API_KEY = get_secret("NTS_API_KEY", ["NTS_KEY", "ODCLOUD_API_KEY"])

missing_files = [f for f in REQUIRED_FILES if not os.path.exists(f)]
if missing_files:
    st.error(f"🚨 필수 파일 누락으로 실행 불가: {', '.join(missing_files)}")
    st.stop()

missing_keys = []
if not API_KEY: missing_keys.append("구글 제미나이 키 (GEMINI_API_KEY)")
if not NTS_API_KEY: missing_keys.append("국세청 사업자조회 키 (NTS_API_KEY)")

if missing_keys:
    st.error(f"🚨 **API 키 누락 감지**: 다음 키를 찾을 수 없습니다 ➔ **{', '.join(missing_keys)}**")
    st.info(
        "💡 **해결 방법 (Streamlit Cloud)**:\n"
        "1. Streamlit Community Cloud 관리 페이지의 앱 설정(`Settings` ➔ `Secrets`)으로 이동합니다.\n"
        "2. 아래와 같이 **두 가지 키가 모두** 작성되어 있는지 확인하세요 (반드시 큰따옴표 포함):\n\n"
        "```toml\n"
        'GEMINI_API_KEY = "구글_AI_스튜디오_API_키"\n'
        'NTS_API_KEY = "국세청_공공데이터_API_키"\n'
        "```\n"
        "*(국세청 키가 누락되어도 이 경고가 뜰 수 있습니다)*"
    )
    st.stop()

ai_client = genai.Client(api_key=API_KEY)

INSPECTOR_INFO = APP_CONFIG.get("inspector_info", DEFAULT_APP_CONFIG["inspector_info"])
BUYER_STAFF = APP_CONFIG.get("buyer_staff", DEFAULT_APP_CONFIG["buyer_staff"])

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
for _bk, _bv in BUYER_STAFF.items():
    USER_DB[_bk] = {"phone": _bv.get("phone", ""), "email": _bv.get("email", "")}
    b_first = _bk.split()[0]
    USER_DB[b_first] = {"phone": _bv.get("phone", ""), "email": _bv.get("email", "")}
CACHE_FILE = "local_cache.json"
GSHEETS_URL = get_secret("GSHEETS_URL") or "https://script.google.com/macros/s/AKfycbx3qKiOuC7t0J5g7ZJWWccXzQTld728WRPqbcyf7jBD9VHFZNYv0XhqqGug--imgH6U/exec"

def save_to_cache(req_no, meta_data, raw_data):
    # 1. 로컬 캐시 파일 저장 (오프라인 / 빠른 로드용)
    cache = {}
    if os.path.exists(CACHE_FILE):
        try:
            with open(CACHE_FILE, "r", encoding="utf-8") as f: cache = json.load(f)
        except: pass
    
    existing_record = cache.get(req_no, {"meta": {}, "raw": {}})
    existing_record["meta"].update(meta_data)
    existing_record["raw"].update(raw_data)
    cache[req_no] = existing_record
    
    try:
        with open(CACHE_FILE, "w", encoding="utf-8") as f:
            json.dump(cache, f, ensure_ascii=False, indent=2)
    except Exception as e:
        log(f"로컬 캐시 저장 오류: {e}")

    # 2. 구글 시트 영구 저장소 (Google Apps Script API) 연동
    if GSHEETS_URL:
        try:
            gsheet_payload = dict(meta_data)
            # Raw_data를 직렬화하여 구글 시트의 Raw_data 열에도 영구 저장
            gsheet_payload["Raw_data"] = json.dumps(raw_data, ensure_ascii=False)
            res = requests.post(GSHEETS_URL, json={"data": gsheet_payload}, timeout=10)
            if res.status_code == 200:
                log(f"구글 시트 저장 성공 ({req_no})")
            else:
                log(f"구글 시트 저장 실패 ({res.status_code}): {res.text}")
        except Exception as e:
            log(f"구글 시트 API 연동 실패: {e}")

def get_cache_list():
    result_map = {}
    
    # 1. 구글 시트 영구 저장소에서 먼저 최신 목록 조회
    if GSHEETS_URL:
        try:
            res = requests.get(GSHEETS_URL, timeout=8)
            if res.status_code == 200:
                rows = res.json()
                if isinstance(rows, list):
                    for row in rows:
                        req_no = str(row.get("구매요구번호", "")).strip()
                        if not req_no or req_no.startswith("TEST"): 
                            continue
                        
                        raw_str = row.get("Raw_data", "")
                        raw_obj = {}
                        if raw_str:
                            try: raw_obj = json.loads(raw_str)
                            except: pass
                            
                        vendor_mgr = row.get("업체담당자정보") or row.get("업체담당자 정보") or ""
                        remarks_val = row.get("비고사항") or row.get("물품비고") or ""
                        if not remarks_val and raw_obj:
                            remarks_val = raw_obj.get("remarks", "")
                            
                        item_record = {
                            "구매요구번호": req_no,
                            "과제번호": row.get("과제번호", ""),
                            "구매요청액": row.get("구매요청액", ""),
                            "연구책임자정보": row.get("연구책임자정보", ""),
                            "물품담당자": row.get("물품담당자", ""),
                            "업체명": row.get("업체명", ""),
                            "업체담당자정보": vendor_mgr,
                            "비고사항": remarks_val
                        }
                        result_map[req_no] = (item_record, raw_obj, row)
        except Exception as e:
            log(f"구글 시트 목록 조회 실패 (로컬 캐시 사용): {e}")

    # 2. 로컬 캐시 병합 (오프라인 상태이거나 미동기화 건 보완)
    if os.path.exists(CACHE_FILE):
        try:
            with open(CACHE_FILE, "r", encoding="utf-8") as f:
                local_db = json.load(f)
            for req_no, v in local_db.items():
                if str(req_no).startswith("TEST"): continue
                meta = v.get("meta", {})
                raw = v.get("raw", {})
                if req_no not in result_map:
                    result_map[req_no] = ({
                        "구매요구번호": meta.get("구매요구번호", req_no),
                        "과제번호": meta.get("과제번호", ""),
                        "구매요청액": meta.get("구매요청액", ""),
                        "연구책임자정보": meta.get("연구책임자정보", ""),
                        "물품담당자": meta.get("물품담당자", ""),
                        "업체명": meta.get("업체명", ""),
                        "업체담당자정보": meta.get("업체담당자정보") or meta.get("업체담당자 정보", ""),
                        "비고사항": meta.get("비고사항") or meta.get("물품비고", "")
                    }, raw, meta)
                else:
                    existing_item, existing_raw, existing_row = result_map[req_no]
                    if not existing_raw and raw:
                        result_map[req_no] = (existing_item, raw, existing_row)
        except Exception as e:
            log(f"로컬 캐시 읽기 실패: {e}")

    # 3. 로컬 파일에 동기화 상태 캐싱
    if result_map:
        try:
            sync_db = {}
            if os.path.exists(CACHE_FILE):
                try:
                    with open(CACHE_FILE, "r", encoding="utf-8") as f: sync_db = json.load(f)
                except: pass
            for req_no, (m, r, raw_meta) in result_map.items():
                if req_no not in sync_db: sync_db[req_no] = {"meta": {}, "raw": {}}
                sync_db[req_no]["meta"].update(raw_meta)
                if r: sync_db[req_no]["raw"].update(r)
            with open(CACHE_FILE, "w", encoding="utf-8") as f:
                json.dump(sync_db, f, ensure_ascii=False, indent=2)
        except: pass

    res_list = [item[0] for item in result_map.values()]
    return res_list[::-1]

def load_from_cache(req_no):
    req_no_str = str(req_no).strip()
    # 1. 로컬 캐시 조회
    if os.path.exists(CACHE_FILE):
        try:
            with open(CACHE_FILE, "r", encoding="utf-8") as f: 
                data = json.load(f).get(req_no_str, {}).get("raw", None)
                if data: return data
        except: pass

    # 2. 로컬에 없으면 구글 시트에서 직접 조회
    if GSHEETS_URL:
        try:
            res = requests.get(GSHEETS_URL, timeout=8)
            if res.status_code == 200:
                for row in res.json():
                    if str(row.get("구매요구번호", "")).strip() == req_no_str:
                        raw_str = row.get("Raw_data", "")
                        if raw_str:
                            return json.loads(raw_str)
        except Exception as e:
            log(f"구글 시트 로드 실패: {e}")

    return None

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

    # 표 헤더에 '비고' 열이 명시되어 있는지 검색
    remark_col_idx = -1
    for line in text.strip().split('\n'):
        if not line.strip(): continue
        tokens = [t.strip() for t in line.split('\t')]
        for i, t in enumerate(tokens):
            if '비고' in t:
                remark_col_idx = i
                break
        if remark_col_idx != -1:
            break

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
                
                # 비고(참고사항) 추출 로직
                remark = ""
                if remark_col_idx != -1 and len(tokens) > remark_col_idx:
                    remark = tokens[remark_col_idx].strip()
                elif len(tokens) > con_idx + 2:
                    # rERP 물품 테이블 컬럼 순서:
                    # [con_idx: 물품담당자] -> [con_idx+1: 용도(예: 회의용 테이블)] -> [con_idx+2: 비고]
                    # 따라서 con_idx+1(용도)을 건너뛰고 con_idx+2의 실제 비고란을 추출합니다.
                    remark = tokens[con_idx + 2].strip()

                # 공란 및 무의미한 텍스트 필터링 (공란이면 비고 추출 안 함)
                if remark in ['-', '--', '없음', '해당없음', 'N/A', 'nan', '미기재', '없음.']:
                    remark = ""

                items.append({
                    'name': name_str, 'category': raw_cat if raw_cat else "공란", 
                    'reg_status': reg_status, 'contact': parse_contact(contact_raw), 
                    'qty': qty, 'unit': unit, 'price': price, 'unit_price': unit_price,
                    'remark': remark
                })
        except Exception as e: 
            log(f"물품 파싱 에러: {e}")
            continue 
    return items, valid_format

def parse_service_erp(text, force_category):
    items = []
    valid_format = False
    if not text.strip(): return items, valid_format

    remark_col_idx = -1
    for line in text.strip().split('\n'):
        if not line.strip(): continue
        tokens = [t.strip() for t in line.split('\t')]
        for i, t in enumerate(tokens):
            if '비고' in t:
                remark_col_idx = i
                break
        if remark_col_idx != -1:
            break

    for line in text.strip().split('\n'):
        if not line.strip(): continue
        tokens = line.split('\t')
        try:
            if len(tokens) >= 15:
                valid_format = True
                name_str = tokens[1].strip() if len(tokens) > 1 else "무명"
                if len(tokens) > 2 and tokens[2].strip(): name_str = f"{name_str} ({tokens[2].strip()})"
                
                remark = ""
                if remark_col_idx != -1 and len(tokens) > remark_col_idx:
                    remark = tokens[remark_col_idx].strip()
                elif len(tokens) > 16:
                    remark = tokens[16].strip()

                if remark in ['-', '--', '없음', '해당없음', 'N/A', 'nan', '미기재', '없음.']:
                    remark = ""

                items.append({
                    'name': name_str, 
                    'category': tokens[9].strip() if tokens[9].strip() else force_category, 
                    'reg_status': "N/A(용역)", 
                    'contact': parse_contact(tokens[14].strip()), 
                    'qty': safe_int(tokens[3]), 'unit': tokens[4].strip(), 
                    'price': safe_int(tokens[8]), 'unit_price': safe_int(tokens[5]),
                    'remark': remark
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
        ("gemini-3.5-flash-lite", 1),
        ("gemini-3.8-flash", 1),
        ("gemini-2.5-flash", 1)
    ]
    
    failure_history = []
    
    for model_name, max_retries in fallback_chain:
        for attempt in range(max_retries):
            attempt_label = f"{model_name} (시도 {attempt+1}/{max_retries})"
            try:
                response = ai_client.models.generate_content(model=model_name, contents=contents, config=gen_config)
                
                resp_text = ""
                finish_reason = None
                
                if hasattr(response, 'candidates') and response.candidates:
                    cand = response.candidates[0]
                    finish_reason = getattr(cand, 'finish_reason', None)
                    if hasattr(finish_reason, 'name'):
                        finish_reason = finish_reason.name
                    elif finish_reason is not None:
                        finish_reason = str(finish_reason)

                try:
                    resp_text = response.text or ""
                except Exception as text_e:
                    log(f"⚠️ {attempt_label} text 추출 불가: {text_e} (finish_reason: {finish_reason})")
                    is_safety = "SAFETY" in str(finish_reason).upper() or "SAFETY" in str(text_e).upper()
                    failure_history.append({
                        "model": model_name,
                        "attempt": attempt + 1,
                        "type": "SAFETY_BLOCKED" if is_safety else "CONTENT_BLOCKED",
                        "summary": f"🛑 안전 정책(Safety Filter)에 의해 응답 생성이 차단되었습니다. (사유: {finish_reason})" if is_safety else f"⚠️ 응답 본문을 읽을 수 없습니다. (사유: {finish_reason or text_e})",
                        "error": f"finish_reason: {finish_reason}, 에러: {str(text_e)}",
                        "response_text": ""
                    })
                    time.sleep(1)
                    continue

                if resp_text.strip():
                    parsed = safe_json_load(resp_text.strip())
                    if parsed: 
                        log(f"✅ AI 분석 성공: {model_name} (시도 {attempt+1})")
                        return parsed, []
                    else: 
                        log(f"⚠️ {attempt_label} JSON 파싱 실패 (응답 텍스트 길이: {len(resp_text)})")
                        failure_history.append({
                            "model": model_name,
                            "attempt": attempt + 1,
                            "type": "PARSE_ERROR",
                            "summary": "⚠️ 제미나이가 응답했으나, 시스템 규격(JSON)이 아닌 일반 텍스트/안내문으로 답변하여 자동 분석에 실패했습니다.",
                            "error": "JSON 파싱 불가 (중괄호 { } 형태의 유효한 JSON 없음)",
                            "response_text": resp_text.strip()
                        })
                else:
                    failure_history.append({
                        "model": model_name,
                        "attempt": attempt + 1,
                        "type": "EMPTY_RESPONSE",
                        "summary": f"⚠️ 제미나이 응답 내용이 비어 있습니다. (finish_reason: {finish_reason})",
                        "error": f"빈 텍스트 응답 (finish_reason: {finish_reason})",
                        "response_text": ""
                    })
            except Exception as e: 
                err_str = str(e)
                log(f"⚠️ AI 호출 실패 ({attempt_label}): {err_str}")
                
                lower_err = err_str.lower()
                if "429" in err_str or "resource_exhausted" in lower_err or "quota" in lower_err or "rate limit" in lower_err:
                    err_type = "QUOTA_EXCEEDED"
                    summary = "🚨 할당량/한도 초과 (429 RESOURCE_EXHAUSTED): 제미나이 API 분당 요청 수(RPM) 또는 일일 무료 한도에 도달했습니다."
                elif "401" in err_str or "403" in err_str or "permission_denied" in lower_err or "api_key" in lower_err:
                    err_type = "AUTH_ERROR"
                    summary = "🔑 API 인증 실패 (401/403): GEMINI_API_KEY가 유효하지 않거나 권한이 없습니다."
                elif "404" in err_str or "not_found" in lower_err:
                    err_type = "MODEL_NOT_FOUND"
                    summary = f"🔍 모델 미지원/미발견 (404 NOT_FOUND): '{model_name}' 모델을 사용할 수 없습니다."
                elif "413" in err_str or "too large" in lower_err or "payload" in lower_err:
                    err_type = "PAYLOAD_TOO_LARGE"
                    summary = "📦 용량/토큰 초과: 첨부파일 용량이 너무 크거나 1회 토큰 한도를 초과했습니다."
                elif "safety" in lower_err:
                    err_type = "SAFETY_BLOCKED"
                    summary = "🛑 안전 정책(Safety Filter)에 의해 요청이 차단되었습니다."
                else:
                    err_type = "API_ERROR"
                    summary = f"❗ 서버/API 호출 오류 ({type(e).__name__}): {err_str[:180]}"

                failure_history.append({
                    "model": model_name,
                    "attempt": attempt + 1,
                    "type": err_type,
                    "summary": summary,
                    "error": err_str,
                    "response_text": ""
                })
                time.sleep(1)

    log("🚨 모든 Fallback 모델 호출 실패. 수동 입력 모드로 전환합니다.")
    return None, failure_history

def extract_contract_info(text):
    clean_text = text.replace('\n', ' ').replace('\r', '').replace('\t', ' ')
    info = {
        'req_no': re.search(r'구매요구번호\s*\*?\s*(\d+)', clean_text),
        'con_name': re.search(r'계약건명\s*(.*?)\s*납품기간', clean_text),
        'con_amount': re.search(r'계약금액\s*([\d,]+)', clean_text),
        'con_deposit': re.search(r'계약보증금\s*([\d,]+)', clean_text),
        'vendor_name': re.search(r'계약업체명\s*(.*?)\s*사업자번호', clean_text),
        'buyer_name': (
            re.search(r'계약담당자\s*(.*?)\s*접수일자', clean_text) or 
            re.search(r'계약담당자\s*([^\s]+)', clean_text) or
            re.search(r'계약자\s*([^\s]+)', clean_text)
        ),
        'email': re.search(r'e-mail주소\s*([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})', clean_text)
    }
    result = {k: v.group(1).strip() if v else "" for k, v in info.items()}
    result['con_amount'] = str(safe_int(result.get('con_amount', 0)))
    if result.get('con_deposit'):
        result['con_deposit'] = str(safe_int(result['con_deposit']))

    # 💡 지원사업 파싱 (신임교원지원사업 여부 판별)
    supp_m = re.search(r'지원사업\s*([^\t\r\n\s]+)', clean_text)
    supp_txt = supp_m.group(1).strip() if supp_m else ""
    is_new_fac = ('신임교원' in supp_txt) or ('신임교원' in clean_text)
    result['support_name'] = supp_txt
    result['project_type'] = "신임교원지원사업" if is_new_fac else "일반 중앙구매"

    # 💡 연구책임자정보 직접 추출 (rERP 본문 복사 시 자동 채우기)
    pi_match = re.search(r'연구책임자정보\s*([^/\s]+)\s*/\s*([^/\s]+)\s*/\s*([^/\s]+)\s*/\s*([0-9\-]+)\s*/\s*([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})', clean_text)
    if pi_match:
        result['pi_name'] = pi_match.group(1).strip()
        result['pi_position'] = pi_match.group(2).strip()
        result['pi_dept'] = pi_match.group(3).strip()
        raw_ph = pi_match.group(4).strip()
        dig = re.sub(r'[^\d]', '', raw_ph)
        if len(dig) == 11 and dig.startswith('01'):
            result['pi_phone'] = f"{dig[:3]}-{dig[3:7]}-{dig[7:]}"
        elif len(dig) == 10 and dig.startswith('01'):
            result['pi_phone'] = f"{dig[:3]}-{dig[3:6]}-{dig[6:]}"
        else:
            result['pi_phone'] = raw_ph
        result['pi_email'] = pi_match.group(5).strip()
    
    loc_match = re.search(r'납품장소\s*(.*?)\s*계약금액', clean_text)
    result['location'] = ""
    if loc_match and '/' in loc_match.group(1):
        parts = [p.strip() for p in loc_match.group(1).split('/')]
        if len(parts) >= 3:
            room_match = re.search(r'(\d+)', parts[-1])
            if room_match: result['location'] = f"{parts[0]} {int(room_match.group(1))}호"

    # 본문 내 실사용자/업체 메모 패턴이 직접 포함된 경우 함께 추출
    result['smart_memo'] = parse_smart_contact_memo(text)
    
    # 💡 사용자가 직접 입력한 메모/정보가 최우선입니다.
    # 사용자가 입력한 실사용자 정보가 없을 때만(공란일 때만) 계약관리의 연구책임자정보를 fallback으로 반영합니다.
    if not result['smart_memo'].get('user_name') and result.get('pi_name'):
        result['smart_memo']['user_name'] = f"{result['pi_name']} {result.get('pi_position', '')}".strip()
        if not result['smart_memo'].get('user_phone'):
            result['smart_memo']['user_phone'] = result.get('pi_phone', '')
        if not result['smart_memo'].get('user_email'):
            result['smart_memo']['user_email'] = result.get('pi_email', '')

    # 💡 rERP 계약업체 대표자가 명시되어 있고 메모상 업체담당자가 없거나 헤더명인 경우 대표자명 반영
    rep_m = re.search(r'대표자\s*([^\t\r\n\s]+(?:\s+[^\t\r\n\s]+)*?)\s*(?:전화번호|법인번호|사업자번호)', clean_text)
    if rep_m:
        cur_vrep = result['smart_memo'].get('vendor_rep', '')
        if not cur_vrep or cur_vrep in ['계약업체정보', '업체']:
            result['smart_memo']['vendor_rep'] = rep_m.group(1).strip()

    return result

def parse_smart_contact_memo(text):
    """
    노션 비고란 등에서 복사한 실사용자/업체담당자 비정형 텍스트를 정밀 분리합니다.
    (예: 실사용자: 정영권 교수 01033246327 ykjung@sejong.ac.kr
         업체담당자 / 김철수 / 010-9999-8888)
    """
    result = {
        'user_name': '', 'user_phone': '', 'user_email': '',
        'vendor_rep': '', 'vendor_phone': '', 'vendor_email': ''
    }
    if not text or not text.strip():
        return result

    phone_pattern = r'(01[016789]-?\d{3,4}-?\d{4}|02-?\d{3,4}-?\d{4}|\b\d{4}\b)'
    email_pattern = r'([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})'

    normalized_text = text
    # 만약 엔터 없이 한 줄에 실사용자와 업체담당자가 함께 입력된 경우 분리
    vendor_split_match = re.search(r'[\r\n\s/]+(?=(?:업체담당자|업체|협력사|거래처|vendor|공급사)[\s:：/\-])', normalized_text, flags=re.IGNORECASE)
    if vendor_split_match and ('실사용자' in normalized_text or '사용자' in normalized_text):
        split_pos = vendor_split_match.start()
        user_part = normalized_text[:split_pos]
        vendor_part = normalized_text[split_pos:]
        lines = [user_part.strip(), vendor_part.strip()]
    else:
        lines = [line.strip() for line in normalized_text.strip().split('\n') if line.strip()]

    def extract_from_line(line, prefix_patterns):
        em = ''
        em_m = re.search(email_pattern, line)
        if em_m:
            em = em_m.group(1).strip()
            line = line.replace(em, ' ')
        
        ph = ''
        ph_m = re.search(phone_pattern, line)
        if ph_m:
            raw_ph = ph_m.group(1).strip()
            digits = re.sub(r'[^\d]', '', raw_ph)
            if len(digits) == 11 and digits.startswith('01'):
                ph = f'{digits[:3]}-{digits[3:7]}-{digits[7:]}'
            elif len(digits) == 10 and digits.startswith('01'):
                ph = f'{digits[:3]}-{digits[3:6]}-{digits[6:]}'
            elif len(digits) == 10 and digits.startswith('02'):
                ph = f'{digits[:2]}-{digits[2:6]}-{digits[6:]}'
            elif len(digits) == 9 and digits.startswith('02'):
                ph = f'{digits[:2]}-{digits[2:5]}-{digits[5:]}'
            elif len(digits) == 4:
                ph = f'02-3408-{digits}'
            else:
                ph = raw_ph
            line = line.replace(raw_ph, ' ')
        
        clean = line.strip()
        clean = re.sub(r'^[^\w가-힣]+', '', clean).strip()
        clean = re.sub(prefix_patterns, '', clean, flags=re.IGNORECASE).strip()
        clean = re.sub(r'^[^\w가-힣]+', '', clean).strip()
        clean = re.sub(r'[\(\)\[\]/,:：]', ' ', clean).strip()
        clean = re.sub(r'\s+', ' ', clean).strip()
        return clean, ph, em

    user_prefix = r'^(실사용자|사용자|담당교수|교수|연구원|신청자|성명|이름)[\s:：/\-]*'
    vendor_prefix = r'^(업체담당자|업체|협력사|거래처|vendor|공급사|담당자|성명|이름)[\s:：/\-]*'

    for idx, line in enumerate(lines):
        # 결재선(기안자/과장/부장 등) 및 rERP 시스템 헤더/칼럼 행은 메모 분석 대상에서 제외
        if any(kw in line for kw in [
            '과장', '부장', '단장', '결재', '구매계약서', '상세보기', '재원정보', 
            '심의정보', '첨부문서', '물품내역', '연구책임자정보', '과제정보', '사업장', 
            '구매요구번호', '진행상태', '보완사항', '구매진행상세', '계약사유', '계약업체명', 
            '계약건명', '계약금액'
        ]):
            continue

        is_vendor = any(kw in line for kw in ['업체담당자', '업체', '협력사', '거래처', 'vendor', '공급사'])
        is_user = any(kw in line for kw in ['실사용자', '사용자', '담당교수', '교수', '연구원', '신청자'])

        target = None
        if is_vendor:
            target = 'vendor'
        elif is_user:
            target = 'user'
        elif len(lines) <= 2:
            target = 'user' if idx == 0 else 'vendor'
        else:
            continue

        if target == 'user':
            nm, ph, em = extract_from_line(line, user_prefix)
            if nm and not result['user_name']: result['user_name'] = nm
            if ph and not result['user_phone']: result['user_phone'] = ph
            if em and not result['user_email']: result['user_email'] = em
        else:
            nm, ph, em = extract_from_line(line, vendor_prefix)
            if nm and not result['vendor_rep']: result['vendor_rep'] = nm
            if ph and not result['vendor_phone']: result['vendor_phone'] = ph
            # 업체담당자 라인에 이메일이 없는 경우 공란 유지 (실사용자 이메일 침범 방지)
            if em and not result['vendor_email']: result['vendor_email'] = em

    return result

# ==========================================
# 5. 화면 UI 구성
# ==========================================
def editable_headline(config_key, default_text, level="h3", prefix="", key_suffix=""):
    """
    관리자 모드이고 ui_edit_mode가 켜져 있으면 인라인 수정 UI를 보여주고,
    평상시에는 일반 텍스트/헤더로 렌더링하는 범용 헬퍼 함수
    """
    current_val = APP_CONFIG.get(config_key, default_text)
    is_edit = st.session_state.get("is_admin_auth", False) and st.session_state.get("ui_edit_mode", False)
    
    if is_edit:
        with st.container(border=True):
            c_in1, c_in2 = st.columns([5, 1])
            with c_in1:
                label_text = f"✏️ [위치: {prefix}]" if prefix else f"✏️ [문구 편집: {config_key}]"
                new_val = st.text_input(label_text, value=current_val, key=f"inline_{config_key}_{key_suffix}")
            with c_in2:
                st.markdown("<div style='height: 28px;'></div>", unsafe_allow_html=True)
                if st.button("💾", key=f"btn_s_{config_key}_{key_suffix}", help="이 위치의 문구 저장"):
                    APP_CONFIG[config_key] = new_val.strip() or default_text
                    save_app_config(APP_CONFIG)
                    st.success("저장됨!")
                    time.sleep(0.3)
                    st.rerun()
    else:
        if level == "h1":
            st.title(current_val)
        elif level == "h2":
            st.markdown(f"## {current_val}")
        elif level == "h3":
            st.markdown(f"### {current_val}")
        elif level == "h4":
            st.markdown(f"#### {current_val}")
        elif level == "h5":
            st.markdown(f"##### {current_val}")
        elif level == "caption":
            st.caption(current_val)
        elif level == "info":
            st.info(current_val)
        else:
            st.write(current_val)

def render_top_header(title_key, default_title):
    title_text = APP_CONFIG.get(title_key, default_title)
    is_auth = st.session_state.get("is_admin_auth", False)
    safe_key_suffix = f"_{abs(hash(title_key)) % 10000}"
    
    if is_auth:
        col_t, col_edit, col_l = st.columns([7, 4, 1])
        with col_t:
            is_edit_mode = st.session_state.get("ui_edit_mode", False)
            if is_edit_mode:
                c_in1, c_in2 = st.columns([5, 1])
                with c_in1:
                    new_t = st.text_input("✏️ [화면 대제목 편집]", value=title_text, key=f"top_title_{title_key}")
                with c_in2:
                    st.markdown("<div style='height: 28px;'></div>", unsafe_allow_html=True)
                    if st.button("💾", key=f"btn_s_top_{title_key}", help="대제목 저장"):
                        APP_CONFIG[title_key] = new_t.strip() or default_title
                        save_app_config(APP_CONFIG)
                        st.rerun()
            else:
                st.title(title_text)
                
        with col_edit:
            st.markdown("<div style='height: 14px;'></div>", unsafe_allow_html=True)
            st.toggle("✏️ 화면 문구 직접 수정 모드", key="ui_edit_mode", help="실제 화면 위치에서 타이틀/헤더를 직접 수정합니다.")
            
        with col_l:
            st.markdown("<div style='height: 14px;'></div>", unsafe_allow_html=True)
            with st.popover("🔓", help="관리자 모드 활성화됨"):
                st.markdown("##### ⚙️ 관리자 모드")
                if st.button("⚙️ 설정 화면 이동", key=f"btn_go{safe_key_suffix}", use_container_width=True):
                    st.session_state["nav_target"] = "⚙️ 관리자 설정"
                    st.rerun()
                if st.button("🔒 로그아웃 (숨김)", key=f"btn_out{safe_key_suffix}", use_container_width=True):
                    st.session_state["is_admin_auth"] = False
                    st.session_state["ui_edit_mode"] = False
                    st.session_state["nav_target"] = APP_CONFIG.get("menu_extractor", "💰 구매추출기")
                    st.rerun()
                    
        if st.session_state.get("ui_edit_mode", False):
            st.info("💡 **화면 문구 직접 수정 모드 ON**: 화면 속 각 제목/헤드라인 위치에 입력창이 열려 있습니다. 원하는 단어로 바꾸고 [💾]를 누르면 즉시 저장됩니다.")
    else:
        col_t, col_l = st.columns([11, 1])
        with col_t:
            st.title(title_text)
        with col_l:
            st.markdown("<div style='height: 14px;'></div>", unsafe_allow_html=True)
            with st.popover("🔒", help="관리자 인증 (클릭)"):
                st.markdown("##### 🔑 관리자 PIN")
                pin_val = st.text_input("PIN 입력", type="password", key=f"pin_in{safe_key_suffix}", placeholder="4자리 PIN")
                if st.button("인증", key=f"btn_auth{safe_key_suffix}", use_container_width=True, type="primary"):
                    saved_pin = str(APP_CONFIG.get("admin_pin", "1234")).strip()
                    if pin_val.strip() == saved_pin:
                        st.session_state["is_admin_auth"] = True
                        st.session_state["nav_target"] = "⚙️ 관리자 설정"
                        st.rerun()
                    else:
                        st.error("PIN 번호 불일치")

with st.sidebar:
    s_title = APP_CONFIG.get("sidebar_title", "💼 구매 업무 보조")
    st.markdown(f"## {s_title}")
    
    m_ext = APP_CONFIG.get("menu_extractor", "💰 구매추출기")
    m_con = APP_CONFIG.get("menu_contract", "🤝 구매계약관리")
    
    # 예약된 이동 타겟이 있으면 radio 생성 전에 반영
    if "nav_target" in st.session_state:
        target = st.session_state.pop("nav_target")
        if target == "💰 구매추출기": target = m_ext
        elif target == "🤝 구매계약관리": target = m_con
        st.session_state["main_menu"] = target
        
    # 관리자 인증 시에만 '⚙️ 관리자 설정' 메뉴가 노출됨
    base_menus = [m_ext, m_con]
    if st.session_state.get("is_admin_auth", False):
        base_menus.append("⚙️ 관리자 설정")
        
    current_menu_choice = st.session_state.get("main_menu", m_ext)
    if current_menu_choice not in base_menus:
        current_menu_choice = m_ext
        st.session_state["main_menu"] = current_menu_choice

    menu = st.radio(
        "메뉴 선택", 
        base_menus, 
        key="main_menu"
    )
    st.divider()
    
    # 💡 과거 작업내역 불러오기는 평소엔 닫혀있는 아코디언으로 보관 (화면 복잡도 최소화)
    with st.expander("🕒 과거 분석내역 불러오기 (보관함)", expanded=False):
        if st.checkbox("보관함 대시보드 열기", key="open_history_db"):
            menu = "🕒 분석내용 불러오기"
            
    if st.session_state.get("is_admin_auth", False):
        st.caption("🔓 관리자 모드 활성화됨")
        if st.button("🔒 관리자 로그아웃", key="btn_side_logout", use_container_width=True):
            st.session_state["is_admin_auth"] = False
            st.session_state["ui_edit_mode"] = False
            st.session_state["nav_target"] = m_ext
            st.rerun()
    else:
        st.caption(APP_CONFIG.get("sidebar_status", "✅ 구매 보조 툴 가동 중"))

# ----------------- [1. 구매요청 분석] -----------------
if menu in [m_ext, "💰 구매추출기", "📋 구매요청 분석"]:
    render_top_header("headline_extractor", "💰 구매추출기")
    
    def on_critical_condition_change():
        if SESSION_REQ_KEY in st.session_state:
            del st.session_state[SESSION_REQ_KEY]
        st.session_state['api_failed'] = False
        st.session_state['ai_errors'] = []
        st.session_state['is_saved'] = False
    
    editable_headline("extractor_step1", "📥 서류 및 정보 입력", level="h3", prefix="1단계 전체 구역 제목")
    
    editable_headline("extractor_sec1", "📂 1. 기본 정보 선택", level="h4", prefix="1번 기본 정보 선택 소제목")
    c_type1, c_type2 = st.columns(2)

    with c_type1:
        req_type = st.radio("구매 유형", ["📦 물품", "🛠️ 용역", "🏗️ 공사"], horizontal=True, key="saved_req_type", on_change=on_critical_condition_change)
    with c_type2:
        con_type = st.radio("계약 방식", ["비교견적", "수의계약"], horizontal=True, key="saved_con_type", on_change=on_critical_condition_change)

    col_in1, col_in2 = st.columns(2)
    with col_in1:
        with st.container(border=True):
            editable_headline("extractor_sec2", "📝 2. ERP 텍스트 복사", level="h4", prefix="2번 ERP 텍스트 복사 소제목")
            
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
            editable_headline("extractor_sec3", "📎 3. 파일 직접 업로드", level="h4", prefix="3번 파일 업로드 소제목")
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
    btn_analyze_txt = APP_CONFIG.get("extractor_btn_analyze", "🚀 전체 데이터 통합 분석 시작")
    with col_btn1: start_all = st.button(btn_analyze_txt, width="stretch", type="primary")
        
    def reset_all_state():
        st.session_state["p_req"] = ""  
        st.session_state["p_fin"] = ""
        st.session_state["p_itm"] = ""
        st.session_state["file_key"] = st.session_state.get("file_key", 0) + 1
        st.session_state["is_saved"] = False
        st.session_state['api_failed'] = False
        st.session_state['ai_errors'] = []
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
                rt, ai_errors = run_ai_analysis(up_files, req_type)
                if rt: 
                    st.session_state[SESSION_AI_KEY] = rt
                    st.session_state['api_failed'] = False
                    st.session_state['ai_errors'] = []
                else: 
                    handle_error("AI", "구글 AI 서버 응답 실패. 수동 입력 모드로 전환합니다.")
                    st.session_state['api_failed'] = True
                    st.session_state['ai_errors'] = ai_errors
        else:
            st.session_state[SESSION_AI_KEY] = {}
            st.session_state['api_failed'] = False
            st.session_state['ai_errors'] = []

    if st.session_state.get(SESSION_REQ_KEY):
        st.divider()
        editable_headline("extractor_step2", "📤 데이터 분석 및 검증 결과", level="h3", prefix="2단계 분석 결과 구역 제목")

        if st.session_state.get('api_failed', False):
            ai_errors = st.session_state.get('ai_errors', [])
            
            has_quota = any(e.get("type") == "QUOTA_EXCEEDED" for e in ai_errors)
            has_parse = any(e.get("type") == "PARSE_ERROR" for e in ai_errors)
            has_safety = any(e.get("type") in ["SAFETY_BLOCKED", "CONTENT_BLOCKED"] for e in ai_errors)
            has_auth = any(e.get("type") == "AUTH_ERROR" for e in ai_errors)
            has_model_nf = any(e.get("type") == "MODEL_NOT_FOUND" for e in ai_errors)
            has_payload = any(e.get("type") == "PAYLOAD_TOO_LARGE" for e in ai_errors)
            
            # 1. 사용자가 바로 원인을 알 수 있는 핵심 진단 배너
            if has_quota:
                st.error(
                    "🚨 **AI 자동 분석 실패 원인: [API 사용량/한도 초과 (429 Quota Exceeded)]**\n\n"
                    "Google Gemini API의 **분당 호출 수(RPM) 또는 일일 무료 토큰 한도**에 도달하여 구글 서버가 분석을 거절했습니다.\n\n"
                    "💡 **해결/대응 방법**:\n"
                    "1. 무료 티어의 경우 약 **1~2분 뒤**에 다시 시도하시면 정상 작동할 수 있습니다.\n"
                    "2. 지금 즉시 처리가 필요하시다면 아래 가이드에 따라 제미나이 웹에서 직접 추출하여 붙여넣으세요."
                )
            elif has_safety:
                st.error(
                    "🛑 **AI 자동 분석 실패 원인: [구글 안전 정책(Safety Filter) 차단]**\n\n"
                    "첨부된 서류 또는 텍스트가 Google AI의 안전 정책(개인정보, 민감정보 등) 기준에 감지되어 분석 생성이 차단되었습니다.\n\n"
                    "💡 **해결/대응 방법**: 서류 내 민감정보(주민등록번호 등)를 가린 뒤 재시도하시거나 아래 수동 모드를 이용하세요."
                )
            elif has_auth:
                st.error(
                    "🔑 **AI 자동 분석 실패 원인: [API 키 인증 실패 (401/403)]**\n\n"
                    ".env 파일의 GEMINI_API_KEY가 올바르지 않거나 만료되었습니다. API 키 설정을 확인해주세요."
                )
            elif has_parse:
                st.warning(
                    "⚠️ **AI 자동 분석 실패 원인: [규격 미준수 응답 / 제미나이 거부 안내문]**\n\n"
                    "제미나이가 응답했으나, 시스템이 요구하는 정형 데이터(JSON) 형태가 아닌 **일반 텍스트나 거부 안내문**으로 답변하여 자동 분석에 실패했습니다.\n\n"
                    "👉 **제미나이가 실제로 뭐라고 답변했는지는 바로 아래 '💬 제미나이 실제 답변 내용'에서 확인하실 수 있습니다.**"
                )
            elif has_model_nf:
                st.error(
                    "🔍 **AI 자동 분석 실패 원인: [지정 모델 미지원 (404 Not Found)]**\n\n"
                    "호출하려는 Gemini 모델을 구글 서버에서 찾을 수 없습니다."
                )
            elif has_payload:
                st.error(
                    "📦 **AI 자동 분석 실패 원인: [파일 용량 또는 토큰 한도 초과]**\n\n"
                    "업로드한 파일 크기나 문서 페이지 수가 너무 많아 1회 요청 제한을 초과했습니다."
                )
            else:
                last_summary = ai_errors[-1].get("summary") if ai_errors else "구글 AI 서버 응답 실패"
                st.error(f"🚨 **AI 자동 검증 실패: {last_summary}**\n\n시스템이 멈추지 않도록 우회(수동) 모드를 가동합니다. 아래 가이드에 따라 제미나이 웹에서 직접 데이터를 추출하여 붙여넣으세요.")

            # 2. 제미나이 실제 답변 원문 (답변이 있는 경우 바로 확인 가능)
            actual_replies = [e["response_text"] for e in ai_errors if e.get("response_text")]
            if actual_replies:
                with st.expander("💬 **제미나이가 실제로 응답한 내용 확인하기 (클릭)**", expanded=True):
                    st.info("💡 제미나이가 서류를 검토한 후 아래와 같이 답변했습니다 (정형 JSON이 아니라 자동 반영되지 못함):")
                    st.text_area("제미나이 답변 원문", value=actual_replies[-1], height=160, disabled=True)

            # 3. 모델별 시도 내역 및 상세 에러 로그 (디버깅용)
            if ai_errors:
                with st.expander("🔍 **단계별 모델 호출 실패 상세 로그 (디버깅용)**", expanded=False):
                    for idx, err in enumerate(ai_errors, 1):
                        st.markdown(f"**[{idx}] {err.get('model')} (시도 {err.get('attempt')})**")
                        st.write(f"- **진단**: {err.get('summary')}")
                        if err.get('error'):
                            st.caption(f"상세 에러 내용: `{str(err.get('error'))[:300]}`")

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
                            st.session_state['ai_errors'] = []
                            st.success("✅ 수동 데이터가 완벽하게 적용되었습니다!")
                            time.sleep(1)
                            st.rerun() 
                        else:
                            st.error("🚨 형식 오류: 중괄호 { } 가 포함된 JSON 결과값을 정확히 붙여넣어주세요.")
                with c_btn2:
                    if st.button("⏭️ AI 검증 포기하고 그냥 넘어가기", width="stretch"):
                        st.session_state[SESSION_AI_KEY] = {}
                        st.session_state['api_failed'] = False
                        st.session_state['ai_errors'] = []
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
        editable_headline("extractor_res_sec_erp", "📝 ERP 기본 정보", level="h4", prefix="ERP 기본 정보 소제목")
        r1_c1, r1_c2, r1_c3 = st.columns(3)
        with r1_c1: display_status_value("🔢 구매요구번호", base_data['req_no'])
        with r1_c2: display_status_value("📌 과제번호", base_data['p_no'])
        with r1_c3: display_status_value("👤 연구책임자 성명", base_data['pi_name'])
        
        r2_c1, r2_c2, r2_c3 = st.columns(3)
        with r2_c1: display_status_value("🏫 연구책임자 학과", base_data['pi_dept'])
        with r2_c2: display_status_value("💰 구매요청액", f"{total_amt:,}원")
        with r2_c3: display_status_value("📧 연구책임자 정보", base_data['pi_formatted'])
        
        st.divider()
        editable_headline("extractor_res_sec_item", "📦 ERP 품목 내역", level="h4", prefix="ERP 품목 내역 소제목")
        if item_list:
            t_data = []
            for i, itm in enumerate(item_list):
                stt = itm['reg_status']
                s_disp = f"🚨 {stt}" if "확인필요" in stt else (f"✅ 예" if stt=="예" else ("➖ 해당없음" if "N/A" in stt else f"⚠️ 아니오"))
                rem_val = itm.get('remark', '')
                t_data.append({
                    "No.": i+1, 
                    "품명": itm['name'], 
                    "단위": itm['unit'], 
                    "등록여부": s_disp, 
                    "물품담당자": itm['contact'],
                    "비고": rem_val if rem_val else "➖"
                })
            st.dataframe(pd.DataFrame(t_data), width="stretch", hide_index=True)

            # 💡 [신규] 비고사항 감지 시 별도 참고용 카드 렌더링
            items_with_remarks = [itm for itm in item_list if itm.get('remark') and itm['remark'].strip() and itm['remark'].strip() not in ['-', '없음', '해당없음']]
            if items_with_remarks:
                with st.container(border=True):
                    st.markdown("##### 📌 [참고용] 물품/용역 비고사항 감지")
                    st.info("💡 ERP 품목 내역의 '비고'란에 기재된 내용입니다. 구매 심사 및 처리 시 참고하세요:")
                    for itm in items_with_remarks:
                        st.markdown(f"- 🔹 **{itm['name']}**: `{itm['remark']}`")
        else:
            st.warning("품목 내역이 없습니다.")
        
        st.divider()
        editable_headline("extractor_res_sec_opinion", "✍️ 결재 의견 (자동 생성)", level="h4", prefix="결재 의견 소제목")
        st.code(generate_opinion(req_type, con_type, total_amt, item_list, raw_fin), language="text")
        
        st.divider()
        
        # 💡 [코드 단축] 데스크탑의 탐색기 기능이 불필요하여 깔끔하게 폴더 이름 텍스트 복사 뷰어만 남김
        editable_headline("extractor_res_sec_vendor", "📁 업체 정보 통합 및 추천 폴더명", level="h4", prefix="업체 정보 소제목")
        
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
        editable_headline("extractor_res_sec_bank", "🏦 통장사본 계좌정보", level="h4", prefix="통장 계좌정보 소제목")
        c_b1, c_b2 = st.columns(2)
        with c_b1:
            t_acc_holder = st.text_input("예금주 (수정 가능)", value=ai.get('예금주', '') if ai else '', key="req_acc_holder")
        with c_b2:
            t_acc_num = st.text_input("계좌번호 (수정 가능)", value=ai.get('계좌번호', '') if ai else '', key="req_acc_num")

        st.divider()
        if st.button("💾 현재 분석 내역을 시스템(DB)에 저장하기", type="primary", width="stretch"):
            item_mgrs = []
            for itm in item_list:
                if itm.get('contact') and "없음" not in itm['contact']:
                    item_mgrs.append(itm['contact'])
            unique_item_mgrs = list(dict.fromkeys(item_mgrs))
            final_item_mgr = " / ".join(unique_item_mgrs) if unique_item_mgrs else "⚠️ 미기재"

            vendor_mgr_info = f"{t_rep} / {t_phn} / {t_eml}"

            # 비고사항 수집 및 통합 정리
            item_remarks = []
            for itm in item_list:
                rem = itm.get('remark', '').strip()
                if rem and rem not in ['-', '없음', '해당없음']:
                    item_remarks.append(f"{itm['name']}: {rem}" if len(item_list) > 1 else rem)
            final_remarks = " / ".join(item_remarks) if item_remarks else ""
            
            raw_to_save = {
                "req_input": raw_req, "fin_input": raw_fin, "itm_input": raw_itm, "ai_res": ai, "req_type": req_type, "con_type": con_type,
                "item_list": item_list, "remarks": final_remarks
            }
            
            meta_to_save = {
                "요청분석일자": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "구매요구번호": base_data['req_no'],
                "과제번호": base_data['p_no'],
                "구매요청액": f"{total_amt:,}원",
                "연구책임자정보": base_data['pi_formatted'], 
                "물품담당자": final_item_mgr,               
                "업체명": t_ven,
                "업체담당자정보": vendor_mgr_info,
                "업체담당자 정보": vendor_mgr_info,
                "사업자번호": t_biz,
                "계좌번호": t_acc_num,
                "예금주": t_acc_holder,
                "비고사항": final_remarks,
                "물품비고": final_remarks,
                "연구책임자_성명": base_data['pi_name'],
                "연구책임자_학과": base_data['pi_dept'],
                "연구책임자_이메일": base_data['pi_email']
            }
            
            save_to_cache(base_data['req_no'], meta_to_save, raw_to_save)
            st.session_state['is_saved'] = True
            st.success(f"✅ 구매요구번호 [{base_data['req_no']}] 건의 핵심 데이터가 로컬 및 구글 시트에 안전하게 저장되었습니다! '분석내용 불러오기' 메뉴에서 확인 가능합니다.")

# ----------------- [2. 구매계약 분석 및 발주서 자동 생성] -----------------
elif menu in [m_con, "🤝 구매계약관리", "🤝 구매계약 분석"]:
    render_top_header("headline_contract", "🤝 구매계약관리")
    editable_headline("contract_desc", "여러 개의 계약 건을 한 번에 입력하고 분석하여, 각각의 워드 발주서를 빠르게 생성할 수 있습니다.", level="info", prefix="상단 안내 설명문")
    
    if 'con_input_count' not in st.session_state:
        st.session_state['con_input_count'] = 2

    raw_cons = []
    raw_memos = []
    for i in range(st.session_state['con_input_count']):
        with st.container(border=True):
            st.markdown(f"##### 📄 [계약 {i+1}] 데이터 입력")
            c_txt1, c_txt2 = st.columns([3, 2])
            with c_txt1:
                txt = st.text_area(f"{i+1}️⃣ 구매계약 화면 본문", height=100, key=f"raw_con_input_{i}", placeholder="rERP 구매계약 화면의 텍스트를 복사해 붙여넣으세요.")
                raw_cons.append(txt)
            with c_txt2:
                memo = st.text_area(f"📌 {i+1}번 노션 메모 (선택)", height=100, key=f"raw_con_memo_{i}", placeholder="예:\n실사용자: 정영권 교수 01033246327 ykjung@sejong.ac.kr\n업체담당자: 김철수 01099998888 cs@vendor.com")
                raw_memos.append(memo)
    
    if st.button("➕ 입력창 1개 더 추가하기", type="secondary"):
        st.session_state['con_input_count'] += 1
        st.rerun()

    st.markdown("<br>", unsafe_allow_html=True) 

    col_btn1, col_btn2 = st.columns(2)
    btn_lbl = APP_CONFIG.get("btn_load_contract", "계약 데이터 불러오기")
    with col_btn1: start_con = st.button(f"📋 {btn_lbl}", width="stretch", type="primary")
    with col_btn2: 
        if st.button("🔄 전체 초기화", width="stretch"): 
            st.session_state.pop('multi_con_data', None)
            st.session_state['con_input_count'] = 2 
            for k in list(st.session_state.keys()):
                if k.startswith("raw_con_input_") or k.startswith("raw_con_memo_") or k.startswith("u_") or k.startswith("v_"):
                    del st.session_state[k]
            st.rerun()

    if start_con:
        valid_data = []
        for idx, txt in enumerate(raw_cons):
            memo_txt = raw_memos[idx] if idx < len(raw_memos) else ""
            if txt.strip() or memo_txt.strip(): 
                cd = extract_contract_info(txt)
                # 별도 메모창에 입력된 내용이 있으면 우선 적용
                if memo_txt.strip():
                    memo_info = parse_smart_contact_memo(memo_txt)
                    if 'smart_memo' not in cd: cd['smart_memo'] = {}
                    for mk, mv in memo_info.items():
                        if mv: cd['smart_memo'][mk] = mv

                valid_data.append(cd)
        
        if valid_data:
            st.session_state['multi_con_data'] = valid_data
            # 세션 스테이트에 스마트 파서 값 세팅
            for i, cd in enumerate(valid_data):
                sm = cd.get('smart_memo', {})
                st.session_state[f"u_nm_{i}"] = sm.get('user_name', '')
                st.session_state[f"u_ph_{i}"] = sm.get('user_phone', '')
                st.session_state[f"u_em_{i}"] = sm.get('user_email', cd.get('email', ''))
                st.session_state[f"v_rep_{i}"] = sm.get('vendor_rep', '')
                st.session_state[f"v_ph_{i}"] = sm.get('vendor_phone', '')
                # 업체담당자 이메일은 부재 시 무조건 공란("") 유지
                st.session_state[f"v_em_{i}"] = sm.get('vendor_email', '')
        else:
            st.warning("입력된 텍스트가 없습니다. 먼저 텍스트를 붙여넣어주세요.")

    if 'multi_con_data' in st.session_state:
        st.divider()
        st.write(f"### 📝 2. 발주요청서 입력 폼 (총 {len(st.session_state['multi_con_data'])}건 추출됨)")
        
        for i, cd in enumerate(st.session_state['multi_con_data']):
            with st.expander(f"📌 [계약 {i+1}] {cd.get('con_name', '건명 미상')}", expanded=True):
                
                # 💡 노션 비고란 실시간 즉시 반영 도구
                with st.expander("⚡ 노션 비고란 메모 즉시 붙여넣기 (실사용자/업체 자동 배분)", expanded=False):
                    quick_memo = st.text_area("노션 비고란 텍스트를 붙여넣고 [적용]을 누르세요", key=f"qmemo_{i}", placeholder="실사용자: 정영권 교수 01033246327 ykjung@sejong.ac.kr\n업체담당자: 김철수 01099998888 cs@vendor.com", height=70)
                    if st.button("✨ 폼에 즉시 반영", key=f"btn_apply_qmemo_{i}"):
                        parsed = parse_smart_contact_memo(quick_memo)
                        if parsed['user_name']: st.session_state[f"u_nm_{i}"] = parsed['user_name']
                        if parsed['user_phone']: st.session_state[f"u_ph_{i}"] = parsed['user_phone']
                        if parsed['user_email']: st.session_state[f"u_em_{i}"] = parsed['user_email']
                        if parsed['vendor_rep']: st.session_state[f"v_rep_{i}"] = parsed['vendor_rep']
                        if parsed['vendor_phone']: st.session_state[f"v_ph_{i}"] = parsed['vendor_phone']
                        # 업체 이메일 없으면 공란으로 갱신
                        st.session_state[f"v_em_{i}"] = parsed.get('vendor_email', '')
                        st.rerun()

                # 💡 사업 구분 (신임교원지원사업 vs 일반 중앙구매) 자동 선택 및 수동 변경
                ptype_options = ["🏢 일반 중앙구매 (산학협력단)", "🏫 신임교원지원사업 (교비)"]
                default_p_idx = 1 if cd.get('project_type') == "신임교원지원사업" else 0
                c_pt1, c_pt2 = st.columns([1, 2])
                with c_pt1:
                    sel_ptype = st.selectbox("사업 구분", ptype_options, index=default_p_idx, key=f"ptype_{i}")
                with c_pt2:
                    if "신임교원" in sel_ptype:
                        st.info("🏫 **신임교원지원사업(교비)** 양식이 적용됩니다. (사업자등록증: 세종대학교)")
                    else:
                        st.caption("🏢 **일반 중앙구매(산학협력단)** 양식이 적용됩니다. (사업자등록증: 산학협력단)")

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

                sm_data = cd.get('smart_memo', {})
                default_u_nm = st.session_state.get(f"u_nm_{i}", sm_data.get('user_name', ''))
                default_u_ph = st.session_state.get(f"u_ph_{i}", sm_data.get('user_phone', ''))
                default_u_em = st.session_state.get(f"u_em_{i}", sm_data.get('user_email', cd.get('email', '')))

                editable_headline("contract_sec_user", "1) 실사용자 정보 (스마트 자동 추출)", level="caption", prefix="실사용자 섹션 소제목", key_suffix=f"{i}")
                c_u1, c_u2, c_u3 = st.columns(3)
                with c_u1: u_nm = st.text_input("실사용자 성명", value=default_u_nm, key=f"u_nm_{i}")
                with c_u2: u_ph = st.text_input("실사용자 연락처", value=default_u_ph, key=f"u_ph_{i}")
                with c_u3: u_em = st.text_input("실사용자 이메일", value=default_u_em, key=f"u_em_{i}")

                default_v_rep = st.session_state.get(f"v_rep_{i}", sm_data.get('vendor_rep', ''))
                default_v_ph = st.session_state.get(f"v_ph_{i}", sm_data.get('vendor_phone', ''))
                # 💡 업체담당자 이메일은 실사용자 이메일(cd['email'])을 덮어쓰지 않고 부재 시 공란으로 둠
                default_v_em = st.session_state.get(f"v_em_{i}", sm_data.get('vendor_email', ''))

                editable_headline("contract_sec_vendor", "2) 협력사(업체) 담당자 정보 (이메일 미기재 시 직접 입력 가능)", level="caption", prefix="협력사 섹션 소제목", key_suffix=f"{i}")
                c_vr1, c_vr2, c_vr3 = st.columns(3)
                with c_vr1: v_rep = st.text_input("업체담당자 성명", value=default_v_rep, key=f"v_rep_{i}")
                with c_vr2: v_ph = st.text_input("업체담당자 연락처", value=default_v_ph, key=f"v_ph_{i}")
                with c_vr3: v_em = st.text_input("업체담당자 이메일", value=default_v_em, key=f"v_em_{i}", placeholder="미기재 시 직접 입력")

                editable_headline("contract_sec_staff", "3) 구매/검수 담당자 및 계약보증금", level="caption", prefix="구매/검수 담당자 섹션 소제목", key_suffix=f"{i}")
                c_b1, c_b2, c_b3 = st.columns(3)
                with c_b1:
                    buyer_options = list(BUYER_STAFF.keys()) + ["직접입력"]
                    extracted_buyer = cd.get('buyer_name', '')
                    default_b_idx = 0
                    for b_idx, b_opt in enumerate(buyer_options):
                        if b_opt != "직접입력":
                            b_st_name = BUYER_STAFF[b_opt].get("name", "")
                            if b_st_name and b_st_name in extracted_buyer:
                                default_b_idx = b_idx
                                break

                    sel_buyer = st.selectbox("구매계약 담당자", buyer_options, index=default_b_idx, key=f"b_sel_{i}")
                    if sel_buyer == "직접입력":
                        b_nm = st.text_input("구매담당자 성명", value=extracted_buyer, key=f"b_nm_custom_{i}")
                        b_ph = st.text_input("구매담당자 연락처", value="", key=f"b_ph_custom_{i}")
                        b_em = st.text_input("구매담당자 이메일", value="", key=f"b_em_custom_{i}")
                    else:
                        staff_info = BUYER_STAFF.get(sel_buyer, {})
                        b_nm = staff_info.get("name", "")
                        b_ph = staff_info.get("phone", "")
                        b_em = staff_info.get("email", "")
                        st.caption(f"📞 {b_ph} | ✉️ {b_em}")

                with c_b2:
                    insp_name_default = INSPECTOR_INFO.get("name", "경제용")
                    i_nm = st.text_input("검수담당자", value=insp_name_default, key=f"i_nm_{i}")
                    st.caption(f"📞 {INSPECTOR_INFO.get('phone', '02-3408-4058')} | ✉️ {INSPECTOR_INFO.get('email', 'rudwpdyd@sejong.ac.kr')}")
                
                amt_val = safe_int(e_amt)
                with c_b3:
                    # 💡 기본 계약보증금: 계약금액의 10% (1000만원 미만도 지급각서에 정확히 반영)
                    auto_deposit = amt_val // 10 if amt_val > 0 else 0
                    default_dep = str(safe_int(cd.get('con_deposit')) or auto_deposit)
                    e_deposit = st.text_input("계약보증금액", value=default_dep, key=f"e_dep_{i}")
                    calc_deposit = safe_int(e_deposit) or auto_deposit
                    if calc_deposit > 0:
                        st.caption(f"￦ {calc_deposit:,} (지급각서 반영)")

                if st.button(f"📄 [{i+1}번] 워드파일 생성 준비", key=f"btn_gen_{i}"):
                    if not e_name or amt_val <= 0 or not e_ven or not u_nm:
                        st.error("🚨 필수 정보(계약건명, 계약금액, 협력사, 실사용자 성명)를 모두 입력해주세요.")
                    else:
                        try:
                            is_new_fac = "신임교원" in sel_ptype
                            new_fac_tpl = "template_new_faculty.docx"

                            if is_new_fac and os.path.exists(new_fac_tpl):
                                tpl_name = new_fac_tpl
                            else:
                                tpl_name = "template_over10mil.docx" if amt_val >= 10000000 else "template_under10mil.docx"
                                if is_new_fac and not os.path.exists(new_fac_tpl):
                                    st.warning(f"⚠️ 신임교원지원사업 전용 양식('{new_fac_tpl}')이 아직 없어 기본 서식({tpl_name})으로 생성되었습니다. 프로젝트 폴더에 {new_fac_tpl}을 추가하시면 자동으로 적용됩니다.")

                            doc = DocxTemplate(tpl_name)
                            
                            final_deposit_str = f"{calc_deposit:,}" if calc_deposit > 0 else ""
                            context = {
                                "구매담당자": b_nm, 
                                "구매담당자_연락처": b_ph, 
                                "구매담당자_메일": b_em,
                                "발주_개요": e_name, 
                                "계약금액": f"{amt_val:,}", 
                                "납품방법": e_deliv, 
                                "납품장소": e_loc,
                                "실사용자": u_nm, 
                                "실사용자_연락처": u_ph, 
                                "실사용자_메일": u_em,
                                "협력사": e_ven,
                                "검수담당자": i_nm, 
                                "검수담당자_연락처": INSPECTOR_INFO.get("phone", "02-3408-4058"), 
                                "검수담당자_메일": INSPECTOR_INFO.get("email", "rudwpdyd@sejong.ac.kr"),
                                "계약보증금액": final_deposit_str
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
    render_top_header("headline_history", "🕒 과거 작업 내역 대시보드")
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
                
                cache_map = {item["구매요구번호"]: item for item in cache_list}
                for _, row in edited_df.iterrows():
                    req_no = str(row["구매요구번호"]) 
                    orig = cache_map.get(req_no, {})
                    changed = any(str(row.get(col, "")) != str(orig.get(col, "")) for col in row.index if col != "구매요구번호")

                    if req_no in cache_db:
                        cache_db[req_no]["meta"]["과제번호"] = row["과제번호"]
                        cache_db[req_no]["meta"]["구매요청액"] = row["구매요청액"]
                        cache_db[req_no]["meta"]["연구책임자정보"] = row["연구책임자정보"]
                        cache_db[req_no]["meta"]["물품담당자"] = row["물품담당자"]
                        cache_db[req_no]["meta"]["업체명"] = row["업체명"]
                        cache_db[req_no]["meta"]["업체담당자정보"] = row["업체담당자정보"]
                        cache_db[req_no]["meta"]["업체담당자 정보"] = row["업체담당자정보"]
                        if "비고사항" in row:
                            cache_db[req_no]["meta"]["비고사항"] = row["비고사항"]
                            cache_db[req_no]["meta"]["물품비고"] = row["비고사항"]

                        # 수정한 행만 구글 시트에 즉시 반영
                        if changed and GSHEETS_URL:
                            try:
                                gsheet_data = dict(cache_db[req_no]["meta"])
                                gsheet_data["Raw_data"] = json.dumps(cache_db[req_no].get("raw", {}), ensure_ascii=False)
                                requests.post(GSHEETS_URL, json={"data": gsheet_data}, timeout=6)
                            except Exception as ge:
                                log(f"구글 시트 수정 반영 에러 ({req_no}): {ge}")
                
                with open(CACHE_FILE, "w", encoding="utf-8") as f:
                    json.dump(cache_db, f, ensure_ascii=False, indent=2)
                    
                st.success("✅ 대시보드 수정사항이 로컬 및 구글 시트에 완벽하게 반영되었습니다!")
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
                st.session_state["nav_target"] = "💰 구매추출기"
                if "open_history_db" in st.session_state:
                    st.session_state["open_history_db"] = False
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

# ----------------- [4. 관리자 설정 화면] -----------------
elif menu == "⚙️ 관리자 설정":
    if not st.session_state.get("is_admin_auth", False):
        st.error("🔒 관리자 전용 화면입니다. 우측 상단의 자물쇠(🔒) 아이콘을 클릭하여 PIN을 입력해주세요.")
        st.stop()

    render_top_header("headline_admin", "⚙️ 관리자 설정")
    st.caption("시스템 명칭, 화면 헤드라인 문구 및 담당자 정보를 손쉽게 변경하고 관리할 수 있습니다.")
    
    # 세션에 임시 설정 버퍼가 없으면 현재 설정으로 로드
    if "admin_cfg" not in st.session_state:
        st.session_state["admin_cfg"] = json.loads(json.dumps(APP_CONFIG))
        
    admin_cfg = st.session_state["admin_cfg"]
    
    tab_screens, tab_buyer, tab_inspector, tab_security = st.tabs([
        "🖥️ 화면별 문구 설정 (위치별 미리보기)", 
        "👥 구매계약 담당자 관리", 
        "🔍 검수 담당자 관리",
        "🔐 관리자 비밀번호(PIN) 변경"
    ])
    
    with tab_screens:
        st.markdown("### 🖥️ 화면 위치별 문구 수정")
        st.info("💡 **화면에서 직접 수정하는 방법**: 우측 상단 자물쇠 옆의 **'✏️ 화면 문구 직접 수정 모드'**를 켜면, 실제 사용하시는 화면 속 각 제목/헤드라인 위치에서 바로 수정하실 수도 있습니다!")
        
        scr_tab1, scr_tab2, scr_tab3 = st.tabs([
            "📱 사이드바 메뉴",
            "💰 1번 화면 (구매추출기)",
            "🤝 2번 화면 (구매계약관리)"
        ])
        
        with scr_tab1:
            st.markdown("##### 📍 [사이드바 영역]")
            with st.container(border=True):
                st.caption("👈 화면 왼쪽 사이드바에 표시되는 명칭들입니다.")
                admin_cfg["sidebar_title"] = st.text_input("1. 사이드바 상단 대제목/로고", value=admin_cfg.get("sidebar_title", "💼 구매 업무 보조"), key="cfg_side_title")
                c_sb1, c_sb2 = st.columns(2)
                with c_sb1:
                    admin_cfg["menu_extractor"] = st.text_input("2. 1번 메뉴 라디오 버튼 명칭", value=admin_cfg.get("menu_extractor", "💰 구매추출기"), key="cfg_menu_ext")
                with c_sb2:
                    admin_cfg["menu_contract"] = st.text_input("3. 2번 메뉴 라디오 버튼 명칭", value=admin_cfg.get("menu_contract", "🤝 구매계약관리"), key="cfg_menu_con")
                admin_cfg["sidebar_status"] = st.text_input("4. 사이드바 하단 가동 상태 문구", value=admin_cfg.get("sidebar_status", "✅ 구매 보조 툴 가동 중"), key="cfg_side_status")
                admin_cfg["app_title"] = st.text_input("5. 웹 브라우저 탭 상단 타이틀", value=admin_cfg.get("app_title", "구매 업무 보조"), key="cfg_app_title")
                
        with scr_tab2:
            st.markdown("##### 📍 [1번 화면: 구매추출기 실제 레이아웃 순서]")
            with st.container(border=True):
                st.markdown("###### 🚩 [상단 영역]")
                admin_cfg["headline_extractor"] = st.text_input("1. 화면 최상단 메인 대제목", value=admin_cfg.get("headline_extractor", "💰 구매추출기"), key="cfg_head_ext")
                
            with st.container(border=True):
                st.markdown("###### 📥 [1단계: 서류 및 정보 입력 구역]")
                admin_cfg["extractor_step1"] = st.text_input("2. 1단계 전체 섹션 제목", value=admin_cfg.get("extractor_step1", "📥 서류 및 정보 입력"), key="cfg_ext_s1")
                
                c_ex1, c_ex2 = st.columns(2)
                with c_ex1:
                    st.caption("👈 좌측 입력 폼")
                    admin_cfg["extractor_sec1"] = st.text_input("3. 기본 정보 선택 소제목", value=admin_cfg.get("extractor_sec1", "📂 1. 기본 정보 선택"), key="cfg_ext_sec1")
                    admin_cfg["extractor_sec2"] = st.text_input("4. ERP 텍스트 복사 소제목", value=admin_cfg.get("extractor_sec2", "📝 2. ERP 텍스트 복사"), key="cfg_ext_sec2")
                with c_ex2:
                    st.caption("👉 우측 파일 첨부 폼")
                    admin_cfg["extractor_sec3"] = st.text_input("5. 파일 직접 업로드 소제목", value=admin_cfg.get("extractor_sec3", "📎 3. 파일 직접 업로드"), key="cfg_ext_sec3")
                    
                admin_cfg["extractor_btn_analyze"] = st.text_input("6. 분석 실행 버튼 문구", value=admin_cfg.get("extractor_btn_analyze", "🚀 전체 데이터 통합 분석 시작"), key="cfg_ext_btn")
                
            with st.container(border=True):
                st.markdown("###### 📤 [2단계: 분석 결과 구역]")
                admin_cfg["extractor_step2"] = st.text_input("7. 결과 섹션 제목", value=admin_cfg.get("extractor_step2", "📤 데이터 분석 및 검증 결과"), key="cfg_ext_s2")
                
                c_res1, c_res2 = st.columns(2)
                with c_res1:
                    admin_cfg["extractor_res_sec_erp"] = st.text_input("8. ERP 기본정보 소제목", value=admin_cfg.get("extractor_res_sec_erp", "📝 ERP 기본 정보"), key="cfg_ext_r_erp")
                    admin_cfg["extractor_res_sec_item"] = st.text_input("9. ERP 품목내역 소제목", value=admin_cfg.get("extractor_res_sec_item", "📦 ERP 품목 내역"), key="cfg_ext_r_itm")
                    admin_cfg["extractor_res_sec_opinion"] = st.text_input("10. 결재의견 소제목", value=admin_cfg.get("extractor_res_sec_opinion", "✍️ 결재 의견 (자동 생성)"), key="cfg_ext_r_opn")
                with c_res2:
                    admin_cfg["extractor_res_sec_vendor"] = st.text_input("11. 업체정보/폴더 소제목", value=admin_cfg.get("extractor_res_sec_vendor", "📁 업체 정보 통합 및 추천 폴더명"), key="cfg_ext_r_ven")
                    admin_cfg["extractor_res_sec_bank"] = st.text_input("12. 계좌정보 소제목", value=admin_cfg.get("extractor_res_sec_bank", "🏦 통장사본 계좌정보"), key="cfg_ext_r_bnk")
                
        with scr_tab3:
            st.markdown("##### 📍 [2번 화면: 구매계약관리 실제 레이아웃 순서]")
            with st.container(border=True):
                st.markdown("###### 🚩 [상단 영역]")
                admin_cfg["headline_contract"] = st.text_input("1. 화면 최상단 메인 대제목", value=admin_cfg.get("headline_contract", "🤝 구매계약관리"), key="cfg_head_con")
                admin_cfg["contract_desc"] = st.text_area("2. 화면 상단 안내 설명문", value=admin_cfg.get("contract_desc", "여러 개의 계약 건을 한 번에 입력하고 분석하여, 각각의 워드 발주서를 빠르게 생성할 수 있습니다."), key="cfg_con_desc", height=70)
                admin_cfg["btn_load_contract"] = st.text_input("3. 계약 데이터 불러오기 버튼 문구", value=admin_cfg.get("btn_load_contract", "계약 데이터 불러오기"), key="cfg_con_btn")
                
            with st.container(border=True):
                st.markdown("###### 📄 [계약 상세 폼 내부 섹션]")
                admin_cfg["contract_sec_user"] = st.text_input("4. 실사용자 정보 소제목", value=admin_cfg.get("contract_sec_user", "1) 실사용자 정보 (스마트 자동 추출)"), key="cfg_con_sec_u")
                admin_cfg["contract_sec_vendor"] = st.text_input("5. 협력사(업체) 정보 소제목", value=admin_cfg.get("contract_sec_vendor", "2) 협력사(업체) 담당자 정보 (이메일 미기재 시 직접 입력 가능)"), key="cfg_con_sec_v")
                admin_cfg["contract_sec_staff"] = st.text_input("6. 구매/검수 담당자 정보 소제목", value=admin_cfg.get("contract_sec_staff", "3) 구매/검수 담당자 및 계약보증금"), key="cfg_con_sec_s")

    with tab_buyer:
        st.markdown("#### 👥 구매계약 담당자 목록")
        st.caption("발주서 작성 시 드롭다운 목록에 자동 표시되는 담당자 명단입니다.")
        
        current_buyers = admin_cfg.get("buyer_staff", {})
        
        staff_keys_to_delete = []
        for b_key, b_info in list(current_buyers.items()):
            with st.container(border=True):
                c_k1, c_k2, c_k3, c_k4 = st.columns([2, 2, 3, 1])
                with c_k1:
                    new_b_name = st.text_input(f"성명 (문서 서명란)", value=b_info.get("name", ""), key=f"cfg_bname_{b_key}")
                with c_k2:
                    new_b_phone = st.text_input(f"연락처", value=b_info.get("phone", ""), key=f"cfg_bphone_{b_key}")
                with c_k3:
                    new_b_email = st.text_input(f"이메일", value=b_info.get("email", ""), key=f"cfg_bemail_{b_key}")
                with c_k4:
                    st.text(" ")
                    if st.button("🗑️ 삭제", key=f"del_b_{b_key}", help=f"{b_key} 삭제"):
                        staff_keys_to_delete.append(b_key)
                
                b_info["name"] = new_b_name.strip()
                b_info["phone"] = new_b_phone.strip()
                b_info["email"] = new_b_email.strip()
        
        for dk in staff_keys_to_delete:
            if dk in current_buyers:
                del current_buyers[dk]
            st.rerun()
            
        st.markdown("---")
        with st.expander("➕ 새 구매담당자 추가", expanded=False):
            c_add1, c_add2 = st.columns(2)
            with c_add1:
                add_label = st.text_input("드롭다운 표시명 (예: 홍길동 사원)", key="add_b_label")
                add_name = st.text_input("실제 성명 (예: 홍길동)", key="add_b_name")
            with c_add2:
                add_phone = st.text_input("연락처 (예: 02-1234-5678)", key="add_b_phone")
                add_email = st.text_input("이메일 (예: hong@sejong.ac.kr)", key="add_b_email")
                
            if st.button("추가하기", key="btn_add_buyer"):
                if add_label.strip() and add_name.strip():
                    current_buyers[add_label.strip()] = {
                        "name": add_name.strip(),
                        "phone": add_phone.strip(),
                        "email": add_email.strip()
                    }
                    st.success(f"'{add_label.strip()}' 담당자가 추가되었습니다.")
                    st.rerun()
                else:
                    st.error("표시명과 실제 성명을 모두 입력해주세요.")

    with tab_inspector:
        st.markdown("#### 🔍 검수 담당자 기본 정보")
        st.caption("발주서 문서에 기본으로 들어갈 검수담당자 정보입니다.")
        insp = admin_cfg.get("inspector_info", {})
        i_name = st.text_input("검수담당자 성명", value=insp.get("name", "경제용"), key="cfg_insp_name")
        i_phone = st.text_input("검수담당자 연락처", value=insp.get("phone", "02-3408-4058"), key="cfg_insp_phone")
        i_email = st.text_input("검수담당자 이메일", value=insp.get("email", "rudwpdyd@sejong.ac.kr"), key="cfg_insp_email")
        
        insp["name"] = i_name.strip()
        insp["phone"] = i_phone.strip()
        insp["email"] = i_email.strip()
        
    with tab_security:
        st.markdown("#### 🔐 관리자 PIN 번호 변경")
        st.caption("우측 상단 자물쇠(🔒) 클릭 시 입력하는 관리자 비밀번호입니다.")
        t_pin = st.text_input("새 관리자 PIN", value=str(admin_cfg.get("admin_pin", "1234")), key="cfg_admin_pin", type="password")
        admin_cfg["admin_pin"] = t_pin.strip() or "1234"
        
    st.divider()
    c_btn1, c_btn2, c_btn3 = st.columns([1, 1, 1])
    with c_btn1:
        if st.button("💾 설정 저장하기", type="primary", use_container_width=True):
            if save_app_config(admin_cfg):
                st.session_state["admin_cfg"] = json.loads(json.dumps(admin_cfg))
                st.success("✅ 설정이 성공적으로 저장되었습니다! 즉시 반영됩니다.")
                time.sleep(0.5)
                st.rerun()
            else:
                st.error("🚨 설정 파일 저장 실패")
    with c_btn2:
        if st.button("🔄 기본값으로 복원", type="secondary", use_container_width=True):
            save_app_config(DEFAULT_APP_CONFIG)
            st.session_state["admin_cfg"] = json.loads(json.dumps(DEFAULT_APP_CONFIG))
            st.info("기본 설정으로 복원되었습니다.")
            time.sleep(0.5)
            st.rerun()
    with c_btn3:
        if st.button("🔒 관리자 로그아웃 (숨김)", type="secondary", use_container_width=True):
            st.session_state["is_admin_auth"] = False
            st.session_state["ui_edit_mode"] = False
            st.session_state["nav_target"] = APP_CONFIG.get("menu_extractor", "💰 구매추출기")
            st.rerun()