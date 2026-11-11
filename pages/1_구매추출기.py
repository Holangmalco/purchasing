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
# 💡 [신규] 국세청 사업자 휴폐업 조회 엔진 (1시간 캐싱으로 반복 호출 방지)
@st.cache_data(ttl=3600, show_spinner=False)
def check_nts_business(biz_no_str, api_key=""):
    clean_b = re.sub(r'[^\d]', '', str(biz_no_str))
    if len(clean_b) != 10: 
        return "⚠️ 번호오류"
    
    key_to_use = api_key or (globals().get("NTS_API_KEY") or "")
    if not key_to_use:
        return "⚠️ API키 누락"
    
    url = f"https://api.odcloud.kr/api/nts-businessman/v1/status?serviceKey={key_to_use}"
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
REQUIRED_FILES = ["template_under10mil.docx", "template_over10mil.docx", "template_new_faculty.docx"]

def get_template_path(filename):
    """
    템플릿 파일의 실제 절대 경로를 탐색하여 반환합니다.
    (프로젝트 루트 폴더, pages 폴더, 현재 작업 디렉토리 순)
    """
    if isinstance(filename, list):
        for f in filename:
            p = get_template_path(f)
            if p: return p
        return None
        
    base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    candidates = [
        os.path.join(base_dir, filename),
        os.path.join(os.path.dirname(os.path.abspath(__file__)), filename),
        os.path.abspath(filename)
    ]
    for c in candidates:
        if os.path.exists(c):
            return c
    return None

def sync_new_faculty_template():
    """
    신임교원용 원본 eformsign 템플릿(«» 메일머지 태그)이 수정되거나 추가되었을 경우,
    python-docx-template 엔진에서 인식 가능한 {{}} 태그로 변환하여 template_new_faculty.docx를 동기화합니다.
    """
    base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    korean_tpl = get_template_path("신임교원용_발주서(eformsign 템플릿)_대학_국가계약법_1천만원 미만.docx")
    target_tpl = os.path.join(base_dir, "template_new_faculty.docx")
    
    if korean_tpl and os.path.exists(korean_tpl):
        if not os.path.exists(target_tpl) or os.path.getmtime(korean_tpl) > os.path.getmtime(target_tpl):
            try:
                import zipfile
                with zipfile.ZipFile(korean_tpl, 'r') as zin:
                    with zipfile.ZipFile(target_tpl, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
                        for item in zin.infolist():
                            data = zin.read(item.filename)
                            if item.filename in ('word/document.xml', 'word/header1.xml', 'word/footer1.xml'):
                                text = data.decode('utf-8')
                                if chr(171) in text or chr(187) in text:
                                    text = text.replace(chr(171), '{{').replace(chr(187), '}}')
                                    data = text.encode('utf-8')
                            zout.writestr(item, data)
            except Exception:
                pass

sync_new_faculty_template()

def get_config_path():
    p1 = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "app_config.json")
    if os.path.exists(p1): return p1
    p2 = os.path.join(os.path.dirname(os.path.abspath(__file__)), "app_config.json")
    if os.path.exists(p2): return p2
    if os.path.exists("app_config.json"): return os.path.abspath("app_config.json")
    return p1

DEFAULT_APP_CONFIG = {
    # 1. 전역 및 사이드바
    "app_title": "구매 업무 보조",
    "sidebar_title": "💼 구매 업무 보조",
    "menu_extractor": "💰 구매추출기",
    "menu_contract": "🤝 구매계약분석",
    "menu_history": "🕒 분석내역 불러오기",
    "admin_pin": "1234",
    # 2. 구매추출기 화면
    "ext_main_title": "💰 구매추출기",
    "ext_step1_title": "📥 정보 입력 및 서류 업로드",
    "ext_sub1_title": "📂 1. 기본 정보 선택",
    "ext_sub2_title": "📝 2. ERP 텍스트 복사",
    "ext_sub3_title": "📎 3. AI OCR",
    "ext_btn_analyze": "🚀 전체 데이터 통합 분석 시작",
    "ext_step2_title": "📤 [STEP 2] 분석 결과",
    "ext_erp_basic_title": "📝 ERP 기본 정보",
    "ext_erp_items_title": "📦 ERP 품목 내역",
    "ext_bank_title": "🏦 통장사본 계좌정보",
    # 3. 구매계약관리 화면
    "con_main_title": "🤝 구매계약분석",
    "con_desc": "여러 개의 계약 건을 한 번에 입력하고 분석하여, 각각의 워드 발주서를 빠르게 생성할 수 있습니다.",
    "con_card_title": "📄 [계약 {no}] 데이터 입력",
    "btn_load_contract": "계약 데이터 불러오기",
    "con_btn_gen": "📄 [{no}번] 워드파일 생성 준비",
    "con_btn_dl": "📥 [{no}번] 완성된 발주요청서 다운로드 (.docx)",
    "work_base_dir": os.path.join(os.path.expanduser("~"), "Desktop", "구매업무"),
    # 4. 담당자 정보
    "buyer_staff": {
        "권순철 사원": {"name": "권순철", "phone": "02-6935-2624", "email": "hagrid0924@sejong.ac.kr"},
        "김민석 사원": {"name": "김민석", "phone": "02-3408-3032", "email": "sejongtwice@sejong.ac.kr"}
    },
    "inspector_info": {
        "name": "경제용",
        "phone": "02-3408-4058",
        "email": "rudwpdyd@sejong.ac.kr"
    },
    # 5. 특수 사업단 소속 교수님 매핑 규칙
    "special_projects": {
        "박재우": {"kw": ["지역혁신중심", "RISE", "라이즈"], "suffix": "(RISE사업단)"},
        "김재호": {"kw": ["사물인터넷"], "suffix": "(사물인터넷혁신융합대학사업단)"},
        "송오영": {"kw": ["SW중심대학"], "suffix": "(SW중심대학사업단)"},
        "이기학": {"kw": ["G-LAMP", "GLAMP", "대학기초연구소"], "suffix": "(G-LAMP사업단)"}
    }
}

def load_app_config():
    c_path = get_config_path()
    merged = DEFAULT_APP_CONFIG.copy()
    if os.path.exists(c_path):
        try:
            with open(c_path, "r", encoding="utf-8") as f:
                data = json.load(f)
                merged.update(data)
                if "headline_extractor" in data and "ext_main_title" not in data:
                    merged["ext_main_title"] = data["headline_extractor"]
                if "headline_contract" in data and "con_main_title" not in data:
                    merged["con_main_title"] = data["headline_contract"]
        except Exception:
            pass
            
    # 로컬/서버 전용 PIN 보존 파일 확인
    pin_file = os.path.join(os.path.dirname(c_path), ".admin_pin.key")
    if os.path.exists(pin_file):
        try:
            with open(pin_file, "r", encoding="utf-8") as pf:
                custom_pin = pf.read().strip()
                if custom_pin:
                    merged["admin_pin"] = custom_pin
        except:
            pass
    return merged

def save_app_config(new_cfg):
    global APP_CONFIG
    c_path = get_config_path()
    try:
        with open(c_path, "w", encoding="utf-8") as f:
            json.dump(new_cfg, f, ensure_ascii=False, indent=2)
            
        if "admin_pin" in new_cfg:
            pin_file = os.path.join(os.path.dirname(c_path), ".admin_pin.key")
            try:
                with open(pin_file, "w", encoding="utf-8") as pf:
                    pf.write(str(new_cfg["admin_pin"]).strip())
            except:
                pass
        APP_CONFIG = load_app_config()
        return True
    except Exception:
        return False

def get_ui_text(key, default_fallback=""):
    return APP_CONFIG.get(key, DEFAULT_APP_CONFIG.get(key, default_fallback))

APP_CONFIG = load_app_config()

try:
    st.set_page_config(page_title=APP_CONFIG.get("app_title", "구매 업무 보조"), page_icon="💼", layout="wide")
except Exception:
    pass

load_dotenv()
dev_env = os.path.join(r"C:\PurchasingExtractor_Dev", ".env")
if os.path.exists(dev_env):
    load_dotenv(dev_env)

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

missing_files = [f for f in REQUIRED_FILES if not get_template_path(f)]
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

ai_client = genai.Client(api_key=API_KEY, http_options={"api_version": "v1beta"})

INSPECTOR_INFO = APP_CONFIG.get("inspector_info", DEFAULT_APP_CONFIG["inspector_info"])
BUYER_STAFF = APP_CONFIG.get("buyer_staff", DEFAULT_APP_CONFIG["buyer_staff"])

# ==========================================
# 2. 공통 유틸리티
# ==========================================
KST = datetime.timezone(datetime.timedelta(hours=9))

def now_kst():
    return datetime.datetime.now(KST)

def today_kst():
    return datetime.datetime.now(KST).date()

def sanitize_win_name(name_str):
    """윈도우 폴더/파일명 금지 문자 제거 및 공백 정돈"""
    cleaned = re.sub(r'[\\/*?:"<>|\r\n\t]', ' ', str(name_str or '')).strip()
    return re.sub(r'\s+', ' ', cleaned)

def build_recommended_folder_name(dept, pi_name, vendor_name, req_no="", date_obj=None):
    """추천 폴더명 생성: (YYYY.MM.DD) 학과 성명 - 업체명 (구매요구번호)"""
    dt = date_obj or now_kst()
    d_str = dt.strftime('%Y.%m.%d')
    c_dept = sanitize_win_name(dept)
    c_pi = sanitize_win_name(pi_name)
    c_ven = sanitize_win_name(vendor_name) or "업체미상"
    c_req = sanitize_win_name(req_no)
    if c_req in ["확인요망", "미검출", "없음"]:
        c_req = ""

    if c_dept and c_pi:
        if c_dept == c_pi:
            left_part = c_pi
        else:
            left_part = f"{c_dept} {c_pi}".strip()
    else:
        left_part = f"{c_dept} {c_pi}".strip()

    if left_part:
        base_folder = f"({d_str}) {left_part} - {c_ven}"
    else:
        base_folder = f"({d_str}) {c_ven}"

    if c_req:
        return f"{base_folder} ({c_req})"
    return base_folder

def resolve_or_create_req_folder(base_dir, req_no="", dept="", pi_name="", vendor_name="", fallback_title=""):
    """
    지정된 상위 폴더(base_dir)에서 구매요구번호(req_no)가 포함된 기존 폴더를 우선 탐색하고,
    없으면 추천 폴더명 규칙으로 새 폴더를 자동 생성한 뒤 (폴더절대경로, 신규생성여부)를 반환합니다.
    """
    target_base = str(base_dir or "").strip() or os.path.join(os.path.expanduser("~"), "Desktop", "구매업무")
    os.makedirs(target_base, exist_ok=True)

    c_req = sanitize_win_name(req_no)
    if c_req in ["확인요망", "미검출", "없음"]:
        c_req = ""

    # 1. 구매요구번호가 있으면 상위 폴더(및 1단계 하위 폴더)에서 해당 번호가 포함된 폴더 탐색
    if c_req and len(c_req) >= 6:
        try:
            entries = sorted(os.listdir(target_base), reverse=True)
            for entry in entries:
                full_p = os.path.join(target_base, entry)
                if os.path.isdir(full_p) and c_req in entry:
                    return full_p, False
            # 1단계 하위 폴더(월별 분류 폴더 등) 내부도 탐색
            for entry in entries:
                sub_p = os.path.join(target_base, entry)
                if os.path.isdir(sub_p):
                    try:
                        for sub_entry in sorted(os.listdir(sub_p), reverse=True):
                            sub_full = os.path.join(sub_p, sub_entry)
                            if os.path.isdir(sub_full) and c_req in sub_entry:
                                return sub_full, False
                    except Exception:
                        pass
        except Exception:
            pass

    # 2. 기존 폴더가 없으면 추천 폴더명 규칙으로 신규 생성
    eff_dept = dept if (dept or pi_name) else fallback_title
    folder_nm = build_recommended_folder_name(eff_dept, pi_name, vendor_name, c_req)
    new_folder_path = os.path.join(target_base, folder_nm)
    os.makedirs(new_folder_path, exist_ok=True)
    return new_folder_path, True

def open_local_folder(folder_path):
    """윈도우 탐색기에서 해당 폴더를 즉시 엽니다."""
    if os.path.exists(folder_path):
        try:
            os.startfile(os.path.abspath(folder_path))
            return True, ""
        except Exception as e:
            return False, str(e)
    return False, "폴더가 존재하지 않습니다."

def log(msg):
    if DEBUG: print(f"[LOG] {msg}")

def safe_int(val):
    try: return int(re.sub(r'[^\d]', '', str(val)))
    except: return 0

def safe_json_load(text):
    if not text or not str(text).strip():
        return None
    raw = str(text).strip()
    
    # 1. 마크다운 코드블록(```json ... ```) 제거
    raw = re.sub(r'^```(?:json)?\s*', '', raw, flags=re.IGNORECASE)
    raw = re.sub(r'\s*```$', '', raw)
    
    # 2. 일반 json.loads 직렬 파싱 시도
    try:
        return json.loads(raw)
    except Exception:
        pass
        
    # 3. 중괄호 { ... } 구간 정규식 추출
    match = re.search(r'\{.*\}', raw, re.DOTALL)
    if match:
        snippet = match.group()
        try:
            return json.loads(snippet)
        except Exception:
            # Trailing comma (예: , } 또는 , ]) 제거
            cleaned = re.sub(r',\s*([\}\]])', r'\1', snippet)
            try:
                return json.loads(cleaned)
            except Exception:
                pass
                
    # 4. 대괄호 배열 [ { ... } ] 형태 추출
    match_arr = re.search(r'\[\s*\{.*\}\s*\]', raw, re.DOTALL)
    if match_arr:
        snippet_arr = match_arr.group()
        try:
            res = json.loads(snippet_arr)
            return res[0] if isinstance(res, list) and len(res) > 0 else res
        except Exception:
            cleaned_arr = re.sub(r',\s*([\}\]])', r'\1', snippet_arr)
            try:
                res = json.loads(cleaned_arr)
                return res[0] if isinstance(res, list) and len(res) > 0 else res
            except Exception:
                pass
                
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
    # 괄호 안 표기((주), (유) 등) 및 특수 법인기호 우선 제거
    s = re.sub(r'\(.*?\)|\[.*?\]|㈜|㈲', '', s) 
    s = re.sub(r'[^\w가-힣]', '', s).lower() 
    # 전체 법인격 명칭만 제거 (낱글자 '주', '유', '등' 제거로 인한 일반 상호 훼손 방지)
    for kw in ['주식회사', '유한회사', '합자회사', '합명회사', '사단법인', '재단법인', '유한책임회사']:
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
            name = str(row.get('성명 및 직위') or row.get('성명') or '').strip()
            dept = str(row.get('소속') or row.get('학과') or '').strip()
            phone = str(row.get('연락처', '')).strip()
            email = str(row.get('이메일', '')).strip()
            if name:
                key = f"{name} ({dept})" if dept else name
                db[key] = {"phone": phone, "email": email}
                pure_name = re.sub(r'(교수|부교수|조교수|사원|과장|대리|연구원|연구교수|박사|팀장|선임|책임)', '', name).strip()
                if pure_name and pure_name not in db:
                    db[pure_name] = {"phone": phone, "email": email}
    except Exception as e: 
        log(f"DB 로드 실패: {e}")
    return db

def lookup_user_contact(raw_name, raw_dept=""):
    """
    USER_DB에서 교직원/연구원을 검색하여 누락된 연락처 및 이메일을 보완합니다.
    """
    if not raw_name or not USER_DB:
        return None
    
    clean_target = clean_name(raw_name).strip()
    clean_dept = clean_name(raw_dept).strip() if raw_dept else ""

    # 1. 완전 일치 또는 포함 일치
    for k, v in USER_DB.items():
        k_clean = clean_name(k)
        if clean_target and (k_clean == clean_target or clean_target in k_clean):
            if clean_dept:
                if clean_dept in k_clean:
                    return v
            else:
                return v

    # 2. 직급 제거 후 성명 2~4글자 매칭
    short_target = re.sub(r'(교수|부교수|조교수|사원|과장|대리|연구원|연구교수|박사|팀장|선임|책임)', '', str(raw_name)).strip()
    if short_target and len(short_target) >= 2:
        for k, v in USER_DB.items():
            if short_target in k:
                return v
                
    return None

def is_internal_sejong_email(email_str):
    """
    유저리스트(USER_DB)에 등록된 이메일 및
    @sejong.ac.kr, @sju.ac.kr, @rcv.sejong.ac.kr 등 우리 측 교내 이메일 여부를 정밀 판별합니다.
    이메일은 절대 업체(협력사) 담당자 이메일로 취급되지 않고 실사용자 이메일로만 인식됩니다.
    """
    if not email_str:
        return False
    em = str(email_str).strip().lower()
    if not em or "@" not in em:
        return False
        
    domain = em.split("@")[-1].strip()
    
    # 1. 세종대 및 산학협력단 / 연구실 도메인 체크 (@sejong.ac.kr, @sju.ac.kr, @rcv.sejong.ac.kr 등)
    internal_domains = [
        "sejong.ac.kr",
        "sju.ac.kr",
        "rcv.sejong.ac.kr"
    ]
    for d in internal_domains:
        if domain == d or domain.endswith("." + d):
            return True
            
    # 2. user_list.xlsx (USER_DB) 및 구매/검수 담당자 목록 이메일 체크
    if 'USER_DB' in globals() and USER_DB:
        for v in USER_DB.values():
            db_em = str(v.get("email", "")).strip().lower()
            if db_em and db_em == em:
                return True
                
    if 'BUYER_STAFF' in globals() and BUYER_STAFF:
        for v in BUYER_STAFF.values():
            b_em = str(v.get("email", "")).strip().lower()
            if b_em and b_em == em:
                return True
                
    if 'INSPECTOR_INFO' in globals() and INSPECTOR_INFO:
        i_em = str(INSPECTOR_INFO.get("email", "")).strip().lower()
        if i_em and i_em == em:
            return True
            
    return False

# ==========================================
# 💡 대학안전관리센터 특수 관리 규칙
# ==========================================
SAFETY_CENTER_INFO = {
    "pi_name": "김수동 선생님",
    "pi_dept": "대학안전관리센터",
    "pi_phone": "02-3408-4081",
    "pi_email": "sudong_kim@sejong.ac.kr",
    "pi_formatted": "김수동 선생님 / 02-3408-4081 / sudong_kim@sejong.ac.kr"
}

def is_safety_center_case(*texts):
    """과제담당자 또는 관리기관이 대학안전관리센터인 건인지 판별"""
    combined = " ".join([str(t or "") for t in texts])
    if not combined.strip():
        return False
    clean = re.sub(r'\s+', '', combined)
    # 1. '대학안전관리센터' 키워드 직접 포함
    if "대학안전관리센터" in clean or "대학안전관리" in clean:
        return True
    # 2. '과제담당자' 또는 '관리기관' 근처에 '안전관리' 또는 '김수동'이 있는 경우
    if ("과제담당자" in clean or "관리기관" in clean) and ("안전관리" in clean or "김수동" in clean):
        return True
    return False

def check_special_project_affiliation(pi_name, pi_dept, context_text=""):
    """특수 사업단 소속 교수님의 학과 뒤에 (사업단명)을 자동 부착합니다."""
    if not pi_name or not pi_dept:
        return pi_dept
    
    spec_map = APP_CONFIG.get("special_projects", DEFAULT_APP_CONFIG.get("special_projects", {}))
    cfg_rule = spec_map.get(pi_name.strip())
    if cfg_rule:
        suffix = cfg_rule.get("suffix", "")
        if suffix and suffix not in pi_dept:
            kws = cfg_rule.get("kw", [])
            clean_ctx = str(context_text).upper().replace(" ", "")
            if any(str(kw).upper() in clean_ctx for kw in kws):
                return f"{pi_dept}{suffix}"
    return pi_dept

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
import base64 as _b64
CACHE_FILE = "local_cache.json"
_DEFAULT_GS_B64 = "aHR0cHM6Ly9zY3JpcHQuZ29vZ2xlLmNvbS9tYWNyb3Mvcy9BS2Z5Y2J4M3FLaU91Qzd0MEo1ZzdZSldXY2NYelFUbGQ3MjhXUlBxYmN5ZjdqQkQ5VkhGWk5ZdjBYaHFxR3VnLS1pbWdINlUvZXhlYw=="
GSHEETS_URL = get_secret("GSHEETS_URL", ["GOOGLE_SHEETS_URL", "APPS_SCRIPT_URL"]) or _b64.b64decode(_DEFAULT_GS_B64).decode("utf-8")

def clear_db_caches():
    """DB 저장/삭제/수정 직후 메모리 캐시를 즉시 비워 최신 상태를 반영합니다."""
    try:
        fetch_all_db_records.clear()
    except Exception:
        pass
    try:
        get_cache_list.clear()
    except Exception:
        pass

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

    clear_db_caches()

def delete_from_cache(req_no):
    req_no_str = str(req_no).strip()
    # 1. 로컬 캐시 삭제
    if os.path.exists(CACHE_FILE):
        try:
            with open(CACHE_FILE, "r", encoding="utf-8") as f:
                cache = json.load(f)
            if req_no_str in cache:
                del cache[req_no_str]
                with open(CACHE_FILE, "w", encoding="utf-8") as f:
                    json.dump(cache, f, ensure_ascii=False, indent=2)
        except Exception as e:
            log(f"로컬 캐시 삭제 오류: {e}")

    # 2. 구글 시트 영구 저장소 삭제 요청
    if GSHEETS_URL:
        try:
            res = requests.post(GSHEETS_URL, json={"action": "delete", "data": {"구매요구번호": req_no_str}}, timeout=8)
            if res.status_code == 200:
                log(f"구글 시트 삭제 요청 성공 ({req_no_str})")
            else:
                log(f"구글 시트 삭제 요청 응답 ({res.status_code}): {res.text}")
        except Exception as e:
            log(f"구글 시트 삭제 API 연동 실패: {e}")

    clear_db_caches()

@st.cache_data(ttl=60, show_spinner=False)
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

@st.cache_data(ttl=60, show_spinner=False)
def fetch_all_db_records():
    records = []
    # 1. 구글 시트 영구 저장소 조회
    if GSHEETS_URL:
        try:
            res = requests.get(GSHEETS_URL, timeout=6)
            if res.status_code == 200:
                rows = res.json()
                if isinstance(rows, list):
                    for r in rows:
                        req_no = str(r.get("구매요구번호", "")).strip()
                        if req_no and not req_no.startswith("TEST"):
                            records.append(r)
        except Exception as e:
            log(f"구글 시트 레코드 조회 실패: {e}")
            
    # 2. 로컬 캐시 파일 조회 및 보완
    if os.path.exists(CACHE_FILE):
        try:
            with open(CACHE_FILE, "r", encoding="utf-8") as f:
                local_data = json.load(f)
            existing_reqs = {str(r.get("구매요구번호", "")).strip() for r in records}
            for req_no, v in local_data.items():
                req_no_str = str(req_no).strip()
                if req_no_str not in existing_reqs and not req_no_str.startswith("TEST"):
                    meta = v.get("meta", {})
                    meta["구매요구번호"] = req_no_str
                    records.append(meta)
        except Exception as e:
            log(f"로컬 캐시 레코드 조회 실패: {e}")
            
    return records

def check_split_procurement_risk(project_no, vendor_name, vendor_biz, current_req_no, current_amt):
    """
    동일 과제번호 기준 1개월(31일) 이내 동일 업체와의 기존 구매요청 건이 있을 경우 사전 확인 문구 반환
    """
    if not project_no or not str(project_no).strip() or "확인" in str(project_no):
        return [], 0, []
        
    clean_curr_pno = re.sub(r'[\s\-_]', '', str(project_no)).upper()
    clean_curr_biz = re.sub(r'[^\d]', '', str(vendor_biz or ''))
    clean_curr_ven = clean_name(str(vendor_name or ''))
    curr_req_str = str(current_req_no or '').strip()
    
    if not clean_curr_biz and not clean_curr_ven:
        return [], 0, []
        
    records = fetch_all_db_records()
    today = today_kst()
    matching_records = []
    seen_reqs = set()
    
    for r in records:
        r_req = str(r.get("구매요구번호", "")).strip()
        if not r_req or r_req.startswith("TEST"):
            continue
        if curr_req_str and r_req == curr_req_str:
            continue
        if r_req in seen_reqs:
            continue
            
        r_pno = str(r.get("과제번호", "")).strip()
        clean_r_pno = re.sub(r'[\s\-_]', '', r_pno).upper()
        if not clean_r_pno or clean_r_pno != clean_curr_pno:
            continue
            
        r_biz = re.sub(r'[^\d]', '', str(r.get("사업자번호", "")))
        r_ven = clean_name(str(r.get("업체명", "")))
        
        is_same_vendor = False
        if clean_curr_biz and len(clean_curr_biz) >= 10 and r_biz and len(r_biz) >= 10:
            if clean_curr_biz == r_biz:
                is_same_vendor = True
        if not is_same_vendor and clean_curr_ven and r_ven and len(clean_curr_ven) >= 2 and len(r_ven) >= 2:
            if clean_curr_ven == r_ven or clean_curr_ven in r_ven or r_ven in clean_curr_ven:
                is_same_vendor = True
            
        if not is_same_vendor:
            continue
            
        # DB 등록일 기준 1개월(31일) 이내 여부 체크
        d_str = str(r.get("요청분석일자", ""))
        m_date = re.search(r'(\d{4})[-./](\d{2})[-./](\d{2})', d_str)
        if not m_date:
            continue
        try:
            r_date = datetime.date(int(m_date.group(1)), int(m_date.group(2)), int(m_date.group(3)))
        except:
            continue
            
        diff_days = (today - r_date).days
        if -2 <= diff_days <= 31:
            raw_nums = re.sub(r'[^\d]', '', str(r.get("구매요청액", "")))
            amt_num = int(raw_nums) if raw_nums else 0
            matching_records.append({
                "req_no": r_req,
                "date": f"{m_date.group(1)}-{m_date.group(2)}-{m_date.group(3)}",
                "amt": amt_num,
                "ven": r.get("업체명", "")
            })
            seen_reqs.add(r_req)
            
    if not matching_records:
        return [], 0, []
        
    # 날짜순 정렬
    matching_records.sort(key=lambda x: x.get('date', ''))
    past_sum = sum(m["amt"] for m in matching_records)
    combined_total = past_sum + int(current_amt or 0)
    
    # 사용자 지정 포맷: 구매요구번호 XXXX건 동일과제, 동일업체 / 총 구매요청액 XXXX원
    # 과거건과 현재 요청건 금액을 모두 합산한 누적 총액(combined_total)으로 기재
    req_nos_str = ", ".join([str(m["req_no"]) for m in matching_records])
    warn_msg = f"구매요구번호 {req_nos_str}건 동일과제, 동일업체 / 총 구매요청액 {combined_total:,}원"
    
    return [warn_msg], combined_total, matching_records

# ==========================================
# 3. 데이터 파싱 엔진
# ==========================================
def extract_base_info(text, extra_text=""):
    info = {
        'p_no': "⚠️ 확인 불가", 'req_no': "⚠️ 확인 불가", 'pi_name': "", 
        'pi_dept': "", 'pi_phone': "", 'pi_formatted': "", 'prj_end': "", 'pi_email': ""
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
            info['pi_phone'] = format_phone(parts[-2])
            info['pi_formatted'] = f"{parts[0]} / {info['pi_phone']} / {parts[-1]}"
            
            clean_text = text.replace(" ", "")
            prj_context_match = re.search(r'과제성격(.*?)(연구책임자정보|연차연구기간)', clean_text, flags=re.DOTALL)
            prj_context = prj_context_match.group(1) if prj_context_match else clean_text
            info['pi_dept'] = check_special_project_affiliation(info['pi_name'], info['pi_dept'], prj_context)
        else:
            info['pi_formatted'] = pi_match.group(1).strip()
            
    # 💡 [특수 규칙] 과제담당자가 대학안전관리센터인 건은 문서 내 연구책임자 기재와 상관없이 강제 오버라이드
    if is_safety_center_case(text, extra_text):
        info['pi_name'] = SAFETY_CENTER_INFO['pi_name']
        info['pi_dept'] = SAFETY_CENTER_INFO['pi_dept']
        info['pi_phone'] = SAFETY_CENTER_INFO['pi_phone']
        info['pi_email'] = SAFETY_CENTER_INFO['pi_email']
        info['pi_formatted'] = SAFETY_CENTER_INFO['pi_formatted']
            
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

def check_bipum_qualify(name, unit_price, erp_cat=""):
    clean_name = name.replace(" ", "").lower()
    clean_erp = erp_cat.replace(" ", "").lower()
    
    # 💡 '시약대', '시약장' 등은 소모품(시약)이 아니라 실험실 집기/비품이므로 소모품 판별 시 예외 처리
    name_for_consumable = clean_name
    for safe_kw in ['시약대', '시약장', '시약보관', '시약선반', '시약함']:
        name_for_consumable = name_for_consumable.replace(safe_kw, "")

    consumable_kws = ['시약', '용액', '항체', '키트', 'kit', '튜브', '팁', '케이블', '부품', '커버', '케이스', '필터', '토너', '잉크', '용지', '리필', '건전지', '배터리', '소모품', '소모']
    fixture_kws = [
        '책상', '의자', '파티션', '캐비닛', '수납장', '테이블', '책장', '칠판', '소파', '침대', 
        '블라인드', '집기', '서랍장', '옷장', '진열대', '선반', '선반대', '탁자', '데스크',
        '실험대', '시약대', '시약장', '작업대', '싱크대', '씽크대', '세척대'
    ]
    machine_kws = [
        '서버', 'pc', '데스크탑', '노트북', '모니터', '분석기', '현미경', '측정기', '펌프', '레이저', 
        '냉장고', '카메라', '프린터', '라우터', '스위치', '오실로스코프', '센서', 'sensor', 
        '장비', '기기', '기구', '시스템', 'system', 'tv', '텔레비전', '스탠드', 
        '소프트웨어', 'sw', 's/w', '라이선스', '라이센스', '구독',
        '후드', '퓸후드', '흄후드', '클린벤치', '인큐베이터', '챔버'
    ]
    
    is_fixture = any(kw in clean_name or kw in clean_erp for kw in fixture_kws)
    is_machine = any(kw in clean_name or kw in clean_erp for kw in machine_kws)
    
    # 1. 집기비품 15만원 이상 -> 비품 자격
    if is_fixture and unit_price >= 150000:
        return True
    # 2. 기계기구 30만원 이상 -> 비품 자격
    if is_machine and unit_price >= 300000:
        return True
    # 3. 소모품 명시 키워드가 아니면서 단가 30만원 이상인 경우 -> 비품 자격
    if unit_price >= 300000 and not any(kw in name_for_consumable for kw in consumable_kws):
        return True
    return False

def guess_category_by_keyword(name, price, unit_price=0, reg_status=""):
    up = unit_price if unit_price > 0 else price
    is_qualify = check_bipum_qualify(name, up)
    if is_qualify:
        if "아니오" in reg_status and "구성품" in name:
            return "구성품"
        return "비품"
    return "소모품"

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

    # 💡 [Part 1 - 5번] 연락처나 이메일이 누락된 경우 USER_DB에서 자동 보완
    if clean_name and (not phone or not email):
        u_info = lookup_user_contact(clean_name)
        if u_info:
            if not phone and u_info.get("phone"):
                phone = u_info["phone"]
            if not email and u_info.get("email"):
                email = u_info["email"]

    if clean_name and phone and email: return f"{clean_name} / {phone} / {email}"
    elif clean_name and phone: return f"{clean_name} / {phone}"
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

def generate_opinion(req_type, contract_type, amount, item_list, raw_finance_text="", raw_req_text="", ai_res=None, project_no="", vendor_name="", vendor_bizno="", current_req_no=""): 
    rules_triggered = []
    
    # 1. 간접비건 확인 (재원내역 또는 요청서 텍스트)
    if "간접비" in raw_finance_text or "간접비" in raw_req_text:
        rules_triggered.append("간접비건")
        
    is_goods = "물품" in req_type
    is_service = "용역" in req_type or "공사" in req_type
    
    # 2. 단위 수정 확인
    if is_goods and any(i.get('unit', '').lower() not in ['', 'u'] for i in item_list):
        rules_triggered.append("단위 'u'로 수정 필요")
    if is_service and any(i.get('unit', '') not in ['', '식'] for i in item_list):
        rules_triggered.append("단위 '식'으로 수정 필요")

    # 3. 물품분류구분 판정 및 규칙
    has_bipum = False
    has_guseong = False
    
    if is_service:
        final_category = "공사" if "공사" in req_type else "용역"
    elif not item_list:
        final_category = "소모품"
    else:
        for itm in item_list:
            erp_cat = itm.get('category', '공란').strip()
            reg_status = itm.get('reg_status', '')
            unit_price = itm.get('unit_price', 0)
            price = itm.get('price', 0)
            up = unit_price if unit_price > 0 else price
            name = itm.get('name', '')
            
            is_bipum = check_bipum_qualify(name, up, erp_cat)
            if is_bipum:
                if "아니오" in reg_status and "예" not in reg_status and ("구성품" in erp_cat or "구성품" in name):
                    has_guseong = True
                else:
                    has_bipum = True
            
            # ERP 상 물품분류구분이 공란인 경우
            if erp_cat in ['', '공란']:
                rules_triggered.append("물품분류구분 입력필요")
            else:
                # ERP 상 분류와 판정 기준 불일치 시
                if "소모품" in erp_cat and is_bipum:
                    rules_triggered.append("물품분류구분 변경필요")
                elif ("비품" in erp_cat or "기계기구" in erp_cat or "집기" in erp_cat) and not is_bipum:
                    rules_triggered.append("물품분류구분 변경필요")
            
            # 물품등록여부 확인/변경
            if not is_bipum and "예" in reg_status:
                rules_triggered.append("물품등록여부 변경필요")
            elif is_bipum and "아니오" in reg_status and "구성품" not in erp_cat:
                rules_triggered.append("물품등록여부 확인필요")

        if has_bipum:
            final_category = "비품"
        elif has_guseong:
            final_category = "구성품"
        else:
            final_category = "소모품"

    # 4. 업체담당자 정보 확인
    if ai_res:
        ai_eml = ai_res.get('이메일', '')
        ai_phn = ai_res.get('연락처', '')
        is_bad_eml = not ai_eml or ai_eml in ["미검출", "없음"] or is_internal_sejong_email(ai_eml)
        is_bad_phn = not ai_phn or ai_phn in ["미검출", "없음"]
        if is_bad_eml or is_bad_phn:
            rules_triggered.append("업체담당자 정보 확인필요")

    # 5. 동일과제/동일업체 1개월 내 구매요청 이력 확인
    p_no_to_check = project_no
    curr_req_to_check = current_req_no
    if (not p_no_to_check or "확인" in str(p_no_to_check) or not curr_req_to_check) and raw_req_text:
        b_info = extract_base_info(raw_req_text)
        if b_info:
            if not p_no_to_check or "확인" in str(p_no_to_check):
                p_no_to_check = b_info.get('p_no', '')
            if not curr_req_to_check:
                c_req = b_info.get('req_no', '')
                if c_req and "확인" not in c_req:
                    curr_req_to_check = c_req
        
    v_name_to_check = vendor_name or (ai_res.get('사업자등록증_상호명') or ai_res.get('업체명', '') if ai_res else '')
    v_biz_to_check = vendor_bizno or (ai_res.get('사업자등록증_등록번호') or ai_res.get('사업자번호', '') if ai_res else '')
    
    dup_warns, _, _ = check_split_procurement_risk(p_no_to_check, v_name_to_check, v_biz_to_check, curr_req_to_check, amount)
    if dup_warns:
        if isinstance(dup_warns, list):
            rules_triggered.extend(dup_warns)
        else:
            rules_triggered.append(dup_warns)

    unique_rules = list(dict.fromkeys(rules_triggered))
    issues_text = " - 특이사항 없음" if not unique_rules else "\n".join([f" - {r}" for r in unique_rules])
    
    return f"1. {contract_type}에 의한 구매건\n - 물품분류구분 : {final_category}\n\n2. 확인사항\n{issues_text}"

# ==========================================
# 4. AI 서류 검증 엔진
# ==========================================
def run_ai_analysis(files, req_type):
    today_str = now_kst().strftime('%Y-%m-%d')
    
    extra_instruction = ""
    if req_type == "🏗️ 공사":
        extra_instruction = "🚨 이 건은 '공사' 건입니다. 첨부 서류 중 '시설공사 확인서' 또는 '장비설치 확인서'가 반드시 포함되어야 하므로 이를 꼼꼼히 확인하세요.\n"
    
    contents = [
        f"🚨 [기준일] 오늘 날짜는 {today_str}입니다. 만료일자 계산 시 참고하세요.\n",
        f"{extra_instruction}",
        """아래 지침에 따라 첨부된 문서들을 분석하고, 반드시 순수한 JSON 형식으로만 응답해. 절대 마크다운(```json)이나 다른 설명 텍스트를 넣지 마.

■ [문서 자동 판별 및 추출 원칙 (파일명 무관)]:
1. 첨부 파일명이 'ㅅㄷ', 'ㅌㅅ', '1', '스캔' 등 임의의 이름이더라도, 파일 본문의 양식/텍스트를 직접 보고 문서 역할을 구분할 것:
   - [사업자등록증]: 국세청 양식, '등록번호', '법인명(상호)', '대표자', '개업연월일', '사업장소재지' 등이 적힌 서식.
   - [통장사본]: 은행명, 예금주, 계좌번호가 적힌 통장 이미지/사본. (1개 PDF 내 여러 페이지에 사업자등록증과 통장사본이 함께 있는 경우 각 페이지를 개별 판별)
   - [견적서/비교견적서]: 공급자 정보, 품목, 단가, 공급가액, 세액, 총금액이 적힌 견적 문서.
   - [기안문/품의서/결재문서]: 대학/기관 내부 기안문, 시행일자, 결재선 등이 있는 문서는 내부 행정서류이므로 견적 비교 대상에서 제외할 것.
2. [핵심] 첨부 문서 중 '사업자등록증'과 '통장사본' 양식이 단 1장이라도 발견되면, 견적서 최저가 선정 여부와 관계없이 해당 서류 본문의 정보(상호, 사업자번호, 대표자명, 예금주, 계좌번호)를 무조건 독립 추출하여 아래 JSON 필드에 정확히 채울 것.

{
  "사업자등록증_상호명": "첨부된 사업자등록증 상의 '법인명(단체명)' 또는 '상호' (사업자등록증이 없으면 견적서 상호)",
  "사업자등록증_등록번호": "첨부된 사업자등록증 상의 사업자등록번호 (반드시 000-00-00000 형식)",
  "사업자등록증_대표자명": "첨부된 사업자등록증 상의 대표자 성명",
  "통장사본_예금주": "첨부된 통장사본 상의 예금주명",
  "통장사본_계좌번호": "첨부된 통장사본 상의 계좌번호",
  "통장사본_은행명": "첨부된 통장사본 상의 은행명",
  "업체명": "최종 선정된 업체명 (사업자등록증 상호 우선)",
  "사업자번호": "최종 선정 업체의 사업자번호 (사업자등록증 번호 우선, 000-00-00000 형식)",
  "대표자명": "사업자등록증 상의 대표자 성명",
  "담당자": "최저가 업체의 견적 담당자명",
  "연락처": "최저가 업체 담당자의 010 전화번호 (없으면 대표번호)",
  "이메일": "최저가 업체의 이메일",
  "예금주": "통장사본의 예금주명",
  "계좌번호": "통장사본의 계좌번호",
  "비교견적리스트": [ {"업체명": "A업체", "사업자번호": "123-45-67890", "제출금액": 10000}, {"업체명": "B업체", "사업자번호": "987-65-43210", "제출금액": 12000} ],
  "최저가판단과정": "위 비교견적리스트를 바탕으로 가장 금액이 낮은 업체를 최종 선정하는 과정을 짧게 작성",
  "수신자명확인": "첨부된 모든 견적서(비교견적서 포함)의 수신자가 모두 '세종대학교 산학협력단' 명의가 맞으면 '예', 하나라도 누락되거나 다르면 '아니오(사유)'",
  "만료일자": "견적서 본문에 유효기간 만료일(또는 '견적일로부터 N일' 등)이 명시적으로 적혀 있는 경우에만 계산하여 YYYY-MM-DD 형식으로 기재. 본문에 유효기간이 전혀 적혀있지 않거나 유효기간 칸이 비어있는 경우 절대로 30일을 임의로 더하지 말고 반드시 '무기한'으로 기재할 것",
  "동일규격확인": "복수 견적서 간 동일 물품이면 '예', 다르면 '아니오(사유)', 1장뿐이면 '단일견적'",
  "사업자일치여부": "최저가 견적서와 사업자등록증 상의 정보가 일치하면 '예', 다르면 '아니오(사유)'",
  "총금액": "최종 선정된 업체의 견적 총액 (반드시 부가가치세/VAT 포함된 최종 합계 금액을 숫자로만 기재)",
  "공사확인서유무": "첨부 파일 중 '시설공사' 또는 '장비설치 확인서' 문서가 있으면 '예', 없으면 '아니오', 비슷한 문서가 있으면 '확인요망'",
  "품목리스트": [ {"품명": "품명 텍스트", "수량": 1, "단가": 1000, "금액": 1000} ]
}
"""
    ]

    for idx, f in enumerate(files, 1):
        if "excel" in f.type or "spreadsheet" in f.type:
            try: contents.append(f"문서 {idx} (엑셀 파일명: {f.name}):\n{pd.read_excel(f).to_csv(index=False)}")
            except: pass
        else:
            contents.append(f"문서 {idx} (파일명: {f.name})")
            contents.append(types.Part.from_bytes(data=f.getvalue(), mime_type=f.type))

    gen_config = types.GenerateContentConfig(response_mime_type="application/json")
            
    fallback_chain = [
        ("gemini-2.5-flash", 2)
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
        'con_name': re.search(r'계약건명\s*(.*?)\s*(?:납품기간|계약금액|총금액|납품장소|$)', clean_text),
        'con_amount': re.search(r'계약금액\s*([\d,]+)', clean_text),
        'con_deposit': re.search(r'계약보증금\s*([\d,]+)', clean_text),
        'vendor_name': re.search(r'계약업체명\s*(.*?)\s*(?:사업자번호|대표자|전화번호|$)', clean_text),
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

    # 💡 연구책임자정보 직접 추출
    pi_match = re.search(r'연구책임자정보\s*([^/\s]+)\s*/\s*([^/\s]+)\s*/\s*([^/\s]+)\s*/\s*([0-9\-]+)\s*/\s*([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})', clean_text)
    if pi_match:
        result['pi_name'] = pi_match.group(1).strip()
        result['pi_position'] = pi_match.group(2).strip()
        raw_dept = pi_match.group(3).strip()
        result['pi_dept'] = check_special_project_affiliation(result['pi_name'], raw_dept, clean_text)
        
        raw_ph = pi_match.group(4).strip()
        dig = re.sub(r'[^\d]', '', raw_ph)
        if len(dig) == 11 and dig.startswith('01'):
            result['pi_phone'] = f"{dig[:3]}-{dig[3:7]}-{dig[7:]}"
        elif len(dig) == 10 and dig.startswith('01'):
            result['pi_phone'] = f"{dig[:3]}-{dig[3:6]}-{dig[6:]}"
        else:
            result['pi_phone'] = raw_ph
        result['pi_email'] = pi_match.group(5).strip()

    # 💡 [특수 규칙] 과제담당자가 대학안전관리센터인 건은 문서 내 연구책임자 기재와 상관없이 강제 오버라이드
    if is_safety_center_case(clean_text, text):
        result['pi_name'] = SAFETY_CENTER_INFO['pi_name']
        result['pi_dept'] = SAFETY_CENTER_INFO['pi_dept']
        result['pi_phone'] = SAFETY_CENTER_INFO['pi_phone']
        result['pi_email'] = SAFETY_CENTER_INFO['pi_email']
        result['pi_position'] = ""
    
    # 💡 납품장소 파싱 개선: B101호, 101-1호 등 방번호 원형 보존 및 슬래시 유연화
    loc_match = re.search(r'납품장소\s*(.*?)\s*(?:계약금액|총금액|$)', clean_text)
    result['location'] = ""
    if loc_match:
        loc_str = loc_match.group(1).strip()
        if '/' in loc_str:
            parts = [p.strip() for p in loc_str.split('/') if p.strip()]
            if len(parts) >= 2:
                bldg = parts[0]
                room_raw = parts[-1]
                room_match = re.search(r'([A-Za-z0-9\-가-힣]+(?:호|실|층)?)', room_raw)
                if room_match:
                    r_str = room_match.group(1).strip()
                    if not r_str.endswith(('호', '실', '층')):
                        r_str += "호"
                    result['location'] = f"{bldg} {r_str}"
                else:
                    result['location'] = f"{bldg} {room_raw}"
            else:
                result['location'] = loc_str
        else:
            result['location'] = loc_str

    # 본문 내 실사용자/업체 메모 패턴이 직접 포함된 경우 함께 추출
    result['smart_memo'] = parse_smart_contact_memo(text)
    
    # 본문에서 오인 추출된 과도하게 긴 vendor_rep 정제
    cur_vrep = result['smart_memo'].get('vendor_rep', '')
    if cur_vrep and (len(cur_vrep) > 20 or any(bad in cur_vrep for bad in ['계약업체정보', '업체', '법인번호', '사업자번호', '주소'])):
        result['smart_memo']['vendor_rep'] = ''

    # 💡 [핵심 가드] 세종대/USER_DB 이메일은 업체 메일 칸에 절대 남지 않도록 보증
    if is_internal_sejong_email(result['smart_memo'].get('vendor_email')):
        if not result['smart_memo'].get('user_email'):
            result['smart_memo']['user_email'] = result['smart_memo']['vendor_email']
        result['smart_memo']['vendor_email'] = ''
    
    # 💡 사용자가 직접 입력한 메모/정보가 최우선입니다.
    # 사용자가 입력한 실사용자 정보가 없을 때만(공란일 때만) 계약관리의 연구책임자정보를 fallback으로 반영합니다.
    # 단, 과제담당자가 대학안전관리센터인 건은 문서 내 연구책임자와 상관없이 무조건 김수동 선생님으로 강제 설정합니다.
    if is_safety_center_case(clean_text, text):
        result['smart_memo']['user_name'] = SAFETY_CENTER_INFO['pi_name']
        result['smart_memo']['user_phone'] = SAFETY_CENTER_INFO['pi_phone']
        result['smart_memo']['user_email'] = SAFETY_CENTER_INFO['pi_email']
    elif not result['smart_memo'].get('user_name') and result.get('pi_name'):
        pos = result.get('pi_position', '').strip()
        pos_str = f" {pos}" if pos and not result['pi_name'].endswith(pos) else ""
        result['smart_memo']['user_name'] = f"{result['pi_name']}{pos_str}".strip()
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
    노션 비고란, rERP 복사문 등에서 실사용자와 업체(협력사) 담당자 정보를 정밀하게 추출합니다.
    줄바꿈(엔터), 슬래시(/), 직급(과장/부장/대리/연구원/교수 등), 이메일 분리 등 다양한 형태를 완벽 지원합니다.
    """
    result = {
        'user_name': '', 'user_phone': '', 'user_email': '',
        'vendor_rep': '', 'vendor_phone': '', 'vendor_email': ''
    }
    if not text or not text.strip():
        return result

    phone_pattern = r'(01[016789]-?\d{3,4}-?\d{4}|02-?\d{3,4}-?\d{4}|\b\d{4}\b)'
    email_pattern = r'([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})'

    def format_phone(raw_ph):
        dig = re.sub(r'[^\d]', '', raw_ph)
        if len(dig) == 11 and dig.startswith('01'):
            return f'{dig[:3]}-{dig[3:7]}-{dig[7:]}'
        elif len(dig) == 10 and dig.startswith('01'):
            return f'{dig[:3]}-{dig[3:6]}-{dig[6:]}'
        elif len(dig) == 10 and dig.startswith('02'):
            return f'{dig[:2]}-{dig[2:6]}-{dig[6:]}'
        elif len(dig) == 9 and dig.startswith('02'):
            return f'{dig[:2]}-{dig[2:5]}-{dig[5:]}'
        elif len(dig) == 4:
            return f'02-3408-{dig}'
        return raw_ph

    def extract_chunk_info(chunk):
        em = ''
        em_m = re.search(email_pattern, chunk)
        if em_m:
            em = em_m.group(1).strip()
            chunk = chunk.replace(em, ' ')

        ph = ''
        ph_m = re.search(phone_pattern, chunk)
        if ph_m:
            raw_ph = ph_m.group(1).strip()
            ph = format_phone(raw_ph)
            chunk = chunk.replace(raw_ph, ' ')

        prefix = r'^(실사용자|사용자|업체담당자|업체\s*담당자|협력사\s*담당자|협력사|거래처|vendor|공급사|담당자|신청자|성명|이름)[\s:：/\-]*'
        
        lines = [l.strip() for l in chunk.split('\n') if l.strip()]
        cleaned_words = []
        for l in lines:
            l = re.sub(prefix, '', l, flags=re.IGNORECASE).strip()
            l = re.sub(r'[\(\)\[\]/,:：]', ' ', l).strip()
            l = re.sub(r'\s+', ' ', l).strip()
            if l:
                cleaned_words.append(l)

        name = ' '.join(cleaned_words).strip()
        name = re.sub(r'^(실사용자|사용자|업체담당자|협력사|업체)[\s:：/\-]*', '', name).strip()
        return name, ph, em

    # 1. 명시적 키워드(실사용자 / 업체담당자) 기준 분할
    u_match = re.search(r'(?:^|[\r\n\s/])(?P<kw>실사용자|사용자|담당교수)[\s:：/\-]*', text)
    v_match = re.search(r'(?:^|[\r\n\s/])(?P<kw>업체담당자|협력사담당자|업체\s*담당자|거래처담당자|업체|협력사|거래처|vendor|공급사)[\s:：/\-]*', text, flags=re.IGNORECASE)

    if u_match and v_match:
        if u_match.start() < v_match.start():
            u_chunk = text[u_match.start():v_match.start()]
            v_chunk = text[v_match.start():]
        else:
            v_chunk = text[v_match.start():u_match.start()]
            u_chunk = text[u_match.start():]
            
        erp_end = re.search(r'[\r\n]\s*(?:상세보기|기안자|사업장|회계단위|지원사업|계약업체정보|물품내역|첨부문서|심의정보|재원정보)', v_chunk)
        if erp_end:
            v_chunk = v_chunk[:erp_end.start()]

        u_nm, u_ph, u_em = extract_chunk_info(u_chunk)
        v_nm, v_ph, v_em = extract_chunk_info(v_chunk)
        result['user_name'] = u_nm
        result['user_phone'] = u_ph
        result['user_email'] = u_em
        result['vendor_rep'] = v_nm
        result['vendor_phone'] = v_ph
        result['vendor_email'] = v_em

    elif u_match and not v_match:
        u_chunk = text[u_match.start():]
        erp_end = re.search(r'[\r\n]\s*(?:상세보기|기안자|사업장|회계단위|지원사업|계약업체정보|물품내역|첨부문서|심의정보|재원정보)', u_chunk)
        if erp_end:
            u_chunk = u_chunk[:erp_end.start()]
        u_nm, u_ph, u_em = extract_chunk_info(u_chunk)
        result['user_name'] = u_nm
        result['user_phone'] = u_ph
        result['user_email'] = u_em

    elif v_match and not u_match:
        v_chunk = text[v_match.start():]
        erp_end = re.search(r'[\r\n]\s*(?:상세보기|기안자|사업장|회계단위|지원사업|계약업체정보|물품내역|첨부문서|심의정보|재원정보)', v_chunk)
        if erp_end:
            v_chunk = v_chunk[:erp_end.start()]
        v_nm, v_ph, v_em = extract_chunk_info(v_chunk)
        result['vendor_rep'] = v_nm
        result['vendor_phone'] = v_ph
        result['vendor_email'] = v_em

    else:
        # 2. 키워드가 없는 경우: 줄바꿈 및 전화번호 패턴 기준 청크 분할
        raw_lines = [l.strip() for l in text.strip().split('\n') if l.strip()]
        clean_lines = [
            l for l in raw_lines 
            if not re.search(r'기안자\s+담당\s+과장|사업장/회계단위|지원사업\s+교내|상세보기|구매계약서', l)
        ]
        
        chunks = []
        current_chunk = []
        for line in clean_lines:
            has_phone = bool(re.search(phone_pattern, line))
            if has_phone and current_chunk and any(re.search(phone_pattern, c) for c in current_chunk):
                chunks.append('\n'.join(current_chunk))
                current_chunk = [line]
            else:
                current_chunk.append(line)
        if current_chunk:
            chunks.append('\n'.join(current_chunk))

        if len(chunks) == 2:
            c0_nm, c0_ph, c0_em = extract_chunk_info(chunks[0])
            c1_nm, c1_ph, c1_em = extract_chunk_info(chunks[1])
            
            # 💡 [핵심] 이메일 도메인 판별로 실사용자 / 업체 순서 자동 스왑
            if is_internal_sejong_email(c1_em) and not is_internal_sejong_email(c0_em):
                u_nm, u_ph, u_em = c1_nm, c1_ph, c1_em
                v_nm, v_ph, v_em = c0_nm, c0_ph, c0_em
            else:
                u_nm, u_ph, u_em = c0_nm, c0_ph, c0_em
                v_nm, v_ph, v_em = c1_nm, c1_ph, c1_em
                
            result['user_name'] = u_nm
            result['user_phone'] = u_ph
            result['user_email'] = u_em
            result['vendor_rep'] = v_nm
            result['vendor_phone'] = v_ph
            result['vendor_email'] = v_em
        elif len(chunks) == 1:
            u_nm, u_ph, u_em = extract_chunk_info(chunks[0])
            result['user_name'] = u_nm
            result['user_phone'] = u_ph
            result['user_email'] = u_em

    # =========================================================================
    # 💡 [보정 가드] 세종대/USER_DB 등록 이메일은 무조건 '실사용자' 메일로 강제 이관!
    # 업체담당자 메일 칸에는 절대 들어가지 못하도록 철저히 차단
    # =========================================================================
    cur_vemail = result.get('vendor_email', '').strip()
    if cur_vemail and is_internal_sejong_email(cur_vemail):
        if not result.get('user_email'):
            result['user_email'] = cur_vemail
        result['vendor_email'] = ''

    # 텍스트 전체에서 세종대/USER_DB 메일이 존재하는데 user_email이 비어있다면 자동 할당
    if not result.get('user_email'):
        all_emails = re.findall(email_pattern, text)
        for em in all_emails:
            if is_internal_sejong_email(em):
                result['user_email'] = em
                break

    # 실사용자 메일과 업체 메일이 동일하게 잡힌 경우 업체 메일 초기화
    if result.get('user_email') and result.get('vendor_email') == result.get('user_email'):
        result['vendor_email'] = ''

    return result

# ==========================================
# 5. 화면 UI 구성
# ==========================================

# ---------------------------------------------------------
# 🎨 [UI 개선] 토스 & 애플 스타일 미니멀 모던 CSS 주입 (아이콘 폰트 완벽 보호)
# ---------------------------------------------------------
st.markdown("""
<style>
/* [1] Pretendard 웹폰트 로드 */
@import url("https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.9/dist/web/static/pretendard.css");

/* [2] 텍스트 요소에만 Pretendard 적용 (아이콘 폰트 절대 건드리지 않음) */
html, body, .stApp {
    font-family: 'Pretendard', -apple-system, BlinkMacSystemFont, "Apple SD Gothic Neo", sans-serif;
    letter-spacing: -0.02em;
    background-color: #F8F9FA !important;
    color: #191F28;
}

/* 텍스트 요소 개별 적용 */
h1, h2, h3, h4, h5, h6, p, label:not([class*="material"]),
div[data-testid="stMarkdownContainer"] p,
div[data-testid="stMarkdownContainer"] span:not([data-testid="stIconMaterial"]):not([class*="material"]),
div.stButton > button,
div.stDownloadButton > button,
input, textarea, select {
    font-family: 'Pretendard', -apple-system, BlinkMacSystemFont, "Apple SD Gothic Neo", sans-serif !important;
}

/* [3] Material Symbols 및 스트림릿 기본 아이콘 폰트 완벽 보존 */
span[data-testid="stIconMaterial"],
[class*="material-symbols"],
[class*="material-icons"],
[data-testid*="Icon"],
[data-testid*="icon"],
div[data-testid="stSidebarCollapseButton"] button span,
div[data-testid="stSidebarExpandButton"] button span,
div[data-testid="stFileUploader"] button span,
div[data-testid="stPopover"] button span {
    font-family: "Material Symbols Rounded", "Material Symbols Outlined", "Material Icons" !important;
    font-feature-settings: "liga" 1 !important;
    font-style: normal !important;
    font-weight: normal !important;
    letter-spacing: normal !important;
    line-height: 1 !important;
    text-transform: none !important;
    display: inline-block !important;
    white-space: nowrap !important;
    word-wrap: normal !important;
    direction: ltr !important;
    -webkit-font-smoothing: antialiased !important;
}

/* [4] 상단 무지개 바 및 불필요한 장식 숨김 */
div[data-testid="stDecoration"] {
    display: none !important;
}
header[data-testid="stHeader"] {
    background: rgba(248, 249, 250, 0.85) !important;
    backdrop-filter: blur(12px) !important;
}

/* [5] 메인 콘텐츠 여백 최적화 (넓고 편안한 좌우 여백) */
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

/* [6] 사이드바 토스 스타일 */
section[data-testid="stSidebar"] {
    background-color: #FFFFFF !important;
    border-right: 1px solid #ECEFF2 !important;
}
div[data-testid="stSidebarUserContent"] {
    padding-top: 1.5rem !important;
}

/* 사이드바 메뉴 캡슐 버튼 */
section[data-testid="stSidebar"] div[role="radiogroup"] {
    gap: 4px;
}
section[data-testid="stSidebar"] div[role="radiogroup"] label {
    background-color: transparent;
    padding: 10px 14px;
    border-radius: 12px;
    font-size: 14.5px;
    font-weight: 500;
    color: #4E5968;
    transition: all 0.15s ease-in-out;
    margin-bottom: 2px;
    border: 1px solid transparent;
    cursor: pointer;
}
section[data-testid="stSidebar"] div[role="radiogroup"] label:hover {
    background-color: #F2F4F6;
    color: #191F28;
}
section[data-testid="stSidebar"] div[role="radiogroup"] label:has(input:checked) {
    background-color: #E8F3FF !important;
    color: #1B64DA !important;
    font-weight: 700 !important;
    border: 1px solid #D0E4FF !important;
}

/* [7] 플로팅 카드 컨테이너 (애플/토스 시그니처 카드) */
div[data-testid="stVerticalBlock"] > div[data-testid="stVerticalBlockBorderWrapper"] {
    background-color: #FFFFFF !important;
    border-radius: 18px !important;
    border: 1px solid #EAECEF !important;
    box-shadow: 0 4px 18px rgba(0, 0, 0, 0.03) !important;
    padding: 22px 26px !important;
    margin-bottom: 1.2rem !important;
    transition: box-shadow 0.2s ease;
}
div[data-testid="stVerticalBlock"] > div[data-testid="stVerticalBlockBorderWrapper"]:hover {
    box-shadow: 0 6px 24px rgba(0, 0, 0, 0.05) !important;
}

/* [8] 버튼 (토스 시그니처 라운드 버튼) */
div.stButton > button, div.stDownloadButton > button {
    border-radius: 12px !important;
    padding: 10px 20px !important;
    font-size: 14.5px !important;
    font-weight: 600 !important;
    border: 1px solid #E5E8EB !important;
    background-color: #FFFFFF !important;
    color: #333D4B !important;
    box-shadow: 0 1px 2px rgba(0,0,0,0.03) !important;
    transition: all 0.18s cubic-bezier(0.4, 0, 0.2, 1) !important;
}
div.stButton > button:hover, div.stDownloadButton > button:hover {
    background-color: #F2F4F6 !important;
    border-color: #D1D6DB !important;
    color: #191F28 !important;
    transform: translateY(-1px);
}
/* Primary 버튼 (토스 블루) */
div.stButton > button[kind="primary"], div.stDownloadButton > button[kind="primary"] {
    background-color: #3182F6 !important;
    color: #FFFFFF !important;
    border: none !important;
    box-shadow: 0 4px 14px rgba(49, 130, 246, 0.3) !important;
}
div.stButton > button[kind="primary"]:hover, div.stDownloadButton > button[kind="primary"]:hover {
    background-color: #1B64DA !important;
    box-shadow: 0 6px 18px rgba(49, 130, 246, 0.4) !important;
    transform: translateY(-2px);
}

/* [9] 텍스트 입력창 & 텍스트에어리어 & 셀렉트박스 (배경 대비 극대화) */
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

/* input & textarea 내부 요소 스타일 */
div[data-baseweb="input"] input, 
div[data-baseweb="textarea"] textarea {
    background-color: transparent !important;
    color: #191F28 !important;
    font-weight: 500 !important;
}

/* 파일 업로더 드래그앤드롭 영역 대비 */
div[data-testid="stFileUploader"] section {
    border-radius: 14px !important;
    border: 1.5px dashed #D1D6DB !important;
    background-color: #F8F9FA !important;
    transition: all 0.2s ease;
}
div[data-testid="stFileUploader"] section:hover {
    border-color: #3182F6 !important;
    background-color: #F2F4F6 !important;
}

/* [10] 탭 (Apple 세그먼트 컨트롤 스타일) */
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

/* [11] 배지 & 칩 스타일 */
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

/* 제목 타이포그래피 정돈 */
h1 {
    font-size: 1.7rem !important;
    font-weight: 700 !important;
    letter-spacing: -0.03em !important;
    color: #191F28 !important;
    margin-bottom: 0.5rem !important;
}
h2 {
    font-size: 1.35rem !important;
    font-weight: 700 !important;
    color: #191F28 !important;
    letter-spacing: -0.02em !important;
}
h3 {
    font-size: 1.15rem !important;
    font-weight: 600 !important;
    color: #333D4B !important;
}

/* 알림 및 배너 */
div[data-testid="stAlert"] {
    border-radius: 14px;
    border: none;
    box-shadow: 0 2px 8px rgba(0,0,0,0.03);
}
</style>
""", unsafe_allow_html=True)

def render_top_header(title_text):
    col_t, col_l = st.columns([11, 1])
    with col_t:
        st.title(title_text)
    with col_l:
        st.markdown("<div style='height: 14px;'></div>", unsafe_allow_html=True)
        is_auth = st.session_state.get("is_admin_auth", False)
        safe_key_suffix = f"_{abs(hash(title_text)) % 10000}"
        if not is_auth:
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
        else:
            with st.popover("🔓", help="관리자 모드 활성화됨"):
                st.markdown("##### ⚙️ 관리자 모드")
                if st.button("설정 화면 이동", key=f"btn_go{safe_key_suffix}", use_container_width=True):
                    st.session_state["nav_target"] = "⚙️ 관리자 설정"
                    st.rerun()
                if st.button("🔒 로그아웃 (숨김)", key=f"btn_out{safe_key_suffix}", use_container_width=True):
                    st.session_state["is_admin_auth"] = False
                    st.session_state["nav_target"] = get_ui_text("menu_extractor", "💰 구매추출기")
                    st.rerun()

def render_admin_edit_toolbar():
    if st.session_state.get("is_admin_auth", False):
        with st.container(border=True):
            c_tb1, c_tb2 = st.columns([7, 3])
            with c_tb1:
                st.toggle(
                    "✏️ 화면 문구 직접 수정 모드 (기존 위치에서 바로 바꾸기)", 
                    key="ui_live_edit_mode", 
                    help="스위치를 켜면 현재 화면의 제목/문구를 그 자리에서 바로 수정할 수 있습니다."
                )
            with c_tb2:
                if st.session_state.get("ui_live_edit_mode", False):
                    if st.button("💾 화면 문구 즉시 저장", type="primary", use_container_width=True, key="btn_save_live_ui"):
                        changed = False
                        for k in list(DEFAULT_APP_CONFIG.keys()):
                            live_k = f"live_{k}"
                            if live_k in st.session_state:
                                val = str(st.session_state[live_k]).strip()
                                if val:
                                    APP_CONFIG[k] = val
                                    changed = True
                        if changed:
                            save_app_config(APP_CONFIG)
                            st.success("✅ 화면 문구가 저장되었습니다!")
                            time.sleep(0.5)
                            st.rerun()

with st.sidebar:
    s_title = get_ui_text("sidebar_title", "💼 구매 업무 보조")
    st.markdown(f"### {s_title}")
    
    m_ext = get_ui_text("menu_extractor", "💰 구매추출기")
    m_con = get_ui_text("menu_contract", "🤝 구매계약분석")
    m_his = get_ui_text("menu_history", "🕒 분석내역 불러오기")
    base_menus = [m_ext, m_con, m_his]
    if st.session_state.get("is_admin_auth", False):
        base_menus.append("⚙️ 관리자 설정")

    # 예약된 이동 타겟이 있으면 radio 생성 전에 반영
    if "nav_target" in st.session_state:
        target = st.session_state.pop("nav_target")
        if target in ["💰 구매추출기", "구매추출기", m_ext]:
            st.session_state["main_menu"] = m_ext
        elif target in ["🤝 구매계약관리", "🤝 구매계약분석", "구매계약분석", m_con]:
            st.session_state["main_menu"] = m_con
        elif target in ["🕒 분석내용 불러오기", "🕒 분석내역 불러오기", "분석내역 불러오기", m_his]:
            st.session_state["main_menu"] = m_his
        elif target == "⚙️ 관리자 설정":
            st.session_state["main_menu"] = "⚙️ 관리자 설정"
        
    current_menu_choice = st.session_state.get("main_menu", m_ext)
    if current_menu_choice not in base_menus:
        current_menu_choice = m_ext
        st.session_state["main_menu"] = current_menu_choice

    menu = st.radio(
        "메뉴 선택", 
        base_menus, 
        key="main_menu",
        label_visibility="collapsed"
    )
    st.divider()
            
    if st.session_state.get("is_admin_auth", False):
        st.caption("🔓 관리자 모드")
        if st.button("🔒 관리자 로그아웃", key="btn_side_logout", use_container_width=True):
            st.session_state["is_admin_auth"] = False
            st.session_state["nav_target"] = m_ext
            st.rerun()

# ----------------- [1. 구매요청 분석] -----------------
m_ext = get_ui_text("menu_extractor", "💰 구매추출기")
m_con = get_ui_text("menu_contract", "🤝 구매계약분석")
m_his = get_ui_text("menu_history", "🕒 분석내역 불러오기")
if menu in [m_ext, "💰 구매추출기", "구매추출기", "📋 구매요청 분석"]:
    ext_t = get_ui_text("ext_main_title", "💰 구매추출기")
    render_top_header(ext_t)
    render_admin_edit_toolbar()
    
    is_live = st.session_state.get("ui_live_edit_mode", False) and st.session_state.get("is_admin_auth", False)
    if is_live:
        st.info("💡 **화면 문구 직접 수정 모드 ON**: 아래 각 위치의 텍스트 상자를 수정한 뒤 상단의 [💾 화면 문구 즉시 저장]을 누르면 바로 반영됩니다.")
        st.text_input("📍 [위치: 최상단 대제목]", value=ext_t, key="live_ext_main_title")
    
    def on_critical_condition_change():
        if SESSION_REQ_KEY in st.session_state:
            del st.session_state[SESSION_REQ_KEY]
        st.session_state['api_failed'] = False
        st.session_state['ai_errors'] = []
        st.session_state['is_saved'] = False
    
    step1_t = get_ui_text("ext_step1_title", "📥 정보 입력 및 서류 업로드")
    if is_live:
        st.text_input("📍 [위치: STEP 1 섹션 제목]", value=step1_t, key="live_ext_step1_title")
    else:
        st.write(f"### {step1_t}")
    
    with st.container(border=True):
        sub1_t = get_ui_text("ext_sub1_title", "📂 1. 기본 정보 선택")
        if is_live:
            st.text_input("📍 [위치: 1번 항목 소제목]", value=sub1_t, key="live_ext_sub1_title")
        else:
            st.markdown(f"#### {sub1_t}")
        c_type1, c_type2 = st.columns(2)

        with c_type1:
            req_type = st.radio("구매 유형", ["📦 물품", "🛠️ 용역", "🏗️ 공사"], horizontal=True, key="saved_req_type", on_change=on_critical_condition_change)
        with c_type2:
            con_type = st.radio("계약 방식", ["비교견적", "수의계약"], horizontal=True, key="saved_con_type", on_change=on_critical_condition_change)

    col_in1, col_in2 = st.columns(2)
    with col_in1:
        with st.container(border=True):
            sub2_t = get_ui_text("ext_sub2_title", "📝 2. ERP 텍스트 복사")
            if is_live:
                st.text_input("📍 [위치: 2번 항목 소제목]", value=sub2_t, key="live_ext_sub2_title")
            else:
                st.markdown(f"#### {sub2_t}")
            
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
            sub3_t = get_ui_text("ext_sub3_title", "📎 3. AI OCR")
            if is_live:
                st.text_input("📍 [위치: 3번 항목 소제목]", value=sub3_t, key="live_ext_sub3_title")
            else:
                st.markdown(f"#### {sub3_t}")
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
                "- ✅ **유효 기간:** 견적서 유효기간 7일 이상 잔여 여부 (미기재 시 무기한 인정)\n"
                "- ✅ **명의 일치:** 세종대학교 산학협력단 및 예금주/대표자 교차 검증\n"
                "- ✅ **공사 확인:** 시설공사/장비설치 확인서 첨부 여부 (해당 시)\n"
            )
        
    btn_ana_t = get_ui_text("ext_btn_analyze", "🚀 전체 데이터 통합 분석 시작")
    if is_live:
        st.text_input("📍 [위치: 분석 버튼 문구]", value=btn_ana_t, key="live_ext_btn_analyze")
    col_btn1, col_btn2 = st.columns([7, 3])
    with col_btn1: start_all = st.button(btn_ana_t, width="stretch", type="primary")
        
    def execute_save_analysis():
        cur_req = st.session_state.get("p_req", "")
        cur_fin = st.session_state.get("p_fin", "")
        cur_itm = st.session_state.get("p_itm", "")
        cur_req_type = st.session_state.get("saved_req_type", "📦 물품")
        cur_con_type = st.session_state.get("saved_con_type", "비교견적")
        
        b_data = extract_base_info(cur_req, cur_fin)
        if not b_data or not b_data.get('req_no'):
            return False, "구매요구번호를 찾을 수 없습니다."
            
        ai_data = st.session_state.get(SESSION_AI_KEY, {})
        if isinstance(ai_data, list):
            ai_data = ai_data[0] if len(ai_data) > 0 else {}
            
        force_cat = "용역" if "용역" in cur_req_type else ("공사" if "공사" in cur_req_type else None)
        if force_cat:
            items, _ = parse_service_erp(cur_itm, force_cat)
        else:
            items, _ = parse_goods_erp(cur_itm)
            
        v_ven = st.session_state.get("req_in_ven") or (ai_data.get('사업자등록증_상호명') or ai_data.get('업체명', '') if ai_data else '')
        v_biz = st.session_state.get("req_in_biz") or (ai_data.get('사업자등록증_등록번호') or ai_data.get('사업자번호', '') if ai_data else '')
        v_rep = st.session_state.get("req_in_rep") or (ai_data.get('담당자') or ai_data.get('사업자등록증_대표자명') or ai_data.get('대표자명', '') if ai_data else '')
        v_phn = st.session_state.get("req_in_phn") or (ai_data.get('연락처', '') if ai_data else '')
        v_eml = st.session_state.get("req_in_eml") or (ai_data.get('이메일', '') if ai_data else '')
        v_holder = st.session_state.get("req_acc_holder") or (ai_data.get('통장사본_예금주') or ai_data.get('예금주', '') if ai_data else '')
        v_acc = st.session_state.get("req_acc_num") or (ai_data.get('통장사본_계좌번호') or ai_data.get('계좌번호', '') if ai_data else '')

        item_mgrs = []
        for itm in items:
            if itm.get('contact') and "없음" not in itm['contact']:
                item_mgrs.append(itm['contact'])
        unique_item_mgrs = list(dict.fromkeys(item_mgrs))
        final_item_mgr = " / ".join(unique_item_mgrs) if unique_item_mgrs else "⚠️ 미기재"
        if final_item_mgr == "⚠️ 미기재" and b_data.get('pi_name') == SAFETY_CENTER_INFO['pi_name']:
            final_item_mgr = b_data['pi_formatted']

        vendor_mgr_info = f"{v_rep} / {v_phn} / {v_eml}"

        item_remarks = []
        for itm in items:
            rem = itm.get('remark', '').strip()
            if rem and rem not in ['-', '없음', '해당없음']:
                item_remarks.append(f"{itm['name']}: {rem}" if len(items) > 1 else rem)
        final_remarks = " / ".join(item_remarks) if item_remarks else ""
        
        saved_ai = dict(ai_data) if ai_data else {}
        saved_ai["업체명"] = v_ven
        saved_ai["사업자번호"] = v_biz
        saved_ai["담당자"] = v_rep
        saved_ai["연락처"] = v_phn
        saved_ai["이메일"] = v_eml
        saved_ai["예금주"] = v_holder
        saved_ai["계좌번호"] = v_acc
        st.session_state[SESSION_AI_KEY] = saved_ai

        t_amt = get_accurate_total(cur_req, "합계액")
        if t_amt == 0:
            t_amt = get_accurate_total(cur_fin, "총액")
        if t_amt == 0 and items:
            t_amt = sum(itm.get('price', 0) for itm in items)

        raw_to_save = {
            "req_input": cur_req, "fin_input": cur_fin, "itm_input": cur_itm, "ai_res": saved_ai, 
            "req_type": cur_req_type, "con_type": cur_con_type,
            "item_list": items, "remarks": final_remarks
        }
        
        meta_to_save = {
            "요청분석일자": now_kst().strftime("%Y-%m-%d %H:%M:%S"),
            "구매요구번호": b_data['req_no'],
            "과제번호": b_data['p_no'],
            "구매요청액": f"{t_amt:,}원",
            "연구책임자정보": b_data['pi_formatted'], 
            "물품담당자": final_item_mgr,               
            "업체명": v_ven,
            "업체담당자정보": vendor_mgr_info,
            "업체담당자 정보": vendor_mgr_info,
            "사업자번호": v_biz,
            "계좌번호": v_acc,
            "예금주": v_holder,
            "비고사항": final_remarks,
            "물품비고": final_remarks,
            "연구책임자_성명": b_data['pi_name'],
            "연구책임자_학과": b_data['pi_dept'],
            "연구책임자_이메일": b_data['pi_email']
        }
        
        save_to_cache(b_data['req_no'], meta_to_save, raw_to_save)
        st.session_state['is_saved'] = True
        return True, b_data['req_no']

    def reset_all_state():
        st.session_state["p_req"] = ""  
        st.session_state["p_fin"] = ""
        st.session_state["p_itm"] = ""
        st.session_state["file_key"] = st.session_state.get("file_key", 0) + 1
        st.session_state["is_saved"] = False
        st.session_state['api_failed'] = False
        st.session_state['ai_errors'] = []
        for k in ["req_in_ven", "req_in_biz", "req_in_rep", "req_in_phn", "req_in_eml", "req_acc_holder", "req_acc_num"]:
            if k in st.session_state: del st.session_state[k]
        if 'show_reset_warning' in st.session_state: del st.session_state['show_reset_warning']
        for k in [SESSION_REQ_KEY, SESSION_AI_KEY]:
            if k in st.session_state: del st.session_state[k]

    @st.dialog("🚨 데이터 초기화 경고")
    def reset_warning_popup():
        st.error("아직 분석 내역을 시스템(DB)에 저장하지 않았습니다!\n지금 초기화하면 작업 중인 데이터가 **모두 날아갑니다.**")
        st.info("원하시는 처리 방식을 선택해 주세요.")
        
        c_warn1, c_warn2, c_warn3 = st.columns([1.2, 1.2, 0.8])
        
        with c_warn1:
            if st.button("💾 저장 후 초기화", width="stretch", type="primary"):
                success, res_msg = execute_save_analysis()
                if success:
                    st.toast(f"✅ [{res_msg}] 데이터가 안전하게 저장된 후 초기화되었습니다.")
                else:
                    st.toast(f"⚠️ 저장 오류: {res_msg}")
                reset_all_state()
                st.rerun()
                
        with c_warn2:
            if st.button("🗑️ 무시하고 초기화", width="stretch"):
                reset_all_state()
                st.rerun() 
                
        with c_warn3:
            if st.button("취소 🔙", width="stretch"):
                st.rerun()

    with col_btn2: 
        if st.button("🔄 전체 초기화", width="stretch"):
            if st.session_state.get(SESSION_REQ_KEY) and not st.session_state.get('is_saved', False):
                reset_warning_popup()
            else:
                reset_all_state()
                st.rerun()

    def sync_ai_to_widgets(rt):
        """AI 분석 또는 수동 JSON 붙여넣기 결과를 입력창 위젯 세션 상태에 즉시 동기화"""
        if not isinstance(rt, dict):
            return
        v_ven = rt.get("사업자등록증_상호명") or rt.get("업체명", "")
        v_biz = rt.get("사업자등록증_등록번호") or rt.get("사업자번호", "")
        v_rep = rt.get("담당자") or rt.get("사업자등록증_대표자명") or rt.get("대표자명", "")
        v_phn = rt.get("연락처", "")
        v_eml = rt.get("이메일", "")
        v_holder = rt.get("통장사본_예금주") or rt.get("예금주", "")
        v_acc = rt.get("통장사본_계좌번호") or rt.get("계좌번호", "")
        
        if v_ven: st.session_state["req_in_ven"] = v_ven
        if v_biz: st.session_state["req_in_biz"] = v_biz
        if v_rep: st.session_state["req_in_rep"] = v_rep
        if v_phn and v_phn != "미검출": st.session_state["req_in_phn"] = v_phn
        if v_eml and v_eml != "미검출" and not is_internal_sejong_email(v_eml): st.session_state["req_in_eml"] = v_eml
        if v_holder: st.session_state["req_acc_holder"] = v_holder
        if v_acc: st.session_state["req_acc_num"] = v_acc

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
                    sync_ai_to_widgets(rt)
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
        step2_t = get_ui_text("ext_step2_title", "📤 [STEP 2] 분석 결과")
        if is_live:
            st.text_input("📍 [위치: STEP 2 섹션 제목]", value=step2_t, key="live_ext_step2_title")
        else:
            st.write(f"### {step2_t}")

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

            today_str = now_kst().strftime('%Y-%m-%d')
            extra_instruction = "🚨 이 건은 '공사' 건입니다. 첨부 서류 중 '시설공사 확인서' 또는 '장비설치 확인서'가 반드시 포함되어야 하므로 이를 꼼꼼히 확인하세요.\n" if req_type == "🏗️ 공사" else ""
            
            fallback_prompt = (
                f"🚨 [기준일] 오늘 날짜는 {today_str}입니다. 만료일자 계산 시 참고하세요.\n"
                f"{extra_instruction}"
                "아래 지침에 따라 첨부된 문서들을 분석하고, 반드시 순수한 JSON 형식으로만 응답해. 절대 마크다운(```json)이나 다른 설명 텍스트를 넣지 마.\n\n"
                "■ [문서 자동 판별 및 추출 원칙 (파일명 무관)]:\n"
                "1. 첨부 파일명이 'ㅅㄷ', 'ㅌㅅ', '1', '스캔' 등 임의의 이름이더라도, 파일 본문의 양식/텍스트를 직접 보고 문서 역할을 구분할 것:\n"
                "   - [사업자등록증]: 국세청 양식, '등록번호', '법인명(상호)', '대표자', '개업연월일', '사업장소재지' 등이 적힌 서식.\n"
                "   - [통장사본]: 은행명, 예금주, 계좌번호가 적힌 통장 이미지/사본. (1개 PDF 내 여러 페이지에 사업자등록증과 통장사본이 함께 있는 경우 각 페이지를 개별 판별)\n"
                "   - [견적서/비교견적서]: 공급자 정보, 품목, 단가, 공급가액, 세액, 총금액이 적힌 견적 문서.\n"
                "   - [기안문/품의서/결재문서]: 대학/기관 내부 기안문, 시행일자, 결재선 등이 있는 문서는 내부 행정서류이므로 견적 비교 대상에서 제외할 것.\n"
                "2. [핵심] 첨부 문서 중 '사업자등록증'과 '통장사본' 양식이 단 1장이라도 발견되면, 견적서 최저가 선정 여부와 관계없이 해당 서류 본문의 정보(상호, 사업자번호, 대표자명, 예금주, 계좌번호, 은행명)를 무조건 독립 추출하여 아래 JSON 필드에 정확히 채울 것.\n\n"
                "{\n"
                '  "사업자등록증_상호명": "첨부된 사업자등록증 상의 \'법인명(단체명)\' 또는 \'상호\' (사업자등록증이 없으면 견적서 상호)",\n'
                '  "사업자등록증_등록번호": "첨부된 사업자등록증 상의 사업자등록번호 (반드시 000-00-00000 형식)",\n'
                '  "사업자등록증_대표자명": "첨부된 사업자등록증 상의 대표자 성명",\n'
                '  "통장사본_예금주": "첨부된 통장사본 상의 예금주명",\n'
                '  "통장사본_계좌번호": "첨부된 통장사본 상의 계좌번호",\n'
                '  "통장사본_은행명": "첨부된 통장사본 상의 은행명",\n'
                '  "업체명": "최종 선정된 업체명 (사업자등록증 상호 우선)",\n'
                '  "사업자번호": "최종 선정 업체의 사업자번호 (사업자등록증 번호 우선, 000-00-00000 형식)",\n'
                '  "대표자명": "사업자등록증 상의 대표자 성명",\n'
                '  "담당자": "최저가 업체의 견적 담당자명",\n'
                '  "연락처": "최저가 업체 담당자의 010 전화번호 (없으면 대표번호)",\n'
                '  "이메일": "최저가 업체의 이메일",\n'
                '  "예금주": "통장사본의 예금주명",\n'
                '  "계좌번호": "통장사본의 계좌번호",\n'
                '  "비교견적리스트": [ {"업체명": "A업체", "사업자번호": "123-45-67890", "제출금액": 10000}, {"업체명": "B업체", "사업자번호": "987-65-43210", "제출금액": 12000} ],\n'
                '  "최저가판단과정": "위 비교견적리스트를 바탕으로 가장 금액이 낮은 업체를 최종 선정하는 과정을 짧게 작성",\n'
                '  "수신자명확인": "첨부된 모든 견적서(비교견적서 포함)의 수신자가 모두 \'세종대학교 산학협력단\' 명의가 맞으면 \'예\', 하나라도 누락되거나 다르면 \'아니오(사유)\'",\n'
                '  "만료일자": "견적서 본문에 유효기간 만료일이 명시적으로 적혀 있는 경우에만 계산하여 YYYY-MM-DD 형식으로 기재. 본문에 유효기간이 없거나 비어있는 경우 절대로 임의 계산하지 말고 반드시 \'무기한\'으로 기재할 것",\n'
                '  "동일규격확인": "복수 견적서 간 동일 물품이면 \'예\', 다르면 \'아니오(사유)\', 1장뿐이면 \'단일견적\'",\n'
                '  "사업자일치여부": "최저가(최종 선정) 견적서와 사업자등록증 상의 정보가 일치하면 \'예\', 다르면 \'아니오(사유)\'",\n'
                '  "총금액": "최종 선정된 업체의 견적 총액 (반드시 부가가치세/VAT 포함된 최종 합계 금액을 숫자로만 기재)",\n'
                '  "공사확인서유무": "첨부 파일 중 \'시설공사\' 또는 \'장비설치 확인서\' 문서가 있으면 \'예\', 없으면 \'아니오\', 비슷한 문서가 있으면 \'확인요망\'",\n'
                '  "품목리스트": [ {"품명": "품명 텍스트", "수량": 1, "단가": 1000, "금액": 1000} ]\n'
                "}"
            )

            st.info("💡 **[수동 분석 가이드]**\n1. 제미나이(Gemini) 공식 웹페이지에 견적서 등 첨부파일을 업로드하세요.\n2. 아래 📋 아이콘을 눌러 프롬프트를 복사한 뒤 제미나이에 붙여넣고 전송합니다.\n3. 결과물로 나온 JSON 데이터( `{` 부터 `}` 까지 )를 복사해 아래 입력칸에 붙여넣습니다.")
            
            st.code(fallback_prompt, language="text")

            with st.container(border=True):
                st.markdown("#### ✍️ 제미나이 분석 결과(JSON) 붙여넣기")
                manual_json = st.text_area("결과값", height=150, label_visibility="collapsed", placeholder="여기에 결과값 { JSON } 형태를 그대로 붙여넣으세요.", key="manual_json_area")
                
                c_btn1, c_btn2 = st.columns(2)
                with c_btn1:
                    if st.button("🚀 수동 데이터 시스템 적용", type="primary", width="stretch"):
                        raw_json_str = manual_json or st.session_state.get("manual_json_area", "")
                        if not raw_json_str.strip():
                            st.warning("⚠️ 입력창에 제미나이 분석 결과(JSON)를 먼저 붙여넣어주세요.")
                            st.toast("⚠️ 입력창에 JSON 데이터를 먼저 붙여넣어주세요.")
                        else:
                            parsed_manual = safe_json_load(raw_json_str)
                            if parsed_manual:
                                if isinstance(parsed_manual, list) and len(parsed_manual) > 0:
                                    parsed_manual = parsed_manual[0]
                                st.session_state[SESSION_AI_KEY] = parsed_manual
                                st.session_state['api_failed'] = False
                                st.session_state['ai_errors'] = []
                                sync_ai_to_widgets(parsed_manual)
                                st.toast("✅ 수동 데이터가 시스템에 완벽하게 적용되었습니다!")
                                st.success("✅ 수동 데이터가 완벽하게 적용되었습니다! 잠시 후 분석 화면으로 전환됩니다...")
                                time.sleep(0.5)
                                st.rerun() 
                            else:
                                st.error("🚨 형식 오류: JSON 결과값을 파싱할 수 없습니다. 중괄호 { } 가 온전히 포함되어 있는지 확인해주세요.")
                                st.toast("🚨 JSON 파싱 실패: 형식을 다시 확인해주세요.")
                with c_btn2:
                    if st.button("⏭️ AI 검증 포기하고 그냥 넘어가기", width="stretch"):
                        st.session_state[SESSION_AI_KEY] = {}
                        st.session_state['api_failed'] = False
                        st.session_state['ai_errors'] = []
                        st.toast("⏩ AI 검증을 건너뛰고 기본 ERP 분석으로 진행합니다.")
                        st.rerun()

            st.stop() 
        
        base_data = extract_base_info(raw_req, raw_fin)
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
                today = today_kst()
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
                
                main_biz = ai.get('사업자등록증_등록번호') or ai.get('사업자번호', '')
                main_ven = ai.get('사업자등록증_상호명') or ai.get('업체명', '최종선정업체')
                if not is_missing(main_biz):
                    ai['사업자번호'] = format_bizno(main_biz)
                    biz_check_list.append({"name": main_ven, "no": ai['사업자번호']})
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
                            status = check_nts_business(biz['no'], NTS_API_KEY)
                            nts_reports.append(f"{biz['name']}: {status}")
                            
                            if "휴업자" in status:
                                auto_reject_reasons.append(f"국세청 조회 결과: '휴업' 상태인 사업자 포함 - {biz['name']} ({biz['no']})")
                            elif "폐업자" in status:
                                auto_reject_reasons.append(f"국세청 조회 결과: '폐업' 상태인 사업자 포함 - {biz['name']} ({biz['no']})")
                
                for q in ai.get('비교견적리스트', []):
                    if is_missing(q.get('사업자번호', '')):
                        nts_reports.append(f"{q.get('업체명', '비교업체')}: ⚠️ 번호 미기재(조회생략)")
                                
                nts_status_res = " / ".join(nts_reports) if nts_reports else "⚠️ 사업자번호 미검출 (전체 조회 불가)"
                
                exp_date_str = str(ai.get('만료일자', '')).strip()
                v_exp_result = "⚠️ 유효기간 확인불가 (직접 확인 필요)"
                
                if any(kw in exp_date_str for kw in ["무기한", "미기재", "없음", "해당없음", "별도 표기 없음", "공란"]):
                    v_exp_result = "✅ 통과 (유효기간 별도 표기 없음 - 무기한 인정)"
                elif not is_missing(exp_date_str):
                    match = re.search(r'\d{4}-\d{2}-\d{2}', exp_date_str)
                    if match:
                        try:
                            exp_date = datetime.datetime.strptime(match.group(), '%Y-%m-%d').date()
                            today = today_kst()
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

                        # 💡 부대비용(운송비/배송비/설치비 등) 녹임 여부 및 견적서 내 단가×수량 단순 확인 안내 (반려 사유 아님)
                        soft_check_notes = []
                        extra_fee_kws = ['운송', '배송', '택배', '설치', '운반', '출장', '부대비용', '할인', '절사', '네고', '단수']
                        detected_fee_items = []
                        arith_notes = []

                        for idx_q, q_itm in enumerate(ai_items, 1):
                            q_nm = str(q_itm.get('품명', '')).strip()
                            q_qty = safe_int(q_itm.get('수량', 0))
                            q_up = safe_int(q_itm.get('단가', 0))
                            q_amt = safe_int(q_itm.get('금액', 0))

                            if any(kw in q_nm for kw in extra_fee_kws):
                                fee_amt_val = q_amt if q_amt != 0 else (q_up * (q_qty or 1))
                                detected_fee_items.append(f"{q_nm}({fee_amt_val:,}원)")

                            if q_qty > 0 and q_up > 0 and q_amt > 0:
                                expected_supply = q_up * q_qty
                                expected_vat_inc = round(expected_supply * 1.1)
                                if abs(expected_supply - q_amt) > 10 and abs(expected_vat_inc - q_amt) > 10:
                                    arith_notes.append(
                                        f"견적 품목 {idx_q}번 `{q_nm}`: 단가({q_up:,}) × 수량({q_qty}) = {expected_supply:,}원 ↔ 표기 금액({q_amt:,}원)"
                                    )

                        if detected_fee_items:
                            soft_check_notes.append(
                                f"📦 **부대비용/조정 항목 감지**: 견적서에 **{', '.join(detected_fee_items)}** 항목이 포함되어 있습니다. ERP 물품내역 단가/금액에 적절히 녹여서 반영되었는지 확인해 주세요."
                            )
                        elif len(ai_items) != len(item_list):
                            soft_check_notes.append(
                                f"📋 **품목 건수 차이 참고**: ERP 품목({len(item_list)}건)과 견적서 품목({len(ai_items)}건)의 행 개수가 다릅니다. (운송비·설치비 합산 반영 또는 세트 품목 여부 확인)"
                            )

                        if arith_notes:
                            soft_check_notes.append(
                                "🧮 **견적서 내 단가×수량 참고 확인** (부대비용 합산·단수절사·AI 오독 여부 눈으로만 확인):\n  - " + "\n  - ".join(arith_notes)
                            )

                        if soft_check_notes:
                            st.info("💡 **[참고 / 육안 확인 사항 (반려 대상 아님)]**\n\n" + "\n\n".join(soft_check_notes))
        else:
            st.info("📄 첨부된 서류가 없어 AI 교차 검증(금액/품목 대조)은 생략하고 ERP 텍스트 분석 결과만 출력합니다.")

        st.divider()
        erp_bas_t = get_ui_text("ext_erp_basic_title", "📝 ERP 기본 정보")
        if is_live:
            st.text_input("📍 [위치: 기본 정보 소제목]", value=erp_bas_t, key="live_ext_erp_basic_title")
        else:
            st.write(f"#### {erp_bas_t}")
        r1_c1, r1_c2, r1_c3 = st.columns(3)
        with r1_c1: display_status_value("🔢 구매요구번호", base_data['req_no'])
        with r1_c2: display_status_value("📌 과제번호", base_data['p_no'])
        with r1_c3: display_status_value("👤 연구책임자 성명", base_data['pi_name'])
        
        r2_c1, r2_c2, r2_c3 = st.columns(3)
        with r2_c1: display_status_value("🏫 연구책임자 학과", base_data['pi_dept'])
        with r2_c2: display_status_value("💰 구매요청액", f"{total_amt:,}원")
        with r2_c3: display_status_value("📧 연구책임자 정보", base_data['pi_formatted'])

        if base_data.get('pi_name') == SAFETY_CENTER_INFO['pi_name']:
            st.info("ℹ️ **[대학안전관리센터 건 자동 적용]** 과제담당자가 대학안전관리센터인 건으로 감지되어 연구책임자 및 실사용자 정보가 **김수동 선생님 (02-3408-4081 / sudong_kim@sejong.ac.kr)**으로 자동 설정되었습니다.")
        
        st.divider()
        erp_itm_t = get_ui_text("ext_erp_items_title", "📦 ERP 품목 내역")
        if is_live:
            st.text_input("📍 [위치: 품목 내역 소제목]", value=erp_itm_t, key="live_ext_erp_items_title")
        else:
            st.write(f"#### {erp_itm_t}")
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
        
        curr_ven = st.session_state.get("req_in_ven") or (ai.get('사업자등록증_상호명') or ai.get('업체명', '') if ai else '')
        curr_biz = st.session_state.get("req_in_biz") or (ai.get('사업자등록증_등록번호') or ai.get('사업자번호', '') if ai else '')
        curr_p_no = base_data.get('p_no', '')
        curr_req_no = base_data.get('req_no', '')

        # 동일과제/동일업체 1개월 내 구매요청 이력 검사 및 UI 배너 렌더링
        dup_warns, dup_total, dup_matches = check_split_procurement_risk(curr_p_no, curr_ven, curr_biz, curr_req_no, total_amt)
        if dup_matches:
            with st.container(border=True):
                if dup_total > 22000000:
                    st.error(
                        f"🚨 **[주의] 동일과제·동일업체 1개월 내 누적 2,200만 원 초과 의심 건**\n\n"
                        f"- **과제번호**: `{curr_p_no}` / **업체명**: `{curr_ven}`\n"
                        f"- **현재 요청액**: `{total_amt:,}원` + **과거 1개월 내 등록액**: `{sum(m['amt'] for m in dup_matches):,}원` = **총 `{dup_total:,}원`**\n"
                        f"- **관련 과거 요구건**: " + ", ".join([f"`{m['req_no']}` ({m['date']}, {m['amt']:,}원)" for m in dup_matches]) + "\n\n"
                        f"💡 DB 등록일 기준 1개월 이내 건이므로, **실제 구매계약일 기준으로 1개월 이내 2,200만 원 초과(입찰 대상) 여부**를 반드시 확인하세요."
                    )
                else:
                    st.warning(
                        f"💡 **[안내] 동일과제·동일업체 1개월 내 구매요청 이력 감지**\n\n"
                        f"- **과제번호**: `{curr_p_no}` / **업체명**: `{curr_ven}`\n"
                        f"- **현재 요청액**: `{total_amt:,}원` + **과거 1개월 내 등록액**: `{sum(m['amt'] for m in dup_matches):,}원` = **총 `{dup_total:,}원`**\n"
                        f"- **관련 과거 요구건**: " + ", ".join([f"`{m['req_no']}` ({m['date']}, {m['amt']:,}원)" for m in dup_matches]) + "\n\n"
                        f"📌 결재의견 '2. 확인사항'에 사전 확인 문구가 자동으로 반영되었습니다."
                    )

        st.divider()
        st.write("#### ✍️ 결재 의견 (자동 생성)")
        st.code(generate_opinion(req_type, con_type, total_amt, item_list, raw_fin, raw_req, ai_res=ai, project_no=curr_p_no, vendor_name=curr_ven, vendor_bizno=curr_biz, current_req_no=curr_req_no), language="text")
        
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

        ai_email = ai.get('이메일', '') if ai else ''
        if ai_email and (ai_email == "미검출" or is_internal_sejong_email(ai_email)):
            ai_email = ""

        def_ven = st.session_state.get("req_in_ven") or (ai.get('사업자등록증_상호명') or ai.get('업체명', '') if ai else '')
        def_biz = st.session_state.get("req_in_biz") or (ai.get('사업자등록증_등록번호') or ai.get('사업자번호', '') if ai else '')
        def_rep = st.session_state.get("req_in_rep") or (ai.get('담당자') or ai.get('사업자등록증_대표자명') or ai.get('대표자명', '') if ai else '')

        c_v1, c_v2, c_v3, c_v4, c_v5 = st.columns(5)
        with c_v1: t_ven = st.text_input("업체명", value=def_ven, key="req_in_ven")
        with c_v2: t_biz = st.text_input("사업자번호", value=def_biz, key="req_in_biz")
        with c_v3: t_rep = st.text_input("업체담당자명", value=def_rep, key="req_in_rep")
        with c_v4: t_phn = st.text_input("연락처", value=ai_phone, key="req_in_phn") 
        with c_v5: t_eml = st.text_input("이메일", value=ai_email, key="req_in_eml")
        
        st.caption("✨ **원클릭 복사 보드**")
        c_cp1, c_cp2, c_cp3 = st.columns([1, 1, 2])
        with c_cp1: st.code(t_ven, language=None)
        with c_cp2: st.code(t_biz, language=None)
        with c_cp3: st.code(f"{t_rep} / {t_phn} / {t_eml}", language=None) 
        
        fold_nm = build_recommended_folder_name(t_dept, t_name, t_ven, base_data.get('req_no', ''))
        
        st.info("💡 오른쪽 아이콘을 눌러 폴더 이름을 복사하거나, 아래 버튼으로 지정 상위 폴더 하위에 즉시 생성하세요.")
        st.code(fold_nm, language="text")

        cur_base_dir = APP_CONFIG.get("work_base_dir") or os.path.join(os.path.expanduser("~"), "Desktop", "구매업무")
        c_fd1, c_fd2, c_fd3 = st.columns([2.5, 1.3, 1.0])
        with c_fd1:
            in_base_dir = st.text_input("📂 고정 상위 폴더 경로", value=cur_base_dir, key="ext_work_base_dir", label_visibility="collapsed", placeholder="예: C:\\Users\\사용자\\Desktop\\구매업무")
        with c_fd2:
            if st.button("📁 상위 폴더에 새 폴더 즉시 생성", key="btn_mk_req_folder", width="stretch"):
                try:
                    if in_base_dir.strip() and in_base_dir.strip() != cur_base_dir:
                        APP_CONFIG["work_base_dir"] = in_base_dir.strip()
                        save_app_config(APP_CONFIG)
                    target_f_dir = os.path.join(in_base_dir.strip() or cur_base_dir, fold_nm)
                    os.makedirs(target_f_dir, exist_ok=True)
                    saved_cnt = 0
                    if up_files:
                        for uf in up_files:
                            uf_name = sanitize_win_name(uf.name)
                            if uf_name:
                                with open(os.path.join(target_f_dir, uf_name), "wb") as f_out:
                                    f_out.write(uf.getvalue())
                                saved_cnt += 1
                    st.session_state["last_created_req_folder"] = target_f_dir
                    msg_extra = f" (첨부파일 {saved_cnt}건 함께 저장됨)" if saved_cnt > 0 else ""
                    st.toast(f"✅ 폴더 생성 완료!{msg_extra}")
                except Exception as e_fd:
                    st.error(f"🚨 폴더 생성 실패: {e_fd}")
        with c_fd3:
            if st.button("🗂️ 폴더 열기", key="btn_open_req_folder", width="stretch"):
                target_open = st.session_state.get("last_created_req_folder") or os.path.join(in_base_dir.strip() or cur_base_dir, fold_nm)
                if not os.path.exists(target_open):
                    target_open = in_base_dir.strip() or cur_base_dir
                    os.makedirs(target_open, exist_ok=True)
                ok_op, err_op = open_local_folder(target_open)
                if not ok_op:
                    st.warning(f"폴더 열기 불가 (로컬 환경에서만 지원): {err_op}")
        if st.session_state.get("last_created_req_folder"):
            st.caption(f"📍 생성된 경로: `{st.session_state['last_created_req_folder']}`")

        st.divider()
        bank_t = get_ui_text("ext_bank_title", "🏦 통장사본 계좌정보")
        if is_live:
            st.text_input("📍 [위치: 계좌 정보 소제목]", value=bank_t, key="live_ext_bank_title")
        else:
            st.write(f"#### {bank_t}")
        def_holder = st.session_state.get("req_acc_holder") or (ai.get('통장사본_예금주') or ai.get('예금주', '') if ai else '')
        def_acc = st.session_state.get("req_acc_num") or (ai.get('통장사본_계좌번호') or ai.get('계좌번호', '') if ai else '')

        c_b1, c_b2 = st.columns(2)
        with c_b1:
            t_acc_holder = st.text_input("예금주 (수정 가능)", value=def_holder, key="req_acc_holder")
        with c_b2:
            t_acc_num = st.text_input("계좌번호 (수정 가능)", value=def_acc, key="req_acc_num")

        st.divider()
        if st.button("💾 현재 분석 내역을 시스템(DB)에 저장하기", type="primary", width="stretch", key="btn_save_analysis_db"):
            with st.spinner("⏳ 로컬 캐시 및 구글 시트 DB에 저장 중입니다..."):
                success, res_msg = execute_save_analysis()
            if success:
                st.toast(f"✅ 구매요구번호 [{res_msg}] 건이 DB 및 구글 시트에 안전하게 저장되었습니다!")
                st.success(f"✅ 구매요구번호 [{res_msg}] 건의 핵심 데이터가 로컬 및 구글 시트에 안전하게 저장되었습니다! '분석내용 불러오기' 메뉴에서 확인 가능합니다.")
            else:
                st.toast(f"🚨 저장 실패: {res_msg}")
                st.error(f"❌ 저장 실패: {res_msg}")

# ----------------- [2. 구매계약 분석 및 발주서 자동 생성] -----------------
elif menu in [m_con, "🤝 구매계약분석", "구매계약분석", "🤝 구매계약관리", "구매계약관리", "🤝 구매계약 분석"]:
    con_main_t = get_ui_text("con_main_title", "🤝 구매계약분석")
    render_top_header(con_main_t)
    render_admin_edit_toolbar()
    
    is_live = st.session_state.get("ui_live_edit_mode", False) and st.session_state.get("is_admin_auth", False)
    if is_live:
        st.info("💡 **화면 문구 직접 수정 모드 ON**: 아래 각 위치의 텍스트 상자를 수정한 뒤 상단의 [💾 화면 문구 즉시 저장]을 누르면 바로 반영됩니다.")
        st.text_input("📍 [위치: 최상단 대제목]", value=con_main_t, key="live_con_main_title")
        
    con_desc_t = get_ui_text("con_desc", "여러 개의 계약 건을 한 번에 입력하고 분석하여, 각각의 워드 발주서를 빠르게 생성할 수 있습니다.")
    if is_live:
        st.text_input("📍 [위치: 대제목 하단 안내 문구]", value=con_desc_t, key="live_con_desc")
    else:
        st.info(con_desc_t)
    
    if 'con_input_count' not in st.session_state:
        st.session_state['con_input_count'] = 2

    con_card_pat = get_ui_text("con_card_title", "📄 [계약 {no}] 데이터 입력")
    if is_live:
        st.text_input("📍 [위치: 계약 카드 상단 제목] ({no}는 계약 번호로 자동 치환)", value=con_card_pat, key="live_con_card_title")

    raw_cons = []
    raw_memos = []
    for i in range(st.session_state['con_input_count']):
        with st.container(border=True):
            disp_card_t = con_card_pat.replace("{no}", str(i+1))
            st.markdown(f"##### {disp_card_t}")
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

    col_btn1, col_btn2 = st.columns([7, 3])
    btn_lbl = get_ui_text("btn_load_contract", "계약 데이터 불러오기")
    if is_live:
        st.text_input("📍 [위치: 계약 불러오기 버튼]", value=btn_lbl, key="live_btn_load_contract")
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

                # -------------------------------------------------------------
                # 💡 [Part 2 - 제안 1] 구매요구번호 기준 캐시(DB) 자동 조회 및 프리필
                # -------------------------------------------------------------
                req_no = str(cd.get('req_no', '')).strip()
                cached_data = load_from_cache(req_no) if req_no else None
                
                cached_meta = {}
                if req_no and os.path.exists(CACHE_FILE):
                    try:
                        with open(CACHE_FILE, "r", encoding="utf-8") as f:
                            local_db = json.load(f)
                            if req_no in local_db:
                                cached_meta = local_db[req_no].get("meta", {})
                    except Exception:
                        pass

                if cached_data or cached_meta:
                    cd['from_cache'] = True
                    ai_c = cached_data.get("ai_res", {}) if cached_data else {}
                    
                    # 1) 업체담당자 정보 연동
                    c_vrep = ai_c.get("담당자") or ""
                    c_vph = ai_c.get("연락처") or ""
                    c_vem = ai_c.get("이메일") or ""
                    if not (c_vrep or c_vph or c_vem) and cached_meta.get("업체담당자정보"):
                        v_parts = [p.strip() for p in str(cached_meta["업체담당자정보"]).split("/")]
                        if len(v_parts) >= 1: c_vrep = v_parts[0]
                        if len(v_parts) >= 2: c_vph = v_parts[1]
                        if len(v_parts) >= 3: c_vem = v_parts[2]

                    # 2) 실사용자 정보 (물품담당자 또는 연구책임자)
                    c_unm, c_uph, c_uem = "", "", ""
                    is_cached_safety = (
                        cached_meta.get("연구책임자_성명") == SAFETY_CENTER_INFO["pi_name"] or 
                        cached_meta.get("연구책임자_학과") == SAFETY_CENTER_INFO["pi_dept"] or
                        is_safety_center_case(txt)
                    )
                    if is_cached_safety:
                        c_unm = SAFETY_CENTER_INFO['pi_name']
                        c_uph = SAFETY_CENTER_INFO['pi_phone']
                        c_uem = SAFETY_CENTER_INFO['pi_email']
                    else:
                        c_mgr_str = cached_meta.get("물품담당자", "")
                        if c_mgr_str and "미기재" not in c_mgr_str and "없음" not in c_mgr_str:
                            u_parts = [p.strip() for p in str(c_mgr_str).split("/")]
                            if len(u_parts) >= 1: c_unm = u_parts[0]
                            if len(u_parts) >= 2:
                                if "@" in u_parts[1]: c_uem = u_parts[1]
                                else: c_uph = u_parts[1]
                            if len(u_parts) >= 3:
                                if "@" in u_parts[2]: c_uem = u_parts[2]
                                else: c_uph = u_parts[2]

                        # 실사용자 연락처/이메일 부재 시 USER_DB 조회
                        if c_unm and (not c_uph or not c_uem):
                            u_lookup = lookup_user_contact(c_unm)
                            if u_lookup:
                                if not c_uph: c_uph = u_lookup.get("phone", "")
                                if not c_uem: c_uem = u_lookup.get("email", "")

                    # 3) 비고사항
                    c_rem = cached_meta.get("비고사항") or (cached_data.get("remarks") if cached_data else "")
                    cd['cached_remarks'] = c_rem

                    # 스마트 메모에 반영 (단, 사용자가 노션 메모 직접 입력하지 않은 경우에만 프리필)
                    if not cd['smart_memo'].get('vendor_rep') and c_vrep:
                        cd['smart_memo']['vendor_rep'] = c_vrep
                    if not cd['smart_memo'].get('vendor_phone') and c_vph:
                        cd['smart_memo']['vendor_phone'] = c_vph
                    if not cd['smart_memo'].get('vendor_email') and c_vem:
                        cd['smart_memo']['vendor_email'] = c_vem

                    if not cd['smart_memo'].get('user_name') and c_unm:
                        cd['smart_memo']['user_name'] = c_unm
                    if not cd['smart_memo'].get('user_phone') and c_uph:
                        cd['smart_memo']['user_phone'] = c_uph
                    if not cd['smart_memo'].get('user_email') and c_uem:
                        cd['smart_memo']['user_email'] = c_uem

                    if not cd.get('vendor_name') and cached_meta.get("업체명"):
                        cd['vendor_name'] = cached_meta.get("업체명")

                    # 4) 연구책임자 학과/성명 연동 (발주서 자동 폴더 생성용)
                    if is_cached_safety:
                        cd['pi_dept'] = SAFETY_CENTER_INFO["pi_dept"]
                        cd['pi_name'] = SAFETY_CENTER_INFO["pi_name"]
                        cd['pi_phone'] = SAFETY_CENTER_INFO["pi_phone"]
                        cd['pi_email'] = SAFETY_CENTER_INFO["pi_email"]
                    else:
                        if not cd.get('pi_dept') and cached_meta.get("연구책임자_학과"):
                            cd['pi_dept'] = cached_meta.get("연구책임자_학과")
                        if not cd.get('pi_name') and cached_meta.get("연구책임자_성명"):
                            cd['pi_name'] = cached_meta.get("연구책임자_성명")
                        if (not cd.get('pi_dept') or not cd.get('pi_name')) and cached_meta.get("연구책임자정보"):
                            pi_info_parts = str(cached_meta.get("연구책임자정보", "")).strip().split()
                            if len(pi_info_parts) >= 2:
                                if not cd.get('pi_dept'): cd['pi_dept'] = " ".join(pi_info_parts[:-1])
                                if not cd.get('pi_name'): cd['pi_name'] = pi_info_parts[-1]

                # 별도 메모창에 입력된 내용이 있으면 최우선 적용 (덮어쓰기)
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
                u_em_c = sm.get('user_email', cd.get('email', ''))
                v_em_c = sm.get('vendor_email', '')
                
                # 💡 [핵심] 세종대/USER_DB 이메일은 업체 메일로 들어가지 못하도록 방어
                if is_internal_sejong_email(v_em_c):
                    if not u_em_c:
                        u_em_c = v_em_c
                    v_em_c = ''

                st.session_state[f"u_nm_{i}"] = sm.get('user_name', '')
                st.session_state[f"u_ph_{i}"] = sm.get('user_phone', '')
                st.session_state[f"u_em_{i}"] = u_em_c
                st.session_state[f"v_rep_{i}"] = sm.get('vendor_rep', '')
                st.session_state[f"v_ph_{i}"] = sm.get('vendor_phone', '')
                st.session_state[f"v_em_{i}"] = v_em_c
        else:
            st.warning("입력된 텍스트가 없습니다. 먼저 텍스트를 붙여넣어주세요.")

    if 'multi_con_data' in st.session_state:
        st.divider()
        st.write(f"### 📝 2. 발주요청서 입력 폼 (총 {len(st.session_state['multi_con_data'])}건 추출됨)")
        
        cur_con_base_dir = APP_CONFIG.get("work_base_dir") or os.path.join(os.path.expanduser("~"), "Desktop", "구매업무")
        with st.container(border=True):
            c_cb1, c_cb2 = st.columns([3.5, 1.2])
            with c_cb1:
                con_base_dir_in = st.text_input(
                    "📂 발주서 자동 분류/저장 상위 폴더 경로 (하위에 구매요구번호 폴더를 자동 탐색하거나 새로 생성하여 저장합니다)",
                    value=cur_con_base_dir,
                    key="con_work_base_dir"
                )
            with c_cb2:
                st.markdown("<div style='height: 28px;'></div>", unsafe_allow_html=True)
                if st.button("💾 기본 경로로 저장", key="btn_save_con_base_dir", width="stretch"):
                    if con_base_dir_in.strip():
                        APP_CONFIG["work_base_dir"] = con_base_dir_in.strip()
                        save_app_config(APP_CONFIG)
                        st.toast("✅ 기본 작업 상위 폴더 경로가 저장되었습니다!")
        
        for i, cd in enumerate(st.session_state['multi_con_data']):
            with st.expander(f"📌 [계약 {i+1}] {cd.get('con_name', '건명 미상')}", expanded=True):
                if cd.get('from_cache'):
                    st.success(f"✨ **[구매추출기 DB 연동 완료]** 구매요구번호 `[{cd.get('req_no')}]`의 분석 데이터(업체담당자, 실사용자)를 자동으로 불러왔습니다.")
                    if cd.get('cached_remarks'):
                        st.info(f"💡 **[구매요청 시 비고사항]** {cd.get('cached_remarks')}")
                
                if cd.get('pi_name') == SAFETY_CENTER_INFO['pi_name']:
                    st.info("ℹ️ **[대학안전관리센터 건 자동 적용]** 과제담당자가 대학안전관리센터인 건으로 감지되어 연구책임자 및 실사용자 정보가 **김수동 선생님 (02-3408-4081 / sudong_kim@sejong.ac.kr)**으로 자동 설정되었습니다.")
                
                # 💡 노션 비고란 실시간 즉시 반영 도구
                with st.expander("⚡ 노션 비고란 메모 즉시 붙여넣기 (실사용자/업체 자동 배분)", expanded=False):
                    quick_memo = st.text_area("노션 비고란 텍스트를 붙여넣고 [적용]을 누르세요", key=f"qmemo_{i}", placeholder="실사용자: 정영권 교수 01033246327 ykjung@sejong.ac.kr\n업체담당자: 김철수 01099998888 cs@vendor.com", height=70)
                    if st.button("✨ 폼에 즉시 반영", key=f"btn_apply_qmemo_{i}"):
                        parsed = parse_smart_contact_memo(quick_memo)
                        q_u_em = parsed.get('user_email', '')
                        q_v_em = parsed.get('vendor_email', '')
                        if is_internal_sejong_email(q_v_em):
                            if not q_u_em: q_u_em = q_v_em
                            q_v_em = ''
                        if parsed['user_name']: st.session_state[f"u_nm_{i}"] = parsed['user_name']
                        if parsed['user_phone']: st.session_state[f"u_ph_{i}"] = parsed['user_phone']
                        if q_u_em: st.session_state[f"u_em_{i}"] = q_u_em
                        if parsed['vendor_rep']: st.session_state[f"v_rep_{i}"] = parsed['vendor_rep']
                        if parsed['vendor_phone']: st.session_state[f"v_ph_{i}"] = parsed['vendor_phone']
                        st.session_state[f"v_em_{i}"] = q_v_em
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

                c4, c5, c6 = st.columns([1.2, 1, 1.5])
                with c4:
                    e_req = st.text_input("구매요구번호", value=cd.get('req_no', ''), key=f"e_req_{i}", placeholder="예: 202603120005")
                with c5:
                    delivery_options = ["납품장소도", "현장설치도", "용역", "공사", "기타"]
                    e_deliv = st.selectbox("납품방법", delivery_options, key=f"e_del_{i}")
                with c6:
                    e_loc = st.text_input("납품장소", value=cd.get('location', ''), key=f"e_loc_{i}")

                sm_data = cd.get('smart_memo', {})
                default_u_nm = st.session_state.get(f"u_nm_{i}", sm_data.get('user_name', ''))
                default_u_ph = st.session_state.get(f"u_ph_{i}", sm_data.get('user_phone', ''))
                default_u_em = st.session_state.get(f"u_em_{i}", sm_data.get('user_email', cd.get('email', '')))

                st.caption("1) 실사용자 정보 (스마트 자동 추출)")
                c_u1, c_u2, c_u3 = st.columns(3)
                with c_u1: u_nm = st.text_input("실사용자 성명", value=default_u_nm, key=f"u_nm_{i}")
                with c_u2: u_ph = st.text_input("실사용자 연락처", value=default_u_ph, key=f"u_ph_{i}")
                with c_u3: u_em = st.text_input("실사용자 이메일", value=default_u_em, key=f"u_em_{i}")

                default_v_rep = st.session_state.get(f"v_rep_{i}", sm_data.get('vendor_rep', ''))
                default_v_ph = st.session_state.get(f"v_ph_{i}", sm_data.get('vendor_phone', ''))
                # 💡 업체담당자 이메일은 실사용자 이메일(cd['email'])을 덮어쓰지 않고 부재 시 공란으로 둠
                default_v_em = st.session_state.get(f"v_em_{i}", sm_data.get('vendor_email', ''))

                st.caption("2) 협력사(업체) 담당자 정보 (이메일 미기재 시 직접 입력 가능)")
                c_vr1, c_vr2, c_vr3 = st.columns(3)
                with c_vr1: v_rep = st.text_input("업체담당자 성명", value=default_v_rep, key=f"v_rep_{i}")
                with c_vr2: v_ph = st.text_input("업체담당자 연락처", value=default_v_ph, key=f"v_ph_{i}")
                with c_vr3: v_em = st.text_input("업체담당자 이메일", value=default_v_em, key=f"v_em_{i}", placeholder="미기재 시 직접 입력")

                st.caption("3) 구매/검수 담당자 및 계약보증금")
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

                btn_gen_pat = get_ui_text("con_btn_gen", "📄 [{no}번] 워드파일 생성 준비")
                btn_dl_pat = get_ui_text("con_btn_dl", "📥 [{no}번] 완성된 발주요청서 다운로드 (.docx)")
                disp_gen_t = btn_gen_pat.replace("{no}", str(i+1))
                disp_dl_t = btn_dl_pat.replace("{no}", str(i+1))
                
                if i == 0 and is_live:
                    c_g1, c_g2 = st.columns(2)
                    with c_g1: st.text_input("📍 [위치: 발주서 생성 준비 버튼] ({no}는 번호)", value=btn_gen_pat, key="live_con_btn_gen")
                    with c_g2: st.text_input("📍 [위치: 완성본 다운로드 버튼] ({no}는 번호)", value=btn_dl_pat, key="live_con_btn_dl")

                if st.button(disp_gen_t, key=f"btn_gen_{i}"):
                    if not e_name or amt_val <= 0 or not e_ven or not u_nm:
                        st.error("🚨 필수 정보(계약건명, 계약금액, 협력사, 실사용자 성명)를 모두 입력해주세요.")
                    else:
                        try:
                            is_new_fac = "신임교원" in sel_ptype
                            new_fac_tpl = get_template_path(["template_new_faculty.docx", "신임교원용_발주서(eformsign 템플릿)_대학_국가계약법_1천만원 미만.docx"])

                            if is_new_fac and new_fac_tpl:
                                tpl_name = new_fac_tpl
                            else:
                                tpl_default = "template_over10mil.docx" if amt_val >= 10000000 else "template_under10mil.docx"
                                tpl_name = get_template_path(tpl_default) or tpl_default
                                if is_new_fac and not new_fac_tpl:
                                    st.warning(f"⚠️ 신임교원지원사업 전용 양식('template_new_faculty.docx')이 아직 없어 기본 서식({tpl_default})으로 생성되었습니다.")

                            doc = DocxTemplate(tpl_name)
                            
                            final_deposit_str = f"{calc_deposit:,}" if calc_deposit > 0 else ""
                            clean_user_for_doc = re.sub(r'\s*교수\s*$', '', u_nm).strip() if is_new_fac else u_nm
                            context = {
                                "구매담당자": b_nm, 
                                "구매담당자_연락처": b_ph, 
                                "구매담당자_메일": b_em,
                                "발주_개요": e_name, 
                                "계약금액": f"{amt_val:,}", 
                                "납품방법": e_deliv, 
                                "납품장소": e_loc,
                                "실사용자": clean_user_for_doc, 
                                "실사용자_연락처": u_ph, 
                                "실사용자_메일": u_em,
                                "협력사": e_ven,
                                "검수담당자": i_nm, 
                                "검수담당자_연락처": INSPECTOR_INFO.get("phone", "02-3408-4058"), 
                                "검수담당자_메일": INSPECTOR_INFO.get("email", "rudwpdyd@sejong.ac.kr"),
                                "계약보증금액": final_deposit_str,
                                "구매요구번호": e_req
                            }
                            doc.render(context)
                            ts = BytesIO()
                            doc.save(ts)
                            ts.seek(0)

                            clean_req = re.sub(r'[\\/*?:"<>|]', '', str(e_req).strip())
                            if clean_req:
                                target_fname = f"발주서 - {clean_req}.docx"
                            else:
                                target_fname = f"발주서 - { sanitize_win_name(e_name) }.docx" if e_name else f"발주서_{ sanitize_win_name(e_ven) }.docx"

                            docx_bytes = ts.getvalue()
                            active_base_dir = st.session_state.get("con_work_base_dir", cur_con_base_dir).strip() or cur_con_base_dir
                            if active_base_dir != APP_CONFIG.get("work_base_dir"):
                                APP_CONFIG["work_base_dir"] = active_base_dir
                                save_app_config(APP_CONFIG)

                            saved_folder_path = ""
                            saved_file_path = ""
                            is_new_folder = False
                            folder_err = ""
                            try:
                                saved_folder_path, is_new_folder = resolve_or_create_req_folder(
                                    base_dir=active_base_dir,
                                    req_no=clean_req,
                                    dept=cd.get('pi_dept', ''),
                                    pi_name=cd.get('pi_name', ''),
                                    vendor_name=e_ven,
                                    fallback_title=e_name
                                )
                                saved_file_path = os.path.join(saved_folder_path, target_fname)
                                with open(saved_file_path, "wb") as f_target:
                                    f_target.write(docx_bytes)
                            except Exception as fe:
                                folder_err = str(fe)

                            st.session_state[f"last_gen_docx_{i}"] = {
                                "bytes": docx_bytes,
                                "name": target_fname,
                                "folder_path": saved_folder_path,
                                "file_path": saved_file_path,
                                "is_new_folder": is_new_folder,
                                "folder_err": folder_err
                            }
                            st.success(f"✅ {i+1}번 발주서 생성 및 폴더 자동 배치 완료! (`{target_fname}`)")
                        except Exception as e:
                            st.error(f"🚨 워드 파일 생성 중 오류 발생: {e}")

                if f"last_gen_docx_{i}" in st.session_state:
                    saved_docx = st.session_state[f"last_gen_docx_{i}"]
                    if saved_docx.get("file_path") and os.path.exists(saved_docx["file_path"]):
                        folder_badge = "🆕 새 폴더 자동 생성됨" if saved_docx.get("is_new_folder") else "🔗 기존 구매요구번호 폴더 매칭됨"
                        st.info(f"📂 **[{folder_badge}]** `{saved_docx['file_path']}`")
                    elif saved_docx.get("folder_err"):
                        st.warning(f"⚠️ 로컬 폴더 자동 저장 실패 (다운로드 버튼을 이용하세요): {saved_docx['folder_err']}")

                    c_ef, c_op, c_dl = st.columns([1.2, 0.9, 1.2])
                    with c_ef:
                        if st.button(f"✍️ [{i+1}번] 이폼싸인(eformsign) 바로 실행", key=f"btn_eform_{i}", type="primary", width="stretch", help="폴더에 저장된 워드 문서를 열고 eformsign 모듈 버튼을 자동으로 실행합니다."):
                            target_docx_path = saved_docx.get("file_path")
                            if not target_docx_path or not os.path.exists(target_docx_path):
                                out_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "output")
                                os.makedirs(out_dir, exist_ok=True)
                                safe_name = re.sub(r'[\\/*?:"<>|]', '_', saved_docx["name"])
                                target_docx_path = os.path.join(out_dir, safe_name)
                                with open(target_docx_path, "wb") as f_out:
                                    f_out.write(saved_docx["bytes"])
                            
                            with st.spinner("⏳ 워드를 실행하고 이폼싸인 모듈을 호출하는 중입니다..."):
                                try:
                                    from local_tools.eformsign_trigger import trigger_eformsign_in_word
                                    ok = trigger_eformsign_in_word(target_docx_path, "실행")
                                    if ok:
                                        st.success("🎉 이폼싸인 모듈이 성공적으로 호출되었습니다! 워드 창에서 전자결재를 계속 진행해 주세요.")
                                    else:
                                        st.error("🚨 이폼싸인 모듈 호출 실패: 워드가 정상 실행되었는지, 상단에 eformsign 탭이 있는지 확인해주세요.")
                                except Exception as err:
                                    st.error(f"🚨 이폼싸인 실행 에러: {err}")
                    with c_op:
                        if st.button(f"📁 [{i+1}번] 저장 폴더 열기", key=f"btn_open_con_dir_{i}", width="stretch"):
                            f_to_open = saved_docx.get("folder_path")
                            if f_to_open and os.path.exists(f_to_open):
                                ok_o, err_o = open_local_folder(f_to_open)
                                if not ok_o:
                                    st.error(f"폴더 열기 실패: {err_o}")
                            else:
                                st.warning("저장된 로컬 폴더를 찾을 수 없습니다.")
                    with c_dl:
                        st.download_button(
                            label=disp_dl_t,
                            data=saved_docx["bytes"],
                            file_name=saved_docx["name"],
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            key=f"btn_dl_{i}",
                            width="stretch"
                        )

# ----------------- [3. 분석내용 불러오기] -----------------
elif menu in [m_his, "🕒 분석내역 불러오기", "분석내역 불러오기", "🕒 분석내용 불러오기", "분석내용 불러오기"]:
    render_top_header("🕒 분석내역 불러오기")
    st.info("이전에 작업하고 저장해둔 분석 내역을 엑셀 대장처럼 확인하고, 언제든 다시 불러와서 이어서 작업할 수 있습니다.")
    
    cache_list = get_cache_list()
    
    if not cache_list:
        st.warning("아직 저장된 작업 내역이 없습니다. '구매요청 분석' 탭에서 작업을 마치고 [저장] 버튼을 눌러주세요.")
    else:
        st.write("#### 📊 내장 데이터베이스 (목록 직접 수정 가능 ✏️)")
        st.info("💡 표 안의 글씨를 **더블클릭**해서 엑셀처럼 바로 수정할 수 있습니다. (고유번호인 구매요구번호는 수정 불가)\n\n수정을 마친 후 반드시 아래의 **[저장]** 버튼을 눌러주세요!")
        
        search_kw = st.text_input("🔎 내역 검색 (구매요구번호 / 과제번호 / 업체명 / 연구책임자 / 물품담당자 / 비고)", placeholder="찾으실 키워드를 입력하면 표가 즉시 필터링됩니다...", key="his_search_kw").strip()
        if search_kw:
            kw_low = search_kw.lower()
            filtered_cache_list = [
                item for item in cache_list
                if any(kw_low in str(v).lower() for v in item.values())
            ]
            st.caption(f"검색 결과: 총 **{len(filtered_cache_list)}**건 (전체 {len(cache_list)}건 중)")
        else:
            filtered_cache_list = cache_list

        edited_df = st.data_editor(
            pd.DataFrame(filtered_cache_list), 
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
                    
                clear_db_caches()
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
                ai_data = loaded_data.get("ai_res", {})
                st.session_state[SESSION_AI_KEY] = ai_data
                
                # 수기 수정했던 업체/계좌 입력창 상태 복구
                st.session_state["req_in_ven"] = ai_data.get("업체명", "")
                st.session_state["req_in_biz"] = ai_data.get("사업자번호", "")
                st.session_state["req_in_rep"] = ai_data.get("담당자", "")
                st.session_state["req_in_phn"] = ai_data.get("연락처", "")
                st.session_state["req_in_eml"] = ai_data.get("이메일", "")
                st.session_state["req_acc_holder"] = ai_data.get("예금주", "")
                st.session_state["req_acc_num"] = ai_data.get("계좌번호", "")
                
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

        st.divider()
        with st.expander("🗑️ 과거 내역 영구 삭제 (로컬 & 구글 시트 동시 삭제)", expanded=False):
            st.caption("선택한 구매요구번호의 내역을 로컬 캐시와 구글 시트 양쪽에서 완전히 삭제합니다.")
            del_targets = [f"{item['구매요구번호']} | {item.get('업체명', '업체미상')} | {item.get('구매요청액', '0')}원" for item in cache_list]
            sel_del_target = st.selectbox("삭제할 내역 선택", del_targets, key="sel_del_target")
            
            c_d1, c_d2 = st.columns([2, 1])
            with c_d2:
                if st.button("🚨 선택 내역 영구 삭제", type="primary", key="btn_delete_cache", width="stretch"):
                    target_req = sel_del_target.split(" | ")[0].strip()
                    delete_from_cache(target_req)
                    st.success(f"✅ [{target_req}] 건이 로컬 캐시 및 구글 시트에서 삭제되었습니다!")
                    time.sleep(1)
                    st.rerun()

# ----------------- [4. 관리자 설정 화면] -----------------
elif menu == "⚙️ 관리자 설정":
    if not st.session_state.get("is_admin_auth", False):
        st.error("🔒 관리자 전용 화면입니다. 우측 상단의 자물쇠(🔒) 아이콘을 클릭하여 PIN을 입력해주세요.")
        st.stop()

    render_top_header("⚙️ 관리자 설정")
    st.caption("시스템 명칭, 화면 헤드라인 문구 및 담당자 정보를 손쉽게 변경하고 관리할 수 있습니다.")
    
    # 세션에 임시 설정 버퍼가 없으면 현재 설정으로 로드
    if "admin_cfg" not in st.session_state:
        st.session_state["admin_cfg"] = json.loads(json.dumps(APP_CONFIG))
        
    admin_cfg = st.session_state["admin_cfg"]
    
    tab_titles, tab_buyer, tab_inspector = st.tabs([
        "📝 명칭 및 헤드라인", 
        "👥 구매계약 담당자 관리", 
        "🔍 검수 담당자 관리"
    ])
    
    with tab_titles:
        st.markdown("#### ✏️ 화면별 시각적 레이아웃 & 문구 설정")
        st.info("💡 **어느 위치인지 한눈에 확인하며 바꾸세요:** 각 화면의 실제 배치 순서대로 영역별로 묶어두었습니다. 원하시는 위치의 텍스트를 수정하고 맨 아래 [💾 설정 저장하기]를 누르시면 됩니다.\n\n*(더 편리하게는 실제 화면 상단에 있는 **'✏️ 화면 문구 직접 수정 모드'**를 켜서 그 자리에서 바로 고칠 수도 있습니다!)*")
        
        sub_tab_ext, sub_tab_con, sub_tab_global = st.tabs([
            "💰 1. 구매추출기 화면 레이아웃",
            "🤝 2. 구매계약관리 화면 레이아웃",
            "🌐 3. 전역 & 사이드바 & 보안"
        ])
        
        with sub_tab_ext:
            with st.container(border=True):
                st.markdown("##### 🖥️ [위치 1] 화면 최상단 영역")
                t_ext = st.text_input("화면 대제목", value=admin_cfg.get("ext_main_title", "💰 구매추출기"), key="cfg_ext_main_title")
                admin_cfg["ext_main_title"] = t_ext.strip() or "💰 구매추출기"
                
            with st.container(border=True):
                st.markdown("##### 📥 [위치 2] STEP 1 서류/텍스트 입력 카드 영역")
                t_step1 = st.text_input("1단계 큰 섹션 제목", value=admin_cfg.get("ext_step1_title", "📥 정보 입력 및 서류 업로드"), key="cfg_ext_step1_title")
                admin_cfg["ext_step1_title"] = t_step1.strip() or "📥 정보 입력 및 서류 업로드"
                
                c_s1, c_s2, c_s3 = st.columns(3)
                with c_s1:
                    t_sub1 = st.text_input("1. 기본 정보 선택 소제목", value=admin_cfg.get("ext_sub1_title", "📂 1. 기본 정보 선택"), key="cfg_ext_sub1_title")
                    admin_cfg["ext_sub1_title"] = t_sub1.strip() or "📂 1. 기본 정보 선택"
                with c_s2:
                    t_sub2 = st.text_input("2. ERP 텍스트 복사 소제목", value=admin_cfg.get("ext_sub2_title", "📝 2. ERP 텍스트 복사"), key="cfg_ext_sub2_title")
                    admin_cfg["ext_sub2_title"] = t_sub2.strip() or "📝 2. ERP 텍스트 복사"
                with c_s3:
                    t_sub3 = st.text_input("3. 첨부 서류 업로드 소제목", value=admin_cfg.get("ext_sub3_title", "📎 3. AI OCR"), key="cfg_ext_sub3_title")
                    admin_cfg["ext_sub3_title"] = t_sub3.strip() or "📎 3. AI OCR"
                    
                t_btn_ana = st.text_input("분석 시작 버튼 문구", value=admin_cfg.get("ext_btn_analyze", "🚀 전체 데이터 통합 분석 시작"), key="cfg_ext_btn_analyze")
                admin_cfg["ext_btn_analyze"] = t_btn_ana.strip() or "🚀 전체 데이터 통합 분석 시작"
                
            with st.container(border=True):
                st.markdown("##### 📤 [위치 3] STEP 2 분석 결과 확인 영역")
                t_step2 = st.text_input("2단계 큰 섹션 제목", value=admin_cfg.get("ext_step2_title", "📤 [STEP 2] 분석 결과"), key="cfg_ext_step2_title")
                admin_cfg["ext_step2_title"] = t_step2.strip() or "📤 [STEP 2] 분석 결과"
                
                c_r1, c_r2, c_r3 = st.columns(3)
                with c_r1:
                    t_ebas = st.text_input("ERP 기본 정보 소제목", value=admin_cfg.get("ext_erp_basic_title", "📝 ERP 기본 정보"), key="cfg_ext_erp_basic_title")
                    admin_cfg["ext_erp_basic_title"] = t_ebas.strip() or "📝 ERP 기본 정보"
                with c_r2:
                    t_eitm = st.text_input("ERP 품목 내역 소제목", value=admin_cfg.get("ext_erp_items_title", "📦 ERP 품목 내역"), key="cfg_ext_erp_items_title")
                    admin_cfg["ext_erp_items_title"] = t_eitm.strip() or "📦 ERP 품목 내역"
                with c_r3:
                    t_ebank = st.text_input("통장사본 계좌정보 소제목", value=admin_cfg.get("ext_bank_title", "🏦 통장사본 계좌정보"), key="cfg_ext_bank_title")
                    admin_cfg["ext_bank_title"] = t_ebank.strip() or "🏦 통장사본 계좌정보"

        with sub_tab_con:
            with st.container(border=True):
                st.markdown("##### 🖥️ [위치 1] 화면 최상단 영역")
                t_con = st.text_input("화면 대제목", value=admin_cfg.get("con_main_title", "🤝 구매계약관리"), key="cfg_con_main_title")
                t_cdesc = st.text_input("대제목 아래 안내 문구", value=admin_cfg.get("con_desc", "여러 개의 계약 건을 한 번에 입력하고 분석하여, 각각의 워드 발주서를 빠르게 생성할 수 있습니다."), key="cfg_con_desc")
                admin_cfg["con_main_title"] = t_con.strip() or "🤝 구매계약관리"
                admin_cfg["con_desc"] = t_cdesc.strip() or "여러 개의 계약 건을 한 번에 입력하고 분석하여, 각각의 워드 발주서를 빠르게 생성할 수 있습니다."
                
            with st.container(border=True):
                st.markdown("##### 📄 [위치 2] 계약 입력 카드 & 불러오기 버튼")
                t_ccard = st.text_input("계약 카드 상단 제목 ({no}는 계약 번호로 자동 치환)", value=admin_cfg.get("con_card_title", "📄 [계약 {no}] 데이터 입력"), key="cfg_con_card_title")
                t_btn = st.text_input("계약 데이터 불러오기 버튼 문구", value=admin_cfg.get("btn_load_contract", "계약 데이터 불러오기"), key="cfg_btn_title")
                admin_cfg["con_card_title"] = t_ccard.strip() or "📄 [계약 {no}] 데이터 입력"
                admin_cfg["btn_load_contract"] = t_btn.strip() or "계약 데이터 불러오기"
                
            with st.container(border=True):
                st.markdown("##### 📥 [위치 3] 발주서 생성 & 다운로드 버튼")
                c_bgen, c_bdl = st.columns(2)
                with c_bgen:
                    t_bgen = st.text_input("워드파일 생성 버튼 문구 ({no}는 번호)", value=admin_cfg.get("con_btn_gen", "📄 [{no}번] 워드파일 생성 준비"), key="cfg_con_btn_gen")
                    admin_cfg["con_btn_gen"] = t_bgen.strip() or "📄 [{no}번] 워드파일 생성 준비"
                with c_bdl:
                    t_bdl = st.text_input("워드파일 다운로드 버튼 문구 ({no}는 번호)", value=admin_cfg.get("con_btn_dl", "📥 [{no}번] 완성된 발주요청서 다운로드 (.docx)"), key="cfg_con_btn_dl")
                    admin_cfg["con_btn_dl"] = t_bdl.strip() or "📥 [{no}번] 완성된 발주요청서 다운로드 (.docx)"

        with sub_tab_global:
            with st.container(border=True):
                st.markdown("##### 🌐 브라우저 및 사이드바 설정")
                t_app = st.text_input("브라우저 탭 제목", value=admin_cfg.get("app_title", "구매 업무 보조"), key="cfg_app_title")
                t_side = st.text_input("사이드바 상단 제목", value=admin_cfg.get("sidebar_title", "💼 구매 업무 보조"), key="cfg_side_title")
                c_m1, c_m2 = st.columns(2)
                with c_m1:
                    t_mext = st.text_input("사이드바 1번 메뉴명", value=admin_cfg.get("menu_extractor", "💰 구매추출기"), key="cfg_menu_extractor")
                with c_m2:
                    t_mcon = st.text_input("사이드바 2번 메뉴명", value=admin_cfg.get("menu_contract", "🤝 구매계약관리"), key="cfg_menu_contract")
                    
                admin_cfg["app_title"] = t_app.strip() or "구매 업무 보조"
                admin_cfg["sidebar_title"] = t_side.strip() or "💼 구매 업무 보조"
                admin_cfg["menu_extractor"] = t_mext.strip() or "💰 구매추출기"
                admin_cfg["menu_contract"] = t_mcon.strip() or "🤝 구매계약관리"
                
            with st.container(border=True):
                st.markdown("##### 🔐 관리자 PIN 번호 변경")
                st.caption("우측 상단 자물쇠(🔒) 클릭 시 입력하는 관리자 비밀번호입니다. (기본값: 1234)")
                t_pin = st.text_input("새 관리자 PIN", value=str(admin_cfg.get("admin_pin", "1234")), key="cfg_admin_pin", type="password")
                admin_cfg["admin_pin"] = t_pin.strip() or "1234"
        
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
            st.session_state["nav_target"] = "💰 구매추출기"
            st.rerun()