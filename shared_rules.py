"""웹 앱(1_구매추출기.py)과 로컬 스크래퍼(scraper.py)가 공통으로 사용하는 핵심 업무 규칙 모듈."""

import datetime
import json
import os
import re

KST = datetime.timezone(datetime.timedelta(hours=9))

DEFAULT_SPECIAL_PROJECTS = {
    "박재우": {"kw": ["지역혁신중심", "RISE", "라이즈"], "suffix": "(RISE사업단)"},
    "김재호": {"kw": ["사물인터넷"], "suffix": "(사물인터넷혁신융합대학사업단)"},
    "송오영": {"kw": ["SW중심대학"], "suffix": "(SW중심대학사업단)"},
    "이기학": {"kw": ["G-LAMP", "GLAMP", "대학기초연구소"], "suffix": "(G-LAMP사업단)"},
}

SAFETY_CENTER_INFO = {
    "pi_name": "김수동 선생님",
    "pi_dept": "대학안전관리센터",
    "pi_phone": "02-3408-4081",
    "pi_email": "sudong_kim@sejong.ac.kr",
    "pi_formatted": "김수동 선생님 / 02-3408-4081 / sudong_kim@sejong.ac.kr",
}


def get_project_root():
    return os.path.dirname(os.path.abspath(__file__))


def load_shared_app_config(default_config=None):
    cfg = dict(default_config or {})
    cfg_path = os.path.join(get_project_root(), "app_config.json")
    if os.path.exists(cfg_path):
        try:
            with open(cfg_path, "r", encoding="utf-8") as f:
                data = json.load(f)
            if isinstance(data, dict):
                cfg.update(data)
        except Exception:
            pass
    return cfg


def now_kst():
    return datetime.datetime.now(KST)


def today_kst():
    return datetime.datetime.now(KST).date()


def parse_krw_amount(val):
    """소수점(.0/.00)이 섞인 금액 문자열이 10배로 부풀려지지 않도록 정제 후 정수 반환."""
    try:
        s = str(val or "").strip()
        if not s:
            return 0
        s_no_dec = re.sub(r'\.\d{1,2}(?=\s*원?\s*$)', '', s)
        digits = re.sub(r'[^\d]', '', s_no_dec)
        return int(digits) if digits else 0
    except Exception:
        return 0


def sanitize_win_name(name_str):
    """윈도우 폴더/파일명 금지 문자 제거 및 공백 정돈."""
    cleaned = re.sub(r'[\\/*?:"<>|\r\n\t]', ' ', str(name_str or '')).strip()
    return re.sub(r'\s+', ' ', cleaned)


def extract_req_no(text):
    """rERP 본문에서 구매요구번호(10~14자리)를 추출합니다."""
    if not text:
        return ""
    clean_text = ' '.join(str(text).split())
    m = (
        re.search(r'구매요구번호\s*\*?\s*[:\s]*\s*(\d{10,14})', clean_text)
        or re.search(r'구매요구번호[^\d\n]{0,15}(\d{10,14})', str(text))
    )
    return m.group(1).strip() if m else ""


def is_safety_center_case(*texts):
    """과제담당자 또는 관리기관이 대학안전관리센터인 건인지 판별합니다."""
    combined = " ".join([str(t or "") for t in texts])
    if not combined.strip():
        return False
    clean = re.sub(r'\s+', '', combined)
    if "대학안전관리센터" in clean or "대학안전관리" in clean:
        return True
    if ("과제담당자" in clean or "관리기관" in clean) and ("안전관리" in clean or "김수동" in clean):
        return True
    return False


def extract_special_project_context(raw_text):
    """과제성격 구간을 우선 추출해 특수 사업단 판별 오탐을 방지합니다."""
    clean_text = str(raw_text or "").replace(" ", "")
    m = re.search(r'과제성격(.*?)(연구책임자정보|연차연구기간)', clean_text, flags=re.DOTALL)
    return m.group(1) if m else clean_text


def check_special_project_affiliation(pi_name, pi_dept, context_text="", special_projects=None):
    """특수 사업단 소속 교수님의 학과 뒤에 (사업단명)을 자동 부착합니다."""
    if not pi_name or not pi_dept:
        return pi_dept

    spec_map = special_projects if isinstance(special_projects, dict) else DEFAULT_SPECIAL_PROJECTS
    cfg_rule = spec_map.get(pi_name.strip())
    if cfg_rule:
        suffix = cfg_rule.get("suffix", "")
        if suffix and suffix not in pi_dept:
            kws = cfg_rule.get("kw", [])
            clean_ctx = str(context_text).upper().replace(" ", "")
            if any(str(kw).upper() in clean_ctx for kw in kws):
                return f"{pi_dept}{suffix}"
    return pi_dept


def build_recommended_folder_name(dept, pi_name, vendor_name, req_no="", date_obj=None, date_str=None):
    """추천 폴더명 생성: (YYYY.MM.DD) 학과 성명 - 업체명 (구매요구번호)."""
    if date_str:
        d_str = str(date_str).strip()
    else:
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


def get_accurate_total(text, keyword=""):
    """rERP 텍스트에서 날짜/규격 숫자를 걸러내고 정확한 금액 최대값을 추출합니다."""
    search_area = str(text or "")
    if keyword:
        search_area = search_area.split(keyword)[-1]
    search_area = re.sub(r'\b\d+(?:-\d+)+\b', '', search_area)
    search_area = re.sub(r'\b\d[\d,]*\s*(mm|cm|m|kg|g|mg|ml|l|v|hz|w|ea|u|%)\b', '', search_area, flags=re.IGNORECASE)
    search_area = re.sub(r'\b[a-zA-Z]+\d+|\d+[a-zA-Z]+\b', '', search_area)
    all_numbers = re.findall(r'\b\d[\d,]*\b', search_area)

    clean_numbers = []
    for n in all_numbers:
        num_str = n.replace(',', '')
        val = parse_krw_amount(num_str)
        if num_str.startswith('0') and len(num_str) > 1:
            continue
        if len(num_str) >= 11:
            continue
        if len(num_str) == 8 and (num_str.startswith('19') or num_str.startswith('20')):
            month = parse_krw_amount(num_str[4:6])
            day = parse_krw_amount(num_str[6:8])
            if 1 <= month <= 12 and 1 <= day <= 31:
                continue
        if val >= 1000 or val == 0:
            clean_numbers.append(val)
    return max(clean_numbers) if clean_numbers else 0


def resolve_request_total_amount(raw_req, raw_fin="", item_list=None):
    """구매요청 화면 표시 및 DB 저장에 공통 적용되는 총액 추출 (합계액 -> 총액 -> 품목합산)."""
    t_amt = get_accurate_total(raw_req, "합계액")
    if t_amt == 0 and raw_fin:
        t_amt = get_accurate_total(raw_fin, "총액")
    if t_amt == 0 and item_list:
        t_amt = sum(parse_krw_amount(itm.get('price', 0)) for itm in item_list)
    return t_amt

