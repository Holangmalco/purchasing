# %%
import os
import sys
import time
import zipfile
import re
import shutil
import json
import traceback
import threading
import queue
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

if sys.platform.startswith('win'):
    try:
        sys.stdout.reconfigure(encoding='utf-8')
        sys.stderr.reconfigure(encoding='utf-8')
    except Exception:
        pass

from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.options import Options as ChromeOptions
from selenium.webdriver.chrome.service import Service as ChromeService

from google import genai
from google.genai import types

# =================================================================
# 0. GUI 및 로그 큐 헬퍼
# =================================================================
stop_requested = False
gui_status_cb = None

class QueueStdout:
    def __init__(self, q, orig=None):
        self.q = q
        self.orig = orig

    def write(self, s):
        if self.orig:
            try:
                self.orig.write(s)
            except Exception:
                pass
        if s:
            self.q.put(s)

    def flush(self):
        if self.orig:
            try:
                self.orig.flush()
            except Exception:
                pass

# =================================================================
# 1. 시스템 및 API 키 세팅 (.env)
# =================================================================
load_dotenv()

# 상위 폴더의 .env까지 뒤져서 환경변수 확인
dev_env = os.path.join(r"C:\PurchasingExtractor_Dev", ".env")
if os.path.exists(dev_env):
    load_dotenv(dev_env)

RERP_ID = os.getenv("RERP_ID", "")
RERP_PW = os.getenv("RERP_PW", "")
DEFAULT_API_KEY = ""
GEMINI_KEY = os.getenv("GEMINI_API_KEY") or os.getenv("GOOGLE_API_KEY") or ""

def check_credentials(interactive=True):
    global RERP_ID, RERP_PW, GEMINI_KEY, ai_client
    if not RERP_ID or RERP_ID == "본인_아이디_입력":
        if not interactive:
            return False
        print("\n" + "="*50)
        print("  [최초 실행 설정] rERP 로그인 정보가 필요합니다.")
        print("="*50)
        RERP_ID = input("👉 rERP 아이디(학번/사번)를 입력하세요: ").strip()
        RERP_PW = input("👉 rERP 비밀번호를 입력하세요: ").strip()
        try:
            with open(".env", "a", encoding="utf-8") as ef:
                ef.write(f"\nRERP_ID={RERP_ID}\nRERP_PW={RERP_PW}\n")
            print("[+] 계정 정보가 저장되었습니다.\n")
        except: pass

    if not GEMINI_KEY and not ai_client:
        if not interactive:
            return False
        print("\n" + "="*50)
        print("  [AI 분석 설정] Gemini API Key를 입력해주세요.")
        print("="*50)
        user_key = input("👉 Gemini API Key 입력 (Enter시 AI 자동분석 건너뜀): ").strip()
        if user_key:
            GEMINI_KEY = user_key
            try:
                with open(".env", "a", encoding="utf-8") as ef:
                    ef.write(f"\nGEMINI_API_KEY={GEMINI_KEY}\n")
                print("[+] API Key가 저장되었습니다.\n")
            except: pass
            ai_client = genai.Client(api_key=GEMINI_KEY)

    global BASE_TARGET_DIR, TEMP_DL_DIR
    if not os.getenv("TARGET_DIR") and interactive:
        print("\n" + "="*50)
        print("  [저장 폴더 설정]")
        print(f"  기본 저장 경로: {BASE_TARGET_DIR}")
        print("="*50)
        user_dir = input("👉 저장할 폴더 경로를 입력하세요 (기본값 사용 시 Enter): ").strip('\'" ')
        if user_dir:
            try:
                os.makedirs(user_dir, exist_ok=True)
                BASE_TARGET_DIR = user_dir
                TEMP_DL_DIR = os.path.join(BASE_TARGET_DIR, "_temp_dl")
                os.makedirs(TEMP_DL_DIR, exist_ok=True)
                with open(".env", "a", encoding="utf-8") as ef:
                    ef.write(f"\nTARGET_DIR={BASE_TARGET_DIR}\n")
                print(f"[+] 저장 경로가 설정되었습니다: {BASE_TARGET_DIR}\n")
            except Exception as e:
                print(f"[!] 폴더 생성 오류 ({e}), 기본 경로를 사용합니다.\n")
        else:
            try:
                with open(".env", "a", encoding="utf-8") as ef:
                    ef.write(f"\nTARGET_DIR={BASE_TARGET_DIR}\n")
            except: pass
    return True

ai_client = genai.Client(api_key=GEMINI_KEY) if GEMINI_KEY else None

# =================================================================
# 2. 지정된 업무 기본 저장 경로 세팅 (사용자 설정 / 환경변수 반영)
# =================================================================
env_target = os.getenv("TARGET_DIR", "").strip()
desktop_dir = os.path.join(os.path.expanduser("~"), "Desktop")
default_dir = os.path.join(desktop_dir, "★업무", "1. 2026학년도 구매요청")

if env_target:
    BASE_TARGET_DIR = env_target
else:
    BASE_TARGET_DIR = default_dir

try:
    os.makedirs(BASE_TARGET_DIR, exist_ok=True)
except Exception:
    BASE_TARGET_DIR = os.path.join(desktop_dir, "구매요청_수집")
    os.makedirs(BASE_TARGET_DIR, exist_ok=True)

TEMP_DL_DIR = os.path.join(BASE_TARGET_DIR, "_temp_dl")
os.makedirs(TEMP_DL_DIR, exist_ok=True)

driver = None 

# =================================================================
# 💡 [유틸] 안전 파일명 & ZIP 한글 깨짐 방지 압축 해제
# =================================================================
def sanitize_filename(name):
    clean = re.sub(r'[\\/*?:"<>|]', '_', str(name)).strip()
    return clean[:60]

def extract_zip_safe(zip_path, target_dir):
    os.makedirs(target_dir, exist_ok=True)
    with zipfile.ZipFile(zip_path, 'r') as zf:
        for member in zf.infolist():
            filename = member.filename
            try:
                filename = member.filename.encode('cp437').decode('cp949')
            except Exception:
                try:
                    filename = member.filename.encode('cp437').decode('euc-kr')
                except Exception:
                    filename = member.filename
            
            clean_name = os.path.basename(filename)
            if not clean_name: continue
            
            dest_path = os.path.join(target_dir, clean_name)
            if not member.is_dir():
                with zf.open(member) as source, open(dest_path, "wb") as target:
                    shutil.copyfileobj(source, target)
    print(f"  [+] 첨부 서류 압축 해제 완료: {target_dir}")

# =================================================================
# 🤖 [이식 완료] Gemini 3.8-Flash 구매추출기 AI 분석 엔진
# =================================================================
def analyze_folder_documents(folder_path):
    if not ai_client:
        print("  [!] GEMINI_API_KEY를 찾지 못해 AI 자동 분석을 건너뜁니다.")
        return None

    # 폴더 내 분석 가능한 문서(PDF, 이미지) 수집
    supported_exts = ('.pdf', '.png', '.jpg', '.jpeg')
    doc_files = [
        os.path.join(folder_path, f) for f in os.listdir(folder_path)
        if f.lower().endswith(supported_exts)
    ]
    
    if not doc_files:
        print("  [!] 분석할 PDF/이미지 견적서 서류가 폴더에 없습니다.")
        return None

    print(f"  [*] Gemini 3.8-Flash 서류 분석 가동 중... (서류 {len(doc_files)}개 전송)")
    today_str = time.strftime('%Y-%m-%d')

    contents = [
        f"🚨 [기준일] 오늘 날짜는 {today_str}입니다. 만료일자 계산 시 참고하세요.\n"
        "분석 후 반드시 아래 형태의 JSON 형식으로만 응답해. 절대 마크다운(```json)이나 다른 텍스트를 넣지 마.\n"
        "{\n"
        '  "비교견적리스트": [ {"업체명": "A업체", "사업자번호": "123-45-67890", "제출금액": 10000} ],\n'
        '  "최저가판단과정": "가장 금액이 낮은 업체를 최종 선정한 과정 요약",\n'
        '  "수신자명확인": "모든 견적서 수신자가 \'세종대학교 산학협력단\' 명의가 맞으면 \'예\', 아니면 사유",\n'
        '  "만료일자": "유효기간 만료일자를 YYYY-MM-DD 형식으로 기재",\n'
        '  "동일규격확인": "복수 견적서 간 동일 물품이면 \'예\', 다르면 사유, 1장이면 \'단일견적\'",\n'
        '  "사업자일치여부": "최저가 견적서와 사업자등록증 상 정보 일치 여부",\n'
        '  "업체명": "첨부된 사업자등록증 상의 \'법인명(단체명)\' 또는 \'상호\'를 최우선으로 추출 (사업자등록증이 없는 경우 견적서의 상호명)",\n'
        '  "사업자번호": "최저가 업체의 사업자등록번호",\n'
        '  "대표자명": "대표자 성명",\n'
        '  "담당자": "견적 담당자명",\n'
        '  "연락처": "담당자 연락처",\n'
        '  "이메일": "이메일",\n'
        '  "총금액": 0,\n'
        '  "품목리스트": [ {"품명": "품명", "수량": 1, "단가": 1000, "금액": 1000} ]\n'
        "}"
    ]

    for fpath in doc_files:
        try:
            with open(fpath, "rb") as f:
                data = f.read()
            mime = "application/pdf" if fpath.lower().endswith(".pdf") else "image/jpeg"
            contents.append(types.Part.from_bytes(data=data, mime_type=mime))
        except Exception as e:
            print(f"  [!] 파일 읽기 실패: {os.path.basename(fpath)} ({e})")

    gen_config = types.GenerateContentConfig(response_mime_type="application/json")
    
    # 3단계 폴백 체인 (3.5-flash-lite -> 3.8-flash -> 2.5-flash)
    for model_name in ["gemini-3.5-flash-lite", "gemini-3.8-flash", "gemini-2.5-flash"]:
        try:
            resp = ai_client.models.generate_content(model=model_name, contents=contents, config=gen_config)
            text = (resp.text or "").strip()
            # JSON 파싱
            data = json.loads(text)
            extracted_vendor = data.get("업체명", "").strip()
            print(f"  🎉 [AI 분석 성공 ({model_name})] 최저가 견적 업체 추출: [{extracted_vendor}]")
            return data
        except Exception:
            continue

    print("  [!] AI 분석 실패 (JSON 파싱 또는 API 오류)")
    return None

# =================================================================
# 📝 노션(Notion) 복사용 마크다운 요약본 생성기
# =================================================================
def create_notion_summary(ai_data, erp_info):
    lines = []
    lines.append(f"> 💡 **[구매요청 검토 완료]** {erp_info.get('date')} / 신청: {erp_info.get('author')} / 책임자: {erp_info.get('pi_dept')} {erp_info.get('pi_name')}")
    lines.append("")
    lines.append("### 📌 기본 정보")
    lines.append(f"- **문서 제목:** {erp_info.get('subject')}")
    lines.append(f"- **문서 번호:** {erp_info.get('doc_no')}")
    
    vendor = ai_data.get('업체명', '확인필요') if ai_data else '서류확인필요'
    biz_no = ai_data.get('사업자번호', '-') if ai_data else '-'
    amount = ai_data.get('총금액', 0) if ai_data else 0
    
    lines.append(f"- **최저가 선정업체:** **{vendor}** (사업자번호: {biz_no})")
    try:
        amt_str = f"{int(amount):,}원"
    except:
        amt_str = f"{amount}원"
    lines.append(f"- **최종 견적 총액:** {amt_str} (VAT 포함)")
    
    if ai_data:
        lines.append(f"- **업체 담당자:** {ai_data.get('담당자', '-')} ({ai_data.get('연락처', '-')}, {ai_data.get('이메일', '-')})")
        lines.append("")
        lines.append("### 📋 품목 내역")
        lines.append("| No | 품명 | 수량 | 단가 | 금액 |")
        lines.append("|:---:|:---|:---:|---:|---:|")
        for i, itm in enumerate(ai_data.get("품목리스트", []), 1):
            p_val = itm.get('단가', 0)
            a_val = itm.get('금액', 0)
            try: p_str = f"{int(p_val):,}"
            except: p_str = str(p_val)
            try: a_str = f"{int(a_val):,}"
            except: a_str = str(a_val)
            lines.append(f"| {i} | {itm.get('품명', '-')} | {itm.get('수량', 1)} | {p_str} | {a_str} |")
        lines.append("")
        lines.append("### 🔍 서류 교차검증 결과")
        lines.append(f"* **수신자 명의 확인:** {ai_data.get('수신자명확인', '-')}")
        lines.append(f"* **규격 동일 여부:** {ai_data.get('동일규격확인', '-')}")
        lines.append(f"* **사업자 일치 여부:** {ai_data.get('사업자일치여부', '-')}")
        lines.append(f"* **최저가 판단 과정:** {ai_data.get('최저가판단과정', '-')}")
        lines.append(f"* **견적 유효기간 만료일:** {ai_data.get('만료일자', '-')}")
    
    return "\n".join(lines)

# =================================================================
# 🔘 [기능 1] 로그인 및 결재대기함 열기
# =================================================================
def start_login(interactive=True):
    global driver
    check_credentials(interactive=interactive)
    try:
        # 기존 임시 다운로드 폴더 청소
        for f in os.listdir(TEMP_DL_DIR):
            try: os.remove(os.path.join(TEMP_DL_DIR, f))
            except: pass

        options = webdriver.ChromeOptions()
        options.add_experimental_option("detach", True)
        options.add_experimental_option("excludeSwitches", ["enable-automation"])
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--disable-popup-blocking")

        prefs = {
            "credentials_enable_service": False,
            "profile.password_manager_enabled": False,
            "download.default_directory": os.path.abspath(TEMP_DL_DIR),
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True
        }
        options.add_experimental_option("prefs", prefs)

        print("[*] 크롬 브라우저를 가동합니다...")
        driver = webdriver.Chrome(options=options)
        driver.maximize_window()

        ERP_URL = "https://iacf.sejong.ac.kr/main_0001_08.act" 
        driver.get(ERP_URL)

        wait = WebDriverWait(driver, 10)
        id_box = wait.until(EC.presence_of_element_located((By.ID, "USER_NM")))
        id_box.clear()
        id_box.send_keys(RERP_ID)
        
        pw_box = driver.find_element(By.ID, "USER_PW")
        pw_box.clear()
        pw_box.send_keys(RERP_PW)
        
        login_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), '로그인')]")))
        login_btn.click()
        print("[*] 로그인 버튼 클릭 완료!")

        # 팝업 방어 1: 새 창 팝업 닫기
        try:
            WebDriverWait(driver, 4).until(EC.number_of_windows_to_be(2))
            handles = driver.window_handles
            driver.switch_to.window(handles[1]) 
            driver.close()                      
            driver.switch_to.window(handles[0]) 
            print("[+] 새 창 팝업 닫기 완료.")
        except TimeoutException:
            pass

        # 팝업 방어 2: 레이어 팝업 닫기
        driver.switch_to.default_content() 
        try:
            close_btn = WebDriverWait(driver, 2).until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="unified_popup_layer"]/div/div/div[1]/button'))
            )
            close_btn.click()
            print("[+] 레이어 팝업 닫기 완료.")
        except TimeoutException:
            iframes = driver.find_elements(By.TAG_NAME, "iframe")
            for iframe in iframes:
                try:
                    driver.switch_to.default_content()
                    driver.switch_to.frame(iframe) 
                    close_btn = WebDriverWait(driver, 1).until(
                        EC.element_to_be_clickable((By.XPATH, '//*[@id="unified_popup_layer"]/div/div/div[1]/button'))
                    )
                    close_btn.click()
                    print("[+] iframe 내부 레이어 팝업 닫기 완료.")
                    break
                except Exception:
                    pass

        # 결재대기함 버튼 클릭
        driver.switch_to.default_content() 
        iframes = driver.find_elements(By.TAG_NAME, "iframe")
        clicked = False
        
        for idx, iframe in enumerate(iframes):
            try:
                driver.switch_to.default_content()
                driver.switch_to.frame(iframe) 
                wait_box = WebDriverWait(driver, 1.5).until(
                    EC.element_to_be_clickable((By.XPATH, '//*[@id="content"]/div/div[2]/div[2]/div[2]/div[1]/table/tbody/tr/td[1]'))
                )
                wait_box.click()
                print(f"[+] {idx}번 iframe 안에서 '결재대기함' 클릭 성공!")
                clicked = True
                break 
            except Exception:
                pass 
        
        driver.switch_to.default_content() 
        if clicked:
            time.sleep(1.5)
            print("[*] 결재대기함 화면 진입 완료!")
        else:
            print("[!] 결재대기함 버튼을 찾지 못했습니다.")

    except Exception as e:
        print(f"[!] 로그인 중 오류: {e}")

# =================================================================
# 🔘 [기능 2] 상세 창 다운로드 + AI 분석 + 최종 폴더명 변경
# =================================================================
def process_current_detail_window(default_folder_hint="구매요청"):
    global driver
    try:
        # 화면 맨 아래로 스크롤하여 모든 요소 렌더링
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(0.5)

        # 1. 화면 전체 텍스트 수집
        body_text = driver.find_element(By.TAG_NAME, "body").text

        # 2. ERP 본문 정보 파싱 (일자, 학과, 연구책임자성명, 기안자, 건명)
        doc_no = ""
        doc_match = re.search(r'문서번호\s*[:\s]*([0-9A-Za-z-]+)', body_text)
        if doc_match: doc_no = doc_match.group(1).strip()

        author = ""
        author_match = re.search(r'기안자\s*[:\s]*(\d{4}[-.]\d{2}[-.]\d{2})?\s*([가-힣]{2,4})', body_text)
        if author_match: author = author_match.group(2).strip()

        # 날짜 추출 (YYYY.MM.DD 포맷으로 변환)
        raw_date = time.strftime('%Y.%m.%d')
        date_match = re.search(r'(\d{4})[-.](\d{2})[-.](\d{2})', body_text)
        if date_match:
            raw_date = f"{date_match.group(1)}.{date_match.group(2)}.{date_match.group(3)}"

        # 연구책임자 및 학과 추출: "연구책임자정보 김성한 / 부교수 / AI로봇학과 ..."
        pi_name = ""
        pi_dept = ""
        pi_match = re.search(r'연구책임자정보\s*(.*?)(?:연차연구기간|\n)', body_text, re.DOTALL)
        if pi_match:
            parts = [p.strip() for p in pi_match.group(1).split('/') if p.strip()]
            if len(parts) >= 1: pi_name = parts[0]
            if len(parts) >= 3: pi_dept = parts[2]
            elif len(parts) >= 2: pi_dept = parts[1]

        # 건명 추출
        subject = default_folder_hint
        subj_match = re.search(r'대표건명\s*[:\s*]*(.*?)(?:\n|보완사항|진행상태)', body_text)
        if subj_match and subj_match.group(1).strip():
            subject = subj_match.group(1).strip()

        erp_info = {
            "date": raw_date,
            "pi_dept": pi_dept if pi_dept else "소속미상",
            "pi_name": pi_name if pi_name else (author if author else "신청자"),
            "author": author if author else "기안자",
            "subject": subject,
            "doc_no": doc_no
        }

        # 3. 임시 1차 폴더 생성: (2026.10.14) 학과 연구책임자 - [분석중]
        temp_folder_name = f"({erp_info['date']}) {erp_info['pi_dept']} {erp_info['pi_name']} - [분석중]"
        temp_dir = os.path.join(BASE_TARGET_DIR, temp_folder_name)
        os.makedirs(temp_dir, exist_ok=True)
        print(f"  [*] 1단계 임시 폴더 생성: {temp_folder_name}")

        # 본문 텍스트 저장 (erp_text.txt)
        txt_path = os.path.join(temp_dir, "erp_text.txt")
        with open(txt_path, "w", encoding="utf-8") as f:
            f.write(body_text)

        # 4. '전체파일 다운로드' 버튼 클릭
        dl_btns = driver.find_elements(By.XPATH, "//*[contains(text(), '전체파일 다운로드')]")
        if not dl_btns:
            iframes = driver.find_elements(By.TAG_NAME, "iframe")
            for ifr in iframes:
                try:
                    driver.switch_to.frame(ifr)
                    dl_btns = driver.find_elements(By.XPATH, "//*[contains(text(), '전체파일 다운로드')]")
                    if dl_btns: break
                    driver.switch_to.default_content()
                except:
                    driver.switch_to.default_content()

        if dl_btns:
            btn = dl_btns[0]
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn)
            time.sleep(0.3)
            
            for f in os.listdir(TEMP_DL_DIR):
                try: os.remove(os.path.join(TEMP_DL_DIR, f))
                except: pass

            print("  [*] '전체파일 다운로드' 클릭...")
            try: btn.click()
            except: driver.execute_script("arguments[0].click();", btn)

            # 다운로드 완료 대기 (최대 15초)
            print("  [*] ZIP 파일 다운로드 대기 중...")
            downloaded_zip = None
            for _ in range(30):
                time.sleep(0.5)
                files = os.listdir(TEMP_DL_DIR)
                zips = [f for f in files if f.endswith('.zip')]
                crdownloads = [f for f in files if '.crdownload' in f]
                if zips and not crdownloads:
                    downloaded_zip = os.path.join(TEMP_DL_DIR, zips[0])
                    break

            if downloaded_zip and os.path.exists(downloaded_zip):
                extract_zip_safe(downloaded_zip, temp_dir)
                try: os.remove(downloaded_zip)
                except: pass
            else:
                print("  [!] ZIP 다운로드 대기 시간 초과 또는 파일 없음")
        else:
            print("  [!] '전체파일 다운로드' 버튼을 찾지 못했습니다.")

        # 5. [핵심] Gemini 3.8-Flash AI 구매추출기 분석 가동!
        ai_data = analyze_folder_documents(temp_dir)
        
        vendor_name = "최저가업체확인요망"
        if ai_data and ai_data.get("업체명"):
            vendor_name = sanitize_filename(ai_data.get("업체명"))

        # 6. 노션 복사용 마크다운 요약본 저장
        notion_md = create_notion_summary(ai_data, erp_info)
        with open(os.path.join(temp_dir, "노션_복사용.txt"), "w", encoding="utf-8") as nf:
            nf.write(notion_md)
        print("  [+] 노션 복사용 요약문 (노션_복사용.txt) 생성 완료")

        # 7. 최종 폴더명 변경 (규칙: (년.월.일) 학과 연구책임자성명 - 최저가업체명)
        final_folder_name = f"({erp_info['date']}) {erp_info['pi_dept']} {erp_info['pi_name']} - {vendor_name}"
        final_dir = os.path.join(BASE_TARGET_DIR, final_folder_name)
        
        # 중복 폴더명 방어
        if temp_dir != final_dir:
            if os.path.exists(final_dir):
                final_dir = os.path.join(BASE_TARGET_DIR, f"{final_folder_name}_{int(time.time())%1000}")
            try:
                shutil.move(temp_dir, final_dir)
                print(f"  ✨ [최종 폴더명 완성] ➔ {os.path.basename(final_dir)}")
            except Exception as ren_err:
                print(f"  [!] 폴더명 변경 실패: {ren_err} (임시 폴더로 유지)")

        return True

    except Exception as e:
        print(f"  [!] 상세 페이지 처리 중 오류: {e}")
        return False

# =================================================================
# 🔘 [기능 3] 대기함의 구매요청 건 자동 순회 (단일 또는 전체)
# =================================================================
def process_purchase_requests(single_only=False):
    global driver
    if driver is None:
        print("[!] 브라우저가 꺼져 있습니다. 1번(로그인)을 먼저 실행하세요.")
        return

    main_window = driver.window_handles[0]
    driver.switch_to.window(main_window)
    driver.switch_to.default_content()

    # 결재대기함이 들어있는 iframe 찾기
    target_iframe = None
    iframes = driver.find_elements(By.TAG_NAME, "iframe")
    for iframe in iframes:
        try:
            driver.switch_to.default_content()
            driver.switch_to.frame(iframe)
            if driver.find_elements(By.XPATH, "//td[contains(@class, 'HideCol0PROC_TYP_NM')]"):
                target_iframe = iframe
                break
        except:
            pass

    if not target_iframe:
        print("[!] 결재대기함 표가 있는 프레임을 찾지 못했습니다.")
        return

    # 대기함 목록 파악
    driver.switch_to.default_content()
    driver.switch_to.frame(target_iframe)
    proc_cells = driver.find_elements(By.XPATH, "//td[contains(@class, 'HideCol0PROC_TYP_NM')]")

    # '구매요청' 행 번호(0-based) 수집
    purchase_indices = []
    for idx, cell in enumerate(proc_cells):
        if "구매요청" in cell.text.strip():
            purchase_indices.append(idx)

    total_count = len(purchase_indices)
    print(f"\n==================================================")
    print(f"📋 대기함에서 총 {total_count}건의 '구매요청' 문서를 발견했습니다!")
    print(f"==================================================")

    if total_count == 0:
        return

    global stop_requested, gui_status_cb
    stop_requested = False

    process_targets = purchase_indices[:1] if single_only else purchase_indices

    for loop_cnt, row_idx in enumerate(process_targets, 1):
        if stop_requested:
            print("\n[🛑] 사용자에 의해 작업이 중단되었습니다.")
            if gui_status_cb:
                gui_status_cb("작업 중단됨")
            break

        if gui_status_cb:
            gui_status_cb(f"[{loop_cnt}/{len(process_targets)}] 문서 처리 중...")

        print(f"\n▶ [{loop_cnt}/{len(process_targets)}] 번째 구매요청 문서 처리 시작 (행 번호: {row_idx + 1})...")

        # 1. 메인 창 및 프레임 재진입
        driver.switch_to.window(main_window)
        driver.switch_to.default_content()
        driver.switch_to.frame(target_iframe)

        # 2. 행 재검색
        current_proc_cells = driver.find_elements(By.XPATH, "//td[contains(@class, 'HideCol0PROC_TYP_NM')]")
        if row_idx >= len(current_proc_cells):
            print(f"[!] {row_idx}번 행을 찾을 수 없습니다. 목록 갱신으로 스킵합니다.")
            continue

        target_proc_td = current_proc_cells[row_idx]
        parent_row = target_proc_td.find_element(By.XPATH, "./..")
        title_td = parent_row.find_element(By.XPATH, ".//td[contains(@class, 'HideCol0PRJ_NM')]")
        row_title = title_td.text.strip()
        print(f"  [*] 클릭 대상 과제명: {row_title}")

        # 3. 새 창 열기 (마우스 물리 클릭)
        before_handles = list(driver.window_handles)
        driver.execute_script("arguments[0].scrollIntoView({block: 'center', inline: 'center'});", title_td)
        time.sleep(0.3)

        try:
            actions = ActionChains(driver)
            actions.move_to_element(title_td).pause(0.2).click().perform()
        except:
            driver.execute_script("arguments[0].click();", title_td)

        # 새 창 대기
        opened = False
        try:
            WebDriverWait(driver, 4).until(lambda d: len(d.window_handles) > len(before_handles))
            opened = True
        except TimeoutException:
            js_event = """
            var elem = arguments[0];
            ['mouseover', 'mousedown', 'mouseup', 'click'].forEach(function(e) {
                elem.dispatchEvent(new MouseEvent(e, {view: window, bubbles: true, cancelable: true, buttons: 1}));
            });
            """
            driver.execute_script(js_event, title_td)
            try:
                WebDriverWait(driver, 3).until(lambda d: len(d.window_handles) > len(before_handles))
                opened = True
            except:
                pass

        if not opened:
            print("  [!] 새 창 열기에 실패하여 다음 문서로 넘어갑니다.")
            continue

        # 4. 새 창으로 전환
        new_handles = list(driver.window_handles)
        detail_window = [h for h in new_handles if h != main_window][-1]
        driver.switch_to.window(detail_window)
        print(f"  [+] 새 창 전환 성공: [{driver.title}]")

        # 5. 본문 텍스트 저장 & 첨부파일 ZIP 다운로드 & AI 분석 & 최종 폴더명 변경
        time.sleep(1)
        process_current_detail_window(default_folder_hint=row_title)

        # 6. 상세 창 닫기 및 메인 창 복귀
        try:
            driver.close()
            print("  [+] 상세 창 닫기 완료.")
        except:
            pass

        driver.switch_to.window(main_window)
        time.sleep(1)

    print(f"\n🎉 모든 구매요청 문서({len(process_targets)}건) 처리 및 폴더링 완료!\n저장 경로: {BASE_TARGET_DIR}")

# =================================================================
# 🔘 [기능 4] 1번째 구매요청 건 열기만 하기 (확인용)
# =================================================================
def open_first_request_only():
    global driver
    if driver is None:
        print("[!] 브라우저가 꺼져 있습니다. 1번(로그인)을 먼저 실행하세요.")
        return

    main_window = driver.window_handles[0]
    driver.switch_to.window(main_window)
    driver.switch_to.default_content()

    target_iframe = None
    for iframe in driver.find_elements(By.TAG_NAME, "iframe"):
        try:
            driver.switch_to.default_content()
            driver.switch_to.frame(iframe)
            if driver.find_elements(By.XPATH, "//td[contains(@class, 'HideCol0PROC_TYP_NM')]"):
                target_iframe = iframe
                break
        except:
            pass

    if not target_iframe:
        print("[!] 결재대기함 목록을 찾지 못했습니다.")
        return

    driver.switch_to.default_content()
    driver.switch_to.frame(target_iframe)
    proc_cells = driver.find_elements(By.XPATH, "//td[contains(@class, 'HideCol0PROC_TYP_NM')]")

    for cell in proc_cells:
        if "구매요청" in cell.text.strip():
            parent_row = cell.find_element(By.XPATH, "./..")
            title_td = parent_row.find_element(By.XPATH, ".//td[contains(@class, 'HideCol0PRJ_NM')]")
            print(f"[*] 1번째 구매요청 건명: {title_td.text.strip()}")
            
            before_handles = list(driver.window_handles)
            driver.execute_script("arguments[0].scrollIntoView({block: 'center', inline: 'center'});", title_td)
            time.sleep(0.3)
            
            try:
                actions = ActionChains(driver)
                actions.move_to_element(title_td).pause(0.2).click().perform()
            except:
                driver.execute_script("arguments[0].click();", title_td)

            try:
                WebDriverWait(driver, 4).until(lambda d: len(d.window_handles) > len(before_handles))
                new_handles = list(driver.window_handles)
                detail_window = [h for h in new_handles if h != main_window][-1]
                driver.switch_to.window(detail_window)
                print(f"🎉 새 창 열림 성공! 창 제목: [{driver.title}]")
                print("💡 창을 닫지 않고 유지합니다. [3번] 메뉴를 누르면 이 창에 대해 다운로드/AI분석을 수행합니다.")
            except:
                print("[!] 새 창이 열리지 않았습니다.")
            return

    print("[!] 결재대기함에 '구매요청' 건이 없습니다.")

# =================================================================
# 📁 [설정] 저장 폴더 변경
# =================================================================
def change_target_directory():
    global BASE_TARGET_DIR, TEMP_DL_DIR
    print("\n" + "="*52)
    print(f" 현재 저장 경로: {BASE_TARGET_DIR}")
    print("="*52)
    new_path = input("👉 새로 지정할 폴더 경로를 입력하세요 (취소는 Enter): ").strip('\'" ')
    if new_path:
        try:
            os.makedirs(new_path, exist_ok=True)
            BASE_TARGET_DIR = new_path
            TEMP_DL_DIR = os.path.join(BASE_TARGET_DIR, "_temp_dl")
            os.makedirs(TEMP_DL_DIR, exist_ok=True)
            
            env_lines = []
            if os.path.exists(".env"):
                with open(".env", "r", encoding="utf-8") as ef:
                    env_lines = ef.readlines()
            new_lines = [l for l in env_lines if not l.startswith("TARGET_DIR=")]
            new_lines.append(f"TARGET_DIR={BASE_TARGET_DIR}\n")
            with open(".env", "w", encoding="utf-8") as ef:
                ef.writelines(new_lines)
            print(f"[+] 저장 경로가 성공적으로 변경되었습니다:\n    {BASE_TARGET_DIR}\n")
        except Exception as e:
            print(f"[!] 폴더 경로 생성 실패: {e}\n")
    else:
        print("[*] 경로 변경을 취소했습니다.\n")

# =================================================================
# 🚀 인터랙티브 메뉴 실행기 (더블클릭 / 터미널용)
# =================================================================
def main_menu():
    global driver
    while True:
        print("\n" + "="*52)
        print("   산학협력단 rERP 구매요청 자동 스크래퍼 + AI")
        print(f"   [저장 경로] {BASE_TARGET_DIR}")
        print("="*52)
        print("  --- [구간별 수동/검증 모드] ---")
        print(" [1] 🔑 1단계: 로그인 및 결재대기함 열기")
        print(" [2] 📑 2단계: 첫 번째 구매요청 문서 열기만 하기 (화면 확인용)")
        print(" [3] 🔍 3단계: 현재 열려있는 상세 창 다운로드 + AI분석 + 폴더링")
        print("  --- [원클릭 자동 모드] ---")
        print(" [4] 🎯 [1건 원클릭] 1번째 건 열기부터 AI분석/폴더링까지 한 번에")
        print(" [5] 🚀 [전체 자동] 대기함의 모든 구매요청 건 순차 일괄 처리!")
        print("  --- [설정 관리] ---")
        print(" [6] 📁 저장 폴더 경로 변경")
        print(" [0] ❌ 프로그램 및 브라우저 종료")
        print("="*52)
        
        choice = input(">> 번호를 입력하고 [Enter]를 누르세요 (기본 5): ").strip()
        
        if choice == "1":
            start_login()
        elif choice == "2":
            if driver is None:
                start_login()
                time.sleep(2)
            open_first_request_only()
        elif choice == "3":
            if driver is None or len(driver.window_handles) <= 1:
                print("[!] 상세조회 새 창이 열려있지 않습니다. 먼저 2번을 실행하세요.")
            else:
                process_current_detail_window()
        elif choice == "4":
            if driver is None:
                start_login()
                time.sleep(2)
            process_purchase_requests(single_only=True)
        elif choice == "" or choice == "5":
            if driver is None:
                start_login()
                time.sleep(2)
            process_purchase_requests(single_only=False)
        elif choice == "6":
            change_target_directory()
        elif choice == "0":
            if driver:
                try: driver.quit()
                except: pass
            print("[*] 스크래퍼를 종료합니다.")
            break
        else:
            print("[!] 잘못된 입력입니다. 1~6 또는 0을 입력하세요.")

# =================================================================
# 🖥️ 윈도우 전용 그래픽 화면 (Tkinter GUI 대시보드)
# =================================================================
def launch_scraper_gui():
    global RERP_ID, RERP_PW, GEMINI_KEY, BASE_TARGET_DIR, TEMP_DL_DIR, ai_client, gui_status_cb, stop_requested
    
    root = tk.Tk()
    root.title("산학협력단 rERP 구매요청 자동 스크래퍼 + AI")
    root.geometry("780x740")
    root.minsize(700, 600)
    root.configure(bg="#F4F6F9")

    # Queue stdout redirection
    log_queue = queue.Queue()
    orig_stdout = sys.stdout
    orig_stderr = sys.stderr
    sys.stdout = QueueStdout(log_queue, orig_stdout)
    sys.stderr = QueueStdout(log_queue, orig_stderr)

    style = ttk.Style()
    style.theme_use("clam")

    # 1. 상단 헤더
    header = tk.Frame(root, bg="#004797", padx=16, pady=12)
    header.pack(fill=tk.X)

    tk.Label(
        header,
        text="🚀 구매요청 자동 스크래퍼",
        font=("Malgun Gothic", 13, "bold"),
        fg="white",
        bg="#004797"
    ).pack(anchor="w")

    # 2. 계정 및 저장 설정 카드
    cfg_frame = tk.LabelFrame(
        root,
        text=" 👤 계정 및 저장 설정 ",
        font=("Malgun Gothic", 10, "bold"),
        bg="#F4F6F9",
        fg="#222",
        padx=14,
        pady=10
    )
    cfg_frame.pack(fill=tk.X, padx=15, pady=(10, 6))

    # Row 0: ID and PW
    row0 = tk.Frame(cfg_frame, bg="#F4F6F9")
    row0.pack(fill=tk.X, pady=3)

    tk.Label(row0, text="rERP 아이디:", font=("Malgun Gothic", 9, "bold"), bg="#F4F6F9").pack(side=tk.LEFT, padx=(0, 5))
    entry_id = ttk.Entry(row0, width=18, font=("Malgun Gothic", 9))
    entry_id.pack(side=tk.LEFT, padx=(0, 20))
    entry_id.insert(0, RERP_ID if RERP_ID != "본인_아이디_입력" else "")

    tk.Label(row0, text="비밀번호:", font=("Malgun Gothic", 9, "bold"), bg="#F4F6F9").pack(side=tk.LEFT, padx=(0, 5))
    entry_pw = ttk.Entry(row0, width=18, font=("Malgun Gothic", 9), show="*")
    entry_pw.pack(side=tk.LEFT, padx=(0, 6))
    entry_pw.insert(0, RERP_PW if RERP_PW else "")

    def toggle_pw_view():
        if entry_pw.cget("show") == "*":
            entry_pw.configure(show="")
            btn_eye.configure(text="🙈")
        else:
            entry_pw.configure(show="*")
            btn_eye.configure(text="👁️")

    btn_eye = tk.Button(row0, text="👁️", font=("Segoe UI Emoji", 8), relief=tk.FLAT, bg="#E9ECEF", cursor="hand2", command=toggle_pw_view)
    btn_eye.pack(side=tk.LEFT, padx=(0, 15))

    def on_save_settings():
        global RERP_ID, RERP_PW, GEMINI_KEY, BASE_TARGET_DIR, TEMP_DL_DIR, ai_client
        RERP_ID = entry_id.get().strip()
        RERP_PW = entry_pw.get().strip()
        GEMINI_KEY = entry_gemini.get().strip()
        BASE_TARGET_DIR = entry_dir.get().strip()
        TEMP_DL_DIR = os.path.join(BASE_TARGET_DIR, "_temp_dl")
        os.makedirs(TEMP_DL_DIR, exist_ok=True)
        if GEMINI_KEY:
            try:
                ai_client = genai.Client(api_key=GEMINI_KEY)
            except Exception:
                pass

        try:
            env_map = {}
            if os.path.exists(".env"):
                with open(".env", "r", encoding="utf-8") as ef:
                    for line in ef:
                        if "=" in line:
                            k, v = line.strip().split("=", 1)
                            env_map[k] = v
            env_map["RERP_ID"] = RERP_ID
            env_map["RERP_PW"] = RERP_PW
            if GEMINI_KEY:
                env_map["GEMINI_API_KEY"] = GEMINI_KEY
            if BASE_TARGET_DIR:
                env_map["TARGET_DIR"] = BASE_TARGET_DIR

            with open(".env", "w", encoding="utf-8") as ef:
                for k, v in env_map.items():
                    ef.write(f"{k}={v}\n")
            messagebox.showinfo("저장 완료", "계정 및 환경설정이 성공적으로 저장되었습니다!")
            print("[+] 계정 및 환경설정 저장 완료.")
        except Exception as e:
            messagebox.showerror("오류", f"설정 저장 실패: {e}")

    btn_save_cfg = tk.Button(row0, text="💾 설정 저장", font=("Malgun Gothic", 9, "bold"), bg="#28A745", fg="white", activebackground="#218838", activeforeground="white", relief=tk.FLAT, padx=8, pady=2, cursor="hand2", command=on_save_settings)
    btn_save_cfg.pack(side=tk.RIGHT)

    # Row 1: Gemini Key
    row1 = tk.Frame(cfg_frame, bg="#F4F6F9")
    row1.pack(fill=tk.X, pady=3)

    tk.Label(row1, text="Gemini API 키:", font=("Malgun Gothic", 9), bg="#F4F6F9").pack(side=tk.LEFT, padx=(0, 5))
    entry_gemini = ttk.Entry(row1, font=("Consolas", 8))
    entry_gemini.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
    entry_gemini.insert(0, GEMINI_KEY if GEMINI_KEY else DEFAULT_API_KEY)

    # Row 2: Target Directory
    row2 = tk.Frame(cfg_frame, bg="#F4F6F9")
    row2.pack(fill=tk.X, pady=3)

    tk.Label(row2, text="저장 폴더 경로:", font=("Malgun Gothic", 9), bg="#F4F6F9").pack(side=tk.LEFT, padx=(0, 5))
    entry_dir = ttk.Entry(row2, font=("Malgun Gothic", 9))
    entry_dir.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6))
    entry_dir.insert(0, BASE_TARGET_DIR)

    def on_browse_dir():
        sel = filedialog.askdirectory(title="문서 저장 폴더 선택", initialdir=entry_dir.get())
        if sel:
            entry_dir.delete(0, tk.END)
            entry_dir.insert(0, sel)

    btn_browse = tk.Button(row2, text="📁 찾아보기", font=("Malgun Gothic", 8), bg="#E9ECEF", relief=tk.FLAT, padx=6, pady=2, cursor="hand2", command=on_browse_dir)
    btn_browse.pack(side=tk.RIGHT)

    # 3. 스크래퍼 실행 제어 카드
    act_frame = tk.LabelFrame(
        root,
        text=" 🚀 스크래퍼 실행 제어 ",
        font=("Malgun Gothic", 10, "bold"),
        bg="#F4F6F9",
        fg="#222",
        padx=14,
        pady=10
    )
    act_frame.pack(fill=tk.X, padx=15, pady=(0, 6))

    def sync_credentials_from_ui():
        global RERP_ID, RERP_PW, GEMINI_KEY, BASE_TARGET_DIR, TEMP_DL_DIR, ai_client
        RERP_ID = entry_id.get().strip()
        RERP_PW = entry_pw.get().strip()
        GEMINI_KEY = entry_gemini.get().strip() or DEFAULT_API_KEY
        BASE_TARGET_DIR = entry_dir.get().strip()
        TEMP_DL_DIR = os.path.join(BASE_TARGET_DIR, "_temp_dl")
        os.makedirs(TEMP_DL_DIR, exist_ok=True)
        if GEMINI_KEY:
            try:
                ai_client = genai.Client(api_key=GEMINI_KEY)
            except Exception:
                pass

        if not RERP_ID or not RERP_PW:
            messagebox.showwarning("입력 필요", "rERP 아이디와 비밀번호를 입력해주세요!")
            return False
        return True

    def do_full_auto():
        if not sync_credentials_from_ui(): return
        run_in_thread(_task_full_auto)

    def _task_full_auto():
        global driver
        if driver is None:
            start_login(interactive=False)
            time.sleep(2)
        process_purchase_requests(single_only=False)

    btn_full_auto = tk.Button(
        act_frame,
        text="🚀 [원클릭] 결재대기함 전체 일괄 자동 처리 (다운로드 + AI분석 + 폴더링)",
        font=("Malgun Gothic", 11, "bold"),
        bg="#004797",
        fg="white",
        activebackground="#003366",
        activeforeground="white",
        relief=tk.FLAT,
        pady=10,
        cursor="hand2",
        command=do_full_auto
    )
    btn_full_auto.pack(fill=tk.X, pady=(0, 8))

    # Sub actions row 1
    sub_row1 = tk.Frame(act_frame, bg="#F4F6F9")
    sub_row1.pack(fill=tk.X, pady=(0, 4))

    def do_single_auto():
        if not sync_credentials_from_ui(): return
        run_in_thread(_task_single_auto)

    def _task_single_auto():
        global driver
        if driver is None:
            start_login(interactive=False)
            time.sleep(2)
        process_purchase_requests(single_only=True)

    def do_login_only():
        if not sync_credentials_from_ui(): return
        run_in_thread(lambda: start_login(interactive=False))

    def do_open_first():
        run_in_thread(open_first_request_only)

    def do_process_current():
        run_in_thread(process_current_detail_window)

    btn_single = tk.Button(sub_row1, text="🎯 1건만 테스트", font=("Malgun Gothic", 9, "bold"), bg="#17A2B8", fg="white", activebackground="#138496", activeforeground="white", relief=tk.FLAT, pady=6, cursor="hand2", command=do_single_auto)
    btn_single.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 3))

    btn_login = tk.Button(sub_row1, text="🔑 1단계: 로그인만", font=("Malgun Gothic", 9), bg="#4A6984", fg="white", activebackground="#3B546A", activeforeground="white", relief=tk.FLAT, pady=6, cursor="hand2", command=do_login_only)
    btn_login.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(3, 3))

    btn_open1 = tk.Button(sub_row1, text="📑 2단계: 첫문서 열기", font=("Malgun Gothic", 9), bg="#4A6984", fg="white", activebackground="#3B546A", activeforeground="white", relief=tk.FLAT, pady=6, cursor="hand2", command=do_open_first)
    btn_open1.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(3, 3))

    btn_proc3 = tk.Button(sub_row1, text="🔍 3단계: 다운+AI분석", font=("Malgun Gothic", 9), bg="#4A6984", fg="white", activebackground="#3B546A", activeforeground="white", relief=tk.FLAT, pady=6, cursor="hand2", command=do_process_current)
    btn_proc3.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(3, 0))

    # Sub actions row 2 (Utility)
    sub_row2 = tk.Frame(act_frame, bg="#F4F6F9")
    sub_row2.pack(fill=tk.X, pady=(4, 0))

    def on_open_folder():
        target = entry_dir.get().strip() or BASE_TARGET_DIR
        if os.path.exists(target):
            os.startfile(target)
        else:
            messagebox.showwarning("확인", f"폴더가 아직 존재하지 않습니다:\n{target}")

    btn_open_target = tk.Button(sub_row2, text="📂 저장 폴더 열기", font=("Malgun Gothic", 9), bg="#6C757D", fg="white", activebackground="#5A6268", activeforeground="white", relief=tk.FLAT, pady=5, cursor="hand2", command=on_open_folder)
    btn_open_target.pack(side=tk.LEFT, padx=(0, 6))

    def on_clear_log():
        txt_log.configure(state=tk.NORMAL)
        txt_log.delete("1.0", tk.END)
        txt_log.configure(state=tk.DISABLED)

    btn_clear_log = tk.Button(sub_row2, text="🧹 로그 비우기", font=("Malgun Gothic", 9), bg="#E2E6EA", fg="#333", relief=tk.FLAT, pady=5, cursor="hand2", command=on_clear_log)
    btn_clear_log.pack(side=tk.LEFT, padx=(0, 6))

    def on_stop_all():
        global stop_requested, driver
        stop_requested = True
        print("\n[🛑] 중단 요청됨: 안전하게 작업을 중단하고 브라우저를 닫습니다...")
        status_var.set("상태: 작업 중단 요청됨")
        if driver:
            try:
                driver.quit()
                driver = None
                print("[+] 브라우저가 종료되었습니다.")
            except: pass

    btn_stop = tk.Button(sub_row2, text="🛑 작업 중단 / 브라우저 닫기", font=("Malgun Gothic", 9, "bold"), bg="#DC3545", fg="white", activebackground="#C82333", activeforeground="white", relief=tk.FLAT, pady=5, cursor="hand2", command=on_stop_all)
    btn_stop.pack(side=tk.RIGHT)

    # 4. 상태 및 프로그레스바
    status_frame = tk.Frame(root, bg="#F4F6F9", padx=15)
    status_frame.pack(fill=tk.X, pady=(2, 2))

    status_var = tk.StringVar(value="상태: 대기 중")
    lbl_status = tk.Label(status_frame, textvariable=status_var, font=("Malgun Gothic", 9, "bold"), fg="#004797", bg="#F4F6F9")
    lbl_status.pack(side=tk.LEFT)

    progress = ttk.Progressbar(status_frame, mode="indeterminate", length=200)
    progress.pack(side=tk.RIGHT)

    def update_gui_status(text):
        status_var.set(f"상태: {text}")

    gui_status_cb = update_gui_status

    # Thread runner & Busy management
    action_buttons = [btn_full_auto, btn_single, btn_login, btn_open1, btn_proc3]

    def set_ui_busy(busy: bool):
        if busy:
            for b in action_buttons:
                b.configure(state=tk.DISABLED)
            progress.start(10)
        else:
            for b in action_buttons:
                b.configure(state=tk.NORMAL)
            progress.stop()
            status_var.set("상태: 대기 중")

    def run_in_thread(fn):
        def worker():
            try:
                set_ui_busy(True)
                status_var.set("상태: 작업 진행 중...")
                fn()
            except Exception as e:
                print(f"\n[!] 실행 중 오류: {e}\n{traceback.format_exc()}")
            finally:
                set_ui_busy(False)
        threading.Thread(target=worker, daemon=True).start()

    # 5. 실시간 진행 로그 영역
    log_card = tk.LabelFrame(
        root,
        text=" 📋 실시간 진행 상황 및 로그 ",
        font=("Malgun Gothic", 9, "bold"),
        bg="#F4F6F9",
        fg="#222",
        padx=10,
        pady=8
    )
    log_card.pack(fill=tk.BOTH, expand=True, padx=15, pady=(2, 12))

    txt_log = tk.Text(
        log_card,
        font=("Consolas", 9),
        bg="#FFFFFF",
        fg="#212529",
        wrap=tk.WORD,
        relief=tk.SOLID,
        borderwidth=1
    )
    sb = tk.Scrollbar(log_card, command=txt_log.yview)
    txt_log.configure(yscrollcommand=sb.set)
    sb.pack(side=tk.RIGHT, fill=tk.Y)
    txt_log.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    txt_log.configure(state=tk.DISABLED)

    def poll_logs():
        while not log_queue.empty():
            try:
                msg = log_queue.get_nowait()
                txt_log.configure(state=tk.NORMAL)
                txt_log.insert(tk.END, msg)
                txt_log.see(tk.END)
                txt_log.configure(state=tk.DISABLED)
            except queue.Empty:
                break
        root.after(80, poll_logs)

    root.after(80, poll_logs)

    def on_window_close():
        global driver
        sys.stdout = orig_stdout
        sys.stderr = orig_stderr
        if driver:
            try: driver.quit()
            except: pass
        root.destroy()

    root.protocol("WM_DELETE_WINDOW", on_window_close)

    print("==================================================")
    print("  산학협력단 rERP 구매요청 자동 스크래퍼 v2.0")
    print(f"  [현재 저장 경로] {BASE_TARGET_DIR}")
    print("  원하시는 작업을 위 버튼으로 선택하세요.")
    print("==================================================")

    root.mainloop()

if __name__ == "__main__":
    if "--cli" in sys.argv:
        main_menu()
    else:
        launch_scraper_gui()
