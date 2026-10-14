# %%
import os
import traceback
from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

# 1. 환경변수(.env) 세팅
load_dotenv()
RERP_ID = os.getenv("RERP_ID", "본인_아이디_입력")
RERP_PW = os.getenv("RERP_PW", "본인_비밀번호_입력")

# =================================================================
# 💡 [핵심] 브라우저 창을 모든 블록이 공유하기 위한 공용 변수
# =================================================================
driver = None 

# =================================================================
# 🔘 [버튼 1 역할] 로그인 및 결재대기함 열기 (순철님 코드 원본 유지)
# =================================================================
def start_login():
    global driver # 밖에 있는 공용 driver를 쓰겠다고 선언
    try:
        options = webdriver.ChromeOptions()
        options.add_experimental_option("detach", True)
        options.add_experimental_option("excludeSwitches", ["enable-automation"])
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")

        prefs = {
            "credentials_enable_service": False,
            "profile.password_manager_enabled": False
        }
        options.add_experimental_option("prefs", prefs)

        print("🤖 기본 구글 크롬 스크래퍼 가동 중...")
        driver = webdriver.Chrome(options=options)

        driver.set_window_position(2000, 0)
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
        print("🚀 로그인 버튼 클릭 완료!")

        # 팝업 방어 1: 새 창 팝업 닫기
        try:
            WebDriverWait(driver, 5).until(EC.number_of_windows_to_be(2))
            handles = driver.window_handles
            driver.switch_to.window(handles[1]) 
            driver.close()                      
            driver.switch_to.window(handles[0]) 
            print("✅ [성공] 새 창 팝업을 닫았습니다.")
        except TimeoutException:
            print("▶ 새 창 팝업 없음")

        # 팝업 방어 2: 아이프레임(방) 전수조사 로직
        driver.switch_to.default_content() 
        
        try:
            close_layer_btn = WebDriverWait(driver, 3).until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="unified_popup_layer"]/div/div/div[1]/button'))
            )
            close_layer_btn.click()
            print("✅ [성공] 거실에서 회색 레이어 팝업을 닫았습니다.")
        except TimeoutException:
            iframes = driver.find_elements(By.TAG_NAME, "iframe")
            closed = False
            
            for index, iframe in enumerate(iframes):
                try:
                    driver.switch_to.default_content()
                    driver.switch_to.frame(iframe) 
                    
                    close_layer_btn = WebDriverWait(driver, 1.5).until(
                        EC.element_to_be_clickable((By.XPATH, '//*[@id="unified_popup_layer"]/div/div/div[1]/button'))
                    )
                    close_layer_btn.click()
                    print(f"✅ [성공] {index}번 방(iframe) 안에서 회색 레이어 팝업을 닫았습니다.")
                    closed = True
                    break 
                except Exception:
                    pass
            
            driver.switch_to.default_content() 
            if not closed:
                print("▶ 회색 레이어 팝업 없음 또는 이미 닫힘")

        # 1단계: 아이프레임(방)을 뒤져서 결재대기함 버튼 클릭
        driver.switch_to.default_content() 
        iframes = driver.find_elements(By.TAG_NAME, "iframe")
        clicked = False
        
        for index, iframe in enumerate(iframes):
            try:
                driver.switch_to.default_content()
                driver.switch_to.frame(iframe) 
                
                wait_box = WebDriverWait(driver, 1.5).until(
                    EC.element_to_be_clickable((By.XPATH, '//*[@id="content"]/div/div[2]/div[2]/div[2]/div[1]/table/tbody/tr/td[1]'))
                )
                wait_box.click()
                print(f"✅ {index}번 방(iframe) 안에서 '결재대기함' 버튼 클릭 완료!")
                clicked = True
                break 
            except Exception:
                pass 
        
        driver.switch_to.default_content() 
        
        if clicked:
            driver.implicitly_wait(2) 
        else:
            print("🚨 모든 방을 뒤졌지만 결재대기함 버튼을 찾지 못했습니다.")

    except Exception as e:
        error_msg = traceback.format_exc()
        print(f"🚨 크롬 실행 중 오류가 발생했습니다.\n{e}\n\n상세 내용:\n{error_msg}")

# =================================================================
# 🔘 [버튼 2 역할] 구매요청 문서 솎아내기 
# =================================================================
def process_purchase_request():
    global driver
    if driver is None:
        print("🚨 브라우저가 꺼져 있습니다. 1번(로그인) 블록을 먼저 실행하세요.")
        return

    try:
        print("🔍 [버튼 2] '구매요청' 문서 스캔 시작...")
        
        # iframe 안에 갇혀있을 경우 뚫고 들어가는 로직 (결재대기함 화면용)
        driver.switch_to.default_content()
        iframes = driver.find_elements(By.TAG_NAME, "iframe")
        target_found = False

        for iframe in iframes:
            try:
                driver.switch_to.default_content()
                driver.switch_to.frame(iframe)
                
                # 표가 있는지 2초간 확인
                WebDriverWait(driver, 2).until(
                    EC.presence_of_element_located((By.XPATH, "//td[contains(@class, 'HideCol0PROC_TYP_NM')]"))
                )
                
                proc_types = driver.find_elements(By.XPATH, "//td[contains(@class, 'HideCol0PROC_TYP_NM')]")
                
                for index, proc_td in enumerate(proc_types):
                    if "구매요청" in proc_td.text.strip():
                        print(f"✅ 위에서 {index + 1}번째 줄에서 '구매요청' 발견!")
                        title_td = proc_td.find_element(By.XPATH, "./..").find_element(By.XPATH, ".//td[contains(@class, 'HideCol0PRJ_NM')]")
                        print(f"📄 문서 제목: {title_td.text}")
                        title_td.click()
                        target_found = True
                        break
                
                if target_found: break # 찾았으면 프레임 탐색 종료
            except:
                pass
                
        if not target_found:
            print("▶ 현재 목록에 '구매요청' 대기 문서가 없습니다.")
            
    except Exception as e:
        print(f"🚨 2번 버튼(구매요청) 오류: {e}")

# =================================================================
# 🚀 파이썬 쾌속 부분 테스트용 실행 영역 (Jupyter Cell 방식)
# =================================================================

# %% 
# 1번 버튼: 로그인 및 대기함 열기 (Run Cell 클릭)
start_login()

# %% 
# 2번 버튼: 구매요청 문서 솎아내기 (Run Cell 클릭)
process_purchase_request()
# %%
