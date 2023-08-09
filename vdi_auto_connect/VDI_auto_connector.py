# Library
import os
import sys
import time

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
# headless 옵션 지정시, .slv 파일 다운로드 불가
options = webdriver.ChromeOptions()
options.add_argument("window-size=1920x1080")
options.add_argument("disable-gpu")

# Path Set
P_Downloads = os.path.join(os.path.expanduser("~"), "Downloads")
P_Documents = os.path.join(os.path.expanduser("~"), "Documents")
P_ChromeDriver = os.path.join(os.getcwd(), "chromedriver.exe") # 115.0.5790.171, https://googlechromelabs.github.io/chrome-for-testing/
VDI_Launcher = os.path.join(P_Downloads, "launch.slv")
#
max_delay_time = 60
SKMR_Portal_URL = "https://vdi.sk-materials.com/intweb/Login/Login?uid=6975151"

# Selenium 기본 동작 응용
class Clickable:
    @staticmethod
    def element(wait, ptype, path):
        if ptype=="xpath": return wait.until(EC.element_to_be_clickable((By.XPATH, path)))
        else: return wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, path)))
    @staticmethod
    def click(wait, ptype, path, f_delay=0, r_delay=0):
        time.sleep(f_delay)
        Clickable.element(wait, ptype, path).click()
        time.sleep(r_delay)

if __name__=="__main__":
    vdi_id, vdi_pw = sys.argv[1], sys.argv[2]
    # SK포털 로그인
    browser = webdriver.Chrome(P_ChromeDriver, options=options)
    browser.get(SKMR_Portal_URL)
    browser.maximize_window()
    browser.minimize_window()
    wait = WebDriverWait(browser, max_delay_time)
    # VDI ID
    Clickable.element(wait, "css", "#wrap > div.login-area > div.login-box > ul.inpbox > li:nth-child(1) > input").clear()
    Clickable.element(wait, "css", "#wrap > div.login-area > div.login-box > ul.inpbox > li:nth-child(1) > input").send_keys(vdi_id)
    # VDI PW
    Clickable.element(wait, "css", "#wrap > div.login-area > div.login-box > ul.inpbox > li:nth-child(2) > input").clear()
    Clickable.element(wait, "css", "#wrap > div.login-area > div.login-box > ul.inpbox > li:nth-child(2) > input").send_keys(vdi_pw)
    # Login
    Clickable.click(wait, "xpath", '//*[@id="wrap"]/div[2]/div[1]/a[1]', r_delay=10)
    # Start
    Clickable.click(wait, "css", "#visual_wrap > div > div > div > div > button", r_delay=10)
    browser.quit()
    # VDI 실행
    os.system(VDI_Launcher)
    try: os.remove(VDI_Launcher)
    except: pass
    
    
