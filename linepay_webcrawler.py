# -*- mode: python ; coding: utf-8 -*-
"""
This program is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with this program. If not, see <https://www.gnu.org/licenses/>.
© 2025, waaaaaaz (fisherw0311@gmail.com)

"""
import os
import time
import re
import json
import logging
from logging.handlers import RotatingFileHandler
import datetime
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import winreg

# log 與 RotatingFileHandler

current_dir = os.path.dirname(os.path.abspath(__file__))
log_dir = os.path.join(current_dir, "log")
if not os.path.exists(log_dir):
    os.makedirs(log_dir)
    logger_temp = logging.getLogger("download_logger_temp")
    logger_temp.debug("建立 log 資料夾：%s", log_dir)

timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
log_file = os.path.join(log_dir, f"download_log_{timestamp}.txt")
logger = logging.getLogger("download_logger")
logger.setLevel(logging.DEBUG)
handler = RotatingFileHandler(log_file, maxBytes=5*1024*1024, backupCount=5, encoding="utf-8")
formatter = logging.Formatter("%(asctime)s - %(levelname)s - %(message)s")
handler.setFormatter(formatter)
logger.addHandler(handler)
logger.propagate = False  

logger.info("程式啟動。")
logger.debug("目前工作目錄: %s", current_dir)

# credentials 

credentials_path = os.path.join(current_dir, 'credentials.json')
if not os.path.exists(credentials_path):
    logger.error("找不到 credentials 檔案: %s", credentials_path)
    raise FileNotFoundError(f"找不到 credentials 檔案: {credentials_path}")

with open(credentials_path, 'r', encoding='utf-8') as f:
    credentials = json.load(f)
logger.debug("credentials 檔案內容讀取完成。")

username_credential = credentials.get("username")
password_credential = credentials.get("password")

if not username_credential or not password_credential:
    logger.error("credentials 檔案中缺少 username 或 password")
    raise ValueError("credentials 檔案中缺少 username 或 password")

logger.debug("credentials 載入成功。")


# 載目錄

download_path = os.path.abspath("downloads")
if not os.path.exists(download_path):
    os.makedirs(download_path)
    logger.debug("建立下載資料夾：%s", download_path)
else:
    logger.debug("下載資料夾已存在：%s", download_path)


chrome_options = Options()
chrome_options.add_experimental_option("prefs", {
    "download.default_directory": download_path,  
    "download.prompt_for_download": False,         
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True
})
logger.debug("設定 Chrome 首選項。")
chrome_options.add_argument("--headless")  # 無頭
logger.debug("啟用 headless 模式。")

# chrome
def chrome_installed():
    reg_path = r"SOFTWARE\WOW6432Node\Google"
    
    try:
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, reg_path):
            return True  
    except FileNotFoundError:
        return False  

if chrome_installed():
    logger.debug("Google Chrome 已安裝.")
    print("Google Chrome 已安裝")
else:
    logger.error("Google Chrome 未找到或未下載.")
    raise Exception("Google Chrome 未找到或未下載.")

#  WebDriver
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
logger.info("Chrome WebDriver 初始化完成。")
logger.debug("WebDriver 版本與相關資訊: %s", driver.capabilities)

try:
    #################################
    # S1
    print("導向登入頁面。")
    logger.info("導向登入頁面。")
    login_url = "https://pay.line.me/portal/tw/auth/login"
    driver.get(login_url)
    logger.debug("已導向 URL: %s", login_url)
    
    username_field = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "id"))
    )
    logger.debug("找到 username 輸入框，元素內容: %s", username_field)
    username_field.send_keys(username_credential)
    print("輸入 username。")
    logger.debug("輸入 username 完成。")
    
    password_field = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "password"))
    )
    logger.debug("找到 password 輸入框，元素內容: %s", password_field)
    password_field.send_keys(password_credential)
    print("輸入 password。")
    logger.debug("輸入 password 完成。")
    
    login_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, "loginActionButton"))
    )
    logger.debug("找到登入按鈕，元素內容: %s", login_button)
    login_button.click()
    print("點擊登入按鈕。")
    logger.info("點擊登入按鈕。")
    
    time.sleep(5)
    logger.debug("等待 5 秒以確保登入流程穩定，當前 URL: %s", driver.current_url)
    print("✅ 登入成功！")
    logger.info("登入成功！")
    
    #################################
    # S2 (CSV)
    existing_csv = set(f for f in os.listdir(download_path) if f.lower().endswith('.csv'))
    print("導向 CSV 下載頁面。")
    logger.info("導向 CSV 下載頁面。")
    csv_download_url = "https://pay.line.me/tw/center/download/dealDownloadView?locale=zh-tw"
    driver.get(csv_download_url)
    logger.debug("已導向 CSV 下載頁面 URL: %s", csv_download_url)
    
    initial_button = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located(
            (By.XPATH, "//button[contains(@class, 'download_btn') and contains(text(), 'Download')]")
        )
    )
    logger.debug("找到初始 CSV 下載按鈕: %s", initial_button)
    initial_onclick = initial_button.get_attribute("onclick")
    logger.debug("初始 CSV 下載按鈕 onclick: %s", initial_onclick)
    match = re.search(r"downloadFile\('(\d+)'\);", initial_onclick)
    if match:
        initial_number = int(match.group(1))
        print("初始下載按鈕編號 (CSV):", initial_number)
        logger.info("初始 CSV 下載按鈕編號: %d", initial_number)
    else:
        raise Exception("無法從初始 CSV 下載按鈕中擷取編號")
    
    csv_span = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//span[text()='CSV']"))
    )
    logger.debug("找到 CSV 選項按鈕: %s", csv_span)
    csv_span.click()
    print("✅ CSV 選項已點選")
    logger.info("已選取 CSV 選項。")
    
    #################################
    # S3 (CSV)
    time.sleep(2)
    def new_download_button_available(driver):
        try:
            reset_span = driver.find_element(By.XPATH, "//span[text()='重設此頁']")
            reset_span.click()
            print("✅ 已點選【重設此頁】(CSV)")
            logger.info("點選【重設此頁】(CSV)。")
        except Exception as e:
            print("重設按鈕錯誤:", e)
            logger.error("點選重設按鈕錯誤: %s", e)
        time.sleep(1)
        buttons = driver.find_elements(By.XPATH, "//button[contains(@class, 'download_btn') and contains(text(), 'Download')]")
        logger.debug("搜尋到 %d 個 CSV Download 按鈕", len(buttons))
        for btn in buttons:
            try:
                onclick_attr = btn.get_attribute("onclick")
                logger.debug("檢查按鈕，onclick 屬性: %s", onclick_attr)
                if onclick_attr:
                    m = re.search(r"downloadFile\('(\d+)'\);", onclick_attr)
                    if m:
                        new_number = int(m.group(1))
                        logger.debug("按鈕下載編號: %d (初始編號: %d)", new_number, initial_number)
                        if new_number > initial_number:
                            logger.debug("找到更新後的 CSV 按鈕，編號: %d", new_number)
                            return btn
            except Exception as e:
                print("處理按鈕錯誤:", e)
                logger.error("處理按鈕錯誤: %s", e)
        print("ℹ️ 新下載按鈕號碼尚未更新，繼續等待...")
        logger.debug("新 CSV 下載按鈕號碼尚未更新，繼續等待...")
        return False

    new_download_button = WebDriverWait(driver, 300, poll_frequency=5).until(new_download_button_available)
    new_number_found = re.search(r"downloadFile\('(\d+)'\);", new_download_button.get_attribute("onclick")).group(1)
    print("✅ 找到新的 CSV 下載按鈕，編號:", new_number_found)
    logger.info("找到新的 CSV 下載按鈕，編號: %s", new_number_found)
    
    new_download_button.click()
    print("✅ CSV Download 按鈕已點擊，開始下載")
    logger.info("點擊 CSV Download 按鈕，開始下載。")
    
    #################################
    # S4 等載
    timeout = 20
    downloaded_file_csv = None
    while timeout > 0:
        try:
            current_csv = set(f for f in os.listdir(download_path) if f.lower().endswith('.csv'))
            new_csv_set = current_csv - existing_csv  # 只看剛新增的檔案
            logger.debug("CSV 下載監控：原有=%d，目前=%d，新增=%d",
                        len(existing_csv), len(current_csv), len(new_csv_set))

            if new_csv_set:
                # 多個新檔時，選修改時間最新的
                newest = max(
                    new_csv_set,
                    key=lambda fn: os.path.getmtime(os.path.join(download_path, fn))
                )
                downloaded_file_csv = newest
                print("✅ CSV 下載成功:", downloaded_file_csv)
                logger.info("CSV 下載成功: %s", downloaded_file_csv)
                break
        except Exception as e:
            print("檢查 CSV 檔案錯誤:", e)
            logger.error("檢查 CSV 檔案錯誤: %s", e)

        time.sleep(1)
        timeout -= 1
        logger.debug("CSV 下載等待中，剩餘 %d 秒", timeout)

    if not downloaded_file_csv:
        print("❌ CSV 下載檔案失敗或逾時")
        logger.error("CSV 下載檔案失敗或逾時")
    
    #################################
    # S5 (EXCEL)
    existing_excel = set(
    f for f in os.listdir(download_path)
    if f.lower().endswith(('.xls', '.xlsx'))
    )

    print("===============\n開始下載 EXCEL 檔案")
    logger.info("=============== 開始下載 EXCEL 檔案")
    driver.get("https://pay.line.me/tw/center/download/dealDownloadView?locale=zh-tw")
    logger.debug("重新導向至 EXCEL 下載頁面。")
    
    initial_button_excel = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located(
            (By.XPATH, "//button[contains(@class, 'download_btn') and contains(text(), 'Download')]")
        )
    )
    logger.debug("找到初始 EXCEL 下載按鈕: %s", initial_button_excel)
    initial_onclick_excel = initial_button_excel.get_attribute("onclick")
    logger.debug("初始 EXCEL 下載按鈕 onclick: %s", initial_onclick_excel)
    match_excel = re.search(r"downloadFile\('(\d+)'\);", initial_onclick_excel)
    if match_excel:
        initial_number_excel = int(match_excel.group(1))
        print("初始下載按鈕編號 (EXCEL):", initial_number_excel)
        logger.info("初始 EXCEL 下載按鈕編號: %d", initial_number_excel)
    else:
        raise Exception("無法從初始 EXCEL 下載按鈕中擷取編號")
    
    excel_span = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//span[text()='EXCEL']"))
    )
    logger.debug("找到 EXCEL 選項按鈕: %s", excel_span)
    excel_span.click()
    print("✅ EXCEL 選項已點選")
    logger.info("已選取 EXCEL 選項。")
    
    time.sleep(2)
    def new_download_button_available_excel(driver):
        try:
            reset_span = driver.find_element(By.XPATH, "//span[text()='重設此頁']")
            reset_span.click()
            print("✅ 已點選【重設此頁】(EXCEL)")
            logger.info("點選【重設此頁】(EXCEL)。")
        except Exception as e:
            print("EXCEL 重設按鈕錯誤:", e)
            logger.error("EXCEL 重設按鈕錯誤: %s", e)
        time.sleep(1)
        buttons = driver.find_elements(By.XPATH, "//button[contains(@class, 'download_btn') and contains(text(), 'Download')]")
        logger.debug("搜尋到 %d 個 EXCEL Download 按鈕", len(buttons))
        for btn in buttons:
            try:
                onclick_attr = btn.get_attribute("onclick")
                logger.debug("檢查 EXCEL 按鈕，onclick 屬性: %s", onclick_attr)
                if onclick_attr:
                    m = re.search(r"downloadFile\('(\d+)'\);", onclick_attr)
                    if m:
                        new_number = int(m.group(1))
                        logger.debug("EXCEL 按鈕下載編號: %d (初始編號: %d)", new_number, initial_number_excel)
                        if new_number > initial_number_excel:
                            logger.debug("找到更新後的 EXCEL 按鈕，編號: %d", new_number)
                            return btn
            except Exception as e:
                print("處理 EXCEL 按鈕錯誤:", e)
                logger.error("處理 EXCEL 按鈕錯誤: %s", e)
        print("ℹ️ 新 EXCEL 下載按鈕號碼尚未更新，繼續等待...")
        logger.debug("新 EXCEL 下載按鈕號碼尚未更新，繼續等待...")
        return False

    new_download_button_excel = WebDriverWait(driver, 300, poll_frequency=5).until(new_download_button_available_excel)
    new_number_found_excel = re.search(r"downloadFile\('(\d+)'\);", new_download_button_excel.get_attribute("onclick")).group(1)
    print("✅ 找到新的 EXCEL 下載按鈕，編號:", new_number_found_excel)
    logger.info("找到新的 EXCEL 下載按鈕，編號: %s", new_number_found_excel)
    
    new_download_button_excel.click()
    print("✅ EXCEL Download 按鈕已點擊，開始下載")
    logger.info("點擊 EXCEL Download 按鈕，開始下載。")
    
    # 等載 (.xls 或 .xlsx)
    timeout = 20
    downloaded_file_excel = None
    while timeout > 0:
        try:
            current_excel = set(
                f for f in os.listdir(download_path)
                if f.lower().endswith(('.xls', '.xlsx'))
            )
            new_excel_set = current_excel - existing_excel
            logger.debug("EXCEL 下載監控：原有=%d，目前=%d，新增=%d",
                        len(existing_excel), len(current_excel), len(new_excel_set))

            if new_excel_set:                
                newest = max(
                    new_excel_set,
                    key=lambda fn: os.path.getmtime(os.path.join(download_path, fn))
                )
                downloaded_file_excel = newest
                print("✅ EXCEL 下載成功:", downloaded_file_excel)
                logger.info("EXCEL 下載成功: %s", downloaded_file_excel)
                break
        except Exception as e:
            print("檢查 EXCEL 檔案錯誤:", e)
            logger.error("檢查 EXCEL 檔案錯誤: %s", e)

        time.sleep(1)
        timeout -= 1
        logger.debug("EXCEL 下載等待中，剩餘 %d 秒", timeout)

    if not downloaded_file_excel:
        print("❌ EXCEL 下載檔案失敗或逾時")
        logger.error("EXCEL 下載檔案失敗或逾時")

except Exception as e:
    print("發生錯誤:", e)
    logger.exception("程式執行過程中發生錯誤:")
finally:
    time.sleep(3)
    driver.quit()
    logger.info("WebDriver 已關閉。程式結束。")
    logger.debug("程式全部流程已執行完畢。")