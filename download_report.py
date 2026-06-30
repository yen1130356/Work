import os
import time
from datetime import datetime, date, timedelta
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager

LOGIN_URL = "https://eporter.jiouhung.com/ePorterController_VGHTC/Login.aspx"
REPORT_BASE_URL = "https://eporter.jiouhung.com/ePorterController_VGHTC/reports/frmReports.aspx"

WEB_USER = "mctg37"
WEB_PASSWORD = "mc0037"

DOWNLOAD_DIR = os.path.join(os.getcwd(), "uploads")
os.makedirs(DOWNLOAD_DIR, exist_ok=True)

def date_to_excel_ordinal(dt):
    """ 精準計算 Excel 核心天數序數（對齊 Excel 1900 閏年 Bug） """
    # 💡 核心修正：統一處理 date 與 datetime 物件，防止型別不匹配的 TypeError
    excel_anchor = datetime(1899, 12, 30).date()
    if isinstance(dt, datetime):
        dt = dt.date()
    return int((dt - excel_anchor).days)

def download_yesterday_report(date_from=None, date_to=None):
    # 確保有 uploads 資料夾
    os.makedirs(DOWNLOAD_DIR, exist_ok=True)
    
    # 清理並確保傳入的日期格式正確
    if date_from and date_to:
        try:
            # 相容處理斜線 / 或橫線 -
            clean_from = date_from.strip().replace('-', '/')
            clean_to = date_to.strip().replace('-', '/')
            
            dt_from = datetime.strptime(clean_from, "%Y/%m/%d")
            dt_to = datetime.strptime(clean_to, "%Y/%m/%d")
            
            # 💡 核心修正：起日往前推 1 天，確保能抓到「任務時間」在前一天但「結束時間」在今天的跨夜任務
            dt_from_buffered = dt_from - timedelta(days=1)
            dt_to_buffered = dt_to + timedelta(days=1)
            
            excel_param_start = date_to_excel_ordinal(dt_from_buffered)
            excel_param_end = date_to_excel_ordinal(dt_to_buffered)
            
            print(f"🚗 執行指定區間下載: {clean_from} 至 {clean_to}")
            print(f"📊 網頁實際帶入序數: {excel_param_start} ~ {excel_param_end} (已向外擴展起訖天數對齊跨夜邊界)")
        except Exception as e:
            print(f"❌ 解析傳入日期失敗: {e}，改用預設昨日排程")
            yesterday_obj = date.today() - timedelta(days=1)
            yesterday_start = yesterday_obj - timedelta(days=1)
            yesterday_end = yesterday_obj + timedelta(days=1)
            excel_param_start = date_to_excel_ordinal(yesterday_start)
            excel_param_end = date_to_excel_ordinal(yesterday_end)
            print(f"📊 降級防禦帶入序數: {excel_param_start} ~ {excel_param_end}")
    else:
        # 自動定時任務（預設昨日）
        yesterday_obj = date.today() - timedelta(days=1)
        yesterday_start = yesterday_obj - timedelta(days=1)
        yesterday_end = yesterday_obj + timedelta(days=1)
        excel_param_start = date_to_excel_ordinal(yesterday_start)
        excel_param_end = date_to_excel_ordinal(yesterday_end)
        print(f"⏰ 自動定時任務發動：自動同步昨日數據...")
        print(f"📊 網頁實際帶入序數: {excel_param_start} ~ {excel_param_end}")

    chrome_options = Options()
    chrome_options.add_argument("--headless=new")  # 使用全新強固型 Headless 模式
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    
    prefs = {
        "download.default_directory": DOWNLOAD_DIR,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    }
    chrome_options.add_experimental_option("prefs", prefs)

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
    
    # 🎯 核心修正：允許在 Headless 模式下進行檔案下載
    driver.execute_cdp_cmd("Page.setDownloadBehavior", {
        "behavior": "allow",
        "downloadPath": DOWNLOAD_DIR
    })
    
    try:
        driver.get(LOGIN_URL)
        time.sleep(3)

        print("⏳ 正在輸入憑證並登入...")
        driver.find_element(By.ID, "txtUserName").send_keys(WEB_USER)      
        driver.find_element(By.ID, "txtPassword").send_keys(WEB_PASSWORD)  
        driver.find_element(By.ID, "LoginButton").click()                 
        time.sleep(5) 

        target_url = (
            f"{REPORT_BASE_URL}?ReportFilter=ALL&ReportTitle=TASK%20ALL"
            f"&ReportStartDate={excel_param_start}&ReportEndDate={excel_param_end}"
            f"&ReportStartTime=0000&ReportEndTime=2359&ReportSort=TaskTime&UrgentType=0"
        )
        
        print(f"⏳ 請求大數據報表產出...")
        # 為了保險起見，先清空目錄下殘留的舊未完成下載檔
        for f in os.listdir(DOWNLOAD_DIR):
            if f.endswith(".crdownload"):
                try: os.remove(os.path.join(DOWNLOAD_DIR, f))
                except: pass

        driver.get(target_url)
        
        # 🎯 核心修正：動態輪詢檢測下載狀態，拒絕盲等 60 秒
        print(f"⏳ 進入智能動態檢測，等待 Excel 寫入硬碟...")
        success_download = False
        timeout = 120  # 最多等待 120 秒
        start_time = time.time()
        
        while time.time() - start_time < timeout:
            time.sleep(3)
            files = os.listdir(DOWNLOAD_DIR)
            # 過濾掉暫存檔，確保有真正的 .xlsx 或 .xls 存在
            actual_files = [f for f in files if (f.endswith(".xlsx") or f.endswith(".xls")) and not f.startswith("~$")]
            crdownload_files = [f for f in files if f.endswith(".crdownload")]
            
            if actual_files and not crdownload_files:
                print(f"✅ 偵測到報表檔案成功落盤: {actual_files}")
                success_download = True
                break
        
        if not success_download:
            print("❌ 下載超時或目錄內未產生合法 Excel 檔案。")
            return False
            
        return True

    except Exception as e:
        print(f"❌ 自動化 RPA 下載異常: {str(e)}")
        return False
    finally:
        driver.quit()