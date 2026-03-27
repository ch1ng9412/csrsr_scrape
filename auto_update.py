from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import os

def auto_upload_to_system():
    # ==================== 1. 設定系統與登入資訊 ====================
    LOGIN_URL = "https://2026niag.ncu.edu.tw/Login.aspx?ReturnUrl=%2fuser%2f"  
    UPLOAD_URL = "https://2026niag.ncu.edu.tw/Manager/Advuser/Volunteer/Volunteer_Add_Excel_M.aspx"
    
    USER_ACCOUNT = "Volunteer_Admin_001"
    USER_PASSWORD = "$heKK5391784"
    
    # 自動取得當前資料夾下「整理後_表單志工名單.xls」的絕對路徑 (免除手動填寫的麻煩)
    current_dir = os.getcwd()
    EXCEL_FILE_PATH = os.path.join(current_dir, "整理後_表單志工名單.xls")
    
    # ==================== 2. 網頁元素 ID (根據您提供的 HTML 精準設定) ====================
    # 登入頁面元素
    ACCOUNT_INPUT_ID = "ctl00_ContentPlaceHolder1_LoginUser_UserName"
    PASSWORD_INPUT_ID = "ctl00_ContentPlaceHolder1_LoginUser_Password"
    
    # 匯入頁面元素
    FILE_INPUT_ID = "ctl00_ContentPlaceHolder1_FileUpload3"
    SUBMIT_BUTTON_ID = "ctl00_ContentPlaceHolder1_Btn_Upload_T"

    # 檢查要上傳的檔案是否存在
    if not os.path.exists(EXCEL_FILE_PATH):
        print(f"【錯誤】找不到待上傳的檔案：{EXCEL_FILE_PATH}")
        print("👉 請確認您已經先執行了整理程式，並產出了「整理後_表單志工名單.xls」")
        return

    # ==================== 3. 啟動機器人流程 ====================
    print("正在啟動瀏覽器自動化機器人...")
    service = Service(ChromeDriverManager().install())
    options = webdriver.ChromeOptions()
    options.add_experimental_option("detach", True) 
    driver = webdriver.Chrome(service=service, options=options)

    try:
        # 加入智慧等待機制，最長等待 10 秒
        wait = WebDriverWait(driver, 10)

        # [步驟 A] 打開登入網頁
        print("前往登入頁面...")
        driver.get(LOGIN_URL)
        time.sleep(2) 

        # [步驟 B] 機器人代填帳號密碼
        print("機器人自動輸入帳號與密碼...")
        # 使用 wait.until 確保元素「可互動」才輸入，避免網頁還沒跑完就打字
        account_input = wait.until(EC.element_to_be_clickable((By.ID, ACCOUNT_INPUT_ID)))
        account_input.send_keys(USER_ACCOUNT)
        
        password_input = wait.until(EC.element_to_be_clickable((By.ID, PASSWORD_INPUT_ID)))
        password_input.send_keys(USER_PASSWORD)
        
        # ---------------------------------------------------------
        # 【人機協作斷點 1】交給人類處理驗證碼
        # ---------------------------------------------------------
        print("\n" + "★"*50)
        print("⚠️ 帳密已填妥，接下來請交給您：")
        print("👉 1. 請在彈出的瀏覽器中，手動輸入 5 碼驗證碼。")
        print("👉 2. 點擊網頁上的「登 入」按鈕。")
        print("👉 3. 登入成功後，請回到這個終端機視窗繼續下一步。")
        print("★"*50 + "\n")
        
        # 程式會停在這裡，直到您在終端機按下 Enter 鍵
        input("✅ 確認登入成功後，請在此處按下【Enter 鍵】，機器人將接手上傳工作...")

        # [步驟 C] 機器人接手，前往匯入資料的頁面
        print("接收到指令！機器人接手中，前往資料匯入頁面...")
        driver.get(UPLOAD_URL)
        time.sleep(2)

        # ---------------------------------------------------------
        # 【人機協作斷點 2】第二階段等待：交給人類手動選擇檔案
        # ---------------------------------------------------------
        print("\n" + "★"*50)
        print("⚠️ 已經為您導覽至匯入頁面，接下來的【選擇檔案】步驟請交給您：")
        print(f"👉 準備要上傳的檔案路徑為：\n   {EXCEL_FILE_PATH}")
        print("👉 (建議您可以直接複製上方路徑，貼到選擇檔案的視窗中)")
        print("👉 1. 請在網頁上手動點選並「選擇該 Excel 檔案」。")
        print("★"*50 + "\n")
        
        print("\n【🎉 大功告成】自動點選上傳指令已送出！")
        print("請於瀏覽器視窗中確認網站是否顯示匯入成功的相關訊息。")

    except Exception as e:
        print(f"【系統錯誤】機器人執行過程中發生問題。")
        print(f"詳細報錯原因：{e}")
        print("👉 可能是網頁載入過慢，或是網站的 ID 有臨時變動。")

if __name__ == "__main__":
    auto_upload_to_system()