import pandas as pd
import os
import xlrd
import xlwt
from xlutils.copy import copy
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import re
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

# ==============================================================================
# 第一階段：Google 表單資料抓取與處理 (原 process_excel_csv.py)
# ==============================================================================
def process_excel_from_cloud():
    """
    從 Google 雲端抓取表單資料，整理後匯入至系統指定的範例格式中。
    回傳 True 表示執行成功，False 表示失敗。
    """
    KEY_FILE = 'credentials.json' 
    SHEET_URL = 'https://docs.google.com/spreadsheets/d/1jcrxOXlxRkeeRdcU_f_5Mo1MkpCaJghWXXXXXGb1bdg/edit?resourcekey=&gid=2106417132#gid=2106417132'
    template_filename = '志工整合匯入範例.xls'  
    output_filename = '整理後_表單志工名單.xls'

    if not os.path.exists(KEY_FILE):
        print(f"【錯誤】找不到金鑰檔案 '{KEY_FILE}'。")
        return False
        
    if not os.path.exists(template_filename):
        print(f"【錯誤】找不到範例格式檔案 '{template_filename}'。")
        return False

    try:
        print("正在連線至 Google 雲端硬碟...")
        scope = [
            'https://spreadsheets.google.com/feeds',
            'https://www.googleapis.com/auth/drive'
        ]
        creds = ServiceAccountCredentials.from_json_keyfile_name(KEY_FILE, scope)
        client = gspread.authorize(creds)
        
        print("成功連線！正在抓取最新資料...")
        sheet = client.open_by_url(SHEET_URL).sheet1
        
        raw_data = sheet.get_all_values()
        if not raw_data or len(raw_data) < 2:
            print("【警告】雲端試算表中目前沒有任何資料。")
            return False
            
        headers = raw_data[0]
        records_data = raw_data[1:]
        df = pd.DataFrame(records_data, columns=headers)

        name_col = None
        phone_col = None
        
        for col in df.columns:
            col_str = str(col).strip()
            if '姓名' in col_str or '名字' in col_str:
                name_col = col
            if '電話' in col_str or '手機' in col_str:
                phone_col = col

        if not name_col or not phone_col:
            print(f"【錯誤】找不到明確的「姓名」或「電話」欄位。")
            return False

        filtered_df = df[[name_col, phone_col]].copy()
        filtered_df.rename(columns={name_col: '姓名', phone_col: '電話'}, inplace=True)

        def fix_phone(x):
            if pd.isna(x) or str(x).strip() == '':
                return ""
            p = str(x).strip()
            # 強制移除所有非數字字元
            p = re.sub(r'\D', '', p)
            if not p:
                return ""
            if len(p) == 9 and not p.startswith('0'):
                return '0' + p
            return p

        filtered_df['電話'] = filtered_df['電話'].apply(fix_phone)

        filtered_df.drop_duplicates(subset=['姓名'], keep='first', inplace=True)
        filtered_df.dropna(subset=['姓名'], inplace=True)
        filtered_df = filtered_df[filtered_df['姓名'].str.strip() != ""]

        print("資料整理完畢，正在匯入至系統範例格式中...")
        rb = xlrd.open_workbook(template_filename, formatting_info=True)
        wb = copy(rb) 
        
        sheet_index = 0
        for i, sheet_obj in enumerate(rb.sheets()):
            if sheet_obj.name == 'ExcelData':
                sheet_index = i
                break
                
        ws = wb.get_sheet(sheet_index)
        old_sheet = rb.sheet_by_index(sheet_index)
        
        names = filtered_df['姓名'].tolist()
        phones = filtered_df['電話'].tolist()
        
        # 建立「純文字」儲存格樣式 (@)，避免電話變空值
        text_style = xlwt.easyxf(num_format_str='@')
        max_rows = max(old_sheet.nrows - 2, len(names))
        
        for i in range(max_rows):
            row_idx = i + 2  
            if i < len(names):
                ws.write(row_idx, 1, names[i], text_style)   
                ws.write(row_idx, 2, phones[i], text_style)  
            else:
                ws.write(row_idx, 1, "", text_style)
                ws.write(row_idx, 2, "", text_style)
            
        wb.save(output_filename)
        print(f"【成功】已匯出 {len(filtered_df)} 位志工資料至 '{output_filename}'！\n")
        return True

    except Exception as e:
        print(f"【資料處理錯誤】發生問題：{e}")
        return False


# ==============================================================================
# 第二階段：自動化上傳至系統 (原 auto_update.py)
# ==============================================================================
def auto_upload_to_system():
    LOGIN_URL = "https://2026niag.ncu.edu.tw/Login.aspx?ReturnUrl=%2fuser%2f"  
    UPLOAD_URL = "https://2026niag.ncu.edu.tw/Manager/Advuser/Volunteer/Volunteer_Add_Excel_M.aspx"
    
    USER_ACCOUNT = "Volunteer_Admin_001"
    USER_PASSWORD = "$heKK5391784"
    
    current_dir = os.getcwd()
    EXCEL_FILE_PATH = os.path.join(current_dir, "整理後_表單志工名單.xls")
    
    ACCOUNT_INPUT_ID = "ctl00_ContentPlaceHolder1_LoginUser_UserName"
    PASSWORD_INPUT_ID = "ctl00_ContentPlaceHolder1_LoginUser_Password"
    SUBMIT_BUTTON_ID = "ctl00_ContentPlaceHolder1_Btn_Upload_T"

    if not os.path.exists(EXCEL_FILE_PATH):
        print(f"【錯誤】找不到待上傳的檔案：{EXCEL_FILE_PATH}")
        return

    print("正在啟動瀏覽器自動化機器人...")
    service = Service(ChromeDriverManager().install())
    options = webdriver.ChromeOptions()
    options.add_experimental_option("detach", True) 
    driver = webdriver.Chrome(service=service, options=options)

    try:
        wait = WebDriverWait(driver, 10)

        # [步驟 A & B] 打開網頁與輸入帳密
        print("前往登入頁面並輸入帳密...")
        driver.get(LOGIN_URL)
        time.sleep(2) 

        account_input = wait.until(EC.element_to_be_clickable((By.ID, ACCOUNT_INPUT_ID)))
        account_input.send_keys(USER_ACCOUNT)
        
        password_input = wait.until(EC.element_to_be_clickable((By.ID, PASSWORD_INPUT_ID)))
        password_input.send_keys(USER_PASSWORD)
        
        # [斷點 1] 處理驗證碼
        print("\n" + "★"*50)
        print("⚠️ 帳密已填妥，接下來請交給您：")
        print("👉 1. 請在彈出的瀏覽器中，手動輸入 5 碼驗證碼。")
        print("👉 2. 點擊網頁上的「登 入」按鈕。")
        print("👉 3. 登入成功後，請回到這個終端機視窗繼續下一步。")
        print("★"*50 + "\n")
        
        input("✅ 確認登入成功後，請在此處按下【Enter 鍵】，機器人將接手上傳工作...")

        # [步驟 C] 前往匯入頁面
        print("接收到指令！機器人接手中，前往資料匯入頁面...")
        driver.get(UPLOAD_URL)
        time.sleep(2)

        # [斷點 2] 手動選擇檔案
        print("\n" + "★"*50)
        print("⚠️ 已經為您導覽至匯入頁面，接下來的【選擇檔案】步驟請交給您：")
        print(f"👉 準備要上傳的檔案路徑為：\n   {EXCEL_FILE_PATH}")
        print("👉 (建議您可以直接複製上方路徑，貼到選擇檔案的視窗中)")
        print("👉 1. 請在網頁上手動點選並「選擇該 Excel 檔案」。")
        print("★"*50 + "\n")
        
        print("\n【🎉 大功告成】自動點選上傳指令已送出！")
        print("請於瀏覽器視窗中確認網站是否顯示匯入成功的相關訊息。")

    except Exception as e:
        print(f"【系統錯誤】機器人執行過程中發生問題：{e}")


# ==============================================================================
# 主程式執行區塊
# ==============================================================================
if __name__ == "__main__":
    print("="*60)
    print("      🚀 志工名單自動化整合與上傳系統啟動")
    print("="*60 + "\n")

    # 執行第一階段：抓取與整理
    print(">>> 進入第一階段：資料處理 <<<")
    process_success = process_excel_from_cloud()

    # 若第一階段成功，則執行第二階段：自動上傳
    if process_success:
        print("\n>>> 進入第二階段：網頁自動化上傳 <<<")
        auto_upload_to_system()
    else:
        print("\n【終止執行】因為第一階段資料整理發生錯誤，機器人已取消上傳任務。")