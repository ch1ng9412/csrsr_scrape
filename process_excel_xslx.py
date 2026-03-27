import pandas as pd
import os
import xlrd
import xlwt
from xlutils.copy import copy
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import re  # 【新增】用來過濾電話號碼中的非數字字元
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

# ==============================================================================
# 第一階段：Google 雲端多工作表抓取與資料統合
# ==============================================================================
def process_service_learning_from_cloud():
    """
    從 Google 雲端抓取「服學班級」表單資料（包含所有分頁），
    提取姓名與電話並統合匯入至範例格式中。
    """
    KEY_FILE = 'credentials_2.json' 
    # 更新為服學班級的雲端試算表網址
    SHEET_URL = 'https://docs.google.com/spreadsheets/d/1LwLIWxVzx6xTw6TSkUzXyxm9mH_kW64p7Cq2Wnrwi0E/edit?gid=1270225478#gid=1270225478'
    
    template_filename = '志工整合匯入範例.xls'  
    output_filename = '整理後_服學班級名單.xls'

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
        
        print("成功連線！正在讀取「服學班級」的所有系級分頁...")
        spreadsheet = client.open_by_url(SHEET_URL)
        
        all_data = []
        
        # 動態遍歷試算表中的所有分頁 (Worksheets)
        for sheet in spreadsheet.worksheets():
            try:
                raw_data = sheet.get_all_values()
                if not raw_data or len(raw_data) < 2:
                    continue
                    
                headers = raw_data[0]
                records_data = raw_data[1:]
                df = pd.DataFrame(records_data, columns=headers)

                name_col = None
                phone_col = None
                
                # 【變更 1】尋找該分頁中的「姓名」與「電話」
                for col in df.columns:
                    col_str = str(col).strip()
                    if '姓名' in col_str or '名字' in col_str:
                        name_col = col
                    if '電話' in col_str or '手機' in col_str:
                        phone_col = col

                # 如果該分頁同時擁有這兩個欄位，則將其提取並加入總集合中
                if name_col and phone_col:
                    temp_df = df[[name_col, phone_col]].copy()
                    temp_df.rename(columns={name_col: '姓名', phone_col: '電話'}, inplace=True)
                    all_data.append(temp_df)
                else:
                    print(f"  [略過] 分頁 '{sheet.title}' 找不到姓名或電話欄位。")
                    
            except Exception as inner_e:
                print(f"  [錯誤] 讀取分頁 '{sheet.title}' 時發生錯誤: {inner_e}")
                continue

        if not all_data:
            print("【錯誤】在所有分頁中都找不到可用的資料。")
            return False

        # 統合所有分頁的資料
        final_df = pd.concat(all_data, ignore_index=True)

        # 姓名清理函數
        def clean_name(x):
            if pd.isna(x) or str(x).strip() in ('', 'nan', 'None'):
                return ""
            return str(x).strip()

        # 【變更 2】電話清理函數 (防呆與補0)
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

        final_df['姓名'] = final_df['姓名'].apply(clean_name)
        final_df['電話'] = final_df['電話'].apply(fix_phone)

        # 移除沒有資料的列，並依姓名去重複
        final_df = final_df[final_df['姓名'] != ""]
        final_df.drop_duplicates(subset=['姓名'], keep='first', inplace=True)

        print(f"資料統合完畢，共取得 {len(final_df)} 筆不重複的服學資料，準備匯入範本...")
        
        rb = xlrd.open_workbook(template_filename, formatting_info=True)
        wb = copy(rb) 
        
        sheet_index = 0
        for i, sheet_obj in enumerate(rb.sheets()):
            if sheet_obj.name == 'ExcelData':
                sheet_index = i
                break
                
        ws = wb.get_sheet(sheet_index)
        old_sheet = rb.sheet_by_index(sheet_index)
        
        names = final_df['姓名'].tolist()
        phones = final_df['電話'].tolist()
        
        # 建立純文字儲存格格式，確保電話0開頭不會變形
        text_style = xlwt.easyxf(num_format_str='@')
        max_rows = max(old_sheet.nrows - 2, len(names))
        
        # 【變更 3】將姓名寫入 index 1 (B欄)，電話寫入 index 2 (C欄)
        for i in range(max_rows):
            row_idx = i + 2  
            if i < len(names):
                ws.write(row_idx, 1, names[i], text_style)  # B欄：姓名
                ws.write(row_idx, 2, phones[i], text_style) # C欄：電話
            else:
                ws.write(row_idx, 1, "", text_style)
                ws.write(row_idx, 2, "", text_style)
            
        wb.save(output_filename)
        print(f"【成功】已匯出 {len(final_df)} 筆資料至 '{output_filename}'！\n")
        return True

    except gspread.exceptions.APIError as e:
        print(f"【權限錯誤】機器人帳號沒有這份「服學班級」試算表的權限！")
        print("👉 請確認已將 credentials.json 內的 client_email 加入該試算表的「共用」名單。")
        return False
    except Exception as e:
        print(f"【資料處理錯誤】發生問題：{e}")
        return False


# ==============================================================================
# 第二階段：網頁自動化上傳 (銜接新的 Excel 檔案)
# ==============================================================================
def auto_upload_to_system():
    LOGIN_URL = "https://2026niag.ncu.edu.tw/Login.aspx?ReturnUrl=%2fuser%2f"  
    UPLOAD_URL = "https://2026niag.ncu.edu.tw/Manager/Advuser/Volunteer/Volunteer_Add_Excel_M.aspx"
    
    USER_ACCOUNT = "Volunteer_Admin_001"
    USER_PASSWORD = "$heKK5391784"
    
    # 抓取新產出的「服學班級」整理檔
    current_dir = os.getcwd()
    EXCEL_FILE_PATH = os.path.join(current_dir, "整理後_服學班級名單.xls")
    
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

        print("前往登入頁面並輸入帳密...")
        driver.get(LOGIN_URL)
        time.sleep(2) 

        account_input = wait.until(EC.element_to_be_clickable((By.ID, ACCOUNT_INPUT_ID)))
        account_input.send_keys(USER_ACCOUNT)
        
        password_input = wait.until(EC.element_to_be_clickable((By.ID, PASSWORD_INPUT_ID)))
        password_input.send_keys(USER_PASSWORD)
        
        print("\n" + "★"*50)
        print("⚠️ 帳密已填妥，接下來請交給您：")
        print("👉 1. 請在彈出的瀏覽器中，手動輸入 5 碼驗證碼。")
        print("👉 2. 點擊網頁上的「登 入」按鈕。")
        print("👉 3. 登入成功後，請回到這個終端機視窗繼續下一步。")
        print("★"*50 + "\n")
        
        input("✅ 確認登入成功後，請在此處按下【Enter 鍵】，機器人將接手上傳工作...")

        print("接收到指令！機器人接手中，前往資料匯入頁面...")
        driver.get(UPLOAD_URL)
        time.sleep(2)

        print("\n" + "★"*50)
        print("⚠️ 已經為您導覽至匯入頁面，接下來的【選擇檔案】步驟請交給您：")
        print(f"👉 準備要上傳的檔案路徑為：\n   {EXCEL_FILE_PATH}")
        print("👉 (建議您可以直接複製上方路徑，貼到選擇檔案的視窗中)")
        print("👉 1. 請在網頁上手動點選並「選擇該 Excel 檔案」。")
        print("★"*50 + "\n")
        
        print("【🎉 大功告成】自動點選上傳指令已送出！")
        print("請於瀏覽器視窗中確認網站是否顯示匯入成功的相關訊息。")

    except Exception as e:
        print(f"【系統錯誤】機器人執行過程中發生問題：{e}")

# ==============================================================================
# 主程式執行區塊
# ==============================================================================
if __name__ == "__main__":
    print("="*60)
    print("      🚀 服學班級名單自動化整合與上傳系統啟動")
    print("="*60 + "\n")

    print(">>> 進入第一階段：資料處理 <<<")
    process_success = process_service_learning_from_cloud()

    if process_success:
        print("\n>>> 進入第二階段：網頁自動化上傳 <<<")
        auto_upload_to_system()
    else:
        print("\n【終止執行】因為第一階段資料整理發生錯誤，機器人已取消上傳任務。")