import pyautogui as ag
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import time, re, pyperclip

# Give the position of button on the interface
# test it by autopygui.position()
##############################################
基本資料設定 = (1646, 93)
使用者 = (1671, 112)
新增 = (2515, 209)
姓名欄 = (2595, 264)
關閉 = (3130, 177)
讀卡機操作 = (1738, 92)
傳送允入資料 = (1767, 221)
##############################################

# The information of the appicants
# ex: 姓名 = "test", 卡號 = "00000000"
# authorization types, see https://developers.google.com/sheets/api/scopes?hl=zh-tw
scopes = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

# give credential(json key from google cloud engine)
creds = ServiceAccountCredentials.from_json_keyfile_name("C:\\Users\\yunyu\\Desktop\\self-learning\\python\\auto_822door_permission\\door-permission-7ca2fd6ae670.json", scopes=scopes)

file = gspread.authorize(creds)
workbook = file.open("測試_822空間借用表單 (回覆)")
sheet = workbook.sheet1
temp_max_row = 1

# make regex for card number
cardNumRegex = re.compile(r'(\d){8}')

# make regex for english names
name_en = re.compile(r'[u4e00-u9fa5]')

# repeatedly detect if there exist new applicants
# press ctrl+c to end the program
while(1):
    # consider the delay time for processing
    # max_col = len(sheet.get_all_values()[0])
    max_row = len(sheet.get_all_values())

# for cell in sheet.range('A1:I'+str(max_col)):
    # print(cell.value)

# for i in range(1, max_col + 1):
#     print(sheet.cell(row=1, col=i).value)

    if temp_max_row < max_row:
        temp_max_row += 1
        姓名 = sheet.cell(temp_max_row, 2).value
        卡號 = sheet.cell(temp_max_row, 8).value
        
        pyperclip.copy(姓名)

        print("姓名:" + 姓名 + "\n卡號:" + 卡號)

        # Step by Step setting process
        if bool(cardNumRegex.search(卡號)):
            ag.click(基本資料設定)
            ag.click(使用者)
            ag.click(新增)
            ag.click(姓名欄)
            ag.typewrite(['backspace'], interval=0.05)
            ag.hotkey('ctrl', 'v')            
            ag.typewrite(['tab', 'tab'], interval=0.05)
            ag.typewrite(卡號, interval=0.05)
            # randomly click somewhere
            ag.click(2700, 400)
            # prevent repetitve data warning window
            time.sleep(0.5)
            ag.press('enter')
            ag.click(關閉)
            time.sleep(0.5)
            ag.press('enter')
            ag.click(讀卡機操作)
            ag.click(傳送允入資料)
            time.sleep(0.5)
            ag.press('enter')

            # email the applicants that the card is now capable of entering 822
            sheet.update_cell(temp_max_row, 10, "0")

        else:
            print("🐼Warning: illegal card number!")
            # email the applicants that something happens to their card number
            # please reach out to 822 teammates for help!
            sheet.update_cell(temp_max_row, 10, "1")
    
    time.sleep(3)
    
