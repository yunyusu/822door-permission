import gspread
from oauth2client.service_account import ServiceAccountCredentials
import time

# choose from https://developers.google.com/sheets/api/scopes?hl=zh-tw
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

while(1):
    # consider the delay time for processing
    time.sleep(0.5)
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
        
        print("姓名:" + 姓名 + "\n卡號:" + 卡號)
