import glob
import gspread
from oauth2client.service_account import ServiceAccountCredentials

scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']

# 秘密鍵（JSONファイル）のファイル名を入力
credentials = ServiceAccountCredentials.from_json_keyfile_name('my-project-sample-292308-cf3de616e9f6.json', scope)
gc = gspread.authorize(credentials)


# ②「キー」で取得
SPREADSHEET_KEY = '1ZyomcKNjDzlruDXNiVvjqUM-LWUaZPLy95160MeTyFM'

wb = gc.open_by_key(SPREADSHEET_KEY)
#print(wb)


# ①「名前」で取得
ws = wb.get_worksheet(0)
#print(ws)


f = open('/Users/mizumuratomoya/Desktop/gspread/pdf.txt', 'r')
data = f.read()
#print(data)
s = data.replace('（軽）','').replace('　', ' ').replace('差出人:', '').replace('株式会社 開陽', '').replace('件名: ', '')
#print(s)
s2 = s.replace('【注文情報】', '').replace('ご注文(カード)', '').replace('日付: ', '').replace('宛先: ', '')
s3 = s2.replace('※ 配送方法 : ', '').replace('※ 送料 : ', '').replace('-----------------------------------------------------------', '')
s4 = s3.replace('は軽減税率対象であることを示します。', '').replace('【決済方法】', '').replace('【注文者情報】', '').replace('━━━━━━ 【商品のお届け先】━━━━━━', '')
s5 = s4.replace('━━━━━━━━━━━━━━━━━━━━━', '')
data1 = s5.splitlines()

f.close()

#print(data1)
list_example4 = [a for a in data1 if a != '']
#print(list_example4)

data2 = len(list_example4)
    #print(data2)
ws.update_acell('A4', data2)

for i in range(data2):
    cell = 'A' + str(i+6)
    print(cell)
    ws.update_acell(cell, list_example4[i])

#数回に一回バグが出る
