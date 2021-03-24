import gspread
import json
#ServiceAccountCredentials：Googleの各サービスへアクセスできるservice変数を生成します。
from oauth2client.service_account import ServiceAccountCredentials 
#2つのAPIを記述しないとリフレッシュトークンを3600秒毎に発行し続けなければならない
scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
#認証情報設定
#ダウンロードしたjsonファイル名をクレデンシャル変数に設定（秘密鍵、Pythonファイルから読み込みしやすい位置に置く）
credentials = ServiceAccountCredentials.from_json_keyfile_name('/Users/masa/dys-mercari/mercariaddress/mercari-284420-8a48bc1a3e9b.json', scope)
#OAuth2の資格情報を使用してGoogle APIにログインします。
gc = gspread.authorize(credentials)
#共有設定したスプレッドシートキーを変数[SPREADSHEET_KEY]に格納する。
SPREADSHEET_KEY = '13H4m38CC7k0-VnhkZZolPu5D1PY4Ekb1NG_-JGL2ZvQ'
#共有設定したスプレッドシートのシート1を開く
worksheet = gc.open_by_key(SPREADSHEET_KEY).sheet1

from selenium import webdriver
from selenium.webdriver import ActionChains
from time import sleep
import webbrowser
from webdriver_manager.chrome import ChromeDriverManager

from datetime import datetime as dt
import datetime
import requests

worksheet.clear() #ワークシートの値を全てクリアする

def mersta(status, lists):
   i = 0
   for elm in elms:
      sta = status[i].text
      if '発送待ち' in sta:
         elm = elms[i].get_attribute("href")
         lists.append(elm)
      i += 1
   return lists

def mercon_values():
   values_list = worksheet.get_all_values() #worksheetの全てのセルの情報を取得
   i = 0
   for minf in lists:
      driver.get(lists[i])
      print('41行目', 'i', i)
      profit = driver.find_elements_by_css_selector("ul.transact-info-table-cell") #クラス名を指定して要素を特定（住所）
      mitm = driver.find_elements_by_css_selector(".transact-info-item.bold") #クラス名を取得して要素を特定（商品名）
      lastrow = len(values_list) + 1 + i #空白の行を取得する（取得を開始する行）
      worksheet.update_cell(lastrow, 1, profit[6].text)
      print(profit[6].text)
      worksheet.update_cell(lastrow, 2, "M " + mitm[0].text)
      print(mitm[0].text)
      if "大" in mitm[0].text:
         worksheet.update_cell(lastrow, 3, 1)
      else:
         worksheet.update_cell(lastrow, 3, 0)
      i += 1
      sleep(1)
try:
   time = datetime.datetime.strftime(dt.now(), '%m月%d日%H時%M分')
   worksheet.update_cell(1, 1, time)
   print('account_1', dt.now())
   options = webdriver.ChromeOptions()
   options.add_argument(
      '--user-data-dir={chrom_dir_path}'.format(chrom_dir_path = '/Users/masa/Library/Application Support/Google/Chrome/Profile 4'))
   options.add_argument('--headless')
   options.add_argument('--no-sandbox')
   options.add_argument('--disable-gpu')
   driver = webdriver.Chrome(ChromeDriverManager().install(), options=options)
   # driver.minimize_window()
   sleep(1)
   print('mercari')
   driver.get("https://www.mercari.com/jp/mypage/listings/in_progress/") # メルカリにアクセス

   # res = requests.get("https://www.mercari.com/jp/mypage/listings/in_progress/")
   # print(res.text)
   elms = driver.find_elements_by_class_name("mypage-item-link") #要素のクラス名を指定して取引ページの発送待ちのみURLを取得
   status = driver.find_elements_by_class_name("mypage-item-body")
   lists = []
   mersta(status, lists)
   mercon_values()
   print('yafuoku')
   sleep(1)
   driver.get('https://auctions.yahoo.co.jp/closeduser/jp/show/mystatus?select=closed&hasWinner=1')
   yoms3 = driver.find_elements_by_class_name("decTxt03") #発送状態のURLを取得する
   yoms1 = driver.find_elements_by_class_name("decTxt01")
   i = 0
   lists = [] 
   for yom3 in yoms3:
      msg3 = yom3.find_elements_by_css_selector("a")[0]
      msgt3 = msg3.text
      if "支払いが完了しました" in msgt3:
         url = msg3.get_attribute("href")
         lists.append(url)
      i += 1
   for yom1 in yoms1:
      msg1 = yom1.find_elements_by_css_selector("a")[0]
      msgt1 = msg1.text
      if "取引" in msgt1:
         url = msg1.get_attribute("href")
         lists.append(url)
      i += 1
   #取得したリンクの住所を取得
   i = 0
   j = 0
   values_list = worksheet.get_all_values() #worksheetの全てのセルの情報を取得
   for yinf in lists:
      driver.get(lists[i])
      lastrow = len(values_list) + 1 + j #空白の行を取得する（取得を開始する行）
      yname = driver.find_elements_by_css_selector(".decCnfWr") #落札者の情報を取得
      yitem = driver.find_elements_by_css_selector(".decItmName")
      ypri = driver.find_elements_by_css_selector(".decPrice")
      yspst = driver.find_elements_by_class_name("acMdStatusImage__chargeText")
      if yspst[2].text == "あなた":
         worksheet.update_cell(lastrow, 1, yname[2].text + '\n' + yname[3].text + '\n' + yname[1].text + " 様")
         print(yname[2].text + '\n' + yname[3].text + '\n' + yname[1].text + " 様")
         worksheet.update_cell(lastrow, 2, "Y " + yitem[0].text)
         print(yitem[0].text)
         if "大" in yitem[0].text:
            worksheet.update_cell(lastrow, 3, 1)
         else:
            worksheet.update_cell(lastrow, 3, 0)
         j += 1
      i += 1
      sleep(1)
   print('rakuma')
   sleep(1)
   driver.get("https://fril.jp/sell#trading")
   sleep(1)
   relms = driver.find_elements_by_class_name("media") # 取引ページの発送待ちのみURLを取得
   i = 0
   lists = []
   for relm in relms:
      relmt = relm.text
      if "商品の発送" in relmt:
         rmsga = relm.find_element_by_css_selector("a")
         url = rmsga.get_attribute("href")
         lists.append(url)
      i += 1
   i = 0
   print('r-1')
   values_list = worksheet.get_all_values() #worksheetの全てのセルの情報を取得
   for rinf in lists:
      driver.get(lists[i])
      lastrow = len(values_list) + 1 + i #空白の行を取得する（取得を開始する行）
      profit = driver.find_elements_by_css_selector('.col.s12') #住所を取得
      ritm = driver.find_elements_by_css_selector('.item_name') #商品名を取得
      lastrow = len(values_list) + 1 + i #空白の行を取得する（取得を開始する行）
      worksheet.update_cell(lastrow, 1, profit[7].text)
      print(profit[7].text)
      worksheet.update_cell(lastrow, 2, "R " + ritm[0].text)
      print(ritm[0].text)
      if "大" in ritm[0].text:
         worksheet.update_cell(lastrow, 3, 1)
      else:
         worksheet.update_cell(lastrow, 3, 0)
      i += 1
      sleep(1)
   sleep(1)
   driver.quit() #ウインドウを閉じる
   print('account_2')
   options = webdriver.ChromeOptions()
   options.add_argument(
      '--user-data-dir={chrom_dir_path}'.format(chrom_dir_path = '/Users/masa/Library/Application Support/Google/Chrome/Profile 5'))
   options.add_argument('--headless')
   options.add_argument('--no-sandbox')
   options.add_argument('--disable-gpu')
   driver = webdriver.Chrome(options=options, executable_path='/Users/masa/bin/chromedriver')
   print('mercari')
   driver.get('https://www.mercari.com/jp/mypage/listings/in_progress/')
   elms = driver.find_elements_by_class_name("mypage-item-link") # 取引ページの発送待ちのみURLを取得
   status = driver.find_elements_by_class_name("mypage-item-body")
   lists = []
   print('m2-1')
   mersta(status, lists)
   mercon_values()
   sleep(1)
   driver.quit() #ウインドウを閉じる
   print('last', dt.now())
except Exception as e:
   print(e)