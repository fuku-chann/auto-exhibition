import gspread
import json

#ServiceAccountCredentials：Googleの各サービスへアクセスできるservice変数を生成します。
from oauth2client.service_account import ServiceAccountCredentials 

#2つのAPIを記述しないとリフレッシュトークンを3600秒毎に発行し続けなければならない
scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']

#認証情報設定
#ダウンロードしたjsonファイル名をクレデンシャル変数に設定（秘密鍵、Pythonファイルから読み込みしやすい位置に置く）
credentials = ServiceAccountCredentials.from_json_keyfile_name('/Users/masa/dys-mercari/auto-exhibition/mercari-284420-8a48bc1a3e9b.json', scope)

#OAuth2の資格情報を使用してGoogle APIにログインします。
gc = gspread.authorize(credentials)

#共有設定したスプレッドシートキーを変数[SPREADSHEET_KEY]に格納する。
SPREADSHEET_KEY = '1EPa-1zsfxXissFWV3BigpHsJCyBuiQHnvfucSWmv1FE'

#共有設定したスプレッドシートのシート1を開く
worksheet = gc.open_by_key(SPREADSHEET_KEY).sheet1
worksheet2 = gc.open_by_key(SPREADSHEET_KEY).worksheet('sales')

from selenium import webdriver
from selenium.webdriver import ActionChains
from time import sleep
import webbrowser
from webdriver_manager.chrome import ChromeDriverManager

from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.select import Select
from datetime import datetime as dt
import datetime
from selenium.webdriver.common.alert import Alert
import re

import item_def

options = webdriver.ChromeOptions()
options.add_argument(
   '--user-data-dir={chrom_dir_path}'.format(chrom_dir_path = '/Users/masa/Library/Application Support/Google/Chrome/Profile 3'))
# options.add_argument('--headless')
# options.add_argument('--no-sandbox')
# options.add_argument('--disable-gpu')
# options.add_argument('--disable-dev-shm-usage')
# options.add_argument('--window-size=1200x600')
driver = webdriver.Chrome(ChromeDriverManager().install(), options=options)
# driver.minimize_window()
# driver.get("https://www.mercari.com/jp/mypage/listings/listing/") # メルカリにアクセス
# sleep(1)
exhibit_col = 4
item_def.abc(driver, exhibit_col)
sleep(1)

try:
   print('rakuma')
   def rexp1():
      rinputElement = driver.find_elements_by_class_name("form-group")[1].find_element_by_css_selector("input")
      rinputElement.send_keys(itm)
      print(329, rinputElement.get_attribute("value"))
      rinputElement3 = driver.find_elements_by_class_name("form-group")[2].find_element_by_css_selector("textarea")
      if ('ダイソン' in rinputElement.get_attribute("value")) and ('大' in rinputElement.get_attribute("value")) and ('スペーサー' in rinputElement.get_attribute("value")):
         print(348)
         rinputElement3.send_keys(rexp1a + item_def.exp3a + item_def.exp0a)
      elif ('ダイソン' in rinputElement.get_attribute("value")) and ('大' in rinputElement.get_attribute("value")):
         print(348)
         rinputElement3.send_keys(rexp1a + item_def.exp4a + item_def.exp0a)
      elif "ダイソン" in rinputElement.get_attribute("value"):
         sleep(1)
         rinputElement3.send_keys(rexp1a + item_def.exp2a + item_def.exp0a)
      else:
         rinputElement3.send_keys(rexp1a)
   #ダイソン共通  rakumacommon_to_dyson: 
   def rcom_to_dys(i):
      rcateElement1 = driver.find_element_by_id("category_name")
      rcateElement1.click()
      rcateElement2 = driver.find_elements_by_class_name("panel.list-group")[8]
      rcateElement2.click()
      sleep(1)
      rcateElement3 = driver.find_elements_by_class_name("panel.list-group")[8].find_elements_by_class_name("list-group-item.small.branch")[9]
      rcateElement3.click()
      sleep(1)
      rcateElement4 = driver.find_elements_by_class_name("panel.list-group")[8].find_elements_by_class_name("sublinks.collapse.in")[1].find_elements_by_class_name("list-group-item.small.leaf")[6]
      rcateElement4.click()
      rcateElement5 = driver.find_elements_by_class_name("btn.btn-select")[1]
      rcateElement5.click()
      sleep(1)
      rcateElement6 = driver.find_elements_by_class_name("form-group")[0].find_element_by_css_selector("input")
      rcateElement6.send_keys("dyson")
      sleep(1)
      rcateElement7 = driver.find_elements_by_class_name("brand-alphabet-name")[0]
      rcateElement7.click()
      rinputElement8 = driver.find_elements_by_class_name("form-group")[11].find_element_by_css_selector("input")
      rinputElement8.send_keys(worksheet.cell(i, 7).value)
   def rclick():
      rcateElement9 = driver.find_element_by_class_name("col-md-offset-4.col-md-5.col-xs-12")
      rcateElement9.click()
      sleep(2)
      rcateElement10 = driver.find_element_by_class_name("col-md-8.col-xs-8")
      sleep(2)
      rcateElement10.click()
      sleep(10)
      driver.get("https://fril.jp/item/new")
      sleep(5)
   driver.get("https://fril.jp/sell")
   sleep(5)
   print('E-1', dt.now())
   time = dt.now()
   a = len(worksheet.col_values(2))
   for i in range(5, a + 1):
      print('i', i, '112')
      sleep(1)
      if dt.strptime(worksheet.cell(i, 3).value, '%Y/%m/%d %H:%M:%S') < dt.strptime(worksheet.cell(3, 2).value, '%Y/%m/%d %H:%M:%S'): #出品してから１日以上経過している場合、trueになります。
         del_item = worksheet.cell(i, 2).value #該当のサンプル名を代入
         urlitems = driver.find_elements_by_class_name("information-pane")[0].find_elements_by_class_name("media") #親要素を取得
         for urlitem in urlitems:
            itemr = urlitem.find_element_by_class_name("media-heading") #商品名を取得
            comr = urlitem.find_elements_by_class_name("item-action-count")[1] #コメントの数を取得
            sleep(1)
            if del_item == itemr.text and comr.text == "0":
               print('344')
               urlr = urlitem.find_elements_by_class_name("btn.btn-default")[1] #削除ボタンを取得
               sleep(5)
               urlr.click()
               sleep(10)
               Alert(driver).accept() #削除しますか？のはいをクリックして削除
               sleep(5)
               break
   print('354')
   ritmlists = driver.find_elements_by_class_name("media-heading")
   print('J-1', dt.now())
   rlists=[]
   for ritmlist in ritmlists:
      rlists.append(ritmlist.text)
   driver.get("https://fril.jp/item/new")
   sleep(3)
   i = 5
   itm = worksheet.cell(i, 2).value #"ダイソン掃除機 タイヤ4個+スペーサー+トルクスドライバー3本セット"
   if not itm in rlists:
      # price = worksheet.cell(i, 7).value
      driver.find_elements_by_xpath("//input[@type='file']")[0].send_keys("/Users/masa/Pictures/torx/torxset.JPG")
      driver.find_elements_by_xpath("//input[@type='file']")[1].send_keys("/Users/masa/Pictures/dyson/tire4.png")
      driver.find_elements_by_xpath("//input[@type='file']")[2].send_keys("/Users/masa/Pictures/dyson/spacers.png")
      driver.find_elements_by_xpath("//input[@type='file']")[3].send_keys("/Users/masa/Pictures/torx/torxdriver.jpeg")
      rexp1a = "【商品紹介】\n画像の2～4枚目が商品になります。ダイソンモーターヘッドの裏側に付いているタイヤと工具です。"
      rexp1()
      rcom_to_dys(i)
      rclick()
      worksheet.update_cell(i, 3, str(time))
   i = 6
   itm = worksheet.cell(i, 2).value #"ダイソン掃除機 タイヤ4個+スペーサー+シャフト4本セット"
   if not itm in rlists:
      # price = worksheet.cell(i, 7).value
      driver.find_elements_by_xpath("//input[@type='file']")[0].send_keys("/Users/masa/Pictures/dyson/shaftset.JPG")
      driver.find_elements_by_xpath("//input[@type='file']")[1].send_keys("/Users/masa/Pictures/dyson/tire4.png")
      driver.find_elements_by_xpath("//input[@type='file']")[2].send_keys("/Users/masa/Pictures/dyson/spacers.png")
      driver.find_elements_by_xpath("//input[@type='file']")[3].send_keys("/Users/masa/Pictures/dyson/shaft.png")
      rexp1a = "【商品紹介】\n画像の2～4枚目が商品になります。ダイソンモーターヘッドの裏側に付いているタイヤと部品です。"
      rexp1()
      rcom_to_dys(i)
      rclick()
      worksheet.update_cell(i, 3, str(time))
   i = 7
   itm = worksheet.cell(i, 2).value #"ダイソン掃除機 タイヤ4個+スペーサー+シャフト4本+トルクスドライバー3本"
   if not itm in rlists:
      # price = worksheet.cell(i, 7).value
      driver.find_elements_by_xpath("//input[@type='file']")[0].send_keys("/Users/masa/Pictures/dyson/4set.JPG")
      driver.find_elements_by_xpath("//input[@type='file']")[1].send_keys("/Users/masa/Pictures/dyson/spacers.jpg")
      driver.find_elements_by_xpath("//input[@type='file']")[2].send_keys("/Users/masa/Pictures/dyson/shaft.png")
      driver.find_elements_by_xpath("//input[@type='file']")[3].send_keys("/Users/masa/Pictures/torx/torxdriver.jpeg")
      rexp1a = "【商品紹介】\n画像の2～4枚目が商品になります。ダイソンモーターヘッドの裏側に付いているタイヤと部品と工具です。"
      rexp1()
      rcom_to_dys(i)
      rclick()
      worksheet.update_cell(i, 3, str(time))
   print('J-5', dt.now())
   i = 8
   itm = worksheet.cell(i, 2).value #"ダイソン掃除機 タイヤ4個+スペーサーセット"
   if not itm in rlists:
      # price = worksheet.cell(i, 7).value
      driver.find_elements_by_xpath("//input[@type='file']")[0].send_keys("/Users/masa/Pictures/dyson/tire4set.JPG")
      driver.find_elements_by_xpath("//input[@type='file']")[1].send_keys("/Users/masa/Pictures/dyson/tire4.png")
      driver.find_elements_by_xpath("//input[@type='file']")[2].send_keys("/Users/masa/Pictures/dyson/spacers.png")
      driver.find_elements_by_xpath("//input[@type='file']")[3].send_keys("/Users/masa/Pictures/dyson/morterhead.jpeg")
      rexp1a = "【商品紹介】\n画像の2～3枚目が商品になります。ダイソンモーターヘッドの裏側に付いているタイヤです。"
      rexp1()
      rcom_to_dys(i)
      rclick()
      worksheet.update_cell(i, 3, str(time))
   print('J-6', dt.now())
   i = 9
   itm = worksheet.cell(i, 2).value #"ダイソン掃除機 タイヤ2個+スペーサーセット" 
   if not itm in rlists:
      # price = worksheet.cell(i, 7).value
      driver.find_elements_by_xpath("//input[@type='file']")[0].send_keys("/Users/masa/Pictures/dyson/tire2set.JPG")
      driver.find_elements_by_xpath("//input[@type='file']")[1].send_keys("/Users/masa/Pictures/dyson/tire2.jpeg")
      driver.find_elements_by_xpath("//input[@type='file']")[2].send_keys("/Users/masa/Pictures/dyson/spacers.png")
      driver.find_elements_by_xpath("//input[@type='file']")[3].send_keys("/Users/masa/Pictures/dyson/morterhead.jpeg")
      rexp1a = "【商品紹介】\n画像の2～3枚目が商品になります。ダイソンモーターヘッドの裏側に付いているタイヤです。"
      rexp1()
      rcom_to_dys(i)
      rclick()
      worksheet.update_cell(i, 3, str(time))
   i = 10
   itm = worksheet.cell(i, 2).value #"トルクスドライバー3本セット（T10 & T8 & T6）"
   if not itm in rlists:
      # price = worksheet.cell(i, 7).value
      driver.find_elements_by_xpath("//input[@type='file']")[0].send_keys("/Users/masa/Pictures/torx/torxdriver.jpeg")
      rinputElement = driver.find_elements_by_class_name("form-group")[1].find_element_by_css_selector("input")
      rinputElement.send_keys(itm)
      rinputElement3 = driver.find_elements_by_class_name("form-group")[2].find_element_by_css_selector("textarea")
      rinputElement3.send_keys("トルクスドライバー3本セットです。（T10 & T8 & T6) ダイソン分解清掃などにお得な3本セットです。")
      rcom_to_dys(i)
      rclick()
      worksheet.update_cell(i, 3, str(time))
   i = 13
   itm = worksheet.cell(i, 2).value #"ダイソン掃除機 タイヤ4個(大)+スペーサーセット"
   if not itm in rlists:
      # price = worksheet.cell(i, 7).value
      driver.find_elements_by_xpath("//input[@type='file']")[0].send_keys("/Users/masa/Pictures/dyson/tire4(large)set.JPG")
      driver.find_elements_by_xpath("//input[@type='file']")[1].send_keys("/Users/masa/Pictures/dyson/tire41mm.jpg")
      driver.find_elements_by_xpath("//input[@type='file']")[2].send_keys("/Users/masa/Pictures/dyson/spacer(縦大).png")
      driver.find_elements_by_xpath("//input[@type='file']")[3].send_keys("/Users/masa/Pictures/dyson/tireimage.png")
      rexp1a = "【商品紹介】\n画像の2～3枚目が商品になります。ダイソンモーターヘッドの裏側に付いているタイヤです。"
      rexp1()
      rcom_to_dys(i)
      rclick()
      worksheet.update_cell(i, 3, str(time))
   i = 14
   itm = worksheet.cell(i, 2).value #"ダイソン掃除機 タイヤ2個(大)+スペーサーセット"
   if not itm in rlists:
      # price = worksheet.cell(i, 7).value
      driver.find_elements_by_xpath("//input[@type='file']")[0].send_keys("/Users/masa/Pictures/dyson/tire2(large)set.JPG")
      driver.find_elements_by_xpath("//input[@type='file']")[1].send_keys("/Users/masa/Pictures/dyson/tire21mm.jpg")
      driver.find_elements_by_xpath("//input[@type='file']")[2].send_keys("/Users/masa/Pictures/dyson/spacer(縦大).png")
      driver.find_elements_by_xpath("//input[@type='file']")[3].send_keys("/Users/masa/Pictures/dyson/tireimage.png")
      rexp1a = "【商品紹介】\n画像の2～3枚目が商品になります。ダイソンモーターヘッドの裏側に付いているタイヤです。"
      rexp1()
      rcom_to_dys(i)
      rclick()
      worksheet.update_cell(i, 3, str(time))
   i = 15
   itm = worksheet.cell(i, 2).value #"ダイソン掃除機 タイヤ4個(大)+スペーサー+トルクスドライバー3本セット"
   if not itm in rlists:
      # price = worksheet.cell(i, 7).value
      driver.find_elements_by_xpath("//input[@type='file']")[0].send_keys("/Users/masa/Pictures/torx/torxset(large)2.JPG")
      driver.find_elements_by_xpath("//input[@type='file']")[1].send_keys("/Users/masa/Pictures/dyson/tire41mm.jpg")
      driver.find_elements_by_xpath("//input[@type='file']")[2].send_keys("/Users/masa/Pictures/dyson/spacer(縦大).png")
      driver.find_elements_by_xpath("//input[@type='file']")[3].send_keys("/Users/masa/Pictures/torx/torxdriver.jpeg")
      rexp1a = "【商品紹介】\n画像の2～4枚目が商品になります。ダイソンモーターヘッドの裏側に付いているタイヤと工具です。"
      rexp1()
      rcom_to_dys(i)
      rclick()
      worksheet.update_cell(i, 3, str(time))
   i = 16
   itm = worksheet.cell(i, 2).value #"ダイソン掃除機 タイヤ4個(大)+トルクスドライバー3本セット"
   if not itm in rlists:
      # price = worksheet.cell(i, 7).value
      driver.find_elements_by_xpath("//input[@type='file']")[0].send_keys("/Users/masa/Pictures/torx/torxset(soft).jpeg")
      driver.find_elements_by_xpath("//input[@type='file']")[1].send_keys("/Users/masa/Pictures/dyson/tire41mm.jpg")
      driver.find_elements_by_xpath("//input[@type='file']")[2].send_keys("/Users/masa/Pictures/torx/torxdriver.jpeg")
      driver.find_elements_by_xpath("//input[@type='file']")[3].send_keys("/Users/masa/Pictures/dyson/soft.jpeg")
      rexp1a = "【商品紹介】\n画像の2～3枚目が商品になります。ダイソン掃除機ソフトローラークリーンヘッドの裏側に付いているタイヤと工具です。"
      rexp1()
      rcom_to_dys(i)
      rclick()
      worksheet.update_cell(i, 3, str(time))
   i = 17
   itm = worksheet.cell(i, 2).value #"ダイソン掃除機 タイヤ4個(大)+シャフト4本セット"
   if not itm in rlists:
      price = worksheet.cell(i, 7).value
      driver.find_elements_by_xpath("//input[@type='file']")[0].send_keys("/Users/masa/Pictures/dyson/shaft4set(soft).JPG")
      driver.find_elements_by_xpath("//input[@type='file']")[1].send_keys("/Users/masa/Pictures/dyson/tire41mm.jpg")
      driver.find_elements_by_xpath("//input[@type='file']")[2].send_keys("/Users/masa/Pictures/dyson/shaft4(soft).png")
      driver.find_elements_by_xpath("//input[@type='file']")[3].send_keys("/Users/masa/Pictures/dyson/soft.jpeg")
      rexp1a = "【商品紹介】\n画像の2～3枚目が商品になります。ダイソン掃除機ソフトローラークリーンヘッドの裏側に付いているタイヤと部品です。"
      rexp1()
      rcom_to_dys(i)
      rclick()
      worksheet.update_cell(i, 3, str(time))
   i = 18
   itm = worksheet.cell(i, 2).value #"ダイソン掃除機 タイヤ2個(大)+シャフト2本セット"
   if not itm in rlists:
      # price = worksheet.cell(i, 7).value
      driver.find_elements_by_xpath("//input[@type='file']")[0].send_keys("/Users/masa/Pictures/dyson/shaft2set(soft).JPG")
      driver.find_elements_by_xpath("//input[@type='file']")[1].send_keys("/Users/masa/Pictures/dyson/tire21mm.jpg")
      driver.find_elements_by_xpath("//input[@type='file']")[2].send_keys("/Users/masa/Pictures/dyson/shaft3.png")
      driver.find_elements_by_xpath("//input[@type='file']")[3].send_keys("/Users/masa/Pictures/dyson/soft.jpeg")
      rexp1a = "【商品紹介】\n画像の2～3枚目が商品になります。ダイソン掃除機ソフトローラークリーンヘッドの裏側に付いているタイヤと部品です。"
      rexp1()
      rcom_to_dys(i)
      rclick()
      worksheet.update_cell(i, 3, str(time))
   i = 19
   itm = worksheet.cell(i, 2).value #"ダイソン掃除機 タイヤ2個(大)+シャフト4本セット+トルクスドライバー3本セット"
   if not itm in rlists:
      # price = worksheet.cell(i, 7).value
      driver.find_elements_by_xpath("//input[@type='file']")[0].send_keys("/Users/masa/Pictures/dyson/tire4(large)3set.JPG")
      driver.find_elements_by_xpath("//input[@type='file']")[1].send_keys("/Users/masa/Pictures/dyson/tire41mm.jpg")
      driver.find_elements_by_xpath("//input[@type='file']")[2].send_keys("/Users/masa/Pictures/dyson/shaft4(soft).png")
      driver.find_elements_by_xpath("//input[@type='file']")[3].send_keys("/Users/masa/Pictures/torx/torxdriver.jpeg")
      rexp1a = "【商品紹介】\n画像の2～3枚目が商品になります。ダイソン掃除機ソフトローラークリーンヘッドの裏側に付いているタイヤと部品と工具のセットです。"
      rexp1()
      rcom_to_dys(i)
      rclick()
      worksheet.update_cell(i, 3, str(time))
   print('J-6', dt.now())
   print("ヤフオク")
   driver.get("https://auctions.yahoo.co.jp/closeduser/jp/show/mystatus") # ヤフオクにアクセス
   sleep(1)
   itm_fin = driver.find_element_by_class_name("decMainLi.decLiLast")
   itm_fin.click()
   sleep(3)
   itm_fin1 = driver.find_elements_by_class_name("wrapper")[1].find_elements_by_tag_name("a")[0]
   itm_fin1.click()
   sleep(3)
   print(318)
   try:
      itm_fins = driver.find_elements_by_class_name("wrapper")[1].find_elements_by_tag_name("table")[10]
   except Exception as e:
      print('e', e)
      print('終了商品がありません')
      driver.quit()
   print(320)
   # print('itm_fins', itm_fins)
   # print('itm_fins.text', itm_fins.text)
   i = 0
   print(322)
   itm_fins_counts = itm_fins.find_elements_by_tag_name("tr")
   print(324)
   # print("itm_fins_counts", itm_fins_counts)
   dicty = {}
   for itm_fin_count in itm_fins_counts:
      print("i", i)
      sleep(1)
      if i > 0:
         item = itm_fin_count.text
         link = itm_fin_count.find_elements_by_tag_name("td")[5].find_element_by_tag_name("a").get_attribute("href")
         dicty[item]=link
      i += 1
   print("dicty", dicty)
   def reproduce():
      sleep(3)
      try:
         print(337)
         driver.find_element_by_xpath("//*[text()=\"出品を続ける\"]").click()
      except Exception as e:
         print('e', e)
         print("出品を続けるボタンがない時")
      print("rpro-1")
      itm_fin3 = driver.find_element_by_class_name("Button.Button--proceed.Inquiry__button.Inquiry__buttonText")
      itm_fin3.click()
      sleep(5)
      print("rpro-2")
      itm_fin4 = driver.find_element_by_class_name("u-displayBlock.u-fontSize20.u-textNormal")
      itm_fin4.click()
   for repro_url in dicty.values():
      driver.get(repro_url)
      sleep(3)
      print(338)
      reproduce()
   def reproduction():
      driver.get("https://auctions.yahoo.co.jp/closeduser/jp/show/mystatus") # ヤフオクにアクセス
      sleep(1)
      itm_fin = driver.find_element_by_class_name("decMainLi.decLiLast")
      itm_fin.click()
      sleep(3)
      itm_fin1 = driver.find_elements_by_class_name("wrapper")[1].find_elements_by_tag_name("a")[0]
      itm_fin1.click()
      sleep(3)
      global itm_fins
      itm_fins = driver.find_elements_by_class_name("wrapper")[1].find_elements_by_tag_name("a")
      i = 12
   reproduction()
   count_itm = len(worksheet.col_values(2)) - 4
   print('a', count_itm)
   for i in range(count_itm):
      i = 0
      for itm_fin in itm_fins:
         itm_fin = driver.find_elements_by_class_name("wrapper")[1].find_elements_by_tag_name("a")[i]
         i += 1
      print('A', i)
      if i > 10:
         print('B', i)
         itm_fin1 = driver.find_elements_by_class_name("wrapper")[1].find_elements_by_tag_name("a")[2]
         itm_fin1.click()
         sleep(3)
         rpro()
         reproduction()
      else:
         break
         print('break')
except Exception as e:
   print(e)
   print("include error in total")
   # item_def.excep()
   # driver.quit()
print('last', dt.now())
# item_def.excep()
driver.quit() #全てのウインドウを閉じる