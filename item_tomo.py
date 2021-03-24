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
   '--user-data-dir={chrom_dir_path}'.format(chrom_dir_path = '/Users/masa/Library/Application Support/Google/Chrome/Profile 2'))
options.add_argument('--headless')
options.add_argument('--no-sandbox')
options.add_argument('--disable-gpu')
driver = webdriver.Chrome(ChromeDriverManager().install(), options=options)
# driver.minimize_window()
# driver.get("https://www.mercari.com/jp/mypage/listings/listing/") # メルカリにアクセス
# sleep(1)
exhibit_col = 5
item_def.abc(driver, exhibit_col)
sleep(1)
print('xyz')
driver.quit( ) #全てのウインドウを閉じる

# # ラクマページにアクセス
# def rexp1(rinputElement1):
#    rinputElement3 = driver.find_elements_by_class_name("form-group")[2].find_element_by_css_selector("textarea")
#    rinputElement3.send_keys(rexp1a + exp0a)
#    print('rexp1')
# # #ダイソン共通  rakumacommon_to_cut: 
# def rcom_to_cut():
#    driver.find_elements_by_xpath("//input[@type='file']")[0].send_keys("/Users/masa/Pictures/tube_cutter/spares.JPG")
#    driver.find_elements_by_xpath("//input[@type='file']")[1].send_keys("/Users/masa/Pictures/tube_cutter/spare.jpeg")
#    driver.find_elements_by_xpath("//input[@type='file']")[2].send_keys("/Users/masa/Pictures/tube_cutter/tube_cutter2.jpeg")
#    # driver.find_elements_by_xpath("//input[@type='file']")[3].send_keys("/Users/masa/Pictures/tube_cutter/")
#    rinputElement1 = driver.find_elements_by_class_name("form-group")[1].find_element_by_css_selector("input")
#    rinputElement1.send_keys(itm_name)
#    rinputElement3 = driver.find_elements_by_class_name("form-group")[2].find_element_by_css_selector("textarea")
#    rinputElement3.send_keys(itm_name + "です。画像の1枚目が商品になります。サイズは写真１枚目の通りです。画像３枚目の取り付け用ケースは付きません。\
#       \n\nTC16SC\
#       \n千代田通商\
#       \nチヨダ\
#       \nchiyoda\
#       \nCHIYODA")
#    rcateElement1 = driver.find_element_by_id("category_name")
#    rcateElement1.click()
#    rcateElement2 = driver.find_elements_by_class_name("panel.list-group")[7]
#    rcateElement2.click()
#    sleep(2)
#    rcateElement3 = rcateElement2.find_elements_by_class_name("list-group-item.small.branch")[11]
#    # rcateElement3 = driver.find_elements_by_class_name("panel.list-group")[7].find_elements_by_class_name("list-group-item.small.branch")[11]
#    rcateElement3.click()
#    print('227', rcateElement3.text)
#    sleep(2)
#    rcateElement4 = driver.find_elements_by_class_name("sublinks.collapse.in")[1].find_elements_by_class_name("list-group-item.small.leaf")[2]
#    print(230, rcateElement4.text)
#    # rcateElement4 = rcateElement2.find_elements_by_class_name("sublinks.collapse.in")[1].find_elements_by_class_name("list-group-item.small.leaf")[2]
#    # rcateElement4 = driver.find_elements_by_class_name("panel.list-group")[7].find_elements_by_class_name("list-group-item.small.branch")[11].find_elements_by_class_name("list-group-item.small.leaf")[1]
#    rcateElement4.click()
#    sleep(2)
#    rcateElement5 = driver.find_element_by_id("delivery_method")
#    # print(349, rcateElement5)
#    rcateElement5.click()
#    sleep(2)
#    rcateElement6 = driver.find_elements_by_class_name("modal-body")[3].find_elements_by_class_name("list-group-item")[12]
#    # print(353, rcateElement6.text)
#    sleep(2)
#    rcateElement6.click()
#    sleep(2)
#    rcateElement7 = driver.find_elements_by_class_name("col-sm-9.col-xs-8")[8]
#    rcateElement8 = rcateElement7.find_element_by_class_name("form-control")
#    # print(242, rcateElement8.text)
#    rcateElement9 = Select(rcateElement8).options
#    print(243, rcateElement9[5].text)
#    rcateElement10 = rcateElement9[11]
#    rcateElement10.click()
#    rinputElement8 = driver.find_elements_by_class_name("form-group")[11].find_element_by_css_selector("input")
#    rinputElement8.send_keys(cost)
#    print('rcom_to_dys')
# def rclick():
#    rcateElement9 = driver.find_element_by_class_name("col-md-offset-4.col-md-5.col-xs-12")
#    rcateElement9.click()
#    sleep(2)
#    rcateElement10 = driver.find_element_by_class_name("col-md-8.col-xs-8")
#    sleep(2)
#    rcateElement10.click()
#    sleep(8)
#    driver.get("https://fril.jp/item/new")
#    sleep(2)
#    print('rclick')

# driver.get("https://fril.jp/sell")
# for i in range(5,7): #商品追加したら数字を増やす
#    print('i', i, '387')
#    if dt.strptime(worksheet.cell(i, 19).value, '%Y/%m/%d %H:%M:%S') < dt.strptime(worksheet.cell(i, 20).value, '%Y/%m/%d %H:%M:%S'): #出品してから１日以上経過している場合、trueになります。
#       del_item = worksheet.cell(i, 18).value #該当のサンプル名を代入
#       urlitems = driver.find_elements_by_class_name("information-pane")[0].find_elements_by_class_name("media") #親要素を取得
#       # print('E-1-1', dt.now())
#       for urlitem in urlitems:
#          itemr = urlitem.find_element_by_class_name("media-heading") #商品名を取得
#          comr = urlitem.find_elements_by_class_name("item-action-count")[1] #コメントの数を取得
#          # print('E-1-2', dt.now(),'itemr', itemr)
#          if del_item == itemr.text and comr.text == "0":
#             print('344')
#             urlr = urlitem.find_elements_by_class_name("btn.btn-default")[1] #削除ボタンを取得
#             # print('E-2', dt.now())
#             sleep(5)
#             urlr.click()
#             # print('E-3', dt.now())
#             sleep(10)
#             Alert(driver).accept() #削除しますか？のはいをクリックして削除
#             # print('E-4', dt.now())
#             sleep(5)
#             break

# print('J-1', dt.now())
# rlists=[]
# sleep(3)
# ritmlists = driver.find_elements_by_class_name("media-heading")
# for ritmlist in ritmlists:
#    rlists.append(ritmlist.text)
# driver.get("https://fril.jp/item/new")
# sleep(3)
# itm_name = "チューブカッター用替刃 5枚セット"
# if not itm_name in rlists:
#    cost = 500
#    rcom_to_cut()
#    rclick()
#    worksheet.update_cell(5, 19, str(time))
# itm_name = "チューブカッター用替刃 10枚セット"
# if not itm_name in rlists:
#    cost = 950
#    rcom_to_cut()
#    rclick()
#    worksheet.update_cell(6, 19, str(time))
# driver.quit() #全てのウインドウを閉じる