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
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.select import Select
from datetime import datetime as dt
import datetime
from selenium.webdriver.common.alert import Alert
import re

def itm1(driver): #ダイソン掃除機 タイヤ4個+スペーサー+トルクスドライバー3本セット
   driver.find_element_by_xpath("//input[@type='file']").send_keys("/Users/masa/Pictures/torx/torxset.JPG \n /Users/masa/Pictures/dyson/tire4.png \n /Users/masa/Pictures/dyson/spacers.png \n /Users/masa/Pictures/torx/torxdriver.jpeg \n /Users/masa/Pictures/torx/torxside.jpg \n /Users/masa/Pictures/torx/torxsize.JPG \n /Users/masa/Pictures/dyson/morterhead.jpeg \n /Users/masa/Pictures/dyson/tire_remove.JPG \n /Users/masa/Pictures/dyson/installation.JPG")
   global exp1a
   exp1a = "【商品紹介】\n画像の2～4枚目が商品になります。ダイソンモーターヘッドの裏側に付いているタイヤと工具です。"
def itm2(driver): #"ダイソン掃除機 タイヤ4個+スペーサー+シャフト4本+トルクスドライバー3本"
   driver.find_element_by_xpath("//input[@type='file']").send_keys("/Users/masa/Pictures/dyson/shaftset.JPG \n /Users/masa/Pictures/dyson/tire4.png \n /Users/masa/Pictures/dyson/spacers.png \n /Users/masa/Pictures/dyson/shaft.png \n /Users/masa/Pictures/dyson/morterhead.jpeg \n /Users/masa/Pictures/dyson/tire_remove.JPG \n /Users/masa/Pictures/dyson/installation.JPG")
   global exp1a
   exp1a = "【商品紹介】\n画像の2～4枚目が商品になります。ダイソンモーターヘッドの裏側に付いているタイヤと部品です。"
def itm3(driver): #"ダイソン掃除機 タイヤ4個+スペーサー+シャフト4本+トルクスドライバー3本"
   driver.find_element_by_xpath("//input[@type='file']").send_keys("/Users/masa/Pictures/dyson/4set.JPG \n /Users/masa/Pictures/dyson/tire4.png \n /Users/masa/Pictures/dyson/spacers.png \n /Users/masa/Pictures/torx/torxdriver.jpeg \n /Users/masa/Pictures/torx/torxside.jpg \n /Users/masa/Pictures/torx/torxsize.JPG \n /Users/masa/Pictures/dyson/shaft.png \n /Users/masa/Pictures/dyson/morterhead.jpeg \n /Users/masa/Pictures/dyson/tire_remove.JPG \n /Users/masa/Pictures/dyson/installation.JPG")
   global exp1a
   exp1a = "【商品紹介】\n画像の2～7枚目が商品になります。ダイソンモーターヘッドの裏側に付いているタイヤと部品と工具です。"
def itm4(driver): #"ダイソン掃除機 タイヤ4個+スペーサーセット"
   driver.find_element_by_xpath("//input[@type='file']").send_keys("/Users/masa/Pictures/dyson/tire4set.JPG \n /Users/masa/Pictures/dyson/tire4.png \n /Users/masa/Pictures/dyson/spacers.png \n /Users/masa/Pictures/dyson/morterhead.jpeg \n /Users/masa/Pictures/dyson/tire_remove.JPG \n /Users/masa/Pictures/dyson/installation.JPG")
   global exp1a
   exp1a = "【商品紹介】\n画像の2～3枚目が商品になります。ダイソンモーターヘッドの裏側に付いているタイヤとスペーサーです。"
def itm5(driver): #"ダイソン掃除機 タイヤ2個+スペーサーセット"
   driver.find_element_by_xpath("//input[@type='file']").send_keys("/Users/masa/Pictures/dyson/tireimage.png \n /Users/masa/Pictures/dyson/tire4.png \n /Users/masa/Pictures/dyson/morterhead.jpeg \n /Users/masa/Pictures/dyson/tire_remove.JPG \n /Users/masa/Pictures/dyson/installation.JPG")
   global exp1a
   exp1a = "【商品紹介】\n画像の2枚目が商品になります。ダイソンモーターヘッドの裏側に付いているタイヤです。"
def itm6(driver): #"トルクスドライバー3本セット（T10 & T8 & T6）"
   driver.find_element_by_xpath("//input[@type='file']").send_keys("/Users/masa/Pictures/dyson/tire2set.JPG \n /Users/masa/Pictures/dyson/tire2.jpeg \n /Users/masa/Pictures/dyson/spacers.png \n /Users/masa/Pictures/dyson/morterhead.jpeg \n /Users/masa/Pictures/dyson/tire_remove.JPG \n /Users/masa/Pictures/dyson/installation.JPG")
   global exp1a
   exp1a = "【商品紹介】\n画像の2～3枚目が商品になります。ダイソンモーターヘッドの裏側に付いているタイヤです。"
def itm7(driver): #"ダイソン掃除機 タイヤ4個（リピーター様専用）"
   driver.find_element_by_xpath("//input[@type='file']").send_keys("/Users/masa/Pictures/dyson/tireimage.png \n /Users/masa/Pictures/dyson/tire2.jpeg \n /Users/masa/Pictures/dyson/morterhead.jpeg \n /Users/masa/Pictures/dyson/tire_remove.JPG \n /Users/masa/Pictures/dyson/installation.JPG")
   global exp1a
   exp1a = "【商品紹介】\n画像の2枚目が商品になります。ダイソンモーターヘッドの裏側に付いているタイヤです。"
def itm8(driver): #"ダイソン掃除機 タイヤ2個（リピーター様専用）"
   driver.find_element_by_xpath("//input[@type='file']").send_keys("/Users/masa/Pictures/torx/torxdriver.jpeg \n /Users/masa/Pictures/torx/torxside.jpg \n /Users/masa/Pictures/torx/torxsize.JPG")
   global exp1a
   exp1a = "トルクスドライバー3本セット（T10 & T8 & T6）\
   \n ダイソン分解清掃などにお得な3本セットです。\n\n六角星形\nヘックスローブ\nヘクスローブ\ntorx\ndyson\nXBOX\nxbox"
def itm9(driver): #"ダイソン掃除機 タイヤ4個(大)+スペーサーセット"
   driver.find_element_by_xpath("//input[@type='file']").send_keys("/Users/masa/Pictures/dyson/tire4(large)set.JPG \n /Users/masa/Pictures/dyson/tire41mm.jpg \n /Users/masa/Pictures/dyson/spacer(縦大).png")
   global exp1a
   exp1a = "【商品紹介】\n画像の2～3枚目が商品になります。ダイソンモーターヘッドの裏側に付いているタイヤです。"
def itm10(driver): #"ダイソン掃除機 タイヤ2個(大)+スペーサーセット"
   driver.find_element_by_xpath("//input[@type='file']").send_keys("/Users/masa/Pictures/dyson/tire2(large)set.JPG \n /Users/masa/Pictures/dyson/tire21mm.jpg \n /Users/masa/Pictures/dyson/spacer(縦大).png")
   global exp1a
   exp1a = "【商品紹介】\n画像の2～3枚目が商品になります。ダイソンモーターヘッドの裏側に付いているタイヤです。"
def itm11(driver): #"ダイソン掃除機 タイヤ4個(大)+スペーサー+トルクスドライバー3本セット"
   driver.find_element_by_xpath("//input[@type='file']").send_keys("/Users/masa/Pictures/torx/torxset(large)2.JPG \n /Users/masa/Pictures/dyson/tire41mm.jpg \n /Users/masa/Pictures/dyson/spacer(縦大).png \n /Users/masa/Pictures/torx/torxdriver.jpeg \n /Users/masa/Pictures/torx/torxside.jpg \n /Users/masa/Pictures/torx/torxsize.JPG")
   global exp1a
   exp1a = "【商品紹介】\n画像の2～4枚目が商品になります。ダイソンモーターヘッドの裏側に付いているタイヤと工具のセットです。"
def itm12(driver): #"ダイソン掃除機 タイヤ4個(大)+トルクスドライバー3本セット"
   driver.find_element_by_xpath("//input[@type='file']").send_keys("/Users/masa/Pictures/torx/torxset(soft).jpeg \n /Users/masa/Pictures/dyson/tire41mm.jpg \n /Users/masa/Pictures/torx/torxdriver.jpeg \n /Users/masa/Pictures/torx/torxside.jpg \n /Users/masa/Pictures/torx/torxsize.JPG \n /Users/masa/Pictures/dyson/soft.jpeg \n /Users/masa/Pictures/dyson/tire_remove2.JPG")
   global exp1a
   exp1a = "【商品紹介】\n画像の1枚目が商品になります。ダイソン掃除機ソフトローラークリーンヘッドの裏側に付いているタイヤと工具のセットです。"
def itm13(driver): #"ダイソン掃除機 タイヤ4個(大)+シャフト4本セット"
   driver.find_element_by_xpath("//input[@type='file']").send_keys("/Users/masa/Pictures/dyson/shaft4set(soft).JPG \n /Users/masa/Pictures/dyson/tire41mm.jpg \n /Users/masa/Pictures/dyson/shaft4(soft).png \n /Users/masa/Pictures/dyson/soft.jpeg \n /Users/masa/Pictures/dyson/tire_remove2.JPG")
   global exp1a
   exp1a = "【商品紹介】\n画像の1枚目が商品になります。ダイソン掃除機ソフトローラークリーンヘッドの裏側に付いているタイヤと部品のセットです。"
def itm14(driver): #"ダイソン掃除機 タイヤ2個(大)+シャフト2本セット"
   driver.find_element_by_xpath("//input[@type='file']").send_keys("/Users/masa/Pictures/dyson/shaft2set(soft).JPG \n /Users/masa/Pictures/dyson/tire21mm.jpg \n /Users/masa/Pictures/dyson/shaft3.png \n /Users/masa/Pictures/dyson/soft.jpeg \n /Users/masa/Pictures/dyson/tire_remove2.JPG")
   global exp1a
   exp1a = "【商品紹介】\n画像の1枚目が商品になります。ダイソン掃除機ソフトローラークリーンヘッドの裏側に付いているタイヤと部品のセットです。"
def itm15(driver): #"ダイソン掃除機 タイヤ4個(大)+シャフト4本+トルクスドライバー3本セット"
   driver.find_element_by_xpath("//input[@type='file']").send_keys("/Users/masa/Pictures/dyson/tire4(large)3set.JPG \n /Users/masa/Pictures/dyson/tire41mm.jpg \n /Users/masa/Pictures/dyson/shaft4(soft).png \n /Users/masa/Pictures/dyson/soft.jpeg \n /Users/masa/Pictures/torx/torxdriver.jpeg \n /Users/masa/Pictures/torx/torxside.jpg \n /Users/masa/Pictures/torx/torxsize.JPG \n /Users/masa/Pictures/dyson/tire_remove2.JPG")
   global exp1a
   exp1a = "【商品紹介】\n画像の1枚目が商品になります。ダイソン掃除機ソフトローラークリーンヘッドの裏側に付いているタイヤと部品と工具のセットです。"
def sct(cateElement,indexNum):
   select = Select(cateElement)
   select.select_by_index(indexNum)
# ダイソン共通  ctd: 1-9.カテゴリー, 10-11.ブランド, 12-13.商品の説明
def com_to_dys(driver, exp1, mc, itm, itm_row): 
   cateElement = driver.find_elements_by_css_selector("select[name=categoryId]")[0]
   indexNum= 8
   sct(cateElement,indexNum)
   cateElement = driver.find_elements_by_css_selector("select[name=categoryId]")[1]
   indexNum= 9
   sct(cateElement,indexNum)
   cateElement = driver.find_elements_by_css_selector("select[name=categoryId]")[2]
   indexNum= 7
   sct(cateElement,indexNum)
   inputElement2 = driver.find_elements_by_class_name("style_inputarea__1mtsA")[1].find_element_by_css_selector("input")
   inputElement2.send_keys("ダイソン")
   exp1(driver, itm, itm_row)
   mc(driver)
def exp1(driver, itm, itm_row):
   inputElement1 = driver.find_elements_by_class_name("style_inputarea__1mtsA")[0].find_element_by_css_selector("input")
   inputElement1.send_keys(itm)
   inputElement3 = driver.find_elements_by_class_name("style_inputarea__1mtsA")[2].find_element_by_css_selector("input")
   inputElement3.send_keys(worksheet.cell(itm_row, 8).value)
   # inputElement3.send_keys(price)
   print(worksheet.cell(itm_row, 8).value)
   inputElement4 = driver.find_elements_by_class_name("style_inputarea__1mtsA")[0].find_element_by_css_selector("input")
   inputElement1 = driver.find_elements_by_class_name("style_body__1OP1S")[2].find_element_by_css_selector("textarea")
   print(174, inputElement4.get_attribute("value"))
   if ('ダイソン' in inputElement4.get_attribute("value")) and ('大' in inputElement4.get_attribute("value")) and ('スペーサー' in inputElement4.get_attribute("value")):
      print(176)
      inputElement1.send_keys(exp1a + exp3a + exp0a)
   elif ('ダイソン' in inputElement4.get_attribute("value")) and ('大' in inputElement4.get_attribute("value")):
      print(179)
      inputElement1.send_keys(exp1a + exp4a + exp0a)
   elif 'ダイソン' in inputElement4.get_attribute("value"):
      print(182)
      inputElement1.send_keys(exp1a + exp2a + exp0a)
   else:
      print(185)
      inputElement1.send_keys(exp1a)
exp0a = "\
   \nタイヤが劣化すると約１年ほどでフローリングや畳が傷つくことがありますので、定期的に交換する必要があります。\
   \n\n【交換方法】\n交換方法などブログで詳しく紹介していますので、必ず購入前にご確認ください。\nブラウザでURLを入力すると検索できます。(URL：fuku-channnel.com)\nブログが検索できたら、サイト内検索で「ダイソン」と入力下さい。\
   \n\n【保証内容】\n対象機種以外で交換をご検討されている場合、下記2点をご対応いただくことで保証が適応されることがあります。(対象機種は保証が適用されます)\n\n1. コメントにてご利用の機種をご連絡ください。\n2. ヘッドの裏の写真(タイヤ部分がわかる写真)を出品下さい。(価格は100,000としてください。)\nこちらで取り付けできると判断した場合のみ、保証が適応されます。保証が適応された場合、万が一取り付けできなかったら返金いたしますので、その旨ご連絡ください。\
   \n\n【交換頻度】\n１年以内の交換をおすすめしています。\
   \n\n【現状のメーカー修理】\nメーカーで修理依頼すると、購入から2年以上経過すると最低5000円の工賃が発生します。\
   \n\n【最後に】\nタイヤは簡単に交換できるので、DIYで行うことをおすすめしています。(機種によって右後輪の取り外しが困難な場合があります)\nメーカーから部品だけの購入が不可と言われたため、自分で作成したものを出品しています。\n最近、ご購入者様からお褒めの言葉を多く頂いております。モチベーションの向上に繋がりますので、評価の方にもコメントいただけたら幸いです。"
exp2a = "\
   \n\n【対象機種】\n実績機種は、DC35, DC48, DC59, DC62, V6です。\nその他の機種でも適合する可能性があります。(その他の機種でタイヤ交換をご検討されている方は次の保証内容をご確認ください。)\
   \n\n【タイヤのサイズ】\n外径: 10 mm\n内径: 3 mm\n幅: 4.5 mm\nシャフトにスペーサーを装着することで取り付けできます。"
exp3a = "\
   \n\n【対象機種】\n実績機種は、DC48, DC63, CY25ダイレクトドライブです。\nその他の機種でも適合する可能性があります。(その他の機種でタイヤ交換をご検討されている方は次の保証内容をご確認ください。)\
   \n\n【タイヤのサイズ】\n外径: 10 mm\n内径: 3 mm\n幅: 10 mm \nシャフトにスペーサーを装着することで取り付けできます"
exp4a = "\
   \n\n【対象機種】\n実績機種は、ソフトローラークリーンヘッド搭載機種です。\nその他の機種でも適合する可能性があります。(その他の機種でタイヤ交換をご検討されている方は次の保証内容をご確認ください。)\
   \n\n【タイヤのサイズ】\n外径: 10 mm\n内径: 3 mm\n幅: 10 mm \n取り外し方などは上記ブログをご覧下さい。"
# メルカリ共通  mc: 1-3.商品の状態, 4-6.配送料の負担, 7-9.配送の方法, 10-12.配送元の地域, 13-15.発送までの日数
def mc(driver):
   cateElement = driver.find_element_by_css_selector("select[name=itemCondition")
   indexNum= 1
   sct(cateElement,indexNum)
   cateElement = driver.find_element_by_css_selector("select[name=shippingPayer")
   indexNum= 1
   sct(cateElement,indexNum)
   cateElement = driver.find_element_by_css_selector("select[name=shippingMethod")
   indexNum= 6
   sct(cateElement,indexNum)
   cateElement = driver.find_element_by_css_selector("select[name=shippingFromArea")
   indexNum= 12
   sct(cateElement,indexNum)
   cateElement = driver.find_element_by_css_selector("select[name=shippingDuration")
   indexNum= 1
   sct(cateElement,indexNum)
def exh(itm_row, exhibit_col, driver, time):
   sl = driver.find_element_by_css_selector("button[type=submit")
   sleep(10)
   sl.click()
   sleep(1)
   driver.get("https://www.mercari.com/jp/sell/")
   sleep(1)
   worksheet.update_cell(itm_row, exhibit_col, str(time))
def abc(driver, exhibit_col):
    driver.get("https://www.mercari.com/jp/mypage/listings/listing/") # メルカリにアクセス
    sleep(1)
    time = dt.now()
    a = len(worksheet.col_values(2))
    exhibit_col_com = exhibit_col
    print(44, exhibit_col)
    worksheet.update_cell(2, 2, str(time))
    dictm = {}
    dictcomm = {}
    urlitems = driver.find_elements_by_class_name("mypage-item-link")
    for urlitem in urlitems:
        sleep(1)
        url = urlitem.get_attribute("href")
        item = urlitem.find_element_by_class_name("mypage-item-text")
        com = urlitem.find_elements_by_class_name("listing-item-count")[1]
        dictm[item.text]=url
        dictcomm[item.text]=com.text
    for i in range(5, a + 1):
        sleep(1)
        if dt.strptime(worksheet.cell(i, exhibit_col).value, '%Y/%m/%d %H:%M:%S') < dt.strptime(worksheet.cell(3, 2).value, '%Y/%m/%d %H:%M:%S'):
            del_item = worksheet.cell(i, 2).value
            sleep(1)
            if del_item in dictm and dictcomm[del_item] == "0":
                driver.get(dictm[del_item])
                sleep(5)
                del_button1 = driver.find_element_by_css_selector("button[data-modal='delete-item']")
                del_button1.click()
                sleep(5)
                del_button3 = driver.find_elements_by_class_name("modal-btn.modal-btn-submit")[1]
                del_button3.click()
    itmlists = driver.find_elements_by_class_name("mypage-item-text") # 出品中の商品リストを作成
    lists = []
    for itmlist in itmlists:
        sleep(1)
        lists.append(itmlist.text)
    driver.get("https://www.mercari.com/jp/sell/")
    sleep(3)
   #  1.画像をアップロード & 商品名 & 販売価格, 2.カテゴリ & ブランド & 商品説明, 商品の状態 & 配送料の負担 & 配送の方法 & 発送元の地域 & 発送までの日数, 3.出品する
    itm_row = 5
    itm = worksheet.cell(itm_row, 2).value #ダイソン掃除機 タイヤ4個+スペーサー+トルクスドライバー3本セット
    if not itm in lists:
        itm1(driver)
        com_to_dys(driver, exp1, mc, itm, itm_row)
        exh(itm_row, exhibit_col, driver, time)
      #   worksheet.update_cell(itm_row, exhibit_col, str(time))
    itm_row = 6
    itm = worksheet.cell(itm_row, 2).value #"ダイソン掃除機 タイヤ4個+スペーサー+シャフト4本セット"
    if not itm in lists:
        itm2(driver)
        com_to_dys(driver, exp1, mc, itm, itm_row)
        exh(itm_row, exhibit_col, driver, time)
      #   worksheet.update_cell(itm_row, exhibit_col, str(time))
    itm_row = 7
    itm = worksheet.cell(itm_row, 2).value #"ダイソン掃除機 タイヤ4個+スペーサー+シャフト4本+トルクスドライバー3本"
    if not itm in lists:
        itm3(driver)
        com_to_dys(driver, exp1, mc, itm, itm_row)
        exh(itm_row, exhibit_col, driver, time)
      #   worksheet.update_cell(itm_row, exhibit_col, str(time))
    print('C-1', dt.now())
    itm_row = 8
    itm = worksheet.cell(itm_row, 2).value #"ダイソン掃除機 タイヤ4個+スペーサーセット"
    if not itm in lists:
        itm4(driver)
        com_to_dys(driver, exp1, mc, itm, itm_row)
        exh(itm_row, exhibit_col, driver, time)
      #   worksheet.update_cell(itm_row, exhibit_col, str(time))
    itm_row = 9
    itm = worksheet.cell(itm_row, 2).value #"ダイソン掃除機 タイヤ2個+スペーサーセット"
    if not itm in lists:
        itm6(driver)
        com_to_dys(driver, exp1, mc, itm, itm_row)
        exh(itm_row, exhibit_col, driver, time)
      #   worksheet.update_cell(itm_row, exhibit_col, str(time))
    itm_row = 10
    itm = worksheet.cell(itm_row, 2).value #"トルクスドライバー3本セット（T10 & T8 & T6）"
    if not itm in lists:
        itm8(driver)
        com_to_dys(driver, exp1, mc, itm, itm_row)
        exh(itm_row, exhibit_col, driver, time)
      #   worksheet.update_cell(itm_row, exhibit_col, str(time))
    itm_row = 11
    itm = worksheet.cell(itm_row, 2).value #"ダイソン掃除機 タイヤ4個（リピーター様専用）"
    if not itm in lists:
        itm5(driver)
        com_to_dys(driver, exp1, mc, itm, itm_row)
        exh(itm_row, exhibit_col, driver, time)
      #   worksheet.update_cell(itm_row, exhibit_col, str(time))
    itm_row = 12
    itm = worksheet.cell(itm_row, 2).value #"ダイソン掃除機 タイヤ2個（リピーター様専用）"
    if not itm in lists:
        itm7(driver)
        com_to_dys(driver, exp1, mc, itm, itm_row)
        exh(itm_row, exhibit_col, driver, time)
      #   worksheet.update_cell(itm_row, exhibit_col, str(time))
    itm_row = 13
    itm = worksheet.cell(itm_row, 2).value #"ダイソン掃除機 タイヤ4個(大)+スペーサーセット"
    if not itm in lists:
        itm9(driver)
        com_to_dys(driver, exp1, mc, itm, itm_row)
        exh(itm_row, exhibit_col, driver, time)
      #   worksheet.update_cell(itm_row, exhibit_col, str(time))
    itm_row = 14
    itm = worksheet.cell(itm_row, 2).value #"ダイソン掃除機 タイヤ2個(大)+スペーサーセット"
    if not itm in lists:
        itm10(driver)
        com_to_dys(driver, exp1, mc, itm, itm_row)
        exh(itm_row, exhibit_col, driver, time)
      #   worksheet.update_cell(itm_row, exhibit_col, str(time))
    itm_row = 15
    itm = worksheet.cell(itm_row, 2).value #"ダイソン掃除機 タイヤ4個(大)+スペーサー+トルクスドライバー3本セット"
    if not itm in lists:
        itm11(driver)
        com_to_dys(driver, exp1, mc, itm, itm_row)
        exh(itm_row, exhibit_col, driver, time)
      #   worksheet.update_cell(itm_row, exhibit_col, str(time))
    print('C-2', dt.now())
    itm_row = 16
    itm = worksheet.cell(itm_row, 2).value #"ダイソン掃除機 タイヤ4個(大)+トルクスドライバー3本セット"
    if not itm in lists:
        itm12(driver)
        com_to_dys(driver, exp1, mc, itm, itm_row)
        exh(itm_row, exhibit_col, driver, time)
      #   worksheet.update_cell(itm_row, exhibit_col, str(time))
    itm_row = 17
    itm = worksheet.cell(itm_row, 2).value #"ダイソン掃除機 タイヤ4個(大)+シャフト4本セット"
    if not itm in lists:
        itm13(driver)
        com_to_dys(driver, exp1, mc, itm, itm_row)
        exh(itm_row, exhibit_col, driver, time)
      #   worksheet.update_cell(itm_row, exhibit_col, str(time))
    itm_row = 18
    itm = worksheet.cell(itm_row, 2).value #"ダイソン掃除機 タイヤ2個(大)+シャフト2本セット"
    if not itm in lists:
        itm14(driver)
        com_to_dys(driver, exp1, mc, itm, itm_row)
        exh(ditm_row, exhibit_col, driver, time)
      #   worksheet.update_cell(itm_row, exhibit_col, str(time))
    itm_row = 19
    itm = worksheet.cell(itm_row, 2).value #"ダイソン掃除機 タイヤ4個(大)+シャフト4本+トルクスドライバー3本セット"
    if not itm in lists:
        itm15(driver)
        com_to_dys(driver, exp1, mc, itm, itm_row)
        exh(itm_row, exhibit_col, driver, time)
      #   worksheet.update_cell(itm_row, exhibit_col, str(time))
    return exhibit_col, itm, time