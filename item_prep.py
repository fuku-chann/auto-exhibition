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

from selenium import webdriver
from selenium.webdriver import ActionChains
from time import sleep
#import chromedriver_binary
import webbrowser
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.select import Select
from datetime import datetime as dt
from selenium.webdriver.common.alert import Alert

print('A-1', dt.now())

# 入力（共通部分）
def itmpage():
   #sleep(1)
   driver.get("https://www.mercari.com/jp/sell/")
   sleep(3)
# itmx: 1.画像をアップロード, 2-3.商品名, 4-5.販売価格
def itm1():
   driver.find_element_by_xpath("//input[@type='file']").send_keys("/Users/masa/Pictures/torx/torxset.JPG \n /Users/masa/Pictures/dyson/tire5.jpg \n /Users/masa/Pictures/dyson/tape.jpeg \n /Users/masa/Pictures/torx/torxdriver.jpeg \n /Users/masa/Pictures/dyson/morterhead.jpeg \n /Users/masa/Pictures/dyson/morterheadback.jpg")
   inputElement1 = driver.find_elements_by_class_name("style_inputarea__1mtsA")[0].find_element_by_css_selector("input")
   inputElement1.send_keys("ダイソン掃除機 タイヤ5個+テフロンテープ+トルクスドライバー3本セット")
   inputElement3 = driver.find_elements_by_class_name("style_inputarea__1mtsA")[2].find_element_by_css_selector("input")
   inputElement3.send_keys("680")
   global exp1a
   exp1a = "【商品紹介】\n画像の2～4枚目が商品になります。ダイソンモーターヘッドの裏側に付いているタイヤと工具です。"
def itm2():
   driver.find_element_by_xpath("//input[@type='file']").send_keys("/Users/masa/Pictures/dyson/shaftset.JPG \n /Users/masa/Pictures/dyson/tire5.jpg \n /Users/masa/Pictures/dyson/tape.jpeg \n /Users/masa/Pictures/dyson/shaft.jpeg \n /Users/masa/Pictures/dyson/morterhead.jpeg \n /Users/masa/Pictures/dyson/morterheadback.jpg")
   inputElement1 = driver.find_elements_by_class_name("style_inputarea__1mtsA")[0].find_element_by_css_selector("input")
   inputElement1.send_keys("ダイソン掃除機 タイヤ5個+テフロンテープ+シャフト4本セット")
   inputElement3 = driver.find_elements_by_class_name("style_inputarea__1mtsA")[2].find_element_by_css_selector("input")
   inputElement3.send_keys("680")
   global exp1a
   exp1a = "【商品紹介】\n画像の2～4枚目が商品になります。ダイソンモーターヘッドの裏側に付いているタイヤと部品です。"
def itm3():
   driver.find_element_by_xpath("//input[@type='file']").send_keys("/Users/masa/Pictures/dyson/4set.JPG \n /Users/masa/Pictures/dyson/tire5.jpg \n /Users/masa/Pictures/dyson/tape.jpeg \n /Users/masa/Pictures/torx/torxdriver.jpeg \n /Users/masa/Pictures/dyson/shaft.jpeg \n /Users/masa/Pictures/dyson/morterhead.jpeg \n /Users/masa/Pictures/dyson/morterheadback.jpg")
   inputElement1 = driver.find_elements_by_class_name("style_inputarea__1mtsA")[0].find_element_by_css_selector("input")
   inputElement1.send_keys("ダイソン掃除機 タイヤ5個+テフロンテープ+シャフト4本+トルクスドライバー3本")
   inputElement3 = driver.find_elements_by_class_name("style_inputarea__1mtsA")[2].find_element_by_css_selector("input")
   inputElement3.send_keys("880")
   global exp1a
   exp1a = "【商品紹介】\n画像の2～5枚目が商品になります。ダイソンモーターヘッドの裏側に付いているタイヤと部品と工具です。"
def itm4():
   driver.find_element_by_xpath("//input[@type='file']").send_keys("/Users/masa/Pictures/dyson/tire5set.JPG \n /Users/masa/Pictures/dyson/tire5.jpg \n /Users/masa/Pictures/dyson/tape.jpeg \n /Users/masa/Pictures/dyson/morterhead.jpeg \n /Users/masa/Pictures/dyson/morterheadback.jpg")
   inputElement1 = driver.find_elements_by_class_name("style_inputarea__1mtsA")[0].find_element_by_css_selector("input")
   inputElement1.send_keys("ダイソン掃除機 タイヤ5個+テフロンテープセット")
   inputElement3 = driver.find_elements_by_class_name("style_inputarea__1mtsA")[2].find_element_by_css_selector("input")
   inputElement3.send_keys("600")
   global exp1a
   exp1a = "【商品紹介】\n画像の2～3枚目が商品になります。ダイソンモーターヘッドの裏側に付いているタイヤです。"
def itm5():
   driver.find_element_by_xpath("//input[@type='file']").send_keys("/Users/masa/Pictures/dyson/tireimage.png \n /Users/masa/Pictures/dyson/tire5.jpg \n /Users/masa/Pictures/dyson/morterhead.jpeg \n /Users/masa/Pictures/dyson/morterheadback.jpg")
   inputElement1 = driver.find_elements_by_class_name("style_inputarea__1mtsA")[0].find_element_by_css_selector("input")
   inputElement1.send_keys("ダイソン掃除機 タイヤ5個")
   inputElement3 = driver.find_elements_by_class_name("style_inputarea__1mtsA")[2].find_element_by_css_selector("input")
   inputElement3.send_keys("500")
   global exp1a
   exp1a = "【商品紹介】\n画像の2枚目が商品になります。ダイソンモーターヘッドの裏側に付いているタイヤです。"
def itm6():
   driver.find_element_by_xpath("//input[@type='file']").send_keys("/Users/masa/Pictures/dyson/tire2set.JPG \n /Users/masa/Pictures/dyson/tire2.jpeg \n /Users/masa/Pictures/dyson/tape.jpeg \n /Users/masa/Pictures/dyson/morterhead.jpeg \n /Users/masa/Pictures/dyson/morterheadback.jpg")
   inputElement1 = driver.find_elements_by_class_name("style_inputarea__1mtsA")[0].find_element_by_css_selector("input")
   inputElement1.send_keys("ダイソン掃除機 タイヤ2個+テフロンテープセット")
   inputElement3 = driver.find_elements_by_class_name("style_inputarea__1mtsA")[2].find_element_by_css_selector("input")
   inputElement3.send_keys("400")
   global exp1a
   exp1a = "【商品紹介】\n画像の2～3枚目が商品になります。ダイソンモーターヘッドの裏側に付いているタイヤです。"
def itm7():
   driver.find_element_by_xpath("//input[@type='file']").send_keys("/Users/masa/Pictures/dyson/tireimage.png \n /Users/masa/Pictures/dyson/tire2.jpeg \n /Users/masa/Pictures/dyson/morterhead.jpeg \n /Users/masa/Pictures/dyson/morterheadback.jpg")
   inputElement1 = driver.find_elements_by_class_name("style_inputarea__1mtsA")[0].find_element_by_css_selector("input")
   inputElement1.send_keys("ダイソン掃除機 タイヤ2個")
   inputElement3 = driver.find_elements_by_class_name("style_inputarea__1mtsA")[2].find_element_by_css_selector("input")
   inputElement3.send_keys("300")
   global exp1a
   exp1a = "【商品紹介】\n画像の2枚目が商品になります。ダイソンモーターヘッドの裏側に付いているタイヤです。"
def itm8():
   driver.find_element_by_xpath("//input[@type='file']").send_keys("/Users/masa/Pictures/torx/torxdriver.jpeg")
   inputElement1 = driver.find_elements_by_class_name("style_inputarea__1mtsA")[0].find_element_by_css_selector("input")
   inputElement1.send_keys("トルクスドライバー3本セット（T10 & T8 & T6）")
   inputElement3 = driver.find_elements_by_class_name("style_inputarea__1mtsA")[2].find_element_by_css_selector("input")
   inputElement3.send_keys("430")
   global exp1a
   exp1a = "トルクスドライバー3本セット（T10 & T8 & T6）\
   #\n\n ダイソン分解清掃などにお得な3本セットです。"
def sct(cateElement,indexNum):
   select = Select(cateElement)
   select.select_by_index(indexNum)
# ダイソン共通  ctd: 1-9.カテゴリー, 10-11.ブランド, 12-13.商品の説明
def com_to_dys(exp1, mc): 
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
   exp1()
   mc()
def exp1():
   inputElement4 = driver.find_elements_by_class_name("style_inputarea__1mtsA")[0].find_element_by_css_selector("input")
   inputElement1 = driver.find_elements_by_class_name("style_body__1OP1S")[2].find_element_by_css_selector("textarea")
   if "ダイソン" in inputElement4.get_attribute("value"):
      inputElement1.send_keys(exp1a + exp0a)
   else:
      inputElement1.send_keys(exp1a)
exp0a = "\
   \nタイヤが劣化すると約１年ほどでフローリングや畳が傷つくことがあります。そのため、定期的に交換する必要があります。\
   \n\n【交換方法】\n交換方法などブログで詳しく紹介していますので、必ず購入前にご確認ください。\nブラウザでURLを入力すると検索できます。(URL：fuku-channnel.com)\nブログが検索できたら、サイト内検索で「ダイソン」と入力下さい。\
   \n\n【対象機種】\n実績機種は、DC35, DC48, DC59, DC62, V6です。\nその他の機種でも適合する可能性があります。(その他の機種でタイヤ交換をご検討されている方は次の保証内容をご確認ください。)\
   \n\n【保証内容】\n対象機種以外で交換をご検討されている場合、下記2点をご対応いただくことで保証が適応されることがあります。(対象機種は保証が適用されます)\n1. コメントにてご利用の機種をご連絡ください。\n2. ヘッドの裏の写真(タイヤ部分がわかる写真)を出品下さい。(価格は100,000としてください。)\nこちらで取り付けできると判断した場合のみ、保証が適応されます。保証が適応された場合、万が一取り付けできなかったら返金いたしますので、その旨ご連絡ください。\
   \n\n【交換頻度】\n１年以内の交換をおすすめしています。\n【タイヤのサイズ】\n外径: 10 mm\n内径: 3 mm\n幅: 4.5 mm (予備タイヤは3.5 mm)\nシャフトにシールを巻くことを前提にしています。詳しくは上記ブログ（タイヤの取り付け方）をご覧下さい。\
   \n\n【現状のメーカー修理】\nメーカーで修理依頼すると、購入から2年以上経過すると最低5000円の工賃が発生します。\
   \n\n【最後に】\nタイヤは簡単に交換できるので、DIYで行うことをおすすめしています。(機種によって右後輪の取り外しが困難な場合があります)\nメーカーから部品だけの購入が不可と言われたため、自分で作成したものを出品しています。\n最近、ご購入者様からお褒めの言葉を多く頂いております。モチベーションの向上に繋がりますので、評価の方にもコメントいただけたら幸いです。"
# メルカリ共通  mc: 1-3.商品の状態, 4-6.配送料の負担, 7-9.配送の方法, 10-12.配送元の地域, 13-15.発送までの日数
def mc():
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
def exh():
   sl = driver.find_element_by_css_selector("button[type=submit")
   sleep(10)
   sl.click()
   sleep(1)
   driver.get("https://www.mercari.com/jp/sell/")
   sleep(1)

# Open web (account 1)
options = webdriver.ChromeOptions()
options.add_argument(
   '--user-data-dir={chrom_dir_path}'.format(chrom_dir_path = '/Users/masa/Library/Application Support/Google/Chrome/Profile 3'))
try:
   driver = webdriver.Chrome(options=options, executable_path='/Users/masa/bin/chromedriver')
   print('B-1', dt.now())
   driver.get("https://www.mercari.com/jp/mypage/listings/listing/") # メルカリにアクセス
   sleep(1)
   time = dt.now()
   worksheet.update('D5:D12', [[str(time)], [str(time)], [str(time)], [str(time)], [str(time)], [str(time)], [str(time)], [str(time)]])
   dictm = {}
   dictcomm = {}
   urlitems = driver.find_elements_by_class_name("mypage-item-link")
   for urlitem in urlitems:
      url = urlitem.get_attribute("href")
      item = urlitem.find_element_by_class_name("mypage-item-text")
      com = urlitem.find_elements_by_class_name("listing-item-count")[1]
      dictm[item.text]=url
      dictcomm[item.text]=com.text
   for i in range(5,13):
      if dt.strptime(worksheet.cell(i,9).value, '%Y/%m/%d %H:%M:%S') < dt.strptime(worksheet.cell(i,10).value, '%Y/%m/%d %H:%M:%S'):
         del_item = worksheet.cell(i, 8).value
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
      lists.append(itmlist.text)
   itmpage()
   # 1.画像をアップロード & 商品名 & 販売価格, 2.カテゴリ & ブランド & 商品説明, 商品の状態 & 配送料の負担 & 配送の方法 & 発送元の地域 & 発送までの日数, 3.出品する
   if not "ダイソン掃除機 タイヤ5個+テフロンテープ+トルクスドライバー3本セット" in lists:
      itm1()
      com_to_dys(exp1, mc)
      exh()
      worksheet.update_cell(5, 9, str(time))
   if not "ダイソン掃除機 タイヤ5個+テフロンテープ+シャフト4本セット" in lists:
      itm2()
      com_to_dys(exp1, mc)
      exh()
      worksheet.update_cell(6, 9, str(time))
   if not "ダイソン掃除機 タイヤ5個+テフロンテープ+シャフト4本+トルクスドライバー3本" in lists:
      itm3()
      com_to_dys(exp1, mc)
      exh()
      worksheet.update_cell(7, 9, str(time))
   print('C-1', dt.now())
   if not "ダイソン掃除機 タイヤ5個+テフロンテープセット" in lists:
      itm4()
      com_to_dys(exp1, mc)
      exh()
      worksheet.update_cell(8, 9, str(time))
   if not "ダイソン掃除機 タイヤ5個" in lists:
      itm5()
      com_to_dys(exp1, mc)
      exh()
      worksheet.update_cell(11, 9, str(time))
   if not "ダイソン掃除機 タイヤ2個+テフロンテープセット" in lists:
      itm6()
      com_to_dys(exp1, mc)
      exh()
      worksheet.update_cell(9, 9, str(time))
   if not "ダイソン掃除機 タイヤ2個" in lists:
      itm7()
      com_to_dys(exp1, mc)
      exh()
      worksheet.update_cell(12, 9, str(time))
   if not "トルクスドライバー3本セット（T10 & T8 & T6）" in lists:
      itm8()
      com_to_dys(exp1, mc)
      exh()
      worksheet.update_cell(10, 9, str(time))
   print('D-1', dt.now())
except Exception as e:
   print(e)
   driver.quit()

# ラクマページにアクセス
def rexp1():
   rinputElement3 = driver.find_elements_by_class_name("form-group")[2].find_element_by_css_selector("textarea")
   rinputElement3.send_keys(rexp1a + exp0a)
   print('rexp1')
#ダイソン共通  rakumacommon_to_dyson: 
def rcom_to_dys():
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
   print('rcom_to_dys')
def rclick():
   rcateElement9 = driver.find_element_by_class_name("col-md-offset-4.col-md-5.col-xs-12")
   rcateElement9.click()
   sleep(2)
   rcateElement10 = driver.find_element_by_class_name("col-md-8.col-xs-8")
   sleep(2)
   rcateElement10.click()
   sleep(8)
   driver.get("https://fril.jp/item/new")
   sleep(2)
   print('rclick')
try:
   driver.get("https://fril.jp/sell")
   sleep(5)
   print('E-1', dt.now())
   for i in range(5,10):
      print('i', i)
      if dt.strptime(worksheet.cell(i,3).value, '%Y/%m/%d %H:%M:%S') < dt.strptime(worksheet.cell(i,5).value, '%Y/%m/%d %H:%M:%S'): #出品してから１日以上経過している場合、trueになります。
         del_item = worksheet.cell(i, 2).value #該当のサンプル名を代入
         urlitems = driver.find_elements_by_class_name("information-pane")[0].find_elements_by_class_name("media") #親要素を取得
         for urlitem in urlitems:
            itemr = urlitem.find_element_by_class_name("media-heading") #商品名を取得
            comr = urlitem.find_elements_by_class_name("item-action-count")[1] #コメントの数を取得
            if del_item == itemr.text and comr.text == "0":
               urlr = urlitem.find_elements_by_class_name("btn.btn-default")[1] #削除ボタンを取得
               print('E-2', dt.now())
               sleep(5)
               urlr.click()
               print('E-3', dt.now())
               sleep(10)
               Alert(driver).accept() #削除しますか？のはいをクリックして削除
               print('E-4', dt.now())
               sleep(5)
               break
   print('I-1', dt.now())
   ritmlists = driver.find_elements_by_class_name("media-heading")
   print('J-1', dt.now())
   rlists=[]
   for ritmlist in ritmlists:
      rlists.append(ritmlist.text)
   driver.get("https://fril.jp/item/new")
   sleep(3)
   print('J-2', dt.now())
   if not "ダイソン掃除機 タイヤ5個+テフロンテープ+トルクスドライバー3本セット" in rlists:
      driver.find_elements_by_xpath("//input[@type='file']")[0].send_keys("/Users/masa/Pictures/torx/torxset.JPG")
      driver.find_elements_by_xpath("//input[@type='file']")[1].send_keys("/Users/masa/Pictures/dyson/tire5.jpg")
      driver.find_elements_by_xpath("//input[@type='file']")[2].send_keys("/Users/masa/Pictures/dyson/tape.jpeg")
      driver.find_elements_by_xpath("//input[@type='file']")[3].send_keys("/Users/masa/Pictures/torx/torxdriver.jpeg")
      rinputElement1 = driver.find_elements_by_class_name("form-group")[1].find_element_by_css_selector("input")
      rinputElement1.send_keys("ダイソン掃除機 タイヤ5個+テフロンテープ+トルクスドライバー3本セット")
      rexp1a = "【商品紹介】\n画像の2～4枚目が商品になります。ダイソンモーターヘッドの裏側に付いているタイヤと工具です。"
      rexp1()
      rcom_to_dys()
      rinputElement8 = driver.find_elements_by_class_name("form-group")[11].find_element_by_css_selector("input")
      rinputElement8.send_keys("650")
      rclick()
      worksheet.update_cell(5, 3, str(time))
   print('J-3', dt.now())
   if not "ダイソン掃除機 タイヤ5個+テフロンテープ+シャフト4本セット" in rlists:
      driver.find_elements_by_xpath("//input[@type='file']")[0].send_keys("/Users/masa/Pictures/dyson/shaftset.JPG")
      driver.find_elements_by_xpath("//input[@type='file']")[1].send_keys("/Users/masa/Pictures/dyson/tire5.jpg")
      driver.find_elements_by_xpath("//input[@type='file']")[2].send_keys("/Users/masa/Pictures/dyson/tape.jpeg")
      driver.find_elements_by_xpath("//input[@type='file']")[3].send_keys("/Users/masa/Pictures/dyson/shaft.jpeg")
      rinputElement1 = driver.find_elements_by_class_name("form-group")[1].find_element_by_css_selector("input")
      rinputElement1.send_keys("ダイソン掃除機 タイヤ5個+テフロンテープ+シャフト4本セット")
      rexp1a = "【商品紹介】\n画像の2～4枚目が商品になります。ダイソンモーターヘッドの裏側に付いているタイヤと部品です。"
      rexp1()
      rcom_to_dys()
      rinputElement8 = driver.find_elements_by_class_name("form-group")[11].find_element_by_css_selector("input")
      rinputElement8.send_keys("650")
      rclick()
      worksheet.update_cell(6, 3, str(time))
   print('J-4', dt.now())
   if not "ダイソン掃除機 タイヤ5個+テフロンテープ+シャフト4本+トルクスドライバー3本" in rlists:
      driver.find_elements_by_xpath("//input[@type='file']")[0].send_keys("/Users/masa/Pictures/dyson/4set.JPG")
      driver.find_elements_by_xpath("//input[@type='file']")[1].send_keys("/Users/masa/Pictures/dyson/tire5set.JPG")
      driver.find_elements_by_xpath("//input[@type='file']")[2].send_keys("/Users/masa/Pictures/dyson/shaft.jpeg")
      driver.find_elements_by_xpath("//input[@type='file']")[3].send_keys("/Users/masa/Pictures/torx/torxdriver.jpeg")
      rinputElement1 = driver.find_elements_by_class_name("form-group")[1].find_element_by_css_selector("input")
      rinputElement1.send_keys("ダイソン掃除機 タイヤ5個+テフロンテープ+シャフト4本+トルクスドライバー3本")
      rexp1a = "【商品紹介】\n画像の2～4枚目が商品になります。ダイソンモーターヘッドの裏側に付いているタイヤと部品と工具です。"
      rexp1()
      rcom_to_dys()
      rinputElement8 = driver.find_elements_by_class_name("form-group")[11].find_element_by_css_selector("input")
      rinputElement8.send_keys("830")
      rclick()
      worksheet.update_cell(7, 3, str(time))
   print('J-5', dt.now())
   if not "ダイソン掃除機 タイヤ5個+テフロンテープセット" in rlists:
      driver.find_elements_by_xpath("//input[@type='file']")[0].send_keys("/Users/masa/Pictures/dyson/tire5set.JPG")
      driver.find_elements_by_xpath("//input[@type='file']")[1].send_keys("/Users/masa/Pictures/dyson/tire5.jpg")
      driver.find_elements_by_xpath("//input[@type='file']")[2].send_keys("/Users/masa/Pictures/dyson/tape.jpeg")
      driver.find_elements_by_xpath("//input[@type='file']")[3].send_keys("/Users/masa/Pictures/dyson/morterheadback.jpg")
      rinputElement1 = driver.find_elements_by_class_name("form-group")[1].find_element_by_css_selector("input")
      rinputElement1.send_keys("ダイソン掃除機 タイヤ5個+テフロンテープセット")
      rexp1a = "【商品紹介】\n画像の2～3枚目が商品になります。ダイソンモーターヘッドの裏側に付いているタイヤです。"
      rexp1()
      rcom_to_dys()
      rinputElement8 = driver.find_elements_by_class_name("form-group")[11].find_element_by_css_selector("input")
      rinputElement8.send_keys("580")
      rclick()
      worksheet.update_cell(8, 3, str(time))
   print('J-6', dt.now())
   if not "ダイソン掃除機 タイヤ2個+テフロンテープセット" in rlists:
      driver.find_elements_by_xpath("//input[@type='file']")[0].send_keys("/Users/masa/Pictures/dyson/tire2set.JPG")
      driver.find_elements_by_xpath("//input[@type='file']")[1].send_keys("/Users/masa/Pictures/dyson/tire2.jpeg")
      driver.find_elements_by_xpath("//input[@type='file']")[2].send_keys("/Users/masa/Pictures/dyson/tape.jpeg")
      driver.find_elements_by_xpath("//input[@type='file']")[3].send_keys("/Users/masa/Pictures/dyson/morterheadback.jpg")
      rinputElement1 = driver.find_elements_by_class_name("form-group")[1].find_element_by_css_selector("input")
      rinputElement1.send_keys("ダイソン掃除機 タイヤ2個+テフロンテープセット")
      rexp1a = "【商品紹介】\n画像の2～3枚目が商品になります。ダイソンモーターヘッドの裏側に付いているタイヤです。"
      rexp1()
      rcom_to_dys()
      rinputElement8 = driver.find_elements_by_class_name("form-group")[11].find_element_by_css_selector("input")
      rinputElement8.send_keys("380")
      rclick()
      worksheet.update_cell(9, 3, str(time))
   print('try-last1', dt.now())
except Exception as e:
   print(e)
   driver.quit()
print('j-8', dt.now())
driver.quit( ) #全てのウインドウを閉じる

# Open web (account 2)
options = webdriver.ChromeOptions()
options.add_argument(
   '--user-data-dir={chrom_dir_path}'.format(chrom_dir_path = '/Users/masa/Library/Application Support/Google/Chrome/Profile 2'))
try:
   driver = webdriver.Chrome(options=options, executable_path='/Users/masa/bin/chromedriver')
   print('K-1', dt.now()) # メルカリTのページにアクセス
   sleep(1)
   driver.get("https://www.mercari.com/jp/mypage/listings/listing/")
   sleep(1)
   dictm = {}
   dictcomm = {}
   urlitems = driver.find_elements_by_class_name("mypage-item-link")
   for urlitem in urlitems:
      url = urlitem.get_attribute("href")
      item = urlitem.find_element_by_class_name("mypage-item-text")
      com = urlitem.find_elements_by_class_name("listing-item-count")[1]
      dictm[item.text]=url
      dictcomm[item.text]=com.text
   for i in range(5,13):
      if dt.strptime(worksheet.cell(i,14).value, '%Y/%m/%d %H:%M:%S') < dt.strptime(worksheet.cell(i,15).value, '%Y/%m/%d %H:%M:%S'):
         del_item = worksheet.cell(i, 13).value
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
      lists.append(itmlist.text)
   itmpage()
   #  1.画像をアップロード & 商品名 & 販売価格, 2.カテゴリ & ブランド & 商品説明, 3. 商品の状態 & 配送料の負担 & 配送の方法 & 発送元の地域 & 発送までの日数, 4.出品する
   if not "ダイソン掃除機 タイヤ5個+テフロンテープ+トルクスドライバー3本セット" in lists:
      itm1()
      com_to_dys(exp1, mc)
      exh()
      worksheet.update_cell(5, 14, str(time))
   print('L-1', dt.now())
   if not "ダイソン掃除機 タイヤ5個+テフロンテープ+シャフト4本セット" in lists:
      itm2()
      com_to_dys(exp1, mc)
      exh()
      worksheet.update_cell(6, 14, str(time))
   if not "ダイソン掃除機 タイヤ5個+テフロンテープ+シャフト4本+トルクスドライバー3本" in lists:
      itm3()
      com_to_dys(exp1, mc)
      exh()
      worksheet.update_cell(7, 14, str(time))
   if not "ダイソン掃除機 タイヤ5個+テフロンテープセット" in lists:
      itm4()
      com_to_dys(exp1, mc)
      exh()
      worksheet.update_cell(8, 14, str(time))
   print('M-1', dt.now())
   if not "ダイソン掃除機 タイヤ5個" in lists:
      itm5()
      com_to_dys(exp1, mc)
      exh()
      worksheet.update_cell(11, 14, str(time))
   if not "ダイソン掃除機 タイヤ2個+テフロンテープセット" in lists:
      itm6()
      com_to_dys(exp1, mc)
      exh()
      worksheet.update_cell(9, 14, str(time))
   if not "ダイソン掃除機 タイヤ2個" in lists:
      itm7()
      com_to_dys(exp1, mc)
      exh()
      worksheet.update_cell(12, 14, str(time))
   if not "トルクスドライバー3本セット（T10 & T8 & T6）" in lists:
      itm8()
      com_to_dys(exp1, mc)
      exh()
      worksheet.update_cell(10, 14, str(time))
   print('try-last2', dt.now())
except Exception as e:
   print(e)
   driver.quit()
driver.quit() #全てのウインドウを閉じる