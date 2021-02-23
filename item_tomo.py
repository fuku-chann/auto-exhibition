print('1')
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
   driver.find_element_by_xpath("//input[@type='file']").send_keys("/Users/masa/Pictures/torx/torxset.JPG \n /Users/masa/Pictures/dyson/tire4.png \n /Users/masa/Pictures/dyson/spacers.png \n /Users/masa/Pictures/torx/torxdriver.jpeg \n /Users/masa/Pictures/torx/torxside.jpg \n /Users/masa/Pictures/torx/torxsize.JPG \n /Users/masa/Pictures/dyson/morterhead.jpeg")
   inputElement1 = driver.find_elements_by_class_name("style_inputarea__1mtsA")[0].find_element_by_css_selector("input")
   inputElement1.send_keys("ダイソン掃除機 タイヤ4個+スペーサー+トルクスドライバー3本セット")
   inputElement3 = driver.find_elements_by_class_name("style_inputarea__1mtsA")[2].find_element_by_css_selector("input")
   inputElement3.send_keys("630")
   global exp1a
   exp1a = "【商品紹介】\n画像の2～4枚目が商品になります。ダイソンモーターヘッドの裏側に付いているタイヤと工具です。"
def itm2():
   driver.find_element_by_xpath("//input[@type='file']").send_keys("/Users/masa/Pictures/dyson/shaftset.JPG \n /Users/masa/Pictures/dyson/tire4.png \n /Users/masa/Pictures/dyson/spacers.png \n /Users/masa/Pictures/dyson/shaft.png \n /Users/masa/Pictures/dyson/morterhead.jpeg")
   inputElement1 = driver.find_elements_by_class_name("style_inputarea__1mtsA")[0].find_element_by_css_selector("input")
   inputElement1.send_keys("ダイソン掃除機 タイヤ4個+スペーサー+シャフト4本セット")
   inputElement3 = driver.find_elements_by_class_name("style_inputarea__1mtsA")[2].find_element_by_css_selector("input")
   inputElement3.send_keys("630")
   global exp1a
   exp1a = "【商品紹介】\n画像の2～4枚目が商品になります。ダイソンモーターヘッドの裏側に付いているタイヤと部品です。"
def itm3():
   driver.find_element_by_xpath("//input[@type='file']").send_keys("/Users/masa/Pictures/dyson/4set.JPG \n /Users/masa/Pictures/dyson/tire4.png \n /Users/masa/Pictures/dyson/spacers.png \n /Users/masa/Pictures/torx/torxdriver.jpeg \n /Users/masa/Pictures/torx/torxside.jpg \n /Users/masa/Pictures/torx/torxsize.JPG \n /Users/masa/Pictures/dyson/shaft.png \n /Users/masa/Pictures/dyson/morterhead.jpeg")
   inputElement1 = driver.find_elements_by_class_name("style_inputarea__1mtsA")[0].find_element_by_css_selector("input")
   inputElement1.send_keys("ダイソン掃除機 タイヤ4個+スペーサー+シャフト4本+トルクスドライバー3本")
   inputElement3 = driver.find_elements_by_class_name("style_inputarea__1mtsA")[2].find_element_by_css_selector("input")
   inputElement3.send_keys("730")
   global exp1a
   exp1a = "【商品紹介】\n画像の2～5枚目が商品になります。ダイソンモーターヘッドの裏側に付いているタイヤと部品と工具です。"
def itm4():
   driver.find_element_by_xpath("//input[@type='file']").send_keys("/Users/masa/Pictures/dyson/tire4set.JPG \n /Users/masa/Pictures/dyson/tire4.png \n /Users/masa/Pictures/dyson/spacers.png \n /Users/masa/Pictures/dyson/morterhead.jpeg")
   inputElement1 = driver.find_elements_by_class_name("style_inputarea__1mtsA")[0].find_element_by_css_selector("input")
   inputElement1.send_keys("ダイソン掃除機 タイヤ4個+スペーサーセット")
   inputElement3 = driver.find_elements_by_class_name("style_inputarea__1mtsA")[2].find_element_by_css_selector("input")
   inputElement3.send_keys("550")
   global exp1a
   exp1a = "【商品紹介】\n画像の2～3枚目が商品になります。ダイソンモーターヘッドの裏側に付いているタイヤです。"
def itm5():
   driver.find_element_by_xpath("//input[@type='file']").send_keys("/Users/masa/Pictures/dyson/tireimage.png \n /Users/masa/Pictures/dyson/tire4.png \n /Users/masa/Pictures/dyson/morterhead.jpeg")
   inputElement1 = driver.find_elements_by_class_name("style_inputarea__1mtsA")[0].find_element_by_css_selector("input")
   inputElement1.send_keys("ダイソン掃除機 タイヤ4個（リピーター様専用）")
   inputElement3 = driver.find_elements_by_class_name("style_inputarea__1mtsA")[2].find_element_by_css_selector("input")
   inputElement3.send_keys("450")
   global exp1a
   exp1a = "【商品紹介】\n画像の2枚目が商品になります。ダイソンモーターヘッドの裏側に付いているタイヤです。"
def itm6():
   driver.find_element_by_xpath("//input[@type='file']").send_keys("/Users/masa/Pictures/dyson/tire2set.JPG \n /Users/masa/Pictures/dyson/tire2.jpeg \n /Users/masa/Pictures/dyson/spacers.png \n /Users/masa/Pictures/dyson/morterhead.jpeg")
   inputElement1 = driver.find_elements_by_class_name("style_inputarea__1mtsA")[0].find_element_by_css_selector("input")
   inputElement1.send_keys("ダイソン掃除機 タイヤ2個+スペーサーセット")
   inputElement3 = driver.find_elements_by_class_name("style_inputarea__1mtsA")[2].find_element_by_css_selector("input")
   inputElement3.send_keys("400")
   global exp1a
   exp1a = "【商品紹介】\n画像の2～3枚目が商品になります。ダイソンモーターヘッドの裏側に付いているタイヤです。"
def itm7():
   driver.find_element_by_xpath("//input[@type='file']").send_keys("/Users/masa/Pictures/dyson/tireimage.png \n /Users/masa/Pictures/dyson/tire2.jpeg \n /Users/masa/Pictures/dyson/morterhead.jpeg")
   inputElement1 = driver.find_elements_by_class_name("style_inputarea__1mtsA")[0].find_element_by_css_selector("input")
   inputElement1.send_keys("ダイソン掃除機 タイヤ2個（リピーター様専用）")
   inputElement3 = driver.find_elements_by_class_name("style_inputarea__1mtsA")[2].find_element_by_css_selector("input")
   inputElement3.send_keys("300")
   global exp1a
   exp1a = "【商品紹介】\n画像の2枚目が商品になります。ダイソンモーターヘッドの裏側に付いているタイヤです。"
def itm8():
   driver.find_element_by_xpath("//input[@type='file']").send_keys("/Users/masa/Pictures/torx/torxdriver.jpeg \n /Users/masa/Pictures/torx/torxside.jpg \n /Users/masa/Pictures/torx/torxsize.JPG")
   inputElement1 = driver.find_elements_by_class_name("style_inputarea__1mtsA")[0].find_element_by_css_selector("input")
   inputElement1.send_keys("トルクスドライバー3本セット（T10 & T8 & T6）")
   inputElement3 = driver.find_elements_by_class_name("style_inputarea__1mtsA")[2].find_element_by_css_selector("input")
   inputElement3.send_keys("430")
   global exp1a
   exp1a = "トルクスドライバー3本セット（T10 & T8 & T6）\
   \n\n ダイソン分解清掃などにお得な3本セットです。"
def itm9():
   driver.find_element_by_xpath("//input[@type='file']").send_keys("/Users/masa/Pictures/dyson/tire4(large)set.JPG \n /Users/masa/Pictures/dyson/tire41mm.jpg \n /Users/masa/Pictures/dyson/spacer(縦大).png")
   inputElement1 = driver.find_elements_by_class_name("style_inputarea__1mtsA")[0].find_element_by_css_selector("input")
   inputElement1.send_keys("ダイソン掃除機 タイヤ4個(大)+スペーサーセット")
   inputElement3 = driver.find_elements_by_class_name("style_inputarea__1mtsA")[2].find_element_by_css_selector("input")
   inputElement3.send_keys("800")
   global exp1a
   exp1a = "【商品紹介】\n画像の2～3枚目が商品になります。ダイソンモーターヘッドの裏側に付いているタイヤです。"
def itm10():
   driver.find_element_by_xpath("//input[@type='file']").send_keys("/Users/masa/Pictures/dyson/tire2(large)set.JPG \n /Users/masa/Pictures/dyson/tire21mm.jpg \n /Users/masa/Pictures/dyson/spacer(縦大).png")
   inputElement1 = driver.find_elements_by_class_name("style_inputarea__1mtsA")[0].find_element_by_css_selector("input")
   inputElement1.send_keys("ダイソン掃除機 タイヤ2個(大)+スペーサーセット")
   inputElement3 = driver.find_elements_by_class_name("style_inputarea__1mtsA")[2].find_element_by_css_selector("input")
   inputElement3.send_keys("650")
   global exp1a
   exp1a = "【商品紹介】\n画像の2～3枚目が商品になります。ダイソンモーターヘッドの裏側に付いているタイヤです。"
def itm11():
   driver.find_element_by_xpath("//input[@type='file']").send_keys("/Users/masa/Pictures/torx/torxset(large)2.JPG \n /Users/masa/Pictures/dyson/tire41mm.jpg \n /Users/masa/Pictures/dyson/spacer(縦大).png \n /Users/masa/Pictures/torx/torxdriver.jpeg \n /Users/masa/Pictures/torx/torxside.jpg \n /Users/masa/Pictures/torx/torxsize.JPG")
   inputElement1 = driver.find_elements_by_class_name("style_inputarea__1mtsA")[0].find_element_by_css_selector("input")
   inputElement1.send_keys("ダイソン掃除機 タイヤ4個(大)+スペーサー+トルクスドライバー3本セット")
   inputElement3 = driver.find_elements_by_class_name("style_inputarea__1mtsA")[2].find_element_by_css_selector("input")
   inputElement3.send_keys("880")
   global exp1a
   exp1a = "【商品紹介】\n画像の2～4枚目が商品になります。ダイソンモーターヘッドの裏側に付いているタイヤです。"
def itm12():
   driver.find_element_by_xpath("//input[@type='file']").send_keys("/Users/masa/Pictures/torx/torxset(soft).jpeg \n /Users/masa/Pictures/dyson/tire41mm.jpg \n /Users/masa/Pictures/torx/torxdriver.jpeg \n /Users/masa/Pictures/torx/torxside.jpg \n /Users/masa/Pictures/torx/torxsize.JPG \n /Users/masa/Pictures/dyson/soft.jpeg")
   inputElement1 = driver.find_elements_by_class_name("style_inputarea__1mtsA")[0].find_element_by_css_selector("input")
   inputElement1.send_keys("ダイソン掃除機 タイヤ4個(大)+トルクスドライバー3本セット")
   inputElement3 = driver.find_elements_by_class_name("style_inputarea__1mtsA")[2].find_element_by_css_selector("input")
   inputElement3.send_keys("780")
   global exp1a
   exp1a = "【商品紹介】\n画像の1枚目が商品になります。ダイソン掃除機ソフトローラークリーンヘッドの裏側に付いているタイヤです。"
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
   \nタイヤが劣化すると約１年ほどでフローリングや畳が傷つくことがあります。そのため、定期的に交換する必要があります。\
   \n\n【交換方法】\n交換方法などブログで詳しく紹介していますので、必ず購入前にご確認ください。\nブラウザでURLを入力すると検索できます。(URL：fuku-channnel.com)\nブログが検索できたら、サイト内検索で「ダイソン」と入力下さい。\
   \n\n【保証内容】\n対象機種以外で交換をご検討されている場合、下記2点をご対応いただくことで保証が適応されることがあります。(対象機種は保証が適用されます)\n\n1. コメントにてご利用の機種をご連絡ください。\n2. ヘッドの裏の写真(タイヤ部分がわかる写真)を出品下さい。(価格は100,000としてください。)\nこちらで取り付けできると判断した場合のみ、保証が適応されます。保証が適応された場合、万が一取り付けできなかったら返金いたしますので、その旨ご連絡ください。\
   \n\n【交換頻度】\n１年以内の交換をおすすめしています。\
   \n\n【現状のメーカー修理】\nメーカーで修理依頼すると、購入から2年以上経過すると最低5000円の工賃が発生します。\
   \n\n【最後に】\nタイヤは簡単に交換できるので、DIYで行うことをおすすめしています。(機種によって右後輪の取り外しが困難な場合があります)\nメーカーから部品だけの購入が不可と言われたため、自分で作成したものを出品しています。\n最近、ご購入者様からお褒めの言葉を多く頂いております。モチベーションの向上に繋がりますので、評価の方にもコメントいただけたら幸いです。"
exp2a = "\
   \n\n【対象機種】\n実績機種は、DC35, DC48, DC59, DC62, V6です。\nその他の機種でも適合する可能性があります。(その他の機種でタイヤ交換をご検討されている方は次の保証内容をご確認ください。)\
   \n\n【タイヤのサイズ】\n外径: 10 mm\n内径: 3 mm\n幅: 4.5 mm\nシャフトにスペーサーを装着することで取り付けできます"
exp3a = "\
   \n\n【対象機種】\n実績機種は、DC48, DC63, CY25ダイレクトドライブです。\nその他の機種でも適合する可能性があります。(その他の機種でタイヤ交換をご検討されている方は次の保証内容をご確認ください。)\
   \n\n【タイヤのサイズ】\n外径: 10 mm\n内径: 3 mm\n幅: 10 mm \nシャフトにスペーサーを装着することで取り付けできます"
exp4a = "\
   \n\n【対象機種】\n実績機種は、ソフトローラークリーンヘッド搭載機種です。\nその他の機種でも適合する可能性があります。(その他の機種でタイヤ交換をご検討されている方は次の保証内容をご確認ください。)\
   \n\n【タイヤのサイズ】\n外径: 10 mm\n内径: 3 mm\n幅: 10 mm \n取り外し方などは上記ブログをご覧下さい。"
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

# Open web (account 2)
options = webdriver.ChromeOptions()
options.add_argument(
   '--user-data-dir={chrom_dir_path}'.format(chrom_dir_path = '/Users/masa/Library/Application Support/Google/Chrome/Profile 2'))
options.add_argument('--headless')
options.add_argument('--no-sandbox')
options.add_argument('--disable-gpu')
try:
   driver = webdriver.Chrome(options=options, executable_path='/Users/masa/bin/chromedriver')
   # driver.minimize_window()
   time = dt.now()
   worksheet.update('D5:D16', [[str(time)], [str(time)], [str(time)], [str(time)], [str(time)], [str(time)], [str(time)], [str(time)], [str(time)], [str(time)], [str(time)], [str(time)]])
   print('K-1', dt.now())
   sleep(1)
   driver.get("https://www.mercari.com/jp/mypage/listings/listing/") # メルカリTのページにアクセス
   sleep(1)
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
   for i in range(5,17):
      sleep(1)
      if dt.strptime(worksheet.cell(i, 14).value, '%Y/%m/%d %H:%M:%S') < dt.strptime(worksheet.cell(i, 15).value, '%Y/%m/%d %H:%M:%S'):
         del_item = worksheet.cell(i, 13).value
         print('del_item', del_item)
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
   if not "ダイソン掃除機 タイヤ4個+スペーサー+トルクスドライバー3本セット" in lists:
      itm1()
      com_to_dys(exp1, mc)
      exh()
      worksheet.update_cell(5, 14, str(time))
   if not "ダイソン掃除機 タイヤ4個+スペーサー+シャフト4本セット" in lists:
      itm2()
      com_to_dys(exp1, mc)
      exh()
      worksheet.update_cell(6, 14, str(time))
   if not "ダイソン掃除機 タイヤ4個+スペーサー+シャフト4本+トルクスドライバー3本" in lists:
      itm3()
      com_to_dys(exp1, mc)
      exh()
      worksheet.update_cell(7, 14, str(time))
   print('C-1', dt.now())
   if not "ダイソン掃除機 タイヤ4個+スペーサーセット" in lists:
      itm4()
      com_to_dys(exp1, mc)
      exh()
      worksheet.update_cell(8, 14, str(time))
   if not "ダイソン掃除機 タイヤ4個（リピーター様専用）" in lists:
      itm5()
      com_to_dys(exp1, mc)
      exh()
      worksheet.update_cell(11, 14, str(time))
   if not "ダイソン掃除機 タイヤ2個+スペーサーセット" in lists:
      itm6()
      com_to_dys(exp1, mc)
      exh()
      worksheet.update_cell(9, 14, str(time))
   if not "ダイソン掃除機 タイヤ2個（リピーター様専用）" in lists:
      itm7()
      com_to_dys(exp1, mc)
      exh()
      worksheet.update_cell(12, 14, str(time))
   if not "トルクスドライバー3本セット（T10 & T8 & T6）" in lists:
      itm8()
      com_to_dys(exp1, mc)
      exh()
      worksheet.update_cell(10, 14, str(time))
   if not "ダイソン掃除機 タイヤ4個(大)+スペーサーセット" in lists:
      itm9()
      com_to_dys(exp1, mc)
      exh()
      worksheet.update_cell(13, 14, str(time))
   if not "ダイソン掃除機 タイヤ2個(大)+スペーサーセット" in lists:
      itm10()
      com_to_dys(exp1, mc)
      exh()
      worksheet.update_cell(14, 14, str(time))
   if not "ダイソン掃除機 タイヤ4個(大)+スペーサー+トルクスドライバー3本セット" in lists:
      itm11()
      com_to_dys(exp1, mc)
      exh()
      worksheet.update_cell(15, 14, str(time))
   print('C-2', dt.now())
   if not "ダイソン掃除機 タイヤ4個(大)+トルクスドライバー3本セット" in lists:
      itm12()
      com_to_dys(exp1, mc)
      exh()
      worksheet.update_cell(16, 14, str(time))
   print('D-1', dt.now())
except Exception as e:
   print(e)
   # driver.quit()


# ラクマページにアクセス
def rexp1(rinputElement1):
   rinputElement3 = driver.find_elements_by_class_name("form-group")[2].find_element_by_css_selector("textarea")
   rinputElement3.send_keys(rexp1a + exp0a)
   print('rexp1')
# #ダイソン共通  rakumacommon_to_cut: 
def rcom_to_cut():
   driver.find_elements_by_xpath("//input[@type='file']")[0].send_keys("/Users/masa/Pictures/tube_cutter/spares.JPG")
   driver.find_elements_by_xpath("//input[@type='file']")[1].send_keys("/Users/masa/Pictures/tube_cutter/spare.jpeg")
   driver.find_elements_by_xpath("//input[@type='file']")[2].send_keys("/Users/masa/Pictures/tube_cutter/tube_cutter2.jpeg")
   # driver.find_elements_by_xpath("//input[@type='file']")[3].send_keys("/Users/masa/Pictures/tube_cutter/")
   rinputElement1 = driver.find_elements_by_class_name("form-group")[1].find_element_by_css_selector("input")
   rinputElement1.send_keys(itm_name)
   rinputElement3 = driver.find_elements_by_class_name("form-group")[2].find_element_by_css_selector("textarea")
   rinputElement3.send_keys(itm_name + "です。画像の1枚目が商品になります。サイズは写真１枚目の通りです。画像３枚目の取り付け用ケースは付きません。\
      \n\nTC16SC\
      \n千代田通商\
      \nチヨダ\
      \nchiyoda\
      \nCHIYODA")
   rcateElement1 = driver.find_element_by_id("category_name")
   rcateElement1.click()
   rcateElement2 = driver.find_elements_by_class_name("panel.list-group")[7]
   rcateElement2.click()
   sleep(2)
   rcateElement3 = rcateElement2.find_elements_by_class_name("list-group-item.small.branch")[11]
   # rcateElement3 = driver.find_elements_by_class_name("panel.list-group")[7].find_elements_by_class_name("list-group-item.small.branch")[11]
   rcateElement3.click()
   print('227', rcateElement3.text)
   sleep(2)
   rcateElement4 = driver.find_elements_by_class_name("sublinks.collapse.in")[1].find_elements_by_class_name("list-group-item.small.leaf")[2]
   print(230, rcateElement4.text)
   # rcateElement4 = rcateElement2.find_elements_by_class_name("sublinks.collapse.in")[1].find_elements_by_class_name("list-group-item.small.leaf")[2]
   # rcateElement4 = driver.find_elements_by_class_name("panel.list-group")[7].find_elements_by_class_name("list-group-item.small.branch")[11].find_elements_by_class_name("list-group-item.small.leaf")[1]
   rcateElement4.click()
   sleep(2)
   rcateElement5 = driver.find_element_by_id("delivery_method")
   # print(349, rcateElement5)
   rcateElement5.click()
   sleep(2)
   rcateElement6 = driver.find_elements_by_class_name("modal-body")[3].find_elements_by_class_name("list-group-item")[12]
   # print(353, rcateElement6.text)
   sleep(2)
   rcateElement6.click()
   sleep(2)
   rcateElement7 = driver.find_elements_by_class_name("col-sm-9.col-xs-8")[8]
   rcateElement8 = rcateElement7.find_element_by_class_name("form-control")
   # print(242, rcateElement8.text)
   rcateElement9 = Select(rcateElement8).options
   print(243, rcateElement9[5].text)
   rcateElement10 = rcateElement9[11]
   rcateElement10.click()
   rinputElement8 = driver.find_elements_by_class_name("form-group")[11].find_element_by_css_selector("input")
   rinputElement8.send_keys(cost)
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

driver.get("https://fril.jp/sell")
for i in range(5,7): #商品追加したら数字を増やす
   print('i', i, '387')
   if dt.strptime(worksheet.cell(i, 19).value, '%Y/%m/%d %H:%M:%S') < dt.strptime(worksheet.cell(i, 20).value, '%Y/%m/%d %H:%M:%S'): #出品してから１日以上経過している場合、trueになります。
      del_item = worksheet.cell(i, 18).value #該当のサンプル名を代入
      urlitems = driver.find_elements_by_class_name("information-pane")[0].find_elements_by_class_name("media") #親要素を取得
      # print('E-1-1', dt.now())
      for urlitem in urlitems:
         itemr = urlitem.find_element_by_class_name("media-heading") #商品名を取得
         comr = urlitem.find_elements_by_class_name("item-action-count")[1] #コメントの数を取得
         # print('E-1-2', dt.now(),'itemr', itemr)
         if del_item == itemr.text and comr.text == "0":
            print('344')
            urlr = urlitem.find_elements_by_class_name("btn.btn-default")[1] #削除ボタンを取得
            # print('E-2', dt.now())
            sleep(5)
            urlr.click()
            # print('E-3', dt.now())
            sleep(10)
            Alert(driver).accept() #削除しますか？のはいをクリックして削除
            # print('E-4', dt.now())
            sleep(5)
            break

print('J-1', dt.now())
rlists=[]
sleep(3)
ritmlists = driver.find_elements_by_class_name("media-heading")
for ritmlist in ritmlists:
   rlists.append(ritmlist.text)
driver.get("https://fril.jp/item/new")
sleep(3)
itm_name = "チューブカッター用替刃 5枚セット"
if not itm_name in rlists:
   cost = 500
   rcom_to_cut()
   rclick()
   worksheet.update_cell(5, 19, str(time))
itm_name = "チューブカッター用替刃 10枚セット"
if not itm_name in rlists:
   cost = 950
   rcom_to_cut()
   rclick()
   worksheet.update_cell(6, 19, str(time))
driver.quit() #全てのウインドウを閉じる