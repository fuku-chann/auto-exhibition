from selenium import webdriver
from selenium.webdriver import ActionChains
from time import sleep
import chromedriver_binary
import webbrowser
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.select import Select
import datetime

# 入力（共通部分）
def itmpage():
   sleep(1)
   driver.get("https://www.mercari.com/jp/sell/")
   sleep(1)
def sct(cateElement,indexNum):
   select = Select(cateElement)
   select.select_by_index(indexNum)
# ダイソン共通  ctd: 1-9.カテゴリー, 10-11.ブランド, 12-13.商品の説明
def ctd(exp1): 
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
   inputElement1 = driver.find_elements_by_class_name("style_body__1OP1S")[2].find_element_by_css_selector("textarea")
   exp1(inputElement1,exp1a)
def exp1(inputElement1,exp1a):
   inputElement1.send_keys(exp1a + "\
   \nタイヤが劣化すると約１年ほどでフローリングや畳が傷つくことがあります。そのため、定期的に交換する必要があります。\
   \n\n【交換方法】\n交換方法などブログで詳しく紹介していますので、必ず購入前にご確認ください。\nブラウザでURLを入力すると検索できます。(URL：fuku-channnel.com)\nブログが検索できたら、サイト内検索で「ダイソン」と入力下さい。\
   \n\n【対象機種】\n実績機種は、DC35, DC48, DC59, DC62, V6です。\nその他の機種でも適合する可能性があります。(その他の機種でタイヤ交換をご検討されている方は次の保証内容をご確認ください。)\
   \n\n【保証内容】\n対象機種以外で交換をご検討されている場合、下記2点をご対応いただくことで保証が適応されることがあります。(対象機種は保証が適用されます)\n1. コメントにてご利用の機種をご連絡ください。\n2. ヘッドの裏の写真(タイヤ部分がわかる写真)を出品下さい。(価格は9,999,999としてください。)\nこちらで取り付けできると判断した場合のみ、保証が適応されます。保証が適応された場合、万が一取り付けできなかったら返金いたしますので、その旨ご連絡ください。\
   \n\n【交換頻度】\n１年以内の交換をおすすめしています。\n【タイヤのサイズ】\n外径: 10 mm\n内径: 3 mm\n幅: 4.5 mm (予備タイヤは3.5 mm)\nシャフトにシールを巻くことを前提にしています。詳しくは上記ブログ（タイヤの取り付け方）をご覧下さい。\
   \n\n【現状のメーカー修理】\nメーカーで修理依頼すると、購入から2年以上経過すると最低5000円の工賃が発生します。\
   \n\n【最後に】\nタイヤは簡単に交換できるので、DIYで行うことをおすすめしています。(機種によって右後輪の取り外しが困難な場合があります)\nメーカーから部品だけの購入が不可と言われたため、自分で作成したものを出品しています。\n最近、ご購入者様からお褒めの言葉を多く頂いております。モチベーションの向上に繋がりますので、評価の方にもコメントいただけたら幸いです。")
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
# itmx: 1.画像をアップロード, 2-3.商品名, 4-5.販売価格
def itm1():
   driver.find_element_by_xpath("//input[@type='file']").send_keys("/Users/masa/Pictures/torx/torxset.JPG \n /Users/masa/Pictures/dyson/tire5.jpg \n /Users/masa/Pictures/dyson/tape.jpeg \n /Users/masa/Pictures/torx/torxdriver.jpeg \n /Users/masa/Pictures/dyson/moeterhead.jpeg \n /Users/masa/Pictures/dyson/morterheadback.jpeg")
   inputElement1 = driver.find_elements_by_class_name("style_inputarea__1mtsA")[0].find_element_by_css_selector("input")
   inputElement1.send_keys("ダイソン掃除機 タイヤ5個+テフロンテープ+トルクスドライバー3本セット")
   inputElement3 = driver.find_elements_by_class_name("style_inputarea__1mtsA")[2].find_element_by_css_selector("input")
   inputElement3.send_keys("680")
def itm2():
   driver.find_element_by_xpath("//input[@type='file']").send_keys("/Users/masa/Pictures/dyson/shaftset.JPG \n /Users/masa/Pictures/dyson/tire5.jpg \n /Users/masa/Pictures/dyson/tape.jpeg \n /Users/masa/Pictures/dyson/shaft.jpeg \n /Users/masa/Pictures/dyson/moeterhead.jpeg \n /Users/masa/Pictures/dyson/morterheadback.jpeg")
   inputElement1 = driver.find_elements_by_class_name("style_inputarea__1mtsA")[0].find_element_by_css_selector("input")
   inputElement1.send_keys("ダイソン掃除機 タイヤ5個+テフロンテープ+シャフト4本セット")
   inputElement3 = driver.find_elements_by_class_name("style_inputarea__1mtsA")[2].find_element_by_css_selector("input")
   inputElement3.send_keys("680")
def itm3():
   driver.find_element_by_xpath("//input[@type='file']").send_keys("/Users/masa/Pictures/dyson/4set.JPG \n /Users/masa/Pictures/dyson/tire5.jpg \n /Users/masa/Pictures/dyson/tape.jpeg \n /Users/masa/Pictures/torx/torxdriver.jpeg \n /Users/masa/Pictures/dyson/shaft.jpeg \n /Users/masa/Pictures/dyson/moeterhead.jpeg \n /Users/masa/Pictures/dyson/morterheadback.jpeg")
   inputElement1 = driver.find_elements_by_class_name("style_inputarea__1mtsA")[0].find_element_by_css_selector("input")
   inputElement1.send_keys("ダイソン掃除機 タイヤ5個+テフロンテープ+シャフト4本+トルクスドライバー3本")
   inputElement3 = driver.find_elements_by_class_name("style_inputarea__1mtsA")[2].find_element_by_css_selector("input")
   inputElement3.send_keys("880")
def itm4():
   driver.find_element_by_xpath("//input[@type='file']").send_keys("/Users/masa/Pictures/dyson/tire5set.JPG \n /Users/masa/Pictures/dyson/tire5.jpg \n /Users/masa/Pictures/dyson/tape.jpeg \n /Users/masa/Pictures/dyson/moeterhead.jpeg \n /Users/masa/Pictures/dyson/morterheadback.jpeg")
   inputElement1 = driver.find_elements_by_class_name("style_inputarea__1mtsA")[0].find_element_by_css_selector("input")
   inputElement1.send_keys("ダイソン掃除機 タイヤ5個+テフロンテープセット")
   inputElement3 = driver.find_elements_by_class_name("style_inputarea__1mtsA")[2].find_element_by_css_selector("input")
   inputElement3.send_keys("600")
def itm5():
   driver.find_element_by_xpath("//input[@type='file']").send_keys("/Users/masa/Pictures/dyson/tireimage.png \n /Users/masa/Pictures/dyson/tire5.jpg \n /Users/masa/Pictures/dyson/moeterhead.jpeg \n /Users/masa/Pictures/dyson/morterheadback.jpeg")
   inputElement1 = driver.find_elements_by_class_name("style_inputarea__1mtsA")[0].find_element_by_css_selector("input")
   inputElement1.send_keys("ダイソン掃除機 タイヤ5個")
   inputElement3 = driver.find_elements_by_class_name("style_inputarea__1mtsA")[2].find_element_by_css_selector("input")
   inputElement3.send_keys("500")
def itm6():
   driver.find_element_by_xpath("//input[@type='file']").send_keys("/Users/masa/Pictures/dyson/tire2set.JPG \n /Users/masa/Pictures/dyson/tire2.jpeg \n /Users/masa/Pictures/dyson/tape.jpeg \n /Users/masa/Pictures/dyson/moeterhead.jpeg \n /Users/masa/Pictures/dyson/morterheadback.jpeg")
   inputElement1 = driver.find_elements_by_class_name("style_inputarea__1mtsA")[0].find_element_by_css_selector("input")
   inputElement1.send_keys("ダイソン掃除機 タイヤ2個+テフロンテープセット")
   inputElement3 = driver.find_elements_by_class_name("style_inputarea__1mtsA")[2].find_element_by_css_selector("input")
   inputElement3.send_keys("400")
def itm7():
   driver.find_element_by_xpath("//input[@type='file']").send_keys("/Users/masa/Pictures/dyson/tireimage.png \n /Users/masa/Pictures/dyson/tire2.jpeg \n /Users/masa/Pictures/dyson/moeterhead.jpeg \n /Users/masa/Pictures/dyson/morterheadback.jpeg")
   inputElement1 = driver.find_elements_by_class_name("style_inputarea__1mtsA")[0].find_element_by_css_selector("input")
   inputElement1.send_keys("ダイソン掃除機 タイヤ2個")
   inputElement3 = driver.find_elements_by_class_name("style_inputarea__1mtsA")[2].find_element_by_css_selector("input")
   inputElement3.send_keys("300")
def itm8():
   driver.find_element_by_xpath("//input[@type='file']").send_keys("/Users/masa/Pictures/torx/torxdriver.jpeg")
   inputElement1 = driver.find_elements_by_class_name("style_inputarea__1mtsA")[0].find_element_by_css_selector("input")
   inputElement1.send_keys("トルクスドライバー3本セット（T10 & T8 & T6）")
   inputElement3 = driver.find_elements_by_class_name("style_inputarea__1mtsA")[2].find_element_by_css_selector("input")
   inputElement3.send_keys("430")
   # 商品の説明
   inputElement1 = driver.find_elements_by_class_name("style_body__1OP1S")[2].find_element_by_css_selector("textarea")
   inputElement1.send_keys("トルクスドライバー3本セット（T10 & T8 & T6）\
   \n\n ダイソン分解清掃などにお得な3本セットです。")

# Open web (account 1)
options = webdriver.ChromeOptions()
options.add_argument(
   '--user-data-dir={chrom_dir_path}'.format(chrom_dir_path = '/Users/masa/Library/Application Support/Google/Chrome/Profile 3'))
driver = webdriver.Chrome(options=options, executable_path='/usr/local/bin/chromedriver')
print(1,datetime.datetime.now())
# メルカリのページにアクセス
sleep(1)
driver.get("https://www.mercari.com/jp/mypage/listings/listing/")
sleep(1)
# 出品中の商品リストを作成
itmlists = driver.find_elements_by_class_name("mypage-item-text")
lists = []
for itmlist in itmlists:
   lists.append(itmlist.text)
exp1a = "【商品紹介】\n画像の2～4枚目が商品になります。ダイソンモーターヘッドの裏側に付いているタイヤと工具です。"
itmpage()
# 1.画像をアップロード & 商品名 & 販売価格, 2.カテゴリ & ブランド & 商品説明, 3. 商品の状態 & 配送料の負担 & 配送の方法 & 発送元の地域 & 発送までの日数, 4.出品する
if not "ダイソン掃除機 タイヤ5個+テフロンテープ+トルクスドライバー3本セット" in lists:
   itm1()
   ctd(exp1)
   mc()
   exh()
exp1a = "【商品紹介】\n画像の2～4枚目が商品になります。ダイソンモーターヘッドの裏側に付いているタイヤと部品です。"
if not "ダイソン掃除機 タイヤ5個+テフロンテープ+シャフト4本セット" in lists:
   itm2()
   ctd(exp1)
   mc()
   exh()
exp1a = "【商品紹介】\n画像の2～5枚目が商品になります。ダイソンモーターヘッドの裏側に付いているタイヤと部品と工具です。"
if not "ダイソン掃除機 タイヤ5個+テフロンテープ+シャフト4本+トルクスドライバー3本" in lists:
   itm3()
   ctd(exp1)
   mc()
   exh()
print(2,datetime.datetime.now())
exp1a = "【商品紹介】\n画像の2～3枚目が商品になります。ダイソンモーターヘッドの裏側に付いているタイヤです。"
if not "ダイソン掃除機 タイヤ5個+テフロンテープセット" in lists:
   itm4()
   ctd(exp1)
   mc()
   exh()
exp1a = "【商品紹介】\n画像の2枚目が商品になります。ダイソンモーターヘッドの裏側に付いているタイヤです。"
if not "ダイソン掃除機 タイヤ5個" in lists:
   itm5()
   ctd(exp1)
   mc()
   exh()
exp1a = "【商品紹介】\n画像の2～3枚目が商品になります。ダイソンモーターヘッドの裏側に付いているタイヤです。"
if not "ダイソン掃除機 タイヤ2個+テフロンテープセット" in lists:
   itm6()
   ctd(exp1)
   mc()
   exh()
exp1a = "【商品紹介】\n画像の2枚目が商品になります。ダイソンモーターヘッドの裏側に付いているタイヤです。"
if not "ダイソン掃除機 タイヤ2個" in lists:
   itm7()
   ctd(exp1)
   mc()
   exh()
print(3,datetime.datetime.now())
# 1.画像をアップロード & 商品名 & 商品説明 & 販売価格, 2-10.カテゴリー, 11-12ブランド, 13.商品の状態 & 配送料の負担 & 配送の方法 & 発送元の地域 & 発送までの日数, 14.出品する
if not "トルクスドライバー3本セット（T10 & T8 & T6）" in lists:
   itm8()
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
   mc()
   exh()
# ラクマページにアクセス
def rexp1():
   rinputElement3 = driver.find_elements_by_class_name("form-group")[2].find_element_by_css_selector("textarea")
   rinputElement3.send_keys(rexp1a + "\
   \nタイヤが劣化すると約１年ほどでフローリングや畳が傷つくことがあります。そのため、定期的に交換する必要があります。\
   \n\n【交換方法】\n交換方法などブログで詳しく紹介していますので、必ず購入前にご確認ください。\nブラウザでURLを入力すると検索できます。(URL：fuku-channnel.com)\nブログが検索できたら、サイト内検索で「ダイソン」と入力下さい。\
   \n\n【対象機種】\n実績機種は、DC35, DC48, DC59, DC62, V6です。\nその他の機種でも適合する可能性があります。(その他の機種でタイヤ交換をご検討されている方は次の保証内容をご確認ください。)\
   \n\n【保証内容】\n対象機種以外で交換をご検討されている場合、下記2点をご対応いただくことで保証が適応されることがあります。(対象機種は保証が適用されます)\n1. コメントにてご利用の機種をご連絡ください。\n2. ヘッドの裏の写真(タイヤ部分がわかる写真)を出品下さい。(価格は9,999,999としてください。)\nこちらで取り付けできると判断した場合のみ、保証が適応されます。保証が適応された場合、万が一取り付けできなかったら返金いたしますので、その旨ご連絡ください。\
   \n\n【交換頻度】\n１年以内の交換をおすすめしています。\n【タイヤのサイズ】\n外径: 10 mm\n内径: 3 mm\n幅: 4.5 mm (予備タイヤは3.5 mm)\nシャフトにシールを巻くことを前提にしています。詳しくは上記ブログ（タイヤの取り付け方）をご覧下さい。\
   \n\n【現状のメーカー修理】\nメーカーで修理依頼すると、購入から2年以上経過すると最低5000円の工賃が発生します。\
   \n\n【最後に】\nタイヤは簡単に交換できるので、DIYで行うことをおすすめしています。(機種によって右後輪の取り外しが困難な場合があります)\nメーカーから部品だけの購入が不可と言われたため、自分で作成したものを出品しています。\n最近、ご購入者様からお褒めの言葉を多く頂いております。モチベーションの向上に繋がりますので、評価の方にもコメントいただけたら幸いです。")
#ダイソン共通  rctd: 
def rctd():
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
sleep(1)
driver.get("https://fril.jp/sell")
sleep(1)
ritmlists = driver.find_elements_by_class_name("media-heading")
sleep(1)
print(4,datetime.datetime.now())
rlists=[]
sleep(1)
for ritmlist in ritmlists:
   rlists.append(ritmlist.text)
sleep(1)
if not "ダイソン掃除機 タイヤ5個+テフロンテープ+トルクスドライバー3本セット" in rlists:
   driver.get("https://fril.jp/item/new")
   sleep(1)
   driver.find_elements_by_xpath("//input[@type='file']")[0].send_keys("/Users/masa/Pictures/torx/torxset.JPG")
   driver.find_elements_by_xpath("//input[@type='file']")[1].send_keys("/Users/masa/Pictures/dyson/tire5.jpg")
   driver.find_elements_by_xpath("//input[@type='file']")[2].send_keys("/Users/masa/Pictures/dyson/tape.jpeg")
   driver.find_elements_by_xpath("//input[@type='file']")[3].send_keys("/Users/masa/Pictures/torx/torxdriver.jpeg")
   rinputElement1 = driver.find_elements_by_class_name("form-group")[1].find_element_by_css_selector("input")
   rinputElement1.send_keys("ダイソン掃除機 タイヤ5個+テフロンテープ+トルクスドライバー3本セット")
   rexp1a = "【商品紹介】\n画像の2～4枚目が商品になります。ダイソンモーターヘッドの裏側に付いているタイヤと工具です。"
   rexp1()
   rctd()
   rinputElement8 = driver.find_elements_by_class_name("form-group")[11].find_element_by_css_selector("input")
   rinputElement8.send_keys("650")
   rcateElement9 = driver.find_element_by_class_name("col-md-offset-4.col-md-5.col-xs-12")
   rcateElement9.click()
   sleep(2)
   rcateElement10 = driver.find_element_by_class_name("col-md-8.col-xs-8")
   rcateElement10.click()
   sleep(5)
if not "ダイソン掃除機 タイヤ5個+テフロンテープ+シャフト4本セット" in rlists:
   driver.get("https://fril.jp/item/new")
   sleep(1)
   driver.find_elements_by_xpath("//input[@type='file']")[0].send_keys("/Users/masa/Pictures/dyson/shaftset.JPG")
   driver.find_elements_by_xpath("//input[@type='file']")[1].send_keys("/Users/masa/Pictures/dyson/tire5.jpg")
   driver.find_elements_by_xpath("//input[@type='file']")[2].send_keys("/Users/masa/Pictures/dyson/tape.jpeg")
   driver.find_elements_by_xpath("//input[@type='file']")[3].send_keys("/Users/masa/Pictures/dyson/shaft.jpeg")
   rinputElement1 = driver.find_elements_by_class_name("form-group")[1].find_element_by_css_selector("input")
   rinputElement1.send_keys("ダイソン掃除機 タイヤ5個+テフロンテープ+シャフト4本セット")
   rexp1a = "【商品紹介】\n画像の2～4枚目が商品になります。ダイソンモーターヘッドの裏側に付いているタイヤと部品です。"
   rexp1()
   rctd()
   rinputElement8 = driver.find_elements_by_class_name("form-group")[11].find_element_by_css_selector("input")
   rinputElement8.send_keys("650")
   rcateElement9 = driver.find_element_by_class_name("col-md-offset-4.col-md-5.col-xs-12")
   rcateElement9.click()
   sleep(2)
   rcateElement10 = driver.find_element_by_class_name("col-md-8.col-xs-8")
   rcateElement10.click()
   sleep(5)
if not "ダイソン掃除機 タイヤ5個+テフロンテープ+シャフト4本+トルクスドライバー3本" in rlists:
   driver.get("https://fril.jp/item/new")
   sleep(1)
   driver.find_elements_by_xpath("//input[@type='file']")[0].send_keys("/Users/masa/Pictures/dyson/4set.JPG")
   driver.find_elements_by_xpath("//input[@type='file']")[1].send_keys("/Users/masa/Pictures/dyson/tire5set.JPG")
   driver.find_elements_by_xpath("//input[@type='file']")[2].send_keys("/Users/masa/Pictures/dyson/shaft.jpeg")
   driver.find_elements_by_xpath("//input[@type='file']")[3].send_keys("/Users/masa/Pictures/torx/torxdriver.jpeg")
   rinputElement1 = driver.find_elements_by_class_name("form-group")[1].find_element_by_css_selector("input")
   rinputElement1.send_keys("ダイソン掃除機 タイヤ5個+テフロンテープ+シャフト4本+トルクスドライバー3本")
   rexp1a = "【商品紹介】\n画像の2～4枚目が商品になります。ダイソンモーターヘッドの裏側に付いているタイヤと部品と工具です。"
   rexp1()
   rctd()
   rinputElement8 = driver.find_elements_by_class_name("form-group")[11].find_element_by_css_selector("input")
   rinputElement8.send_keys("830")
   rcateElement9 = driver.find_element_by_class_name("col-md-offset-4.col-md-5.col-xs-12")
   rcateElement9.click()
   sleep(2)
   rcateElement10 = driver.find_element_by_class_name("col-md-8.col-xs-8")
   rcateElement10.click()
   sleep(5)
if not "ダイソン掃除機 タイヤ5個+テフロンテープセット" in rlists:
   driver.get("https://fril.jp/item/new")
   sleep(1)
   driver.find_elements_by_xpath("//input[@type='file']")[0].send_keys("/Users/masa/Pictures/dyson/tire5set.jpg")
   driver.find_elements_by_xpath("//input[@type='file']")[1].send_keys("/Users/masa/Pictures/dyson/tire5.jpg")
   driver.find_elements_by_xpath("//input[@type='file']")[2].send_keys("/Users/masa/Pictures/dyson/tape.jpeg")
   driver.find_elements_by_xpath("//input[@type='file']")[3].send_keys("/Users/masa/Pictures/dyson/morterheadback.jpeg")
   rinputElement1 = driver.find_elements_by_class_name("form-group")[1].find_element_by_css_selector("input")
   rinputElement1.send_keys("ダイソン掃除機 タイヤ5個+テフロンテープセット")
   rexp1a = "【商品紹介】\n画像の2～3枚目が商品になります。ダイソンモーターヘッドの裏側に付いているタイヤです。"
   rexp1()
   rctd()
   rinputElement8 = driver.find_elements_by_class_name("form-group")[11].find_element_by_css_selector("input")
   rinputElement8.send_keys("580")
   rcateElement9 = driver.find_element_by_class_name("col-md-offset-4.col-md-5.col-xs-12")
   rcateElement9.click()
   sleep(2)
   rcateElement10 = driver.find_element_by_class_name("col-md-8.col-xs-8")
   rcateElement10.click()
   sleep(5)
if not "ダイソン掃除機 タイヤ2個+テフロンテープセット" in rlists:
   driver.get("https://fril.jp/item/new")
   sleep(1)
   driver.find_elements_by_xpath("//input[@type='file']")[0].send_keys("/Users/masa/Pictures/dyson/tire2set.JPG")
   driver.find_elements_by_xpath("//input[@type='file']")[1].send_keys("/Users/masa/Pictures/dyson/tire2.jpeg")
   driver.find_elements_by_xpath("//input[@type='file']")[2].send_keys("/Users/masa/Pictures/dyson/tape.jpeg")
   driver.find_elements_by_xpath("//input[@type='file']")[3].send_keys("/Users/masa/Pictures/dyson/morterheadback.jpeg")
   rinputElement1 = driver.find_elements_by_class_name("form-group")[1].find_element_by_css_selector("input")
   rinputElement1.send_keys("ダイソン掃除機 タイヤ2個+テフロンテープセット")
   rexp1a = "【商品紹介】\n画像の2～3枚目が商品になります。ダイソンモーターヘッドの裏側に付いているタイヤです。"
   rexp1()
   rctd()
   rinputElement8 = driver.find_elements_by_class_name("form-group")[11].find_element_by_css_selector("input")
   rinputElement8.send_keys("380")
   rcateElement9 = driver.find_element_by_class_name("col-md-offset-4.col-md-5.col-xs-12")
   rcateElement9.click()
   sleep(2)
   rcateElement10 = driver.find_element_by_class_name("col-md-8.col-xs-8")
   rcateElement10.click()
   sleep(5)
#全てのウインドウを閉じる
sleep(4)
driver.quit()

# Open web (account 2)
options = webdriver.ChromeOptions()
options.add_argument(
   '--user-data-dir={chrom_dir_path}'.format(chrom_dir_path = '/Users/masa/Library/Application Support/Google/Chrome/Profile 2'))
driver = webdriver.Chrome(options=options, executable_path='/usr/local/bin/chromedriver')

# メルカリのページにアクセス
print(5,datetime.datetime.now())
sleep(1)
driver.get("https://www.mercari.com/jp/mypage/listings/listing/")
sleep(1)
# 出品中の商品リストを作成
itmlists = driver.find_elements_by_class_name("mypage-item-text")
lists = []
for itmlist in itmlists:
   lists.append(itmlist.text)
exp1a = "【商品紹介】\n画像の2～4枚目が商品になります。ダイソンモーターヘッドの裏側に付いているタイヤと工具です。"
itmpage()
#  1.画像をアップロード & 商品名 & 販売価格, 2.カテゴリ & ブランド & 商品説明, 3. 商品の状態 & 配送料の負担 & 配送の方法 & 発送元の地域 & 発送までの日数, 4.出品する
if not "ダイソン掃除機 タイヤ5個+テフロンテープ+トルクスドライバー3本セット" in lists:
   itm1()
   ctd(exp1)
   mc()
   exh()
print(6,datetime.datetime.now())
exp1a = "【商品紹介】\n画像の2～4枚目が商品になります。ダイソンモーターヘッドの裏側に付いているタイヤと部品です。"
if not "ダイソン掃除機 タイヤ5個+テフロンテープ+シャフト4本セット" in lists:
   itm2()
   ctd(exp1)
   mc()
   exh()
exp1a = "【商品紹介】\n画像の2～5枚目が商品になります。ダイソンモーターヘッドの裏側に付いているタイヤと部品と工具です。"
if not "ダイソン掃除機 タイヤ5個+テフロンテープ+シャフト4本+トルクスドライバー3本" in lists:
   itm3()
   ctd(exp1)
   mc()
   exh()
exp1a = "【商品紹介】\n画像の2～3枚目が商品になります。ダイソンモーターヘッドの裏側に付いているタイヤです。"
if not "ダイソン掃除機 タイヤ5個+テフロンテープセット" in lists:
   itm4()
   ctd(exp1)
   mc()
   exh()
print(7,datetime.datetime.now())
exp1a = "【商品紹介】\n画像の2枚目が商品になります。ダイソンモーターヘッドの裏側に付いているタイヤです。"
if not "ダイソン掃除機 タイヤ5個" in lists:
   itm5()
   ctd(exp1)
   mc()
   exh()
exp1a = "【商品紹介】\n画像の2～3枚目が商品になります。ダイソンモーターヘッドの裏側に付いているタイヤです。"
if not "ダイソン掃除機 タイヤ2個+テフロンテープセット" in lists:
   itm6()
   ctd(exp1)
   mc()
   exh()
exp1a = "【商品紹介】\n画像の2枚目が商品になります。ダイソンモーターヘッドの裏側に付いているタイヤです。"
if not "ダイソン掃除機 タイヤ2個" in lists:
   itm7()
   ctd(exp1)
   mc()
   exh()
print(8,datetime.datetime.now())
# 1.出品ページへ, 2.画像をアップロード & 商品名 & 商品説明 & 販売価格, 3-11.カテゴリー, 12-13ブランド, 14.商品の状態 & 配送料の負担 & 配送の方法 & 発送元の地域 & 発送までの日数, 15.出品する
if not "トルクスドライバー3本セット（T10 & T8 & T6）" in lists:
   itm8()
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
   mc()
   exh()
#全てのウインドウを閉じる
print(9,datetime.datetime.now())
sleep(3)
driver.quit()