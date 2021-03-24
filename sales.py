import gspread
import json
#ServiceAccountCredentials：Googleの各サービスへアクセスできるservice変数を生成します。
from oauth2client.service_account import ServiceAccountCredentials 
#2つのAPIを記述しないとリフレッシュトークンを3600秒毎に発行し続けなければならない
scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
#認証情報設定 ダウンロードしたjsonファイル名をクレデンシャル変数に設定（秘密鍵、Pythonファイルから読み込みしやすい位置に置く）
credentials = ServiceAccountCredentials.from_json_keyfile_name('/Users/masa/dys-mercari/auto-exhibition/mercari-284420-8a48bc1a3e9b.json', scope)
gc = gspread.authorize(credentials) #OAuth2の資格情報を使用してGoogle APIにログインします。
SPREADSHEET_KEY = '1EPa-1zsfxXissFWV3BigpHsJCyBuiQHnvfucSWmv1FE' #共有設定したスプレッドシートキーを変数[SPREADSHEET_KEY]に格納する。
worksheet = gc.open_by_key(SPREADSHEET_KEY).worksheet('sales') #共有設定したスプレッドシートのシート1を開く
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager

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

def mercari():
    driver.get("https://www.mercari.com/jp/mypage/listings/in_progress/") #取引中ページに移動
    sleep(3)
    slditems = driver.find_elements_by_class_name("js-mypage-item") #全ての取引中の商品名を取得
    i = 0
    for slditem in slditems:
        i += 1
        print(31, 'i', i)
    f = 0
    for d in range(0, i): #各商品の進捗状況を取得
        values_list = worksheet.get_all_values() #worksheetの全てのセルの情報を取得
        print('len(values_list)', len(values_list), 'd-f', d-f)
        slditem = driver.find_elements_by_class_name("js-mypage-item")[d-f] #全ての取引中の商品名を取得
        prog = slditem.find_element_by_class_name("mypage-item-status.action-required") #取引の進捗状況を取得（発送待ち、受取評価待ちなど）
        print('prog.text', prog.text)
        if prog.text == "受取評価待ち": #進捗状況が「受取評価待ち」の時に次に進む
            url = slditem.find_element_by_class_name("mypage-item-link.has-button").get_attribute("href") #該当のurlを取得
            driver.get(url) #該当のurlへ移動する
            sleep(3)
            sldinf = driver.find_element_by_class_name("transact-info-table") #商品名、利益、販売日付を取得する
            itm_id = sldinf.find_elements_by_class_name("transact-info-table-row")[5].find_elements_by_class_name("transact-info-table-cell")[1] #商品IDを取得する
            print('itm_id', itm_id.text)
            worksheet.update_cell(3, 6, itm_id.text) #ワークシートに記入する
            if worksheet.cell(2, 6).value == "0": #D列に同じIDがない場合
                itm_name = sldinf.find_elements_by_class_name("transact-info-table-row")[0].find_element_by_class_name("transact-info-item.bold") #商品名を取得する
                itm_name2 = itm_name.text
                print(49, itm_name2)
                itm_name3 = itm_name2.replace('\n','')
                print(51, itm_name3)
                itm_name4 = itm_name3.split('¥')[0]
                print(51, itm_name4)
                itm_pri = sldinf.find_elements_by_class_name("transact-info-table-row")[3].find_elements_by_class_name("transact-info-table-cell")[1] #利益を取得する
                itm_pri2 = re.sub("\\D", "", itm_pri.text)
                itm_dat = sldinf.find_elements_by_class_name("transact-info-table-row")[4].find_elements_by_class_name("transact-info-table-cell")[1] #販売日を取得する
                print('itm_dat', itm_dat)
                itm_dat2 = datetime.datetime.strftime(dt.now(), '%Y年')
                print('itm_dat2', itm_dat2)
                itm_dat3 = itm_dat2 + itm_dat.text
                print('itm_dat3', itm_dat3, type(itm_dat3))
                lastrow = len(values_list) + 1
                print(57, 'lastrow', lastrow) 
                worksheet.update_cell(lastrow, 1, itm_dat3) #ワークシートに記入する
                worksheet.update_cell(lastrow, 2, itm_name4)
                worksheet.update_cell(lastrow, 3, itm_pri2)
                worksheet.update_cell(lastrow, 4, itm_id.text)
                worksheet.update_cell(lastrow, 5, 'mer_m')
                driver.get("https://www.mercari.com/jp/mypage/listings/in_progress/") #取引中ページに移動する
                sleep(3)
        elif prog.text == "評価待ち":
            f -= 1
            print("評価, f", f)
            url = slditem.find_element_by_class_name("mypage-item-link.has-button").get_attribute("href") #該当のurlを取得
            driver.get(url) #該当のurlへ移動する
            sleep(3)
            icon = driver.find_element_by_class_name("transact-review-icon")
            icon.click()
            review = driver.find_element_by_class_name("textarea-default")
            review.send_keys("この度はお取引いただき誠にありがとうございました。また機会がございましたら、よろしくお願い致します。")
            btn = driver.find_element_by_class_name("btn-default.btn-red")
            btn.click()
        driver.get("https://www.mercari.com/jp/mypage/listings/in_progress/") #取引中ページに移動する
        sleep(3)

try:
    print('account_1', dt.now())
    options = webdriver.ChromeOptions()
    options.add_argument(
    '--user-data-dir={chrom_dir_path}'.format(chrom_dir_path = '/Users/masa/Library/Application Support/Google/Chrome/Profile 7'))
    options.add_argument('--headless')
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-gpu')
    driver = webdriver.Chrome(ChromeDriverManager().install(), options=options)
    print('mercari')
    mercari()
except Exception as e:
    print(e)
    print("error in mercari")

try:
    print('rakuma')
    driver.get("https://fril.jp/sell")
    driver.find_element_by_xpath("/html/body/div[2]/div/div/div/div/ul/li[2]/a").click() #取引中タブを選択
    sleep(1)
    rprogs = driver.find_elements_by_class_name("tab-pane.fade")[1].find_elements_by_class_name("media")
    i = 0
    for rprog in rprogs: #取引中の数を取得
        prog = rprog.find_element_by_class_name("media-body")
        i += 1
        print(99, 'i',i)
    for i in range(0, i): #受取確認待ちの情報をスプレッドシートに転記
        rprog = driver.find_elements_by_class_name("tab-pane.fade")[1].find_elements_by_class_name("media")[i]
        prog = rprog.find_element_by_class_name("media-body")
        print('prog', prog.text)
        if "受取確認待ち" in prog.text:
            values_list = worksheet.get_all_values() #worksheetの全てのセルの情報を取得
            url = driver.find_elements_by_class_name("tab-pane.fade")[1].find_elements_by_class_name("media")[i].find_element_by_tag_name("a").get_attribute("href") #進捗状況を取得
            driver.get(url)
            sleep(1)
            ele = driver.find_elements_by_class_name("contents-wrapper")[1]
            if not "取引に問題" in ele.text: #取引に問題があるとrowの数が変わるので除外
                order_id = ele.find_element_by_class_name("large-text")
                worksheet.update_cell(3, 6, order_id.text) #ワークシートに記入する
                if worksheet.cell(2, 6).value == "0": #D列に同じIDがない場合
                    name = ele.find_element_by_class_name("item_name")
                    price = ele.find_elements_by_class_name("row")[6].find_elements_by_class_name("col.s6.right-align")[2]
                    price2 = re.sub("\\D", "", price.text)
                    lastrow = len(values_list) + 1
                    date = ele.find_elements_by_class_name("row")[4].find_element_by_class_name("col.s12")
                    print('date', date)
                    date2 = re.sub("[^0-9]", "", date.text)
                    print('date2', date2)
                    date3 = datetime.datetime.strptime(date2, '%Y%m%d%H%M')
                    print('date3', date3)
                    date4 = date3.strftime('%Y年%-m月%-d日 %H:%M')
                    print('date4', date4, type(date4))
                    worksheet.update_cell(lastrow, 1, date4) #ワークシートに記入する
                    worksheet.update_cell(lastrow, 2, name.text)
                    worksheet.update_cell(lastrow, 3, price2)
                    print(133)
                    worksheet.update_cell(lastrow, 4, order_id.text)
                    worksheet.update_cell(lastrow, 5, 'rak_m')
        elif "評価" in prog.text:
            url = driver.find_elements_by_class_name("tab-pane.fade")[1].find_elements_by_class_name("media")[i].find_element_by_tag_name("a").get_attribute("href") #進捗状況を取得
            driver.get(url)
            sleep(3)
            review = driver.find_element_by_class_name("review-textarea")
            review.send_keys("この度はお取引いただき誠にありがとうございました。また機会がございましたら、よろしくお願い致します。")
            btn = driver.find_element_by_class_name("btn.waves-light.fril-pink.modal-trigger.btn-style01")
            btn.click()
            btn = driver.find_element_by_class_name("one-tap.btn-flat.fril-pink-text.modal-action")
            btn.click()
            sleep(3)
        driver.get("https://fril.jp/sell") #取引中ページに移動する
        sleep(3)
        driver.find_element_by_xpath("/html/body/div[2]/div/div/div/div/ul/li[2]/a").click() #取引中タブを選択
        sleep(1)
    print("ヤフオク")
    def yafuoku():
        driver.get("https://auctions.yahoo.co.jp/closeduser/jp/show/mystatus") # ヤフオクにアクセス
        sleep(1)
        itm_fin1 = driver.find_element_by_class_name("decMainLi.decLiLast")
        itm_fin1.click()
        sleep(1)
    yafuoku()
    itm_fins = driver.find_elements_by_class_name("auc_del_style")
    ylists=[]
    for itm_fin in itm_fins:
        print(159)
        values_list = worksheet.get_all_values() #worksheetの全てのセルの情報を取得
        lastrow = len(values_list) + 1
        itm_id = itm_fin.find_element_by_css_selector("input").get_attribute("value")
        print(itm_id)
        worksheet.update_cell(3, 6, itm_id)
        if worksheet.cell(2, 6).value == "0": #D列に同じIDがない場合
            itm_date = itm_fin.find_elements_by_css_selector("td")[4]
            worksheet.update_cell(lastrow, 1, itm_date.text)
            itm_name = itm_fin.find_element_by_css_selector("a")
            worksheet.update_cell(lastrow, 2, itm_name.text) 
            itm_price = itm_fin.find_elements_by_css_selector("td")[3]
            itm_price2 = re.sub("\\D", "", itm_price.text) # 文字(円)を削除
            itm_price3 = float(itm_price2) # 文字列から数値に変換
            itm_price4 = itm_price3 * 0.9 # 手数料を除く
            itm_price5 = int(itm_price4) # 整数に変換
            worksheet.update_cell(lastrow, 3, itm_price5)
            worksheet.update_cell(lastrow, 4, itm_id)
            worksheet.update_cell(lastrow, 5, 'yahoo_m')
            itm_fin1 = itm_fin.find_elements_by_css_selector("td")[10].find_element_by_css_selector("a").get_attribute("href")
            print("itm_fin1", itm_fin1)
            ylists.append(itm_fin1)
    print(ylists)
    i = 0
    for link in ylists:
        driver.get(ylists[i])
        try:
            driver.find_element_by_xpath("//*[text()=\"出品を続ける\"]").click()
        except Exception as e:
            print(e)
            print("出品を続けるボタンがない時")
        sleep(1)
        itm_fin3 = driver.find_element_by_class_name("Button.Button--proceed.Inquiry__button.Inquiry__buttonText")
        itm_fin3.click()
        itm_fin4 = driver.find_element_by_class_name("u-displayBlock.u-fontSize20.u-textNormal")
        itm_fin4.click()
        i += 1
    driver.quit()
    print('last', dt.now())
except Exception as e:
    print(e)
    print("include error in total")