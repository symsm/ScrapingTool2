from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.action_chains import ActionChains
import os
import tqdm
import winsound
import pyautogui
import datetime

#ドライバー設定
options = Options()
options.add_argument('--disable-gpu')
options.add_argument('--disable-extensions')
options.add_argument('--proxy-server="direct://"')
options.add_argument('--proxy-bypass-list=*')
options.add_argument('--start-maximized')

#画面スクロールの際の調整
ELEMENT_NUM = 44
#記録ファイルパス（前回で何ページまで読み込んだか記録してある）
idxDir = './indexFiles'
idxFile = ''
history = 1 #記録ファイルの内容をここに入れる。初期値は１

try:
    #対象県コード入力
    print('\nスクレイピング対象の県コードを入力してください')
    PC = input()
    if not PC.isdecimal():
        print('数値で入力してください')
        input()
        os._exit(0)
    
    idxFile = idxDir + '/indexFile_' + str(PC) + '.txt'

    #ディレクトリチェック／記録ファイルチェックと読み込み
    if not os.path.exists(idxDir):
      print('「indexFiles」フォルダを作成してください')
      os._exit(0)
    elif os.path.exists(idxFile):
      fp = open(idxFile, 'r', encoding='utf-8')
      history = fp.readlines()
      if len(history[len(history) - 1].split(':')) == 2:
        histroy = int(history[len(history) - 1].split(':')[0])
      elif len(str(history[len(history) - 1]).split(':')) == 3:
        print('スクレイピングは完了しています。やり直す場合は「indexFiles」内の該当する記録ファイルを削除してください。')
        os._exit(0)
      else:
        print('無効な記録ファイルです。「indexFiles」フォルダ内を確認してください。')
        os._exit(0)
    
    #処理開始メッセージ
    print("\n情報を読み込んでいます…")

    #ドライバーのパス
    DRIVER_PATH = 'chromedriver.exe'

    # ブラウザの起動
    driver = webdriver.Chrome(executable_path=DRIVER_PATH, chrome_options=options)
    url = 'https://etsuran.mlit.go.jp/TAKKEN/sokatsuKensaku.do'
    driver.get(url)

    #本店指定
    selector = '#choice'
    element = driver.find_element_by_css_selector(selector)
    element.send_keys(Keys.DOWN)

    #県コード指定
    selector = '#kenCode'
    element = driver.find_element_by_css_selector(selector)
    element.send_keys(PC)

    #ソート項目を指定
    selector = '#sortValue'
    element = driver.find_element_by_css_selector(selector)
    element.send_keys(Keys.DOWN, Keys.DOWN)

    #ソートを降順に指定
    selector = '#rdoSelectSort2'
    element = driver.find_element_by_css_selector(selector)
    element.click()
    
    #表示件数を５０件に指定
    selector = '#dispCount'
    element = driver.find_element_by_css_selector(selector)
    element.send_keys('50', Keys.ENTER)

    #スクレイピング開始ページ
    if history == 1:
      START_PAGE = 1
    else: 
      START_PAGE = histroy + 1

    #最大ページ数取得
    selector = '#pageListNo1'
    element = driver.find_element_by_css_selector(selector)
    MAX_PAGE = int(str(element.text).split('\n')[1].split('/')[1])
    #一回で取得できる上限は30000件。ページ数で600ページ。
    if MAX_PAGE - START_PAGE >= 600:
      END_PAGE = START_PAGE + 599
    else:
      END_PAGE = MAX_PAGE

    #開始ページまで移動
    selector = '#pageListNo1'
    element = driver.find_element_by_css_selector(selector)
    select = Select(element)
    select.select_by_index(START_PAGE - 1)

    #ファイルオープン
    filePath = 'result_' + PC + '.csv'
    fp = open(filePath, 'a', encoding='utf-8')
    #fp.write('会社名,電話番号,住所\n')

    #スクレイピング
    for i in tqdm.tqdm(range(START_PAGE, END_PAGE + 1)):
        scrlHeight = 0  #スクロールする高さ
        crrScrl = 0  #スクロールした高さ
        #個別の情報を取得
        for i in range(3, 53):
          #スクロールする
          driver.execute_script("window.scrollTo(0, " + str(scrlHeight) + ");")
          #要素を指定
          #element = driver.find_element_by_xpath('//tbody/tr[' + str(i) + ']/td[3]')
          #mode = element.text #「許可等行政庁」が空かどうかで個別ページのレイアウトが違う
          element = driver.find_element_by_xpath('//tbody/tr[' + str(i) + ']/td[5]')
          #要素の座標を取得
          loc = element.location
          y_relative_coord = loc['y']
          browser_navigation_panel_height = driver.execute_script('return window.outerHeight - window.innerHeight;')
          y_absolute_coord = y_relative_coord + browser_navigation_panel_height - scrlHeight
          x_absolute_coord = loc['x']
          #次回スクロールする高さを決める
          if scrlHeight == 0:
            scrlHeight += element.rect['height']
          elif scrlHeight < element.rect['height'] * ELEMENT_NUM:
            scrlHeight += element.rect['height']
          #要素のところまでマウスを移動してクリック
          pyautogui.moveTo(x_absolute_coord + element.rect['height']/4, y_absolute_coord + element.rect['height']/2)  #要素のtop-leftをクリックしても反応しないので+10して補正
          pyautogui.click()
          
          #個別ページのレイアウトを調べる
          selector = '.re_summ'
          element = driver.find_element_by_css_selector(selector)
          mode = '許可番号' in element.text

          #情報をファイルに書き込み
          if mode:  #「許可等行政庁」が空かどうかで個別ページのレイアウトが違う
            #情報を取得
            element = driver.find_element_by_xpath("//table[@class='re_summ']/tbody/tr[2]/td")
            fp.write(str(element.text).split('\n')[1] + ',')  #会社名
            element = driver.find_element_by_xpath("//table[@class='re_summ']/tbody/tr[5]/td")
            fp.write(str(element.text) + ',')  #電話番号
            element = driver.find_element_by_xpath("//table[@class='re_summ']/tbody/tr[4]/td")
            fp.write(str(element.text).split('\n')[0] + ',')  #郵便番号
            fp.write(str(element.text).split('\n')[1] + str(element.text).split('\n')[2] + '\n')  #住所
          else:
            element = driver.find_element_by_xpath("//table[@class='re_summ6']/tbody/tr[1]/td")
            fp.write(str(element.text).split('\n')[1] + ',')  #会社名
            element = driver.find_element_by_xpath("//table[@class='re_stan']/tbody/tr/td")
            fp.write(str(element.text) + ',')  #電話番号
            element = driver.find_element_by_xpath("//table[@class='re_summ6']/tbody/tr[3]/td")
            fp.write(',') #郵便番号（空になる）
            fp.write(str(element.text) + '\n')  #住所

          #ブラウザバック
          driver.back()

        #次ページへ
        selector = '.result :nth-child(6)'
        element = driver.find_element_by_xpath("//img[contains(@src,'/TAKKEN/images/result_move_r.jpg')]")
        element.click()

    fp.close

    #記録ファイル
    fp = open(idxFile, 'a', encoding='utf-8')
    if MAX_PAGE == END_PAGE:
      fp.write('\n' + str(END_PAGE) + ':' + str(datetime.date.today) + ':done!')
    else:
      fp.write('\n' + str(END_PAGE) + ':' + str(datetime.date.today))

    #終了音
    winsound.Beep(1500, 100)
    winsound.Beep(2000, 100)
    winsound.Beep(2400, 100)
    winsound.Beep(3000, 100)

    print("\n\n/////////////////////////////////////////////////")
    print("///          スクレイピング処理完了!          ///")
    print("/////////////////////////////////////////////////")
    input('キーを押してください...')
    driver.close()
    driver.quit()
    os._exit(0)

except IOError as e:
  print("\n\nファイルに書き出せません！：")
  print('異常終了：' + str(e))
  input('キーを押してください...')
  driver.close()
  driver.quit()
  os._exit(1)
except Exception as e:
  print('異常終了：' + str(e))
  input('キーを押してください...')
  driver.close()
  driver.quit()
  os._exit(1)
