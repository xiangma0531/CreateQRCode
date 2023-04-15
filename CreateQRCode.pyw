import PySimpleGUI as sg
from selenium import webdriver
from selenium.webdriver.common.by import By
import openpyxl
import time
import configparser as cp

# iniファイル読込
inifile = cp.SafeConfigParser()
inifile.read('settings.ini')

# right関数
def right(infolder):
  return infolder[len(infolder) - 1]

def createQRCode(driver_path, url, file_name, infolder):
  msg = ""
  # webdriverの作成
  driver = webdriver.Chrome(executable_path=driver_path)
  # 要素が見つからない場合は10秒待つように設定
  driver.implicitly_wait(10)
  # QRコードをブラウザに表示
  driver.get("https://quickchart.io/qr?text=" + url)
  # imgタグを取得
  img = driver.find_element(By.TAG_NAME, "img")
  # imgタグだけをスクリーンショット撮影
  img.screenshot(infolder + file_name + ".png")
  # 2秒待つ
  time.sleep(2)
  # ブラウザを閉じる
  driver.quit()
  msg = infolder + file_name + ".pngを作成しました"
  return msg

# 変数定義
title = "QRコード生成アプリ"
driver_path = inifile.get('DEFAULT', 'driver')
label1, label2, label3, label4 = "ドライバー保存場所", "URL", "ファイル名", "保存先"

# 実行用関数
def execute():
  driver_path = values["driver_path"]
  url = values["URL"]
  file_name = values["file_name"]
  infolder = values["infolder"]
  # 作成先フォルダの最終文字が「/」でない場合「/」を結合
  if right(infolder) != "/":
    infolder = infolder + '/'
  msg = createQRCode(driver_path, url, file_name, infolder)
  window["result"].update(msg)


# アプリレイアウト
layout = [
  # ドライバー指定
  [sg.Text(label1, size=(14,1)), sg.Input(driver_path, key="driver_path"), sg.FileBrowse("選択")],
  # 作成したいQRコードのURL
  [sg.Text(label2, size=(14,1)), sg.Input("", key="URL")],
  # 作成したいQRコードのファイル名
  [sg.Text(label3, size=(14,1)), sg.Input("", key="file_name")],
  # QRコード保存先のパス
  [sg.Text(label4, size=(14,1)), sg.Input(".", key="infolder"), sg.FolderBrowse("選択")],
  # 実行ボタン
  [sg.Button("実行", size=(20,1), pad=(5,15), bind_return_key=True)],
  # 実行結果出力
  [sg.Multiline(key="result", size=(60,10))]
]

# アプリ実行
window = sg.Window(title, layout, font=(None, 14))
while True:
  event, values = window.read()
  if event == None:
    break
  if event == "実行":
    execute()