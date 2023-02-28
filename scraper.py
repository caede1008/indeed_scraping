import os
import sys
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from services.process import Process
import time
from selenium.webdriver.common.by import By
import openpyxl
from tkinter import *
from tkinter import ttk
import requests
from PIL import Image
import glob

# GUI設定
root = Tk()
root.title('Indeedスクレイピングツール')
root.geometry("400x300+100+300")
frame = ttk.Frame(root, padding=16)

keywordlbl = ttk.Label(text='キーワード')
keywordlbl.place(x=20, y=30)

txtbox = ttk.Entry()
txtbox.configure(state='normal', width=35)
txtbox.place(x=110, y=30)

wkplacelbl = ttk.Label(text='勤務地')
wkplacelbl.place(x=20, y=60)

wkplacebox = ttk.Entry()
wkplacebox.configure(state='normal', width=35)
wkplacebox.place(x=110, y=60)

exclusionlbl = ttk.Label(text='除外ワード')
exclusionlbl.place(x=20, y=90)

exclusionbox = ttk.Entry()
exclusionbox.configure(state='normal', width=35)
exclusionbox.place(x=110, y=90)

scrapingbutton = ttk.Button(root, text="スクレイピング実行", command=lambda:main(), width=40)
scrapingbutton.pack(pady=130)

def main():
    BrowserPath=ResourcePath("./browser/chrome.exe") # ブラウザ
    DriverPath=ResourcePath("./driver/chromedriver.exe") # ウェブドライバ

    # ウェブドライバ設定
    options=Options()
    options.binary_location=BrowserPath
    # options.add_argument("--headless") # 動きを見たい場合はコメントアウトする。
    driver=webdriver.Chrome(DriverPath, options=options)

    # 変数宣言
    names = []
    companynames =[]
    contents = []

    # スクレイピング準備
    ProcessC=Process(driver)
    ProcessC.goPage()
    time.sleep(3)

    # キーワード設定
    keywordtxt = txtbox.get()
    if len(keywordtxt) != 0:
        kwsearchbox = driver.find_element(By.XPATH, "//*[@id='text-input-what']")
        kwsearchbox.send_keys(keywordtxt)

    # 勤務地
    wkplacetext = wkplacebox.get()
    if len(wkplacetext) != 0:
        wkp_searchbox = driver.find_element(By.XPATH, "//*[@id='text-input-where']")
        wkp_searchbox.send_keys(wkplacetext)

    # 除外ワード
    exclusiontxt = exclusionbox.get()
    if len(exclusiontxt) != 0:
        kwsearchbox = driver.find_element(By.XPATH, "//*[@id='text-input-what']")
        kwsearchbox.send_keys(" -" + exclusiontxt)

    # 検索ボタン押下
    searchbutton = driver.find_element(By.XPATH, "//*[@id='jobsearch']/button")
    searchbutton.click()
    time.sleep(5)

    # 各案件ボタン
    buttons = []
    buttons = driver.find_elements(By.CLASS_NAME, "css-1m4cuuf.e37uo190")

    # Cookie同意
    ckbutton = driver.find_element(By.CLASS_NAME, "gnav-CookiePrivacyNoticeButton")
    ckbutton.click()

    # スクレイピング実行
    wknames = driver.find_elements(By.CLASS_NAME, "jobTitle.css-1h4a4n5.eu4oa1w0")
    for wkname in wknames:
        names.append(wkname.text)

    wkcompanynames = driver.find_elements(By.CLASS_NAME, "companyName")
    for wkcompanyname in wkcompanynames:
        companynames.append(wkcompanyname.text)

    for button in buttons:
        button.click()
        time.sleep(2)
        wkcontents = driver.find_elements(By.CLASS_NAME, "jobsearch-JobComponent-embeddedBody")
        for wkcontent in wkcontents:
            if len(wkcontent.text)==0:
                contents.append('無し')
            else:
                contents.append(wkcontent.text)

    # エクセル準備
    path = './indeed_list.xlsx'
    wb = openpyxl.load_workbook(path)
    ws = wb["list"]

    # Excel入力
    rownumber = 2
    idx = 0
    for name in names:
        ws.cell(rownumber, 1).value = companynames[idx]
        ws.cell(rownumber, 2).value = name
        ws.cell(rownumber, 3).value = contents[idx]
        idx += 1
        rownumber += 1

    wb.save(path)
    wb.close()

    # クローズ処理
    time.sleep(10)
    driver.close()
    driver.quit()


def ResourcePath(relativePath):
    try:
        basePath=sys._MEIPASS
    except Exception:
        basePath=os.path.dirname(__file__)
    return os.path.join(basePath, relativePath)

#if __name__=="__main__":
    #main()


root.mainloop()
