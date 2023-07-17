import pyautogui
import datetime
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager

#Use File Rename
import os,time
import shutil
from pathlib import Path
import zipfile

#Use from CSV to EXCEL and copy&paste
import os
import openpyxl as px
import csv
from openpyxl.utils import column_index_from_string

#複数のファイルの一括削除
import glob
def remove_glob(pathname, recursive=True):
    for p in glob.glob(pathname, recursive=recursive):
        if os.path.isfile(p):
            os.remove(p)

p_path = r'C:/Users/H-INOUE/'

#ここからfunc1-func4をとにかく動かしてDLするように操作する。
def func1():
    #Period(choose latest year )
    pyautogui.sleep(1)
    for i in range(3):
        pyautogui.press('tab')

    for i in range(5):
        pyautogui.press('up')

    #Search
    pyautogui.keyDown('shift')
    for i in range(5):
        pyautogui.press('tab')
    pyautogui.keyUp('shift')

    pyautogui.press('enter')

    #CSV DL
    pyautogui.sleep(3)
    #モニタにより調整（本当はもう少しスマートに操作させたい）
    pyautogui.click(x=70, y=460)

def func2(n):
    pyautogui.sleep(5)
    pyautogui.keyDown('alt')
    pyautogui.press('tab')
    pyautogui.keyUp('alt')

    print("[debug]alt+tab pressed...")

    pyautogui.sleep(1)
    pyautogui.click(x=40, y=375)

    pyautogui.sleep(2)
    for i in range(10):
        p
    pyautogui.press('tab')

    pyautogui.press('delete')
    pyautogui.typewrite(n)

    #Search
    pyautogui.press('enter')

    #CSV DL
    pyautogui.sleep(3)
    #モニタにより調整（本当はスマートに操作させたい）
    pyautogui.click(x=70, y=460)

def func3(country_num,n):
    #TAB+ALT
    pyautogui.sleep(5)
    pyautogui.keyDown('alt')
    pyautogui.press('tab')
    pyautogui.keyUp('alt')

    pyautogui.sleep(1)
    pyautogui.click(x=40, y=375)

    #NumberINPUT
    pyautogui.sleep(2)
    for i in range(6):
        pyautogui.press('tab')
    pyautogui.typewrite(country_num)

    for i in range(4):
        pyautogui.press('tab')

    #Type Keyboard
    pyautogui.typewrite(n)

    #Search
    pyautogui.press('enter')

    #CSV DL
    pyautogui.sleep(3)
    pyautogui.click(x=70, y=460)

def func4(n):
    #TAB+ALT
    pyautogui.sleep(5)
    pyautogui.keyDown('alt')
    pyautogui.press('tab')
    pyautogui.keyUp('alt')

    pyautogui.sleep(1)
    pyautogui.click(x=40, y=375)

    #NumberINPUT
    #Select All Country
    pyautogui.sleep(2)
    for i in range(5):
        pyautogui.press('tab')
    pyautogui.press('up')
    pyautogui.press('tab')
    pyautogui.press('tab')

    #Type Keyboard
    pyautogui.press('delete')
    pyautogui.typewrite(n)

    #Search
    pyautogui.press('enter')

    #CSV DL
    pyautogui.sleep(3)
    pyautogui.click(x=70, y=460)

def dl():
    #run chrome
    driver = webdriver.Chrome(ChromeDriverManager().install())

    driver.get('https://www.customs.go.jp/toukei/srch/index.htm?M=77&P=0')
    driver.maximize_window()

    #ここから目的に合わせて今回はfunc1-func4で動作させる。以下はfunc2の例
    #適宜データ用意し、同様に処理
    data = ['293212000','293213000','',...,'']
    for index,item in enumerate(data):
        print(index,item)
        print("Call func2")
        func2(item)

    #DLOPEN
    pyautogui.sleep(5)
    pyautogui.click(x=100, y=1050)

    driver.close()

if __name__ == "__main__":
    try:

        remove_glob(p_path+'Downloads/*.zip')
        x,y = pyautogui.position()

        #dl funcへ
        dl()

        #DL file save the tempfolder and rename
        today = datetime.date.today() #今日の日付を取得
        Y = str(today.year)
        M = str(today.month)
        D = str(today.day)

        #Path 形式でDownloads以下のzipファイルを l_1にlistで格納
        p = Path(p_path+"Downloads/")
        l_1 = list(p.glob('*.zip'))
        re_l = ['fl.csv','fa.csv']

        Temp_P = Path(p_path+"DESKTOP/temp_python")

        #list 要素数から、zipfile.ZipFileリストをmyzipに格納
        for i in range(0,len(l_1)):
            with zipfile.ZipFile(str(l_1[i])) as myzip:

                #Temp_P フォルダへ取り出し
                myzip.extractall(Temp_P)

                #ソートして、temp_pythonフォルダの中のファイルを取得
                files = sorted(Temp_P.glob('*.csv'))
        i = 0
        #ディレクトリ無い場合、新規作成
        NEW_P = Path(p_path+"DESKTOP/temp_python/RENAME(PY_AUTO)/")

        #格納先フォルダ not available -> mkdir
        if not os.path.exists(NEW_P):
            os.mkdir(NEW_P)

        #available -> 一度全て空にして、新しく作成
        else:
            shutil.rmtree(NEW_P)
            os.mkdir(NEW_P)

        #フォルダ格納
        #for file in files -> files の中にあるファイル全てに対して実行 ->ここでerrror ->num of files check!!!
        for file in files:
            #隠しファイル処理実行せず
            if file.name.startswith('.'):
                continue

            #old Path CHECK
            oldFilePath = file.__str__()
            print("oldFilePath",oldFilePath)

            #new Path CHECK
            newFileName = re_l[i]
            print("newFileName -> ",re_l[i],newFileName)

            newFilePath = str(NEW_P) + '/' + newFileName
            print("newFilePath -> ",newFilePath)

            os.rename(oldFilePath,newFilePath)
            i = i+1

        #CSV to EXCEL and copy & paste

        with open (p_path+'Desktop/temp_python/RENAME(PY_AUTO)/.csv', newline='') as tempfile:
            data_list_fl = list(csv.reader(tempfile, delimiter=','))
            wb = px.Workbook()
            ws = wb.active
            ws.append([]) #A1,B1 空行

        #store the data
        for record in data_list_fl:
            ws.append(record)

        #store the type -> xlxs
        wb.save('.xlsx')

        #copy temp sheet data to output
        temp = '.xlsx'
        output = p_path+'Desktop/test.xlsx'
        #How to open wb and Indicate the ws
        #別のブックへセルのコピー

        wb_temp = px.load_workbook(filename=temp)
        ws_temp = wb_temp['Sheet']

    except KeyboardInterrupt:
        print('\n終了')


