import os
import pyautogui
import signal
import sys
import time
import tkinter as tk
from tkinter import messagebox
import datetime

# デスクトップのパスを取得
desktop_path = os.path.expanduser("~/Desktop")

# pyautoguiの動作速度を設定
pyautogui.PAUSE = 0.5  # 秒数を調整してください（例: 0.5秒）

def calculate_date():
    """期中、直近１ヶ月期間を算出"""

    today = datetime.date.today()
    year = today.year
    month = today.month
    day = today.day

    first_day_of_previous_month = today.replace(day=1) - datetime.timedelta(days=1)
    if today.month == 1:
        first_day_of_previous_month = today.replace(year=today.year - 1, month=12, day=1)
        last_day_of_previous_month = first_day_of_previous_month.replace(month = 12, day=31)
    else:
        last_day_of_previous_month = today.replace(day=1) - datetime.timedelta(days=1)

    y1 = first_day_of_previous_month.year
    m1 = first_day_of_previous_month.month
    d1 = 1

    y2 = last_day_of_previous_month.year
    m2 = last_day_of_previous_month.month
    d2 = last_day_of_previous_month.day

    current_year = today.year
    current_month = today.month

    if current_month <= 8:
        y1_mid = current_year - 1
        m1_mid = 8
        d1_mid = 1

        last_day_of_previous_month_mid = today.replace(day=1) - datetime.timedelta(days=1)
        y2_mid = last_day_of_previous_month_mid.year
        m2_mid = last_day_of_previous_month_mid.month
        d2_mid = last_day_of_previous_month_mid.day
    else:
        y1_mid = current_year
        m1_mid = 8
        d1_mid = 1

        last_day_of_previous_month_mid = today.replace(day=1) - datetime.timedelta(days=1)
        y2_mid = last_day_of_previous_month_mid.year
        m2_mid = last_day_of_previous_month_mid.month
        d2_mid = last_day_of_previous_month_mid.day

    print("直近月報")
    print(f"y1 = {y1}, m1 = {m1}, d1 = {d1}")
    print(f"y2 = {y2}, m2 = {m2}, d2 = {d2}")
    print("\n期中")
    print(f"y1 = {y1_mid}, m1 = {m1_mid}, d1 = {d1_mid}")
    print(f"y2 = {y2_mid}, m2 = {m2_mid}, d2 = {d2_mid}")

    # 出力
    time.sleep(2)
    print("\n期中データ")
    time.sleep(1)
    pyautogui.typewrite(str(y1_mid))
    time.sleep(1)
    pyautogui.press('right')
    pyautogui.typewrite(str(m1_mid))
    time.sleep(1)
    pyautogui.press('right')
    pyautogui.typewrite(str(d1_mid))
    time.sleep(1)
    pyautogui.press('right')

    pyautogui.typewrite(str(y2_mid))
    time.sleep(1)
    pyautogui.press('right')
    pyautogui.typewrite(str(m2_mid))
    time.sleep(1)
    pyautogui.press('right')
    pyautogui.typewrite(str(d2_mid))
    time.sleep(1)
    pyautogui.press('right')
    pyautogui.press('down')
    pyautogui.press('right')
    pyautogui.press('enter')
    time.sleep(1)

    path = r"\\\\192.168.1.240\\中部産商(共有)\\02裕章\\01_tokuisaki.xlsx"

    print(path)

def user_geppo():
    # 得意先月報自由設計から選択
    in_a = '24'
    in_b = '13'

    pyautogui.click(100, 100)
    print("click")

    time.sleep(1)
    print("delete")
    pyautogui.press('delete')
    time.sleep(1)
    print("24")
    pyautogui.typewrite(str(in_a))
    pyautogui.press('enter')
    time.sleep(1)
    pyautogui.typewrite(str(in_b))
    pyautogui.press('enter')

    time.sleep(2)
    pyautogui.keyDown('shift')
    pyautogui.press('tab')
    pyautogui.keyUp('shift')

    time.sleep(1)
    pyautogui.press('space')
    time.sleep(1)
    for i in range(3):
        pyautogui.press('up')
    time.sleep(2)
    pyautogui.press('down')
    time.sleep(1)
    pyautogui.press('enter')
    time.sleep(1)

    print("得意先月報が選択されました")
    pyautogui.press('f2')
    time.sleep(3)
    pyautogui.press('down')
    calculate_date()

def zaiko():
    # zaiko data
    print("zaiko data")

    pyautogui.press('delete')
    pyautogui.typewrite('31')
    time.sleep(1)
    pyautogui.typewrite('31')
    pyautogui.press('enter')


def copy_to_clipboard(text):
    """
    テキストをクリップボードにコピーする関数
    """
    command = 'echo ' + text.strip() + '| clip'
    os.system(command)

def run_mstsc():
    """
    mstsc.exeを実行する関数
    """
    mstsc_path = os.path.join(os.environ['windir'], 'system32', 'mstsc.exe')
    os.system(mstsc_path)
    time.sleep(4)
    pyautogui.press('enter')
    time.sleep(2)
    pyautogui.press('left')
    pyautogui.press('enter')
    time.sleep(7)


    print("Maxmum Windows & Win+R")
    pyautogui.hotkey('win', 'up')
    pyautogui.hotkey('win', 'r')
    time.sleep(1)

    # 'C:\\Program Files (x86)\\OHKEN\\HBWIN\\BIN\\HSERP.EXE' をクリップボードにコピー
    copy_to_clipboard('C:\\Program Files (x86)\\OHKEN\\HBWIN\\BIN\\HSERP.EXE')
    time.sleep(1)
    pyautogui.keyDown('ctrl')
    pyautogui.press('v')
    pyautogui.keyUp('ctrl')

    time.sleep(1)
    pyautogui.press('enter')

    time.sleep(2)
    pyautogui.press('enter')
    time.sleep(1)
    pyautogui.press('enter')

    # 大臣起動しましたか？
    print("大臣起動しましたか？")
    root = tk.Tk()
    root.withdraw()
    answer = messagebox.askquestion("確認", "販売大臣が起動しましたか？")
    if answer == "yes":
        # データ操作　
        # 販売大臣　→　得意先月報、商品月報
        # 旬報、前期会議資料コピー
        print("user_geppo!!")
        user_geppo()

    else:
        messagebox.showinfo("終了", "プログラムを終了します")
        sys.exit(0)

def signal_handler(signal, frame):
    """
    Ctrl+Cが押された場合のシグナルハンドラ関数
    """
    answer = input("Are you sure you want to exit? (y/n): ")
    if answer.lower() == "y":
        print("Exiting program...")
        sys.exit(0)

def exit_program():
    """
    プログラムを終了する関数
    """
    print("Exiting program...")
    sys.exit(0)

def main():
    """
    メインの実行関数
    """
    root = tk.Tk()
    root.withdraw()
    answer = messagebox.askquestion("確認", "販売大臣からデータを抽出しますか？")
    if answer == "yes":
        print("Waiting for mstsc.exe to start...")
        run_mstsc()
        print("OK!!!")

    else:
        messagebox.showinfo("終了", "プログラムを終了します")
        sys.exit(0)

    # Ctrl+Cを押してプログラムを終了させるための無限ループ
    while True:
        try:
            # 画面の左上隅にマウスカーソルが来た場合にプログラムを終了する
            current_mouse_pos = pyautogui.position()
            if current_mouse_pos == (0, 0):
                exit_program()
        except pyautogui.FailSafeException:
            pass

if __name__ == "__main__":
    # Ctrl+Cのシグナルハンドラを設定
    signal.signal(signal.SIGINT, signal_handler)

    # メインの実行関数を呼び出す
    print("Starting program...")
    main()
