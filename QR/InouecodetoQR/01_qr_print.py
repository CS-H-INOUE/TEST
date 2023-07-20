import os
import datetime
from PIL import Image
import tkinter as tk
from tkinter import messagebox
import win32api
import sys
import requests
import time
import pyautogui

def send_teams_notification(webhook_url, message):
    headers = {'Content-Type': 'application/json'}
    data = {'text': message}

    response = requests.post(webhook_url, json=data, headers=headers)

    if response.status_code == 200:
        print("Notification sent successfully.")
    else:
        print(f"Failed to send notification. Status code: {response.status_code}, Response: {response.text}")

def debug_print(message):
    # デバッグメッセージを表示
    print("[DEBUG]", message)

    # ログファイルにメッセージを追記
    with open(f"print_log_{datetime.date.today()}.txt", "a") as log_file:
        log_file.write("[DEBUG] " + message + "\n")

def print_and_rename_images(folder_name):
    debug_print("print_and_rename_images 関数の開始")

    folder_path = os.path.join(os.getcwd(), folder_name)
    debug_print(f"フォルダパス: {folder_path}")

    if not os.path.exists(folder_path):
        messagebox.showerror("エラー", "指定されたフォルダが存在しません。")
        sys.exit(0)
        return

    # フォルダ内のファイルを取得
    files = os.listdir(folder_path)
    debug_print(f"ファイル一覧: {files}")

    # 画像ファイルのみを選択
    image_files = [file for file in files if file.endswith(('.jpg', '.jpeg', '.png'))]
    debug_print(f"画像ファイル一覧: {image_files}")

    # 画像を印刷してファイル名を変更
    for image_file in image_files:
        image_path = os.path.join(folder_path, image_file)

        debug_print(f"印刷中の画像: {image_path}")

        try:
            # 画像を印刷
            print_image(image_path)

            # ファイル名に"p_"を追加
            new_image_file = "p_" + image_file
            new_image_path = os.path.join(folder_path, new_image_file)
            os.rename(image_path, new_image_path)
            print(f"Renamed file: {image_file} -> {new_image_file}")

        except win32api.error as e:
            debug_print(f"印刷エラー: {e}")
            messagebox.showerror("印刷エラー", "画像の印刷中にエラーが発生しました。")
            continue

    debug_print("print_and_rename_images 関数の終了")

def print_image(image_path):
    debug_print(f"print_image 関数の開始: {image_path}")

    # 画像を開く
    image = Image.open(image_path)
    try:

        #印刷は７０点。どうにかサイズを変えたい＋



        # 画像を印刷する処理
        #
        #
        #
        # win32api.ShellExecute関数は、Windowsのシェルコマンドを実行するための関数で、
        # 第1引数には親ウィンドウのハンドル（0を指定するとバックグラウンドで実行）、
        # p第2引数には操作（'print'を指定すると印刷）、
        # 第3引数には印刷対象のファイルパス、
        # 第4引数には印刷するプリンター名またはIPアドレスを指定します

        debug_print(f"**************  Printing image: {image_path}")
        win32api.ShellExecute(0, 'print', image_path, None, '.', 0)



        # テストのため、印刷注意すること！！！
        time.sleep(4)
        pyautogui.press('enter')
        time.sleep(5)


        debug_print(f"**************  Printing Finished..............................................: {image_path}")

    except win32api.error as e:
        debug_print(f"印刷エラー: {e}")
        raise e

    finally:
        # 画像を閉じる
        image.close()

    debug_print("print_image 関数の終了")

def main():
    # 当日の日付を取得
    today = datetime.date.today()
    folder_name = today.strftime("%Y%m%d")

    # テスト用例
    folder_path = os.path.join(os.getcwd(), folder_name)
    debug_print("メイン処理の開始")
    print_and_rename_images(folder_path)

    #
    #
    #Teamsフォルダのテストコードより
    #
    #
    # TeamsのWebhook URLを設定してください

    # 前処理
    # folder_path内のすべてのファイル名を取得します
    file_names = os.listdir(folder_path)
    today = datetime.date.today().strftime("%Y%m%d")
    # ファイル名から拡張子を除いた部分を格納するリストを作成します
    file_names_without_extension = []
    for file_name in file_names:
        # {today} の日付を含む部分を削除します
        name_without_today = file_name.split(today)[-1]
        # 拡張子以降の部分を取得します
        name_without_extension = os.path.splitext(name_without_today)[0]
        file_names_without_extension.append(name_without_extension)

    # messageを作成します
    message = "Hello from Python! This is a test notification.<br>本日のプレス一覧です。<br><br>"

    # ファイル名の一覧をmessageに追加します（改行で区切ります）
    for name in file_names_without_extension:
        message += name + "<br><br>"

    # 本処理
    teams_webhook_url = "https://chubusunsho.webhook.office.com/webhookb2/bc534cb4-a282-41bc-a9e8-3f30791f72ea@e60fb0e2-e4db-4ea3-a60f-888216fb15cd/IncomingWebhook/cf2d3b99513e4876bd0b7a55eecd87cc/d5dbd0cb-710e-44bd-8d17-4b3be9317246"

    send_teams_notification(teams_webhook_url, message)

    debug_print("メイン処理の終了")

if __name__ == "__main__":
    main()
