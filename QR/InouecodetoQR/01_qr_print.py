import os
import datetime
from PIL import Image
import tkinter as tk
from tkinter import messagebox
import win32api
import win32print

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
        # 画像を印刷する処理
        debug_print(f"**************  Printing image: {image_path}")
        printer_ip= "192.168.1.117"
        win32api.ShellExecute(0, 'print', image_path, f'/d:\\\\{printer_ip}', '.', 0)
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
    debug_print("メイン処理の終了")

if __name__ == "__main__":
    main()
