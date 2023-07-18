import tkinter as tk
import qrcode
import openpyxl
import os
from datetime import date
from PIL import ImageGrab

def generate_qr_code(qr_text, qr_directory):
    """
    テキストを使用してQRコードを生成して保存する関数

    Args:
        qr_text (str): QRコードのテキスト
        qr_directory (str): QRコードの保存先ディレクトリ名

    Returns:
        None
    """
    # QRコードの保存先ディレクトリのパスを作成
    qr_directory_path = os.path.join(os.getcwd(), qr_directory)
    # QRコードの保存先ディレクトリが存在しない場合は作成
    if not os.path.exists(qr_directory_path):
        os.makedirs(qr_directory_path)

    # QRコードを生成
    qr = qrcode.QRCode(
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_H,
        box_size=10,
        border=4,
    )
    qr.add_data(qr_text)
    qr.make(fit=True)

    # QRコードをPIL Imageオブジェクトに変換
    qr_image = qr.make_image(fill_color="black", back_color="white")

    # QRコードを保存
    qr_filename = f"qr_{qr_text}.png"
    qr_filepath = os.path.join(qr_directory_path, qr_filename)
    qr_image.save(qr_filepath)

def display_text_and_qr_code(text, qr_text, qr_directory, save_screenshot=False):
    """
    テキストとQRコードを同時に表示するGUIウィンドウを作成する関数

    Args:
        text (str): 表示するテキスト
        qr_text (str): QRコードのテキスト
        qr_directory (str): QRコードの保存先ディレクトリ名

    Returns:
        None
    """
    # GUIウィンドウを作成
    window = tk.Tk()
    window.title("Text and QR Code Display")

    # テキストの表示
    text_label = tk.Label(window, text=text, font=("Arial", 16))
    text_label.pack(pady=10)

    # QRコードの表示
    generate_qr_code(qr_text, qr_directory)  # QRコードを生成して保存
    qr_filepath = os.path.join(qr_directory, f"qr_{qr_text}.png")
    qr_image = tk.PhotoImage(file=qr_filepath)
    qr_label = tk.Label(window, image=qr_image)
    qr_label.image = qr_image  # 参照を保持するために必要
    qr_label.pack(pady=10)

    # GUIのメインループを開始
    window.mainloop()

    # スクリーンショットを保存
    if save_screenshot:
        screenshot_filename = f"screenshot_{qr_text}.png"
        screenshot_filepath = os.path.join(qr_directory, screenshot_filename)
        window.update()  # ウィンドウを更新してキャプチャ対象にする
        window.withdraw()  # ウィンドウを非表示にする
        window.update_idletasks()  # イベントの処理が完了するまで待機
        window.deiconify()  # ウィンドウを再表示する
        window.attributes("-topmost", True)  # ウィンドウを最前面に表示する
        window.update()  # ウィンドウを更新してキャプチャ対象にする
        window.geometry(window.winfo_geometry())  # ウィンドウのサイズを固定
        window.update()  # ウィンドウを更新してキャプチャ対象にする
        window.update_idletasks()  # イベントの処理が完了するまで待機
        window.update()  # ウィンドウを更新してキャプチャ対象にする
        window.after(100)  # ウィンドウの再描画のための短い待機時間
        window.update()  # ウィンドウを更新してキャプチャ対象にする
        window.update_idletasks()  # イベントの処理が完了するまで待機
        window.update()  # ウィンドウを更新してキャプチャ対象にする
        window.grab_set()  # ウィンドウの外部からの操作をブロックする
        window.update()  # ウィンドウを更新してキャプチャ対象にする
        window.update_idletasks()  # イベントの処理が完了するまで待機
        window.update()  # ウィンドウを更新してキャプチャ対象にする
        window.focus_set()  # ウィンドウをフォーカスする
        window.update()  # ウィンドウを更新してキャプチャ対象にする
        window.grab_release()  # ウィンドウの外部からの操作のブロックを解除する
        window.update()  # ウィンドウを更新してキャプチャ対象にする
        window.attributes("-topmost", False)  # ウィンドウを最前面でなくする
        window.update()  # ウィンドウを更新してキャプチャ対象にする

        # スクリーンショットを保存
        window_image = ImageGrab.grab()
        window_image.save(screenshot_filepath)

def generate_qr_codes(excel_file, sheet_name, qr_directory, save_screenshot=False):
    """
    エクセルファイルからデータを読み取り、テキストとQRコードを生成する関数

    Args:
        excel_file (str): エクセルファイルのパス
        sheet_name (str): シート名
        qr_directory (str): QRコードの保存先ディレクトリ名

    Returns:
        None
    """
    # エクセルファイルを読み込み
    workbook = openpyxl.load_workbook(excel_file)
    sheet = workbook[sheet_name]

    # 今日の日付を取得
    today = date.today().strftime("%Y%m%d")

    for row in sheet.iter_rows(values_only=True):
        no = str(row[0])
        code = str(row[1])
        name = str(row[2])
        quantity = str(row[3])
        text = f"No: {no}\nCode: {code}\nName: {name}\nq: {quantity}個入"
        qr_text = f"{today}{no}"
        # display_text_and_qr_code(text, qr_text, qr_directory)
        display_text_and_qr_code(text, qr_text, qr_directory, save_screenshot=True)

    # エクセルファイルを閉じる
    workbook.close()

# メインの処理
if __name__ == "__main__":
    # エクセルファイルのパスとシート名
    excel_file = 'data.xlsx'
    sheet_name = 'Sheet'

    # QRコードの保存先ディレクトリ
    qr_directory = 'QRディレクトリ'

    # QRコード生成とテキスト＆QRコード表示の関数を呼び出す
    generate_qr_codes(excel_file, sheet_name, qr_directory)
