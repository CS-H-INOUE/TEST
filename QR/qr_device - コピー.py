import tkinter as tk
import qrcode
import os
from datetime import date

def generate_qr_code(qr_text, qr_directory):
    """
    QRコードを生成して保存する関数

    Args:
        qr_text (str): QRコードのテキスト
        qr_directory (str): QRコードの保存先ディレクトリ名

    Returns:
        None
    """
    # QRコードの保存先ディレクトリのパスを作成
    qr_directory_path = os.path.join(os.getcwd(), qr_directory)

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

def generate_qr_code_wrapper(qr_window, text_entry, date_entry):
    """
    QRコード生成を行うラッパー関数

    Args:
        qr_window (tk.Toplevel): QRコード生成用のウィンドウ
        text_entry (tk.Entry): 管理対象者入力フィールド
        date_entry (tk.Entry): 購入日付入力フィールド

    Returns:
        None
    """
    # 入力値を取得
    text = text_entry.get()
    qr_date = date_entry.get()

    # QRコードのテキストを生成
    qr_text = f"中部産商_管理デバイスPC_購入日付{qr_date}_管理対象者{text}"

    # QRコードを生成して保存
    qr_directory = 'QRディレクトリ'
    generate_qr_code(qr_text, qr_directory)

    # ウィンドウを閉じる
    qr_window.destroy()

def create_qr_code_window():
    """
    QRコード生成用のウィンドウを生成する関数

    Args:
        None

    Returns:
        None
    """
    # QRコード生成用のウィンドウを生成
    qr_window = tk.Toplevel(window)
    qr_window.title("QRコード生成")
    qr_window.geometry("400x200")

    # 管理対象者入力フィールド
    text_label = tk.Label(qr_window, text="管理対象者：")
    text_label.pack()
    text_entry = tk.Entry(qr_window)
    text_entry.pack()

    # 購入日付入力フィールド
    date_label = tk.Label(qr_window, text="購入日付（YYYYMMDD）：")
    date_label.pack()
    date_entry = tk.Entry(qr_window)
    date_entry.pack()

    # ボタン
    generate_button = tk.Button(qr_window, text="QRコード生成", command=lambda: generate_qr_code_wrapper(qr_window, text_entry, date_entry))
    generate_button.pack()

def open_qr_code_window():

    # QRコード生成ウィンド
    # ウィンドウを表示
    create_qr_code_window()

def display_qr_code(qr_text):
    """
    QRコードを別のウィンドウで表示する関数

    Args:
        qr_text (str): 表示するQRコードのテキスト

    Returns:
        None
    """
    qr_directory = 'QRディレクトリ'
    qr_filename = f"qr_{qr_text}.png"
    qr_filepath = os.path.join(os.getcwd(), qr_directory, qr_filename)

    # 別のウィンドウでQRコードを表示
    display_window = tk.Toplevel(window)
    display_window.title("QRコード表示")
    display_window.geometry("300x300")

    # QRコード画像を表示
    qr_image = tk.PhotoImage(file=qr_filepath)
    qr_label = tk.Label(display_window, image=qr_image)
    qr_label.pack()

# GUIを生成
window = tk.Tk()
window.title("QRコード生成")
window.geometry("400x200")

# QRコード生成ボタン
generate_button = tk.Button(window, text="QRコード生成", command=open_qr_code_window)
generate_button.pack()

# イベントループ開始
window.mainloop()
