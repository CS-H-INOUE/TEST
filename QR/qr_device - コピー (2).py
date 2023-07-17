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
    qr_directory_path = os.path.join(os.path.expanduser("~"), "Desktop", qr_directory)

    # ディレクトリが存在しない場合は新規作成
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

def generate_qr_code_wrapper(qr_window, text_entry, date_entry, company_var, device_var):
    """
    QRコード生成を行うラッパー関数

    Args:
        qr_window (tk.Toplevel): QRコード生成用のウィンドウ
        text_entry (tk.Entry): 管理対象者入力フィールド
        date_entry (tk.Entry): 購入日付入力フィールド
        company_var (tk.StringVar): 会社名選択用の変数
        device_var (tk.StringVar): 管理デバイス選択用の変数

    Returns:
        None
    """
    # 入力値を取得
    text = text_entry.get()
    qr_date = date_entry.get()
    company = company_var.get()
    device = device_var.get()

    # QRコードのテキストを生成
    qr_text = f"{company}_管理デバイス{device}_購入日付{qr_date}_管理対象者{text}"

    # QRコードを生成して保存
    qr_directory = f"QRディレクトリ_{date.today().strftime('%Y%m%d')}"
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
    qr_window.geometry("400x400")

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

    # 会社名選択用のラジオボタン
    company_label = tk.Label(qr_window, text="会社名：")
    company_label.pack()

    company_var = tk.StringVar(qr_window)
    company_var.set("中部産商")  # デフォルトの選択肢

    company_radio_a = tk.Radiobutton(qr_window, text="中部産商", variable=company_var, value="中部産商")
    company_radio_a.pack()
    company_radio_b = tk.Radiobutton(qr_window, text="イノウエ", variable=company_var, value="イノウエ")
    company_radio_b.pack()

    # 管理デバイス選択用のラジオボタン
    device_label = tk.Label(qr_window, text="管理デバイス：")
    device_label.pack()

    device_var = tk.StringVar(qr_window)
    device_var.set("PC")  # デフォルトの選択肢

    device_radio_a = tk.Radiobutton(qr_window, text="PC", variable=device_var, value="PC")
    device_radio_a.pack()
    device_radio_b = tk.Radiobutton(qr_window, text="iPad", variable=device_var, value="iPad")
    device_radio_b.pack()
    device_radio_c = tk.Radiobutton(qr_window, text="ルータ", variable=device_var, value="ルータ")
    device_radio_c.pack()
    device_radio_d = tk.Radiobutton(qr_window, text="AP", variable=device_var, value="AP")
    device_radio_d.pack()
    device_radio_e = tk.Radiobutton(qr_window, text="NAS", variable=device_var, value="NAS")
    device_radio_e.pack()
    device_radio_f = tk.Radiobutton(qr_window, text="その他", variable=device_var, value="その他")
    device_radio_f.pack()

    # ボタン
    generate_button = tk.Button(qr_window, text="QRコード生成", command=lambda: generate_qr_code_wrapper(qr_window, text_entry, date_entry, company_var, device_var))
    generate_button.pack()

# GUIを生成
window = tk.Tk()
window.title("QRコード生成")
window.geometry("300x200")

# QRコード生成ボタン
generate_button = tk.Button(window, text="QRコード生成", command=create_qr_code_window)
generate_button.pack()

# イベントループ開始
window.mainloop()
