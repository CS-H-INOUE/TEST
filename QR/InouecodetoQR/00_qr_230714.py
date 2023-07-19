import tkinter as tk
from tkinter import messagebox, font
import openpyxl
import os
import os.path
import qrcode
from datetime import date
import datetime
from PIL import Image, ImageDraw, ImageFont
import shutil
import pyautogui

def create_daily_folder():
    # 実行ファイルの直下パスを取得
    base_path = os.path.dirname(os.path.abspath(__file__))
    # 当日の日付を取得
    today = datetime.date.today()
    folder_name = today.strftime("%Y%m%d")

    # 当日の日付を取得
    folder_path = os.path.join(os.getcwd(), folder_name)
    # 当日の日付フォルダが存在しない場合、作成する
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)
        print(f"Created folder: {folder_path}")

    # 前日以前の日付フォルダを削除する
    for item in os.listdir(base_path):
        item_path = os.path.join(base_path, item)
        if os.path.isdir(item_path):
            try:
                item_date = datetime.datetime.strptime(item, "%Y%m%d").date()
                if item_date < today:
                    shutil.rmtree(item_path)
                    print(f"Deleted folder: {item_path}")
            except ValueError:
                pass  # フォルダ名が日付形式ではない場合はスキップ

    # 実行ファイルの直下パスを返す
    return folder_path

def create_image_with_text(text, image_path, output_path):
    print("text,image_path,output_path",text,image_path,output_path)

    # 画像の読み込み
    image = Image.open(image_path)

    # 画像の上に白い領域を作成
    text_height = 200
    expanded_width = image.width + 200  # 両サイドに100ピクセルずつ広げる
    expanded_image = Image.new("RGB", (expanded_width, image.height + text_height), "white")
    expanded_image.paste(image, (100, text_height))  # 中央に画像を配置

    # テキストを描画する
    draw = ImageDraw.Draw(expanded_image)
    # メイリオフォント（meiryo.ttc）xを指定する
    font = ImageFont.truetype("meiryo.ttc", 28)

    # 左上にtextの
    # 左から最初のスペースまでを配置
    first_space_index = text.index(" ")
    text_position = (10, 10)
    draw.text(text_position,"コード："+ text[:first_space_index], font=font, fill="black")

    # 右上に品番を配置（次のスペースからその次のスペースまで）
    second_space_index = text.index(" ", first_space_index + 1)
    product_number_position = (expanded_width - 280, 10)
    draw.text(product_number_position, "品　番："+text[first_space_index+1:second_space_index], font=font, fill="black")

    # 商品サイズ
    description_font = ImageFont.truetype("meiryo.ttc", 40)  # テキストのフォントサイズを調整
    description_position = (10 , 100)
    draw.text(description_position, "商品サイズ", font=description_font, fill="black")

    # 板入数
    description_font = ImageFont.truetype("meiryo.ttc", 40)  # テキストのフォントサイズを調整
    description_position = (expanded_width -140 , 100)
    draw.text(description_position, "板入数", font=description_font, fill="black")

    # 11文字以降、右から5文字目までを表示する
    description_font = ImageFont.truetype("meiryo.ttc", 32)  # テキストのフォントサイズを調整
    description_position = (10,170)
    third_space_index = text.find(' ', second_space_index + 1)
    draw.text(description_position, text[second_space_index+1:third_space_index], font=description_font, fill="black")

    # 11桁以降を品番の下に右揃えで配置
    description_font = ImageFont.truetype("meiryo.ttc", 32)  # テキストのフォントサイズを調整
    description_position = (expanded_width-130 , 170)
    # extracted_text = text[-4:]  # 右から4文字を抽出
    extracted_text = text.rstrip().split()[-1]
    draw.text(description_position, extracted_text, font=description_font, fill="black")

    # 画像を保存する
    expanded_image.save(output_path)

def load_excel_file(excel_file, sheet_name):
    """指定されたエクセルファイルとシートを読み込みます。"""
    print("[debug] load_excel_file")
    workbook = openpyxl.load_workbook(excel_file)
    sheet = workbook[sheet_name]
    return sheet

def get_product_list(sheet):
    """指定されたシートから商品リストを取得します。"""
    print("[debug] get_product_list")
    products = []
    for row in sheet.iter_rows(min_row=2, values_only=True):
        if row[0] is not None:
            product = str(row[0]) + " " + str(row[1]) + " " + str(row[2]) + " " + str(row[4])
            products.append(product)
    return products

def search_products(products, search_input):
    """指定された商品リストから部分一致する商品を検索します。"""
    print("[debug] search_products")
    matching_products = []
    for product in products:
        if search_input.lower() in product.lower():
            matching_products.append(product)
    return matching_products

def show_matching_products(matching_products, qr_directory, main_window):
    """部分一致する商品の一覧を表示します。"""
    print("[debug] show_matching_products")
    window = tk.Toplevel(main_window)
    window.title("検索結果")

    # ウィンドウの幅と高さ
    window_width = 800
    window_height = 800

    # 画面の幅と高さを取得
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()
    # ウィンドウを画面の中央に配置
    x = (screen_width - window_width) // 2
    y = (screen_height - window_height) // 2
    window.geometry(f"{window_width}x{window_height}+{x}+{y}")

    # メイリオフォントを使用するための設定
    font_name = "Meiryo"
    bold_font = font.Font(family=font_name, size=16, weight="bold")

    result_label = tk.Label(window, text="部分一致する商品:", font=bold_font)
    result_label.pack(pady=10)

    listbox = tk.Listbox(window, font=(font_name, 12))
    listbox.pack(fill=tk.BOTH, expand=True)

    for product in matching_products:
        listbox.insert(tk.END, product)

    def select_product():
        selected_index = listbox.curselection()
        if selected_index:
            selected_product = matching_products[selected_index[0]]
            print(selected_product, "is selected!!! & will generate")
            qr_output_filename = generate_qr_code(selected_product, qr_directory)

            qr_directory_path = os.path.join(os.getcwd(), qr_directory)
            # qr_filename = f"qr_{selected_product}.png"
            qr_filepath = os.path.join(qr_directory_path, qr_output_filename)

            # first_five_digits = selected_product[:5]
            # qr_output_filename = f"qr_{first_five_digits}.png"
            # output_path = create_daily_folder() + "\\" + qr_output_filename
            output_path = create_daily_folder() + "\\" + qr_output_filename
            print("[Debug: output_path]",output_path)
            print("[debug]: create image with text")
            create_image_with_text(selected_product, qr_filepath, output_path)

            messagebox.showinfo("商品選択", "選択された商品が保存されました。")
            window.destroy()  # 検索画面を閉じる
            main_window.deiconify()  # メインウィンドウを再表示
        else:
            messagebox.showinfo("エラー", "商品が選択されていません。")

    def select_cancel():
            messagebox.showinfo("キャンセル", "メイン画面へ戻ります")

            window.destroy()  # 検索画面を閉じる
            main_window.deiconify()  # メインウィンドウを再表示

    def handle_key(event):
        if event.keysym == "Up":
            if listbox.curselection():
                current_index = listbox.curselection()[0]
                if current_index != 0:
                    listbox.selection_clear(0, tk.END)
                    listbox.selection_set(current_index - 1)
                    listbox.yview_scroll(-1, "units")
        elif event.keysym == "Down":
            if listbox.curselection():
                current_index = listbox.curselection()[0]
                if current_index != listbox.size() - 1:
                    listbox.selection_clear(0, tk.END)
                    listbox.selection_set(current_index + 1)
                    listbox.yview_scroll(1, "units")

    def on_closing():
        """ウィンドウが閉じられるときの処理を行います。"""
        if messagebox.askokcancel("終了", "本当に終了しますか？"):
            window.destroy()

    # Enterキーを検索ボタンとしてバインド
    window.bind("<Return>", lambda event: select_product())

    select_button = tk.Button(window, text="選択", command=select_product, font=bold_font)
    select_button.pack(pady=10)

    select_button = tk.Button(window, text="戻る", command=select_cancel, font=bold_font)
    select_button.pack(pady=10)

    # キーボードイベントを処理するためのバインド
    listbox.bind("<Up>", handle_key)
    listbox.bind("<Down>", handle_key)

    print("listbox.focus_get()",listbox.focus_get())
    main_window.focus_force()
    listbox.select_set(0)  # 一番上のアイテムを選択状態に設定
    listbox.activate(0)  # リストボックスをアクティブ化
    listbox.focus_set()  # リストボックスにキーボードフォーカスを設定
    print("listbox.focus_set()",listbox.focus_set())

    # ウィンドウの"×"ボタンを押したときの処理を設定
    window.protocol("WM_DELETE_WINDOW", on_closing)

    window.mainloop()

def generate_qr_code(selected_text, qr_directory):
    """
    テキストを使用してQRコードを生成して保存する関数

    Args:
        qr_text (str): QRコードのテキスト
        qr_directory (str): QRコードの保存先ディレクトリ名

    Returns:
        None
    """
    print("[debug] generate_qr_code")
    # QRコードの保存先ディレクトリのパスを作成
    qr_directory_path = os.path.join(os.getcwd(), qr_directory)
    # QRコードの保存先ディレクトリが存在しない場合は作成
    if not os.path.exists(qr_directory_path):
        os.makedirs(qr_directory_path)

    today = datetime.date.today()
    date_text = today.strftime("%Y%m%d")

    # コードのみ抽出するとき
    # first_space_index = selected_text.index(" ")
    qr_text = date_text+selected_text

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

    # QRコードのファイル名を作成
    qr_filename = f"qr_{qr_text}.png"

    print("[debug:qr_filename]",qr_filename)
    print("[debug:qr_text]",qr_text)

    # QRコードを保存
    qr_filepath = os.path.join(qr_directory_path, qr_filename)
    qr_image.save(qr_filepath)

    print("[debug:]: qr_image is saved.........",qr_filepath)

    return qr_filename

def main():
    # エクセルファイルのパスとシート名を指定
    excel_file = 'codedata.xlsx'
    sheet_name = 'Sheet1'

    # QRコードの保存先ディレクトリを指定
    qr_directory = 'QRディレクトリ'

    # エクセルファイルを読み込む
    sheet = load_excel_file(excel_file, sheet_name)

    # 商品の検索と選択を行うGUIウィンドウを作成
    window = tk.Tk()
    window.title("商品検索")

    # ウィンドウの幅と高さ
    window_width = 500
    window_height = 400

    # 画面の幅と高さを取得
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()

    # ウィンドウを画面の中央に配置
    x = (screen_width - window_width) // 2
    y = (screen_height - window_height) // 2
    window.geometry(f"{window_width}x{window_height}+{x}+{y}")

    # メイリオフォントを使用するための設定
    font_name = "Meiryo"
    bold_font = font.Font(family=font_name, size=16, weight="bold")
    # テキストを検索フィールドの前に表示
    label = tk.Label(window, text="検索したい文字列を入力してください.", font=("Meiryo", 14))
    label.pack(pady=10)
    label = tk.Label(window, text="（例. 50x 、T50、など）※半角数字で入力NG:５ OK:5", font=("Meiryo", 12))
    label.pack(pady=10)

    # 検索キーワードの入力フィールド
    search_entry = tk.Entry(window, font=("Meiryo", 16, "bold"))
    search_entry.pack(pady=10, ipadx=20, ipady=20)
    search_entry.focus_set()  # テキストフィールドにフォーカスを設定
    window.update()  # ウィンドウの更新
    search_entry.focus_set()  # 再度フォーカスを設定

    def search_button_click():
        """検索ボタンがクリックされたか、Enterキーが押されたときの処理を行います。"""
        search_input = search_entry.get().strip()

        if search_input:
            # 商品の一覧を取得
            products = get_product_list(sheet)

            # 部分一致する商品を検索
            matching_products = search_products(products, search_input)
            # print("[debug: matching products", matching_products)

            if matching_products:
                window.withdraw()  # メインウィンドウを非表示にする
                # 部分一致する商品の一覧を表示して、QR作成
                show_matching_products(matching_products, qr_directory, window)
            else:
                messagebox.showinfo("検索結果", "部分一致する商品はありません。")
        else:
            messagebox.showinfo("エラー", "検索キーワードを入力してください。")

    def on_closing():
        """ウィンドウが閉じられるときの処理を行います。"""
        if messagebox.askokcancel("終了", "本当に終了しますか？"):
            window.destroy()

    # Enterキーを検索ボタンとしてバインド
    search_entry.bind("<Return>", lambda event: search_button_click())
    # 終了ボタン
    exit_button = tk.Button(window, text="終了", command=on_closing, font=bold_font)
    exit_button.pack(side=tk.BOTTOM, pady=10)

    # ウィンドウの"×"ボタンを押したときの処理を設定
    window.protocol("WM_DELETE_WINDOW", on_closing)

    # GUIのメインループを開始
    window.mainloop()

if __name__ == "__main__":
    main()
