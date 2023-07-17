import tkinter as tk
from tkinter import messagebox, font
import openpyxl
import os

def load_excel_file(excel_file, sheet_name):
    """指定されたエクセルファイルとシートを読み込みます。"""
    workbook = openpyxl.load_workbook(excel_file)
    sheet = workbook[sheet_name]
    return sheet


def get_product_list(sheet):
    """指定されたシートから商品リストを取得します。"""
    products = []
    for row in sheet.iter_rows(min_row=2, values_only=True):
        if row[0] is not None:
            product = str(row[0]) + " " + str(row[1])+ " " +str(row[2])+ " " +str(row[4])
            products.append(product)
    return products


def search_products(products, search_input):
    """指定された商品リストから部分一致する商品を検索します。"""
    matching_products = []
    for product in products:
        if search_input.lower() in product.lower():
            matching_products.append(product)
    return matching_products


def show_matching_products(matching_products):
    """部分一致する商品の一覧を表示します。"""
    window = tk.Toplevel()
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
        """商品の選択を行います。"""
        selected_index = listbox.curselection()
        if selected_index:
            selected_product = matching_products[selected_index[0]]
            save_selected_product(selected_product)
            messagebox.showinfo("商品選択", "選択された商品が保存されました。")
            window.destroy()
        else:
            messagebox.showinfo("エラー", "商品が選択されていません。")

    # Enterキーを検索ボタンとしてバインド
    window.bind("<Return>", lambda event:select_product())

    select_button = tk.Button(window, text="選択", command=select_product, font=bold_font)
    select_button.pack(pady=10)

    window.mainloop()

def save_selected_product(selected_product):
    """選択された商品をファイルに保存します。"""
    with open("selected_product.txt", "a") as file:
        file.write(selected_product + "\n")

def main():

    file_path = "selected_product.txt"
    # ファイルが存在しない場合のみ作成
    if not os.path.isfile(file_path):
        with open(file_path, "w") as file:
            file.write("")
    with open(file_path, "w") as file:
        file.write("")

    # エクセルファイルのパスとシート名を指定
    excel_file = 'codedata.xlsx'
    sheet_name = 'Sheet1'

    # エクセルファイルを読み込む
    sheet = load_excel_file(excel_file, sheet_name)

    # GUIウィンドウを作成
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

    def search_button_click():
        """検索ボタンがクリックされたか、Enterキーが押されたときの処理を行います。"""
        search_input = search_entry.get().strip()

        if search_input:
            # 商品の一覧を取得
            products = get_product_list(sheet)

            # 部分一致する商品を検索
            matching_products = search_products(products, search_input)

            if matching_products:
                # 部分一致する商品の一覧を表示
                show_matching_products(matching_products)
            else:
                messagebox.showinfo("検索結果", "部分一致する商品はありません。")

            # 再度検索を行う
            search_entry.delete(0, tk.END)  # 検索キーワードの入力フィールドをクリア
        else:
            messagebox.showinfo("エラー", "検索キーワードを入力してください。")

    # 検索キーワードの入力フィールド
    search_entry = tk.Entry(window, font=("Meiryo", 16, "bold"))
    search_entry.pack(pady=10, ipadx=20, ipady=20)

    # Enterキーを検索ボタンとしてバインド
    search_entry.bind("<Return>", lambda event: search_button_click())
    # 検索ボタン
    search_button = tk.Button(window, text="検索", command=search_button_click, font=bold_font)
    search_button.pack(side=tk.BOTTOM, pady=10)


    # GUIのメインループを開始
    window.mainloop()


if __name__ == "__main__":
    main()

