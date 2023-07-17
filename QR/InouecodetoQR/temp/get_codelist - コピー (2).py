import tkinter as tk
from tkinter import simpledialog, messagebox
import openpyxl


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
            product = str(row[0]) + " " + "x".join(str(cell) for cell in row[2:])
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
    window.geometry("300x300")

    result_label = tk.Label(window, text="部分一致する商品:")
    result_label.pack(pady=10)

    listbox = tk.Listbox(window)
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

    select_button = tk.Button(window, text="選択", command=select_product)
    select_button.pack(pady=10)


def save_selected_product(selected_product):
    """選択された商品をファイルに保存します。"""
    with open("selected_product.txt", "w") as file:
        file.write(selected_product)


def main():
    # エクセルファイルのパスとシート名を指定
    excel_file = 'codedata.xlsx'
    sheet_name = 'Sheet1'

    # エクセルファイルを読み込む
    sheet = load_excel_file(excel_file, sheet_name)

    # GUIウィンドウを作成
    window = tk.Tk()
    window.title("商品検索")
    window.geometry("300x150")

    def search_button_click():
        """検索ボタンがクリックされたときの処理を行います。"""
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
    search_entry = tk.Entry(window)
    search_entry.pack(pady=10)

    # 検索ボタン
    search_button = tk.Button(window, text="検索", command=search_button_click)
    search_button.pack()

    # GUIのメインループを開始
    window.mainloop()


if __name__ == "__main__":
    main()
