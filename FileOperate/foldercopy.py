import os
import tkinter as tk
from tkinter import filedialog, messagebox


def select_folders():
    """エクスプローラでフォルダを選択するダイアログを表示し、選択されたフォルダのパスを返す関数"""
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    folder_path = filedialog.askdirectory(
        parent=root,
        title='フォルダを選択してください'
    )
    return folder_path


def select_destination():
    """エクスプローラで移動先のフォルダを選択するダイアログを表示し、選択されたフォルダのパスを返す関数"""
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    folder_path = filedialog.askdirectory(
        parent=root,
        title='移動先のフォルダを選択してください'
    )
    return folder_path


def move_folder(source_path, destination_path):
    """指定されたフォルダを移動する関数"""
    folder_name = os.path.basename(source_path)
    destination_folder_path = os.path.join(destination_path, folder_name)
    os.rename(source_path, destination_folder_path)


# フォルダの移動
while True:
    # 移動するフォルダの選択
    print("移動するフォルダを選択してください...")
    source_folder = select_folders()

    # 移動先のフォルダの選択
    print("移動先のフォルダを選択してください...")
    destination_folder = select_destination()

    # フォルダの移動
    print("フォルダを移動しています...")
    move_folder(source_folder, destination_folder)
    print("移動が完了しました。")

    # 追加の移動を行うか確認
    choice = messagebox.askquestion("追加の移動", "他にも移動するフォルダがありますか？")
    if choice != "yes":
        break

print("移動処理が終了しました。")