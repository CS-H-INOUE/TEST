import subprocess
import os
import tkinter as tk
from tkinter import filedialog

def change_directory(directory):
    """指定されたディレクトリに移動します。"""
    os.chdir(directory)

def select_file():
    """ファイル選択ダイアログを表示し、選択されたファイルパスを返します。"""
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(title="Select Python File")
    return file_path

def create_venv(file_path):
    """選択されたファイル名を利用して仮想環境を構築します。"""
    print("選択されたファイル名を利用して仮想環境を構築します。", file_path)
    file_name = os.path.splitext(os.path.basename(file_path))[0]  # ファイル名（拡張子を除く）を取得
    command = f"py -m venv {file_name}_env"
    subprocess.call(command, shell=True)

def activate_venv(file_path, name):
    """仮想環境をアクティベートします。"""
    file_name = os.path.splitext(os.path.basename(file_path))[0]  # ファイル名（拡張子を除く）を取得
    activate_script = f"C:\\Users\\{name}\\Desktop\\PythonScripts\\{file_name}_env\\Scripts\\activate.ps1"
    print("Activating...", activate_script)
    command = f"powershell -ExecutionPolicy RemoteSigned -File {activate_script}"
    subprocess.call(command, shell=True)

def main():
    # ユーザー名の入力
        # name = input("ユーザー名を入力してください: ")
    name = "H-INOUE"

    # ディレクトリ移動処理
    path_cd = f"C:\\Users\\{name}\\Desktop\\PythonScripts"
    change_directory(path_cd)

    # Pythonファイル選択処理
    python_file = select_file()
    if python_file.endswith('.py'):
        # 仮想環境構築処理
        create_venv(python_file)
        # 仮想環境アクティベート処理
        activate_venv(python_file, name)
    else:
        print("Invalid file selected. Please choose a .py file.")

if __name__ == "__main__":
    main()
