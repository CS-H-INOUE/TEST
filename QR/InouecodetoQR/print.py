from PIL import Image
import os
import win32print
import win32ui

def debug_print(message):
    # デバッグモードの場合にメッセージを表示する関数
    DEBUG_MODE = True  # デバッグモードをオンにするか切り替えてください
    if DEBUG_MODE:
        print("[DEBUG]:", message)

def get_printer_physical_size(printer_name):
    hPrinter = win32print.OpenPrinter(printer_name)
    try:
        level = 2
        buf = win32print.GetPrinter(hPrinter, level)
        pDevModeObj, needed = win32print.DocumentProperties(None, hPrinter, printer_name, None, None, 0)
        physical_width = pDevModeObj.PaperWidth
        physical_height = pDevModeObj.PaperLength
        return physical_width, physical_height
    finally:
        win32print.ClosePrinter(hPrinter)

# 画像印刷関数
def print_image(file_path, printer_name):
    try:
        # 画像を開く
        img = Image.open(file_path)

        # 印刷対象のプリンターを選択
        debug_print(f"Selected printer: {printer_name}")

        hdc = win32ui.CreateDC()
        hdc.CreatePrinterDC(printer_name)

        # 画像サイズを取得
        img_width, img_height = img.size
        debug_print(f"Image size: {img_width}x{img_height}")

        # 印刷の設定
        printable_area_width, printable_area_height = get_printer_physical_size(printer_name)
        printer_scale = min(printable_area_width / img_width, printable_area_height / img_height)
        debug_print(f"Printable area size: {printable_area_width}x{printable_area_height}")
        debug_print(f"Printer scale: {printer_scale}")

        # 印刷を開始
        hdc.StartDoc(file_path)
        hdc.StartPage()

        # 画像を印刷
        debug_print("Printing image...")
        hdc.StretchBlt(
            (printable_area_width - int(printer_scale * img_width)) // 2,
            (printable_area_height - int(printer_scale * img_height)) // 2,
            int(printer_scale * img_width),
            int(printer_scale * img_height),
            img,
            0,
            0,
            img_width,
            img_height,
            win32print.SRCCOPY
        )

        hdc.EndPage()
        hdc.EndDoc()

        debug_print("Printing complete.")

    except Exception as e:
        debug_print(f"Error printing {file_path}: {e}")

def print_all_images_in_folder(folder_path, printer_ip):
    if not os.path.exists(folder_path):
        print(f"Folder '{folder_path}' does not exist.")
        return

    printer_name = fr"\\{printer_ip}"  # UNCパスにIPアドレスを指定
    physical_width, physical_height = get_printer_physical_size(printer_name)
    debug_print(f"Printer physical size: {physical_width}x{physical_height}")

    for filename in os.listdir(folder_path):
        file_path = os.path.join(folder_path, filename)
        if os.path.isfile(file_path) and filename.lower().endswith(('.png', '.jpg', '.jpeg', '.gif')):
            debug_print(f"Printing {file_path}")
            print_image(file_path, printer_name)

# 使用例
folder_name = "20230719"  # フォルダ名を設定してください
folder_path = os.path.join(os.getcwd(), folder_name)
printer_ip = "192.168.1.117"  # プリンターのIPアドレスを設定してください

print_all_images_in_folder(folder_path, printer_ip)
