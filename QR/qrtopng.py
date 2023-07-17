
import os
import openpyxl
from PIL import Image, ImageFont, ImageDraw

def print_qr_codes(excel_file, sheet_name, qr_directory):
    """
    エクセルファイルからデータを読み取り、一致するQRコード画像を印刷する関数

    Args:
        excel_file (str): エクセルファイルのパス
        sheet_name (str): シート名
        qr_directory (str): QRコード画像が保存されているディレクトリのパス

    Returns:
        None

    Memo:
        if change codes_per_pages , qr_width //a ,qr_height //b, x= c,y =d will be changed!!!

    """
    # メイリオフォントのパス
    font_path = "C:/Windows/Fonts/meiryo.ttc"

    # エクセルファイルを読み込み
    workbook = openpyxl.load_workbook(excel_file)
    sheet = workbook[sheet_name]

    # QRコード画像のファイルリストを取得
    qr_files = [file for file in os.listdir(qr_directory) if file.endswith('.png')]

    # QRコード画像の数を取得
    num_qr_codes = len(qr_files)

    # 1ページに印刷するQRコードの数
    codes_per_page = 4 * 5

    # A4用紙のサイズ
    page_width = 2480  # ピクセル単位
    page_height = 3508  # ピクセル単位

    # 1つのQRコードの幅と高さ
    qr_width = page_width // 4
    qr_height = page_height // 5

    # 印刷するQRコード画像の座標
    x = 0
    y = 80

    # メイリオフォントを指定
    title_font = ImageFont.truetype(font_path, 48)

    # 新しいページのイメージを作成
    page_image = Image.new('RGB', (page_width, page_height), 'white')
    draw = ImageDraw.Draw(page_image)

    # ページカウンタ
    page_count = 1

    # QRコードを1つずつ配置していく
    for i in range(num_qr_codes):
        qr_code_file = qr_files[i]
        qr_code_path = os.path.join(qr_directory, qr_code_file)

        # QRコード画像を開く
        qr_code_image = Image.open(qr_code_path)

        # QRコード画像を指定の位置に貼り付け
        # page_image.paste(qr_code_image, (x, y))

        # QRコード画像を指定の位置に貼り付け
        qr_code_image_resized = qr_code_image.resize((qr_width, qr_height))
        page_image.paste(qr_code_image_resized, (x, y))


        # QRコード画像のファイル名を取得
        qr_code_name = os.path.splitext(qr_code_file)[0]

        # QRコードのファイル名から検索ワードを抽出
        search_word = qr_code_name[3:]  # ファイル名の4文字目以降から.pngまでを抽出
        print("qr_code_name:",qr_code_name)
        print("search:",search_word)

        # エクセルのA列と比較して一致するC列の値を検索
        data_c = None
        for row in sheet.iter_rows(values_only=True):
            if int(row[0]) - int(search_word) == 0:
                print("search data_c ",row[0],search_word)
                data_c = row[2]
                break

        print("cant search ",row[0],search_word)

        # 画像にタイトルを追加
        try:
            draw.text((x, y), qr_code_name, font=title_font, fill='black')

            # データが存在する場合はデータを追加
            if data_c:
                # 修正前
                # draw.text((x, y - title_font.getsize(data_c)[1]), data_c, font=title_font, fill='black')
                # 修正後
                bbox = draw.textbbox((x, y-50), data_c, font=title_font)
                draw.text(bbox[:2], data_c, font=title_font, fill='black')

        except UnicodeEncodeError:
            print("データを追加できません:", qr_code_name, data_c)

        x += qr_width

        if x >= page_width:
            x = 0
            y += qr_height

        # ページが一杯になったら次のページへ
        if (i + 1) % codes_per_page == 0:
            page_image.save(f"printed_page_{page_count}.png")
            print(f"ページ {page_count} が保存されました。")
            page_count += 1
            x = 0
            y = 80
            page_image = Image.new('RGB', (page_width, page_height), 'white')
            draw = ImageDraw.Draw(page_image)

    # 残りのQRコードを保存
    if num_qr_codes % codes_per_page != 0:
        page_image.save(f"printed_page_{page_count}.png")
        print(f"ページ {page_count} が保存されました。")

if __name__ == "__main__":
    # エクセルファイルのパス
    excel_file = 'data.xlsx'

    # 使用するシート名
    sheet_name = 'Sheet1'

    # QRコードの保存先ディレクトリ
    qr_directory = 'QRディレクトリ'

    # QRコードを印刷する関数を呼び出す
    print_qr_codes(excel_file, sheet_name, qr_directory)

