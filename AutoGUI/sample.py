import os
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
import time

# ChromeOptionsを作成してシークレットモードを有効にする
# chrome_options = Options()
# chrome_options.add_argument("--incognito")
# デスクトップのパスを取得
desktop_path = os.path.expanduser("~/Desktop")

# 
# 
# Err　バージョンが異なるため、動作しない。仕方なく
# 
# 
# ユーザーデータディレクトリのパスを作成
# user_data_dir = os.path.join(desktop_path, "AppData", "Local", "Google", "Chrome", "User Data")
# chrome_options.add_argument(user_data_dir)

# ChromeDriverのパスを指定してWebDriverを作成
chrome_path = os.path.join(desktop_path, "Python","driver","chromedriver.exe")
# driver = webdriver.Chrome(chrome_path, options=chrome_options)
driver = webdriver.Chrome(chrome_path)

# ウェブページにアクセス
url = "https://www.customs.go.jp/toukei/srch/index.htm?M=77&P=0,2,,,,1,,,,2,,2017,%EF%BC%922023,,,7,293212000,,,,,,,,,,1,,,,,,,,,,,1,,,,,,,,,,,"
driver.get(url)

# 操作
time.sleep(2)  # ページが読み込まれるのを待つ

driver.find_element_by_tag_name('body').send_keys(Keys.TAB)  # Tabキーを押す
driver.find_element_by_tag_name('body').send_keys(Keys.TAB)  # Tabキーを押す
driver.find_element_by_tag_name('body').send_keys(Keys.SPACE)  # Spaceキーを押す
driver.find_element_by_tag_name('body').send_keys(Keys.ENTER)  # Enterキーを押す
time.sleep(2)  # ページが読み込まれるのを待つ

driver.find_element_by_tag_name('body').send_keys(Keys.TAB)  # Tabキーを押す
driver.find_element_by_tag_name('body').send_keys(Keys.TAB)  # Tabキーを押す
driver.find_element_by_tag_name('body').send_keys(Keys.SPACE)  # Spaceキーを押す
driver.find_element_by_tag_name('body').send_keys(Keys.DOWN)  # 下矢印キーを押す
driver.find_element_by_tag_name('body').send_keys(Keys.ENTER)  # Enterキーを押す
time.sleep(2)  # ページが読み込まれるのを待つ

driver.find_element_by_tag_name('body').send_keys(Keys.TAB)  # Tabキーを押す
driver.find_element_by_tag_name('body').send_keys('105')  # 105を入力
driver.find_element_by_tag_name('body').send_keys(Keys.TAB)  # Tabキーを押す
driver.find_element_by_tag_name('body').send_keys(Keys.TAB)  # Tabキーを押す
driver.find_element_by_tag_name('body').send_keys(Keys.TAB)  # Tabキーを押す
driver.find_element_by_tag_name('body').send_keys(Keys.SPACE)  # Spaceキーを押す
time.sleep(2)  # ページが読み込まれるのを待つ

driver.find_element_by_tag_name('body').send_keys(Keys.DOWN)  # 下矢印キーを押す
driver.find_element_by_tag_name('body').send_keys(Keys.ENTER)  # Enterキーを押す
driver.find_element_by_tag_name('body').send_keys(Keys.TAB)  # Tabキーを押す
driver.find_element_by_tag_name('body').send_keys('293212000')  # 293212000を入力

driver.find
