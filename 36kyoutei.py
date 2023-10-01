import sys
import time
import pandas as pd
import pyautogui
import pyperclip
import pygetwindow as gw

def activate_window(title):
    # 指定されたタイトルを含むウィンドウを検索してアクティブ化、最大化する
    target_windows = gw.getWindowsWithTitle(title)
    if not target_windows:
        print(f"指定されたウィンドウ({title})が見つかりませんでした。")
        sys.exit()

    target_window = target_windows[0]  # 最初に見つかったウィンドウを選択
    target_window.activate()
    time.sleep(1)  # ウィンドウがアクティブになるのを待つ
    target_window.maximize()
    time.sleep(1)  # ウィンドウが最大化されるのを待つ

def click_on_image(image, max_attempts=2):
    # 指定された画像を検索してクリックする
    for attempt in range(max_attempts):
        location = pyautogui.locateCenterOnScreen(image)
        if location:
            pyautogui.click(location.x, location.y)
            time.sleep(1)
            return
        time.sleep(1)

    print(f"{image}が見つかりませんでした。")

def paste_text(data, start=0, length=None):
    """
    指定されたデータの一部または全部をコピーしてペーストし、Tabキーを押す関数

    Parameters:
    data (str): ペーストしたい文字列
    start (int, optional): ペーストを開始する位置（デフォルトは0）1文字目は0となる。
    length (int, optional): ペーストする文字数（デフォルトはNone、文字列の末尾まで）
    """
    # start位置からlength文字のサブストリングを取得
    substring = str(data)[start : start + length] if length is not None else str(data)[start:]
    
    # サブストリングをコピーしてペーストし、Tabキーを押す
    pyperclip.copy(substring)
    print(substring)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(0.5)
    pyautogui.press("tab")
    time.sleep(0.5)

def enter_text(text):
    # 指定されたテキストを1文字ずつ入力し、Enterキーを押してからTabキーを押す
    for char in str(text):
        pyautogui.press(char)
        print(char)
        time.sleep(0.5)
    pyautogui.press("enter")
    time.sleep(0.5)
    pyautogui.press("tab")
    time.sleep(0.3)

def select_era(era):
    """
    指定された和暦を選択する関数

    Parameters:
    era (str): 選択したい和暦（"令和"または"平成"）
    """
    # 和暦を選択
    if era == "令和":
        pyautogui.press("up")
        time.sleep(0.3)
    elif era == "平成":
        pyautogui.press("up")
        time.sleep(0.3)
        pyautogui.press("up")
        time.sleep(0.3)
    pyautogui.press("tab")
    time.sleep(1)

def main():
    # Excelファイルを読み込む
    df = pd.read_excel('36協定データ.xlsx')

    # 準備がOKか確認するダイアログを表示
    res = pyautogui.confirm(text='e-Gov電子申請マイページ 準備OK？', title='確認', buttons=['OK', 'Cancel'])
    if res != 'OK':
        sys.exit()

    time.sleep(2)

    # e-Gov電子申請マイページのウィンドウをアクティブ化
    activate_window('e-Gov電子申請マイページ')

    # 画像を検索してクリック
    click_on_image('bookmark.png')
    click_on_image('36-ikkatu-tokubetujoukou.png')

    # ページダウンキーを押してスクロール
    for _ in range(10):
        pyautogui.press('pgdn')
        time.sleep(0.3)

    # 画像を検索してクリック
    click_on_image('sinseisho-nyuuryoku.png')
    time.sleep(3)

    # 労働保険番号の入力欄（都道府県）に移動
    pyautogui.press("tab")
    time.sleep(1)

    # 日本語入力をOFFにする
    pyautogui.hotkey('hanja')
    time.sleep(1)

    # DataFrameの各行に対して処理
    for _, r in df.iterrows():
        # 労働保険番号(例：13512345678-001)を5回ペースト
        paste_text(r['労働保険番号'], 0, 2)     # 都道府県 2桁
        paste_text(r['労働保険番号'], 2, 1)     # 所掌 1桁
        paste_text(r['労働保険番号'], 3, 2)     # 管轄 2桁
        paste_text(r['労働保険番号'], 5, 6)     # 基幹番号 6桁
        paste_text(r['労働保険番号'], 12, 3)     # 枝番号 3桁

        # その他のデータをペースト
        paste_text(r['被一括事業場番号'])
        paste_text(r['法人番号'])

        # 協定の有効期間(自)の和暦を選択
        select_era(r['協定の有効期間(自)和暦'])

        # 日本語入力をOFFにする
        pyautogui.hotkey('hanja')
        time.sleep(1)

        # 年月日を入力
        enter_text(r['協定の有効期間(自)年'])
        enter_text(r['協定の有効期間(自)月'])
        enter_text(r['協定の有効期間(自)日'])

        # 協定の有効期間(至)の和暦を選択
        select_era(r['協定の有効期間(至)和暦'])

        # 年月日を入力
        enter_text(r['協定の有効期間(至)年'])
        enter_text(r['協定の有効期間(至)月'])
        enter_text(r['協定の有効期間(至)日'])

        # 8回タブキーを押す
        for _ in range(8):
            pyautogui.press("tab")
            time.sleep(0.3)

        # 起算日和暦を選択
        select_era(r['起算日和暦'])

        # 起算日年月日を入力
        enter_text(r['起算日年'])
        enter_text(r['起算日月'])
        enter_text(r['起算日日'])

        # 1回タブキーを押す
        pyautogui.press("tab")
        time.sleep(0.3)
        paste_text(r['時間外労働をさせる必要のある具体的自由(自由記入)'])
        pyautogui.press("tab")
        time.sleep(0.3)

        # 1回タブキーを押す
        pyautogui.press("tab")
        time.sleep(0.3)

        # 日本語入力をOFFにする
        pyautogui.hotkey('hanja')
        time.sleep(1)

        paste_text(r['業務の種類(自由記入)'])
        pyautogui.press("tab")
        time.sleep(0.3)

        enter_text(r['所定労働時間(1日)時間'])
        enter_text(r['所定労働時間(1日)分'])

        pyautogui.press("tab")
        time.sleep(0.3)
        pyautogui.press("tab")
        time.sleep(0.3)

        enter_text(r['1日法定労働時間を超える時間数時間'])
        enter_text(r['1日法定労働時間を超える時間数分'])


if __name__ == "__main__":
    main()
