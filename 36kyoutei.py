import pyautogui
import pyperclip
import time
import pandas as pd
import sys
import pygetwindow as gw


df = pd.read_excel('36協定データ.xlsx') 
res = pyautogui.confirm(text='e-Gov電子申請マイページ 準備OK？', title='確認', buttons=['OK', 'Cancel'])
if res == 'OK':
    time.sleep(2)
else:
    sys.exit()

# # ウィンドウのタイトルに「e-Gov電子申請マイページ」という文字列が含まれるウィンドウを検索
# target_windows = gw.getWindowsWithTitle('e-Gov電子申請マイページ')
# ウィンドウのタイトルに「申請書入力｜e-Gov電子申請」という文字列が含まれるウィンドウを検索
target_windows = gw.getWindowsWithTitle('申請書入力｜e-Gov電子申請')

if not target_windows:
    print("指定されたウィンドウが見つかりませんでした。")
else:
    # 1. 見つかったウィンドウをアクティブにする
    target_window = target_windows[0]  # 最初に見つかったウィンドウを選択
    target_window.activate()
    time.sleep(1)  # ウィンドウがアクティブになるのを待つ

    # 2. ウィンドウを最大化
    target_window.maximize()
    time.sleep(1)  # ウィンドウが最大化されるのを待つ

    # # 3. 「手続きブックマーク」という文字列をクリック
    # # 以下は画像検索の例です（「手続きブックマーク」のスクリーンショットを 'bookmark.png' として保存しておく必要があります）
    # location = pyautogui.locateCenterOnScreen('bookmark.png')
    # if location:
    #     pyautogui.click(location.x, location.y)
    #     time.sleep(1)  # クリックされるのを待つ
    # else:
    #     print("「手続きブックマーク」の文字列が見つかりませんでした。")

    # # 4. 「時間外労働・休日労働に関する協定届（本社一括届）（特別条項付き）」という文字列をクリック
    # # 以下は画像検索の例です（「時間外労働・休日労働に関する協定届（本社一括届）（特別条項付き）」のスクリーンショットを '36-ikkatu-tokubetujoukou.png' として保存しておく必要があります）
    # location = pyautogui.locateCenterOnScreen('36-ikkatu-tokubetujoukou.png')
    # if location:
    #     pyautogui.click(location.x, location.y)
    #     time.sleep(1)  # クリックされるのを待つ
    # else:
    #     print("「時間外労働・休日労働に関する協定届（本社一括届）（特別条項付き）」の文字列が見つかりませんでした。")
    #     location = pyautogui.locateCenterOnScreen('36-ikkatu-tokubetujoukou.png')
    #     if location:
    #         pyautogui.click(location.x, location.y)
    #         time.sleep(1)  # クリックされるのを待つ
    #     else:
    #         print("「時間外労働・休日労働に関する協定届（本社一括届）（特別条項付き）」の文字列が見つかりませんでした。")

    # # 5. ページダウンキーで一番下までスクロールする
    # for _ in range(10):  # 適切な回数を設定してください
    #     pyautogui.press('pgdn')  # ページダウンキーを押す
    #     time.sleep(0.1)  # キーストローク間の待機時間

    # # 6. 「申請書入力へ」という文字列をクリック
    # # 以下は画像検索の例です（「申請書入力へ」のスクリーンショットを 'sinseisho-nyuuryoku.png' として保存しておく必要があります）
    # location = pyautogui.locateCenterOnScreen('sinseisho-nyuuryoku.png')
    # if location:
    #     pyautogui.click(location.x, location.y)
    #     time.sleep(3)  # クリックされるのを待つ
    # else:
    #     print("「申請書入力へ」の文字列が見つかりませんでした。")

    # 最初の入力位置へ移動（労働保険番号 都道府県）
    pyautogui.press("tab")

    for i, r in df.iterrows():
        pyperclip.copy(str(r['労働保険番号']))  
        pyautogui.hotkey('ctrl', 'v')
        pyautogui.press("tab")
        time.sleep(0.3)  # キーストローク間の待機時間

        pyperclip.copy(str(r['労働保険番号']))  
        pyautogui.hotkey('ctrl', 'v')
        pyautogui.press("tab")
        time.sleep(0.3)  # キーストローク間の待機時間

        pyperclip.copy(str(r['労働保険番号']))  
        pyautogui.hotkey('ctrl', 'v')
        pyautogui.press("tab")
        time.sleep(0.3)  # キーストローク間の待機時間

        pyperclip.copy(str(r['労働保険番号']))  
        pyautogui.hotkey('ctrl', 'v')
        pyautogui.press("tab")
        time.sleep(0.3)  # キーストローク間の待機時間

        pyperclip.copy(str(r['労働保険番号']))  
        pyautogui.hotkey('ctrl', 'v')
        pyautogui.press("tab")
        time.sleep(0.3)  # キーストローク間の待機時間

        pyperclip.copy(str(r['被一括事業場番号']))  
        pyautogui.hotkey('ctrl', 'v')
        pyautogui.press("tab")
        time.sleep(0.3)  # キーストローク間の待機時間

        pyperclip.copy(str(r['法人番号']))  
        pyautogui.hotkey('ctrl', 'v')
        pyautogui.press("tab")
        time.sleep(0.3)  # キーストローク間の待機時間

        pyperclip.copy(str(r['協定の有効期間(自)和暦']))  
        pyautogui.press("up")
        pyautogui.press("tab")
        time.sleep(0.3)  # キーストローク間の待機時間

        text = str(r['協定の有効期間(自)年'])
        # 文字列を1文字ずつ対象アプリケーションに送信
        for char in text:
            print(char)
            pyautogui.press(char)
            # オプション: 各文字の間に遅延を入れる
            time.sleep(0.5)
        pyautogui.press("enter")
        time.sleep(0.5)
        pyautogui.press("tab")
        time.sleep(0.3)  # キーストローク間の待機時間

        text = str(r['協定の有効期間(自)月'])
        # 文字列を1文字ずつ対象アプリケーションに送信
        for char in text:
            print(char)
            pyautogui.press(char)
            # オプション: 各文字の間に遅延を入れる
            time.sleep(0.5)
        pyautogui.press("enter")
        time.sleep(0.5)
        pyautogui.press("tab")
        time.sleep(0.3)  # キーストローク間の待機時間

        text = str(r['協定の有効期間(自)日'])
        # 文字列を1文字ずつ対象アプリケーションに送信
        for char in text:
            print(char)
            pyautogui.press(char)
            # オプション: 各文字の間に遅延を入れる
            time.sleep(0.5)
        pyautogui.press("enter")
        time.sleep(0.5)
        pyautogui.press("tab")
        time.sleep(0.3)  # キーストローク間の待機時間

