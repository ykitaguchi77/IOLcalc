import pyautogui

image_location = pyautogui.locateOnScreen('robot.png')
if image_location is not None:
    left, top, width, height = image_location
    click_x = left +int(width/2)
    click_y = top +int(height/2)
    pyautogui.moveTo(click_x/2, click_y/2)  # マウスを移動
    pyautogui.click()  # クリックを実行
    pyautogui.click()  # クリックを実行

    # pyautogui.click(click_x/2, click_y/2)
   
    print("見つかりました")
else:
    print("画像が見つかりませんでした")