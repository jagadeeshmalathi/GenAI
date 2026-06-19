import pyautogui

pyautogui.FAILSAFE = True

screenshot = pyautogui.screenshot()
screenshot.save("screenshot.png")

print("Screenshot saved as screenshot.png")