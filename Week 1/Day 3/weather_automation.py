import os
import time
import pyautogui

pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.3

folder = os.path.dirname(os.path.abspath(__file__))
screenshot_path = os.path.join(folder, "new_york_weather.png")
notepad_save_path = os.path.join(folder, "new_york_weather.txt")

print("Starting in 3 seconds...")
time.sleep(3)

# Open Chrome with New York weather page
pyautogui.hotkey("win", "r")
time.sleep(0.8)
pyautogui.write(
    "chrome https://www.accuweather.com/en/us/new-york/10021/weather-forecast/349727?type=locality&city=new%20york",
    interval=0.02
)
pyautogui.press("enter")
time.sleep(8)  # Wait for page to fully load

# Click inside page body, then select and copy all text
pyautogui.click(x=760, y=400)
time.sleep(0.5)
pyautogui.hotkey("ctrl", "a")
pyautogui.hotkey("ctrl", "c")
time.sleep(1)

# Open Notepad and paste
pyautogui.hotkey("win", "r")
time.sleep(0.8)
pyautogui.write("notepad")
pyautogui.press("enter")
time.sleep(2)
pyautogui.hotkey("ctrl", "v")
time.sleep(1)

# Save the Notepad file
pyautogui.hotkey("ctrl", "s")
time.sleep(1)
pyautogui.write(notepad_save_path, interval=0.02)
pyautogui.press("enter")
time.sleep(1)

# Take screenshot
pyautogui.screenshot(screenshot_path)

print("Done.")
print("Text saved at:", notepad_save_path)
print("Screenshot saved at:", screenshot_path)
