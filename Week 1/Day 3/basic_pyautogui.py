import pyautogui
pyautogui.FAILSAFE
pyautogui.PAUSE = 1.0

import pyautogui

print("PyAutoGUI program started")

print("Screen size:", pyautogui.size())
print("Current mouse position:", pyautogui.position())

print("Moving mouse...")
pyautogui.moveTo(500, 300, duration=2)

print("Program completed")