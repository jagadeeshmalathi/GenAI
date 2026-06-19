import time
import pyautogui

# Safety settings
pyautogui.FAILSAFE = True
pyautogui.PAUSE = 1.0

print("PyAutoGUI playground starting...")
print("Move your mouse to the TOP-LEFT corner to stop the program.")

# Give yourself time before automation starts
for seconds in range(3, 0, -1):
    print(f"Starting in {seconds}...")
    time.sleep(1)

# Get screen information
screen_width, screen_height = pyautogui.size()
print("Screen size:", screen_width, "x", screen_height)

# Get current mouse position
current_position = pyautogui.position()
print("Current mouse position:", current_position)

# Move mouse to the center of the screen
center_x = screen_width // 2
center_y = screen_height // 2

print("Moving mouse to screen center...")
pyautogui.moveTo(center_x, center_y, duration=2)

# Move the mouse slightly
pyautogui.moveRel(100, 0, duration=1)
pyautogui.moveRel(0, 100, duration=1)
pyautogui.moveRel(-100, 0, duration=1)
pyautogui.moveRel(0, -100, duration=1)

print("Mouse movement test completed.")

print("Click inside Notepad. Typing starts in 5 seconds...")
time.sleep(5)

pyautogui.write("Hello Jagadeesh! This text was typed using PyAutoGUI.", interval=0.05)
pyautogui.press("enter")
pyautogui.write("I am learning Python automation.", interval=0.05)