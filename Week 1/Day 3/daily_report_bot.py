"""
daily_report_bot.py

A simple PyAutoGUI bot that automates a daily status report:
1. Opens Chrome and loads a public website.
2. Copies the key piece of info from the page (via select-all + copy).
3. Opens Excel and creates a new row with date/time, the fetched data,
   and a short comment.
4. Saves the Excel file with today's date in the filename.
5. Takes a screenshot of the final sheet and saves it.

NOTE: PyAutoGUI controls your REAL mouse and keyboard. Before running:
- Close other windows / save your work, since the bot will be clicking
  and typing on whatever is on screen.
- Don't touch the mouse/keyboard while it runs.
- You may need to adjust the sleep() values below if your machine is
  slower/faster, or if apps take longer to open.

Install dependencies first:
    pip install pyautogui pyperclip
"""

import time
import datetime
import os
import subprocess
import sys
import platform

import pyautogui
import pyperclip

# Safety: moving the mouse to a screen corner will raise FailSafeException
# and stop the bot if something goes wrong.
pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.5  # small delay after every PyAutoGUI call

# ---------------------------------------------------------------------------
# CONFIG — edit these for your machine / preference
# ---------------------------------------------------------------------------
WEBSITE_URL = "https://www.timeanddate.com/weather/"   # any public site works
COMMENT = "Good for outdoor activities"

IS_MAC = platform.system() == "Darwin"
IS_WINDOWS = platform.system() == "Windows"

CMD_KEY = "command" if IS_MAC else "ctrl"


def wait(seconds, reason=""):
    """Simple wrapper so wait times are easy to find/tune, with a printed reason."""
    if reason:
        print(f"Waiting {seconds}s — {reason}")
    time.sleep(seconds)


# ---------------------------------------------------------------------------
# STEP 1: Open Chrome and navigate to a public website
# ---------------------------------------------------------------------------
def open_chrome_and_get_data():
    print("Opening Chrome...")

    if IS_MAC:
        subprocess.Popen(["open", "-a", "Google Chrome"])
    elif IS_WINDOWS:
        os.startfile("chrome")  # relies on Chrome being on PATH / default handler
    else:
        subprocess.Popen(["google-chrome"])

    wait(3, "letting Chrome fully open")

    # Open a new tab and go to the address bar, just in case a stale window
    # came to focus instead of a fresh window.
    pyautogui.hotkey(CMD_KEY, "t")
    wait(1, "new tab opening")

    pyautogui.hotkey(CMD_KEY, "l")  # focus address bar (Cmd/Ctrl+L works in Chrome)
    wait(0.5)
    pyautogui.typewrite(WEBSITE_URL, interval=0.02)
    pyautogui.press("enter")

    wait(4, "letting the page load")

    # Select all visible text on the page and copy it.
    # This is far more reliable with PyAutoGUI than clicking an exact pixel,
    # since page layouts shift between runs.
    pyautogui.hotkey(CMD_KEY, "a")
    wait(0.5)
    pyautogui.hotkey(CMD_KEY, "c")
    wait(1, "clipboard copy to register")

    page_text = pyperclip.paste()

    # Pull out the actual data point (e.g. a temperature) instead of
    # whatever junk line happens to come first on the page.
    fetched_data = extract_snippet(page_text)
    print(f"Fetched snippet: {fetched_data}")
    return fetched_data


# Lines that are almost never the content we want — skip-links, cookie
# banners, nav menus, etc. Add to this list if you see other junk show up.
JUNK_PATTERNS = [
    "skip to main content",
    "accessibility help",
    "cookie",
    "subscribe",
    "sign in",
    "log in",
    "advertisement",
    "menu",
    "navigation",
    "search",
]


def is_junk_line(line):
    lower = line.lower()
    return any(pattern in lower for pattern in JUNK_PATTERNS)


def extract_snippet(text, max_len=120):
    """
    Try to find the actual data point on the page (e.g. a temperature
    like "72°F" or "22°C"), skipping over nav/boilerplate lines.
    Falls back to the first clean non-junk line if no temperature-style
    value is found, so this still works on news/stock sites.
    """
    import re

    lines = [line.strip() for line in text.splitlines() if line.strip()]
    clean_lines = [line for line in lines if not is_junk_line(line)]

    # 1) Look for a temperature-shaped value first (works great for
    #    weather sites): e.g. "72°F", "22°C", "-3°"
    temp_pattern = re.compile(r"-?\d{1,3}\s*°\s*[CF]?")
    for line in clean_lines:
        match = temp_pattern.search(line)
        if match:
            return match.group().strip()

    # 2) Otherwise, fall back to the first reasonably-sized clean line
    #    (useful for headline/stock-price sites).
    for line in clean_lines:
        if 5 < len(line) <= max_len:
            return line

    # 3) Last resort: just truncate whatever we got
    return text.strip()[:max_len] if text.strip() else "No data found"


# ---------------------------------------------------------------------------
# STEP 2: Open Excel and write the report row
# ---------------------------------------------------------------------------
def open_excel():
    print("Opening Excel...")

    if IS_MAC:
        subprocess.Popen(["open", "-a", "Microsoft Excel"])
    elif IS_WINDOWS:
        os.startfile("excel")
    else:
        # Linux fallback — LibreOffice Calc
        subprocess.Popen(["libreoffice", "--calc"])

    wait(6, "letting Excel fully launch (Excel is often slow to open)")

    # Make sure a blank workbook is active. If Excel opened to a start
    # screen instead of a blank sheet, Cmd/Ctrl+N creates a new workbook.
    pyautogui.hotkey(CMD_KEY, "n")
    wait(3, "letting the new workbook open")


def write_report_row(fetched_data):
    print("Writing report row...")

    now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # --- Type the header row (A1:C1) ---
    pyautogui.typewrite("Date & Time", interval=0.02)
    pyautogui.press("tab")
    pyautogui.typewrite("Fetched Data", interval=0.02)
    pyautogui.press("tab")
    pyautogui.typewrite("Comment", interval=0.02)
    pyautogui.press("enter")

    # Move back to column A of the next row before typing the data row
    pyautogui.hotkey(CMD_KEY, "home") if IS_MAC else pyautogui.press("home")
    wait(0.3)

    # --- Type the data row (A2:C2) ---
    pyautogui.typewrite(now_str, interval=0.02)
    pyautogui.press("tab")
    pyautogui.typewrite(fetched_data, interval=0.02)
    pyautogui.press("tab")
    pyautogui.typewrite(COMMENT, interval=0.02)
    pyautogui.press("enter")

    wait(0.5)

    # Everything below makes the table look like a clean report instead
    # of plain typed text: bold + colored header, borders around the
    # whole table, and columns wide enough to actually show the data.
    format_report_table()


def format_report_table():
    """
    Selects A1:C2 and applies formatting using only keyboard shortcuts —
    no clicking needed, so this stays reliable across screen sizes.
    """
    print("Formatting table (bold header, borders, fill, column width)...")

    # --- Step 1: Select the whole table range A1:C2 ---
    # Go to cell A1 first (Ctrl/Cmd+Home jumps to the very start of the sheet)
    pyautogui.hotkey(CMD_KEY, "home")
    wait(0.3)
    # Shift+Right twice extends selection to column C, Shift+Down extends to row 2
    pyautogui.keyDown("shift")
    pyautogui.press("right")
    pyautogui.press("right")
    pyautogui.press("down")
    pyautogui.keyUp("shift")
    wait(0.3)

    # --- Step 2: Add borders around every cell in the selection ---
    # Excel shortcut: Ctrl/Cmd+Shift+& applies an outline+gridlines border
    pyautogui.hotkey(CMD_KEY, "shift", "ampersand") if IS_MAC else pyautogui.hotkey("ctrl", "shift", "7")
    wait(0.3)

    # --- Step 3: Bold + fill color just the header row (A1:C1) ---
    pyautogui.hotkey(CMD_KEY, "home")
    wait(0.2)
    pyautogui.keyDown("shift")
    pyautogui.press("right")
    pyautogui.press("right")
    pyautogui.keyUp("shift")
    wait(0.2)

    pyautogui.hotkey(CMD_KEY, "b")  # bold the header text
    wait(0.2)

    # Fill color: Excel's Fill Color button isn't reachable by a default
    # keyboard shortcut, so we use the Ribbon's Alt-key access keys on
    # Windows. On Mac, there's no reliable keyboard-only fill shortcut,
    # so we skip fill there and rely on bold + borders for contrast.
    if IS_WINDOWS:
        pyautogui.press("alt")
        wait(0.2)
        pyautogui.typewrite("h", interval=0.05)   # Home tab
        wait(0.2)
        pyautogui.typewrite("h", interval=0.05)   # Fill Color button
        wait(0.3)
        pyautogui.press("down")   # move into the color palette
        pyautogui.press("enter")  # pick a color swatch
        wait(0.3)

    # --- Step 4: Auto-fit all three columns so no text is cut off ---
    pyautogui.hotkey(CMD_KEY, "a")  # select entire sheet
    wait(0.2)
    if IS_WINDOWS:
        pyautogui.press("alt")
        pyautogui.typewrite("h", interval=0.05)  # Home tab
        wait(0.2)
        pyautogui.typewrite("o", interval=0.05)  # Format menu
        wait(0.2)
        pyautogui.typewrite("i", interval=0.05)  # AutoFit Column Width
        wait(0.3)
    else:
        # On Mac Excel, double-clicking a column border auto-fits it;
        # there's no clean keyboard-only equivalent, so we widen columns
        # manually instead via Format > Column Width using the menu bar
        # shortcut (Option key opens menu mnemonics in some versions).
        # Simplest reliable fallback: just leave default width — the
        # borders/bold still make it look like a clean table.
        pass

    # Click back on A1 equivalent (keyboard) so the next screenshot
    # doesn't show a big blue selection highlight over everything
    pyautogui.hotkey(CMD_KEY, "home")
    wait(0.3)


# ---------------------------------------------------------------------------
# STEP 3: Save the Excel file with today's date in the filename
# ---------------------------------------------------------------------------
def save_excel_file():
    today_str = datetime.date.today().strftime("%Y-%m-%d")
    filename = f"daily_report_{today_str}"

    print(f"Saving as {filename}.xlsx ...")

    pyautogui.hotkey(CMD_KEY, "s")
    wait(1.5, "letting Save dialog open")

    # Type the full filename (without extension; Excel adds .xlsx)
    pyautogui.typewrite(filename, interval=0.02)
    wait(0.5)
    pyautogui.press("enter")
    wait(2, "letting the save complete / dismiss format dialogs")

    # Some Excel versions show an extra "Keep format" confirmation dialog
    pyautogui.press("enter")
    wait(1)

    return filename


# ---------------------------------------------------------------------------
# STEP 4: Screenshot the final sheet
# ---------------------------------------------------------------------------
def take_screenshot(filename_base):
    screenshot_name = f"{filename_base}_screenshot.png"
    print(f"Taking screenshot: {screenshot_name}")
    wait(1, "settling before screenshot")
    screenshot = pyautogui.screenshot()
    screenshot.save(screenshot_name)
    print(f"Screenshot saved to {os.path.abspath(screenshot_name)}")


# ---------------------------------------------------------------------------
# MAIN
# ---------------------------------------------------------------------------
def main():
    print("Starting daily report bot. You have 5 seconds to switch away "
          "if you need to (e.g. close other apps) ...")
    wait(5)

    fetched_data = open_chrome_and_get_data()

    open_excel()
    write_report_row(fetched_data)
    filename_base = save_excel_file()
    take_screenshot(filename_base)

    print("Done! Daily report created and screenshot saved.")


if __name__ == "__main__":
    main()