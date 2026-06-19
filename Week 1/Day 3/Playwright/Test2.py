from pathlib import Path
from urllib.parse import quote

from playwright.sync_api import expect, sync_playwright


html = """
<html>
<body>
    <h1>Playwright Test Page</h1>

    <label for="name">Enter Name:</label>
    <input id="name" type="text">

    <button onclick="showResult()">Submit</button>

    <p id="result"></p>

    <script>
        function showResult() {
            const name = document.getElementById("name").value;

            setTimeout(() => {
                document.getElementById("result").textContent =
                    "Hello, " + name;
            }, 2000);
        }
    </script>
</body>
</html>
"""

page_url = "data:text/html;charset=utf-8," + quote(html)

screenshot_path = Path(__file__).with_name(
    "playwright_screenshot.png"
)

with sync_playwright() as p:
    browser = p.chromium.launch(
        headless=False,
        slow_mo=500
    )

    page = browser.new_page()

    # 1. Navigation
    page.goto(page_url)
    print("Navigation completed")

    # 2. Typing / Filling form
    page.get_by_label("Enter Name:").fill("Jagadeesh")
    print("Form filled")

    # 3. Clicking
    page.get_by_role("button", name="Submit").click()
    print("Button clicked")

    # 4. Waiting
    result = page.locator("#result")

    expect(result).to_have_text(
        "Hello, Jagadeesh",
        timeout=5000
    )

    print("Waiting completed")

    # 5. Extracting data
    extracted_data = result.inner_text()
    print("Extracted Data:", extracted_data)

    # 6. Screenshot
    page.screenshot(
        path=str(screenshot_path),
        full_page=True
    )

    print("Screenshot saved:", screenshot_path)

    page.wait_for_timeout(3000)
    browser.close()