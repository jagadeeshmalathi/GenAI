import html
import re
from datetime import datetime
from pathlib import Path

from playwright.sync_api import sync_playwright


URL = (
    "https://www.espn.com/soccer/story/_/id/48939282/"
    "2026-fifa-world-cup-fixtures-results-match-schedule-"
    "group-stage-knockout-rounds-bracket"
)

today = datetime.now().date()

# Save screenshot in the same folder as this Python file
screenshot_path = Path(__file__).with_name(
    f"fifa_results_{today.strftime('%Y-%m-%d')}.png"
)

# Example date: Thursday, June 18
date_pattern = re.compile(
    r"^(Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday), "
    r"(January|February|March|April|May|June|July|August|"
    r"September|October|November|December) \d{1,2}:?$"
)

# Match any line containing a score like "2-0" or "1–1"
score_pattern = re.compile(
    r"^.*\b\d+\s*[-–]\s*\d+\b.*$"
)


with sync_playwright() as p:
    browser = p.chromium.launch(
        headless=False,
        slow_mo=200
    )

    page = browser.new_page(
        viewport={"width": 1440, "height": 1000}
    )

    # 1. Navigation
    print("Opening ESPN World Cup results page...")

    page.goto(
        URL,
        wait_until="domcontentloaded",
        timeout=60000
    )

    # 2. Dismiss cookie/consent popup if present
    try:
        accept_btn = page.get_by_role("button", name=re.compile(r"accept|agree|continue|ok", re.IGNORECASE))
        accept_btn.first.click(timeout=5000)
        page.wait_for_timeout(1000)
    except Exception:
        pass  # No popup, continue

    # 3. Wait for article content
    page.get_by_text(
        "2026 FIFA World Cup fixtures, results and features",
        exact=False
    ).first.wait_for(timeout=30000)

    print("Page loaded. Scrolling to load all content...")

    # Scroll to bottom to trigger lazy-loaded content
    for _ in range(8):
        page.keyboard.press("End")
        page.wait_for_timeout(800)

    print("Scroll complete.")

    # 4. Extract page data
    article = page.locator("article")

    if article.count() > 0:
        page_text = article.first.inner_text()
    else:
        page_text = page.locator("body").inner_text()

    lines = [
        line.strip()
        for line in page_text.splitlines()
        if line.strip()
    ]

    completed_matches = []
    current_match_date = None

    for line in lines:
        # Check whether the current line is a date heading
        if date_pattern.match(line):
            date_text = line.rstrip(":")

            parsed_date = datetime.strptime(
                date_text,
                "%A, %B %d"
            )

            current_match_date = parsed_date.replace(
                year=today.year
            ).date()

            continue

        # Keep only completed matches up to today's date
        if (
            current_match_date is not None
            and current_match_date <= today
            and score_pattern.match(line)
        ):
            completed_matches.append(
                (current_match_date, line)
            )

    if not completed_matches:
        browser.close()
        raise RuntimeError(
            "No completed matches found. ESPN page structure may have changed — "
            "check the score_pattern regex or article locator."
        )

    print(
        f"Found {len(completed_matches)} completed matches."
    )

    # 4. Build a clean report for the screenshot
    match_rows = ""

    last_date = None

    for match_date, match_text in completed_matches:
        if match_date != last_date:
            match_rows += f"""
                <h2>{match_date.strftime("%A, %B %d, %Y")}</h2>
            """
            last_date = match_date

        match_rows += f"""
            <div class="match">
                {html.escape(match_text)}
            </div>
        """

    report_html = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="UTF-8">

        <style>
            body {{
                font-family: Arial, sans-serif;
                background: #f4f6f8;
                margin: 0;
                padding: 40px;
                color: #222;
            }}

            .report {{
                max-width: 900px;
                margin: auto;
                background: white;
                padding: 35px;
                border-radius: 12px;
            }}

            h1 {{
                margin-top: 0;
                text-align: center;
            }}

            .generated {{
                text-align: center;
                color: #666;
                margin-bottom: 30px;
            }}

            h2 {{
                margin-top: 28px;
                padding-bottom: 8px;
                border-bottom: 2px solid #ddd;
            }}

            .match {{
                padding: 12px;
                margin: 8px 0;
                background: #f7f7f7;
                border-radius: 6px;
                font-size: 17px;
            }}
        </style>
    </head>

    <body>
        <div class="report">
            <h1>2026 FIFA World Cup Results</h1>

            <div class="generated">
                Completed matches through
                {today.strftime("%B %d, %Y")}
            </div>

            {match_rows}
        </div>
    </body>
    </html>
    """

    report_page = browser.new_page(
        viewport={"width": 1100, "height": 900}
    )

    report_page.set_content(report_html)

    # 5. Screenshot
    report_page.screenshot(
        path=str(screenshot_path),
        full_page=True
    )

    print("Screenshot saved at:")
    print(screenshot_path)

    report_page.wait_for_timeout(3000)

    browser.close()

print("Completed successfully.")