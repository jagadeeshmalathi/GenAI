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

screenshot_path = Path(__file__).with_name(
    f"fifa_results_{today.strftime('%Y-%m-%d')}.png"
)

debug_path = Path(__file__).with_name("espn_extracted_text.txt")


# Matches headings such as:
# Thursday, June 18
# Friday, June 19:
date_pattern = re.compile(
    r"^(Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday),\s+"
    r"(January|February|March|April|May|June|July|August|"
    r"September|October|November|December)\s+"
    r"(\d{1,2}):?$"
)


# Matches completed games such as:
# Group A: Mexico 2-0 South Africa (Mexico City)
score_pattern = re.compile(
    r"^Group\s+([A-L]):\s+"
    r"(.+?)\s+"
    r"(\d+)\s*[-–—]\s*(\d+)\s+"
    r"(.+?)"
    r"(?:\s+\(([^)]*)\))?$"
)


with sync_playwright() as p:
    browser = p.chromium.launch(
        headless=False,
        slow_mo=100
    )

    page = browser.new_page(
        viewport={
            "width": 1440,
            "height": 1000
        }
    )

    # 1. Navigation
    print("Opening ESPN World Cup results page...")

    page.goto(
        URL,
        wait_until="domcontentloaded",
        timeout=60000
    )

    # 2. Waiting
    page.get_by_text(
        "2026 FIFA World Cup fixtures, results and features",
        exact=True
    ).last.wait_for(timeout=30000)

    page.wait_for_timeout(3000)

    print("Page loaded.")

    # 3. Extract the full visible page text
    page_text = page.locator("body").inner_text()

    section_title = (
        "2026 FIFA World Cup fixtures, results and features"
    )

    end_title = "What is the 2026 FIFA World Cup format?"

    start_position = page_text.find(section_title)
    end_position = page_text.find(end_title)

    if start_position == -1:
        browser.close()
        raise RuntimeError(
            "Could not find the World Cup fixtures section."
        )

    if end_position == -1:
        fixtures_text = page_text[start_position:]
    else:
        fixtures_text = page_text[
            start_position:end_position
        ]

    # Save extracted text to help with debugging
    debug_path.write_text(
        fixtures_text,
        encoding="utf-8"
    )

    lines = [
        line.strip()
        for line in fixtures_text.splitlines()
        if line.strip()
    ]

    completed_matches = []
    current_match_date = None

    for line in lines:
        date_match = date_pattern.match(line)

        if date_match:
            date_without_year = line.rstrip(":")

            parsed_date = datetime.strptime(
                date_without_year,
                "%A, %B %d"
            )

            current_match_date = parsed_date.replace(
                year=2026
            ).date()

            continue

        score_match = score_pattern.match(line)

        if (
            score_match
            and current_match_date is not None
            and current_match_date <= today
        ):
            group = score_match.group(1)
            team_one = score_match.group(2)
            score_one = score_match.group(3)
            score_two = score_match.group(4)
            team_two = score_match.group(5)
            location = score_match.group(6)

            match_text = (
                f"Group {group}: "
                f"{team_one} {score_one}-{score_two} {team_two}"
            )

            if location:
                match_text += f" ({location})"

            completed_matches.append(
                {
                    "date": current_match_date,
                    "match": match_text
                }
            )

    if not completed_matches:
        browser.close()

        raise RuntimeError(
            "No completed matches were found. "
            "Check espn_extracted_text.txt to see "
            "the text Playwright extracted."
        )

    print(
        f"Found {len(completed_matches)} completed matches."
    )

    for match in completed_matches:
        print(
            match["date"],
            "-",
            match["match"]
        )

    # 4. Create a clean report
    report_rows = ""
    previous_date = None

    for match in completed_matches:
        match_date = match["date"]
        match_text = match["match"]

        if match_date != previous_date:
            report_rows += f"""
                <h2>
                    {match_date.strftime("%A, %B %d, %Y")}
                </h2>
            """

            previous_date = match_date

        report_rows += f"""
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
                background: #eeeeee;
                margin: 0;
                padding: 35px;
                color: #222222;
            }}

            .report {{
                max-width: 900px;
                margin: auto;
                padding: 35px;
                background: white;
            }}

            h1 {{
                text-align: center;
                margin-top: 0;
            }}

            .summary {{
                text-align: center;
                margin-bottom: 30px;
                color: #555555;
            }}

            h2 {{
                margin-top: 28px;
                padding-bottom: 8px;
                border-bottom: 2px solid #dddddd;
            }}

            .match {{
                margin: 8px 0;
                padding: 12px;
                background: #f5f5f5;
                font-size: 17px;
            }}
        </style>
    </head>

    <body>
        <div class="report">
            <h1>2026 FIFA World Cup Results</h1>

            <div class="summary">
                Completed matches through
                {today.strftime("%B %d, %Y")}
                <br>
                Total completed matches:
                {len(completed_matches)}
            </div>

            {report_rows}
        </div>
    </body>
    </html>
    """

    # 5. Display the report in a new page
    report_page = browser.new_page(
        viewport={
            "width": 1100,
            "height": 900
        }
    )

    report_page.set_content(report_html)

    # 6. Screenshot
    report_page.screenshot(
        path=str(screenshot_path),
        full_page=True
    )

    print("\nScreenshot saved at:")
    print(screenshot_path)

    print("\nExtracted ESPN text saved at:")
    print(debug_path)

    report_page.wait_for_timeout(3000)

    browser.close()

print("\nCompleted successfully.")