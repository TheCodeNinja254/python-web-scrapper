"""
Hello,

Here is a quick guide before you edit this code.
Project Name: Manitoba Child Care Search — Full Scraper
==========================================
Scrapes all licensed child care facilities from childcaresearch.gov.mb.ca
and writes them to manitoba_childcare.xlsx.

Requirements:
    pip install playwright openpyxl
    python -m playwright install chromium

    Run each of the above, one at a time, preferably

Usage:
    python mb_childcare_scraper.py                  # Full scrape → saves to manitoba_childcare.xlsx
    python mb_childcare_scraper.py --test           # Quick test: scrapes only 10 facilities

Wishing you well,
Fredrick M
"""

import argparse
import time
import re
from datetime import datetime
from pathlib import Path

from playwright.sync_api import sync_playwright
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

OUTPUT_FILE = "manitoba_childcare.xlsx"
BASE_URL = "https://childcaresearch.gov.mb.ca"

def make_browser_context(p):
    return p.chromium.launch(headless=True).new_context(
        user_agent=(
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/124.0.0.0 Safari/537.36"
        )
    )


def intercept_facility_api(page):
    """Return the JSON payload from the facility-list API call."""
    result = {}

    def handle(response):
        url = response.url
        if "/api/" in url.lower() or "facilit" in url.lower():
            ct = response.headers.get("content-type", "")
            if "json" in ct:
                try:
                    result["data"] = response.json()
                    result["url"] = url
                except Exception:
                    pass

    page.on("response", handle)
    return result


# ── Scraping ───────────────────────────────────────────────────────────────────

def get_all_facility_ids(page, context_result):
    """Load the search page and collect every facility ID from the rendered list."""
    page.goto(BASE_URL, wait_until="networkidle", timeout=60000)
    page.wait_for_timeout(6000)  # let JS fully render

    # If we caught a JSON API response, extract IDs from it
    if context_result.get("data"):
        data = context_result["data"]
        if isinstance(data, list):
            ids = [str(f.get("childCareFacilityId") or f.get("id") or "") for f in data]
            return [i for i in ids if i]

    # Fallback: parse IDs from anchor tags in the rendered DOM
    html = page.content()
    ids = re.findall(r'ChildCareFacilityId=(\d+)', html)
    return list(dict.fromkeys(ids))  # deduplicate, preserve order


def parse_facility_page(page, facility_id):
    """Load one facility profile and extract all available fields."""
    url = f"{BASE_URL}/en/Facility/Index?ChildCareFacilityId={facility_id}"
    try:
        page.goto(url, wait_until="networkidle", timeout=30000)
        page.wait_for_timeout(2000)
    except Exception:
        return None

    def txt(selector):
        try:
            return page.locator(selector).first.inner_text(timeout=2000).strip()
        except Exception:
            return ""

    def checked(selector):
        try:
            return "Yes" if page.locator(selector).first.is_visible(timeout=1000) else "No"
        except Exception:
            return "No"

    # Extract page title / facility name from <h4> or meta
    name = txt("h4") or txt("h3") or txt(".facility-name")

    # Meta description often contains name + address
    meta = ""
    try:
        meta = page.locator('meta[name="description"]').get_attribute("content") or ""
    except Exception:
        pass

    # Pull structured fields
    facility_type = txt(".facility-type") or txt('[data-label="Facility Type"] + *')
    business_type = txt(".business-type") or txt('[data-label="Business Type"] + *')
    funding_type = txt(".funding-type") or txt('[data-label="Funding Type"] + *')
    language = txt(".language") or ""
    last_update = txt(".last-update") or txt('[data-label="Last Update"] + *')
    address_raw = txt(".address") or txt('[data-label="Address"] + *') or ""
    region = txt(".region") or ""
    area = txt(".area") or ""
    neighbourhood = txt(".neighbourhood") or ""
    phone = txt(".phone") or txt('[data-label="Phone"] + *') or ""
    email = txt(".email") or txt('[data-label="Email"] + *') or ""
    website = txt(".website") or txt('[data-label="Website"] + *') or ""

    # Vacancies
    def vacancy(label):
        try:
            el = page.locator(f'text="{label}"').first
            sib = el.locator("..").locator("span, td, div").last
            return sib.inner_text(timeout=1000).strip()
        except Exception:
            return ""

    vac_infant = vacancy("0 to 2")
    vac_preschool = vacancy("2 to 6 (Preschool)")
    vac_nursery = vacancy("2 to 6 (Nursery)")
    vac_school = vacancy("6 to 12")

    # Operating times
    weekdays = checked('text="Weekdays"')
    weekends = checked('text="Weekends"')
    evenings = checked('text="Evenings"')
    overnight = checked('text="Overnight"')

    # Employment
    jobs = []
    for role in ["CCA", "ECE II", "ECE III"]:
        try:
            if page.locator(f'text="{role}"').first.is_visible(timeout=500):
                jobs.append(role)
        except Exception:
            pass
    employment = ", ".join(jobs) if jobs else ""

    # Fallback name extraction from meta
    if not name and meta:
        name = meta.split(" is ")[0].strip()

    return {
        "Facility ID": facility_id,
        "Facility Name": name,
        "Facility Type": facility_type,
        "Business Type": business_type,
        "Funding Type": funding_type,
        "Language": language,
        "Address": address_raw,
        "Region": region,
        "Area": area,
        "Neighbourhood": neighbourhood,
        "Phone": phone,
        "Email": email,
        "Website": website,
        "Infant (0–2) Vacancies": vac_infant,
        "Preschool (2–6) Vacancies": vac_preschool,
        "Nursery (2–6) Vacancies": vac_nursery,
        "School Age (6–12) Vacancies": vac_school,
        "Open Weekdays": weekdays,
        "Open Weekends": weekends,
        "Open Evenings": evenings,
        "Open Overnight": overnight,
        "Employment Opportunities": employment,
        "Last Updated": last_update,
        "Profile URL": f"{BASE_URL}/en/Facility/Index?ChildCareFacilityId={facility_id}",
        "Scraped At": datetime.now().strftime("%Y-%m-%d %H:%M"),
    }


# ── Excel output ───────────────────────────────────────────────────────────────

COLUMNS = [
    "Facility ID", "Facility Name", "Facility Type", "Business Type",
    "Funding Type", "Language", "Address", "Region", "Area", "Neighbourhood",
    "Phone", "Email", "Website",
    "Infant (0–2) Vacancies", "Preschool (2–6) Vacancies",
    "Nursery (2–6) Vacancies", "School Age (6–12) Vacancies",
    "Open Weekdays", "Open Weekends", "Open Evenings", "Open Overnight",
    "Employment Opportunities", "Last Updated", "Profile URL", "Scraped At",
]

HEADER_FILL   = PatternFill("solid", start_color="1F4E79")
HEADER_FONT   = Font(name="Arial", bold=True, color="FFFFFF", size=10)
ALT_FILL      = PatternFill("solid", start_color="D6E4F0")
NORMAL_FILL   = PatternFill("solid", start_color="FFFFFF")
BODY_FONT     = Font(name="Arial", size=9)
CENTER        = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT          = Alignment(horizontal="left",   vertical="center", wrap_text=True)
THIN          = Side(style="thin", color="BFBFBF")
BORDER        = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

COL_WIDTHS = {
    "Facility ID": 12, "Facility Name": 35, "Facility Type": 15,
    "Business Type": 15, "Funding Type": 15, "Language": 12,
    "Address": 40, "Region": 18, "Area": 18, "Neighbourhood": 20,
    "Phone": 16, "Email": 28, "Website": 30,
    "Infant (0–2) Vacancies": 18, "Preschool (2–6) Vacancies": 22,
    "Nursery (2–6) Vacancies": 20, "School Age (6–12) Vacancies": 22,
    "Open Weekdays": 14, "Open Weekends": 14,
    "Open Evenings": 14, "Open Overnight": 14,
    "Employment Opportunities": 24, "Last Updated": 16,
    "Profile URL": 55, "Scraped At": 18,
}


def write_excel(facilities, output_path):
    # Load existing file if it exists (for incremental updates)
    if Path(output_path).exists():
        wb = load_workbook(output_path)
        ws = wb["Facilities"] if "Facilities" in wb.sheetnames else wb.active
        # Remove old data rows, keep header
        ws.delete_rows(2, ws.max_row)
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = "Facilities"

        # Instructions sheet
        info = wb.create_sheet("How to Refresh")
        info.column_dimensions["A"].width = 80
        info["A1"] = "HOW TO REFRESH THIS SPREADSHEET"
        info["A1"].font = Font(name="Arial", bold=True, size=14, color="1F4E79")
        steps = [
            "",
            "This file is populated by the Python scraper: mb_childcare_scraper.py",
            "",
            "STEPS TO REFRESH:",
            "  1.  Make sure Python is installed (python.org)",
            "  2.  Install dependencies (run once in terminal):",
            "        pip install playwright openpyxl",
            "        python -m playwright install chromium",
            "  3.  Run the scraper:",
            "        python mb_childcare_scraper.py",
            "  4.  The script overwrites the 'Facilities' sheet with fresh data.",
            "",
            "SCHEDULED REFRESH (Windows Task Scheduler / macOS launchd / Linux cron):",
            "  Add a scheduled task that runs:  python mb_childcare_scraper.py",
            "  e.g. daily at 6 AM to keep the sheet current.",
            " My Assumption is that you are using a Windows Machine, instructions are based on that. "
            "",
            "SOURCE:",
            "  https://childcaresearch.gov.mb.ca/",
            "  Data is sourced from the Manitoba Government Child Care Search portal.",
        ]
        for i, s in enumerate(steps, 2):
            info[f"A{i}"] = s
            info[f"A{i}"].font = Font(name="Arial", size=10)

    # Header row
    ws.row_dimensions[1].height = 30
    for col_idx, col_name in enumerate(COLUMNS, 1):
        cell = ws.cell(row=1, column=col_idx, value=col_name)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = CENTER
        cell.border = BORDER
        ws.column_dimensions[get_column_letter(col_idx)].width = COL_WIDTHS.get(col_name, 16)

    ws.freeze_panes = "A2"

    # Data rows
    for row_idx, fac in enumerate(facilities, 2):
        fill = ALT_FILL if row_idx % 2 == 0 else NORMAL_FILL
        ws.row_dimensions[row_idx].height = 20
        for col_idx, col_name in enumerate(COLUMNS, 1):
            val = fac.get(col_name, "")
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            cell.font = BODY_FONT
            cell.fill = fill
            cell.border = BORDER
            cell.alignment = LEFT

    # Summary row
    summary_row = len(facilities) + 2
    ws.cell(row=summary_row, column=1, value=f"Total facilities: {len(facilities)}")
    ws.cell(row=summary_row, column=1).font = Font(name="Arial", bold=True, size=9)
    ws.cell(row=summary_row, column=len(COLUMNS), value=f"Last scraped: {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    ws.cell(row=summary_row, column=len(COLUMNS)).font = Font(name="Arial", italic=True, size=9)

    wb.save(output_path)
    print(f"\n✅  Saved {len(facilities)} facilities → {output_path}")


# ── Main. This is the project start.

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--test", action="store_true", help="Scrape only 10 facilities for testing")
    parser.add_argument("--output", default=OUTPUT_FILE, help="Output Excel file path")
    args = parser.parse_args()

    print("🔍  Starting Manitoba Child Care scraper...")
    print(f"    Target: {BASE_URL}")
    print(f"    Mode:   {'TEST (10 facilities)' if args.test else 'FULL (all Manitoba)'}\n")

    with sync_playwright() as p:
        context = make_browser_context(p)
        page = context.new_page()
        api_result = intercept_facility_api(page)

        print("📋  Loading search page to collect facility IDs...")
        ids = get_all_facility_ids(page, api_result)
        print(f"    Found {len(ids)} facility IDs")

        if args.test:
            ids = ids[:10]
            print(f"    [TEST MODE] Limiting to {len(ids)} facilities\n")

        facilities = []
        errors = []

        for i, fid in enumerate(ids, 1):
            print(f"  [{i:>4}/{len(ids)}] Scraping facility {fid}...", end="\r")
            data = parse_facility_page(page, fid)
            if data:
                facilities.append(data)
            else:
                errors.append(fid)
            # Polite delay to avoid hammering the server with requests also so we do not get lkocked out.
            time.sleep(0.5)

        context.browser.close()

    print(f"\n✅  Scraped {len(facilities)} facilities  ({len(errors)} errors)")
    if errors:
        print(f"    Failed IDs: {errors[:10]}{'...' if len(errors) > 10 else ''}")

    write_excel(facilities, args.output)


if __name__ == "__main__":
    main()
