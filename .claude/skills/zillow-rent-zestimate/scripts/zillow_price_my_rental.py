#!/usr/bin/env python3
"""Batch-fill Zillow rent estimates by driving Zillow's Price My Rental tool.

For each row that needs one, this opens
https://www.zillow.com/rental-manager/price-my-rental/ in a real (headed)
browser, enters the full address including any unit, reads the Rent
Zestimate it shows, and writes it into the workbook (saving after every
row so nothing is lost if it stops).

Zillow uses bot protection ("Press & Hold" challenges). The browser runs
headed with a persistent profile so you can clear a challenge or sign in
once; when the script can't find an estimate on its own it pauses and
asks you to finish the step in the browser window (--assist, on by default).

Usage:
  python zillow_price_my_rental.py MKE.xlsm                 # Sheet1, blank rows only
  python zillow_price_my_rental.py MKE.xlsm --sheet Favorties
  python zillow_price_my_rental.py MKE.xlsm --rows 5-12 --overwrite
  python zillow_price_my_rental.py MKE.xlsm --dry-run       # read only, write nothing
"""
import argparse
import csv
import os
import random
import re
import sys
import time
from pathlib import Path

from playwright.sync_api import TimeoutError as PWTimeout
from playwright.sync_api import sync_playwright

sys.path.insert(0, str(Path(__file__).resolve().parent))
import rent_sheet  # noqa: E402

TOOL_URL = "https://www.zillow.com/rental-manager/price-my-rental/"
PROFILE_DIR = Path.home() / ".zillow-rent-profile"
MONEY = r"\$\s?(\d{1,2},?\d{3}|\d{3})"
MIN_RENT, MAX_RENT = 300, 15000

ADDRESS_INPUT_SELECTORS = [
    'input[placeholder*="address" i]',
    'input[aria-label*="address" i]',
    'input[role="combobox"]',
    'input[type="search"]',
    'input[type="text"]',
]
SUBMIT_BUTTON_TEXT = re.compile(
    r"get (started|estimate|rent estimate|my estimate)|price my rental|see (rent|estimate)|continue|next|search|submit",
    re.I,
)
CHALLENGE_TEXT = re.compile(r"press\s*&\s*hold|are you a human|verify you are|access to this page has been denied", re.I)


def human_pause(a=0.6, b=1.6):
    time.sleep(random.uniform(a, b))


def body_text(page):
    try:
        return page.inner_text("body", timeout=5000)
    except PWTimeout:
        return ""


def extract_rent(text):
    """Pull the Rent Zestimate from page text. Returns int or None."""
    # Prefer an amount that follows the words "Rent Zestimate" / "estimated rent".
    for label in (r"rent\s*zestimate", r"estimated\s+(monthly\s+)?rent", r"rental\s+estimate"):
        m = re.search(label + r"[\s\S]{0,200}?" + MONEY, text, re.I)
        if m:
            val = int(m.group(m.lastindex).replace(",", ""))
            if MIN_RENT <= val <= MAX_RENT:
                return val
    # Otherwise the first "$1,650/mo"-style amount on the page.
    for m in re.finditer(MONEY + r"\s*(/\s*mo|per month|a month)", text, re.I):
        val = int(m.group(1).replace(",", ""))
        if MIN_RENT <= val <= MAX_RENT:
            return val
    return None


def wait_for_human(msg):
    print(f"\n>>> {msg}")
    return input(">>> Press Enter when done, type a rent (e.g. 1650) to record it yourself, or 's' to skip: ").strip()


def clear_challenge(page, assist):
    if CHALLENGE_TEXT.search(body_text(page)):
        if not assist:
            raise RuntimeError("Zillow bot challenge shown; rerun with --assist")
        wait_for_human("Zillow is showing a bot check. Complete it in the browser window.")


def find_address_input(page):
    for sel in ADDRESS_INPUT_SELECTORS:
        loc = page.locator(sel).filter(visible=True)
        if loc.count():
            return loc.first
    return None


def pick_suggestion(page, address):
    """Choose the autocomplete option that best matches the address (unit included)."""
    opts = page.locator('[role="option"], [role="listbox"] li').filter(visible=True)
    try:
        opts.first.wait_for(timeout=6000)
    except PWTimeout:
        return False
    unit = re.search(r"(?:unit|apt|#)\s*([\w-]+)", address, re.I)
    number = address.split()[0]
    best = None
    for i in range(min(opts.count(), 8)):
        t = opts.nth(i).inner_text()
        if number not in t:
            continue
        if unit and unit.group(1).lower() in t.lower():
            best = opts.nth(i)
            break
        best = best or opts.nth(i)
    (best or opts.first).click()
    return True


def click_submit(page):
    btn = page.get_by_role("button", name=SUBMIT_BUTTON_TEXT).filter(visible=True)
    if btn.count():
        btn.first.click()
        return True
    return False


def lookup(page, address, assist, wait_s):
    page.goto(TOOL_URL, wait_until="domcontentloaded")
    human_pause(1.5, 3)
    clear_challenge(page, assist)

    box = find_address_input(page)
    if box is None:
        if not assist:
            return None, page.url
        page.wait_for_timeout(500)
        ans = wait_for_human(f"Couldn't find the address box. Enter this address and get the estimate:\n    {address}")
        return _answer(ans, page)

    box.click()
    box.fill("")
    box.type(address, delay=random.randint(40, 90))
    human_pause()
    if not pick_suggestion(page, address):
        box.press("Enter")
    human_pause()
    click_submit(page)

    # The tool can take a few screens; keep pressing obvious "continue" buttons
    # and polling for an amount until the timeout.
    deadline = time.time() + wait_s
    while time.time() < deadline:
        clear_challenge(page, assist)
        rent = extract_rent(body_text(page))
        if rent:
            return rent, page.url
        click_submit(page)
        page.wait_for_timeout(1500)

    if assist:
        ans = wait_for_human(
            f"No estimate found yet for:\n    {address}\n"
            "    Finish any remaining steps (property details, unit, etc.) in the browser."
        )
        return _answer(ans, page)
    return None, page.url


def _answer(ans, page):
    if ans.lower() == "s":
        return None, page.url
    if ans:
        try:
            return int(float(ans.replace("$", "").replace(",", ""))), page.url
        except ValueError:
            pass
    return extract_rent(body_text(page)), page.url


def main():
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("workbook")
    p.add_argument("--sheet", default="Sheet1")
    p.add_argument("--column", help="Target column letter; default is the 'Zillow Rent Zestimate' column")
    p.add_argument("--force", action="store_true", help="Allow writing into a formula column")
    p.add_argument("--rows", help="Only these sheet rows, e.g. 5-20,31")
    p.add_argument("--overwrite", action="store_true", help="Re-check rows that already have a value")
    p.add_argument("--limit", type=int, help="Stop after this many lookups")
    p.add_argument("--out", help="Save to this file instead of in place")
    p.add_argument("--dry-run", action="store_true", help="Look up rents but don't write the workbook")
    p.add_argument("--no-assist", dest="assist", action="store_false", help="Never pause for manual help")
    p.add_argument("--wait", type=int, default=25, help="Seconds to wait for an estimate per address")
    p.add_argument("--channel", default="chrome", help="Browser channel: chrome (installed Chrome), msedge, or '' for bundled Chromium")
    p.add_argument("--executable-path", help="Path to a Chrome/Chromium binary (overrides --channel)")
    p.add_argument("--profile", default=str(PROFILE_DIR), help="Persistent browser profile directory")
    p.add_argument("--log", default="zillow_rent_log.csv")
    args = p.parse_args()

    wb = rent_sheet.load(args.workbook)
    ws = wb[args.sheet]
    hr = rent_sheet.find_header_row(ws)
    target, source_col, city_col = rent_sheet.resolve_columns(ws, hr, args.column, args.force)
    todo = list(rent_sheet.iter_rows(ws, hr, target, city_col, args.overwrite, rent_sheet.parse_rows(args.rows)))
    if args.limit:
        todo = todo[: args.limit]
    print(f"{len(todo)} row(s) to look up on '{ws.title}', writing to column {rent_sheet.get_column_letter(target)}")
    if not todo:
        return

    out_path = args.out or args.workbook
    new_log = not os.path.exists(args.log)
    with sync_playwright() as pw, open(args.log, "a", newline="") as logf:
        log = csv.writer(logf)
        if new_log:
            log.writerow(["sheet", "row", "address", "rent", "source_url", "time"])
        ctx = pw.chromium.launch_persistent_context(
            args.profile,
            channel=None if args.executable_path else (args.channel or None),
            executable_path=args.executable_path,
            headless=False,
            viewport={"width": 1280, "height": 900},
        )
        page = ctx.pages[0] if ctx.pages else ctx.new_page()
        done = 0
        try:
            for item in todo:
                print(f"\n[{item['row']}] {item['address']}")
                try:
                    rent, url = lookup(page, item["address"], args.assist, args.wait)
                except Exception as e:  # keep going on one bad address
                    print(f"    error: {e}")
                    rent, url = None, page.url
                print(f"    -> {'$' + format(rent, ',') if rent else 'no estimate'}")
                log.writerow([ws.title, item["row"], item["address"], rent or "", url, time.strftime("%Y-%m-%d %H:%M")])
                logf.flush()
                if rent and not args.dry_run:
                    rent_sheet.write_value(ws, item["row"], target, source_col, rent, url)
                    wb.save(out_path)
                    done += 1
                human_pause(2, 5)
        except KeyboardInterrupt:
            print("\nStopped by user.")
        finally:
            ctx.close()
    print(f"\nWrote {done} estimate(s) to {out_path}. Log: {args.log}")


if __name__ == "__main__":
    main()
