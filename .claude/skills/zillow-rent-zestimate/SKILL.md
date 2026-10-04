---
name: zillow-rent-zestimate
description: Look up Zillow rent estimates for property addresses (including any unit number) with Zillow's Price My Rental tool (zillow.com/rental-manager/price-my-rental) and write them into the Zillow rent column of a rental screening spreadsheet, such as MKE_with_rent_screen_updated.xlsm. Use when the user asks to price rentals, get Zillow rent estimates or Rent Zestimates, fill the Zillow rent column, or run the Zillow rental tool on the addresses in their sheet.
---

# Zillow rent estimates → spreadsheet

Enter each address from the sheet into Zillow's Price My Rental tool and record
the monthly estimate it shows, one unit at a time.

Tool URL: https://www.zillow.com/rental-manager/price-my-rental/

## The workbook

The default workbook is `MKE_with_rent_screen_updated.xlsm` in Google Drive
(file ID `1QJ1MupyjujxZK6QHdEwL0JhrJlLhIUne`). Its layout is:

| Sheet | Header row | Address | City | Zillow rent column | Source URL column |
|---|---|---|---|---|---|
| `Sheet1` | 4 | A | C | **AB** "Zillow Rent Zestimate / Unit ($/mo)" | none |
| `Favorties` (spelled that way) | 7 | A | C | **AC** "Zillow Rent Zestimate / Unit ($/mo)" | AD "Zillow Source" |

- Column A already includes the unit where there is one, for example
  `2025 South 76th Street, Unit 2027`. Never drop the unit. The full address
  sent to Zillow is `<A>, <City>, WI`.
- **Column AA on both sheets is a formula** ("Nick Rent Gap vs 9%" = V − X).
  If the user says "put it in AA", they almost certainly mean the Zillow rent
  column next to it. Write to the Zillow column, and tell them you did. Only
  overwrite AA if they confirm that's what they want (`--column AA --force`).
- The helper scripts find the target column by its header text, so they still
  work if columns move. Pass `--column <letter>` to override.

## Steps

### 1. Get the workbook locally

- If you have a Drive link or ID, download it with the Google Drive connector
  (`download_file_content`). The result is base64; if it's too large to show
  inline, it's saved to a file. Decode it:
  `jq -r .content <saved-result> | base64 -d > MKE.xlsm`
- If you have a local path, use that.
- Install the dependencies once: `pip install -r scripts/requirements.txt`.
  The Playwright path in step 3b also needs `playwright install chrome`, or an
  installed Chrome.

### 2. List the rows that need an estimate

```bash
python scripts/rent_sheet.py MKE.xlsm --sheet Sheet1 list            # blank Zillow cells only
python scripts/rent_sheet.py MKE.xlsm --sheet Sheet1 list --overwrite --rows 5-20
```

The output is JSON: `row`, `address` (full, with unit and city), and `current`.
If the user named particular addresses, filter to those rows.

### 3. Look up each address on Zillow

Zillow blocks most automated traffic, and many cloud sandboxes can't reach
zillow.com at all. Use a real browser on the user's computer. Pick the first
option that's available:

**a. A browser you can drive (preferred):** Claude in Chrome, the desktop
app's built-in browser, or computer use. Read that tool's skill first. Then, for
each row:

1. Open the tool URL.
2. Type the full address, including the unit, into the address box. Choose the
   autocomplete suggestion that matches the street number **and** the unit.
3. Continue through any prompts. If Zillow asks for property details (beds,
   baths, sq ft), use the row's E/F/I columns on Sheet1. For a duplex those
   columns are whole-building totals, so give the per-unit share if Zillow asks
   per unit, and say that you did.
4. Read the **Rent Zestimate** (monthly $). Record the midpoint, not the
   range ends.
5. Write it straight away, so progress is saved even if a later lookup fails:
   ```bash
   python scripts/rent_sheet.py MKE.xlsm --sheet Sheet1 write --row 5 --value 1650 --source "<page url>"
   ```
6. If you hit a "Press & Hold" or other bot check, ask the user to clear it in
   their browser, then carry on. Don't try to get around it.

**b. Batch script (Claude Code on the user's own machine):**

```bash
python scripts/zillow_price_my_rental.py MKE.xlsm --sheet Sheet1             # fills blank rows
python scripts/zillow_price_my_rental.py MKE.xlsm --sheet Favorties --rows 8-12
python scripts/zillow_price_my_rental.py MKE.xlsm --dry-run --limit 3         # trial run, writes nothing
```

The script opens a visible Chrome window with a persistent profile
(`~/.zillow-rent-profile`; pass `--executable-path` to use a specific Chrome/Chromium binary), types each address, reads the estimate, and saves
the workbook after every row. It also appends each lookup to
`zillow_rent_log.csv`. When it can't find an estimate on its own (a bot check,
an extra form step, or a changed page layout), it pauses so the user can finish
the step in the window or type the number. Run it in a terminal the user can
see, or ask the user to run it.

**c. Neither is available** (for example, a cloud session with no linked
browser): don't guess the numbers. Give the user the `list` output and the
step 3b commands to run locally.

### 4. Rules for the numbers

- Record only what Zillow shows. If Zillow shows no estimate for an address,
  leave the cell blank and report it. Never estimate or copy a neighbouring
  value. (The workbook's own note says: "no estimate was invented.")
- Values are monthly dollars **per unit**, written as whole numbers with `$#,##0`
  formatting.
- Don't touch the formula columns (T, W–AA, and AB on Favorties). The scripts
  refuse to overwrite them unless you pass `--force`.

### 5. Return the workbook

- Report a short table: row, address, estimate (or "none shown").
- Before you overwrite the original in Drive, ask the user. Uploading replaces
  their file. With approval, update it through the Drive connector. Otherwise,
  upload a copy named `MKE_with_rent_screen_updated (Zillow rents).xlsm`, or
  hand back the local file.
- openpyxl doesn't store cached formula results. Excel and Google Sheets
  recalculate on open, so the W–AA columns refill automatically.
