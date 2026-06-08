# AGENTS.md

## Project Purpose

This directory is for automating Gmarket product checks and packaging that automation for non-technical Windows users.

The main script, `gmarket.py`, reads target sellers and product names from `빅스_판매량_20260604.xlsx`, opens Gmarket search pages with Selenium and `undetected_chromedriver`, matches each search result by normalized product title and seller name, collects the visible purchase/sales count, price, and reward/point amount, and writes the collected values back into the same tracking workbook.

The desktop entrypoint, `gmarket_app.py`, provides a simple Tkinter UI for choosing an Excel workbook, saving a daily run time, registering a Windows Task Scheduler job, and running the crawler from the saved config without requiring the end user to know Python.

The current product set is focused on robot vacuums, cordless vacuums, wet/dry cleaning products, and related home appliance items from brands or sellers such as Roborock, Narwal, Ecovacs, Dreame, MOVA, LG, and Samsung.

## Important Files

- `gmarket.py`: main crawler script.
- `gmarket_app.py`: Windows desktop UI and scheduled-run entrypoint. Normal launch opens the UI; `--run` executes the crawler from saved config.
- `빅스_판매량_20260604.xlsx`: main tracking workbook. The crawler reads C열 seller names and D열 Gmarket search product names only, then writes date-based sales counts from F열 onward. Price/reward trend columns start at X/Y and then continue as 2 used columns plus 1 skipped column, for example X/Y, skip Z, AA/AB, skip AC, AD/AE. If today's price/reward headers already exist, the crawler updates those same columns instead of appending another pair.
- `requirements.txt`: Python packages needed for development and packaging.
- `build_exe.ps1`: PyInstaller build script for creating `dist\GmarketCrawler\GmarketCrawler.exe`.
- `installer.iss`: optional Inno Setup script for creating `installer\GmarketCrawlerSetup.exe` after the PyInstaller build.
- `README_배포.md`: Korean distribution and user instructions.
- Other `*.xlsx`: generated or reference Excel output files. Treat these as user data unless the user explicitly asks to regenerate or overwrite them.
- `.vscode/`: local editor settings.

## Current Script Behavior

`gmarket.py` currently:

1. Reads up to the first 10 target rows from `빅스_판매량_20260604.xlsx`, using C열 as `판매자` and D열 as `상품명`. Rows without either value are skipped.
2. Builds a Gmarket search URL for each product name read from the workbook.
3. Launches Chrome through `undetected_chromedriver`.
4. Waits for Gmarket product result containers.
5. Extracts item title, seller text or seller logo alt text, purchase count, price, and reward/point amount.
6. Appends one result row per target product.
7. Opens the selected workbook, writes the result back to the same Excel row, finds or creates the current date column from F열 onward, and saves the purchase count into that cell.
8. Writes price and reward/point values to today's matching price/reward pair when it exists. If today's pair does not exist yet, it creates the next blank pair from X/Y onward. The pattern intentionally skips one column between pairs so a sales-count column can sit between price/reward pairs.

`gmarket_app.py` currently:

1. Stores user config in `%APPDATA%\GmarketCrawler\config.json`.
2. Stores run logs in `%APPDATA%\GmarketCrawler\logs`.
3. Registers or overwrites a daily Windows Task Scheduler task named `GmarketDailyCrawler`.
4. Uses `GmarketCrawler.exe --run` when frozen by PyInstaller, or `python gmarket_app.py --run` during source development.
5. Requires the user to be logged in for scheduled browser automation to run reliably.

## Editing Guidelines

- Keep changes narrow and preserve the existing crawling flow unless the user asks for a redesign.
- Product targets should be managed in the workbook, not hardcoded in Python. Add or remove crawl targets by editing C열 `판매자` and D열 `상품명`.
- Keep the 10-product crawl limit unless the user explicitly asks to expand it.
- Keep the UI simple for non-technical users: Excel file picker, time selection, save/register, run now, and log folder are the intended surface.
- Be careful with Korean text. The current `gmarket.py` appears to contain mojibake/encoding-corrupted Korean strings and may also contain syntax damage around dictionary keys and string literals. Before editing product names, seller names, column names, or output filenames, inspect the file encoding and avoid blind search-and-replace.
- Do not overwrite unrelated `.xlsx` files. The intended output target for this crawler is `빅스_판매량_20260604.xlsx`.
- Keep selectors centralized or clearly documented if Gmarket markup changes.
- Keep browser cleanup behavior intentional. The current script leaves `driver.quit()` commented out so the browser may remain open for inspection.
- Avoid adding unrelated crawling targets or marketplaces to `gmarket.py`; create a separate script if the scope changes beyond Gmarket.

## Runtime Notes

Expected Python packages:

- `openpyxl` for Excel output
- `selenium`
- `undetected_chromedriver`
- `pyinstaller` for packaging

The script depends on a working local Chrome installation. Driver creation first lets `undetected_chromedriver` auto-detect the matching ChromeDriver version; if that fails, it reads the installed local Chrome major version and retries with that `version_main`.

Build the distributable UI with:

```powershell
.\build_exe.ps1
```

## Verification

For code-only edits, first check syntax without launching the browser:

```powershell
python -m py_compile .\gmarket.py
```

For UI/scheduler helper tests:

```powershell
python .\test_gmarket_app.py
```

For crawler behavior, run the script manually from this directory:

```powershell
python .\gmarket.py
```

Because this script opens a real browser and connects to Gmarket, do not run full crawler verification unless the user asks for live crawling or the change requires it.
