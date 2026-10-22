# Enterprise Emissions Analysis Tool

Upload one or many Excel/CSV files → get the Chapter-3 statistics, tables, charts and written findings, recalculated from the raw rows.

## Easiest: Google Colab (no installation, no web page)
1. Open https://colab.research.google.com → File → Upload notebook → choose `Emissions_Analysis_Colab.ipynb`.
2. Run **Cell 1** and upload `emissions_tool.zip`. Run **Cell 2** and upload your Excel files - all tables, statistics, findings and charts appear below the cell.
3. Run **Cell 3** to download `results.zip` (Excel with tables + charts, chart images, key findings text).
(Files go to Google's servers - do not use for confidential data.)

Same thing without Colab: `python batch.py file1.xlsx file2.xlsx --out results`

## Optional: web-page version on your own PC (Windows)
1. Install Python 3.10+ from https://www.python.org/downloads/ (tick **"Add Python to PATH"** in the installer).
2. Unzip this folder anywhere, e.g. `C:\emissions_tool`.
3. Double-click **`run_tool.bat`**. The first run installs the libraries (needs internet, ~2 minutes); the tool then opens in your browser at http://localhost:8501.
   Next times it starts in a few seconds.

Mac/Linux: run `bash run_tool.sh`.

To let colleagues on the same office network use it: run it on one office PC and share `http://<that-PC's-IP>:8501`
(the .bat already allows this). Data never leaves that PC.

## Using it
1. Drag files into the sidebar (or tick "Use the example workbook").
2. Open **① Check how columns were recognised** if anything looks wrong - columns are matched by header name, not position.
3. Choose rules in the sidebar (leave out Large, treatment of N/A, outlier rule).
4. Read the tabs; press **Download full results** for an Excel file with every table and charts.

## What it calculates (all from the raw Scope 1/2/3 values)
Mean, median, sample standard deviation, min, max, quartiles/IQR, totals · scope totals, shares, means, medians, maxima and
"dominant in X of N" · the same by Small/Medium/Large and by sector, location, process · ratios between sizes · rankings,
share of total, outliers (IQR rule) · solar offsets / net footprint (if solar columns exist) · automatic findings text.

## Rules you should know
* Totals are recomputed (Scope 1 + 2 + 3); the workbook's own totals are only used to cross-check.
* **N/A / blank / "-"** count as 0 by default, because that is how the report's statistics were calculated. Switch in the sidebar to leave them out instead.
* Obvious typos such as `4.51.` are repaired **and logged** (Data quality tab). Unreadable values are treated as missing and logged as errors.
* Sample standard deviation (n−1). Quartiles interpolate linearly (Excel QUARTILE.INC).
* If a combined sheet and the separate sector sheets are both selected, exact duplicates are removed; the same enterprise spelled differently in two sheets cannot be detected automatically.

## Tests (folder `tests/`)
`verify_report.py` (engine vs numbers printed in Chapter 3) · `verify_workbook.py` (engine vs independent recomputation, all 5 sectors) ·
`test_robustness.py` (shuffled columns, header on another row, missing scope, 4th scope, duplicates, bad values, CSV, 200 files).
