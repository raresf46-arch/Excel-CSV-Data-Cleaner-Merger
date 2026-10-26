# Excel / CSV Data Cleaner & Merger

Turns a folder of messy spreadsheet exports into **one clean, validated Excel report**, in one command.

Built for the common real-world situation where every branch, colleague or system exports data differently,
and someone loses hours each month fixing it by hand.

![Before and after](docs/before_after.png)

## What it fixes automatically

| Problem in the input files | What the script does |
|---|---|
| Different column names (`Client`, `customer`, `Client Name`) | Maps them to one standard set |
| Header not on row 1 (title rows above the table) | Finds the real header row |
| `,` or `;` separated CSV, or XLSX | Detects the format |
| Mixed dates: `2026-01-04`, `04.01.2026`, `04 Jan 2026` | Converts to real dates |
| Mixed numbers: `385.19`, `1.250,50 RON` | Converts to real numbers |
| `ALPHA SRL`, `alpha srl  `, `Alpha SRL` | Normalises to `Alpha SRL` |
| Blank rows, exact duplicates | Removes them |
| Missing client, bad date, negative quantity | Moves the row to an **Issues** sheet with the reason and original row number |

## Output (one Excel file, 4 sheets)

- **Processing Log**: rows read, removed, rejected and kept, plus total value
- **Summary by Client**: sales per client per month, with a grand total
- **Clean Data**: merged, sorted, formatted table with filters and a frozen header
- **Issues**: every rejected row, with its file, row number and problem, so nothing disappears silently

## Run it

```bash
pip install pandas openpyxl
python make_sample_data.py                     # creates 3 messy demo files
python clean_merge.py sample_data -o output/clean_sales_report.xlsx
```

Example run:

```
  read branch_A_export.csv                   42 rows
  read branch_B_vanzari.csv                  35 rows
  read branch_C_report_FINAL_v2.xlsx         32 rows

                     Step     Result
    Rows read (all files)        109
       Blank rows removed          2
Invalid rows (see Issues)          4
       Duplicates removed          2
     Clean rows in report        101
```

## Adapting it to new data

Column names are configured in one dictionary (`COLUMN_ALIASES`) at the top of `clean_merge.py`.
Adding a new source format is usually a one-line change.

## Tech

Python · pandas · openpyxl · argparse
