"""
Excel / CSV Data Cleaner & Merger
---------------------------------
Merges messy spreadsheet exports (CSV with , or ; separators, XLSX with title rows)
into ONE clean, validated Excel report.

What it handles automatically:
  * different column names for the same field  (Client / customer / Client Name ...)
  * header rows that don't start on row 1
  * mixed date formats   (2026-01-04, 04.01.2026, 04 Jan 2026)
  * mixed number formats (1250.50, "1.250,50 RON")
  * inconsistent text    ("ALPHA SRL", "alpha srl  ", "Alpha SRL")
  * blank rows and exact duplicates
  * invalid rows (missing client, bad date, negative quantity) -> moved to an Issues sheet

Usage:
    python clean_merge.py sample_data -o output/clean_sales_report.xlsx
"""
from __future__ import annotations

import argparse
import re
from pathlib import Path

import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# --- Column mapping: every known alias -> standard column name -----------------
COLUMN_ALIASES = {
    "date": ["date", "data", "order date", "invoice date"],
    "client": ["client", "customer", "client name", "customer name"],
    "product": ["product", "item", "product name", "article"],
    "qty": ["qty", "quantity", "units", "pcs"],
    "unit_price": ["unit price", "price", "price/unit", "unit_price"],
}
ALIAS_LOOKUP = {alias: std for std, aliases in COLUMN_ALIASES.items() for alias in aliases}
REQUIRED = list(COLUMN_ALIASES)


# --- Reading ---------------------------------------------------------------------
def _find_header_row(raw: pd.DataFrame) -> int:
    """Return the index of the first row that contains at least 3 known column names."""
    for i, row in raw.iterrows():
        cells = [str(c).strip().lower() for c in row if pd.notna(c)]
        if sum(c in ALIAS_LOOKUP for c in cells) >= 3:
            return i
    raise ValueError("No recognisable header row found")


def read_any(path: Path) -> pd.DataFrame:
    """Read a CSV (auto-detect separator) or XLSX, locating the real header row."""
    if path.suffix.lower() in {".xlsx", ".xls"}:
        raw = pd.read_excel(path, header=None, dtype=str)
    else:
        raw = pd.read_csv(path, header=None, dtype=str, sep=None, engine="python")
    h = _find_header_row(raw)
    df = raw.iloc[h + 1:].copy()
    df.columns = [str(c).strip().lower() for c in raw.iloc[h]]
    df = df.rename(columns=ALIAS_LOOKUP)
    missing = [c for c in REQUIRED if c not in df.columns]
    if missing:
        raise ValueError(f"{path.name}: missing columns {missing}")
    df = df[REQUIRED]
    df["source_file"] = path.name
    df["source_row"] = df.index + 1  # 1-based row number in the original file
    return df


# --- Cleaning helpers ------------------------------------------------------------
def parse_number(value) -> float | None:
    """'1.250,50 RON' -> 1250.5 ; '385.19' -> 385.19 ; '' -> None"""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    s = re.sub(r"[^\d,.\-]", "", str(value))
    if not s:
        return None
    if "," in s and "." in s:  # both present: the last one is the decimal separator
        s = s.replace(".", "").replace(",", ".") if s.rfind(",") > s.rfind(".") else s.replace(",", "")
    elif "," in s:
        s = s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None


def parse_date(value):
    if value is None or pd.isna(value) or str(value).strip() == "":
        return pd.NaT
    s = str(value).strip()
    for fmt in ("%Y-%m-%d", "%d.%m.%Y", "%d/%m/%Y", "%d %b %Y", "%Y-%m-%d %H:%M:%S"):
        try:
            return pd.to_datetime(s, format=fmt)
        except ValueError:
            continue
    return pd.NaT


def normalise_text(value) -> str:
    if value is None or pd.isna(value):
        return ""
    s = re.sub(r"\s+", " ", str(value)).strip()
    # Title-case, but keep company suffixes in upper case
    suffixes = {"SRL": "SRL", "SA": "SA", "LTD": "Ltd", "GMBH": "GmbH"}
    words = [suffixes.get(w.upper(), w.capitalize()) for w in s.split(" ")]
    return " ".join(words)


# --- Pipeline --------------------------------------------------------------------
def clean(df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    df = df.dropna(how="all", subset=REQUIRED).copy()
    df["date"] = df["date"].apply(parse_date)
    df["client"] = df["client"].apply(normalise_text)
    df["product"] = df["product"].apply(normalise_text)
    df["qty"] = df["qty"].apply(parse_number)
    df["unit_price"] = df["unit_price"].apply(parse_number)

    problems = pd.Series("", index=df.index)
    problems[df["date"].isna()] += "invalid date; "
    problems[df["client"] == ""] += "missing client; "
    problems[df["qty"].isna()] += "missing quantity; "
    problems[df["qty"] < 0] += "negative quantity; "
    problems[df["unit_price"].isna()] += "invalid price; "

    issues = df[problems != ""].assign(issue=problems[problems != ""].str.rstrip("; "))
    good = df[problems == ""].copy()

    before = len(good)
    good = good.drop_duplicates(subset=REQUIRED)
    dup_count = before - len(good)

    good["qty"] = good["qty"].astype(int)
    good["total"] = (good["qty"] * good["unit_price"]).round(2)
    good["month"] = good["date"].dt.to_period("M").astype(str)
    good = good.sort_values(["date", "client"]).reset_index(drop=True)
    good.attrs["duplicates_removed"] = dup_count
    return good, issues


# --- Excel output ----------------------------------------------------------------
HEADER_FILL = PatternFill("solid", fgColor="1F7A4D")
HEADER_FONT = Font(bold=True, color="FFFFFF")


def _style_sheet(ws, money_cols=(), date_cols=()):
    for cell in ws[1]:
        cell.fill, cell.font = HEADER_FILL, HEADER_FONT
        cell.alignment = Alignment(horizontal="center")
    for col_idx, col in enumerate(ws.iter_cols(min_row=1, max_row=ws.max_row), start=1):
        width = max(len(str(c.value)) if c.value is not None else 0 for c in col) + 3
        ws.column_dimensions[get_column_letter(col_idx)].width = min(width, 45)
        header = col[0].value
        for c in col[1:]:
            if header in money_cols:
                c.number_format = "#,##0.00"
            elif header in date_cols:
                c.number_format = "yyyy-mm-dd"
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions


def write_report(good: pd.DataFrame, issues: pd.DataFrame, raw_count: int, out: Path):
    out.parent.mkdir(parents=True, exist_ok=True)
    clean_view = good[["date", "client", "product", "qty", "unit_price", "total", "source_file"]]
    clean_view.columns = ["Date", "Client", "Product", "Qty", "Unit Price", "Total", "Source File"]

    pivot = good.pivot_table(index="client", columns="month", values="total", aggfunc="sum", fill_value=0)
    pivot["Grand Total"] = pivot.sum(axis=1)
    pivot = pivot.sort_values("Grand Total", ascending=False).round(2).reset_index()
    pivot = pivot.rename(columns={"client": "Client"})

    issues_view = issues[["source_file", "source_row", "client", "product", "issue"]]
    issues_view.columns = ["Source File", "Row In File", "Client", "Product", "Issue"]

    log = pd.DataFrame({
        "Step": ["Rows read (all files)", "Blank rows removed", "Invalid rows (see Issues)",
                 "Duplicates removed", "Clean rows in report", "Total sales value"],
        "Result": [str(x) for x in [raw_count, raw_count - len(good) - len(issues) - good.attrs["duplicates_removed"],
                   len(issues), good.attrs["duplicates_removed"], len(good), f"{good['total'].sum():,.2f}"]],
    })

    with pd.ExcelWriter(out, engine="openpyxl") as w:
        log.to_excel(w, sheet_name="Processing Log", index=False)
        pivot.to_excel(w, sheet_name="Summary by Client", index=False)
        clean_view.to_excel(w, sheet_name="Clean Data", index=False)
        issues_view.to_excel(w, sheet_name="Issues", index=False)
        _style_sheet(w.sheets["Processing Log"])
        _style_sheet(w.sheets["Summary by Client"], money_cols=set(pivot.columns) - {"Client"})
        _style_sheet(w.sheets["Clean Data"], money_cols={"Unit Price", "Total"}, date_cols={"Date"})
        _style_sheet(w.sheets["Issues"])
    return log


def main():
    ap = argparse.ArgumentParser(description="Merge and clean messy Excel/CSV exports into one report.")
    ap.add_argument("input_folder", type=Path)
    ap.add_argument("-o", "--output", type=Path, default=Path("output/clean_report.xlsx"))
    args = ap.parse_args()

    files = sorted(p for p in args.input_folder.iterdir() if p.suffix.lower() in {".csv", ".xlsx", ".xls"})
    if not files:
        raise SystemExit(f"No CSV/XLSX files found in {args.input_folder}")
    frames = []
    for f in files:
        df = read_any(f)
        print(f"  read {f.name:<35} {len(df):>4} rows")
        frames.append(df)
    raw = pd.concat(frames, ignore_index=True)
    good, issues = clean(raw)
    log = write_report(good, issues, len(raw), args.output)
    print("\n" + log.to_string(index=False))
    print(f"\nReport saved to {args.output}")


if __name__ == "__main__":
    main()
