"""
===============================================================
QuantStockLab
US Market Dashboard

Part 1/4

Loads the newest Seeking Alpha export and prepares the data
for dashboard analysis.

Python 3.11+

Required packages

pip install pandas numpy rich openpyxl python-calamine
===============================================================
"""

from __future__ import annotations

from pathlib import Path
from datetime import datetime
import atexit
import html
import re
import sys

import numpy as np
import pandas as pd
import openpyxl

pd.set_option("future.no_silent_downcasting", True)

# -------------------------------------------------------------
# Configuration
# -------------------------------------------------------------

DATA_FOLDER = Path("/Users/elliott/Projects/StockData/")
XLSX_REPORT_FILE = Path("reports/Daily_Watchlist.xlsx")
HTML_REPORT_FILE = Path("reports/Daily_Dashboard.html")

# -------------------------------------------------------------
# Terminal Output Capture
# -------------------------------------------------------------

_terminal_output = []


class TeeOutput:

    def __init__(self, stream):
        self.stream = stream

    def write(self, text):
        _terminal_output.append(text)
        return self.stream.write(text)

    def flush(self):
        return self.stream.flush()


def write_html_report():

    output = "".join(_terminal_output)

    document = f"""<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>QuantStockLab Daily Dashboard</title>
  <style>
    body {{
      margin: 0;
      background: #f6f7f9;
      color: #111827;
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
    }}
    main {{
      max-width: 1180px;
      margin: 0 auto;
      padding: 32px 24px;
    }}
    h1 {{
      margin: 0 0 8px;
      font-size: 28px;
      font-weight: 700;
    }}
    .meta {{
      margin: 0 0 24px;
      color: #4b5563;
      font-size: 14px;
    }}
    pre {{
      margin: 0;
      padding: 24px;
      overflow-x: auto;
      background: #111827;
      color: #f9fafb;
      border-radius: 8px;
      line-height: 1.45;
      font-size: 13px;
      white-space: pre;
    }}
  </style>
</head>
<body>
  <main>
    <h1>QuantStockLab Daily Dashboard</h1>
    <p class="meta">Generated {html.escape(datetime.now().strftime("%Y-%m-%d %H:%M:%S"))}</p>
    <pre>{html.escape(output)}</pre>
  </main>
</body>
</html>
"""

    HTML_REPORT_FILE.write_text(document, encoding="utf-8")


sys.stdout = TeeOutput(sys.stdout)
sys.stderr = TeeOutput(sys.stderr)
atexit.register(write_html_report)

# =============================================================
# Universe Filter
#
# MIN_MARKET_CAP_B = 0 means no lower limit
# MAX_MARKET_CAP_B = 0 means no upper limit
#
# Examples:
#
# 0,0       All stocks
# 2,0       >= $2B
# 10,0      >= $10B
# 200,0     >= $200B
# 2,10      $2B ~ $10B
# 10,200    $10B ~ $200B
#
# =============================================================

MIN_MARKET_CAP_B = 0
MAX_MARKET_CAP_B = 0

MISSING_VALUES = {
    "-",
    "--",
    "—",
    "",
    "N/A",
    "n/a",
    "NA",
    "NM",
    "None",
    "null",
}

MARKET_CAP_MULTIPLIER = {
    "K": 1e3,
    "M": 1e6,
    "B": 1e9,
    "T": 1e12,
}


# -------------------------------------------------------------
# Excel Loader
# -------------------------------------------------------------

def newest_excel():

    files = [
        file
        for file in DATA_FOLDER.glob("*.xlsx")
        if not file.name.startswith("~$")
    ]

    if not files:
        raise FileNotFoundError(
            f"No xlsx found inside {DATA_FOLDER}"
        )

    return max(files, key=lambda f: f.stat().st_mtime)


def read_excel(path):

    try:

        return pd.read_excel(path, engine="calamine")

    except Exception:

        return pd.read_excel(path, engine="openpyxl")


# -------------------------------------------------------------
# Parsing helpers
# -------------------------------------------------------------

def parse_market_cap(x):

    if pd.isna(x):
        return np.nan

    if isinstance(x, (int, float)):
        return float(x)

    s = str(x).strip().upper()

    if s in MISSING_VALUES:
        return np.nan

    m = re.match(r"([0-9.]+)\s*([KMBT])$", s)

    if m:

        return (
            float(m.group(1))
            * MARKET_CAP_MULTIPLIER[m.group(2)]
        )

    try:

        return float(s)

    except:

        return np.nan


def parse_percent(x):

    if pd.isna(x):
        return np.nan

    if isinstance(x, (int, float)):
        return float(x)

    s = str(x).strip()

    if s in MISSING_VALUES:
        return np.nan

    s = s.replace("%", "")

    try:

        return float(s)

    except:

        return np.nan


def parse_ratio(x):

    if pd.isna(x):
        return np.nan

    if isinstance(x, (int, float)):
        return float(x)

    s = str(x).strip()

    if s in MISSING_VALUES:
        return np.nan

    s = s.replace("x", "")

    try:

        return float(s)

    except:

        return np.nan


# -------------------------------------------------------------
# Cleaning
# -------------------------------------------------------------

def clean_dataframe(df):

    df = df.copy()

    df = df.replace(list(MISSING_VALUES), np.nan)

    for col in df.columns:

        name = col.lower()

        # Market Cap

        if "market cap" in name:

            df[col] = df[col].apply(parse_market_cap)

            continue

        # Percentages

        if (
            "%" in col
            or "growth" in name
            or "margin" in name
            or "yield" in name
            or "return" in name
            or "perf" in name
            or "yoy" in name
        ):

            df[col] = df[col].apply(parse_percent)

            continue

        # Ratios

        if (
            "ratio" in name
            or "p/e" in name
            or "peg" in name
        ):

            df[col] = df[col].apply(parse_ratio)

            continue

        converted = pd.to_numeric(
            df[col],
            errors="coerce"
        )

        # only convert when most rows are numeric

        if converted.notna().sum() > len(df) * 0.80:

            df[col] = converted

    return df


# -------------------------------------------------------------
# Market Cap Universe Filter
# -------------------------------------------------------------

def filter_market_cap(df, min_b=0, max_b=0):

    if "Market Cap" not in df.columns:
        return df

    total = len(df)

    mask = pd.Series(True, index=df.index)

    if min_b > 0:
        mask &= df["Market Cap"] >= min_b * 1e9

    if max_b > 0:
        mask &= df["Market Cap"] <= max_b * 1e9

    filtered = df.loc[mask].copy()

    print()
    print("=" * 62)
    print("Market Cap Universe")
    print("=" * 62)

    if min_b == 0:
        print("Minimum Market Cap : None")
    else:
        print(f"Minimum Market Cap : ${min_b:,}B")

    if max_b == 0:
        print("Maximum Market Cap : None")
    else:
        print(f"Maximum Market Cap : ${max_b:,}B")

    print(f"Stocks Selected    : {len(filtered):,}")
    print(f"Stocks Removed     : {total-len(filtered):,}")
    print(f"Universe Size      : {total:,}")

    print("=" * 62)

    return filtered


# -------------------------------------------------------------
# Utility Functions
# -------------------------------------------------------------

def safe_mean(df, column):

    if column not in df.columns:
        return np.nan

    return df[column].mean()


def safe_median(df, column):

    if column not in df.columns:
        return np.nan

    return df[column].median()


def safe_count(df, column, condition):

    if column not in df.columns:
        return 0

    return int(condition(df[column]).sum())


def fmt(x, digits=2):

    if pd.isna(x):

        return "-"

    return f"{x:,.{digits}f}"


def fmt_pct(x):

    if pd.isna(x):

        return "-"

    return f"{x:.2f}%"


# -------------------------------------------------------------
# Pretty printing
# -------------------------------------------------------------

def title(text):

    print()
    print("=" * 62)
    print(text)
    print("=" * 62)


def section(text):

    print()
    print(text)
    print("-" * 62)


def line(name, value):

    print(f"{name:<35}{value:>25}")


# -------------------------------------------------------------
# Load data
# -------------------------------------------------------------

excel_file = newest_excel()

print()

df = read_excel(excel_file)

df = clean_dataframe(df)

df = filter_market_cap(
    df,
    MIN_MARKET_CAP_B,
    MAX_MARKET_CAP_B
)

print("Done.")

print(f"Rows    : {len(df):,}")

print(f"Columns : {len(df.columns)}")


# =============================================================
# PART 2 - Core Market Dashboard
# =============================================================

def market_breadth(df):

    title("US MARKET BREADTH")

    total = len(df)

    adv = safe_count(df, "Change %", lambda s: s > 0)
    dec = safe_count(df, "Change %", lambda s: s < 0)
    flat = total - adv - dec

    ratio = adv / dec if dec else np.nan

    line("Stocks", f"{total:,}")
    line("Advancers", f"{adv:,}")
    line("Decliners", f"{dec:,}")
    line("Unchanged", f"{flat:,}")
    line("Advance / Decline", fmt(ratio))
    line("Average Change", fmt_pct(safe_mean(df, "Change %")))


def momentum(df):

    title("MOMENTUM")

    sma_columns = [
        ("10D SMA", "Last Price Vs. 10D SMA"),
        ("50D SMA", "Last Price Vs. 50D SMA"),
        ("100D SMA", "Last Price Vs. 100D SMA"),
        ("200D SMA", "Last Price Vs. 200D SMA"),
    ]

    for label, col in sma_columns:

        if col not in df.columns:
            continue

        above = int((df[col] > 0).sum())

        pct = above / len(df) * 100

        line(
            f"Above {label}",
            f"{above:,} ({pct:.1f}%)"
        )

    print()

    rsi = safe_mean(df, "RSI")

    line("Average RSI", fmt(rsi))

    if "RSI" in df.columns:

        over70 = int((df["RSI"] >= 70).sum())

        under30 = int((df["RSI"] <= 30).sum())

        line("RSI > 70", f"{over70:,}")

        line("RSI < 30", f"{under30:,}")


def valuation(df):

    title("VALUATION")

    cols = [
        "P/E TTM",
        "P/E FWD",
        "PEG TTM",
        "PEG FWD",
        "Price / Sales",
        "EV / Sales",
        "EV / EBITDA",
        "Price / Book",
        "Price / Cash Flow",
    ]

    for col in cols:

        if col not in df.columns:
            continue

        value = safe_median(df, col)

        line(col, fmt(value))


def growth(df):

    title("GROWTH")

    cols = [
        "Revenue YoY",
        "Revenue FWD",
        "Revenue 3Y",
        "Revenue 5Y",
        "EPS YoY",
        "EPS Growth (FWD)",
        "EPS 3Y",
        "EBITDA YoY",
        "EBITDA FWD",
        "FCF 3Y",
    ]

    for col in cols:

        if col not in df.columns:
            continue

        value = safe_median(df, col)

        line(col, fmt_pct(value))


def quality(df):

    title("QUALITY")

    cols = [
        "Profit Margin",
        "EBIT Margin",
        "EBITDA Margin",
        "Net Income Margin",
        "FCF Margin",
        "Return on Equity",
        "Return on Assets",
        "Return on Total Capital",
        "ROE Growth YoY",
    ]

    for col in cols:

        if col not in df.columns:
            continue

        line(
            col,
            fmt_pct(safe_mean(df, col))
        )


# =============================================================
# Run Dashboard
# =============================================================

print()

print("=" * 62)
print("QuantStockLab Daily Dashboard")
print(datetime.now().strftime("%Y-%m-%d %H:%M"))
print("=" * 62)

market_breadth(df)

momentum(df)

valuation(df)

growth(df)

quality(df)

print()

print("=" * 62)
print("End of Dashboard")
print("=" * 62)



# =============================================================
# PART 3 - Opportunities & Rankings
# =============================================================

def print_table(title_text, dataframe, columns, n=10):

    title(title_text)

    available = [c for c in columns if c in dataframe.columns]

    if len(available) == 0:
        print("No data.")
        return

    print(dataframe[available].head(n).to_string(index=False))


# -------------------------------------------------------------
# Top Gainers
# -------------------------------------------------------------

def top_gainers(df, n=20):

    if "Change %" not in df.columns:
        return

    cols = [
        "Symbol",
        "Company Name",
        "Change %",
        "Price",
        "Market Cap",
        "Quant Rating"
    ]

    d = df.sort_values(
        "Change %",
        ascending=False
    )

    print_table("TOP GAINERS", d, cols, n)


# -------------------------------------------------------------
# Top Losers
# -------------------------------------------------------------

def top_losers(df, n=20):

    if "Change %" not in df.columns:
        return

    cols = [
        "Symbol",
        "Company Name",
        "Change %",
        "Price",
        "Market Cap",
        "Quant Rating"
    ]

    d = df.sort_values(
        "Change %",
        ascending=True
    )

    print_table("TOP LOSERS", d, cols, n)


# -------------------------------------------------------------
# Highest Quant Rating
# -------------------------------------------------------------

def top_quant(df, n=20):

    if "Quant Rating" not in df.columns:
        return

    cols = [
        "Symbol",
        "Company Name",
        "Quant Rating",
        "Valuation",
        "Growth",
        "Profitability",
        "Momentum",
    ]

    d = df.sort_values(
        "Quant Rating",
        ascending=False
    )

    print_table("TOP QUANT RATING", d, cols, n)


# -------------------------------------------------------------
# Highest Revenue Growth
# -------------------------------------------------------------

def top_revenue_growth(df, n=20):

    if "Revenue YoY" not in df.columns:
        return

    cols = [
        "Symbol",
        "Company Name",
        "Revenue YoY",
        "EPS YoY",
        "Market Cap",
        "Quant Rating",
    ]

    d = df.sort_values(
        "Revenue YoY",
        ascending=False
    )

    print_table("FASTEST REVENUE GROWTH", d, cols, n)


# -------------------------------------------------------------
# Highest EPS Growth
# -------------------------------------------------------------

def top_eps_growth(df, n=20):

    if "EPS YoY" not in df.columns:
        return

    cols = [
        "Symbol",
        "Company Name",
        "EPS YoY",
        "Revenue YoY",
        "Quant Rating",
    ]

    d = df.sort_values(
        "EPS YoY",
        ascending=False
    )

    print_table("FASTEST EPS GROWTH", d, cols, n)


# -------------------------------------------------------------
# Highest Dividend Yield
# -------------------------------------------------------------

def top_dividend(df, n=20):

    if "Yield FWD" not in df.columns:
        return

    d = df[df["Yield FWD"] > 0]

    cols = [
        "Symbol",
        "Company Name",
        "Yield FWD",
        "Years of Growth",
        "Payout Ratio",
        "Quant Rating",
    ]

    d = d.sort_values(
        "Yield FWD",
        ascending=False
    )

    print_table("HIGHEST DIVIDEND YIELD", d, cols, n)


# -------------------------------------------------------------
# Best Value
# -------------------------------------------------------------

def best_value(df, n=20):

    if "PEG FWD" not in df.columns:
        return

    d = df.copy()

    d = d[d["PEG FWD"] > 0]

    cols = [
        "Symbol",
        "Company Name",
        "PEG FWD",
        "P/E FWD",
        "Revenue YoY",
        "Quant Rating",
    ]

    d = d.sort_values(
        "PEG FWD"
    )

    print_table("LOWEST PEG", d, cols, n)


# -------------------------------------------------------------
# Sector Summary
# -------------------------------------------------------------

def sector_summary(df):

    if "Sector & Industry" not in df.columns:
        return

    title("SECTOR SUMMARY")

    sector = (
        df["Sector & Industry"]
        .fillna("Unknown")
        .astype(str)
        .str.split("|")
        .str[0]
        .str.strip()
    )

    counts = sector.value_counts()

    print(counts.head(20).to_string())


# -------------------------------------------------------------
# Quant Rating Distribution
# -------------------------------------------------------------

def quant_distribution(df):

    if "Quant Rating" not in df.columns:
        return

    title("QUANT RATING DISTRIBUTION")

    bins = [0,1,2,3,4,5.01]

    labels = [
        "0-1",
        "1-2",
        "2-3",
        "3-4",
        "4-5"
    ]

    groups = pd.cut(
        df["Quant Rating"],
        bins=bins,
        labels=labels
    )

    print(groups.value_counts().sort_index())


# =============================================================
# Execute Part 3
# =============================================================

quant_distribution(df)

sector_summary(df)

top_gainers(df)

top_losers(df)

top_quant(df)

top_revenue_growth(df)

top_eps_growth(df)

best_value(df)

top_dividend(df)


# =============================================================
# PART 4 - Smart Scoring & Daily Watchlist
# =============================================================

def normalize(series):
    """Normalize a numeric series to 0-100."""
    s = pd.to_numeric(series, errors="coerce")

    if s.notna().sum() == 0:
        return pd.Series(50.0, index=series.index)

    mn = s.min()
    mx = s.max()

    if pd.isna(mn) or pd.isna(mx) or mx == mn:
        return pd.Series(50.0, index=series.index)

    return (s - mn) / (mx - mn) * 100


def normalize_inverse(series):
    """Lower is better."""
    return 100 - normalize(series)


def build_watchlist(df):

    d = df.copy()

    score = pd.Series(0.0, index=d.index)

    # ----------------------------
    # Quant Rating
    # ----------------------------
    if "Quant Rating" in d.columns:
        score += normalize(d["Quant Rating"]) * 0.30

    # ----------------------------
    # Growth
    # ----------------------------
    if "Revenue YoY" in d.columns:
        score += normalize(d["Revenue YoY"]) * 0.15

    if "EPS YoY" in d.columns:
        score += normalize(d["EPS YoY"]) * 0.15

    # ----------------------------
    # Profitability
    # ----------------------------
    if "Profit Margin" in d.columns:
        score += normalize(d["Profit Margin"]) * 0.10

    if "Return on Equity" in d.columns:
        score += normalize(d["Return on Equity"]) * 0.10

    # ----------------------------
    # Momentum
    # ----------------------------
    if "6M Perf" in d.columns:
        score += normalize(d["6M Perf"]) * 0.10

    if "Last Price Vs. 200D SMA" in d.columns:
        score += normalize(d["Last Price Vs. 200D SMA"]) * 0.05

    # ----------------------------
    # Valuation
    # ----------------------------
    if "PEG FWD" in d.columns:
        score += normalize_inverse(d["PEG FWD"]) * 0.05

    # ----------------------------
    # Analyst Revisions
    # ----------------------------
    if "EPS Rev." in d.columns:
        score += normalize(d["EPS Rev."]) * 0.10

    d["Composite Score"] = score.round(2)

    return d.sort_values(
        "Composite Score",
        ascending=False
    )


# -------------------------------------------------------------
# Daily Watchlist
# -------------------------------------------------------------

def daily_watchlist(df, n=25):

    d = build_watchlist(df)

    cols = [
        "Symbol",
        "Company Name",
        "Composite Score",
        "Quant Rating",
        "Revenue YoY",
        "EPS YoY",
        "PEG FWD",
        "6M Perf",
    ]

    cols = [c for c in cols if c in d.columns]

    print_table(
        "TODAY'S WATCHLIST",
        d,
        cols,
        n,
    )


# -------------------------------------------------------------
# Market Health Score
# -------------------------------------------------------------

def market_health(df):

    score = 0

    maximum = 100

    if "Last Price Vs. 200D SMA" in df.columns:

        pct = (
            (df["Last Price Vs. 200D SMA"] > 0)
            .mean()
            * 100
        )

        score += pct * 0.30

    if "Last Price Vs. 50D SMA" in df.columns:

        pct = (
            (df["Last Price Vs. 50D SMA"] > 0)
            .mean()
            * 100
        )

        score += pct * 0.20

    if "RSI" in df.columns:

        rsi = df["RSI"].mean()

        score += min(rsi, 100) * 0.20

    if "Quant Rating" in df.columns:

        q = df["Quant Rating"].mean()

        score += q / 5 * 100 * 0.30

    title("MARKET HEALTH")

    print()

    print(f"Overall Score : {score:.1f} / {maximum}")

    if score >= 80:
        print("Condition     : VERY STRONG")
    elif score >= 65:
        print("Condition     : STRONG")
    elif score >= 50:
        print("Condition     : NEUTRAL")
    elif score >= 35:
        print("Condition     : WEAK")
    else:
        print("Condition     : VERY WEAK")


# -------------------------------------------------------------
# Interesting Statistics
# -------------------------------------------------------------

def interesting_statistics(df):

    title("INTERESTING STATISTICS")

    if "Market Cap" in df.columns:

        mega = (df["Market Cap"] >= 200e9).sum()
        large = ((df["Market Cap"] >= 10e9) &
                 (df["Market Cap"] < 200e9)).sum()
        mid = ((df["Market Cap"] >= 2e9) &
               (df["Market Cap"] < 10e9)).sum()
        small = (df["Market Cap"] < 2e9).sum()

        line("Mega Cap", f"{mega:,}")
        line("Large Cap", f"{large:,}")
        line("Mid Cap", f"{mid:,}")
        line("Small Cap", f"{small:,}")

    if "Country" in df.columns:

        print()

        print("Top Countries")

        print(
            df["Country"]
            .value_counts()
            .head(10)
            .to_string()
        )


# -------------------------------------------------------------
# Export Watchlist
# -------------------------------------------------------------

def export_watchlist(df):

    d = build_watchlist(df)

    outfile = XLSX_REPORT_FILE

    d.to_excel(outfile, index=False)

    print()

    print(f"Watchlist exported -> {outfile}")


# =============================================================
# Execute Part 4
# =============================================================

market_health(df)

interesting_statistics(df)

daily_watchlist(df)

export_watchlist(df)

print()

print("=" * 62)
print("Dashboard Completed Successfully")
print("=" * 62)
