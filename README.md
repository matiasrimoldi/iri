# SEC Metrics Extractor

This repository contains a Python script (`sec_metrics.py`) that fetches an HTML/iXBRL filing (such as a 10‑K or 20‑F) from the SEC and attempts to extract a few financial metrics for the last three fiscal years. The script then writes these values and a few calculated ratios into a simple Excel file.

## Requirements

- Python 3.8+
- `requests`
- `beautifulsoup4`
- `lxml`
- `pandas`
- `openpyxl`

These packages can be installed with:

```bash
pip install requests beautifulsoup4 lxml pandas openpyxl
```

## Usage

```bash
python sec_metrics.py <ixbrl-url> <output.xlsx>
```

Example:

```bash
python sec_metrics.py https://www.sec.gov/Archives/.../example.htm report.xlsx
```

The resulting workbook will contain two sheets:

1. **metrics** – a table with one row per fiscal year and columns for revenue, EBITDA, debt, cash, interest, net cash from operations, dividends, and the calculated ratios.
2. **raw** – the raw values captured from the filing used to compute the metrics. This sheet is optional and can be used for verification.

The script relies on the XBRL tags present in the filing. If a tag is not found, the value will be left blank. EBITDA will be computed from operating loss and depreciation/amortization items if it is not reported directly.
