# TakeoffNSW Formatter

This repository contains a Streamlit app that converts raw mechanical takeoff CSV/XLSX files into a styled, grouped Takeoff file
with product/tag grouping, subtotals, and Excel styling.

Features:
- Upload CSV or Excel
- Column mapping with presets
- AD-GRD grouped by TAG with per-tag subtotals
- Other products grouped by PRODUCT with a single subtotal
- Brand autofill rules (AD-GRD -> PRICE, FAN -> LOREN COOK, SPLIT SYSTEM -> SAMSUNG)
- Option to aggregate exact duplicate rows (sums QTY)
- Excel export with styling (header/subtotal/grand fills, bold columns)
- Unit-splitting (ZIP one file per unit)

## Run locally

1. Create a virtualenv and install dependencies:

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

2. Run the Streamlit app:

```bash
streamlit run takeoffnsw_formatter.py
```

## Files
- `takeoffnsw_formatter.py` - main Streamlit app
- `requirements.txt` - Python dependencies
- `README.md` - this file
