# Hedge Strip Tracker

Tracks the daily price of forward hedge strips for propane programs. For each trading day it blends three hub forward curves (Mont Belvieu Enterprise, Mont Belvieu LST, Conway), weights each delivery month by the program's demand shape, and reports one volume-weighted $/gal price per strip. The output is an Excel workbook with a summary, charts, daily history, the latest month-by-month curve, and a methodology sheet.

## What it shows

- Building forward curves from daily Argus vendor files (`build_price_index.py`).
- Combining hub curves into a weighted blend, priced by delivery month.
- Pricing each strip by demand shape, with rules for partial and expired months.
- A look-back with 1D/1W/1M/1Y changes, the high and low, and where today's price sits in that range (percentile).
- Writing Excel with openpyxl: date-axis line charts, number formats, and live formulas for the totals.
- An incremental daily update: new vendor files are added to the parquet store, with a `.bak` copy kept.

## Run it on synthetic data

```bash
pip install -r requirements.txt
python make_sample_data.py     # random-walk curves + demand shapes in data/
python hedge_tracker.py        # writes Hedge Product Tracker.xlsx
```

The prices and shapes are all made up. To use real data, put the `ArgusFC_YYYYMMDD.csv` exports in `data/LPG/`, then run `build_price_index.py` followed by `hedge_tracker.py`.

## Configuration

`hedge_products.json` defines:
- the origin blend (which index feeds each leg, and its weight)
- each product: start month, length, strip years, hedge deadline, volume and coverage

Every product in the file is an example.

## Pricing rules

- A strip is priced on a given day only when every one of its delivery months is quoted.
- A delivery month that has rolled off the curve keeps its last forward price, so the strip stays priced through delivery.
- Shape weights are normalised within each strip. Hedge coverage changes the gallons only, not the $/gal price.
