# Landed-Cost Variance by Supply Route

Explains why a district's delivered cost per gallon (landed cost above the market index) differs from its monthly forecast. It breaks the gap down by supply route and checks each load against its contract.

All data is synthetic. The district, suppliers, locations, hubs, volumes and prices are invented, and `generate_sample_data.py` builds them from a fixed seed.

## What it gets right

Pulling a landed-cost table straight from a warehouse and comparing it to forecast gives wrong answers for four reasons. This project handles each one:

| Pitfall | How it is handled |
|---|---|
| The warehouse `basis` label is coarse (every row says `HUB`), so TET vs Non-TET is invisible. | The index comes from the contract table. Each load is measured against the index its contract prices off. |
| Contracts price on different days: same day, or the prior business day. | The pricing date follows the contract rule before the index price is looked up. |
| The warehouse forecast column holds one month's value on every row. | It is ignored. Each load is matched to the forecast for its own month. |
| Seasonal premiums switch on Oct 1. | Each load's implied premium is compared with the contract premium for its season, and mismatches are flagged. |

It also shows the error from using the wrong index (`Other index gap`), and it keeps storage out of the headline number.

## Calculation

```
differential         = product cost - index price (on the pricing date)
landed cost above index = differential + freight + overhead
variance             = landed cost above index - forecast for the load's month
```

Remote-rail freight is a fixed cost per car, so its $/gal moves with gallons shipped.

## Run

```bash
pip install -r requirements.txt
python generate_sample_data.py      # writes data/*.csv
python route_variance_report.py     # writes Output/landed_cost_variance_*.xlsx
```

The workbook has four tabs: `Summary` (by month, by route, chart, checks), `Loads` (every load with its working), `Contracts` and `Method`.

`python route_variance_report.py --source snowflake` runs the same report from the queries in `sql/`. Those queries use an illustrative schema, so adjust table and column names to your warehouse. Connection settings come from environment variables (see `snowflake_connect.py`).

## Sample output (synthetic)

| Group | Loads | Above index | Forecast | Variance | Variance $ |
|---|---|---|---|---|---|
| Sep 2026 | 27 | 0.2550 | 0.2400 | +0.0150 | +$7,657 |
| Oct 2026 (to date) | 6 | 0.3043 | 0.3310 | -0.0267 | -$2,485 |
| Delivered rail | 11 | 0.2329 | 0.2562 | -0.0234 | -$7,655 |
| Terminal trucks | 15 | 0.2705 | 0.2576 | +0.0129 | +$1,727 |
| Port spot trucks | 3 | 0.3892 | 0.2703 | +0.1188 | +$3,066 |
| Remote rail | 4 | 0.3093 | 0.2400 | +0.0693 | +$8,034 |
| All loads | 33 | 0.2626 | 0.2540 | +0.0086 | +$5,172 |

Route rows blend both months, so their forecast is a gallon-weighted mix of the September and October forecasts. The premium check flags one port truck on Sep 29: it was priced at the +0.30 winter premium although the contract says +0.22 until Oct 1.

## Files

| File | Purpose |
|---|---|
| `generate_sample_data.py` | Builds the synthetic loads, contracts, index prices and forecast. |
| `route_variance_report.py` | The calculation and the Excel report. |
| `snowflake_connect.py` | Connection helper (SSO by default, settings from env vars). |
| `sql/` | Example queries for each input table. |
