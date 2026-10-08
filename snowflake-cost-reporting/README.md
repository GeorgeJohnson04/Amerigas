# Snowflake Landed-Cost Reporting

Python scripts that pull a landed-cost table (one row per load) from Snowflake. They run data-quality checks, flag outliers, archive snapshots, and compare loads against a finance cost workbook. All output goes to formatted Excel with charts.

| Script | Purpose |
|---|---|
| `snowflake_connect.py` | Connection helper. Uses SSO (external browser) by default, or a password if one is set. All settings come from env vars. |
| `zone_detail_report.py` | Data-quality report covering rule checks, IQR outliers, blanks, placeholder values, spelling variants, date gaps and duplicates. |
| `pull_history.py` | Archives every snapshot still reachable through Time Travel, then combines them into a single history. |
| `district_report.py` | Splits one district's loads into outliers and non-outliers, with charts. |
| `district_cost_check.py` | Compares one district's per-load cost with the finance workbook. |
| `district_vs_finance.py` | Shows Snowflake and finance values side by side, without outlier marking. |

## Setup

```bash
pip install -r requirements.txt
```

Environment variables:
- `SNOWFLAKE_ACCOUNT` and `SNOWFLAKE_USER` (required)
- `SNOWFLAKE_PASSWORD`; leave it unset to sign in with SSO
- `SNOWFLAKE_WAREHOUSE`, `SNOWFLAKE_DATABASE`, `SNOWFLAKE_SCHEMA`, `SNOWFLAKE_ROLE`
- `SNOWFLAKE_TABLE` (default `ANALYTICS_DB.REPORTING.ZONE_DETAIL`)
- `FINANCE_WORKBOOK`, `PRICE_INDEX_FILE` (used by the finance comparisons)

## Examples

```bash
python zone_detail_report.py
python pull_history.py --no-pull
python district_report.py --match "SPRINGFIELD"
python district_cost_check.py --match "SPRINGFIELD" --finance finance_costs.xlsx --fiscal-year 2026
```

Use `--from-file <saved .parquet>` to rerun a report without connecting to Snowflake. No data is included in this repo.
