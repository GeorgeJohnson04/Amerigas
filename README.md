# Energy Supply Analytics Portfolio

Python tools I built for propane supply planning, hedging and cost analytics. This repo holds the code only. It contains no company data, and every name, location, volume and price in the examples is fictional.

| Project | Summary | Stack |
|---|---|---|
| [supply-optimization-model](supply-optimization-model/) | A linear-programming supply and transport model that replaces Excel + OpenSolver. It reads each area's workbook, solves a 12-month buy/ship/store plan with HiGHS, and writes region workbooks. It supports a scenario override deck, and two local web apps let you edit and compare plans. | Python, SciPy/HiGHS, openpyxl, Flask, vanilla JS |
| [hedge-strip-tracker](hedge-strip-tracker/) | Prices forward hedge strips each day from blended hub curves, weighted by demand shape, and builds an Excel dashboard. Includes a synthetic data generator so it runs out of the box. | pandas, pyarrow, openpyxl |
| [snowflake-cost-reporting](snowflake-cost-reporting/) | Landed-cost data-quality and outlier reporting on Snowflake, with Time Travel snapshot archiving and reconciliation against finance data. | snowflake-connector-python, pandas, openpyxl |

Each folder has its own README and `requirements.txt`.
