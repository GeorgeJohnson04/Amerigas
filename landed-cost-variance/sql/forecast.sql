-- Monthly forecast of landed cost above index, one row per district per month (Apr..Mar labels).
-- Illustrative schema; in practice this often comes from a finance workbook rather than the warehouse.
SELECT
    district,
    month,
    forecast
FROM ANALYTICS_DB.FINANCE.LANDED_COST_FORECAST;
