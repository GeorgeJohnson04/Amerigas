-- Landed-cost detail, one row per delivered load.
-- Illustrative schema. Aliases match data/loads.csv so the report runs unchanged on Snowflake.
SELECT
    load_id,
    order_date,
    DATE_TRUNC('month', order_date) AS year_month,
    district,
    origin,
    destination,
    move_type,
    gallons,
    contract_id,
    product_cost,
    overhead,
    storage,
    freight,
    grand_total,
    basis,
    lic AS warehouse_lic
FROM ANALYTICS_DB.REPORTING.ZONE_DETAIL
WHERE load_type = 'Actual';   -- the table also holds forecast rows for the rest of the month
