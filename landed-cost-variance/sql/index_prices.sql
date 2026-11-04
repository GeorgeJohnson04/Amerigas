-- Daily hub prices, one row per index per day. Illustrative schema.
SELECT
    price_date,
    index_name,
    price
FROM ANALYTICS_DB.MARKET.INDEX_PRICES
WHERE price_date >= DATEADD(month, -3, CURRENT_DATE);
