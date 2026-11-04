-- One row per supply contract: which index it prices off, which day, and the seasonal premiums.
-- Illustrative schema. The index lives on the contract's pricing provision, not on the load table.
SELECT
    deal_id            AS contract_id,
    supplier_name      AS supplier,
    route_name         AS route,
    pricing_location   AS index_name,     -- e.g. 'Gulf Hub TET' vs 'Gulf Hub Non-TET'
    pricing_rule,                         -- same_day | prior_business_day | none
    summer_premium,
    winter_premium
FROM ANALYTICS_DB.CONTRACTS.DEAL_PRICING;
