WITH ttype AS (
  SELECT mtt.transaction_type_id
  FROM   apps.mtl_transaction_types mtt
  WHERE  mtt.transaction_action_id = 1   -- Issue
    AND  mtt.transaction_source_type_id = 13 -- Inventory (misc txn family)
),
scoped AS (
  SELECT
      mmt.*,
      MAX(mmt.transaction_date)
        OVER (PARTITION BY mmt.organization_id, mmt.inventory_item_id) AS last_dt
  FROM apps.mtl_material_transactions mmt
  WHERE mmt.transaction_type_id IN (SELECT transaction_type_id FROM ttype)
),
windowed AS (
  SELECT
    s.organization_id,
    s.inventory_item_id,
    s.transaction_date,
    COALESCE(s.transaction_quantity, s.primary_quantity) AS qty
  FROM scoped s
  WHERE s.transaction_date BETWEEN (s.last_dt - INTERVAL '1' DAY) AND s.last_dt
)
SELECT
  mp.organization_code                   AS org,
  msib.segment1                          AS item_name,
  MAX(w.transaction_date)                AS newest_transaction_date,
  SUM(w.qty)                             AS total_qty_last_24h
FROM windowed w
JOIN apps.mtl_system_items_b msib
  ON msib.organization_id   = w.organization_id
 AND msib.inventory_item_id = w.inventory_item_id
JOIN apps.mtl_parameters mp
  ON mp.organization_id     = w.organization_id
WHERE mp.organization_id = 2035
GROUP BY
  mp.organization_code,
  msib.segment1
ORDER BY
  org,
  item_name
