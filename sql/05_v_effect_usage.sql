-- エフェクトマスタの1件ごとに、何人のキャラが取得しているかと平均SL。
-- マスタ（名称）とキャラごとのエフェクト（明細）を名称で結合する。
CREATE OR REPLACE VIEW `{project}.{dataset}.v_effect_usage` AS
SELECT
  m.effect_name,
  m.syndrome,
  m.timing,
  m.erosion,
  COUNT(DISTINCT e.character_id) AS characters,
  ROUND(AVG(e.level), 2) AS avg_level,
  MAX(e.level) AS max_level
FROM `{project}.{dataset}.effect_master` AS m
LEFT JOIN `{project}.{dataset}.effects` AS e
  ON e.effect_name = m.effect_name
GROUP BY m.effect_name, m.syndrome, m.timing, m.erosion
