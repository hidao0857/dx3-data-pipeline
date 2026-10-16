-- キャラごとのエフェクト構成と侵蝕値。
-- 侵蝕値には「効果参照」のような文字も入るため、SAFE_CAST で数値にできるものだけを合計する。
CREATE OR REPLACE VIEW `{project}.{dataset}.v_effect_cost` AS
SELECT
  c.character_id,
  c.name,
  COUNT(*) AS effects,
  SUM(e.level) AS total_level,
  COUNTIF(e.timing = 'メジャーアクション') AS major_effects,
  COUNTIF(e.timing = 'マイナーアクション') AS minor_effects,
  COUNTIF(e.timing = 'オートアクション') AS auto_effects,
  SUM(IF(e.timing IN ('メジャーアクション', 'マイナーアクション'), SAFE_CAST(e.erosion AS INT64), 0)) AS attack_turn_erosion,
  COUNTIF(SAFE_CAST(e.erosion AS INT64) IS NULL) AS erosion_not_numeric
FROM `{project}.{dataset}.effects` AS e
JOIN `{project}.{dataset}.characters` AS c USING (character_id)
GROUP BY c.character_id, c.name
