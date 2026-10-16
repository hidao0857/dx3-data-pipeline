-- 技能ごとの取得人数と平均Lv。RANK() で取得人数の多い順に順位を付ける。
CREATE OR REPLACE VIEW `{project}.{dataset}.v_skill_ranking` AS
SELECT
  ability,
  skill,
  COUNT(DISTINCT character_id) AS characters,
  ROUND(AVG(level), 2) AS avg_level,
  MAX(level) AS max_level,
  RANK() OVER (ORDER BY COUNT(DISTINCT character_id) DESC) AS rank_by_characters
FROM `{project}.{dataset}.skills`
WHERE level > 0
GROUP BY ability, skill
