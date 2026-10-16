-- キャラクターごとの能力値・副能力値・経験点・ロイス数。
-- xlsx の「キャラクター」シートにある計算式を SQL で再現したビュー。
CREATE OR REPLACE VIEW `{project}.{dataset}.v_character_status` AS
WITH ability_totals AS (
  -- 能力値シートは「1キャラ4行（縦持ち）」なので、キャラごとに1行（横持ち）へ並べ替える
  SELECT
    character_id,
    SUM(IF(ability = '肉体', total, 0)) AS body,
    SUM(IF(ability = '感覚', total, 0)) AS sense,
    SUM(IF(ability = '精神', total, 0)) AS mind,
    SUM(IF(ability = '社会', total, 0)) AS social
  FROM (
    SELECT
      character_id,
      ability,
      IFNULL(from_syndrome, 0) + IFNULL(from_works, 0) + IFNULL(from_growth, 0) + IFNULL(from_other, 0) AS total
    FROM `{project}.{dataset}.abilities`
  )
  GROUP BY character_id
),
procurement AS (
  SELECT character_id, SUM(level) AS procurement_level
  FROM `{project}.{dataset}.skills`
  WHERE skill = '調達'
  GROUP BY character_id
),
exp_used AS (
  SELECT character_id, SUM(points) AS used_exp
  FROM `{project}.{dataset}.exp_logs`
  GROUP BY character_id
),
exp_gained AS (
  SELECT character_id, SUM(exp_gained) AS gained_exp
  FROM `{project}.{dataset}.sessions`
  GROUP BY character_id
),
lois_counts AS (
  SELECT
    character_id,
    COUNTIF(status = 'ロイス') AS lois_count,
    COUNTIF(status = 'タイタス') AS titus_count,
    MAX(IF(is_d_lois, target, NULL)) AS d_lois
  FROM `{project}.{dataset}.loises`
  GROUP BY character_id
)
SELECT
  c.character_id,
  c.name,
  c.code_name,
  c.breed,
  c.syndrome1,
  c.syndrome2,
  c.syndrome3,
  c.works,
  IFNULL(c.awakening_erosion, 0) + IFNULL(c.impulse_erosion, 0) AS base_erosion,
  c.current_erosion,
  CASE
    WHEN c.current_erosion >= 160 THEN 5
    WHEN c.current_erosion >= 130 THEN 4
    WHEN c.current_erosion >= 100 THEN 3
    WHEN c.current_erosion >= 80 THEN 2
    WHEN c.current_erosion >= 60 THEN 1
    ELSE 0
  END AS dice_bonus,
  CASE
    WHEN c.current_erosion >= 160 THEN 2
    WHEN c.current_erosion >= 100 THEN 1
    ELSE 0
  END AS effect_level_bonus,
  a.body,
  a.sense,
  a.mind,
  a.social,
  IFNULL(p.procurement_level, 0) AS procurement_level,
  a.body * 2 + a.mind + 20 AS max_hp,
  c.current_hp,
  a.sense * 2 + a.mind AS initiative,
  a.sense * 2 + a.mind + 5 AS move,
  (a.sense * 2 + a.mind + 5) * 2 AS full_move,
  a.social * 2 + IFNULL(p.procurement_level, 0) * 2 AS stock_points,
  c.wealth_points,
  c.initial_exp,
  IFNULL(g.gained_exp, 0) AS gained_exp,
  IFNULL(u.used_exp, 0) AS used_exp,
  IFNULL(c.initial_exp, 0) + IFNULL(g.gained_exp, 0) - IFNULL(u.used_exp, 0) AS remaining_exp,
  IFNULL(l.lois_count, 0) AS lois_count,
  IFNULL(l.titus_count, 0) AS titus_count,
  l.d_lois
FROM `{project}.{dataset}.characters` AS c
LEFT JOIN ability_totals AS a USING (character_id)
LEFT JOIN procurement AS p USING (character_id)
LEFT JOIN exp_used AS u USING (character_id)
LEFT JOIN exp_gained AS g USING (character_id)
LEFT JOIN lois_counts AS l USING (character_id)
