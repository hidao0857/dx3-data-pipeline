-- シンドローム別の人数と平均ステータス。
-- 1キャラが最大3つのシンドロームを持つため、UNNEST で「キャラ×シンドローム」の行に展開してから集計する。
CREATE OR REPLACE VIEW `{project}.{dataset}.v_syndrome_summary` AS
SELECT
  syndrome,
  COUNT(DISTINCT s.character_id) AS characters,
  COUNTIF(s.breed = 'ピュア') AS pure_characters,
  ROUND(AVG(s.body), 1) AS avg_body,
  ROUND(AVG(s.sense), 1) AS avg_sense,
  ROUND(AVG(s.mind), 1) AS avg_mind,
  ROUND(AVG(s.social), 1) AS avg_social,
  ROUND(AVG(s.max_hp), 1) AS avg_max_hp,
  ROUND(AVG(s.initiative), 1) AS avg_initiative
FROM `{project}.{dataset}.v_character_status` AS s,
  UNNEST([s.syndrome1, s.syndrome2, s.syndrome3]) AS syndrome
WHERE syndrome IS NOT NULL
GROUP BY syndrome
