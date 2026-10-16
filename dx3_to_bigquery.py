"""キャラ管理DB（xlsx）の各シートを BigQuery のテーブルに読み込み、集計用のビューを作るスクリプト。

処理の流れ
  1. Extract   : xlsx の各シートを読み込む（入力セルの列だけ。自動計算の列は読まない）
  2. Transform : 列名を英語にそろえ、型（整数・はい/いいえ・日付）を変換する
  3. Check     : キャラIDの重複や、存在しないキャラIDを参照している行がないかを確かめる
  4. Load      : BigQuery のテーブルを丸ごと入れ替える（サンドボックスでも使えるロードジョブ方式）
  5. View      : sql/ フォルダの SQL で、xlsx の計算式と同じ値や集計結果を出すビューを作る

使い方
  python dx3_to_bigquery.py              # 作業フォルダの DX3_キャラ管理DB.xlsx を BigQuery（dx3-portfolio）に読み込む
  python dx3_to_bigquery.py --dry-run    # BigQuery には送らず、output フォルダに CSV を出力して確認
"""

import argparse
import re
import sys
from pathlib import Path

import pandas as pd

BASE_DIR = Path(__file__).resolve().parent  # 作業フォルダ（このスクリプトがあるフォルダ）
DATASET = "dx3"
LOCATION = "asia-northeast1"  # 東京リージョン
SQL_DIR = BASE_DIR / "sql"

# シート名 → (テーブル名, [(xlsx の見出し, 列名, 型), ...])
# 自動計算の列（キャラ名・合計・HP最大値など）はここに書かない。計算は BigQuery のビューでやり直す。
TABLES = {
    "キャラクター": ("characters", [
        ("キャラID", "character_id", "STRING"), ("キャラ名", "name", "STRING"),
        ("コードネーム", "code_name", "STRING"), ("PL名", "player_name", "STRING"),
        ("年齢", "age", "INT64"), ("性別", "gender", "STRING"), ("星座", "constellation", "STRING"),
        ("身長", "height", "INT64"), ("体重", "weight", "INT64"), ("血液型", "blood_type", "STRING"),
        ("ワークス", "works", "STRING"), ("カヴァー", "cover", "STRING"), ("ブリード", "breed", "STRING"),
        ("シンドローム1", "syndrome1", "STRING"), ("シンドローム2", "syndrome2", "STRING"),
        ("シンドローム3", "syndrome3", "STRING"),
        ("出自", "origin", "STRING"), ("経験", "experience", "STRING"), ("邂逅/欲望", "encounter", "STRING"),
        ("覚醒", "awakening", "STRING"), ("覚醒 侵蝕値", "awakening_erosion", "INT64"),
        ("衝動", "impulse", "STRING"), ("衝動 侵蝕値", "impulse_erosion", "INT64"),
        ("現在侵蝕率", "current_erosion", "INT64"), ("現在HP", "current_hp", "INT64"),
        ("財産ポイント", "wealth_points", "INT64"), ("初期経験点", "initial_exp", "INT64"),
        ("メモ", "memo", "STRING"),
    ]),
    "能力値": ("abilities", [
        ("キャラID", "character_id", "STRING"), ("能力値", "ability", "STRING"),
        ("シンドローム", "from_syndrome", "INT64"), ("ワークス", "from_works", "INT64"),
        ("成長", "from_growth", "INT64"), ("その他", "from_other", "INT64"), ("メモ", "memo", "STRING"),
    ]),
    "技能": ("skills", [
        ("キャラID", "character_id", "STRING"), ("能力値", "ability", "STRING"),
        ("技能", "skill", "STRING"), ("Lv", "level", "INT64"), ("メモ", "memo", "STRING"),
    ]),
    "ロイス": ("loises", [
        ("キャラID", "character_id", "STRING"), ("対象", "target", "STRING"), ("関係", "relation", "STRING"),
        ("感情P", "emotion_positive", "STRING"), ("感情N", "emotion_negative", "STRING"),
        ("表の感情", "surface_emotion", "STRING"), ("固定ロイス", "is_fixed", "BOOL"),
        ("状態", "status", "STRING"), ("Dロイス", "is_d_lois", "BOOL"), ("メモ", "memo", "STRING"),
    ]),
    "エフェクト": ("effects", [
        ("キャラID", "character_id", "STRING"), ("エフェクト名", "effect_name", "STRING"),
        ("SL", "level", "INT64"), ("最大SL", "max_level", "INT64"), ("タイミング", "timing", "STRING"),
        ("技能", "skill", "STRING"), ("難易度", "difficulty", "STRING"), ("対象", "target", "STRING"),
        ("射程", "effect_range", "STRING"), ("侵蝕値", "erosion", "STRING"), ("制限", "restriction", "STRING"),
        ("シンドローム", "syndrome", "STRING"), ("効果", "effect_text", "STRING"), ("出典", "source", "STRING"),
    ]),
    "コンボ": ("combos", [
        ("キャラID", "character_id", "STRING"), ("コンボ名", "combo_name", "STRING"),
        ("構成エフェクト", "effects", "STRING"), ("タイミング", "timing", "STRING"), ("技能", "skill", "STRING"),
        ("難易度", "difficulty", "STRING"), ("対象", "target", "STRING"), ("射程", "effect_range", "STRING"),
        ("侵蝕値", "erosion", "STRING"), ("制限", "restriction", "STRING"),
        ("100%未満 ダイス", "dice_under100", "INT64"), ("100%未満 C値", "critical_under100", "INT64"),
        ("100%未満 攻撃力", "attack_under100", "INT64"), ("100%未満 達成値修正", "bonus_under100", "INT64"),
        ("100%以上 ダイス", "dice_over100", "INT64"), ("100%以上 C値", "critical_over100", "INT64"),
        ("100%以上 攻撃力", "attack_over100", "INT64"), ("100%以上 達成値修正", "bonus_over100", "INT64"),
        ("メモ", "memo", "STRING"),
    ]),
    "アイテム": ("items", [
        ("キャラID", "character_id", "STRING"), ("種別", "item_type", "STRING"), ("名称", "item_name", "STRING"),
        ("技能", "skill", "STRING"), ("命中", "accuracy", "INT64"), ("攻撃力", "attack", "INT64"),
        ("ガード値", "guard", "INT64"), ("射程", "effect_range", "STRING"), ("ドッジ", "dodge", "INT64"),
        ("行動", "initiative", "INT64"), ("装甲値", "armor", "INT64"), ("常備化", "procurement", "INT64"),
        ("経験点", "exp_cost", "INT64"), ("装備中", "is_equipped", "BOOL"), ("効果", "effect_text", "STRING"),
    ]),
    "メモリー": ("memories", [
        ("キャラID", "character_id", "STRING"), ("対象", "target", "STRING"), ("関係", "relation", "STRING"),
        ("感情", "emotion", "STRING"), ("取得セッション", "session", "STRING"), ("エピソード", "episode", "STRING"),
    ]),
    "経験点": ("exp_logs", [
        ("キャラID", "character_id", "STRING"), ("区分", "category", "STRING"), ("内容", "content", "STRING"),
        ("点数", "points", "INT64"), ("時期", "timing", "STRING"), ("メモ", "memo", "STRING"),
    ]),
    "セッション": ("sessions", [
        ("キャラID", "character_id", "STRING"), ("日付", "session_date", "DATE"),
        ("シナリオ名", "scenario", "STRING"), ("GM", "gm", "STRING"),
        ("終了時侵蝕率", "end_erosion", "INT64"), ("獲得経験点", "exp_gained", "INT64"), ("メモ", "memo", "STRING"),
    ]),
}


# マスタシート → (テーブル名, 列の定義)。キャラIDを持たず、1列目の名称がキー。
MASTER_TABLES = {
    "エフェクトマスタ": ("effect_master", [
        ("エフェクト名", "effect_name", "STRING"), ("タイミング", "timing", "STRING"), ("技能", "skill", "STRING"),
        ("難易度", "difficulty", "STRING"), ("対象", "target", "STRING"), ("射程", "effect_range", "STRING"),
        ("侵蝕値", "erosion", "STRING"), ("制限", "restriction", "STRING"), ("シンドローム", "syndrome", "STRING"),
        ("効果", "effect_text", "STRING"), ("出典", "source", "STRING"),
        ("初出キャラID", "first_character_id", "STRING"), ("登録日", "registered_date", "DATE"),
    ]),
    "アイテムマスタ": ("item_master", [
        ("名称", "item_name", "STRING"), ("種別", "item_type", "STRING"), ("技能", "skill", "STRING"),
        ("命中", "accuracy", "INT64"), ("攻撃力", "attack", "INT64"), ("ガード値", "guard", "INT64"),
        ("射程", "effect_range", "STRING"), ("ドッジ", "dodge", "INT64"), ("行動", "initiative", "INT64"),
        ("装甲値", "armor", "INT64"), ("常備化", "procurement", "INT64"), ("経験点", "exp_cost", "INT64"),
        ("効果", "effect_text", "STRING"),
        ("初出キャラID", "first_character_id", "STRING"), ("登録日", "registered_date", "DATE"),
    ]),
    "覚醒マスタ": ("awakening_master", [
        ("覚醒", "awakening", "STRING"), ("侵蝕値", "erosion", "INT64"), ("メモ", "memo", "STRING"),
        ("初出キャラID", "first_character_id", "STRING"), ("登録日", "registered_date", "DATE"),
    ]),
    "衝動マスタ": ("impulse_master", [
        ("衝動", "impulse", "STRING"), ("侵蝕値", "erosion", "INT64"), ("メモ", "memo", "STRING"),
        ("初出キャラID", "first_character_id", "STRING"), ("登録日", "registered_date", "DATE"),
    ]),
}


# ---------------------------------------------------------------------------
# 1. Extract / 2. Transform
# ---------------------------------------------------------------------------

def to_string(value):
    if pd.isna(value):
        return None
    if isinstance(value, float) and value.is_integer():  # 2.0 → "2"
        value = int(value)
    text = str(value).strip()
    return text or None


def to_bool(value):
    return {"はい": True, "いいえ": False}.get(to_string(value))


def convert(series, bq_type):
    if bq_type == "INT64":
        return pd.to_numeric(series, errors="coerce").round().astype("Int64")
    if bq_type == "BOOL":
        return series.map(to_bool).astype("boolean")
    if bq_type == "DATE":
        return pd.to_datetime(series, errors="coerce").dt.date
    return series.map(to_string).astype("object")


def extract(xlsx_path, include_examples):
    frames = {}
    for sheet, (table, columns) in TABLES.items():
        raw = pd.read_excel(xlsx_path, sheet_name=sheet, dtype=object)
        missing = [jp for jp, _, _ in columns if jp not in raw.columns]
        if missing:
            sys.exit(f"[エラー] シート「{sheet}」に見出し {missing} がありません")
        raw = raw[raw["キャラID"].notna()]  # キャラIDが空の行（未使用の行）は読まない
        df = pd.DataFrame({en: convert(raw[jp], t) for jp, en, t in columns})
        frames[table] = df.reset_index(drop=True)

    if not include_examples:
        # キャラ名が「（例）」で始まるキャラは書き方の見本なので除外する
        chars = frames["characters"]
        example_ids = set(chars.loc[chars["name"].fillna("").str.startswith("（例）"), "character_id"])
        frames = {t: df[~df["character_id"].isin(example_ids)].reset_index(drop=True) for t, df in frames.items()}
        if example_ids:
            print(f"[除外] 見本のキャラ: {sorted(example_ids)}")

    sheet_names = pd.ExcelFile(xlsx_path).sheet_names
    for sheet, (table, columns) in MASTER_TABLES.items():
        if sheet not in sheet_names:
            print(f"[注意] シート「{sheet}」が無いため、{table} は空で作ります（dx3_loader.py を一度実行すると作られます）")
            raw = pd.DataFrame(columns=[jp for jp, _, _ in columns])
        else:
            raw = pd.read_excel(xlsx_path, sheet_name=sheet, dtype=object)
        raw = raw[raw[columns[0][0]].notna()]  # 名称（キー）が空の行は読まない
        frames[table] = pd.DataFrame({en: convert(raw[jp], t) for jp, en, t in columns}).reset_index(drop=True)
    return frames


# ---------------------------------------------------------------------------
# 3. Check（データ品質チェック）
# ---------------------------------------------------------------------------

def check(frames):
    errors = []
    chars = frames["characters"]
    ids = set(chars["character_id"])

    dup = chars["character_id"][chars["character_id"].duplicated()].tolist()
    if dup:
        errors.append(f"キャラクター: キャラIDが重複しています {dup}")
    bad = [i for i in ids if not re.fullmatch(r"C\d{3,}", i)]
    if bad:
        errors.append(f"キャラクター: キャラIDの形式が C001 のようになっていません {bad}")
    if chars["name"].isna().any():
        errors.append("キャラクター: キャラ名が空の行があります")

    master_tables = {table for table, _ in MASTER_TABLES.values()}
    for table, df in frames.items():
        if table in master_tables:
            key = df.columns[0]
            dup = df[key][df[key].duplicated()].tolist()
            if dup:
                errors.append(f"{table}: 名称が重複しています {dup}")
            continue
        orphan = sorted(set(df["character_id"]) - ids)
        if table != "characters" and orphan:
            errors.append(f"{table}: キャラクター表に無いキャラIDを参照しています {orphan}")

    ab = frames["abilities"]
    for cid in ids:
        names = sorted(ab.loc[ab["character_id"] == cid, "ability"].dropna())
        if names != sorted(["肉体", "感覚", "精神", "社会"]):
            errors.append(f"abilities: {cid} の能力値が肉体/感覚/精神/社会の4行になっていません {names}")

    if (frames["skills"]["level"].dropna() < 0).any():
        errors.append("skills: Lv がマイナスの行があります")

    for table, df in frames.items():
        print(f"  {table:<17} {len(df):>4} 行")
    for e in errors:
        print(f"  [NG] {e}")
    return not errors


# ---------------------------------------------------------------------------
# 4. Load / 5. View
# ---------------------------------------------------------------------------

def load(frames, project):
    from google.cloud import bigquery  # --dry-run のときは不要なので、ここで読み込む

    client = bigquery.Client(project=project, location=LOCATION)
    dataset = bigquery.Dataset(f"{project}.{DATASET}")
    dataset.location = LOCATION
    client.create_dataset(dataset, exists_ok=True)

    for table, columns in list(TABLES.values()) + list(MASTER_TABLES.values()):
        job_config = bigquery.LoadJobConfig(
            schema=[bigquery.SchemaField(en, t) for _, en, t in columns],
            write_disposition="WRITE_TRUNCATE",  # 毎回まるごと入れ替える（何度実行しても同じ結果になる）
        )
        client.load_table_from_dataframe(frames[table], f"{project}.{DATASET}.{table}", job_config=job_config).result()
        print(f"  [ロード] {DATASET}.{table}")

    for sql_file in sorted(SQL_DIR.glob("*.sql")):
        sql = sql_file.read_text(encoding="utf-8").format(project=project, dataset=DATASET)
        client.query(sql).result()
        print(f"  [ビュー] {sql_file.name}")


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("xlsx", type=Path, nargs="?", default=BASE_DIR / "DX3_キャラ管理DB.xlsx",
                        help="キャラ管理DB の xlsx（省略時は作業フォルダの DX3_キャラ管理DB.xlsx）")
    parser.add_argument("--project", default="dx3-portfolio", help="Google Cloud のプロジェクトID")
    parser.add_argument("--include-examples", action="store_true", help="（例）のキャラも読み込む")
    parser.add_argument("--dry-run", action="store_true", help="BigQuery に送らず、output/ に CSV を出力する")
    args = parser.parse_args()

    print(f"[読込] {args.xlsx.name}")
    frames = extract(args.xlsx, args.include_examples)
    print("[チェック]")
    if not check(frames):
        sys.exit("[中止] データ品質チェックで問題が見つかったため、BigQuery には送りません")

    if args.dry_run:
        out = BASE_DIR / "output"
        out.mkdir(exist_ok=True)
        for table, df in frames.items():
            df.to_csv(out / f"{table}.csv", index=False, encoding="utf-8-sig")
        print(f"[出力] {out.resolve()} に CSV を保存しました（BigQuery には送っていません）")
    else:
        print(f"[BigQuery] {args.project}.{DATASET}（{LOCATION}）")
        load(frames, args.project)
    print("[完了]")


if __name__ == "__main__":
    main()
