# Python
import re
import os
import pandas as pd
import numpy as np
from datetime import datetime
from pathlib import Path

# 簡易 .env ローダー（外部依存なし）
# 優先順位: 既存の環境変数 > .env(CWD) > .env(リポジトリルート)
def _load_env_from_dotenv():
    def parse_and_set(dotenv_path: Path):
        if not dotenv_path.exists():
            return
        try:
            for line in dotenv_path.read_text(encoding="utf-8").splitlines():
                s = line.strip()
                if not s or s.startswith("#"):  # コメント/空行
                    continue
                if "=" not in s:
                    continue
                k, v = s.split("=", 1)
                k = k.strip()
                v = v.strip()
                # 値の両端ダブル/シングルクォートを剥がす
                if (v.startswith('"') and v.endswith('"')) or (v.startswith("'") and v.endswith("'")):
                    v = v[1:-1]
                # 既に環境変数に設定されている場合は上書きしない
                if k and (k not in os.environ or os.environ.get(k) is None):
                    os.environ[k] = v
        except Exception:
            # 読み込み失敗は致命ではないため無視（デフォルト値で継続）
            pass

    # カレントディレクトリの .env
    parse_and_set(Path.cwd() / ".env")
    # リポジトリルート（このファイルから2つ上）
    repo_root = Path(__file__).resolve().parents[2]
    parse_and_set(repo_root / ".env")

_load_env_from_dotenv()

# 1) 読み込み
df = pd.read_excel(Path(__file__).parent / "survey.xlsx", engine="openpyxl")

# 2) 「回答」列の例外的な処理（仕様変更後）
# 仕様：
# - 「回答」列（重複名: 回答, 回答.1 など）の左の列名を "元の設問名" とする
# - 左の列（設問列）の列名を "補足説明" + 元の設問名 に変更
# - 「回答」列の列名を 元の設問名 に変更（＝実データは直感的な設問名で参照できる）
# - 先頭列が「回答」の場合は左が存在しないため、そのままにする
# 実装ノート：
# - 元の列順（cols）から新しい列名リスト（new_names）を構成して一括 rename することで副作用を防ぐ
# - 他の列はそのまま

def map_answer_columns(frame: pd.DataFrame) -> pd.DataFrame:
    cols = list(frame.columns)
    new_names = list(cols)  # 初期は同名

    for i, c in enumerate(cols):
        if c == "回答" or re.match(r"^回答\.\d+$", str(c)):
            if i == 0:
                # 先頭が回答なら変更不可、スキップ
                continue
            original_left = cols[i - 1]
            # 左列は補足説明プレフィックスを付ける
            new_names[i - 1] = f"補足説明{original_left}"
            # 回答列は左列の元名にする
            new_names[i] = original_left

    # 重複名が発生する可能性はあるが仕様上許容する（必要なら後段で個別対応）
    rename_map = {old: new for old, new in zip(cols, new_names) if old != new}
    return frame.rename(columns=rename_map)

df = map_answer_columns(df)

# 3) 文字列のトリミング・NaN整備（最低限）
def strip_series(s: pd.Series) -> pd.Series:
    return s.astype(str).str.replace(r"\u3000", " ", regex=True).str.strip().replace({"nan": np.nan})
for c in df.columns:
    if df[c].dtype == object:
        df[c] = strip_series(df[c])

# 4) 生年月日を日時化（yyyyMMddやExcel数値に耐える）
def parse_birth(x):
    if pd.isna(x): return pd.NaT
    if isinstance(x, (int, float)) and not pd.isna(x):
        s = str(int(x))
        return pd.to_datetime(s, format="%Y%m%d", errors="coerce")
    if isinstance(x, str) and re.fullmatch(r"\d{8}", x):
        return pd.to_datetime(x, format="%Y%m%d", errors="coerce")
    return pd.to_datetime(x, errors="coerce")

df["birth_dt"] = df["生年月日"].apply(parse_birth)

# 5) 2024年度の「4/1時点学年」を算出
FISCAL_YEAR = 2024
APRIL1 = pd.Timestamp(f"{FISCAL_YEAR}-04-01")

def age_on(d, ref):
    if pd.isna(d): return np.nan
    return ref.year - d.year - ((ref.month, ref.day) < (d.month, d.day))

def grade_ja_on_april1(birth_dt):
    a = age_on(birth_dt, APRIL1)
    if pd.isna(a): return "不明"
    a = int(a)
    # 4/1時点の年齢 → 学年
    mapping = {
        6: "小1", 7: "小2", 8: "小3", 9: "小4", 10: "小5", 11: "小6",
        12: "中1", 13: "中2", 14: "中3"
    }
    return mapping.get(a, "対象外")

df["grade_2024"] = df["birth_dt"].apply(grade_ja_on_april1)

# 未就学児（2024/04/01時点で6歳未満）を除外するための年齢計算とフィルタ
# age_onはNaNを返すことがあるため、いったん列にしてから判定
df["age_2024"] = df["birth_dt"].apply(lambda d: age_on(d, APRIL1))
# 未就学児: 年齢が6歳未満、または学年が不明（生年月日不明等）に加えて「対象外」も除外
preschool_mask = (
    (df["age_2024"].notna() & (df["age_2024"] < 6))
    | (df["grade_2024"] == "不明")
    | (df["grade_2024"] == "対象外")
)

# 除外数（「組」=1行1組想定）
n_preschool = int(preschool_mask.sum())
# 集計に用いる有効データ
df_eff = df.loc[~preschool_mask].copy()

# 6) 地域区分（東京23区 / 三多摩島しょ / 埼玉 / 神奈川 / 千葉 / その他）
TOKYO_23 = {
    "千代田区","中央区","港区","新宿区","文京区","台東区","墨田区","江東区","品川区","目黒区",
    "大田区","世田谷区","渋谷区","中野区","杉並区","豊島区","北区","荒川区","板橋区","練馬区",
    "足立区","葛飾区","江戸川区"
}
def region_bucket(pref, city):
    if pref == "東京都":
        if isinstance(city, str) and any(city.startswith(ku) for ku in TOKYO_23):
            return "東京23区"
        return "三多摩島しょ"
    if pref == "埼玉県": return "埼玉県"
    if pref == "神奈川県": return "神奈川県"
    if pref == "千葉県": return "千葉県"
    return "その他"

df_eff["region_bucket"] = [region_bucket(p, c) for p, c in zip(df_eff.get("都道府県"), df_eff.get("市区町村"))]

# 7) 複数回答の縦持ち化（例: 情報経路・習い事）
def split_multiselect(series: pd.Series) -> pd.Series:
    return (
        series.fillna("")
        .str.replace("\r\n", "\n")
        .str.split(r"[\n]+", regex=True)
        .apply(lambda xs: [x.strip() for x in xs if x and x.strip()])
    )

# 列名は実ファイルに合わせてください（例に基づく想定）
col_channel = "本イベントを何でお知りになりましたか？（複数回答可）"
col_learning = "現在習い事や塾などに通われていますか？（複数回答可）"

df_eff["channel_list"] = split_multiselect(df_eff.get(col_channel))
df_eff["learning_list"] = split_multiselect(df_eff.get(col_learning))

channel_long = (
    df_eff[["性別", "region_bucket", "grade_2024", "channel_list"]]
    .explode("channel_list", ignore_index=True)
    .rename(columns={"channel_list": "channel"})
)
channel_long = channel_long[channel_long["channel"].notna() & (channel_long["channel"] != "")]

# 8) 集計例（全体／地域別／学年別）
# 全体トップN（情報経路）
top_channel_overall = (
    channel_long["channel"].value_counts().head(10)
)

# 地域別クロス（情報経路 × 地域）
channel_by_region = pd.crosstab(channel_long["channel"], channel_long["region_bucket"])

# 学年別（小1〜中3に限定）
grades_order = ["小1","小2","小3","小4","小5","小6","中1","中2","中3"]
channel_by_grade = (
    pd.crosstab(channel_long["channel"], channel_long["grade_2024"])
    [ [g for g in grades_order if g in channel_long["grade_2024"].unique()] ]
)

# 9) HTMLレポート（A4縦）出力 — サマリ（回答人数／男女比／小学校・中学校比）

# 設問一覧を返す関数
# - 除外: 性別, 生年月日, 郵便番号, 都道府県, 市区町村
# - 除外: 列名が「補足説明」で始まる列
# - 除外: 集計のためにプログラム上で作成したアルファベットの列（ASCII識別子の列名）
#   例: birth_dt, grade_2024, age_2024, region_bucket, channel_list, learning_list, gender_norm, school_level など
# - 入力DataFrame中の列順を維持して返す

def get_question_columns(frame: pd.DataFrame) -> list:
    excluded_exact = {"性別", "生年月日", "郵便番号", "都道府県", "市区町村"}

    def is_ascii_identifier(name: str) -> bool:
        # 先頭は英字またはアンダースコア、以降は英数字またはアンダースコアのみ
        return re.fullmatch(r"[A-Za-z_][A-Za-z0-9_]*", str(name)) is not None

    questions = []
    for col in frame.columns:
        col_str = str(col)
        if col_str in excluded_exact:
            continue
        if col_str.startswith("補足説明"):
            continue
        if is_ascii_identifier(col_str):
            continue
        questions.append(col_str)
    return questions

# 設問の選択肢一覧を返す関数
# - パラメータ: question_col = 設問の列名（文字列）
# - 仕様: 列内の文字列を正規化（trim）し、改行区切りも考慮して個別の選択肢に分解
#         非空の文字列をユニーク（初出順）にして返す
# - 備考: 列が存在しない場合は空リスト

def get_question_options(frame: pd.DataFrame, question_col: str) -> list:
    if question_col not in frame.columns:
        return []
    series = frame[question_col]
    # 文字列化とNaN除去
    series = series.dropna().astype(str)
    seen = set()
    options = []
    for cell in series:
        cell = cell.strip()
        if not cell:
            continue
        # 改行で分割（複数回答セルに対応）。改行が無ければそのまま1件として扱う
        parts = re.split(r"[\r\n]+", cell)
        for p in parts:
            s = p.strip()
            if not s:
                continue
            if s not in seen:
                seen.add(s)
                options.append(s)
    # ソート: 辞書順。ただし「その他」は常に最後に配置
    def sort_key(x: str):
        return (1 if x == "その他" else 0, x)
    options_sorted = sorted(options, key=sort_key)
    return options_sorted

def normalize_gender(x: str) -> str:
    if pd.isna(x) or str(x).strip() == "":
        return "未回答・その他"
    s = str(x)
    if "男" in s:
        return "男性"
    if "女" in s:
        return "女性"
    return "未回答・その他"

# 性別の正規化
if "性別" in df_eff.columns:
    df_eff["gender_norm"] = df_eff["性別"].apply(normalize_gender)
else:
    df_eff["gender_norm"] = "未回答・その他"

# 小学校/中学校の区分（grade_2024が小x/中xで判定）
def school_level_from_grade(g: str) -> str:
    if pd.isna(g):
        return "不明"
    g = str(g)
    if g.startswith("小"):
        return "小学校"
    if g.startswith("中"):
        return "中学校"
    return "不明"

df_eff["school_level"] = df_eff["grade_2024"].apply(school_level_from_grade)

# 集計（未就学児を除いた有効データに対して）
n_total = len(df_eff)

# 性別
gender_counts = df_eff["gender_norm"].value_counts().to_dict()
male = int(gender_counts.get("男性", 0))
female = int(gender_counts.get("女性", 0))
other = int(gender_counts.get("未回答・その他", 0))

def pct(n, d):
    return 0 if d == 0 else round(n * 100.0 / d, 1)

# 学校区分
level_counts = df_eff["school_level"].value_counts().to_dict()
prim = int(level_counts.get("小学校", 0))
mid = int(level_counts.get("中学校", 0))
unknown_lv = int(level_counts.get("不明", 0))

# HTML生成
now_str = datetime.now().strftime("%Y-%m-%d %H:%M")

a4_css = f"""
  @page {{ size: A4 portrait; }}  /* 余白はChromeデフォルトを使うためmargin指定はしない */
  /* html, body の固定高さは印刷でオーバーフローを誘発するため外す */
  html, body {{ /* height: 100%; 削除 */ }}
  body {{ font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', 'Hiragino Kaku Gothic ProN', 'Hiragino Sans', Meiryo, sans-serif; color: #222; }}
  /* A4の印字可能領域に収まるよう安全側の最大幅に。min-heightも外す */
  .page {{ max-width: 180mm; margin: 0 auto; background: white; }}
  h1 {{ font-size: 18pt; margin: 0 0 8mm; }}
  h2 {{ font-size: 13pt; margin: 6mm 0 3mm; border-bottom: 2px solid #eee; padding-bottom: 2mm; }}
  /* 章内の項目インデント */
  section > *:not(h2) {{ margin-left: 6mm; }}
  h3 {{ font-size: 11pt; margin: 3mm 0 2mm; }}
  .title {{ margin: 0 0 6mm; }}
  .title .line1 {{ font-size: 9pt; color: #555; }}
  .title .line2 {{ font-size: 20pt; font-weight: 800; margin-top: 1mm; }}
  .muted {{ color: #777; font-size: 8pt; }}
  .kpis {{ display: grid; grid-template-columns: repeat(3, 1fr); gap: 8mm; margin-bottom: 6mm; }}
  .kpi {{ border: 1px solid #e5e5e5; border-radius: 6px; padding: 6mm; }}
  .kpi .label {{ font-size: 9pt; color: #666; }}
  .kpi .value {{ font-size: 22pt; font-weight: 700; margin-top: 2mm; }}
  .bars {{ display: grid; grid-template-columns: 1fr; gap: 3mm; margin-top: 4mm; }}
  /* Visible track even when backgrounds are not printed */
  .bar {{ background: #f2f4f8; border: 1px solid #d0d7e2; border-radius: 999px; overflow: hidden; height: 10px; position: relative; }}
  /* Fill segment */
  .bar > span {{ display: block; height: 100%; background: #4c8bf5; }}
  .bar.secondary > span {{ background: #f58b4c; }}
  .bar.other > span {{ background: #b5b5b5; }}
  .legend {{ display: flex; gap: 6mm; flex-wrap: wrap; margin-top: 2mm; font-size: 8pt; color: #555; padding-left: 6mm; }}
  .legend .item::before {{ content: ''; display: inline-block; width: 10px; height: 10px; border-radius: 2px; margin-right: 4px; vertical-align: middle; }}
  .legend .male::before {{ background: #4c8bf5; }}
  .legend .female::before {{ background: #f58b4c; }}
  .legend .other::before {{ background: #b5b5b5; }}

  /* テーブル */
  table.simple {{ border-collapse: collapse; width: 90%; font-size: 10pt; margin-left: 6mm; }}
  table.simple th, table.simple td {{ border: 1px solid #e0e0e0; padding: 6px 8px; text-align: right; }}
  table.simple th {{ background: #f9fafb; color: #444; text-align: center; }}
  table.simple tfoot td {{ font-weight: 700; background: #fafafa; }}
  table.simple td.label {{ text-align: left; }}

  /* 選択肢×区分 割合バー */
  .option-pct td {{ vertical-align: middle; }}
  .pct-bar {{ position: relative; height: 12px; background: #f2f4f8; border: 1px solid #d0d7e2; border-radius: 6px; overflow: hidden; }}
  .pct-bar-fill {{ position: absolute; top: 0; left: 0; bottom: 0; background: #4c8bf5; }}
  .pct-bar-label {{ position: absolute; top: 50%; right: 6px; transform: translateY(-50%); font-size: 8pt; color: #333; text-shadow: 0 1px 0 rgba(255,255,255,0.6); }}

  /* 固定カラム幅（全ての.simpleテーブルで列幅を揃える）*/
  table.simple th:nth-child(1), table.simple td:nth-child(1) {{ width: 28%; }}
  table.simple th:nth-child(2), table.simple td:nth-child(2) {{ width: 18%; }}
  table.simple th:nth-child(3), table.simple td:nth-child(3) {{ width: 18%; }}
  table.simple th:nth-child(4), table.simple td:nth-child(4) {{ width: 18%; }}
  table.simple th:nth-child(5), table.simple td:nth-child(5) {{ width: 18%; }}

  /* 概要レイアウト */
  .overview-list {{ display: grid; grid-template-columns: 38mm 1fr; column-gap: 6mm; row-gap: 2mm; font-size: 10pt; }}
  .overview-list .label {{ color: #555; }}
  .overview-list .value {{ font-weight: 600; }}

  /* Ensure colors are preserved when printing */
  @media print {{
    * {{ -webkit-print-color-adjust: exact !important; print-color-adjust: exact !important; }}
    /* 印刷時も高さ固定を避ける */
    html, body {{ height: auto !important; }}
    .page {{ max-width: 180mm; min-height: auto; }}
    /* 末尾の余白で2ページ目に回り込まないように */
    .page > :last-child {{ margin-bottom: 0 !important; padding-bottom: 0 !important; }}

    .bar {{ background: #f2f4f8 !important; border-color: #d0d7e2 !important; }}
    .bar > span {{ background: #4c8bf5 !important; }}
    .bar.secondary > span {{ background: #f58b4c !important; }}
    .bar.other > span {{ background: #b5b5b5 !important; }}
  }}

  /* セクション改ページ */
  section.page-break {{
    break-before: page;
    page-break-before: always; /* Fallback for older engines */
  }}

  /* 設問セクション用 部品 */
  .supplement {{ font-size: 10pt; color: #555; margin: 2mm 0 2mm; white-space: pre-wrap; }}
  .note-box {{ display: block; max-width: 100%; padding: 3mm 5mm; border: 1px solid #e0e6ef; background: #f7f9fc; border-radius: 8px; }}
  .note-box h3 {{ margin: 0 0 2mm; }}
  .note-box .options {{ display: flex; flex-wrap: wrap; gap: 2mm 4mm; align-items: flex-start; }}
  .note-box .option-item {{ display: inline-flex; white-space: nowrap; font-size: 10pt; }}

  /* 設問用: 積み上げ横棒 */
  .q-subheading {{ font-size: 10pt; margin: 3mm 0 2mm; color: #333; }}
  .bar-row {{ display: flex; align-items: center; gap: 6mm; margin: 0.8mm 0 2.5mm 0; }}
  
  /* 外側ラベル対応のバーコンテナ */
  .bar-container {{ display: flex; align-items: center; gap: 6mm; margin: 0.8mm 0 2.5mm 0; position: relative; }}
  .bar-content {{ flex: 1 1 auto; position: relative; }}
  
  /* 外側ラベル領域 */
  .outside-labels-top {{ position: relative; margin-bottom: 0; }}
  .outside-labels-bottom {{ position: relative; margin-top: 0; }}
  
  /* 上下のラベル配置を調整 */
  .outside-labels-top .label-layer-1 {{ display: flex; align-items: end; }}  /* 下端揃え */
  .outside-labels-bottom .label-layer-1 {{ display: flex; align-items: start; }}  /* 上端揃え */
  
  /* ラベル層（動的多層対応） */
  .label-layer-1 {{ height: 12px; position: relative; }}
  .label-layer-2 {{ height: 16px; position: relative; }}
  .label-layer-3 {{ height: 16px; position: relative; }}
  .label-layer-4 {{ height: 16px; position: relative; }}
  .label-layer-5 {{ height: 16px; position: relative; }}
  .label-layer-6 {{ height: 16px; position: relative; }}
  .label-layer-7 {{ height: 16px; position: relative; }}
  .label-layer-8 {{ height: 16px; position: relative; }}
  .label-layer-9 {{ height: 16px; position: relative; }}
  .label-layer-10 {{ height: 16px; position: relative; }}
  
  /* 外側ラベル */
  .outside-label {{ position: absolute; font-size: 8pt; color: #333; background: rgba(255,255,255,0.9); 
    padding: 1px 4px; border-radius: 3px; border: 1px solid #ddd; white-space: nowrap; z-index: 10; }}
  
  /* リード線 */
  .leader-line {{ position: absolute; border-left: 1px solid #999; z-index: 5; }}
  .leader-line.to-top {{ bottom: 0; }}
  .leader-line.to-bottom {{ top: 0; }}
  
  .stacked-bar {{ flex: 1 1 auto; position: relative; height: 16px; border-radius: 8px; overflow: hidden; background: #f2f4f8; border: 1px solid #d0d7e2; }}
  .stacked-bar .seg {{ position: absolute; top: 0; bottom: 0; display: flex; align-items: center; justify-content: center; white-space: nowrap; font-size: 9pt; color: #fff; padding: 0 4px; }}
  .stacked-bar .seg .seg-label {{ font-weight: 600; text-shadow: 0 1px 0 rgba(0,0,0,0.25); }}
  .bar-right {{ flex: 0 0 22mm; font-size: 9pt; color: #444; text-align: right; }}
  .legend2 {{ display: flex; flex-wrap: wrap; gap: 5mm; font-size: 10pt; font-weight: 600; color: #333; margin: 1mm 0 2mm; padding-left: 6mm; }}
  .legend2 .item {{ display: inline-flex; align-items: center; gap: 4px; }}
  .legend2 .swatch {{ width: 10px; height: 10px; border-radius: 2px; display: inline-block; border: 1px solid rgba(0,0,0,0.05); }}
"""

male_pct = pct(male, n_total)
female_pct = pct(female, n_total)
other_pct = pct(other, n_total)

prim_pct = pct(prim, n_total)
mid_pct = pct(mid, n_total)
unknown_lv_pct = pct(unknown_lv, n_total)

# 男女 × 学校区分（小学校/中学校）クロス集計（未就学児除外データで）
rows_order = ["男性", "女性"]
cols_order = ["小学校", "中学校"]
ct = pd.crosstab(df_eff["gender_norm"], df_eff["school_level"])
ct = ct.reindex(index=rows_order, columns=cols_order, fill_value=0)
# 合計
row_totals = ct.sum(axis=1)
col_totals = ct.sum(axis=0)
grand_total = int(ct.values.sum())

# 各行の割合（総数に対する％）
row_pct = row_totals.apply(lambda n: pct(int(n), grand_total)) if grand_total else row_totals.apply(lambda n: 0)

# 地域別 × 学校区分（小学校/中学校）クロス集計
region_rows_order = ["東京23区", "三多摩島しょ", "埼玉県", "神奈川県", "千葉県", "その他"]
region_ct = pd.crosstab(df_eff["region_bucket"], df_eff["school_level"])
region_ct = region_ct.reindex(index=region_rows_order, columns=cols_order, fill_value=0)
region_row_totals = region_ct.sum(axis=1)
region_col_totals = region_ct.sum(axis=0)
# 行割合（総数に対する％）
region_row_pct = region_row_totals.apply(lambda n: pct(int(n), grand_total)) if grand_total else region_row_totals.apply(lambda n: 0)

# 数値のフォーマット
def fmt_int(n: int) -> str:
    return f"{int(n):,}"

# 地域別HTML行生成
region_rows_html = "\n".join([
    f"          <tr>\n            <td class=\"label\">{label}</td>\n            <td>{fmt_int(region_ct.loc[label, '小学校'])}</td>\n            <td>{fmt_int(region_ct.loc[label, '中学校'])}</td>\n            <td>{fmt_int(region_row_totals.loc[label])}</td>\n            <td>{region_row_pct.loc[label]}%</td>\n          </tr>"
    for label in region_rows_order
])

# HTML出力直前に設問一覧を取得してコンソール出力
question_columns = get_question_columns(df)
print("設問一覧（候補）:")
for q in question_columns:
    print(f"- {q}")

# 設問ごとの章（改ページあり）のHTMLを生成

def first_non_empty_value(series: pd.Series):
    if series is None:
        return None
    for v in series:
        if pd.isna(v):
            continue
        s = str(v).strip()
        if s:
            return s
    return None

def escape_html(s: str) -> str:
    # 最低限のエスケープ（選択肢や補足に <, >, & が含まれる場合に備える）
    return (
        str(s)
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
    )

# 補足説明の前処理（HTML埋め込み用）
# 仕様:
# - 入力テキストをまず escape_html で安全にエスケープ
# - その後、全角の開き括弧「（」の直前に <span style="white-space: nowrap;"> を付与
# - 全角の閉じ括弧「）」の直後に </span> を付与
# 備考:
# - 「（」「）」の数が一致しない場合、span が不均衡になる可能性がありますが、仕様通り付与します。
# - この関数は安全な HTML 文字列を返すため、呼び出し側で追加のエスケープは不要です。

def preprocess_supplement_html(raw: str) -> str:
    if raw is None:
        return ""
    escaped = escape_html(raw)
    # 「（」の前にノーブレーク用 span 開始タグを置く（文字自体は保持）
    escaped = escaped.replace("（", "<span style=\"white-space: nowrap;\">（")
    # 「）」の後に span の終了タグを置く
    escaped = escaped.replace("）", "）</span>")
    return escaped

# ---- 設問集計用ヘルパ ----
REGION_ORDER = ["東京23区", "三多摩島しょ", "埼玉県", "神奈川県", "千葉県", "その他"]
GRADE_ORDER = ["小1", "小2", "小3", "小4", "小5", "小6", "中1", "中2", "中3"]

# 積み上げ棒グラフ設定
MIN_SEGMENT_WIDTH_PCT = 1.0  # セグメントの最小幅（％）
OUTSIDE_LABEL_THRESHOLD_PCT = 24.0  # 外側ラベル表示閾値（％）
OUTSIDE_LABEL_WITH_INNER_PCT_THRESHOLD = 5.0  # 外側ラベル+内側割合表示の閾値（％）

# セルから一人分の選択肢セット（重複正規化済み）を取得
# - 改行区切りを分割し、空を除去し、同一セル内重複を1つにする
# - 返却は set

def cell_to_unique_set(val) -> set:
    if pd.isna(val):
        return set()
    s = str(val).strip()
    if not s:
        return set()
    parts = re.split(r"[\r\n]+", s)
    uniq = []
    seen = set()
    for p in parts:
        t = p.strip()
        if not t:
            continue
        if t not in seen:
            seen.add(t)
            uniq.append(t)
    return set(uniq)

# 複数回答可の推定
# - 見出しに「複数」などの語が含まれる場合は True
# - それ以外でも、実データで1セル内に2つ以上の選択が存在する場合は True

def is_multiselect(frame: pd.DataFrame, qcol: str) -> bool:
    name = str(qcol)
    if ("複数" in name) or ("複数回答" in name):
        return True
    series = frame[qcol].dropna()
    for v in series.head(200):  # 全件でなくても傾向は分かる
        if len(cell_to_unique_set(v)) >= 2:
            return True
    return False

# 集計（1グループ）: counts辞書とSを返す

def aggregate_group(frame: pd.DataFrame, qcol: str, options: list):
    opt_set = set(options)
    counts = {opt: 0 for opt in options}
    S = 0
    for v in frame[qcol].dropna():
        chosen = [o for o in cell_to_unique_set(v) if o in opt_set]
        if not chosen:
            continue
        for o in chosen:
            counts[o] += 1
        S += len(chosen)
    return counts, S

# オプションの色割り当て（設問内で一貫）
# - 「その他」は常に #b5b5b5
# - それ以外はパレットを順番に

def color_map_for_options(options: list) -> dict:
    palette = [
        "#4c8bf5", "#f58b4c", "#57b26a", "#9166cc", "#e04f5f",
        "#39c0cf", "#f2c94c", "#7f8c8d", "#2ecc71", "#e67e22",
        "#9b59b6", "#1abc9c", "#e84393", "#0984e3", "#6c5ce7"
    ]
    # その他は最後扱い
    base = [o for o in options if o != "その他"]
    cmap = {}
    i = 0
    for o in base:
        cmap[o] = palette[i % len(palette)]
        i += 1
    if "その他" in options:
        cmap["その他"] = "#b5b5b5"
    return cmap

# オプション並び順: 全体割合降順。同率は辞書順。「その他」は常に最後

def order_options_by_overall(counts: dict, S_overall: int) -> list:
    opts = list(counts.keys())
    def key(o):
        pct_val = 0.0 if S_overall == 0 else counts[o] / S_overall
        return (1 if o == "その他" else 0, -pct_val, o)
    return sorted(opts, key=key)

# ラベル幅推定機能
def estimate_label_width_px(text: str, font_size_pt: int = 8) -> float:
    """
    ラベルテキストの推定幅をピクセル単位で計算
    日本語文字を考慮した簡易的な幅推定
    """
    if not text:
        return 0.0
    
    # フォントサイズをピクセルに変換（1pt ≈ 1.33px）
    font_size_px = font_size_pt * 1.33
    
    # 文字種別による幅係数
    ascii_count = sum(1 for c in text if ord(c) < 128)
    japanese_count = len(text) - ascii_count
    
    # ASCII文字: フォントサイズの約0.6倍, 日本語文字: フォントサイズの約1倍
    estimated_width = (ascii_count * font_size_px * 0.6) + (japanese_count * font_size_px * 1.0)
    
    # パディング（左右4px + 境界線等）を追加
    padding = 10
    
    return estimated_width + padding

def estimate_label_width_percent(text: str, bar_width_px: float, font_size_pt: int = 8) -> float:
    """
    ラベルテキストの幅を棒グラフ全体に対する割合（%）で計算
    """
    if bar_width_px <= 0:
        return 0.0
    
    width_px = estimate_label_width_px(text, font_size_pt)
    return (width_px / bar_width_px) * 100.0

# 改良されたFlexbox風配置アルゴリズム
def calculate_flexbox_positions(labels_data: list, available_width_percent: float = 100.0) -> list:
    """
    改良されたFlexbox風のラベル配置を計算
    - 重なりがない場合は元の位置を保持
    - 重なりがある場合のみ最小限の調整を実行
    Args:
        labels_data: [(center_pos_percent, text, option), ...]
        available_width_percent: 利用可能な幅（%）
    Returns:
        [(adjusted_pos_percent, text, option), ...]
    """
    if not labels_data:
        return []
    
    # 1つだけの場合は元の位置をそのまま保持
    if len(labels_data) == 1:
        return labels_data
    
    # 各ラベルの推定幅を計算（棒グラフ幅を180mm ≈ 680pxと仮定）
    bar_width_px = 680.0
    label_widths = []
    
    for center_pos, text, option in labels_data:
        width_percent = estimate_label_width_percent(text, bar_width_px)
        label_widths.append(width_percent)
    
    # 重なりを検出
    positions_and_widths = []
    for i, (center_pos, text, option) in enumerate(labels_data):
        width = label_widths[i]
        left_edge = center_pos - width / 2
        right_edge = center_pos + width / 2
        positions_and_widths.append((center_pos, left_edge, right_edge, text, option))
    
    # 重なりがあるかチェック（余裕を持たせて判定）
    OVERLAP_MARGIN = 2.0  # ラベル間の最小マージン（%）
    has_significant_overlap = False
    
    for i in range(len(positions_and_widths)):
        for j in range(i + 1, len(positions_and_widths)):
            left1, right1 = positions_and_widths[i][1], positions_and_widths[i][2]
            left2, right2 = positions_and_widths[j][1], positions_and_widths[j][2]
            # マージンを考慮した重なり判定
            if not (right1 + OVERLAP_MARGIN <= left2 or right2 + OVERLAP_MARGIN <= left1):
                has_significant_overlap = True
                break
        if has_significant_overlap:
            break
    
    # 重なりがない、または軽微な場合は元の位置を保持
    if not has_significant_overlap:
        return [(pos, text, option) for pos, _, _, text, option in positions_and_widths]
    
    # 重なりがある場合: 最小限の調整で解決を試行
    # まず、元の位置順序を保持しつつ、最小限の移動で重なりを解消
    sorted_positions = sorted(positions_and_widths, key=lambda x: x[0])  # center_posでソート
    
    adjusted_positions = []
    for i, (original_center, _, _, text, option) in enumerate(sorted_positions):
        width = label_widths[labels_data.index((original_center, text, option))]
        
        if i == 0:
            # 最初のラベル: 左端制約のみ考慮
            pos = max(width / 2, original_center)
        else:
            # 前のラベルとの重なりを避ける最小位置
            prev_pos = adjusted_positions[i-1][0]
            prev_width = label_widths[labels_data.index((sorted_positions[i-1][0], sorted_positions[i-1][3], sorted_positions[i-1][4]))]
            min_pos = prev_pos + prev_width / 2 + width / 2 + OVERLAP_MARGIN
            
            # 元の位置と最小位置の大きい方を選択（可能な限り元位置に近づける）
            pos = max(min_pos, original_center)
            
            # ただし、右端を超える場合は制限
            if pos + width / 2 > available_width_percent:
                pos = available_width_percent - width / 2
        
        adjusted_positions.append((pos, text, option))
    
    return adjusted_positions

# 1本の積み上げ棒HTMLを生成

def render_stacked_bar(title: str, counts: dict, order: list[str], colors: dict, unit: str, show_total_right: bool = True) -> str:
    S = sum(counts.values())
    if S == 0:
        return ""
    
    # 1. 実際の割合を計算
    actual_widths = {}
    for o in order:
        c = counts.get(o, 0)
        if c > 0:
            actual_widths[o] = (c / S) * 100.0
    
    # 2. 外側ラベル対象を判定
    outside_labels = []
    inside_segments = []
    
    for o in order:
        if o not in actual_widths:
            continue
        actual_w = actual_widths[o]
        if actual_w < OUTSIDE_LABEL_THRESHOLD_PCT:
            outside_labels.append(o)
        else:
            inside_segments.append(o)
    
    # 3. 最小幅保証の調整（内側ラベルのセグメントに対して）
    adjusted_widths = {}
    adjustment_needed = 0.0
    
    for o, actual_w in actual_widths.items():
        if o in inside_segments and actual_w < MIN_SEGMENT_WIDTH_PCT:
            adjusted_widths[o] = MIN_SEGMENT_WIDTH_PCT
            adjustment_needed += MIN_SEGMENT_WIDTH_PCT - actual_w
        else:
            adjusted_widths[o] = actual_w
    
    # 4. 最小幅保証により増えた分を、他のセグメントから比例配分で減らす
    if adjustment_needed > 0:
        reducible_total = sum(w for o, w in actual_widths.items() if w >= MIN_SEGMENT_WIDTH_PCT)
        
        if reducible_total > 0:
            reduction_ratio = adjustment_needed / reducible_total
            for o in adjusted_widths:
                if actual_widths[o] >= MIN_SEGMENT_WIDTH_PCT:
                    adjusted_widths[o] = actual_widths[o] * (1 - reduction_ratio)
    
    # 5. 外側ラベル配置の計算（新しいFlexbox風アルゴリズム）
    def calculate_outside_labels_html():
        if not outside_labels:
            return "", ""
        
        # 各外側ラベルの基本情報を集める
        label_data = []
        for o in outside_labels:
            # セグメント中央位置を計算
            left_pos = 0.0
            for prev_o in order:
                if prev_o == o:
                    break
                if prev_o in adjusted_widths:
                    left_pos += adjusted_widths[prev_o]
            
            center_pos = left_pos + (adjusted_widths.get(o, 0) / 2)
            label_pct = round((counts.get(o, 0) / S) * 100.0, 1)
            
            # 外側ラベルの表示内容を判定
            if label_pct >= OUTSIDE_LABEL_WITH_INNER_PCT_THRESHOLD:
                # 5%以上: 外側は選択肢名のみ、内側に割合表示
                label_text = f"{escape_html(o)}"
            else:
                # 5%未満: 外側に選択肢名+割合表示
                label_text = f"{escape_html(o)} {label_pct}%"
            
            label_data.append((center_pos, label_text, o))
        
        # Flexbox風配置で位置調整
        adjusted_label_data = calculate_flexbox_positions(label_data, 100.0)
        
        # 調整後の配置で層分けを実行
        # 可能な限り同じ高さ（layer1）に配置し、重なりがある場合のみ分散
        from collections import defaultdict
        top_layers = defaultdict(list)
        bottom_layers = defaultdict(list)
        
        # まずは全て上layer1に配置を試行
        for i, (pos, text, option) in enumerate(adjusted_label_data):
            if i < len(adjusted_label_data) // 2:  # 前半は上側
                top_layers[1].append((pos, text, option))
            else:  # 後半は下側
                bottom_layers[1].append((pos, text, option))
        
        # HTML生成：動的に層を構築
        def render_label_layer(labels, layer_num, has_leader_line=False, line_direction=""):
            if not labels:
                return ""
            layer_class = f"label-layer-{layer_num}"
            layer_html = f'<div class="{layer_class}">'
            for pos, text, option in labels:
                label_style = f"left:{pos:.2f}%; transform:translateX(-50%);"
                layer_html += f'<div class="outside-label" style="{label_style}">{text}</div>'
                if has_leader_line:
                    line_style = f"left:{pos:.2f}%; height:100%;"
                    layer_html += f'<div class="leader-line {line_direction}" style="{line_style}"></div>'
            layer_html += '</div>'
            return layer_html
        
        # 上側の層を逆順で配置（遠い層から先に配置）
        top_html = ""
        max_top_layer = max(top_layers.keys()) if top_layers else 0
        for layer_num in range(max_top_layer, 0, -1):  # 大きい層番号から小さい層番号へ
            if layer_num in top_layers:
                has_leader = (layer_num > 1)  # layer1のみリード線なし
                top_html += render_label_layer(top_layers[layer_num], layer_num, has_leader, "to-bottom")
        
        # 下側の層を正順で配置（近い層から先に配置）
        bottom_html = ""
        max_bottom_layer = max(bottom_layers.keys()) if bottom_layers else 0
        for layer_num in range(1, max_bottom_layer + 1):  # 小さい層番号から大きい層番号へ
            if layer_num in bottom_layers:
                has_leader = (layer_num > 1)  # layer1のみリード線なし
                bottom_html += render_label_layer(bottom_layers[layer_num], layer_num, has_leader, "to-top")
        
        top_container = f'<div class="outside-labels-top">{top_html}</div>' if top_html else ""
        bottom_container = f'<div class="outside-labels-bottom">{bottom_html}</div>' if bottom_html else ""
        
        return top_container, bottom_container
    
    top_labels_html, bottom_labels_html = calculate_outside_labels_html()
    
    # 6. セグメントHTML生成（内側ラベルのみ）
    left = 0.0
    segs = []
    for o in order:
        if o not in adjusted_widths:
            continue
        
        c = counts.get(o, 0)
        w = adjusted_widths[o]
        label_pct = round((c / S) * 100.0, 1)
        
        style = f"left:{left:.6f}%;width:{w:.6f}%;background:{colors.get(o,'#999')};"
        
        # 内側ラベル表示判定
        if o in inside_segments:
            # 内側セグメント: 選択肢名+割合表示（従来通り）
            segs.append(f"<div class=\"seg\" style=\"{style}\" title=\"{escape_html(o)} {label_pct}%\"><span class=\"seg-label\">{escape_html(o)} {label_pct}%</span></div>")
        else:
            # 外側ラベル対象
            if label_pct >= OUTSIDE_LABEL_WITH_INNER_PCT_THRESHOLD:
                # 10%以上: 棒内部に割合のみ表示
                segs.append(f"<div class=\"seg\" style=\"{style}\" title=\"{escape_html(o)} {label_pct}%\"><span class=\"seg-label\">{label_pct}%</span></div>")
            else:
                # 10%未満: 棒内部にラベル非表示
                segs.append(f"<div class=\"seg\" style=\"{style}\" title=\"{escape_html(o)} {label_pct}%\"></div>")
        
        left += w
    
    s_text = f"{S:,}{unit}"
    
    # 7. 最終HTML構造
    if outside_labels:
        right_html = f'<div class="bar-right">{s_text}</div>' if show_total_right else ''
        return f"""<div class="bar-container">
  <div class="bar-content">
    {top_labels_html}
    <div class="stacked-bar">{''.join(segs)}</div>
    {bottom_labels_html}
  </div>
  {right_html}
</div>"""
    else:
        # 外側ラベルがない場合は従来通り
        if show_total_right:
            return f'<div class="bar-row"><div class="stacked-bar">{"".join(segs)}</div><div class="bar-right">{s_text}</div></div>'
        else:
            return f'<div class="bar-row"><div class="stacked-bar">{"".join(segs)}</div></div>'

# グループごとの棒群HTML（見出し＋棒複数 or データなし）

def render_group_bars(group_label: str, frames: list, qcol: str, order: list, colors: dict, unit: str) -> str:
    bars = []
    for name, fr in frames:
        counts, S = aggregate_group(fr, qcol, order)
        if S > 0:
            bars.append(f"<div><div class=\"q-subheading\">{escape_html(name)}</div>{render_stacked_bar(name, counts, order, colors, unit)}</div>")
    if not bars:
        return f"<div class=\"q-subheading\">{escape_html(group_label)}</div><div class=\"muted\">データなし</div>"
    # group_label as a heading, then stacked bars listed vertically
    inner = []
    # 単位から表示サフィックス（人/回）を決定
    suffix = "回" if unit.endswith("回中") else "人"
    for name, fr in frames:
        counts, S = aggregate_group(fr, qcol, order)
        if S == 0:
            continue
        bar_html = render_stacked_bar(name, counts, order, colors, unit, show_total_right=False)
        label_text = f"{escape_html(name)} = {S:,}{suffix}"
        inner.append(f"<div><div class=\"muted\" style=\"margin-bottom:0.3mm;\">{label_text}</div>{bar_html}</div>")
    return f"<div class=\"q-subheading\">{escape_html(group_label)}</div>" + "".join(inner)

# 伝説（凡例）

def render_legend(order: list[str], colors: dict) -> str:
    items = []
    for o in order:
        items.append(f"<span class=\"item\"><span class=\"swatch\" style=\"background:{colors.get(o,'#999')}\"></span>{escape_html(o)}</span>")
    return f"<div class=\"legend2\">{''.join(items)}</div>"

# 指定したフレーム群に対して、選択肢ごとのカウント表を生成
# columns: [sub_label] + options
# rows: 先頭に「全体」、続いて各カテゴリ

def render_option_count_table(sub_label: str, header_label: str, frames: list[tuple[str, pd.DataFrame]], qcol: str, options: list[str]) -> str:
    # 列ヘッダー（転置版）: 選択肢 | 全体 | 各カテゴリ名
    col_headers = ["選択肢", "全体"] + [escape_html(name) for name, _ in frames]
    thead = "<thead><tr>" + "".join([f"<th>{h}</th>" for h in col_headers]) + "</tr></thead>"

    # 全体のフレーム（全カテゴリ結合）
    all_frame = pd.concat([fr for _, fr in frames], axis=0) if frames else df_eff
    overall_counts, _S_overall = aggregate_group(all_frame, qcol, options)
    # 分母（回答者数）: 設問によらず、全体は n_total、各カテゴリは len(fr)
    overall_denom = n_total

    # 各カテゴリの集計を事前計算
    per_frame_counts = []  # [(name, counts_dict, denom)]
    for name, fr in frames:
        counts, _S = aggregate_group(fr, qcol, options)
        denom = len(fr)
        per_frame_counts.append((name, counts, denom))

    # 行: 各選択肢
    body_rows = []
    for o in options:
        tds = [f"<td class=\"label\">{escape_html(o)}</td>"]
        # 全体（分母: 回答者数）
        ov_num = overall_counts.get(o, 0)
        ov_pct = 0 if overall_denom == 0 else round(ov_num * 100.0 / overall_denom, 1)
        tds.append(f"<td>{fmt_int(ov_num)}<div class=\"muted\" style=\"font-size:8pt;\">{ov_pct}%</div></td>")
        # 各カテゴリ（分母: そのカテゴリの回答者数）
        for name, counts, denom in per_frame_counts:
            num = counts.get(o, 0)
            pct_val = 0 if denom == 0 else round(num * 100.0 / denom, 1)
            tds.append(f"<td>{fmt_int(num)}<div class=\"muted\" style=\"font-size:8pt;\">{pct_val}%</div></td>")
        body_rows.append("<tr>" + "".join(tds) + "</tr>")

    tbody = "<tbody>" + "".join(body_rows) + "</tbody>"

    return f"<div class=\"q-subheading\">{escape_html(sub_label)}</div><table class=\"simple\">{thead}{tbody}</table>"

# 選択肢ごと × 区分ごとの割合を横棒で示すテーブル（A/B/Cレイアウト風）
# - 3列: 選択肢 | 区分（人数） | 割合（横棒）
# - 各選択肢で行グループ化（rowspan）

def render_option_category_pct_table(sub_label: str, frames: list[tuple[str, pd.DataFrame]], qcol: str, options: list[str], colors: dict) -> str:
    if not frames or not options:
        return ""

    # ヘッダ
    thead = (
        "<thead><tr>"
        "<th>選択肢</th>"
        "<th>区分（人数）</th>"
        "<th>割合</th>"
        "</tr></thead>"
    )

    # 事前計算（各区分の分母）
    frame_denoms = [(name, len(fr)) for name, fr in frames]
    # 各区分ごとの選択肢カウント
    per_frame_counts = []  # (name, counts)
    for name, fr in frames:
        counts, _S = aggregate_group(fr, qcol, options)
        per_frame_counts.append((name, counts))

    # name -> (counts, denom) lookup
    counts_map = {name: counts for name, counts in per_frame_counts}
    denom_map = {name: denom for name, denom in frame_denoms}

    # 行生成
    body_rows = []
    for o in options:
        first = True
        # 表示対象の行数（全区分を表示。必要なら0件も表示）
        for name, _ in frames:
            denom = denom_map.get(name, 0)
            num = counts_map.get(name, {}).get(o, 0)
            pct_val = 0 if denom == 0 else round(num * 100.0 / denom, 1)
            # A列
            if first:
                a_cell = f"<td class=\"label\" rowspan=\"{len(frames)}\">{escape_html(o)}</td>"
                first = False
            else:
                a_cell = ""
            # B列（区分名 = 人数）
            b_text = f"{escape_html(name)} = {fmt_int(num)}人"
            b_cell = f"<td>{b_text}</td>"
            # C列（横棒）
            bar_color = colors.get(o, "#4c8bf5")
            c_cell = (
                f"<td>"
                f"  <div class=\"pct-bar\">"
                f"    <div class=\"pct-bar-fill\" style=\"width:{pct_val}%;background:{bar_color};\"></div>"
                f"    <div class=\"pct-bar-label\">{pct_val}%</div>"
                f"  </div>"
                f"</td>"
            )
            body_rows.append("<tr>" + a_cell + b_cell + c_cell + "</tr>")

    tbody = "<tbody>" + "".join(body_rows) + "</tbody>"

    return f"<div class=\"q-subheading\">{escape_html(sub_label)}</div><table class=\"simple option-pct\">{thead}{tbody}</table>"

sections = []
for idx, q in enumerate(question_columns):
    # 補足説明列が存在すれば、最初の非空データを拾って表示（角丸矩形の中に配置）
    supp_col = f"補足説明{q}"
    inner_sup = ""
    if supp_col in df.columns:
        first_val = first_non_empty_value(df[supp_col])
        if first_val is not None:
            inner_sup = f"<h3 class=\"supplement\">{preprocess_supplement_html(first_val)}</h3>"

    # 設問の選択肢（ユニーク）を取得して列挙
    opts = get_question_options(df, q)
    def alpha_label(i: int) -> str:
        # A..Z, それ以降はAA, AB...（簡易実装）
        letters = []
        i0 = i
        while True:
            letters.append(chr(ord('A') + (i0 % 26)))
            i0 = i0 // 26 - 1
            if i0 < 0:
                break
        return "".join(reversed(letters))

    options_html = ""
    if opts:
        items = []
        for i, opt in enumerate(opts):
            label = alpha_label(i)
            items.append(f"<span class=\"option-item\">【{label}】{escape_html(opt)}</span>")
        options_html = f"<div class=\"options\">{''.join(items)}</div>"

    note_box_html = f"<div class=\"note-box\">{inner_sup}<h3>選択肢</h3>{options_html}</div>"

    # オプションなしの場合はスキップ（安全）
    if not opts:
        section_html = f"""
        <section class=\"page-break\">
          <h2>Q{idx+1} {q}</h2>
          {note_box_html}
          <div class=\"muted\">データなし</div>
        </section>
        """.strip()
        sections.append(section_html)
        continue

    # 集計：全体
    overall_counts, S_overall = aggregate_group(df_eff, q, opts)
    # 並び順（全体割合降順。同率は辞書順。その他は最後）
    order = order_options_by_overall(overall_counts, S_overall)
    # 色割当（設問内で固定）
    colors = color_map_for_options(order)

    # 単一/複数の判定と単位・説明文
    multi = is_multiselect(df_eff, q)
    unit = "回中" if multi else "人中"
    if multi:
        explain_text = "以下のグラフの割合は、各区分の選択回数の合計を母数とし、合計は100%になります\n棒の右の数値は、その区分の合計回数です。"
    else:
        explain_text = "以下のグラフの割合は、各区分の回答者数を母数にしています。棒の右の数値は、その区分の合計人数です。"

    # 全体棒（他と同じネスト構造にして開始位置を揃える）
    overall_bar_html = ""
    if S_overall > 0:
        # 全体も他のサブセクションと同様にラベルに合計を表示し、右側の合計は非表示
        overall_suffix = "回" if unit.endswith("回中") else "人"
        overall_label = f"全体 = {S_overall:,}{overall_suffix}"
        overall_bar_html = f"<div><div class=\"muted\" style=\"margin-bottom:0.3mm;\">{overall_label}</div>{render_stacked_bar('全体', overall_counts, order, colors, unit, show_total_right=False)}</div>"
    else:
        overall_bar_html = "<div><div class=\"muted\">データなし</div></div>"

    legend_html = render_legend(order, colors)
    explain_html = f"<div class=\"muted\" style=\"margin:2mm 0 2mm; white-space: pre-wrap;\">{escape_html(explain_text)}</div>"

    # 地域別
    region_frames = [(lab, df_eff[df_eff["region_bucket"] == lab]) for lab in REGION_ORDER]
    region_html = render_group_bars("地域別", region_frames, q, order, colors, unit)

    # 学年別
    grade_frames = [(lab, df_eff[df_eff["grade_2024"] == lab]) for lab in GRADE_ORDER]
    grade_html = render_group_bars("学年別", grade_frames, q, order, colors, unit)

    # テーブル（角丸矩形の直下に配置）
    region_pct_table_html = render_option_category_pct_table("地域別（選択肢×地域の割合）", region_frames, q, order, colors)
    region_table_html = render_option_count_table("地域別", "地域", region_frames, q, order)
    grade_table_html = render_option_count_table("学年別", "学年", grade_frames, q, order)

    section_html = f"""
    <section class=\"page-break\">
      <h2>Q{idx+1} {escape_html(q)}</h2>
      {note_box_html}
      {region_pct_table_html}
      {region_table_html}
      {grade_table_html}
      {legend_html}
      {explain_html}
      <div class=\"q-subheading\">全体</div>
      {overall_bar_html}
      {region_html}
      {grade_html}
    </section>
    """.strip()
    sections.append(section_html)

question_sections_html = "\n".join(sections)

# 先頭ページの文言を環境変数から取得（.env で設定可能）
organizer = os.getenv("REPORT_ORGANIZER", "サンプル主催者")
survey_name = os.getenv("REPORT_SURVEY_NAME", "サンプルイベント名")
participating_schools = os.getenv("REPORT_PARTICIPATING_SCHOOLS", "参加校 100校")
venue = os.getenv("REPORT_VENUE", "サンプル会場 A")
event_dates = os.getenv("REPORT_EVENT_DATES", "9月1日（日）")

html = f"""
<!doctype html>
<html lang=\"ja\">
<head>
  <meta charset=\"utf-8\" />
  <meta name=\"viewport\" content=\"width=device-width, initial-scale=1\" />
  <title>サマリレポート</title>
  <style>
  {a4_css}
  </style>
</head>
<body>
  <div class=\"page\">
    <header>
      <div class="title">
        <div class="line1">{organizer}</div>
        <div class="line2">{survey_name}</div>
      </div>
    </header>

    <section>
      <h2>概要</h2>
      <div class=\"overview-list\">
        <div class=\"label\">参加校</div>
        <div class=\"value\">{participating_schools}</div>
        <div class=\"label\">会場</div>
        <div class=\"value\">{venue}</div>
        <div class=\"label\">開催日</div>
        <div class=\"value\">{event_dates}</div>
        <div class=\"label\">アンケート回答数</div>
        <div class=\"value\">{n_total:,}組（未就学児{n_preschool:,}組を除く）</div>
      </div>
    </section>

    <section>
      <h2>回答者属性</h2>
      <h3>男女別</h3>
      <table class="simple">
        <thead>
          <tr>
            <th></th>
            <th>小学校</th>
            <th>中学校</th>
            <th>合計</th>
            <th>割合</th>
          </tr>
        </thead>
        <tbody>
          <tr>
            <td class="label">男性</td>
            <td>{fmt_int(ct.loc['男性','小学校'])}</td>
            <td>{fmt_int(ct.loc['男性','中学校'])}</td>
            <td>{fmt_int(row_totals.loc['男性'])}</td>
            <td>{row_pct.loc['男性']}%</td>
          </tr>
          <tr>
            <td class="label">女性</td>
            <td>{fmt_int(ct.loc['女性','小学校'])}</td>
            <td>{fmt_int(ct.loc['女性','中学校'])}</td>
            <td>{fmt_int(row_totals.loc['女性'])}</td>
            <td>{row_pct.loc['女性']}%</td>
          </tr>
        </tbody>
        <tfoot>
          <tr>
            <td class="label">合計</td>
            <td>{fmt_int(col_totals.loc['小学校'])}</td>
            <td>{fmt_int(col_totals.loc['中学校'])}</td>
            <td>{fmt_int(grand_total)}</td>
            <td>100%</td>
          </tr>
        </tfoot>
      </table>

      <h3>地域別</h3>
      <table class="simple">
        <thead>
          <tr>
            <th></th>
            <th>小学校</th>
            <th>中学校</th>
            <th>合計</th>
            <th>割合</th>
          </tr>
        </thead>
        <tbody>
{region_rows_html}
        </tbody>
        <tfoot>
          <tr>
            <td class="label">合計</td>
            <td>{fmt_int(region_col_totals.loc['小学校'])}</td>
            <td>{fmt_int(region_col_totals.loc['中学校'])}</td>
            <td>{fmt_int(grand_total)}</td>
            <td>100%</td>
          </tr>
        </tfoot>
      </table>
    </section>

    {question_sections_html}

  </div>
</body>
</html>
"""

# 保存
with open("report.html", "w", encoding="utf-8") as f:
    f.write(html)

print(f"HTMLレポートを出力しました: report.html  （{n_total}件、未就学児{n_preschool}組を除く）")