# Python
import re
import pandas as pd
import numpy as np
from datetime import datetime
from pathlib import Path

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
  h1 {{ font-size: 20pt; margin: 0 0 8mm; }}
  h2 {{ font-size: 14pt; margin: 6mm 0 3mm; border-bottom: 2px solid #eee; padding-bottom: 2mm; }}
  /* 章内の項目インデント */
  section > *:not(h2) {{ margin-left: 6mm; }}
  h3 {{ font-size: 12pt; margin: 3mm 0 2mm; }}
  .title {{ margin: 0 0 6mm; }}
  .title .line1 {{ font-size: 10pt; color: #555; }}
  .title .line2 {{ font-size: 22pt; font-weight: 800; margin-top: 1mm; }}
  .muted {{ color: #777; font-size: 9pt; }}
  .kpis {{ display: grid; grid-template-columns: repeat(3, 1fr); gap: 8mm; margin-bottom: 6mm; }}
  .kpi {{ border: 1px solid #e5e5e5; border-radius: 6px; padding: 6mm; }}
  .kpi .label {{ font-size: 10pt; color: #666; }}
  .kpi .value {{ font-size: 24pt; font-weight: 700; margin-top: 2mm; }}
  .bars {{ display: grid; grid-template-columns: 1fr; gap: 3mm; margin-top: 4mm; }}
  /* Visible track even when backgrounds are not printed */
  .bar {{ background: #f2f4f8; border: 1px solid #d0d7e2; border-radius: 999px; overflow: hidden; height: 10px; position: relative; }}
  /* Fill segment */
  .bar > span {{ display: block; height: 100%; background: #4c8bf5; }}
  .bar.secondary > span {{ background: #f58b4c; }}
  .bar.other > span {{ background: #b5b5b5; }}
  .legend {{ display: flex; gap: 6mm; flex-wrap: wrap; margin-top: 2mm; font-size: 9pt; color: #555; }}
  .legend .item::before {{ content: ''; display: inline-block; width: 10px; height: 10px; border-radius: 2px; margin-right: 4px; vertical-align: middle; }}
  .legend .male::before {{ background: #4c8bf5; }}
  .legend .female::before {{ background: #f58b4c; }}
  .legend .other::before {{ background: #b5b5b5; }}

  /* テーブル */
  table.simple {{ border-collapse: collapse; width: 100%; font-size: 10.5pt; }}
  table.simple th, table.simple td {{ border: 1px solid #e0e0e0; padding: 6px 8px; text-align: right; }}
  table.simple th {{ background: #f9fafb; color: #444; text-align: center; }}
  table.simple tfoot td {{ font-weight: 700; background: #fafafa; }}
  table.simple td.label {{ text-align: left; }}

  /* 概要レイアウト */
  .overview-list {{ display: grid; grid-template-columns: 38mm 1fr; column-gap: 6mm; row-gap: 2mm; font-size: 11pt; }}
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
  .note-box {{ display: inline-block; padding: 3mm 5mm; border: 1px solid #e0e6ef; background: #f7f9fc; border-radius: 8px; }}
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

sections = []
for idx, q in enumerate(question_columns):
    # 補足説明列が存在すれば、最初の非空データを拾って表示（角丸矩形の中に配置）
    supp_col = f"補足説明{q}"
    inner_sup = ""
    if supp_col in df.columns:
        first_val = first_non_empty_value(df[supp_col])
        if first_val is not None:
            inner_sup = f"<h3 class=\"supplement\">{escape_html(first_val)}</h3>"

    note_box_html = f"<div class=\"note-box\">{inner_sup}<h3>選択肢</h3></div>"

    section_html = f"""
    <section class=\"page-break\">
      <h2>Q{idx+1} {q}</h2>
      {note_box_html}
      <p>この設問の集計は準備中です。（ダミー）</p>
    </section>
    """.strip()
    sections.append(section_html)

question_sections_html = "\n".join(sections)

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
        <div class="line1">東京私立中学高等学校協会 主催</div>
        <div class="line2">2024東京都私立学校展(進学相談会)</div>
      </div>
    </header>

    <section>
      <h2>概要</h2>
      <div class=\"overview-list\">
        <div class=\"label\">参加校</div>
        <div class=\"value\">東京私立中学校・高等学校 415校</div>
        <div class=\"label\">会場</div>
        <div class=\"value\">東京国際フォーラム ホールE</div>
        <div class=\"label\">開催日</div>
        <div class=\"value\">8月17日（土）18日（日）</div>
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