# Python
import re
import pandas as pd
import numpy as np
from datetime import datetime

# 1) 読み込み
df = pd.read_excel("survey.xlsx", engine="openpyxl")

# 2) 「回答」列を直前の設問列に対応付けてリネーム
# - 原則: 「回答」という列名（重複含む）は、その左の列名 + "_回答" にする
# - 左の列（設問列）の各行の値は補足情報で全行同じ → 列名としての設問タイトルを使う
def map_answer_columns(frame: pd.DataFrame) -> pd.DataFrame:
    cols = list(frame.columns)
    new_cols = {}
    for i, c in enumerate(cols):
        if c == "回答" or re.match(r"^回答\.\d+$", str(c)):
            if i == 0:
                # 念のため: 先頭が回答ならそのまま
                new_cols[c] = c
            else:
                q_col_name = cols[i - 1]
                new_cols[c] = f"{q_col_name}_回答"
        else:
            new_cols[c] = c
    return frame.rename(columns=new_cols)

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

df["region_bucket"] = [region_bucket(p, c) for p, c in zip(df.get("都道府県"), df.get("市区町村"))]

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

df["channel_list"] = split_multiselect(df.get(col_channel))
df["learning_list"] = split_multiselect(df.get(col_learning))

channel_long = (
    df[["性別", "region_bucket", "grade_2024", "channel_list"]]
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

# 9) 可視化（必要に応じて）
# 例: channel_by_region.plot.barh(stacked=True), 等