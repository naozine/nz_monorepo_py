#!/usr/bin/env python3
"""
Excel generator for testing the "複数回答の縦持ち化（エクスプロード）" feature in apps/pdai/app.py.

This script creates an .xlsx file containing synthetic survey-like data with multiple
columns that intentionally include multiple answers within one cell using various delimiters:
- Newline (\n)
- Comma ("," and Japanese "、，")
- Semicolon (";；")
- Tab ("\t")
- Middle dot ("・･")
- Slash ("/／")

It also injects edge cases such as:
- Leading/trailing spaces
- Consecutive delimiters (empty elements)
- Duplicated options within a cell
- Mixed full-width/half-width characters
- Case differences (UPPER/lower)
- NaN/empty values
- Non-string dtypes (numbers/dates) alongside text columns

Usage:
  python apps/pdai/generate_explode_test_excel.py --rows 50 --seed 42 --out apps/pdai/sample_explode.xlsx

The primary sheet is named "survey". You can load it in the Streamlit app and
use the sidebar "複数回答の縦持ち化（エクスプロード）" to validate behavior.
"""
import argparse
import random
from datetime import datetime, timedelta
from pathlib import Path
from typing import List

import numpy as np
import pandas as pd


CATEGORIES_A = [
    "メール", "チャット", "電話", "対面", "Web会議", "SNS",
]
CATEGORIES_B = [
    "Python", "Java", "C++", "JavaScript", "Go", "Rust",
]
# introduce variants for normalization checks
VARIANT_MAP = {
    "メール": ["ﾒｰﾙ", " メール ", "ﾒ ｰ ル"],
    "チャット": ["ﾁｬｯﾄ", " チャット"],
    "電話": ["TEL", "  電話"],
    "対面": ["対  面", " 対面 "],
    "Web会議": ["Web 会議", "WEB会議", "ｗｅｂ会議"],
    "SNS": ["ＳＮＳ", "sns", " Sns "],
}

DELIM_VARIANTS = {
    "newline": "\n",
    "comma": ",",
    "jp_comma": "、",
    "jp_comma_fw": "，",
    "semicolon": ";",
    "semicolon_jp": "；",
    "tab": "\t",
    "middledot": "・",
    "middledot_half": "･",
    "slash": "/",
    "slash_fw": "／",
}


def _random_multi_answers(options: List[str], k_min=1, k_max=4, allow_dupe=True) -> List[str]:
    k = random.randint(k_min, k_max)
    pick = random.choices(options, k=k) if allow_dupe else random.sample(options, k=min(k, len(options)))
    # inject duplicates sometimes
    if allow_dupe and random.random() < 0.3 and pick:
        pick.append(random.choice(pick))
    # map to variants sometimes
    out = []
    for item in pick:
        if item in VARIANT_MAP and random.random() < 0.4:
            out.append(random.choice(VARIANT_MAP[item]))
        else:
            out.append(item)
    return out


def _join_with_noise(parts: List[str], delimiter: str) -> str:
    # add spaces and consecutive delimiters intentionally
    s = delimiter.join(parts)
    # random spaces
    if random.random() < 0.6:
        s = s.replace(delimiter, f" {delimiter} ")
    # consecutive delimiters
    if random.random() < 0.4:
        s = s.replace(delimiter, delimiter * random.randint(2, 3), 1)
    # leading / trailing whitespace
    if random.random() < 0.5:
        s = " " * random.randint(0, 2) + s + " " * random.randint(0, 2)
    return s


def make_dataframe(n_rows: int, seed: int) -> pd.DataFrame:
    random.seed(seed)
    np.random.seed(seed)

    rows = []
    base_date = datetime(2025, 1, 1)

    for i in range(n_rows):
        # respondent id (explicit column)
        respondent_id = f"R{i+1:05d}"

        # simple attributes for filtering/grouping
        gender = random.choice(["男", "女", "その他", None])
        age_band = random.choice(["20代", "30代", "40代", "50代", "60代", None])
        dept = random.choice(["営業", "開発", "人事", "総務", "マーケ", None])
        score = np.random.normal(loc=70, scale=12)
        score = max(0, min(100, round(float(score), 1)))
        date = base_date + timedelta(days=random.randint(0, 120))

        # multi-answer fields with different delimiters
        ans_a_list = _random_multi_answers(CATEGORIES_A)
        ans_b_list = _random_multi_answers(CATEGORIES_B)

        # Choose random delimiter variant for each row/column
        delim_keys = list(DELIM_VARIANTS.keys())
        d1 = DELIM_VARIANTS[random.choice(delim_keys)]
        d2 = DELIM_VARIANTS[random.choice(delim_keys)]

        col_a = _join_with_noise(ans_a_list, d1)
        col_b = _join_with_noise(ans_b_list, d2)

        # occasionally inject NaN/empty
        if random.random() < 0.1:
            col_a = None
        if random.random() < 0.1:
            col_b = ""

        # a third column mixing two possible separators in same cell
        d3a = DELIM_VARIANTS[random.choice(delim_keys)]
        d3b = DELIM_VARIANTS[random.choice(delim_keys)]
        mix_list = _random_multi_answers(CATEGORIES_A + CATEGORIES_B)
        mixed = _join_with_noise(mix_list[: max(1, len(mix_list)//2)], d3a)
        if len(mix_list) > 1:
            mixed += d3b + _join_with_noise(mix_list[max(1, len(mix_list)//2):], d3b)
        if random.random() < 0.15:
            mixed = None

        rows.append({
            "respondent_id": respondent_id,
            "性別": gender,
            "年代": age_band,
            "部署": dept,
            "スコア": score,
            "回答日": date.date(),
            # multi-answer columns
            "Q1_利用チャネル（複数回答）": col_a,
            "Q2_得意言語（複数回答）": col_b,
            "Q3_ミックス区切り（複数回答）": mixed,
        })

    df = pd.DataFrame(rows)
    return df


def write_excel(df: pd.DataFrame, out_path: Path):
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        # Main sheet used by the app
        df.to_excel(writer, index=False, sheet_name="survey")
        # Additional sheets to test various pure delimiters (optional small subsets)
        # They can help quick testing by picking a specific sheet in the app.
        for key, delim in DELIM_VARIANTS.items():
            small = df.head(20).copy()
            def rebuild(colname: str):
                vals = []
                for x in small[colname].fillna(""):
                    if not isinstance(x, str) or x == "":
                        vals.append(x)
                        continue
                    # split by wide regex over common set, then re-join by chosen delimiter
                    parts = [p.strip() for p in pd.Series([x]).str.split(r"[,、，;；\t/／・･\r?\n]+").iloc[0] if p.strip()]
                    vals.append(delim.join(parts))
                small[colname] = vals
            for cname in ["Q1_利用チャネル（複数回答）", "Q2_得意言語（複数回答）", "Q3_ミックス区切り（複数回答）"]:
                rebuild(cname)
            small.to_excel(writer, index=False, sheet_name=f"survey_{key}")


def main():
    parser = argparse.ArgumentParser(description="Generate Excel test data for multi-answer explode feature")
    parser.add_argument("--rows", type=int, default=50, help="Number of rows to generate (default: 50)")
    parser.add_argument("--seed", type=int, default=42, help="Random seed (default: 42)")
    parser.add_argument("--out", type=str, default="apps/pdai/sample_explode.xlsx", help="Output .xlsx path")
    args = parser.parse_args()

    df = make_dataframe(args.rows, args.seed)
    out_path = Path(args.out)
    write_excel(df, out_path)

    print(f"Wrote test Excel: {out_path}  (rows={len(df)})")
    print("Sheets:")
    print(" - survey")
    for key in DELIM_VARIANTS.keys():
        print(f" - survey_{key}")


if __name__ == "__main__":
    main()
