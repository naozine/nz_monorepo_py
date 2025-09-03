# streamlit アンケート集計アプリ
# 依存: streamlit, pandas, openpyxl, matplotlib
# 任意: openai（サイドバーでAPIキー設定時のみ使用）
import io
import json
import re
from dataclasses import dataclass, asdict
from typing import Any, Dict, List, Literal, Optional, Tuple

import matplotlib
import pandas as pd
import pandas.api.types as ptypes
import streamlit as st

matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib import font_manager as _fm

# 日本語フォント設定（環境にあるものを自動選択）
def _setup_japanese_font():
    try:
        candidates = [
            "IPAexGothic", "IPAGothic", "Noto Sans CJK JP", "Noto Sans JP",
            "Yu Gothic", "YuGothic", "Hiragino Sans", "Hiragino Kaku Gothic ProN",
            "Meiryo", "TakaoGothic", "MotoyaGothic", "MS Gothic", "MS PGothic"
        ]
        available = {f.name for f in _fm.fontManager.ttflist}
        for name in candidates:
            if name in available:
                matplotlib.rcParams["font.family"] = name
                break
        # マイナス記号が豆腐になるのを防ぐ
        matplotlib.rcParams["axes.unicode_minus"] = False
    except (OSError, FileNotFoundError, RuntimeError, ValueError):
        # フォント探索に失敗してもアプリは継続
        matplotlib.rcParams["axes.unicode_minus"] = False

_setup_japanese_font()

APP_TITLE = "アンケート集計アプリ"
HISTORY_LIMIT = 5
DATA_PREVIEW_ROWS = 30

# ========== 型定義 ==========
Comparator = Literal["=", "≠", ">", "≥", "<", "≤", "contains", "in-list"]
LogicOp = Literal["AND", "OR"]
AggFunc = Literal["count", "sum", "mean", "median", "min", "max"]
ChartType = Literal["棒", "横棒", "折れ線", "円"]

@dataclass
class FilterCond:
    column: str
    op: Comparator
    value: Any  # スカラー or リスト（in-list）
    dtype: str  # "number" | "string" | "date"

@dataclass
class VizConfig:
    chart_type: ChartType
    x_label: str = ""
    y_label: str = ""
    legend: bool = True
    percent: bool = False
    sort: str = "自動"  # "自動" | "値昇順" | "値降順" | "ラベル昇順" | "ラベル降順"
    top_n: Optional[int] = None  # 可視性担保

@dataclass
class RunConfig:
    mode: str  # "単純集計" | "グループ集計" | "クロス集計" | "上位N" | "プロンプト集計"
    # 共通
    filters: List[FilterCond]
    logic: LogicOp
    exclude_columns: List[str]
    # 単純集計
    simple_col: Optional[str] = None
    simple_normalize: bool = False
    # グループ集計
    groupby_cols: Optional[List[str]] = None
    agg_map: Optional[Dict[str, List[AggFunc]]] = None
    # クロス集計
    pivot_index: Optional[List[str]] = None
    pivot_columns: Optional[List[str]] = None
    pivot_values: Optional[List[str]] = None
    pivot_aggfunc: Optional[AggFunc] = None
    pivot_margins: bool = False
    # 上位N
    topn_col: Optional[str] = None
    topn_n: Optional[int] = None
    # 可視化
    viz: Optional[VizConfig] = None
    # プロンプト解釈
    prompt_raw: Optional[str] = None
    prompt_interpreted: Optional[Dict[str, Any]] = None

# ========== ユーティリティ ==========

def _to_datetime_if_possible(s: pd.Series) -> pd.Series:
    if ptypes.is_datetime64_any_dtype(s):
        return s
    if ptypes.is_object_dtype(s):
        try:
            out = pd.to_datetime(s, errors="coerce")
            # 変換成功率で判断
            if out.notna().mean() > 0.7:
                return out
        except (TypeError, ValueError):
            pass
    return s

def dtype_optimize(df: pd.DataFrame) -> Tuple[pd.DataFrame, Dict[str, str]]:
    """簡易メモリ最適化: 低ユニーク率のobject→category, 日付っぽい→datetime"""
    df2 = df.copy()
    suggestions: Dict[str, str] = {}
    for col in df2.columns:
        s = df2[col]
        if ptypes.is_object_dtype(s):
            nunique = s.nunique(dropna=True)
            ratio = nunique / max(len(s), 1)
            if ratio < 0.5:  # 適度な閾値
                df2[col] = s.astype("category")
                suggestions[col] = "object→category"
            else:
                # 日付推定
                s2 = _to_datetime_if_possible(s)
                if ptypes.is_datetime64_any_dtype(s2) and not ptypes.is_datetime64_any_dtype(s):
                    df2[col] = s2
                    suggestions[col] = "object→datetime(推定)"
        elif ptypes.is_numeric_dtype(s):
            # 数値はそのまま
            pass
        elif not ptypes.is_datetime64_any_dtype(s):
            # 日付推定
            s2 = _to_datetime_if_possible(s)
            if ptypes.is_datetime64_any_dtype(s2):
                df2[col] = s2
                suggestions[col] = "→datetime(推定)"
    return df2, suggestions

def read_excel_file(uploaded_file, sheet_name=None) -> Dict[str, pd.DataFrame]:
    xls = pd.ExcelFile(uploaded_file, engine="openpyxl")
    sheets = xls.sheet_names
    targets = [sheet_name] if sheet_name else sheets
    data = {}
    for sh in targets:
        try:
            df = pd.read_excel(uploaded_file, sheet_name=sh, engine="openpyxl")
        except (ImportError, ModuleNotFoundError, ValueError):
            df = pd.read_excel(uploaded_file, sheet_name=sh)
        data[sh] = df
    return data

def summarize_df(df: pd.DataFrame) -> Dict[str, Any]:
    missing = df.isna().sum()
    return {
        "rows": len(df),
        "cols": len(df.columns),
        "columns": list(df.columns),
        "missing": missing.to_dict(),
    }

def apply_filters(df: pd.DataFrame, filters: List[FilterCond], logic: LogicOp) -> pd.DataFrame:
    if not filters:
        return df
    masks = []
    for f in filters:
        if f.column not in df.columns:
            masks.append(pd.Series(False, index=df.index))
            continue
        s = df[f.column]
        val = f.value
        # 型合わせ
        if f.dtype == "number":
            try:
                sval = pd.to_numeric(s, errors="coerce")
                if isinstance(val, list):
                    vlist = [pd.to_numeric([v], errors="coerce")[0] for v in val]
                else:
                    vlist = None
                    val = pd.to_numeric([val], errors="coerce")[0]
                s = sval
            except (TypeError, ValueError):
                pass
        elif f.dtype == "date":
            s = _to_datetime_if_possible(s)
            if isinstance(val, list):
                vlist = [pd.to_datetime(v, errors="coerce") for v in val]
            else:
                vlist = None
                val = pd.to_datetime(val, errors="coerce")
        else:
            # string
            vlist = [str(v) for v in val] if isinstance(val, list) else None
            val = str(val) if not isinstance(val, list) else val

        # オペレータ
        if f.op == "=":
            mask = s.eq(val)
        elif f.op == "≠":
            mask = ~s.eq(val)
        elif f.op == ">":
            mask = s > val
        elif f.op == "≥":
            mask = s >= val
        elif f.op == "<":
            mask = s < val
        elif f.op == "≤":
            mask = s <= val
        elif f.op == "contains":
            mask = s.astype(str).str.contains(str(val), case=False, na=False)
        elif f.op == "in-list":
            if vlist is None:
                vlist = [val]
            mask = s.isin(vlist)
        else:
            mask = pd.Series(True, index=df.index)
        masks.append(mask.fillna(False))
    if not masks:
        return df
    out_mask = masks[0]
    for m in masks[1:]:
        if logic == "AND":
            out_mask = out_mask & m
    if logic == "OR":
        out_mask = pd.Series(False, index=df.index)
        for m in masks:
            out_mask = out_mask | m
    return df[out_mask]

def simple_value_counts(df: pd.DataFrame, col: str, normalize: bool) -> pd.DataFrame:
    vc = df[col].value_counts(dropna=False, normalize=normalize)
    out = vc.rename("割合" if normalize else "件数").reset_index().rename(columns={"index": col})
    return out

def group_aggregate(df: pd.DataFrame, group_cols: List[str], agg_map: Dict[str, List[AggFunc]]) -> pd.DataFrame:
    if not group_cols:
        raise ValueError("グループ化する列を1つ以上選択してください。")
    # agg map pandas形式へ変換
    agg_dict: Dict[str, List[str]] = {}
    for k, v in agg_map.items():
        agg_dict[k] = v
    g = df.groupby(group_cols, dropna=False)
    res = g.agg(agg_dict)
    # 列名整形
    res.columns = ["_".join([c for c in tup if c]).strip("_") if isinstance(tup, tuple) else str(tup) for tup in res.columns.values]
    return res.reset_index()

def pivot_aggregate(df: pd.DataFrame, index: List[str], columns: List[str], values: List[str], aggfunc: AggFunc, margins: bool) -> pd.DataFrame:
    if not index or not columns or not values:
        raise ValueError("行・列・値はすべて指定してください。")
    if len(values) == 1:
        pv = pd.pivot_table(df, index=index, columns=columns, values=values[0], aggfunc=aggfunc, margins=margins, margins_name="合計", dropna=False)
    else:
        pv = pd.pivot_table(df, index=index, columns=columns, values=values, aggfunc=aggfunc, margins=margins, margins_name="合計", dropna=False)
    return pv.reset_index()

def top_n_categories(df: pd.DataFrame, col: str, n: int) -> pd.DataFrame:
    vc = df[col].value_counts(dropna=False).head(n)
    return vc.rename("件数").reset_index().rename(columns={"index": col})

def sort_dataframe_for_viz(df: pd.DataFrame, value_col: str, label_col: Optional[str], order: str) -> pd.DataFrame:
    if order == "自動":
        return df
    if order == "値昇順":
        return df.sort_values(by=value_col, ascending=True)
    if order == "値降順":
        return df.sort_values(by=value_col, ascending=False)
    if label_col is not None and order == "ラベル昇順":
        return df.sort_values(by=label_col, ascending=True, key=lambda s: s.astype(str))
    if label_col is not None and order == "ラベル降順":
        return df.sort_values(by=label_col, ascending=False, key=lambda s: s.astype(str))
    return df

def plot_with_matplotlib(df: pd.DataFrame, chart_type: ChartType, x: Optional[str], y: Optional[str], series_col: Optional[str], percent: bool, legend: bool, x_label: str, y_label: str) -> bytes:
    plt.close("all")
    fig, ax = plt.subplots(figsize=(8, 4.5), dpi=150)

    if chart_type in ("棒", "横棒"):
        # 単系列 or 多系列（series_colがある場合はピボット）
        if series_col and x and y:
            pivot = df.pivot(index=x, columns=series_col, values=y)
            if chart_type == "棒":
                pivot.plot(kind="bar", ax=ax)
            else:
                pivot.plot(kind="barh", ax=ax)
        else:
            if x and y:
                if chart_type == "棒":
                    ax.bar(df[x].astype(str), df[y])
                else:
                    ax.barh(df[x].astype(str), df[y])
            elif y is None and x is None and df.shape[1] >= 2:
                # fallback: 1列目x,2列目y
                if chart_type == "棒":
                    ax.bar(df.iloc[:, 0].astype(str), df.iloc[:, 1])
                else:
                    ax.barh(df.iloc[:, 0].astype(str), df.iloc[:, 1])
    elif chart_type == "折れ線":
        if series_col and x and y:
            pivot = df.pivot(index=x, columns=series_col, values=y)
            pivot.plot(kind="line", marker="o", ax=ax)
        else:
            if x and y:
                ax.plot(df[x].astype(str), df[y], marker="o")
            else:
                ax.plot(df.index, df.iloc[:, 0], marker="o")
    elif chart_type == "円":
        # 円は単系列のみ
        if x and y:
            ax.pie(df[y], labels=df[x].astype(str), autopct="%1.1f%%" if percent else None)
        else:
            ax.pie(df.iloc[:, 1], labels=df.iloc[:, 0].astype(str), autopct="%1.1f%%" if percent else None)

    ax.set_xlabel(x_label or (x or ""))
    ax.set_ylabel((y_label or (("割合(%)" if percent else "値"))) if chart_type != "円" else "")
    if legend and chart_type != "円":
        ax.legend(loc="best")
    fig.tight_layout()

    buf = io.BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight")
    buf.seek(0)
    return buf.read()

def df_to_csv_bytes(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False).encode("utf-8-sig")

def df_to_excel_bytes(df: pd.DataFrame, sheet_name="result") -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    buf.seek(0)
    return buf.read()

def normalize_text_variants(df: pd.DataFrame, mapping: Dict[str, str], target_col: str) -> pd.DataFrame:
    if not mapping or target_col not in df.columns:
        return df
    s = df[target_col].astype(str)
    df[target_col] = s.replace(mapping)
    return df

# ========== ルールベース・プロンプトパーサ（簡易） ==========

def parse_prompt_jp(prompt: str, columns: List[str]) -> Tuple[Optional[Dict[str, Any]], List[Dict[str, Any]]]:
    """
    日本語の自然文から集計指示を抽出（簡易）。
    返り値: (best, candidates)
    best/candidate 例:
    {
      "mode": "pivot"|"group"|"simple"|"topn",
      "index": [...], "columns": [...], "values": [...], "agg":"mean", "percent": True,
      "sort": "降順"|"昇順", "sort_by": "値"|"ラベル", "topn": 5,
      "filters": [{"col":"部署","op":"=","val":"営業"}]
    }
    """
    txt = prompt.strip()
    if not txt:
        return None, []
    cand: List[Dict[str, Any]] = []
    best: Optional[Dict[str, Any]] = None

    # 正規化
    t = txt.replace("％", "%").replace("×", " x ").replace("＊", "*")
    percent = bool(re.search(r"(割合|%|パーセント)", t))
    # 上位N
    m_top = re.search(r"(上位|トップ)\s*(\d+)", t)
    topn = int(m_top.group(2)) if m_top else None

    # フィルタ（簡易）: 「X=Y」「X は Y」「X in [A,B]」「Xのみ」
    filters = []
    eqs = re.findall(r"([\w一-龠ぁ-んァ-ンー]+)\s*(?:=|が|は)\s*([^\s、，]+)", t)
    for k, v in eqs:
        if k in columns:
            filters.append({"col": k, "op": "=", "val": v})
    in_list = re.findall(r"([\w一-龠ぁ-んァ-ンー]+)\s*(?:in|IN|In)\s*\[([^\]]+)\]", t)
    for k, body in in_list:
        if k in columns:
            vals = [s.strip() for s in re.split(r"[,\s、，]+", body) if s.strip()]
            filters.append({"col": k, "op": "in-list", "val": vals})
    only = re.findall(r"([\w一-龠ぁ-んァ-ンー]+)\s*のみ", t)
    for k in only:
        if k in columns:
            # 「部署のみ」は値が欠けるので候補止まり
            filters.append({"col": k, "op": "≠", "val": ""})

    # クロス/ピボット
    if ("クロス" in t or "ピボット" in t or " x " in t):
        # 形: A x B / AとB / A×B
        m_pair = re.search(r"([\w一-龠ぁ-んァ-ンー]+)\s*(?:x|×|と)\s*([\w一-龠ぁ-んァ-ンー]+)", t)
        if m_pair:
            a, b = m_pair.group(1), m_pair.group(2)
            if a in columns and b in columns:
                best = {
                    "mode": "pivot",
                    "index": [a],
                    "columns": [b],
                    "values": [],  # 値未指定→件数
                    "agg": "count",
                    "percent": percent,
                    "sort": "降順" if ("降順" in t) else ("昇順" if "昇順" in t else None),
                    "sort_by": "値" if ("値" in t) else ("ラベル" if "ラベル" in t else None),
                    "topn": topn,
                    "filters": filters,
                }
                cand.append(best)

    # 平均・合計など + 「年代別」「部署別」など
    m_stat = None
    for key, agg in [("平均", "mean"), ("合計", "sum"), ("中央値", "median"), ("最小", "min"), ("最大", "max"), ("件数", "count")]:
        if key in t:
            m_stat = agg
            break
    m_by = re.findall(r"([\w一-龠ぁ-んァ-ンー]+)別", t)
    nums = [c for c in columns if is_numeric_dtype_candidate_name(c)]
    if m_stat and m_by:
        # 最初の数値列を値に仮置き
        val_col = nums[0] if nums else None
        group_cols = [c for c in m_by if c in columns]
        if group_cols:
            cand.append({
                "mode": "group",
                "groupby": group_cols,
                "values": [val_col] if val_col else [],
                "agg": m_stat,
                "percent": percent,
                "topn": topn,
                "filters": filters,
            })
            if best is None:
                best = cand[-1]

    # 単純集計（value_counts）
    for c in columns:
        if c in t and ("割合" in t or "件数" in t or "頻度" in t):
            cand.append({
                "mode": "simple",
                "column": c,
                "percent": percent,
                "topn": topn,
                "filters": filters,
            })
            if best is None:
                best = cand[-1]
            break

    # 上位Nカテゴリ（列名推定）
    if topn and not best:
        for c in columns:
            cand.append({"mode": "topn", "column": c, "topn": topn, "filters": filters})
        if cand:
            best = cand[0]

    return best, cand

def is_numeric_dtype_candidate_name(name: str) -> bool:
    # 名称で数値列らしさを推定（スコア/点/金額/数/回/量/時間）
    return any(key in name for key in ["スコア", "点", "金額", "数", "回", "量", "時間", "score", "amt", "num"])

def interpreted_to_runconfig(interp: Dict[str, Any]) -> RunConfig:
    viz = VizConfig(
        chart_type="棒",
        percent=bool(interp.get("percent")),
        sort=("値降順" if interp.get("sort") == "降順" else ("値昇順" if interp.get("sort") == "昇順" else "自動")),
    )
    filters: List[FilterCond] = []
    for f in interp.get("filters", []):
        filters.append(FilterCond(column=f["col"], op=f.get("op", "="), value=f.get("val"), dtype="string"))
    logic: LogicOp = "AND"
    mode = interp.get("mode")
    rc = RunConfig(mode="プロンプト集計", filters=filters, logic=logic, exclude_columns=[], viz=viz, prompt_raw=None, prompt_interpreted=interp)
    if mode == "pivot":
        rc.mode = "クロス集計"
        rc.pivot_index = interp.get("index", [])
        rc.pivot_columns = interp.get("columns", [])
        rc.pivot_values = interp.get("values", [])
        rc.pivot_aggfunc = interp.get("agg", "count")  # type: ignore
    elif mode == "group":
        rc.mode = "グループ集計"
        rc.groupby_cols = interp.get("groupby", [])
        agg = interp.get("agg", "mean")
        values = interp.get("values", [])
        rc.agg_map = {v: [agg] for v in values if v}
    elif mode == "simple":
        rc.mode = "単純集計"
        rc.simple_col = interp.get("column")
        rc.simple_normalize = bool(interp.get("percent"))
    elif mode == "topn":
        rc.mode = "上位N"
        rc.topn_col = interp.get("column")
        rc.topn_n = interp.get("topn")
    return rc

# ========== Streamlit UI ==========

def init_session():
    if "history" not in st.session_state:
        st.session_state["history"]: List[Dict[str, Any]] = []
    if "export_blob" not in st.session_state:
        st.session_state["export_blob"] = None
    if "normalized_mapping" not in st.session_state:
        st.session_state["normalized_mapping"] = {}
    if "llm_enabled" not in st.session_state:
        st.session_state["llm_enabled"] = False
    if "llm_provider" not in st.session_state:
        st.session_state["llm_provider"] = "openai"
    if "llm_api_key" not in st.session_state:
        st.session_state["llm_api_key"] = ""
    if "last_result" not in st.session_state:
        st.session_state["last_result"] = None  # {"df": DataFrame JSON, "viz_png": bytes}
    if "selected_sheet" not in st.session_state:
        st.session_state["selected_sheet"] = None

def sidebar_file_and_options():
    st.sidebar.header("ファイル入力")
    file = st.sidebar.file_uploader("Excelファイル（.xlsx）をアップロード", type=["xlsx"])
    sheet_name = None
    df_map = {}
    if file:
        try:
            xls = pd.ExcelFile(file, engine="openpyxl")
            sheets = xls.sheet_names
            sheet_name = st.sidebar.selectbox("シートを選択", sheets, index=0)
            st.session_state["selected_sheet"] = sheet_name
            df_map = read_excel_file(file, sheet_name=sheet_name)
        except Exception as e:
            st.sidebar.error(f"読み込みエラー: {e}")
    st.sidebar.caption("アップロードしたデータはこの端末内で処理されます。外部送信は行いません。")
    return df_map, sheet_name

def sidebar_main_controls(df: pd.DataFrame):
    st.sidebar.header("操作パネル")
    mode = st.sidebar.selectbox("集計モード", ["単純集計", "グループ集計", "クロス集計（ピボット）", "上位N", "プロンプト集計"])
    st.sidebar.markdown("---")

    # 不要列の除外提案
    with st.sidebar.expander("不要列の除外・型最適化", expanded=False):
        suggest_df, suggestions = dtype_optimize(df)
        to_exclude = st.multiselect("除外する列（計算から外す）", options=list(df.columns))
        st.caption("型最適化の提案: " + (", ".join([f"{k}:{v}" for k, v in suggestions.items()]) if suggestions else "なし"))
        apply_opt = st.checkbox("型最適化を適用（カテゴリ化・日付推定）", value=bool(suggestions))
    if apply_opt:
        df = suggest_df

    # 条件絞り込み
    filters_ui: List[FilterCond] = []
    with st.sidebar.expander("条件絞り込み", expanded=False):
        logic: LogicOp = st.radio("条件の結合", ["AND", "OR"], horizontal=True)
        add_rows = st.number_input("条件の数", min_value=0, max_value=20, value=0, step=1)
        for i in range(add_rows):
            st.write(f"条件 {i+1}")
            col = st.selectbox(f"列{i+1}", options=list(df.columns), key=f"f_col_{i}")
            # dtype推定
            if ptypes.is_numeric_dtype(df[col]):
                dtype = "number"
                ops = ["=", "≠", ">", "≥", "<", "≤", "in-list"]
            elif ptypes.is_datetime64_any_dtype(df[col]):
                dtype = "date"
                ops = ["=", "≠", ">", "≥", "<", "≤", "in-list"]
            else:
                dtype = "string"
                ops = ["=", "≠", "contains", "in-list"]
            op = st.selectbox(f"演算子{i+1}", options=ops, key=f"f_op_{i}")
            if op == "in-list":
                val = st.text_input(f"値（カンマ区切り）{i+1}", key=f"f_val_{i}")
                vals = [v.strip() for v in re.split(r"[,\s、，]+", val) if v.strip()]
                value = vals
            else:
                value = st.text_input(f"値{i+1}", key=f"f_val_{i}")
            filters_ui.append(FilterCond(column=col, op=op, value=value, dtype=dtype))
    return mode, df, filters_ui, logic, to_exclude

def run_simple(df: pd.DataFrame, filters: List[FilterCond], logic: LogicOp, exclude: List[str]) -> Tuple[pd.DataFrame, Dict[str, Any]]:
    cols = [c for c in df.columns if c not in exclude]
    if not cols:
        raise ValueError("列がありません。除外を見直してください。")
    target = st.selectbox("対象列（単純集計）", options=cols)
    normalize = st.checkbox("割合（%）で表示", value=True)
    fdf = apply_filters(df, filters, logic)
    if fdf.empty:
        raise ValueError("フィルタ後のデータが空です。")
    res = simple_value_counts(fdf, target, normalize=normalize)
    return res, {"simple_col": target, "normalize": normalize}

def run_group(df: pd.DataFrame, filters: List[FilterCond], logic: LogicOp, exclude: List[str]) -> Tuple[pd.DataFrame, Dict[str, Any]]:
    cols = [c for c in df.columns if c not in exclude]
    group_cols = st.multiselect("グループ化する列", options=cols)
    num_candidates = [c for c in cols if ptypes.is_numeric_dtype(df[c])]
    agg_target_cols = st.multiselect("集計する数値列", options=num_candidates)
    agg_funcs: List[AggFunc] = st.multiselect("集計関数", options=["count", "sum", "mean", "median", "min", "max"], default=["count", "mean"])
    agg_map = {c: agg_funcs for c in agg_target_cols}
    fdf = apply_filters(df, filters, logic)
    if fdf.empty:
        raise ValueError("フィルタ後のデータが空です。")
    if not group_cols:
        raise ValueError("グループ化する列を選択してください。")
    if not agg_map:
        raise ValueError("集計する数値列を1つ以上選択してください。")
    res = group_aggregate(fdf, group_cols, agg_map)
    return res, {"groupby_cols": group_cols, "agg_map": agg_map}

def run_pivot(df: pd.DataFrame, filters: List[FilterCond], logic: LogicOp, exclude: List[str]) -> Tuple[pd.DataFrame, Dict[str, Any]]:
    cols = [c for c in df.columns if c not in exclude]
    index = st.multiselect("行（index）", options=cols)
    columns = st.multiselect("列（columns）", options=cols)
    values = st.multiselect("値（values; 空の場合は件数カウント）", options=cols)
    aggfunc: AggFunc = st.selectbox("集計関数", options=["count", "sum", "mean", "median", "min", "max"], index=0)
    margins = st.checkbox("合計行列（margins）を表示", value=True)
    fdf = apply_filters(df, filters, logic)
    if fdf.empty:
        raise ValueError("フィルタ後のデータが空です。")
    if not values:
        # 値未指定 → 件数用にダミー列を作って集計
        fdf = fdf.copy()
        fdf["_count_"] = 1
        res = pivot_aggregate(fdf, index, columns, ["_count_"], "sum", margins)
        # 列名調整
        res = res.rename(columns={c: ("件数" if c == "_count_" else c) for c in res.columns})
    else:
        res = pivot_aggregate(fdf, index, columns, values, aggfunc, margins)
    return res, {"pivot_index": index, "pivot_columns": columns, "pivot_values": values, "pivot_aggfunc": aggfunc, "pivot_margins": margins}

def run_topn(df: pd.DataFrame, filters: List[FilterCond], logic: LogicOp, exclude: List[str]) -> Tuple[pd.DataFrame, Dict[str, Any]]:
    cols = [c for c in df.columns if c not in exclude]
    target = st.selectbox("対象列（Top-N）", options=cols)
    n = st.number_input("N", min_value=1, max_value=1000, value=10, step=1)
    fdf = apply_filters(df, filters, logic)
    if fdf.empty:
        raise ValueError("フィルタ後のデータが空です。")
    res = top_n_categories(fdf, target, int(n))
    return res, {"topn_col": target, "topn_n": int(n)}

def viz_controls(default_percent: bool) -> VizConfig:
    st.subheader("可視化設定")
    chart: ChartType = st.selectbox("グラフ種類", ["棒", "横棒", "折れ線", "円"])
    xlabel = st.text_input("X軸ラベル（任意）", "")
    ylabel = st.text_input("Y軸ラベル（任意）", "")
    legend = st.checkbox("凡例を表示", value=True)
    percent = st.checkbox("単位を%として扱う（軸・注釈）", value=default_percent)
    sort = st.selectbox("ソート順", ["自動", "値昇順", "値降順", "ラベル昇順", "ラベル降順"])
    topn_for_viz = st.number_input("グラフ表示の上位N（視認性）", min_value=0, max_value=1000, value=0, step=1)
    return VizConfig(chart_type=chart, x_label=xlabel, y_label=ylabel, legend=legend, percent=percent, sort=sort, top_n=(None if topn_for_viz == 0 else int(topn_for_viz)))

def render_chart_and_downloads(result_df: pd.DataFrame, viz: VizConfig, label_col: Optional[str] = None, value_col: Optional[str] = None, series_col: Optional[str] = None):
    # サンプリング/上位N
    dfv = result_df.copy()
    if viz.top_n and value_col and viz.top_n < len(dfv):
        dfv = dfv.nlargest(viz.top_n, value_col)

    # 可視化列推定
    if value_col is None:
        # 候補: 数値列
        num_cols = [c for c in dfv.columns if ptypes.is_numeric_dtype(dfv[c])]
        value_col = num_cols[0] if num_cols else (dfv.columns[1] if dfv.shape[1] >= 2 else dfv.columns[0])
    if label_col is None:
        label_col = dfv.columns[0] if dfv.shape[1] >= 1 else None

    # ソート
    if value_col:
        dfv = sort_dataframe_for_viz(dfv, value_col, label_col, viz.sort)

    # 表示
    st.dataframe(dfv, use_container_width=True, height=350)

    # グラフ
    png = plot_with_matplotlib(
        dfv, viz.chart_type, x=label_col, y=value_col, series_col=series_col, percent=viz.percent, legend=viz.legend, x_label=viz.x_label, y_label=viz.y_label
    )
    st.image(png, caption="グラフプレビュー", use_column_width=True)

    # ダウンロード
    c1, c2, c3 = st.columns(3)
    with c1:
        st.download_button("CSVダウンロード", data=df_to_csv_bytes(dfv), file_name="result.csv", mime="text/csv")
    with c2:
        st.download_button("Excelダウンロード", data=df_to_excel_bytes(dfv), file_name="result.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    with c3:
        st.download_button("図（PNG）ダウンロード", data=png, file_name="chart.png", mime="image/png")

    # 最新結果保存（履歴用）
    st.session_state["last_result"] = {"df": dfv.to_json(orient="split", force_ascii=False), "png": png}

def export_config(config: RunConfig):
    blob = json.dumps(asdict(config), ensure_ascii=False, indent=2)
    st.download_button("現在の設定をJSONエクスポート", data=blob.encode("utf-8"), file_name="config.json", mime="application/json")

def import_config_ui() -> Optional[RunConfig]:
    up = st.file_uploader("設定JSONをインポート", type=["json"], key="config_importer")
    if up:
        try:
            data = json.load(up)
            # 簡易バリデーション
            if "mode" in data and "filters" in data and "logic" in data:
                # FilterCond再構築
                data["filters"] = [FilterCond(**f) for f in data.get("filters", [])]
                if "viz" in data and data["viz"] is not None:
                    data["viz"] = VizConfig(**data["viz"])
                rc = RunConfig(**data)
                st.success("設定を読み込みました。左のパネルで再設定してください。")
                return rc
        except Exception as e:
            st.error(f"設定の読み込みに失敗しました: {e}")
    return None

def push_history(config: RunConfig):
    hist = st.session_state["history"]
    payload = json.dumps(asdict(config), ensure_ascii=False)
    # 重複防止
    if payload in hist:
        hist.remove(payload)
    hist.insert(0, payload)
    del hist[HISTORY_LIMIT:]  # 5件まで

def render_history_restore() -> Optional[RunConfig]:
    st.subheader("実行履歴（直近5件）")
    hist = st.session_state["history"]
    if not hist:
        st.info("履歴はまだありません。")
        return None
    for i, payload in enumerate(hist, start=1):
        with st.expander(f"履歴 {i}", expanded=False):
            st.code(payload, language="json")
            if st.button("この設定を復元", key=f"restore_{i}"):
                data = json.loads(payload)
                data["filters"] = [FilterCond(**f) for f in data.get("filters", [])]
                if "viz" in data and data["viz"] is not None:
                    data["viz"] = VizConfig(**data["viz"])
                return RunConfig(**data)
    return None

def main():
    st.set_page_config(page_title=APP_TITLE, layout="wide")
    init_session()

    st.title(APP_TITLE)
    st.caption("Excelのアンケート回答データをアップロードし、GUIや自然文プロンプトで集計・可視化します。")

    df_map, sheet = sidebar_file_and_options()
    if not df_map:
        st.info("左のサイドバーから .xlsx ファイルをアップロードしてください。")
        return
    df = list(df_map.values())[0]

    # 概要
    st.subheader("データ概要")
    sm = summarize_df(df)
    c1, c2, c3, c4 = st.columns([2, 2, 3, 3])
    with c1:
        st.metric("件数（行）", sm["rows"])
    with c2:
        st.metric("列数", sm["cols"])
    with c3:
        st.write("列名")
        st.write(", ".join(sm["columns"][:50]) + (" ..." if len(sm["columns"]) > 50 else ""))
    with c4:
        st.write("欠損状況（上位）")
        miss_sorted = sorted(sm["missing"].items(), key=lambda x: x[1], reverse=True)[:10]
        st.write({k: v for k, v in miss_sorted})

    # データプレビュー
    st.subheader("データプレビュー（先頭30行）")
    st.dataframe(df.head(DATA_PREVIEW_ROWS), use_container_width=True, height=320)

    # サイドバー：モード・フィルタ他
    mode, df_eff, filters, logic, excluded = sidebar_main_controls(df)

    # 表記ゆれ正規化支援
    with st.expander("表記ゆれ正規化（簡易マッピング）", expanded=False):
        target_col = st.selectbox("対象列", options=list(df_eff.columns))
        st.caption("例: 男性→男,  女性→女 など。1行に 置換前=置換後 を記載します。")
        mapping_text = st.text_area("置換マッピング", value="")
        mapping = {}
        if mapping_text.strip():
            for line in mapping_text.splitlines():
                if "=" in line:
                    src, dst = line.split("=", 1)
                    mapping[src.strip()] = dst.strip()
        if st.button("正規化を適用"):
            df_eff = normalize_text_variants(df_eff, mapping, target_col)
            st.success("正規化を適用しました。")

    # 実行領域
    result_df: Optional[pd.DataFrame] = None
    run_config: Optional[RunConfig] = None
    error_msg: Optional[str] = None

    try:
        if mode == "単純集計":
            res, extra = run_simple(df_eff, filters, logic, excluded)
            result_df = res
            viz = viz_controls(default_percent=extra["normalize"])
            run_config = RunConfig(mode="単純集計", filters=filters, logic=logic, exclude_columns=excluded, simple_col=extra["simple_col"], simple_normalize=extra["normalize"], viz=viz)
            st.subheader("結果")
            render_chart_and_downloads(result_df, viz, label_col=extra["simple_col"], value_col=("割合" if extra["normalize"] else "件数"))

        elif mode == "グループ集計":
            res, extra = run_group(df_eff, filters, logic, excluded)
            result_df = res
            viz = viz_controls(default_percent=False)
            run_config = RunConfig(mode="グループ集計", filters=filters, logic=logic, exclude_columns=excluded, groupby_cols=extra["groupby_cols"], agg_map=extra["agg_map"], viz=viz)
            st.subheader("結果")
            # 可視化列推定: 最初の集計列
            agg_cols = [c for c in result_df.columns if c not in (extra["groupby_cols"] or [])]
            val_col = None
            if agg_cols:
                # mean優先
                mean_cols = [c for c in agg_cols if c.endswith("_mean")]
                val_col = mean_cols[0] if mean_cols else agg_cols[0]
            label_col = (extra["groupby_cols"] or [None])[0]
            render_chart_and_downloads(result_df, viz, label_col=label_col, value_col=val_col)

        elif mode == "クロス集計（ピボット）":
            res, extra = run_pivot(df_eff, filters, logic, excluded)
            result_df = res
            viz = viz_controls(default_percent=False)
            run_config = RunConfig(mode="クロス集計", filters=filters, logic=logic, exclude_columns=excluded,
                                   pivot_index=extra["pivot_index"], pivot_columns=extra["pivot_columns"],
                                   pivot_values=extra["pivot_values"], pivot_aggfunc=extra["pivot_aggfunc"], pivot_margins=extra["pivot_margins"], viz=viz)
            st.subheader("結果")
            # 可視化: 値が多列になるので、indexをラベルに、columnsを系列とする想定に変換
            df_v = result_df.copy()
            if (extra["pivot_values"] and len(extra["pivot_values"]) == 1) or (not extra["pivot_values"]):
                # columnsが複数レベルの場合のflatten
                df_v.columns = [c if not isinstance(c, tuple) else "_".join([str(x) for x in c if x != ""]) for c in df_v.columns]
                # 可能なら index と columns の2軸で melt して可視化
                if extra["pivot_columns"]:
                    idx = extra["pivot_index"] or []
                    melt_id = idx
                    melt_val_vars = [c for c in df_v.columns if c not in idx]
                    dfm = df_v.melt(id_vars=melt_id, value_vars=melt_val_vars, var_name="系列", value_name="値")
                    # ラベルは index の最初
                    label_col = melt_id[0] if melt_id else "index"
                    result_for_viz = dfm.rename(columns={melt_id[0]: label_col}) if melt_id else dfm
                    st.dataframe(result_df, use_container_width=True, height=300)
                    st.subheader("可視化")
                    render_chart_and_downloads(result_for_viz, viz, label_col=label_col, value_col="値", series_col="系列")
                else:
                    st.dataframe(result_df, use_container_width=True, height=350)
                    st.info("列（columns）が未指定のため、テーブルのみ表示しました。")
            else:
                st.dataframe(result_df, use_container_width=True, height=350)
                st.info("複数の値列があるため、表のみを表示しています。")

        elif mode == "上位N":
            res, extra = run_topn(df_eff, filters, logic, excluded)
            result_df = res
            viz = viz_controls(default_percent=False)
            run_config = RunConfig(mode="上位N", filters=filters, logic=logic, exclude_columns=excluded, topn_col=extra["topn_col"], topn_n=extra["topn_n"], viz=viz)
            st.subheader("結果")
            render_chart_and_downloads(result_df, viz, label_col=extra["topn_col"], value_col="件数")

        elif mode == "プロンプト集計":
            st.subheader("自然文プロンプト")
            prompt = st.text_input("例: 性別×満足度のクロス集計を割合で / 年代別に平均スコア / 営業部のみで上位5カテゴリ", "")
            # 任意のLLM（APIキー入力時のみ有効）
            with st.expander("LLM連携（任意）", expanded=False):
                st.caption("APIキー入力時のみ有効。入力したデータは送信されませんが、プロンプト文は外部送信されます。")
                provider = st.selectbox("プロバイダ", ["openai"], index=0)
                api_key = st.text_input("APIキー（任意）", type="password")
                st.session_state["llm_enabled"] = bool(api_key)
                st.session_state["llm_provider"] = provider
                st.session_state["llm_api_key"] = api_key

            best, candidates = parse_prompt_jp(prompt, list(df_eff.columns))
            if not best:
                st.warning("プロンプトの解釈に失敗しました。以下の解釈案を参考に修正してください。")
                st.code(json.dumps(candidates[:3], ensure_ascii=False, indent=2), language="json")
                return
            st.write("解釈内容（編集可能）")
            best_json = st.text_area("解釈JSON", value=json.dumps(best, ensure_ascii=False, indent=2), height=220)
            try:
                best_checked = json.loads(best_json)
            except Exception as e:
                st.error(f"JSONが不正です: {e}")
                return
            # 実行前確認
            st.info("内容を確認し、「実行」ボタンを押してください。")
            if st.button("実行"):
                rc = interpreted_to_runconfig(best_checked)
                rc.prompt_raw = prompt
                # 実行（rc.modeにより分岐）
                fdf = apply_filters(df_eff, rc.filters, "AND")
                if fdf.empty:
                    raise ValueError("フィルタ後のデータが空です。")

                viz = rc.viz or VizConfig(chart_type="棒")
                if rc.mode == "クロス集計":
                    if not rc.pivot_values:
                        # 件数
                        fdf = fdf.copy(); fdf["_count_"] = 1
                        res = pivot_aggregate(fdf, rc.pivot_index or [], rc.pivot_columns or [], ["_count_"], "sum", True)
                        result_df_local = res.rename(columns={"_count_": "件数"})
                        value_col = "件数"
                    else:
                        res = pivot_aggregate(fdf, rc.pivot_index or [], rc.pivot_columns or [], rc.pivot_values or [], rc.pivot_aggfunc or "count", True)
                        result_df_local = res
                        # 可視化用推定
                        value_col = None
                    st.subheader("結果")
                    # 可視化は run_pivot と同様のロジック
                    df_v = result_df_local.copy()
                    df_v.columns = [c if not isinstance(c, tuple) else "_".join([str(x) for x in c if x != ""]) for c in df_v.columns]
                    if rc.pivot_columns:
                        idx = rc.pivot_index or []
                        melt_id = idx
                        melt_val_vars = [c for c in df_v.columns if c not in idx]
                        dfm = df_v.melt(id_vars=melt_id, value_vars=melt_val_vars, var_name="系列", value_name="値")
                        label_col = melt_id[0] if melt_id else "index"
                        st.dataframe(result_df_local, use_container_width=True, height=300)
                        st.subheader("可視化")
                        render_chart_and_downloads(dfm, viz, label_col=label_col, value_col="値", series_col="系列")
                    else:
                        st.dataframe(result_df_local, use_container_width=True, height=350)
                        st.info("列（columns）が未指定のため、テーブルのみ表示しました。")
                    result_df = result_df_local
                    run_config = rc

                elif rc.mode == "グループ集計":
                    if not rc.groupby_cols:
                        raise ValueError("グループ化列が解釈されていません。")
                    if not rc.agg_map:
                        # 値未指定なら件数のみ
                        agg_map = {}
                        for c in fdf.columns:
                            if ptypes.is_numeric_dtype(fdf[c]):
                                agg_map[c] = ["mean"]
                        if not agg_map:
                            raise ValueError("集計対象の数値列が見つかりません。")
                        rc.agg_map = agg_map
                    res = group_aggregate(fdf, rc.groupby_cols, rc.agg_map)
                    result_df = res
                    st.subheader("結果")
                    agg_cols = [c for c in result_df.columns if c not in (rc.groupby_cols or [])]
                    val_col = None
                    if agg_cols:
                        mean_cols = [c for c in agg_cols if c.endswith("_mean")]
                        val_col = mean_cols[0] if mean_cols else agg_cols[0]
                    label_col = (rc.groupby_cols or [None])[0]
                    render_chart_and_downloads(result_df, viz, label_col=label_col, value_col=val_col)
                    run_config = rc

                elif rc.mode == "単純集計":
                    if not rc.simple_col:
                        raise ValueError("対象列が解釈されていません。")
                    res = simple_value_counts(fdf, rc.simple_col, normalize=rc.simple_normalize)
                    result_df = res
                    st.subheader("結果")
                    render_chart_and_downloads(result_df, viz, label_col=rc.simple_col, value_col=("割合" if rc.simple_normalize else "件数"))
                    run_config = rc

                elif rc.mode == "上位N":
                    if not rc.topn_col or not rc.topn_n:
                        raise ValueError("上位Nの対象列またはNが解釈されていません。")
                    res = top_n_categories(fdf, rc.topn_col, rc.topn_n)
                    result_df = res
                    st.subheader("結果")
                    render_chart_and_downloads(result_df, viz, label_col=rc.topn_col, value_col="件数")
                    run_config = rc

    except Exception as e:
        error_msg = str(e)
        st.error(error_msg)

    # エクスポート/インポート
    st.subheader("設定のエクスポート/インポート（再現性）")
    if result_df is not None:
        export_target = run_config or RunConfig(mode="単純集計", filters=[], logic="AND", exclude_columns=[])
        export_config(export_target)
    imported = import_config_ui()
    if imported:
        st.write("インポート内容（確認用）")
        st.code(json.dumps(asdict(imported), ensure_ascii=False, indent=2), language="json")

    # 履歴
    if run_config and result_df is not None and not error_msg:
        push_history(run_config)
    restored = render_history_restore()
    if restored:
        st.info("復元した設定をサイドバーに反映し、同様の操作を行ってください。")
        st.code(json.dumps(asdict(restored), ensure_ascii=False, indent=2), language="json")

    st.markdown("---")
    st.caption("目安: 10万行規模までストレスなく操作可能。大量データはサンプリング・上位Nでの可視化を推奨。")

if __name__ == "__main__":
    main()