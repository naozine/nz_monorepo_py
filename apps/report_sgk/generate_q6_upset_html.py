from __future__ import annotations
import os
import sys
from pathlib import Path
from typing import Any, Dict, List, Set

import pandas as pd

# Ensure project root is importable for `apps.*`
PROJECT_ROOT = Path(__file__).resolve().parents[2]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

# Reuse classes and helpers from main.py
from apps.report_sgk.main import (
    ReportConfig,
    ReportDataPreparator,
    _load_env_from_dotenv,
    cell_to_unique_set,
)

# Q6: 現在習い事や塾などに通われていますか？（複数回答可）
LEARNING_Q_COL = "現在習い事や塾などに通われていますか？（複数回答可）"
LEARNING_OPTIONS = [
    "学習塾(集団)",
    "学習塾(個別)",
    "家庭教師",
    "語学教室",
    "その他",
    "通っていない",
]


def _normalize_learning_choices(raw_choices: Set[str]) -> Set[str]:
    """値のゆらぎ対策（例: 「英会話」「英会話・語学教室」などを「語学教室」に正規化）"""
    norm: Set[str] = set()
    # 語学系（代表ラベル: 語学教室）
    if any((("語学" in t) or ("英会話" in t)) for t in raw_choices):
        norm.add("語学教室")
    # 既存の選択肢と完全一致するものはそのまま追加
    for t in raw_choices:
        if t in LEARNING_OPTIONS:
            norm.add(t)
    return norm


def extract_q6_memberships(df: pd.DataFrame) -> List[Set[str]]:
    """各回答者のQ6選択肢セット（正規化後）を返す。"""
    memberships: List[Set[str]] = []
    if LEARNING_Q_COL not in df.columns:
        return memberships
    for _, row in df.iterrows():
        val = row.get(LEARNING_Q_COL)
        try:
            raw_set = set(cell_to_unique_set(val)) if pd.notna(val) else set()
            selected_set = _normalize_learning_choices(raw_set)
        except Exception:
            selected_set = set()
        memberships.append(selected_set)
    return memberships


def _build_upset_html(memberships: List[Set[str]], options: List[str]) -> str:
    """
    外部ライブラリに依存せず、アップセット図風の可視化を行うシンプルなHTML/JSを返す。
    - 上段: 各セット（選択肢）の単独出現数のバー
    - 中央: 組合せ（交差）の行ごとのドットマトリクス + 左に交差サイズのバー
    - コントロール: 上位N件 / 0件除外 / 「通っていない」を含めるか
    """
    # 埋め込み用データをPython側で前計算（集合→カウント）
    from collections import Counter

    # set size counts (singletons)
    set_sizes: Dict[str, int] = {opt: 0 for opt in options}
    # combination counts (use tuple of sorted options as key)
    combo_counter: Counter = Counter()

    for s in memberships:
        # increment set sizes
        for opt in s:
            if opt in set_sizes:
                set_sizes[opt] += 1
        # combinations: ignore empty set
        if len(s) > 0:
            key = tuple(sorted(x for x in s if x in options))
            if key:
                combo_counter[key] += 1

    # Convert to lists for JSON embedding
    set_sizes_list = [{"name": k, "count": v} for k, v in set_sizes.items()]
    combo_list = [{"sets": list(k), "count": c} for k, c in combo_counter.items()]

    import json
    data_json = json.dumps({
        "options": options,
        "setSizes": set_sizes_list,
        "combos": combo_list,
    }, ensure_ascii=False)

    # HTML/JS rendering (vanilla) - use placeholder to avoid f-string brace escaping
    html_tpl = """<!doctype html>
<html lang=\"ja\">
<head>
  <meta charset=\"utf-8\" />
  <meta name=\"viewport\" content=\"width=device-width, initial-scale=1\" />
  <title>Q6 アップセット図</title>
  <style>
    :root { --bar-color: #4e79a7; --bar2-color: #f28e2b; --dot: #333; --line: #999; }
    body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', 'Hiragino Kaku Gothic ProN', 'Hiragino Sans', Meiryo, sans-serif; color: #222; padding: 16px; }
    h1 { font-size: 20px; margin: 0 0 12px; }
    .controls { display: flex; flex-wrap: wrap; gap: 16px; align-items: center; margin: 10px 0 14px; font-size: 13px; }
    .chart-wrap { display: grid; grid-template-columns: 220px 1fr; grid-template-rows: auto auto; gap: 8px 12px; align-items: end; }
    .set-bars-title { grid-column: 2; font-size: 12px; color: #666; }
    .set-bars { grid-column: 2; display: flex; align-items: flex-end; gap: 12px; height: 140px; border-bottom: 1px solid #eee; }
    .set-bar { width: 36px; background: var(--bar2-color); display: flex; align-items: flex-end; justify-content: center; position: relative; }
    .set-bar .value { position: absolute; bottom: 100%; transform: translateY(-4px); font-size: 11px; color: #444; }
    .set-labels { grid-column: 2; display: flex; gap: 12px; }
    .set-label { width: 36px; writing-mode: vertical-rl; transform: rotate(180deg); text-align: left; font-size: 12px; color: #333; }

    .matrix { grid-column: 2; }
    .matrix-row { display: grid; grid-template-columns: repeat(var(--ncols), 28px); gap: 12px; align-items: center; margin: 6px 0; }
    .dot { width: 10px; height: 10px; border-radius: 50%; background: var(--dot); margin: 0 auto; position: relative; }
    .line { height: 2px; background: var(--line); position: relative; top: -6px; grid-column: var(--line-start) / var(--line-end); }

    .combo-area { grid-column: 1 / span 2; display: grid; grid-template-columns: 220px 1fr; column-gap: 12px; }
    .combo-left { border-right: 1px solid #eee; padding-right: 8px; }
    .combo-bar { height: 16px; background: var(--bar-color); margin: 8px 0; position: relative; }
    .combo-bar .value { position: absolute; left: 100%; margin-left: 6px; top: 50%; transform: translateY(-50%); font-size: 12px; color: #444; }
    .combo-label { font-size: 12px; color: #333; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
    .muted { color: #777; }
  </style>
</head>
<body>
  <h1>Q6 アップセット図</h1>
  <div class=\"controls\">
    <label><input type=\"checkbox\" id=\"toggleNone\" checked /> 「通っていない」を含める</label>
    <label><input type=\"checkbox\" id=\"hideZero\" checked /> 0件のセット/交差は非表示</label>
    <label>上位交差: <input type=\"number\" id=\"topN\" value=\"20\" min=\"1\" max=\"200\" style=\"width:60px\" /></label>
  </div>
  <div id=\"chart\"></div>

  <script>
    const RAW = __DATA__;

    function computeMax(arr, key) {
      return arr.reduce((m, x) => Math.max(m, x[key]||0), 0);
    }

    function render() {
      const includeNone = document.getElementById('toggleNone').checked;
      const hideZero = document.getElementById('hideZero').checked;
      const topN = Math.max(1, parseInt(document.getElementById('topN').value || '20'));

      const options = RAW.options.slice();
      const noneIdx = options.indexOf('通っていない');
      let opts = options.slice();
      if (!includeNone && noneIdx >= 0) opts.splice(noneIdx, 1);

      const setSizes = RAW.setSizes.filter(d => opts.includes(d.name));
      const combos = RAW.combos
        .map(c => ({ sets: c.sets.filter(s => opts.includes(s)), count: c.count }))
        .filter(c => c.sets.length > 0);

      const setSizesFiltered = hideZero ? setSizes.filter(d => d.count > 0) : setSizes;

      // sort sets by count desc
      setSizesFiltered.sort((a, b) => b.count - a.count || opts.indexOf(a.name) - opts.indexOf(b.name));
      const colOrder = setSizesFiltered.map(d => d.name);

      // filter and sort combos
      let combosFiltered = combos.filter(c => c.sets.every(s => colOrder.includes(s)));
      combosFiltered = hideZero ? combosFiltered.filter(c => c.count > 0) : combosFiltered;
      combosFiltered.sort((a, b) => b.count - a.count || a.sets.length - b.sets.length || a.sets.join('\\u0001').localeCompare(b.sets.join('\\u0001'), 'ja'));
      combosFiltered = combosFiltered.slice(0, topN);

      // Build HTML
      const mount = document.getElementById('chart');
      mount.innerHTML = '';

      // Top set size bars
      const wrap = document.createElement('div');
      wrap.className = 'chart-wrap';

      const setBarsTitle = document.createElement('div');
      setBarsTitle.className = 'set-bars-title';
      setBarsTitle.textContent = '各セット（選択肢）の出現数';
      wrap.appendChild(setBarsTitle);

      const setBars = document.createElement('div');
      setBars.className = 'set-bars';
      const maxSet = Math.max(1, computeMax(setSizesFiltered, 'count'));
      for (const s of setSizesFiltered) {
        const bar = document.createElement('div');
        bar.className = 'set-bar';
        bar.style.height = (Math.round((s.count / maxSet) * 100)) + '%';
        const v = document.createElement('div'); v.className = 'value'; v.textContent = s.count;
        bar.appendChild(v);
        setBars.appendChild(bar);
      }
      wrap.appendChild(setBars);

      const setLabels = document.createElement('div');
      setLabels.className = 'set-labels';
      for (const name of colOrder) {
        const l = document.createElement('div');
        l.className = 'set-label';
        l.textContent = name;
        setLabels.appendChild(l);
      }
      wrap.appendChild(setLabels);

      // Combo matrix + left bars
      const comboArea = document.createElement('div');
      comboArea.className = 'combo-area';

      const left = document.createElement('div');
      left.className = 'combo-left';
      const maxCombo = Math.max(1, computeMax(combosFiltered, 'count'));
      for (const c of combosFiltered) {
        const lbl = document.createElement('div');
        lbl.className = 'combo-label muted';
        lbl.title = c.sets.join(' ∩ ');
        lbl.textContent = c.sets.join(' ∩ ');
        left.appendChild(lbl);
        const bar = document.createElement('div');
        bar.className = 'combo-bar';
        bar.style.width = (Math.round((c.count / maxCombo) * 100)) + '%';
        const v = document.createElement('div'); v.className = 'value'; v.textContent = c.count;
        bar.appendChild(v);
        left.appendChild(bar);
      }
      comboArea.appendChild(left);

      const matrix = document.createElement('div');
      matrix.className = 'matrix';
      matrix.style.setProperty('--ncols', String(colOrder.length));

      for (const c of combosFiltered) {
        const row = document.createElement('div');
        row.className = 'matrix-row';
        const present = new Set(c.sets);
        const firstIdx = Math.min(...c.sets.map(s => colOrder.indexOf(s)));
        const lastIdx = Math.max(...c.sets.map(s => colOrder.indexOf(s)));
        for (let i = 0; i < colOrder.length; i++) {
          const cell = document.createElement('div');
          if (present.has(colOrder[i])) {
            const dot = document.createElement('div'); dot.className = 'dot'; cell.appendChild(dot);
          } else {
            cell.innerHTML = '&nbsp;'
          }
          row.appendChild(cell);
        }
        // connecting line
        const line = document.createElement('div');
        line.className = 'line';
        // CSS grid columns are 1-based; add 1 for start; +1 for end because it's exclusive
        line.style.setProperty('--line-start', String(firstIdx + 1));
        line.style.setProperty('--line-end', String(lastIdx + 2));
        row.appendChild(line);
        matrix.appendChild(row);
      }
      comboArea.appendChild(matrix);

      wrap.appendChild(comboArea);
      mount.appendChild(wrap);
    }

    document.addEventListener('DOMContentLoaded', () => {
      document.getElementById('toggleNone').addEventListener('change', render);
      document.getElementById('hideZero').addEventListener('change', render);
      document.getElementById('topN').addEventListener('input', render);
      render();
    });
  </script>
</body>
</html>
"""
    return html_tpl.replace("__DATA__", data_json)


def generate(output_path: Path | None = None) -> Path:
    """Q6のアップセット図HTMLを生成する。"""
    _load_env_from_dotenv()
    report_config = ReportConfig.from_env()
    preparator = ReportDataPreparator(report_config)

    survey_file = os.getenv('SURVEY_EXCEL_FILE', 'survey.xlsx')
    excel_path = Path(__file__).parent / survey_file
    if not excel_path.exists():
        raise FileNotFoundError(f"Excelファイルが見つかりません: {excel_path}")

    processed = preparator.prepare_data(excel_path)
    df = processed.df_original

    memberships = extract_q6_memberships(df)
    html = _build_upset_html(memberships, LEARNING_OPTIONS)

    if output_path is None:
        output_path = Path(__file__).parent / 'q6_upset.html'
    else:
        output_path = Path(output_path)

    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html)

    return output_path


if __name__ == '__main__':
    p = generate()
    print(f"Q6 アップセット図HTMLを出力しました: {p}")
