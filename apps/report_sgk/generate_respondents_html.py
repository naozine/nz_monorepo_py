from __future__ import annotations
import json
import os
import sys
from dataclasses import asdict
from datetime import datetime
from pathlib import Path
from typing import Any, List, Dict

import pandas as pd

# Ensure project root is importable for `apps.*`
PROJECT_ROOT = Path(__file__).resolve().parents[2]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

# Reuse classes and helpers from main.py
from apps.report_sgk.main import ReportConfig, ReportDataPreparator, _load_env_from_dotenv, cell_to_unique_set

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


def _format_birthdate(val: Any) -> str:
    if pd.isna(val):
        return ""
    # Accept pandas Timestamp, datetime, date, or string
    try:
        if isinstance(val, pd.Timestamp):
            dt = val.to_pydatetime()
        elif isinstance(val, datetime):
            dt = val
        else:
            # Try parse from string or excel serial
            if isinstance(val, (int, float)):
                # Excel serial to datetime via pandas
                dt = pd.Timestamp('1899-12-30') + pd.to_timedelta(int(val), unit='D')
                dt = dt.to_pydatetime()
            else:
                dt = pd.to_datetime(str(val), errors='coerce')
                if pd.isna(dt):
                    return str(val)
                dt = dt.to_pydatetime()
        return dt.strftime('%Y-%m-%d')
    except Exception:
        return str(val)


def extract_rows(df: pd.DataFrame) -> List[Dict[str, Any]]:
    # Ensure required columns exist; use empty if missing
    base_cols = {
        '生年月日': '生年月日',
        '学年': 'grade_2024',
        '都道府県': '都道府県',
        '市区町村': '市区町村',
    }
    for src in base_cols.values():
        if src not in df.columns:
            # Create empty column if not present
            df[src] = None

    # Ensure Q6 column exists gracefully
    has_q6 = LEARNING_Q_COL in df.columns

    out = []
    for _, row in df.iterrows():
        birth_display = row.get('生年月日')
        # Prefer formatted birth_dt if available
        if 'birth_dt' in df.columns and pd.notna(row.get('birth_dt')):
            birth_display = row.get('birth_dt')
        birth_str = _format_birthdate(birth_display)

        rec: Dict[str, Any] = {
            '生年月日': birth_str,
            '学年': row.get('grade_2024') if pd.notna(row.get('grade_2024')) else '',
            '都道府県': row.get('都道府県') if pd.notna(row.get('都道府県')) else '',
            '市区町村': row.get('市区町村') if pd.notna(row.get('市区町村')) else '',
        }

        # Q6 options as checkmarks and count
        selected_set = set()
        if has_q6:
            val = row.get(LEARNING_Q_COL)
            try:
                selected_set = set(cell_to_unique_set(val)) if pd.notna(val) else set()
            except Exception:
                selected_set = set()
        # Fill columns for each option
        cnt = 0
        for opt in LEARNING_OPTIONS:
            checked = '✅' if opt in selected_set else ''
            if checked:
                cnt += 1
            rec[opt] = checked
        rec['回答数'] = cnt if has_q6 else 0

        out.append(rec)
    return out


def build_html(data: List[Dict[str, Any]]) -> str:
    # Build a self-contained HTML with embedded JSON, sortable and filterable per column (vanilla JS)
    columns = ['生年月日', '学年', '都道府県', '市区町村'] + LEARNING_OPTIONS + ['回答数']
    json_data = json.dumps(data, ensure_ascii=False)
    return f"""<!doctype html>
<html lang=\"ja\">
<head>
  <meta charset=\"utf-8\" />
  <meta name=\"viewport\" content=\"width=device-width, initial-scale=1\" />
  <title>回答者一覧</title>
  <style>
    body {{ font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', 'Hiragino Kaku Gothic ProN', 'Hiragino Sans', Meiryo, sans-serif; color: #222; padding: 16px; }}
    h1 {{ font-size: 20px; margin: 0 0 12px; }}
    .controls {{ display: flex; gap: 16px; align-items: center; margin: 8px 0 12px; }}
    .table-wrap {{ overflow-x: auto; }}
    table {{ border-collapse: collapse; width: 100%; min-width: 900px; }}
    th, td {{ border: 1px solid #e0e0e0; padding: 6px 8px; font-size: 14px; text-align: left; }}
    th {{ background: #f7f7f7; position: sticky; top: 0; z-index: 1; cursor: pointer; user-select: none; }}
    .filters input {{ width: 100%; box-sizing: border-box; padding: 4px 6px; font-size: 13px; }}
    .count {{ color: #555; margin-bottom: 8px; }}
    .hint {{ color: #777; font-size: 12px; margin-bottom: 0; }}
    .sortable::after {{ content: ' \25B4\25BE'; font-size: 10px; color: #999; margin-left: 4px; }}
  </style>
</head>
<body>
  <h1>回答者一覧</h1>
  <div class=\"controls\">
    <div class=\"hint\">各列のテキストボックスでフィルタ、ヘッダークリックでソートできます。</div>
    <label style=\"font-size:13px; user-select:none;\"><input type=\"checkbox\" id=\"excludeNonTarget\" /> 対象外の学年を除く</label>
  </div>
  <div class=\"count\" id=\"count\"></div>
  <div class=\"table-wrap\">
    <table id=\"respTable\">
      <thead>
        <tr id=\"headerRow\"></tr>
        <tr class=\"filters\" id=\"filterRow\"></tr>
      </thead>
      <tbody id=\"tbody\"></tbody>
    </table>
  </div>
  <script>
    const DATA = {json.dumps({'rows': data}, ensure_ascii=False)}.rows; // embedded
    const COLUMNS = {json.dumps(columns, ensure_ascii=False)};

    const state = {{ sortKey: null, sortDir: 'asc', filters: Object.fromEntries(COLUMNS.map(c => [c, ''])), excludeNonTarget: false }};

    function compare(a, b, key) {{
      const va = (a[key] ?? '').toString();
      const vb = (b[key] ?? '').toString();
      // Detect YYYY-MM-DD date
      const re = /^\d{{4}}-\d{{2}}-\d{{2}}$/;
      if (re.test(va) && re.test(vb)) {{
        if (va < vb) return -1; if (va > vb) return 1; return 0;
      }}
      // Numeric compare if both are numeric
      const na = parseFloat(va.replace(/[^0-9.-]/g, ''));
      const nb = parseFloat(vb.replace(/[^0-9.-]/g, ''));
      const isNa = !isNaN(na) && va.trim() !== '' && /^[-+]?\d/.test(va);
      const isNb = !isNaN(nb) && vb.trim() !== '' && /^[-+]?\d/.test(vb);
      if (isNa && isNb) {{
        return na - nb;
      }}
      return va.localeCompare(vb, 'ja');
    }}

    function applyFilters(rows) {{
      return rows.filter(r => {{
        // When checked, exclude rows with 学年 === '対象外'
        if (state.excludeNonTarget && (r['学年'] ?? '') === '対象外') return false;
        return COLUMNS.every(col => {{
          const f = (state.filters[col] || '').trim();
          if (!f) return true;
          const v = (r[col] ?? '').toString();
          // Simple case-insensitive contains for Latin; normal contains for Japanese
          return v.toLowerCase().includes(f.toLowerCase());
        }});
      }});
    }}

    function sortRows(rows) {{
      if (!state.sortKey) return rows;
      const arr = rows.slice().sort((a,b) => compare(a,b,state.sortKey));
      if (state.sortDir === 'desc') arr.reverse();
      return arr;
    }}

    function renderTable() {{
      let rows = applyFilters(DATA);
      rows = sortRows(rows);

      const tbody = document.getElementById('tbody');
      tbody.innerHTML = '';
      const frag = document.createDocumentFragment();
      for (const r of rows) {{
        const tr = document.createElement('tr');
        for (const col of COLUMNS) {{
          const td = document.createElement('td');
          td.textContent = r[col] ?? '';
          tr.appendChild(td);
        }}
        frag.appendChild(tr);
      }}
      tbody.appendChild(frag);
      document.getElementById('count').textContent = `表示件数: ${{rows.length}} / 総件数: ${{DATA.length}}`;
    }}

    function setupHeader() {{
      const headerRow = document.getElementById('headerRow');
      const filterRow = document.getElementById('filterRow');
      for (const col of COLUMNS) {{
        const th = document.createElement('th');
        th.textContent = col;
        th.classList.add('sortable');
        th.addEventListener('click', () => {{
          if (state.sortKey === col) {{
            state.sortDir = state.sortDir === 'asc' ? 'desc' : 'asc';
          }} else {{
            state.sortKey = col;
            state.sortDir = 'asc';
          }}
          renderTable();
        }});
        headerRow.appendChild(th);

        const fth = document.createElement('th');
        const input = document.createElement('input');
        input.type = 'text';
        input.placeholder = 'フィルタ';
        input.value = state.filters[col] || '';
        input.addEventListener('input', (e) => {{
          state.filters[col] = e.target.value;
          renderTable();
        }});
        fth.appendChild(input);
        filterRow.appendChild(fth);
      }}
    }}

    // Hook up checkbox for excluding non-target grades
    const chk = document.getElementById('excludeNonTarget');
    if (chk) {{
      chk.checked = !!state.excludeNonTarget;
      chk.addEventListener('change', (e) => {{
        state.excludeNonTarget = e.target.checked;
        renderTable();
      }});
    }}

    setupHeader();
    renderTable();
  </script>
</body>
</html>
"""


def generate(output_path: Path | None = None) -> Path:
    _load_env_from_dotenv()
    report_config = ReportConfig.from_env()
    preparator = ReportDataPreparator(report_config)

    survey_file = os.getenv('SURVEY_EXCEL_FILE', 'survey.xlsx')
    excel_path = Path(__file__).parent / survey_file
    if not excel_path.exists():
        raise FileNotFoundError(f"Excelファイルが見つかりません: {excel_path}")

    processed = preparator.prepare_data(excel_path)
    # Use df_original which contains computed columns and includes all respondents
    rows = extract_rows(processed.df_original)

    html = build_html(rows)

    if output_path is None:
        output_path = Path(__file__).parent / 'respondents.html'
    else:
        output_path = Path(output_path)

    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html)

    return output_path


if __name__ == '__main__':
    path = generate()
    print(f"回答者一覧HTMLを出力しました: {path}")
