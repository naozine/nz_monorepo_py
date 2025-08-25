from pathlib import Path
from typing import Union, List, Dict, Any
import sys

try:
    # openpyxl is the de-facto library for .xlsx read/write
    from openpyxl import load_workbook
    from openpyxl.utils import get_column_letter
    from openpyxl.styles import PatternFill
except Exception as e:  # pragma: no cover
    raise RuntimeError("openpyxl が必要です。`pip install openpyxl` を実行してください。") from e

try:
    import yaml  # PyYAML
except Exception as e:  # pragma: no cover
    yaml = None

try:
    from main import ReportDataPreparator, ReportConfig, ProcessedData
except Exception as e:  # pragma: no cover
    print(f"警告: main.py からのインポートに失敗しました: {e}")
    print("survey_data 機能は利用できません。")
    ReportDataPreparator = None


def fill_ac14(template_path: Union[str, Path], output_path: Union[str, Path], value: str = "xxx") -> Path:
    """
    report_template.xlsx を読み込み、シート p1 の AC14 セルを指定値に更新して保存します。

    :param template_path: 入力テンプレートのパス（.xlsx）
    :param output_path: 出力先ファイルのパス（.xlsx）
    :param value: AC14 に書き込む値（デフォルト: "xxx"）
    :return: 作成したファイルの Path
    """
    tpath = Path(template_path)
    if not tpath.exists():
        raise FileNotFoundError(f"テンプレートが見つかりません: {tpath}")

    wb = load_workbook(tpath)

    if "p1" not in wb.sheetnames:
        raise KeyError("テンプレートにシート 'p1' が見つかりません")

    ws = wb["p1"]
    # 値を書き込み、変更有無に応じて背景色を付与（変更: クリーム／未変更: 薄い水色）
    write_with_cream(ws, "AC14", value)

    out = Path(output_path)
    # 出力ディレクトリが存在しない場合に備える
    out.parent.mkdir(parents=True, exist_ok=True)

    wb.save(out)
    return out


def get_survey_data_value(survey_data_type: str, excel_path: Union[str, Path] = "survey.xlsx") -> Any:
    """
    survey_data_type に応じて、main.py の集計データから値を取得します。
    
    :param survey_data_type: 取得するデータタイプ
        - "total_responses": 総回答者数
        - "invalid_responses": 無効回答者数（未就学児等）
        - "effective_responses": 有効回答者数
    :param excel_path: アンケートExcelファイルのパス
    :return: 対応する値
    """
    if ReportDataPreparator is None:
        raise RuntimeError("main.py からのインポートが失敗しているため、survey_data 機能は利用できません。")
    
    try:
        config = ReportConfig()
        preparator = ReportDataPreparator(config)
        processed_data: ProcessedData = preparator.prepare_data(Path(excel_path))
        
        if survey_data_type == "total_responses":
            return processed_data.n_total + processed_data.n_preschool
        elif survey_data_type == "invalid_responses":
            return processed_data.n_preschool
        elif survey_data_type == "effective_responses":
            return processed_data.n_total
        else:
            raise ValueError(f"不明な survey_data_type: {survey_data_type}")
            
    except Exception as e:
        raise RuntimeError(f"集計データの取得に失敗しました: {e}")


def get_survey_data_series(series_type: str, excel_path: Union[str, Path] = "survey.xlsx") -> List[int]:
    """
    指定のシリーズ種別に応じて、人数配列を返します。
    
    サポートする series_type:
      - "elementary_boys": 小1〜小6の男子の回答者数（6個）
      - "elementary_girls": 小1〜小6の女子の回答者数（6個）
      - "junior_boys": 中1〜中3の男子の回答者数（3個）
      - "junior_girls": 中1〜中3の女子の回答者数（3個）
      - "region_grades:<地域名>": 指定地域の小1〜中3の回答者数（9個）
        例: "region_grades:東京23区" / "region_grades:三多摩島しょ" / "region_grades:埼玉県" など
      - "responses:q=<番号>;choice=<選択肢>;class=<grade|region>": 指定設問の特定選択肢の回答数を、学年または地域の順で返す
        例: "responses:q=1;choice=知っている;class=grade"
    
    :return: 人数リスト（series_typeに応じた順序）
    """
    if ReportDataPreparator is None:
        raise RuntimeError("main.py からのインポートが失敗しているため、survey_series 機能は利用できません。")

    try:
        config = ReportConfig()
        preparator = ReportDataPreparator(config)
        processed_data: ProcessedData = preparator.prepare_data(Path(excel_path))
        df = processed_data.df_effective

        # 地域名の正規化関数
        def normalize_region(name: str) -> str:
            nm = str(name).strip()
            if nm == "東京都下":
                return "三多摩島しょ"
            return nm

        # 新機能: 設問×選択肢の回答数シリーズ（学年または地域の順）
        if series_type.startswith("responses"):
            # 記法: "responses:q=<番号>;choice=<選択肢>;class=<grade|region>"
            # パーズ
            params_str = ""
            parts = series_type.split(":", 1)
            if len(parts) == 2:
                params_str = parts[1]
            # セミコロン区切りの key=value
            params: Dict[str, str] = {}
            if params_str:
                for token in params_str.split(";"):
                    if "=" in token:
                        k, v = token.split("=", 1)
                        params[k.strip().lower()] = v.strip()
            # 必須の3要素
            q_str = params.get("q") or params.get("question")
            choice = params.get("choice")
            class_type = params.get("class")
            if not q_str or not choice or not class_type:
                raise ValueError("responses シリーズには q(またはquestion) / choice / class を指定してください。例: responses:q=1;choice=知っている;class=grade")
            try:
                q_idx = int(q_str)
            except Exception:
                raise ValueError(f"responses の q は整数で指定してください（指定値: {q_str}）")
            class_type = class_type.lower()
            if class_type not in ("grade", "region"):
                raise ValueError("responses の class は 'grade' または 'region' を指定してください。")

            # 設問列名の解決（1開始）
            questions = processed_data.question_columns
            if q_idx < 1 or q_idx > len(questions):
                raise IndexError(f"responses: question 番号が範囲外です（1〜{len(questions)}）。指定: {q_idx}")
            qcol = questions[q_idx - 1]
            if qcol not in df.columns:
                raise RuntimeError(f"指定の設問列が見つかりません: {qcol}")

            # セルから選択肢集合を得る関数（改行区切りにも対応）
            def cell_to_set(val) -> set:
                if val is None:
                    return set()
                try:
                    import pandas as _pd  # 局所importでNaN検出
                    if _pd.isna(val):
                        return set()
                except Exception:
                    pass
                s = str(val).strip()
                if not s:
                    return set()
                import re as _re
                parts = _re.split(r"[\r\n]+", s)
                return {p.strip() for p in parts if p.strip()}

            # 集計
            if class_type == "grade":
                order = ["小1", "小2", "小3", "小4", "小5", "小6", "中1", "中2", "中3"]
                if "grade_2024" not in df.columns:
                    raise RuntimeError("必要な列 'grade_2024' が見つかりません。")
                counts: List[int] = []
                for g in order:
                    sub = df[df["grade_2024"] == g]
                    n = 0
                    for v in sub[qcol].dropna():
                        if choice in cell_to_set(v):
                            n += 1
                    counts.append(int(n))
                return counts
            else:
                order = ["東京23区", "三多摩島しょ", "埼玉県", "神奈川県", "千葉県", "その他"]
                if "region_bucket" not in df.columns:
                    raise RuntimeError("必要な列 'region_bucket' が見つかりません。")
                counts: List[int] = []
                for r in order:
                    sub = df[df["region_bucket"] == r]
                    n = 0
                    for v in sub[qcol].dropna():
                        if choice in cell_to_set(v):
                            n += 1
                    counts.append(int(n))
                return counts

        # 地域シリーズ（region_grades:...）の特別処理
        if series_type.startswith("region_grades"):
            # 記法1: "region_grades:地域名"
            region_name = None
            parts = series_type.split(":", 1)
            if len(parts) == 2 and parts[1].strip():
                region_name = parts[1].strip()
            # 記法2: YAML 側で region キーを使う場合は fill_from_yaml 側で地域名を補完して再呼び出しする想定
            # ここでは parts に地域名があればそれを使う
            if not region_name:
                raise ValueError("region_grades の地域名が指定されていません。'region_grades:東京23区' のように指定してください。")
            reg = normalize_region(region_name)
            allowed = {"東京23区", "三多摩島しょ", "埼玉県", "神奈川県", "千葉県", "その他"}
            if reg not in allowed:
                raise ValueError(f"不明な地域名: {region_name}（許可: 東京23区, 三多摩島しょ, 埼玉県, 神奈川県, 千葉県, その他）")

            # 必要な列確認
            for col in ["region_bucket", "grade_2024"]:
                if col not in df.columns:
                    raise RuntimeError(f"必要な列 '{col}' が見つかりません。Excelや前処理の仕様をご確認ください。")

            grades = ["小1", "小2", "小3", "小4", "小5", "小6", "中1", "中2", "中3"]
            counts: List[int] = []
            mask_reg = (df["region_bucket"] == reg)
            for g in grades:
                counts.append(int(((df["grade_2024"] == g) & mask_reg).sum()))
            return counts

        # 既存：性別×学年シリーズ
        series_def = {
            "elementary_boys": {
                "level": "小学校",
                "gender": "男性",
                "grades": ["小1", "小2", "小3", "小4", "小5", "小6"],
            },
            "elementary_girls": {
                "level": "小学校",
                "gender": "女性",
                "grades": ["小1", "小2", "小3", "小4", "小5", "小6"],
            },
            "junior_boys": {
                "level": "中学校",
                "gender": "男性",
                "grades": ["中1", "中2", "中3"],
            },
            "junior_girls": {
                "level": "中学校",
                "gender": "女性",
                "grades": ["中1", "中2", "中3"],
            },
        }

        if series_type not in series_def:
            raise ValueError(f"不明な survey_series: {series_type}")

        defn = series_def[series_type]
        level = defn["level"]
        gender = defn["gender"]
        grades = defn["grades"]

        # フィルタ列の存在確認
        for col in ["school_level", "gender_norm", "grade_2024"]:
            if col not in df.columns:
                raise RuntimeError(f"必要な列 '{col}' が見つかりません。Excelや前処理の仕様をご確認ください。")

        mask_level_gender = (df["school_level"] == level) & (df["gender_norm"] == gender)
        counts: List[int] = []
        for g in grades:
            counts.append(int(((df["grade_2024"] == g) & mask_level_gender).sum()))
        return counts

    except Exception as e:
        raise RuntimeError(f"シリーズ集計の取得に失敗しました: {e}")


# クリーム色（変更時）と薄い水色（未変更時）の塗りつぶし
CREAM_FILL = PatternFill(fill_type="solid", start_color="FFF2CC", end_color="FFF2CC")
BLUE_FILL = PatternFill(fill_type="solid", start_color="DDEBF7", end_color="DDEBF7")


def write_with_cream(ws, addr: str, new_value: Any):
    """セルの値を設定し、
    - 値が変わった場合: クリーム色（FFF2CC）
    - 値が変わらない場合: 薄い水色（DDEBF7）
    の背景を付与する。
    """
    try:
        old_value = ws[addr].value
    except Exception:
        old_value = None
    ws[addr] = new_value
    try:
        if old_value != new_value:
            ws[addr].fill = CREAM_FILL
        else:
            ws[addr].fill = BLUE_FILL
    except Exception:
        # スタイル設定に失敗しても処理を継続
        pass


def fill_from_yaml(config_path: Union[str, Path]) -> Path:
    """
    YAML 設定を読み込み、指定のシート/セルへ値を書き込みます。

    想定する YAML 例:

    template: report_template.xlsx
    output: report_result.xlsx
    survey_excel: survey.xlsx  # オプション：アンケートExcelファイルのパス
    writes:
      - sheet: p1
        cell: AC14
        value: xxx
      - sheet: p2
        row: 5
        column: B   # または 2 （数値でも可）
        value: Hello
      - sheet: p1
        cell: B10
        survey_data: total_responses      # 総回答者数
      - sheet: p1
        cell: B11
        survey_data: invalid_responses    # 無効回答者数
      - sheet: p1
        cell: B12
        survey_data: effective_responses  # 有効回答者数

      # 新機能: 連続セルに学年別の人数を書き込み（デフォルトは下方向）
      - sheet: p1
        cell: D20      # このセルに小1男子、その下に小2男子…小6男子まで
        survey_series: elementary_boys   # elementary_boys | elementary_girls | junior_boys | junior_girls
        # direction: down  # 省略可（right も指定可能）

      # 新機能: 地域別（小1〜中3）を横方向に書き込み（デフォルトは右方向）
      - sheet: p1
        cell: E25
        survey_series: region_grades
        region: 東京23区  # 許可: 東京23区 / 三多摩島しょ / 埼玉県 / 神奈川県 / 千葉県 / その他
        # direction: right  # 省略可（未指定なら右）
        # または、region_grades:東京23区 のように survey_series に地域名を含めても可

      # 追加機能: 指定設問の指定選択肢の回答数（学年または地域ごと）を縦方向に書き込み
      - sheet: Q1
        cell: T10
        survey_series: responses
        question: 1          # main.py の設問一覧順（1開始）
        choice: 知っている
        class: grade         # grade | region
        # direction: down    # 省略可（responses は縦方向がデフォルト）

    :param config_path: YAML ファイルへのパス
    :return: 出力されたファイルの Path
    """
    if yaml is None:
        raise RuntimeError("PyYAML が必要です。`pip install pyyaml` を実行してください。")

    cpath = Path(config_path)
    if not cpath.exists():
        raise FileNotFoundError(f"設定ファイルが見つかりません: {cpath}")

    with cpath.open("r", encoding="utf-8") as f:
        data = yaml.safe_load(f) or {}

    if not isinstance(data, dict):
        raise ValueError("YAML のルートはマッピングである必要があります（dict）。")

    # 必須キー
    template = data.get("template")
    output = data.get("output")
    writes = data.get("writes")
    survey_excel = data.get("survey_excel", "survey.xlsx")  # デフォルトは survey.xlsx

    if not template:
        raise ValueError("YAML に 'template' が必要です。")
    if not output:
        raise ValueError("YAML に 'output' が必要です。")
    if not isinstance(writes, list) or not writes:
        raise ValueError("YAML の 'writes' は1件以上のリストである必要があります。")

    # パスは設定ファイルの場所からの相対も許可
    base = cpath.parent
    tpath = (base / template) if not Path(str(template)).is_absolute() else Path(str(template))
    out = (base / output) if not Path(str(output)).is_absolute() else Path(str(output))
    survey_path = (base / survey_excel) if not Path(str(survey_excel)).is_absolute() else Path(str(survey_excel))

    if not tpath.exists():
        raise FileNotFoundError(f"テンプレートが見つかりません: {tpath}")

    wb = load_workbook(tpath)

    for i, w in enumerate(writes, start=1):
        if not isinstance(w, dict):
            raise ValueError(f"writes[{i}] はマッピングである必要があります。")
        sheet = w.get("sheet")
        if not sheet:
            raise ValueError(f"writes[{i}] に 'sheet' がありません。")
        if sheet not in wb.sheetnames:
            raise KeyError(f"シートが見つかりません: '{sheet}'（writes[{i}]）")

        # アドレスの決定: cell 優先、なければ row+column
        addr = w.get("cell")
        if not addr:
            row = w.get("row")
            col = w.get("column")
            if row is None or col is None:
                raise ValueError(f"writes[{i}] には 'cell' か 'row'+'column' のいずれかが必要です。")
            # column は文字列（例: 'AC'）または数値（例: 29）を許可
            if isinstance(col, int):
                col_letter = get_column_letter(col)
            else:
                col_letter = str(col).strip()
                if not col_letter:
                    raise ValueError(f"writes[{i}] の column が不正です。")
            addr = f"{col_letter}{int(row)}"

        # シリーズ書き込みか単一値かを判定
        series_type = w.get("survey_series")
        value = w.get("value")
        survey_data_type = w.get("survey_data")

        # 相互排他
        specified = [x is not None for x in (value, survey_data_type, series_type)]
        if sum(specified) != 1:
            raise ValueError(f"writes[{i}] では 'value' / 'survey_data' / 'survey_series' のいずれか1つだけを指定してください。")

        ws = wb[sheet]

        if series_type is not None:
            # region_grades の地域名を補完（survey_series が文字列 'region_grades' かつ region キーが与えられている場合）
            region_param = w.get("region")
            if isinstance(series_type, str) and series_type.strip() == "region_grades":
                if region_param:
                    series_type = f"region_grades:{str(region_param).strip()}"
                else:
                    raise ValueError(f"writes[{i}] の region_grades には 'region' キーで地域名を指定してください（例: region: 東京23区）。")

            # responses の補完（question / choice / class を付与）
            if isinstance(series_type, str) and series_type.strip() == "responses":
                q = w.get("question")
                ch = w.get("choice")
                cls = w.get("class")
                if q is None or ch is None or cls is None:
                    raise ValueError(f"writes[{i}] の responses には question / choice / class を指定してください。")
                try:
                    q_int = int(q)
                except Exception:
                    raise ValueError(f"writes[{i}] の question は整数で指定してください（指定値: {q}）。")
                series_type = f"responses:q={q_int};choice={str(ch).strip()};class={str(cls).strip()}"

            # 連続セルにシリーズを書き込み
            # 方向のデフォルト: 通常は down、region_grades は right
            default_direction = "right" if isinstance(series_type, str) and series_type.startswith("region_grades") else "down"
            direction = (w.get("direction") or default_direction).lower()
            if direction not in ("down", "right"):
                raise ValueError(f"writes[{i}] の direction は 'down' または 'right' で指定してください。")
            try:
                values = get_survey_data_series(series_type, survey_path)
            except Exception as e:
                raise RuntimeError(f"writes[{i}] の survey_series '{series_type}' の取得に失敗: {e}")

            # アドレス分解
            col_letters = ''.join([c for c in addr if c.isalpha()])
            row_digits = ''.join([c for c in addr if c.isdigit()])
            if not col_letters or not row_digits:
                raise ValueError(f"writes[{i}] の cell アドレスが不正です: {addr}")
            base_col = col_letters
            base_row = int(row_digits)

            if direction == "down":
                for idx, v in enumerate(values):
                    target = f"{base_col}{base_row + idx}"
                    write_with_cream(ws, target, v)
            else:  # right
                # 右方向も対応（必要なら）。
                # 列文字を番号に変換してインクリメント
                def col_to_num(s: str) -> int:
                    num = 0
                    for ch in s:
                        num = num * 26 + (ord(ch.upper()) - ord('A') + 1)
                    return num
                def num_to_col(n: int) -> str:
                    res = ""
                    while n > 0:
                        n, rem = divmod(n - 1, 26)
                        res = chr(ord('A') + rem) + res
                    return res
                base_col_num = col_to_num(base_col)
                for idx, v in enumerate(values):
                    col_letter = num_to_col(base_col_num + idx)
                    target = f"{col_letter}{base_row}"
                    write_with_cream(ws, target, v)
        else:
            # 単一セル書き込み
            if value is not None:
                final_value = value
            else:
                try:
                    final_value = get_survey_data_value(survey_data_type, survey_path)
                except Exception as e:
                    raise RuntimeError(f"writes[{i}] の survey_data '{survey_data_type}' の取得に失敗: {e}")
            write_with_cream(ws, addr, final_value)

    # 出力ディレクトリが存在しない場合に備える
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out)
    return out


def main():
    """
    YAML 設定で一括書き込みを実行します（デフォルト）。

    使い方:
      python fill_template_excel.py <config.yaml>

    例:
      python fill_template_excel.py apps/report_sgk/sample_fill.yaml

    備考:
      以前の「引数なしで p1!AC14 に 'xxx' を書き込む」動作は削除されました。
    """
    if len(sys.argv) < 2:
        print("エラー: YAML 設定ファイルへのパスを指定してください。\n"
              "使い方: python fill_template_excel.py <config.yaml>\n"
              "例:     python fill_template_excel.py apps/report_sgk/sample_fill.yaml")
        sys.exit(2)

    cfg = Path(sys.argv[1])
    out_path = fill_from_yaml(cfg)
    print(f"YAML 設定に従い書き出し完了: {out_path}")


if __name__ == "__main__":
    main()
