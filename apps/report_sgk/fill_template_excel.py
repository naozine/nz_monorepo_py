from pathlib import Path
from typing import Union, List, Dict, Any
import sys

try:
    # openpyxl is the de-facto library for .xlsx read/write
    from openpyxl import load_workbook
    from openpyxl.utils import get_column_letter
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
    ws["AC14"] = value

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
    指定のシリーズ種別に応じて、学年別×性別の人数配列を返します。
    
    サポートする series_type:
      - "elementary_boys": 小1〜小6の男子の回答者数（6個）
      - "elementary_girls": 小1〜小6の女子の回答者数（6個）
      - "junior_boys": 中1〜中3の男子の回答者数（3個）
      - "junior_girls": 中1〜中3の女子の回答者数（3個）
    
    :return: 各学年の人数リスト（順序は学年の昇順）
    """
    if ReportDataPreparator is None:
        raise RuntimeError("main.py からのインポートが失敗しているため、survey_series 機能は利用できません。")

    try:
        config = ReportConfig()
        preparator = ReportDataPreparator(config)
        processed_data: ProcessedData = preparator.prepare_data(Path(excel_path))
        df = processed_data.df_effective

        # 定義（日本語ラベルに依存）
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
            # 連続セル（デフォルト: 縦）にシリーズを書き込み
            direction = (w.get("direction") or "down").lower()
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
                    ws[target] = v
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
                    ws[target] = v
        else:
            # 単一セル書き込み
            if value is not None:
                final_value = value
            else:
                try:
                    final_value = get_survey_data_value(survey_data_type, survey_path)
                except Exception as e:
                    raise RuntimeError(f"writes[{i}] の survey_data '{survey_data_type}' の取得に失敗: {e}")
            ws[addr] = final_value

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
