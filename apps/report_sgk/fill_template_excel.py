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


def fill_from_yaml(config_path: Union[str, Path]) -> Path:
    """
    YAML 設定を読み込み、指定のシート/セルへ値を書き込みます。

    想定する YAML 例:

    template: report_template.xlsx
    output: report_result.xlsx
    writes:
      - sheet: p1
        cell: AC14
        value: xxx
      - sheet: p2
        row: 5
        column: B   # または 2 （数値でも可）
        value: Hello

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

        value = w.get("value")
        ws = wb[sheet]
        ws[addr] = value

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
