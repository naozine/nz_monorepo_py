from pathlib import Path
from typing import Union

try:
    # openpyxl is the de-facto library for .xlsx read/write
    from openpyxl import load_workbook
except Exception as e:  # pragma: no cover
    raise RuntimeError("openpyxl が必要です。`pip install openpyxl` を実行してください。") from e


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


def main():
    """簡易CLI: テンプレート→結果ファイルに AC14 を 'xxx' で出力"""
    base = Path(__file__).parent
    template = base / "report_template.xlsx"
    output = base / "report_result.xlsx"

    out_path = fill_ac14(template, output, value="xxx")
    print(f"書き出し完了: {out_path}")


if __name__ == "__main__":
    main()
