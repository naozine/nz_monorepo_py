# サマリレポート生成ツール

このツールは、アンケートデータからHTMLレポートを生成します。

## 機能

- アンケートデータの読み込みと前処理
- 回答者属性の集計（性別、地域別、学年別）
- 設問ごとの詳細集計と可視化
- 積み上げ棒グラフによる視覚的表示
- A4印刷対応のHTMLレポート出力

## 使用方法

```bash
python main.py
```

実行すると `report.html` が生成されます。

### 先頭ページの文言を .env で設定する

レポート先頭ページの以下の項目は、環境変数でパラメータ化されています。プロジェクトのルート（または実行ディレクトリ）に `.env` を置くと自動で読み込まれます。

- REPORT_ORGANIZER （例: サンプル主催者）
- REPORT_SURVEY_NAME （例: サンプルイベント名）
- REPORT_PARTICIPATING_SCHOOLS （例: 参加校 100校）
- REPORT_VENUE （例: サンプル会場 A）
- REPORT_EVENT_DATES （例: 9月1日（日））

例: `.env`

```
REPORT_ORGANIZER=サンプル主催者
REPORT_SURVEY_NAME=サンプルイベント名
REPORT_PARTICIPATING_SCHOOLS=参加校 100校
REPORT_VENUE=サンプル会場 A
REPORT_EVENT_DATES=9月1日（日）
```

.env が存在しない場合や項目が未設定の場合は、既定値（現行ハードコード値）で出力されます。

## グラフ表示の設定変数

### セグメント最小幅保証

```python
MIN_SEGMENT_WIDTH_PCT = 1.0  # セグメントの最小幅（％）
```

**説明**: 積み上げ棒グラフで、非常に小さい割合のセグメントでも最低限の幅を保証します。

- **用途**: 0.5%などの極小セグメントを視認可能にする
- **効果**: 指定値未満のセグメントは、この値まで幅が拡張される
- **調整**: より大きくすると小さいセグメントが見やすくなるが、他のセグメントが圧縮される

### 外側ラベル表示閾値

```python
OUTSIDE_LABEL_THRESHOLD_PCT = 30.0  # 外側ラベル表示閾値（％）
```

**説明**: セグメント内にラベルが収まらない場合の外側表示判定基準です。

- **用途**: 小さいセグメントのラベル文字が見切れる問題を解決
- **効果**: この値未満の割合のセグメントは、ラベルが棒の外側（上下）に表示される
- **配置ルール**: 
  - 1つ目: グラフ直上
  - 2つ目: グラフ直下  
  - 3つ目以降: 上下交互に層を追加（リード線付き）
- **調整**: 値を下げると外側ラベルが増え、上げると内側ラベルが増える

### 外側ラベル内容分割閾値

```python
OUTSIDE_LABEL_WITH_INNER_PCT_THRESHOLD = 10.0  # 外側ラベル+内側割合表示の閾値（％）
```

**説明**: 外側ラベル対象のセグメントで、表示内容を分割する基準値です。

- **用途**: 外側ラベルが長くなりすぎる問題を解決し、棒内部のスペースを有効活用
- **効果**: 
  - **10%以上**: 外側に選択肢名のみ表示、棒内部に割合（%）を表示
  - **10%未満**: 外側に選択肢名+割合の両方を表示、棒内部は空
- **メリット**: 中程度の大きさのセグメントで外側ラベルがすっきりし、視認性が向上
- **調整**: 値を下げると内側割合表示が増え、上げると外側完全表示が増える

## ファイル構成

- `main.py` - メインスクリプト
- `survey.xlsx` - アンケートデータ（Excel形式）
- `report.html` - 生成されるHTMLレポート

## 仕様

- 未就学児（6歳未満）は集計から除外
- 複数回答可の設問に対応
- 地域区分: 東京23区、三多摩島しょ、埼玉県、神奈川県、千葉県、その他
- 学年: 小1〜中3
- 印刷時のカラー保持対応
# report_sgk

このアプリはアンケート Excel (survey.xlsx) を集計し、HTML レポートとテンプレート Excel への自動書き込みを行います。

## YAML によるテンプレート Excel への書き込み

apps/report_sgk/fill_template_excel.py の `fill_from_yaml` を利用して、テンプレート Excel に値を書き込めます。

### 例

```
template: report_template.xlsx
output: report_result_from_yaml.xlsx
survey_excel: survey.xlsx
writes:
  - sheet: Q1
    cell: T10
    survey_series: responses
    question: 1          # main.py の設問一覧順（1開始）
    choice: 知っている   # 指定設問の選択肢文字列
    class: grade         # grade | region
    # direction: down    # 任意（responses は down がデフォルト）
```

- survey_series: responses
  - 指定の設問(question)における指定の選択肢(choice)の「回答数」を、
    - class: grade の場合は [小1..中3] の順
    - class: region の場合は [東京23区, 三多摩島しょ, 埼玉県, 神奈川県, 千葉県, その他] の順
    で配列化し、セルから縦方向（デフォルト）に順に書き込みます。
- 設問番号は main.py の `get_question_columns()` が返すリストの 1 始まりの番号です。

既存の series も利用できます:
- elementary_boys / elementary_girls / junior_boys / junior_girls
- region_grades（`region:` で地域を指定、または `region_grades:東京23区` の形式）
