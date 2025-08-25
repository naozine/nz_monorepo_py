# HTMLコンポーネント化計画

現在のmain.pyのHTML生成部分をコンポーネント化し、保守性・拡張性を向上させる計画書

## 1. 核となるコンポーネントクラス設計

### HTMLComponents (基底クラス)

```python
class HTMLComponents:
    def __init__(self, styles: str):
        self.styles = styles
    
    # === 基本HTML要素 ===
    def escape_html(self, text: str) -> str
    def render_page(self, title: str, sections: list) -> str
    def render_section(self, title: str, content: str, css_class: str = "") -> str
    
    # === テーブル系 ===
    def render_simple_table(self, headers: list, rows: list, css_class: str = "simple") -> str
    def render_crosstab_table(self, data: pd.DataFrame, title: str) -> str
    def render_option_count_table(...) -> str  # 既存関数をメソッド化
    def render_option_pct_table(...) -> str    # 既存関数をメソッド化
    
    # === チャート系 ===
    def render_stacked_bar(...) -> str         # 既存関数をメソッド化
    def render_group_bars(...) -> str          # 既存関数をメソッド化
    def render_legend(order: list, colors: dict) -> str
    
    # === UI要素 ===
    def render_kpi_cards(self, kpis: list) -> str
    def render_note_box(self, title: str, content: str) -> str
    def render_overview_list(self, items: dict) -> str
```

## 2. 特化コンポーネント

### QuestionComponent (設問専用)

```python
class QuestionComponent(HTMLComponents):
    def render_question_section(self, question_data: dict) -> str
    def render_question_header(self, idx: int, title: str, supplement: str, options: list) -> str
    def render_question_analysis(self, question_data: dict) -> str
```

### DemographicsComponent (属性専用)

```python
class DemographicsComponent(HTMLComponents):
    def render_demographics_section(self, demo_data: dict) -> str
    def render_gender_table(self, crosstab_data: pd.DataFrame) -> str
    def render_region_table(self, crosstab_data: pd.DataFrame) -> str
```

## 3. データフロー設計

### データ準備クラス

```python
class ReportDataPreparator:
    def prepare_overview_data(self) -> dict
    def prepare_demographics_data(self) -> dict  
    def prepare_question_data(self, question: str) -> dict
```

### データフロー

```
Raw Data → DataPreparator → Component → HTML
```

## 4. コンポーネント階層

```
HTMLComponents (基底)
├── QuestionComponent (設問専用)
├── DemographicsComponent (属性専用)  
├── ChartComponent (チャート専用)
└── TableComponent (テーブル専用)

ReportGenerator (統合)
├── HTMLComponents群を使用
└── 最終HTMLを組み立て
```

## 5. 設定管理

### 設定クラス

```python
@dataclass
class ComponentConfig:
    min_segment_width: float = 1.0
    outside_label_threshold: float = 24.0
    chart_colors: list = field(default_factory=lambda: ["#4c8bf5", "#f58b4c", ...])
    
@dataclass  
class ReportConfig:
    organizer: str = "サンプル主催者"
    survey_name: str = "サンプルイベント名"
    participating_schools: str = "参加校 100校"
    venue: str = "サンプル会場 A"
    event_dates: str = "9月1日（日）"
    # 環境変数のデフォルト値統合
```

## 6. 段階的移行計画

### Phase 1: 基底コンポーネント作成
1. `HTMLComponents`基底クラス作成
2. 既存の`render_*`関数をメソッドとして移植
   - `render_stacked_bar()` → `HTMLComponents.render_stacked_bar()`
   - `render_group_bars()` → `HTMLComponents.render_group_bars()`
   - `render_legend()` → `HTMLComponents.render_legend()`
   - `render_option_count_table()` → `HTMLComponents.render_option_count_table()`
   - `render_option_category_pct_table()` → `HTMLComponents.render_option_pct_table()`
3. `escape_html`などユーティリティ統合

### Phase 2: 専用コンポーネント分離
1. `QuestionComponent`作成・設問関連ロジック移植
   - 設問セクション全体の生成ロジック
   - 選択肢表示・補足説明の処理
   - 地域別・学年別分析の統合
2. `DemographicsComponent`作成・属性関連ロジック移植
   - 男女別クロス集計テーブル
   - 地域別クロス集計テーブル
   - 概要セクションの生成
3. 既存コードから段階的に置き換え

### Phase 3: データ分離
1. `ReportDataPreparator`作成
   - データ前処理ロジックの統合
   - HTML生成に必要な形式でのデータ準備
2. HTML生成とデータ処理を分離
3. 設定クラス(`ComponentConfig`, `ReportConfig`)導入

### Phase 4: 統合・最適化
1. `ReportGenerator`でコンポーネント統合
2. 不要なグローバル変数・関数の削除
3. テスト追加・コード整理

## 現在の関数とコンポーネントのマッピング

| 現在の関数 | 移行先コンポーネント | 備考 |
|-----------|-------------------|------|
| `escape_html()` | `HTMLComponents.escape_html()` | 基底クラス |
| `render_stacked_bar()` | `HTMLComponents.render_stacked_bar()` | チャート系 |
| `render_group_bars()` | `HTMLComponents.render_group_bars()` | チャート系 |
| `render_legend()` | `HTMLComponents.render_legend()` | チャート系 |
| `render_option_count_table()` | `HTMLComponents.render_option_count_table()` | テーブル系 |
| `render_option_category_pct_table()` | `HTMLComponents.render_option_pct_table()` | テーブル系 |
| 設問セクション生成ロジック | `QuestionComponent.render_question_section()` | 新規作成 |
| 属性セクション生成ロジック | `DemographicsComponent.render_demographics_section()` | 新規作成 |

## メリット

### 保守性
- 機能ごとに分離されたコンポーネント
- 関連する処理が同じクラスに集約
- HTMLとロジックの責任分離

### 再利用性
- 他の種類のレポートでも利用可能
- コンポーネント単位での部分利用

### テスト性
- 各コンポーネント単体でテスト可能
- モックデータでの動作確認が容易

### 拡張性
- 新しいチャートタイプの追加が簡単
- 新しいテーブル形式の追加が容易
- 設定による動作カスタマイズ

## 実装時の注意点

1. **後方互換性の維持**: 既存の出力HTMLと同じ結果を保持
2. **段階的移行**: 一度に大きく変更せず、フェーズごとに確認
3. **テスト追加**: 各フェーズでリグレッションテストを実行
4. **設定の外部化**: ハードコードされた値を設定クラスに移行
5. **エラーハンドリング**: コンポーネントレベルでの適切なエラー処理

## 実装完了ステータス

### Phase 1: 基底コンポーネント作成 ✅ **完了**
- `HTMLComponents`基底クラスの作成完了
- 既存の`render_*`関数をメソッドとして移植完了
- `escape_html`などユーティリティの統合完了

### Phase 2: 専用コンポーネント分離 ✅ **完了**
- `QuestionComponent`の作成・設問関連ロジック移植完了
- `DemographicsComponent`の作成・属性関連ロジック移植完了
- 既存コードから段階的な置き換え完了

### Phase 3: データ分離 ✅ **完了**
- `ReportDataPreparator`の作成完了
- HTML生成とデータ処理の分離完了
- 設定クラス(`ComponentConfig`, `ReportConfig`)の導入完了

### Phase 4: 統合・最適化 ✅ **完了**
- `ReportGenerator`によるコンポーネント統合完了
- 不要なグローバル変数・関数の削除完了
- エラーハンドリングとバリデーションの追加完了
- コード整理とインポートの最適化完了
- 統合システムのテストとレポート生成の確認完了

## 完了日時

**実装完了:** 2025年8月24日

## 最終確認事項

✅ HTMLレポート出力が正常に動作することを確認済み（18,923件の処理、473KB のHTMLファイル生成）
✅ 全ての機能が期待通りに動作することを確認済み
✅ エラーハンドリングと入力検証が適切に実装済み
✅ コードの保守性・拡張性が大幅に向上

## 次のステップ

1. ~~Phase 1の`HTMLComponents`基底クラスの実装から開始~~
2. ~~既存の`render_*`関数を一つずつメソッドに移行~~
3. ~~単体テストを追加して動作確認~~
4. ~~Phase 2以降に進む前に十分な検証を実施~~

**🎉 全フェーズの実装が完了しました！**