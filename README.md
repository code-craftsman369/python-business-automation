# Python業務自動化ツール集

## 📋 概要
Pythonを使った業務データの自動集計・レポート生成ツールです。
CSV形式の売上データを自動で処理し、Excel形式のレポートを出力します。

## ✨ 機能

### 基本版（automation_basic.py）
- CSV読み込み
- 商品別・担当者別の売上集計
- Excel形式での出力（複数シート対応）

### 拡張版（automation_advanced.py）
- **複数CSVファイルの一括処理**
- 自動ファイル検出
- データ結合・統合
- 多角的集計（支店別・商品別・担当者別・日別）
- ランキング機能
- エラーハンドリング

## 🛠️ 使用技術
- Python 3.13
- pandas（データ処理）
- openpyxl（Excel出力）

## 🚀 使い方

### 1. 環境構築
```bash
pip install pandas openpyxl
```

### 2. 基本版の実行
```bash
python3 automation_basic.py
```

### 3. 拡張版の実行（複数ファイル処理）
```bash
# sample_dataフォルダ内のすべてのsales_*.csvを自動処理
python3 automation_advanced.py
```

## 📊 出力例

### 基本版
- `sales_report_YYYYMMDD.xlsx`
  - シート1: 詳細データ
  - シート2: 商品別集計
  - シート3: 担当者別集計

### 拡張版
- `consolidated_report_YYYYMMDD_HHMMSS.xlsx`
  - シート1: 全データ（結合後）
  - シート2: 支店別ランキング
  - シート3: 商品別ランキング
  - シート4: 担当者別詳細
  - シート5: 日別推移

## 💼 実務での活用例
- 月次売上レポート自動生成
- 複数支店のデータ統合
- 日次・週次の定型レポート作成
- データ入力作業の自動化

## 📝 対応可能な案件例
- CSV/Excelデータの加工・集計
- 定型レポートの自動生成
- 複数ファイルの一括処理
- 業務効率化ツールの開発

## 📜 License
MIT

## 👤 Author
**Tatsu**

GitHub: [@code-craftsman369](https://github.com/code-craftsman369)  
X: [@web3_builder369](https://twitter.com/web3_builder369)

## 🙏 Acknowledgments
- pandas community for excellent data processing library
- openpyxl developers for Excel file handling