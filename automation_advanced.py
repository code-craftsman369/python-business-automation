import pandas as pd
from datetime import datetime
import glob
import os

print("🚀 複数ファイル一括処理を開始します...")
print()

# 1. すべてのCSVファイルを自動検出
csv_files = glob.glob('sales_*.csv')

if not csv_files:
    print("❌ CSVファイルが見つかりません")
    exit()

print(f"📁 検出されたファイル数: {len(csv_files)}")
for file in csv_files:
    print(f"  - {file}")
print()

# 2. 全ファイルを読み込んで結合
all_data = []

for file in csv_files:
    try:
        df = pd.read_csv(file, encoding='utf-8')
        all_data.append(df)
        print(f"✅ 読み込み成功: {file} ({len(df)}行)")
    except Exception as e:
        print(f"❌ エラー: {file} - {e}")

# 3. データを1つに結合
df_combined = pd.concat(all_data, ignore_index=True)
print()
print(f"📊 結合後のデータ: {len(df_combined)}行")
print()

# 4. 売上計算
df_combined['売上'] = df_combined['数量'] * df_combined['単価']

# 5. 多角的な集計
# 支店別集計
branch_summary = df_combined.groupby('支店').agg({
    '数量': 'sum',
    '売上': 'sum'
}).reset_index()
branch_summary = branch_summary.sort_values('売上', ascending=False)

# 商品別集計
product_summary = df_combined.groupby('商品名').agg({
    '数量': 'sum',
    '売上': 'sum'
}).reset_index()
product_summary = product_summary.sort_values('売上', ascending=False)

# 担当者別集計
staff_summary = df_combined.groupby(['支店', '担当者']).agg({
    '数量': 'sum',
    '売上': 'sum'
}).reset_index()
staff_summary = staff_summary.sort_values('売上', ascending=False)

# 日別集計
daily_summary = df_combined.groupby('日付').agg({
    '数量': 'sum',
    '売上': 'sum'
}).reset_index()

print("=== 支店別売上ランキング ===")
print(branch_summary)
print()

print("=== 商品別売上ランキング ===")
print(product_summary)
print()

# 6. 高度なExcel出力
output_file = f'consolidated_report_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'

with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    # 全データ
    df_combined.to_excel(writer, sheet_name='全データ', index=False)
    
    # 支店別
    branch_summary.to_excel(writer, sheet_name='支店別ランキング', index=False)
    
    # 商品別
    product_summary.to_excel(writer, sheet_name='商品別ランキング', index=False)
    
    # 担当者別
    staff_summary.to_excel(writer, sheet_name='担当者別詳細', index=False)
    
    # 日別
    daily_summary.to_excel(writer, sheet_name='日別推移', index=False)

print(f"✅ 統合レポート出力完了: {output_file}")
print()
print("📈 生成されたシート:")
print("  1. 全データ（結合後）")
print("  2. 支店別ランキング")
print("  3. 商品別ランキング")
print("  4. 担当者別詳細")
print("  5. 日別推移")