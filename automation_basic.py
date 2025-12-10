import pandas as pd
from datetime import datetime

# 1. CSVファイル読み込み
df = pd.read_csv('sales_data.csv', encoding='utf-8')

# 売上金額を計算
df['売上'] = df['数量'] * df['単価']

print("=== 元データ ===")
print(df)
print()

# 2. 基本的なデータ集計
# 商品別の売上合計
product_summary = df.groupby('商品名').agg({
    '数量': 'sum',
    '売上': 'sum'
}).reset_index()

print("=== 商品別集計 ===")
print(product_summary)
print()

# 担当者別の売上合計
staff_summary = df.groupby('担当者').agg({
    '数量': 'sum',
    '売上': 'sum'
}).reset_index()

print("=== 担当者別集計 ===")
print(staff_summary)
print()

# 3. Excel出力
# 複数シートに分けて出力
output_file = f'sales_report_{datetime.now().strftime("%Y%m%d")}.xlsx'

with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    df.to_excel(writer, sheet_name='詳細データ', index=False)
    product_summary.to_excel(writer, sheet_name='商品別集計', index=False)
    staff_summary.to_excel(writer, sheet_name='担当者別集計', index=False)

print(f"✅ レポート出力完了: {output_file}")


