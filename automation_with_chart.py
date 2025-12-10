import pandas as pd
from datetime import datetime
import glob
import matplotlib.pyplot as plt
import matplotlib
matplotlib.use('Agg')  # GUIなしでグラフ生成

# 日本語フォント設定
plt.rcParams['font.sans-serif'] = ['Arial Unicode MS', 'Hiragino Sans']
plt.rcParams['axes.unicode_minus'] = False

print("🚀 複数ファイル一括処理＋グラフ生成を開始します...")
print()

# 1. すべてのCSVファイルを自動検出
csv_files = glob.glob('sample_data/sales_*.csv')

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

# 5. 集計
# 支店別
if '支店' in df_combined.columns:
    branch_summary = df_combined.groupby('支店').agg({
        '数量': 'sum',
        '売上': 'sum'
    }).reset_index()
    branch_summary = branch_summary.sort_values('売上', ascending=False)
else:
    branch_summary = None

# 商品別
product_summary = df_combined.groupby('商品名').agg({
    '数量': 'sum',
    '売上': 'sum'
}).reset_index()
product_summary = product_summary.sort_values('売上', ascending=False)

# 日別
daily_summary = df_combined.groupby('日付').agg({
    '数量': 'sum',
    '売上': 'sum'
}).reset_index()

print("=== 商品別売上ランキング ===")
print(product_summary)
print()

# 6. グラフ生成
print("📈 グラフを生成中...")

# 4つのグラフを作成
fig, axes = plt.subplots(2, 2, figsize=(14, 10))
fig.suptitle('Sales Analysis Dashboard', fontsize=16, fontweight='bold')

# グラフ1: 支店別売上（棒グラフ）
if branch_summary is not None:
    axes[0, 0].bar(branch_summary['支店'], branch_summary['売上'], color='skyblue')
    axes[0, 0].set_title('Branch Sales Ranking', fontweight='bold')
    axes[0, 0].set_xlabel('Branch')
    axes[0, 0].set_ylabel('Sales (JPY)')
    axes[0, 0].grid(axis='y', alpha=0.3)
    
    # 値を表示
    for i, v in enumerate(branch_summary['売上']):
        axes[0, 0].text(i, v + 1000, f'¥{int(v):,}', ha='center', va='bottom')

# グラフ2: 商品別売上（横棒グラフ）
axes[0, 1].barh(product_summary['商品名'], product_summary['売上'], color='lightcoral')
axes[0, 1].set_title('Product Sales Ranking', fontweight='bold')
axes[0, 1].set_xlabel('Sales (JPY)')
axes[0, 1].set_ylabel('Product')
axes[0, 1].grid(axis='x', alpha=0.3)

# 値を表示
for i, v in enumerate(product_summary['売上']):
    axes[0, 1].text(v + 1000, i, f'¥{int(v):,}', ha='left', va='center')

# グラフ3: 日別売上推移（折れ線グラフ）
axes[1, 0].plot(daily_summary['日付'], daily_summary['売上'], 
                marker='o', linewidth=2, markersize=8, color='green')
axes[1, 0].set_title('Daily Sales Trend', fontweight='bold')
axes[1, 0].set_xlabel('Date')
axes[1, 0].set_ylabel('Sales (JPY)')
axes[1, 0].grid(alpha=0.3)
axes[1, 0].tick_params(axis='x', rotation=45)

# グラフ4: 商品別売上割合（円グラフ）
colors = ['gold', 'lightblue', 'lightgreen', 'pink', 'orange']
axes[1, 1].pie(product_summary['売上'], labels=product_summary['商品名'], 
               autopct='%1.1f%%', startangle=90, colors=colors)
axes[1, 1].set_title('Product Sales Share', fontweight='bold')

plt.tight_layout()

# グラフを画像として保存
chart_file = f'sales_chart_{datetime.now().strftime("%Y%m%d_%H%M%S")}.png'
plt.savefig(chart_file, dpi=150, bbox_inches='tight')
print(f"✅ グラフ保存完了: {chart_file}")

# 7. Excel出力（グラフは別途画像として保存済み）
output_file = f'report_with_chart_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'

with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    # 全データ
    df_combined.to_excel(writer, sheet_name='全データ', index=False)
    
    # 支店別
    if branch_summary is not None:
        branch_summary.to_excel(writer, sheet_name='支店別ランキング', index=False)
    
    # 商品別
    product_summary.to_excel(writer, sheet_name='商品別ランキング', index=False)
    
    # 日別
    daily_summary.to_excel(writer, sheet_name='日別推移', index=False)

print(f"✅ Excelレポート出力完了: {output_file}")
print()
print("📊 生成されたファイル:")
print(f"  1. {output_file} (Excelレポート)")
print(f"  2. {chart_file} (グラフ画像)")
print()
print("💡 グラフ画像をExcelレポートに手動で挿入することも可能です")