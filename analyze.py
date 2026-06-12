"""
蒸汽平台（Steam）游戏市场数据分析
==================================
课程：数据思维与人工智能 大作业

分析维度：
  1. 综合评价分析 — 价格/评分/热度分布、相关性矩阵
  2. 游戏定价策略 — 免费vs付费、价格区间与销量关系
  3. 游戏类型趋势 — 类型分布、类型评分、年度变化
  4. 玩家评价分析 — 好评率分布、影响因素探究

输出：控制台报告 + PNG图表 + 分析报告文本
"""

import pandas as pd
import numpy as np
import matplotlib
matplotlib.use('Agg')  # 非交互式后端，适合脚本运行
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
from matplotlib.gridspec import GridSpec
import seaborn as sns
import warnings
from datetime import datetime

warnings.filterwarnings('ignore')

# ============================================================
# 全局设置
# ============================================================
plt.rcParams['font.family'] = ['SimHei', 'DejaVu Sans', 'sans-serif']
plt.rcParams['axes.unicode_minus'] = False
plt.rcParams['figure.dpi'] = 120
plt.rcParams['savefig.dpi'] = 150
plt.rcParams['savefig.bbox'] = 'tight'
sns.set_style("whitegrid")

# 中文字体回退方案（使用 font_manager 实际检测字体是否可用）
import matplotlib.font_manager as fm
font_found = False
for font_name in ['SimHei', 'Microsoft YaHei', 'WenQuanYi Micro Hei', 'Noto Sans CJK SC']:
    try:
        fm.findfont(font_name, fallback_to_default=False)
        plt.rcParams['font.family'] = [font_name, 'sans-serif']
        print(f'使用字体: {font_name}')
        font_found = True
        break
    except (ValueError, RuntimeError):
        continue
if not font_found:
    print('警告: 未找到中文字体，图表中文可能无法正常显示')

print("库导入完成")


# ============================================================
# 1. 数据加载与初步探索
# ============================================================
print("\n" + "=" * 70)
print("  第一阶段：数据加载与探索")
print("=" * 70)

df = pd.read_csv('Steam Games 2024.csv', low_memory=False)

print(f"\n数据集规模: {df.shape[0]:,} 行 × {df.shape[1]} 列")

# 列出所有列名方便参考
print(f"\n全部列名 ({len(df.columns)}个):")
for i, col in enumerate(df.columns):
    print(f"  [{i:2d}] {col}")

# 缺失值概览
print(f"\n关键列缺失值情况:")
key_cols = ['Name', 'Release date', 'Price', 'Estimated owners', 'Peak CCU',
            'Positive', 'Negative', 'Genres', 'Categories', 'Tags',
            'Metacritic score', 'User score', 'Achievements', 'Recommendations',
            'Average playtime forever', 'DLC count']
for col in key_cols:
    if col in df.columns:
        miss = df[col].isnull().sum()
        pct = miss / len(df) * 100
        print(f"  {col:30s}: {miss:>6,} 缺失 ({pct:.1f}%)")
    else:
        print(f"  {col:30s}: 列不存在！")

# ============================================================
# 2. 数据清洗与特征工程
# ============================================================
print("\n" + "=" * 70)
print("  第二阶段：数据清洗与特征工程")
print("=" * 70)

# 2.1 解析 Estimated owners（字符串范围 → 数值中位数）
def parse_owners(val):
    """ '0 - 20000' → 10000, '50000000 - 100000000' → 75000000 """
    if pd.isna(val):
        return np.nan
    parts = str(val).split(' - ')
    if len(parts) == 2:
        try:
            low = int(parts[0].replace(',', '').strip())
            high = int(parts[1].replace(',', '').strip())
            return (low + high) / 2
        except (ValueError, TypeError):
            return np.nan
    return np.nan

df['owners_num'] = df['Estimated owners'].apply(parse_owners)
print(f"  拥有者数解析完毕: 有效 {df['owners_num'].notna().sum():,} / {len(df):,}")

# 2.2 价格处理
df['Price'] = pd.to_numeric(df['Price'], errors='coerce')
df['is_free'] = df['Price'] <= 0
n_free = df['is_free'].sum()
print(f"  免费游戏: {n_free:,} ({n_free/len(df)*100:.1f}%)")

# 2.3 发行日期解析
def parse_date(val):
    if pd.isna(val):
        return pd.NaT
    try:
        return pd.to_datetime(val)
    except (ValueError, TypeError):
        return pd.NaT

df['release_datetime'] = df['Release date'].apply(parse_date)
df['release_year'] = df['release_datetime'].dt.year
valid_dates = df['release_year'].notna().sum()
print(f"  有效发行日期: {valid_dates:,} / {len(df):,}")

# 2.4 好评率
df['positive'] = pd.to_numeric(df['Positive'], errors='coerce').fillna(0)
df['negative'] = pd.to_numeric(df['Negative'], errors='coerce').fillna(0)
df['total_reviews'] = df['positive'] + df['negative']
df['positive_rate'] = np.where(
    df['total_reviews'] > 0,
    df['positive'] / df['total_reviews'] * 100,
    np.nan
)
n_with_reviews = (df['total_reviews'] > 0).sum()
print(f"  有评价的游戏: {n_with_reviews:,} ({n_with_reviews/len(df)*100:.1f}%)")

# 2.5 峰值在线人数
df['Peak CCU'] = pd.to_numeric(df['Peak CCU'], errors='coerce')

# 2.6 支持平台（安全转换：处理字符串 'TRUE'/'FALSE' 及缺失值）
def _to_bool(val):
    """将字符串 'TRUE'/'FALSE' 正确转为布尔，缺失/异常值视为 False"""
    s = str(val).strip().upper() if not pd.isna(val) else ''
    return s == 'TRUE'

for col in ['Windows', 'Mac', 'Linux']:
    df[col] = df[col].apply(_to_bool)
df['platform_count'] = df[['Windows', 'Mac', 'Linux']].sum(axis=1)

# 2.7 折扣标记
df['Discount'] = pd.to_numeric(df['Discount'], errors='coerce').fillna(0)
df['has_discount'] = df['Discount'] > 0

# 2.8 游戏内容量指标
df['DLC count'] = pd.to_numeric(df['DLC count'], errors='coerce').fillna(0)
df['Achievements'] = pd.to_numeric(df['Achievements'], errors='coerce').fillna(0)

# 2.9 Metacritic 评分
df['Metacritic score'] = pd.to_numeric(df['Metacritic score'], errors='coerce')
df['User score'] = pd.to_numeric(df['User score'], errors='coerce')

# 2.10 价格区间分类
def price_category(p):
    if pd.isna(p):
        return '未知'
    if p <= 0:
        return '免费'
    elif p < 5:
        return '低价(0-5)'
    elif p < 15:
        return '中低价(5-15)'
    elif p < 30:
        return '中等(15-30)'
    elif p < 60:
        return '中高价(30-60)'
    else:
        return '高价(60+)'

df['price_category'] = df['Price'].apply(price_category)

# 2.11 拥有者量级分类
def owners_category(v):
    if pd.isna(v):
        return '未知'
    if v < 50000:
        return '小型(<5万)'
    elif v < 200000:
        return '中小(5-20万)'
    elif v < 1000000:
        return '中型(20-100万)'
    elif v < 5000000:
        return '中大型(100-500万)'
    else:
        return '大型(500万+)'

df['owners_category'] = df['owners_num'].apply(owners_category)

print(f"\n清洗完成！数据集现包含 {len(df.columns)} 列")
print(f"新增特征: owners_num, is_free, release_year, positive_rate, total_reviews,")
print(f"          platform_count, price_category, owners_category, has_discount")

# 数据清洗摘要
print(f"\n数据清洗摘要:")
print(f"  总游戏数:       {len(df):,}")
print(f"  有效价格:       {df['Price'].notna().sum():,}")
print(f"  有效发行年份:   {df['release_year'].notna().sum():,}")
print(f"  有效拥有者数据: {df['owners_num'].notna().sum():,}")
print(f"  有评价的游戏:   {n_with_reviews:,}")
print(f"  平均好评率:     {df['positive_rate'].mean():.1f}%" if n_with_reviews > 0 else "")


# ============================================================
# 3. 维度一：综合评价分析
# ============================================================
print("\n" + "=" * 70)
print("  维度一：综合评价分析")
print("=" * 70)

# --- 图1: 综合仪表盘（4子图） ---
fig, axes = plt.subplots(2, 2, figsize=(14, 10))
fig.suptitle('Steam 游戏市场综合概览', fontsize=16, fontweight='bold')

# 子图1: 价格分布
ax1 = axes[0, 0]
paid = df[df['is_free'] == False]['Price']
paid_under_100 = paid[paid <= 100]
ax1.hist(paid_under_100, bins=80, color='steelblue', edgecolor='white', alpha=0.8)
ax1.axvline(paid.median(), color='red', linestyle='--', linewidth=2, label=f'中位数: ${paid.median():.1f}')
ax1.set_title('付费游戏价格分布 (≤$100)', fontsize=12)
ax1.set_xlabel('价格 (USD)')
ax1.set_ylabel('游戏数量')
ax1.legend()
print(f"  价格统计: 均值=${paid.mean():.2f}, 中位数=${paid.median():.2f}, 最大=${paid.max():.2f}")

# 子图2: 好评率分布
ax2 = axes[0, 1]
valid_reviews = df[df['total_reviews'] >= 50]['positive_rate'].dropna()
ax2.hist(valid_reviews, bins=50, color='forestgreen', edgecolor='white', alpha=0.8)
ax2.axvline(valid_reviews.median(), color='red', linestyle='--', linewidth=2,
            label=f'中位数: {valid_reviews.median():.1f}%')
ax2.set_title('游戏好评率分布 (评价≥50)', fontsize=12)
ax2.set_xlabel('好评率 (%)')
ax2.set_ylabel('游戏数量')
ax2.legend()
print(f"  好评率统计: 均值={valid_reviews.mean():.1f}%, 中位数={valid_reviews.median():.1f}%")

# 子图3: 拥有者数分布（对数）
ax3 = axes[1, 0]
owners_valid = df['owners_num'].dropna()
log_owners = np.log10(owners_valid[owners_valid > 0])
ax3.hist(log_owners, bins=50, color='darkorange', edgecolor='white', alpha=0.8)
ax3.set_title('游戏拥有者数量分布（对数尺度）', fontsize=12)
ax3.set_xlabel('log10(拥有者数)')
ax3.set_ylabel('游戏数量')
xticks = [3, 4, 5, 6, 7, 8]
ax3.set_xticks(xticks)
ax3.set_xticklabels(['1千', '1万', '10万', '100万', '1000万', '1亿'])
print(f"  拥有者统计: 中位数={owners_valid.median():,.0f}, 均值={owners_valid.mean():,.0f}")

# 子图4: 发行年份趋势
ax4 = axes[1, 1]
year_counts = df['release_year'].value_counts().sort_index()
year_counts = year_counts[(year_counts.index >= 2006) & (year_counts.index <= 2025)]
ax4.fill_between(year_counts.index, year_counts.values, alpha=0.5, color='purple')
ax4.plot(year_counts.index, year_counts.values, color='purple', linewidth=2)
ax4.set_title('Steam 游戏发行数量年度趋势', fontsize=12)
ax4.set_xlabel('发行年份')
ax4.set_ylabel('游戏数量')
print(f"  年份范围: {int(df['release_year'].min())} - {int(df['release_year'].max())}")

plt.tight_layout()
plt.savefig('01_综合概览.png', dpi=150)
plt.close()
print("  → 图表已保存: 01_综合概览.png")

# --- 相关性分析 ---
print(f"\n  相关性分析（关键数值变量）:")
corr_cols = ['Price', 'owners_num', 'Peak CCU', 'positive_rate', 'total_reviews',
             'Metacritic score', 'DLC count', 'Achievements', 'platform_count']
corr_data = df[corr_cols].copy()
corr_matrix = corr_data.corr()

# 打印关键相关系数
key_pairs = [
    ('Price', 'positive_rate'),
    ('Price', 'owners_num'),
    ('owners_num', 'positive_rate'),
    ('owners_num', 'Peak CCU'),
    ('Metacritic score', 'positive_rate'),
    ('total_reviews', 'positive_rate'),
]
for a, b in key_pairs:
    if a in corr_matrix.columns and b in corr_matrix.index:
        r = corr_matrix.loc[a, b]
        print(f"    {a} ↔ {b}: r = {r:.3f}")

# 相关性热力图
fig, ax = plt.subplots(figsize=(10, 8))
mask = np.triu(np.ones_like(corr_matrix, dtype=bool), k=1)
sns.heatmap(corr_matrix, mask=mask, annot=True, fmt='.2f', cmap='RdBu_r',
            center=0, square=True, linewidths=0.5,
            cbar_kws={'shrink': 0.8, 'label': '相关系数'},
            ax=ax)
ax.set_title('Steam 游戏关键指标相关性矩阵', fontsize=14, fontweight='bold')
plt.tight_layout()
plt.savefig('01_相关性热力图.png', dpi=150)
plt.close()
print("  → 图表已保存: 01_相关性热力图.png")

# Top游戏列表
print(f"\n  好评数最多的10款游戏:")
top_positive = df.nlargest(10, 'positive')[['Name', 'positive', 'negative', 'positive_rate', 'owners_num', 'Price']]
for _, row in top_positive.iterrows():
    print(f"    {row['Name'][:45]:45s} | 好评:{row['positive']:>10,.0f} | "
          f"好评率:{row['positive_rate']:.1f}% | 拥有者:{row['owners_num']:>12,.0f} | ${row['Price']:.2f}")


# ============================================================
# 4. 维度二：游戏定价策略分析
# ============================================================
print("\n" + "=" * 70)
print("  维度二：游戏定价策略分析")
print("=" * 70)

# --- 图2: 定价策略 ---
fig, axes = plt.subplots(2, 2, figsize=(14, 10))
fig.suptitle('Steam 游戏定价策略分析', fontsize=16, fontweight='bold')

# 子图1: 价格区间分布
ax1 = axes[0, 0]
price_order = ['免费', '低价(0-5)', '中低价(5-15)', '中等(15-30)', '中高价(30-60)', '高价(60+)']
price_counts = df['price_category'].value_counts()
price_counts = price_counts.reindex([p for p in price_order if p in price_counts.index])
colors1 = ['#2ecc71', '#3498db', '#9b59b6', '#e67e22', '#e74c3c', '#c0392b']
ax1.bar(range(len(price_counts)), price_counts.values, color=colors1[:len(price_counts)])
ax1.set_xticks(range(len(price_counts)))
ax1.set_xticklabels(price_counts.index, rotation=30, ha='right', fontsize=9)
ax1.set_title('各价格区间游戏数量', fontsize=12)
ax1.set_ylabel('游戏数量')
for i, v in enumerate(price_counts.values):
    ax1.text(i, v + max(price_counts.values)*0.01, f'{v:,}', ha='center', fontsize=8)
print("  价格区间分布:")
for cat in price_counts.index:
    print(f"    {cat:20s}: {price_counts[cat]:>8,} 款")

# 子图2: 价格区间 vs 好评率
ax2 = axes[0, 1]
valid_reviews = df[df['total_reviews'] >= 50]['positive_rate'].dropna()
price_rate = df[df['total_reviews'] >= 50].groupby('price_category')['positive_rate'].median()
price_rate = price_rate.reindex([p for p in price_order if p in price_rate.index])
ax2.bar(range(len(price_rate)), price_rate.values, color=colors1[:len(price_rate)])
ax2.set_xticks(range(len(price_rate)))
ax2.set_xticklabels(price_rate.index, rotation=30, ha='right', fontsize=9)
ax2.set_title('各价格区间中位数好评率', fontsize=12)
ax2.set_ylabel('中位数好评率 (%)')
ax2.axhline(y=valid_reviews.median(), color='red', linestyle='--', alpha=0.7, label=f'整体中位数: {valid_reviews.median():.1f}%')
ax2.legend(fontsize=8)
for i, v in enumerate(price_rate.values):
    ax2.text(i, v + 1, f'{v:.1f}%', ha='center', fontsize=8)
print("\n  各价格区间中位数好评率:")
for cat in price_rate.index:
    print(f"    {cat:20s}: {price_rate[cat]:.1f}%")

# 子图3: 免费 vs 付费
ax3 = axes[1, 0]
free_paid = df[df['total_reviews'] >= 50].copy()
x = np.arange(2)
width = 0.3
free_rate = free_paid[free_paid['is_free']]['positive_rate'].median()
paid_rate = free_paid[~free_paid['is_free']]['positive_rate'].median()
free_owners = free_paid[free_paid['is_free']]['owners_num'].median() / 10000
paid_owners = free_paid[~free_paid['is_free']]['owners_num'].median() / 10000
ax3.bar(x - width/2, [free_rate, paid_rate], width, label='中位好评率(%)', color='#3498db')
ax3_twin = ax3.twinx()
ax3_twin.bar(x + width/2, [free_owners, paid_owners], width, label='中位拥有者(万)', color='#e74c3c')
ax3.set_xticks(x)
ax3.set_xticklabels(['免费游戏', '付费游戏'], fontsize=11)
ax3.set_title('免费 vs 付费 关键指标对比', fontsize=12)
ax3.set_ylabel('好评率 (%)', color='#3498db')
ax3_twin.set_ylabel('中位拥有者 (万)', color='#e74c3c')
lines1, labels1 = ax3.get_legend_handles_labels()
lines2, labels2 = ax3_twin.get_legend_handles_labels()
ax3.legend(lines1 + lines2, labels1 + labels2, loc='upper right', fontsize=8)
free_g = free_paid[free_paid['is_free']]
paid_g = free_paid[~free_paid['is_free']]
print(f"\n  免费 vs 付费对比:")
print(f"    免费游戏: {len(free_g):,}款, 中位好评率={free_g['positive_rate'].median():.1f}%, 中位拥有者={free_g['owners_num'].median():,.0f}")
print(f"    付费游戏: {len(paid_g):,}款, 中位好评率={paid_g['positive_rate'].median():.1f}%, 中位拥有者={paid_g['owners_num'].median():,.0f}")

# 子图4: 价格 vs 拥有者
ax4 = axes[1, 1]
scatter_data = df[(df['Price'] > 0) & (df['Price'] <= 70) & (df['owners_num'] > 0)].copy()
if len(scatter_data) > 5000:
    scatter_data = scatter_data.sample(5000, random_state=42)
sc = ax4.scatter(scatter_data['Price'], np.log10(scatter_data['owners_num']),
                 c=scatter_data['positive_rate'], cmap='RdYlGn', alpha=0.5, s=10,
                 vmin=50, vmax=100)
ax4.set_xlabel('价格 (USD)')
ax4.set_ylabel('log10(拥有者数)')
ax4.set_title('价格 vs 拥有者数（颜色=好评率）', fontsize=12)
cbar = plt.colorbar(sc, ax=ax4, shrink=0.8)
cbar.set_label('好评率 (%)')

plt.tight_layout()
plt.savefig('02_定价策略分析.png', dpi=150)
plt.close()
print("  -> 图表已保存: 02_定价策略分析.png")


# ============================================================
# 5. 维度三：游戏类型趋势分析
# ============================================================
print("\n" + "=" * 70)
print("  维度三：游戏类型趋势分析")
print("=" * 70)

# 展开 Genres 列
genre_rows = []
for _, row in df.iterrows():
    genres_str = row['Genres']
    if pd.isna(genres_str) or str(genres_str).strip() == '':
        continue
    for g in str(genres_str).split(','):
        g = g.strip()
        if g:
            genre_rows.append({
                'AppID': row['AppID'],
                'genre': g,
                'positive_rate': row['positive_rate'],
                'owners_num': row['owners_num'],
                'Price': row['Price'],
                'release_year': row['release_year'],
            })
genre_df = pd.DataFrame(genre_rows)
print(f"  展开后类型条目数: {len(genre_df):,}")

# --- 图3: 类型分析 ---
fig, axes = plt.subplots(2, 2, figsize=(14, 10))
fig.suptitle('Steam 游戏类型分析', fontsize=16, fontweight='bold')

# 子图1: 类型数量
ax1 = axes[0, 0]
genre_counts = genre_df['genre'].value_counts().head(15)
colors_genre = plt.cm.viridis(np.linspace(0.1, 0.9, len(genre_counts)))
ax1.barh(range(len(genre_counts)), genre_counts.values, color=colors_genre)
ax1.set_yticks(range(len(genre_counts)))
ax1.set_yticklabels(genre_counts.index, fontsize=9)
ax1.set_title('游戏数量 Top 15 类型', fontsize=12)
ax1.set_xlabel('游戏款数')
ax1.invert_yaxis()
for i, v in enumerate(genre_counts.values):
    ax1.text(v + max(genre_counts.values)*0.01, i, f'{v:,}', va='center', fontsize=8)
print("\n  Top 15 游戏类型:")
for g, c in genre_counts.items():
    print(f"    {g:20s}: {c:>8,} 款")

# 子图2: 类型好评率
ax2 = axes[0, 1]
genre_rate = genre_df[genre_df['positive_rate'].notna()].groupby('genre')['positive_rate'].median()
top_genres = genre_counts.head(15).index
genre_rate_top = genre_rate.reindex(top_genres).sort_values()
ax2.barh(range(len(genre_rate_top)), genre_rate_top.values,
         color=colors_genre[np.argsort(genre_counts.values)][:len(genre_rate_top)])
ax2.set_yticks(range(len(genre_rate_top)))
ax2.set_yticklabels(genre_rate_top.index, fontsize=9)
ax2.set_title('Top 15 类型中位数好评率', fontsize=12)
ax2.set_xlabel('中位数好评率 (%)')
ax2.invert_yaxis()
print("\n  各类型中位数好评率:")
for g in genre_rate_top.index:
    print(f"    {g:20s}: {genre_rate_top[g]:.1f}%")

# 子图3: 类型趋势
ax3 = axes[1, 0]
top6_genres = genre_counts.head(6).index.tolist()
yearly_genre = genre_df[genre_df['genre'].isin(top6_genres)].copy()
yearly_genre = yearly_genre[(yearly_genre['release_year'] >= 2010) & (yearly_genre['release_year'] <= 2025)]
yearly_pivot = yearly_genre.groupby(['release_year', 'genre']).size().unstack(fill_value=0)
for genre in top6_genres:
    if genre in yearly_pivot.columns:
        ax3.plot(yearly_pivot.index, yearly_pivot[genre], linewidth=2,
                 label=genre, marker='.', markersize=4)
ax3.set_title('Top 6 类型年度发行趋势', fontsize=12)
ax3.set_xlabel('年份')
ax3.set_ylabel('游戏数量')
ax3.legend(fontsize=8, loc='upper left')

# 子图4: 类型平均价格
ax4 = axes[1, 1]
genre_price = genre_df[genre_df['Price'] > 0].groupby('genre')['Price'].mean()
genre_price_top = genre_price.reindex(top_genres).sort_values()
ax4.barh(range(len(genre_price_top)), genre_price_top.values,
         color=colors_genre[np.argsort(genre_counts.values)][:len(genre_price_top)])
ax4.set_yticks(range(len(genre_price_top)))
ax4.set_yticklabels(genre_price_top.index, fontsize=9)
ax4.set_title('Top 15 类型平均价格', fontsize=12)
ax4.set_xlabel('平均价格 (USD)')
ax4.invert_yaxis()
print("\n  各类型平均价格:")
for g in genre_price_top.index:
    print(f"    {g:20s}: ${genre_price_top[g]:.2f}")

plt.tight_layout()
plt.savefig('03_类型分析.png', dpi=150)
plt.close()
print("  -> 图表已保存: 03_类型分析.png")


# ============================================================
# 6. 维度四：玩家评价分析
# ============================================================
print("\n" + "=" * 70)
print("  维度四：玩家评价分析")
print("=" * 70)

review_df = df[df['total_reviews'] >= 50].copy()

# --- 图4: 评价分析 ---
fig, axes = plt.subplots(2, 2, figsize=(14, 10))
fig.suptitle('Steam 玩家评价深度分析', fontsize=16, fontweight='bold')

# 子图1: 评价数 vs 好评率
ax1 = axes[0, 0]
sample = review_df.sample(min(5000, len(review_df)), random_state=42)
sc1 = ax1.scatter(np.log10(sample['total_reviews']), sample['positive_rate'],
                  c=sample['Price'], cmap='plasma', alpha=0.4, s=8, vmin=0, vmax=50)
ax1.axhline(y=80, color='red', linestyle='--', alpha=0.5, label='80% 好评线')
ax1.set_xlabel('log10(总评价数)')
ax1.set_ylabel('好评率 (%)')
ax1.set_title('评价数量 vs 好评率（颜色=价格）', fontsize=12)
ax1.legend(fontsize=8)
cbar1 = plt.colorbar(sc1, ax=ax1, shrink=0.8)
cbar1.set_label('价格 (USD)')

# 子图2: Metacritic vs Steam
ax2 = axes[0, 1]
mc_data = review_df[review_df['Metacritic score'].notna() & review_df['positive_rate'].notna()]
if len(mc_data) > 10:
    ax2.scatter(mc_data['Metacritic score'], mc_data['positive_rate'],
                alpha=0.3, s=8, color='teal')
    z = np.polyfit(mc_data['Metacritic score'], mc_data['positive_rate'], 1)
    p = np.poly1d(z)
    x_line = np.linspace(mc_data['Metacritic score'].min(), mc_data['Metacritic score'].max(), 100)
    mc_corr = mc_data['Metacritic score'].corr(mc_data['positive_rate'])
    ax2.plot(x_line, p(x_line), 'r--', linewidth=2, label=f'趋势线 r={mc_corr:.2f}')
    ax2.set_title('Metacritic媒体评分 vs Steam玩家好评率', fontsize=12)
    ax2.set_xlabel('Metacritic 评分')
    ax2.set_ylabel('Steam 好评率 (%)')
    ax2.legend(fontsize=8)
    print(f"\n  Metacritic vs Steam 相关系数: {mc_corr:.3f}")
else:
    ax2.text(0.5, 0.5, 'Metacritic数据不足', ha='center', va='center', transform=ax2.transAxes)
    mc_corr = np.nan
    print('\n  Metacritic 数据不足，无法进行相关性分析')

# 子图3: 分组对比
ax3 = axes[1, 0]
box_groups = []
box_labels = []
# 平台数
for pc in [1, 2, 3]:
    data = review_df[review_df['platform_count'] == pc]['positive_rate'].dropna()
    if len(data) > 20:
        box_groups.append(data)
        box_labels.append(f'支持{pc}平台')
# DLC
for has_dlc, label in [(True, '有DLC'), (False, '无DLC')]:
    cond = review_df['DLC count'] > 0 if has_dlc else review_df['DLC count'] == 0
    data = review_df[cond]['positive_rate'].dropna()
    if len(data) > 20:
        box_groups.append(data)
        box_labels.append(label)
if box_groups:
    bp = ax3.boxplot(box_groups, labels=box_labels, patch_artist=True, showmeans=True)
    colors_box = plt.cm.Set2(np.linspace(0, 1, len(box_groups)))
    for patch, color in zip(bp['boxes'], colors_box):
        patch.set_facecolor(color)
    ax3.set_title('好评率影响因素分组对比', fontsize=12)
    ax3.set_ylabel('好评率 (%)')
    ax3.tick_params(axis='x', rotation=20)
    print("\n  分组好评率统计:")
    for label, data in zip(box_labels, box_groups):
        print(f"    {label:15s}: 中位数={data.median():.1f}%, 均值={data.mean():.1f}%, 标准差={data.std():.1f}%")

# 子图4: 成就与好评率
ax4 = axes[1, 1]
achievement_data = review_df[review_df['Achievements'] > 0].copy()
if len(achievement_data) > 20:
    # 使用显式对数分箱边界，确保标签与实际区间对应
    bin_edges = [0, 1, 2, 3, 4, 5, 6, 7, 8]  # log10 尺度: 1-10, 10-100, ...
    bin_labels = ['1-10', '10-100', '100-1千', '1千-1万', '1万-10万',
                  '10万-100万', '100万-1000万', '1000万-1亿']
    achievement_data['achieve_bin'] = pd.cut(
        np.log10(achievement_data['Achievements'].clip(1)),
        bins=bin_edges,
        labels=bin_labels,
        include_lowest=True
    )
    ach_rate = achievement_data.groupby('achieve_bin')['positive_rate'].median()
    valid_bins = [b for b in ach_rate.index if not pd.isna(ach_rate[b]) and ach_rate[b] > 0]
    ax4.bar(range(len(valid_bins)), [ach_rate[b] for b in valid_bins], color='steelblue')
    ax4.set_xticks(range(len(valid_bins)))
    ax4.set_xticklabels(valid_bins, rotation=45, ha='right', fontsize=8)
    ax4.set_title('成就数量与好评率关系', fontsize=12)
    ax4.set_xlabel('成就数量范围')
    ax4.set_ylabel('中位数好评率 (%)')
    ax4.axhline(y=review_df['positive_rate'].median(), color='red', linestyle='--', alpha=0.5)

plt.tight_layout()
plt.savefig('04_评价分析.png', dpi=150)
plt.close()
print("  -> 图表已保存: 04_评价分析.png")

# 高评价 vs 低评价
print(f"\n  高评价(>=90%) vs 低评价(<60%) 游戏特征对比:")
high_rate = review_df[review_df['positive_rate'] >= 90]
low_rate = review_df[review_df['positive_rate'] < 60]
for label, grp in [('高评价(>=90%)', high_rate), ('低评价(<60%)', low_rate)]:
    print(f"    {label}: {len(grp)}款, 均价${grp['Price'].mean():.2f}, "
          f"中位拥有者{grp['owners_num'].median():,.0f}, "
          f"中位评价数{grp['total_reviews'].median():,.0f}")


# ============================================================
# 7. 分析总结与报告输出
# ============================================================
print("\n" + "=" * 70)
print("  分析总结")
print("=" * 70)

total = len(df)
with_reviews = len(review_df)
year_min = int(df['release_year'].min())
year_max = int(df['release_year'].max())
peak_year = int(year_counts.idxmax())
peak_count = int(year_counts.max())

summary_report = f"""
{'='*60}
     Steam 游戏市场数据分析报告
     生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M')}
{'='*60}

【数据集概览】
  总游戏/应用数:       {total:,}
  时间跨度:             {year_min} - {year_max}
  有用户评价的游戏:    {with_reviews:,} ({with_reviews/total*100:.1f}%)
  付费游戏占比:         {(~df['is_free']).sum()/total*100:.1f}%
  支持 Windows:         {df['Windows'].sum()/total*100:.1f}%
  支持 Mac:             {df['Mac'].sum()/total*100:.1f}%
  支持 Linux:           {df['Linux'].sum()/total*100:.1f}%

【维度一：综合评价】
  付费游戏价格中位数:   ${paid.median():.2f}
  整体好评率中位数:     {df[df['total_reviews']>=50]['positive_rate'].median():.1f}%
  拥有者数中位数:       {owners_valid.median():,.0f}
  年度发行量峰值:       {peak_year} 年 ({peak_count} 款)
  价格与好评率相关性:   r = {corr_matrix.loc['Price', 'positive_rate']:.3f}
  评价数与好评率相关性: r = {corr_matrix.loc['total_reviews', 'positive_rate']:.3f}

【维度二：定价策略】
  免费游戏好评率中位数: {free_g['positive_rate'].median():.1f}%
  付费游戏好评率中位数: {paid_g['positive_rate'].median():.1f}%
  免费游戏拥有者中位数: {free_g['owners_num'].median():,.0f}
  付费游戏拥有者中位数: {paid_g['owners_num'].median():,.0f}

【维度三：类型分析】
  最常见类型:           {genre_counts.index[0]} ({genre_counts.iloc[0]:,} 款)
  好评率最高类型(Top15): {genre_rate_top.index[-1]} ({genre_rate_top.iloc[-1]:.1f}%)
  平均价格最高类型:     {genre_price_top.index[-1]} (${genre_price_top.iloc[-1]:.2f})

【维度四：玩家评价】
  高评价游戏(>=90%):    {len(high_rate)} 款
  低评价游戏(<60%):     {len(low_rate)} 款

{'='*60}
"""

print(summary_report)

with open('分析报告.txt', 'w', encoding='utf-8') as f:
    f.write(summary_report)

print("\n" + "=" * 70)
print("  全部分析完成!")
print("  生成的图表文件:")
print("    01_综合概览.png")
print("    01_相关性热力图.png")
print("    02_定价策略分析.png")
print("    03_类型分析.png")
print("    04_评价分析.png")
print("    分析报告.txt")
print("=" * 70)
