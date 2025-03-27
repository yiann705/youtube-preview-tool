import pandas as pd
import numpy as np

"""
@description 廣告數據分析腳本
@author AI Assistant
@date 2024
@version 1.0.0

功能說明：
- 讀取廣告數據CSV文件並進行分析
- 計算各項廣告指標（KPI）
- 生成分析報告

輸入：
- GG.csv：包含廣告數據的CSV文件（Big5編碼）
  欄位：月份、聯播網、曝光、點擊、安裝、費用等

輸出：
- ad_stats_result.xlsx：包含以下分析結果的Excel文件
  1. 曝光：整數，千分位
  2. 費用：整數，千分位
  3. CPI：整數，千分位
  4. IR：保留三位小數的百分比 (x.xxx%)
  5. CPM：整數，千分位
  6. CTR：保留三位小數的百分比 (x.xxx%)

計算公式：
- CPI = 費用 / 安裝數
- IR = (安裝數 / 曝光數) * 100%
- CPM = (費用 / 曝光數) * 1000
- CTR = (點擊數 / 曝光數) * 100%

備註：
- 橫向總計：使用當月所有平台的原始數據總和重新計算
- 縱向總計：使用該平台所有月份的原始數據總和重新計算
- Google 搜尋數據不包含搜尋夥伴
"""

# 讀取 CSV 文件，使用 UTF-8-SIG 編碼
df = pd.read_csv('GG.csv', encoding='utf-8-sig')

# 將日期列轉換為日期格式
df['月'] = pd.to_datetime(df['月'])

# 創建日期範圍
date_range = pd.date_range(start='2024-01-01', end='2025-02-01', freq='MS')

# 創建所有指標的 DataFrame
metrics = ['曝光', '費用', 'CPI', 'IR', 'CPM', 'CTR']
all_results = {}

# 創建原始數據的 DataFrame（用於計算總計）
raw_data = {
    '曝光': pd.DataFrame(index=date_range),
    '點擊': pd.DataFrame(index=date_range),
    '安裝': pd.DataFrame(index=date_range),
    '費用': pd.DataFrame(index=date_range)
}

# 初始化 DataFrame
platforms = ['Google 多媒體廣告聯播網', 'Google 搜尋', 'YouTube', '總計']
for df_name in raw_data:
    raw_data[df_name].index = raw_data[df_name].index.strftime('%Y/%m/1')
    for platform in platforms:
        raw_data[df_name][platform] = 0.0

for metric in metrics:
    all_results[metric] = pd.DataFrame(index=date_range)
    all_results[metric].index = all_results[metric].index.strftime('%Y/%m/1')
    for platform in platforms:
        all_results[metric][platform] = 0.0

# 處理每個月的數據
for date in date_range:
    date_str = date.strftime('%Y/%m/1')
    month_data = df[df['月'].dt.strftime('%Y/%m/1') == date_str]
    
    # Google 多媒體廣告聯播網
    display_data = month_data[month_data['聯播網 (及搜尋夥伴)'] == 'Google 多媒體廣告聯播網']
    
    # Google 搜尋（排除搜尋夥伴）
    search_data = month_data[month_data['聯播網 (及搜尋夥伴)'] == 'Google 搜尋']
    
    # YouTube
    youtube_data = month_data[month_data['聯播網 (及搜尋夥伴)'] == 'YouTube']
    
    # 處理每個平台的指標
    platforms_data = {
        'Google 多媒體廣告聯播網': display_data,
        'Google 搜尋': search_data,
        'YouTube': youtube_data
    }
    
    # 計算每個平台的基礎數據
    for platform, data in platforms_data.items():
        # 儲存原始數據
        raw_data['曝光'].loc[date_str, platform] = data['曝光'].sum()
        raw_data['點擊'].loc[date_str, platform] = data['點擊'].sum()
        raw_data['安裝'].loc[date_str, platform] = data['安裝'].sum()
        raw_data['費用'].loc[date_str, platform] = data['費用'].sum()
        
        # 計算各項指標
        impressions = data['曝光'].sum()
        clicks = data['點擊'].sum()
        installs = data['安裝'].sum()
        cost = data['費用'].sum()
        
        # 儲存基礎指標
        all_results['曝光'].loc[date_str, platform] = impressions
        all_results['費用'].loc[date_str, platform] = cost
        
        # 計算比率指標
        all_results['CPI'].loc[date_str, platform] = cost / installs if installs > 0 else 0
        all_results['IR'].loc[date_str, platform] = (installs / impressions * 100) if impressions > 0 else 0
        all_results['CPM'].loc[date_str, platform] = (cost / impressions * 1000) if impressions > 0 else 0
        all_results['CTR'].loc[date_str, platform] = (clicks / impressions * 100) if impressions > 0 else 0

    # 計算每月的總計
    for metric in raw_data:
        raw_data[metric].loc[date_str, '總計'] = sum(raw_data[metric].loc[date_str, p] for p in platforms_data.keys())
    
    # 計算每月總計的比率指標
    total_impressions = raw_data['曝光'].loc[date_str, '總計']
    total_clicks = raw_data['點擊'].loc[date_str, '總計']
    total_installs = raw_data['安裝'].loc[date_str, '總計']
    total_cost = raw_data['費用'].loc[date_str, '總計']
    
    all_results['曝光'].loc[date_str, '總計'] = total_impressions
    all_results['費用'].loc[date_str, '總計'] = total_cost
    all_results['CPI'].loc[date_str, '總計'] = total_cost / total_installs if total_installs > 0 else 0
    all_results['IR'].loc[date_str, '總計'] = (total_installs / total_impressions * 100) if total_impressions > 0 else 0
    all_results['CPM'].loc[date_str, '總計'] = (total_cost / total_impressions * 1000) if total_impressions > 0 else 0
    all_results['CTR'].loc[date_str, '總計'] = (total_clicks / total_impressions * 100) if total_impressions > 0 else 0

# 添加總計行
for platform in platforms:
    # 計算每個平台所有月份的總和
    total_impressions = raw_data['曝光'][platform].sum()
    total_clicks = raw_data['點擊'][platform].sum()
    total_installs = raw_data['安裝'][platform].sum()
    total_cost = raw_data['費用'][platform].sum()
    
    # 添加總計行
    for metric in metrics:
        if metric == '曝光':
            all_results[metric].loc['總計', platform] = total_impressions
        elif metric == '費用':
            all_results[metric].loc['總計', platform] = total_cost
        elif metric == 'CPI':
            all_results[metric].loc['總計', platform] = total_cost / total_installs if total_installs > 0 else 0
        elif metric == 'IR':
            all_results[metric].loc['總計', platform] = (total_installs / total_impressions * 100) if total_impressions > 0 else 0
        elif metric == 'CPM':
            all_results[metric].loc['總計', platform] = (total_cost / total_impressions * 1000) if total_impressions > 0 else 0
        elif metric == 'CTR':
            all_results[metric].loc['總計', platform] = (total_clicks / total_impressions * 100) if total_impressions > 0 else 0

# 格式化數據
for metric, df in all_results.items():
    df = df.reset_index().rename(columns={'index': '月份'})
    
    # 根據不同指標設置不同格式
    if metric in ['曝光', '費用', 'CPI', 'CPM']:
        for col in df.columns:
            if col != '月份':
                df[col] = df[col].apply(lambda x: f"{int(x):,}" if x != 0 else '')
    elif metric in ['IR', 'CTR']:
        for col in df.columns:
            if col != '月份':
                df[col] = df[col].apply(lambda x: f"{x:.3f}%" if x != 0 else '')
    
    all_results[metric] = df

# 保存到 Excel
with pd.ExcelWriter('ad_stats_result.xlsx', engine='openpyxl') as writer:
    for metric, df in all_results.items():
        df.to_excel(writer, sheet_name=metric, index=False)

print("分析完成，結果已保存到 ad_stats_result.xlsx")

# 顯示曝光數據預覽
print("\n曝光數據預覽：")
print(all_results['曝光'].to_string()) 