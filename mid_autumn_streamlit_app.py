# -*- coding: utf-8 -*-
"""
Created on Sat Jun 28 23:07:19 2025

@author: User
"""

# mid_autumn_streamlit_app.py

import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

st.set_page_config(page_title="中秋銷售分析", layout="wide")
st.title("🎑 中秋檔期銷售分析儀表板")

# -------- 函式區塊 --------

def prepare_data(xls):
    sheets = {sheet: xls.parse(sheet) for sheet in xls.sheet_names}
    week_cols = [col for col in sheets['每週銷售數量'].columns if '–' in col]
    sorted_weeks = sorted(week_cols, key=lambda x: pd.to_datetime(x.split('–')[0] + '/2024'))
    for key in sheets:
        df = sheets[key]
        sheets[key] = df[['產品代號', '產品名稱'] + sorted_weeks]
    return sheets, sorted_weeks

def merge_product_analysis(sheets, product_id, weeks):
    df_weekly = sheets['每週銷售數量']
    df_cumulative = sheets['累計每週銷售數量']
    df_growth = sheets['每週銷售成長率']

    row_weekly = df_weekly[df_weekly['產品代號'] == product_id].iloc[0]
    row_cumulative = df_cumulative[df_cumulative['產品代號'] == product_id].iloc[0]
    row_growth = df_growth[df_growth['產品代號'] == product_id].iloc[0]

    df_analysis = pd.DataFrame({
        '週別': weeks,
        '單週銷售量': row_weekly[weeks].astype(float).values,
        '累積銷售量': row_cumulative[weeks].astype(float).values,
        '週成長率(%)': row_growth[weeks].astype(float).values
    })
    return df_analysis

def plot_product_sales(df_analysis, product_name):
    import matplotlib
    from matplotlib import font_manager

    font_path = "NotoSansTC-VariableFont_wght.ttf"
    prop = font_manager.FontProperties(fname=font_path)

    matplotlib.rcParams['axes.unicode_minus'] = False

    fig, axes = plt.subplots(3, 1, figsize=(14, 10), sharex=True)

    sns.barplot(x='週別', y='單週銷售量', data=df_analysis, ax=axes[0])
    axes[0].set_title(f'{product_name} 單週銷售量', fontproperties=prop)
    axes[0].set_ylabel("單週銷售量", fontproperties=prop)
    axes[0].tick_params(axis='x', rotation=45)
    for label in axes[0].get_xticklabels():
        label.set_fontproperties(prop)
    for label in axes[0].get_yticklabels():
        label.set_fontproperties(prop)

    sns.lineplot(x='週別', y='累積銷售量', data=df_analysis, marker='o', ax=axes[1])
    axes[1].set_title(f'{product_name} 累積銷售量', fontproperties=prop)
    axes[1].set_ylabel("累積銷售量", fontproperties=prop)
    for label in axes[1].get_xticklabels():
        label.set_fontproperties(prop)
    for label in axes[1].get_yticklabels():
        label.set_fontproperties(prop)

    sns.lineplot(x='週別', y='週成長率(%)', data=df_analysis, marker='o', ax=axes[2], color='red')
    axes[2].set_title(f'{product_name} 週成長率 (%)', fontproperties=prop)
    axes[2].set_ylabel("週成長率(%)", fontproperties=prop)
    axes[2].axhline(0, ls='--', color='gray')
    for label in axes[2].get_xticklabels():
        label.set_fontproperties(prop)
    for label in axes[2].get_yticklabels():
        label.set_fontproperties(prop)

    plt.tight_layout()
    return fig

# -------- 主程式區塊 --------

uploaded_file = st.file_uploader("請上傳中秋彙整報表（含三張工作表）", type=["xlsx"])

if uploaded_file:
    xls = pd.ExcelFile(uploaded_file)
    sheets, weeks = prepare_data(xls)

    product_ids = sheets['每週銷售數量']['產品代號'].tolist()
    product_names = sheets['每週銷售數量']['產品名稱'].tolist()
    product_map = dict(zip(product_ids, product_names))

    product_id = st.selectbox(
        "請選擇要分析的產品代號：",
        product_ids,
        format_func=lambda x: f"{x} - {product_map.get(x, '')}"
    )
    product_name = product_map[product_id]

    df_analysis = merge_product_analysis(sheets, product_id, weeks)
    st.subheader(f"📊 {product_name} 銷售數據表")
    st.dataframe(df_analysis, use_container_width=True)

    st.subheader("📈 趨勢圖")
    fig = plot_product_sales(df_analysis, product_name)
    st.pyplot(fig)
else:
    st.info("請先上傳檔案後再開始分析。")
