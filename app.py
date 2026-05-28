"""
产品安全库存管理系统
功能：支持多种文件格式，自动计算安全库存，月度更新
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO
import re
from datetime import datetime

# ==================== 页面设置 ====================
st.set_page_config(
    page_title="产品安全库存管理系统", 
    layout="wide", 
    page_icon="📦",
    initial_sidebar_state="expanded"
)

# ==================== 清爽配色CSS样式 ====================
st.markdown("""
<style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    
    .main .block-container {
        padding-top: 1rem;
        padding-bottom: 0rem;
    }
    
    .main-title {
        font-size: 1.8rem;
        font-weight: 600;
        color: #1e3a8a;
        margin-bottom: 0.2rem;
    }
    
    .subtitle {
        font-size: 0.85rem;
        color: #64748b;
        margin-bottom: 1.5rem;
        padding-bottom: 10px;
        border-bottom: 2px solid #e2e8f0;
    }
    
    .kpi-card {
        background: white;
        border-radius: 16px;
        padding: 16px;
        text-align: center;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
        border: 1px solid #e2e8f0;
    }
    
    .metric-value {
        font-size: 1.8rem;
        font-weight: 700;
        color: #1e3a8a;
    }
    .metric-label {
        font-size: 0.8rem;
        color: #64748b;
    }
    
    .upload-area {
        border: 2px dashed #cbd5e1;
        border-radius: 16px;
        padding: 40px;
        text-align: center;
        background: #fafcff;
    }
</style>
""", unsafe_allow_html=True)

# ==================== 页面头部 ====================
st.markdown('<div class="main-title">📦 产品安全库存管理系统</div>', unsafe_allow_html=True)
st.markdown('<div class="subtitle">⚡ 支持月度数据更新 | 🔄 自动计算安全库存 | 🎯 智能预警 | 📁 支持Excel/CSV</div>', unsafe_allow_html=True)

# ==================== 数据加载函数 ====================
@st.cache_data
def load_safety_data(file):
    """加载安全库存数据"""
    file_name = file.name.lower()
    
    try:
        if file_name.endswith('.csv'):
            df = pd.read_csv(file, encoding='utf-8')
            if df.columns.str.contains('Unnamed').any():
                df = pd.read_csv(file, encoding='gbk')
            return df
        else:
            # 直接读取Excel，跳过前2行
            df = pd.read_excel(file, sheet_name="安全库存（202509月）", header=2)
            # 清理列名
            df.columns = [str(col).replace('\n', '').strip() for col in df.columns]
            return df
    except Exception as e:
        st.error(f"读取数据失败：{e}")
        return None

# ==================== 数据处理函数 ====================
def process_data(df_safety):
    """处理数据，计算库存状态"""
    if df_safety is None or len(df_safety) == 0:
        return None
    
    # 识别关键列
    material_col = None
    for col in df_safety.columns:
        if '物料编码' in str(col) or '物料代码' in str(col):
            material_col = col
            break
    if material_col is None:
        material_col = df_safety.columns[0]
    
    safety_col = None
    for col in df_safety.columns:
        if '安全库存' in str(col):
            safety_col = col
            break
    
    actual_col = None
    for col in df_safety.columns:
        if '实际库存' in str(col) or '6月末实际库存' in str(col):
            actual_col = col
            break
    
    # 重命名
    df_safety = df_safety.rename(columns={material_col: "物料编码"})
    if safety_col:
        df_safety = df_safety.rename(columns={safety_col: "安全库存"})
    if actual_col:
        df_safety = df_safety.rename(columns={actual_col: "实际库存"})
    
    for col in ["安全库存", "实际库存"]:
        if col in df_safety.columns:
            df_safety[col] = pd.to_numeric(df_safety[col], errors="coerce").fillna(0)
        else:
            df_safety[col] = 0
    
    df_safety = df_safety.dropna(subset=["物料编码"], how="all")
    df_safety = df_safety[df_safety["物料编码"].notna()]
    df_safety = df_safety[df_safety["物料编码"].astype(str).str.len() > 0]
    
    # 计算库存状态
    df_safety["库存状态"] = df_safety.apply(
        lambda row: "🔴 低于安全库存" if row["实际库存"] < row["安全库存"] and row["安全库存"] > 0 
        else "🟡 临界状态" if row["实际库存"] <= row["安全库存"] * 1.2 and row["安全库存"] > 0
        else "🟢 库存充足" if row["安全库存"] > 0
        else "⚪ 未设置",
        axis=1
    )
    
    df_safety["缺货风险"] = ((df_safety["安全库存"] - df_safety["实际库存"]) / df_safety["安全库存"] * 100).clip(lower=0).round(1)
    df_safety["缺货风险"] = df_safety["缺货风险"].fillna(0).astype(int)
    
    # 月份识别
    month_pattern = re.compile(r'(\d{4})[年M](\d{1,2})')
    month_cols = [col for col in df_safety.columns if month_pattern.search(str(col))]
    if len(month_cols) >= 3:
        latest_months = month_cols[-3:]
        df_safety["最近3个月月均用量"] = df_safety[latest_months].mean(axis=1).round(0)
    
    return df_safety

# ==================== 初始化Session State ====================
if 'selected_menu' not in st.session_state:
    st.session_state.selected_menu = "概览仪表盘"
if 'uploaded_file' not in st.session_state:
    st.session_state.uploaded_file = None
if 'df_processed' not in st.session_state:
    st.session_state.df_processed = None

# ==================== 侧边栏 ====================
with st.sidebar:
    st.markdown("### 🧭 导航菜单")
    
    menu_options = ["概览仪表盘", "数据上传", "特征分析", "需求预测", "库存优化", "决策建议"]
    menu_icons = {
        "概览仪表盘": "📊",
        "数据上传": "📂",
        "特征分析": "📈",
        "需求预测": "📉",
        "库存优化": "🎯",
        "决策建议": "💡"
    }
    
    for option in menu_options:
        icon = menu_icons.get(option, "📌")
        if st.button(f"{icon} {option}", key=f"nav_{option}", use_container_width=True):
            st.session_state.selected_menu = option
            st.rerun()
    
    st.markdown("---")
    st.markdown("### 📂 文件上传")
    
    # 文件上传器
    uploaded_file = st.file_uploader(
        "选择文件",
        type=["xlsx", "xls", "csv"],
        help="支持 .xlsx, .xls, .csv 格式",
        key="sidebar_uploader"
    )
    
    # 当上传新文件时，处理数据
    if uploaded_file is not None:
        if st.session_state.uploaded_file != uploaded_file.name:
            st.session_state.uploaded_file = uploaded_file.name
            with st.spinner("正在处理数据..."):
                df_raw = load_safety_data(uploaded_file)
                if df_raw is not None:
                    st.session_state.df_processed = process_data(df_raw)
                    st.success(f"✅ 已加载: {uploaded_file.name}")
                    st.rerun()
    
    # 显示当前文件状态
    if st.session_state.uploaded_file is not None:
        st.info(f"📄 当前文件: {st.session_state.uploaded_file}")

# ==================== 获取处理后的数据 ====================
df = st.session_state.df_processed
current_menu = st.session_state.selected_menu

# ==================== 页面1：概览仪表盘 ====================
if current_menu == "概览仪表盘":
    st.markdown("### 📊 概览仪表盘")
    
    if df is not None and len(df) > 0:
        # KPI卡片
        col1, col2, col3, col4 = st.columns(4)
        total = len(df)
        risk = len(df[df["库存状态"] == "🔴 低于安全库存"])
        avg_safety = df["安全库存"].mean()
        valid = df[df["安全库存"] > 0]
        coverage = (valid["实际库存"] / valid["安全库存"]).mean() if len(valid) > 0 else 0
        
        with col1:
            st.markdown(f"""
            <div class="kpi-card">
                <div style="font-size: 2rem;">📦</div>
                <div class="metric-value">{total}</div>
                <div class="metric-label">物料总数</div>
            </div>
            """, unsafe_allow_html=True)
        with col2:
            st.markdown(f"""
            <div class="kpi-card">
                <div style="font-size: 2rem;">⚠️</div>
                <div class="metric-value">{risk}</div>
                <div class="metric-label">低库存预警 (占{round(risk/total*100)}%)</div>
            </div>
            """, unsafe_allow_html=True)
        with col3:
            st.markdown(f"""
            <div class="kpi-card">
                <div style="font-size: 2rem;">📊</div>
                <div class="metric-value">{avg_safety:.0f}</div>
                <div class="metric-label">平均安全库存(件)</div>
            </div>
            """, unsafe_allow_html=True)
        with col4:
            st.markdown(f"""
            <div class="kpi-card">
                <div style="font-size: 2rem;">⏱️</div>
                <div class="metric-value">{coverage:.1f}</div>
                <div class="metric-label">平均库存覆盖倍数</div>
            </div>
            """, unsafe_allow_html=True)
        
        st.markdown("---")
        
        # 库存监控看板
        st.markdown("### 📈 库存监控看板")
        
        col_chart1, col_chart2 = st.columns(2)
        with col_chart1:
            status_counts = df["库存状态"].value_counts()
            if len(status_counts) > 0:
                colors = {"🔴 低于安全库存": "#ef4444", "🟡 临界状态": "#f59e0b", "🟢 库存充足": "#10b981", "⚪ 未设置": "#94a3b8"}
                fig_pie = px.pie(
                    values=status_counts.values,
                    names=status_counts.index,
                    title="库存状态分布",
                    color=status_counts.index,
                    color_discrete_map=colors,
                    hole=0.4
                )
                fig_pie.update_traces(textposition="inside", textinfo="percent+label")
                fig_pie.update_layout(height=400)
                st.plotly_chart(fig_pie, use_container_width=True)
        
        with col_chart2:
            safety_data = df[df["安全库存"] > 0]["安全库存"]
            if len(safety_data) > 0:
                fig_hist = px.histogram(
                    safety_data, nbins=30,
                    title="安全库存值分布",
                    labels={"value": "安全库存 (件)", "count": "物料数量"},
                    color_discrete_sequence=["#3b82f6"],
                    opacity=0.7
                )
                fig_hist.update_layout(height=400)
                st.plotly_chart(fig_hist, use_container_width=True)
        
        # 物料列表
        st.markdown("### 📋 物料安全库存列表")
        
        col_f1, col_f2 = st.columns(2)
        with col_f1:
            status_filter = st.multiselect(
                "按状态筛选",
                options=df["库存状态"].unique(),
                default=df["库存状态"].unique()
            )
        with col_f2:
            search = st.text_input("搜索物料编码", placeholder="输入物料编码...")
        
        filtered_df = df[df["库存状态"].isin(status_filter)]
        if search:
            filtered_df = filtered_df[filtered_df["物料编码"].astype(str).str.contains(search, na=False)]
        
        display_cols = ["物料编码", "安全库存", "实际库存", "库存状态", "缺货风险"]
        if "最近3个月月均用量" in filtered_df.columns:
            display_cols.append("最近3个月月均用量")
        available_cols = [c for c in display_cols if c in filtered_df.columns]
        
        st.dataframe(filtered_df[available_cols], use_container_width=True, height=400)
        
        # 导出
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            filtered_df.to_excel(writer, index=False, sheet_name="安全库存明细")
        st.download_button(
            label="📎 导出数据报表",
            data=output.getvalue(),
            file_name=f"安全库存明细_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
        st.success("✅ 数据加载成功")
    else:
        st.markdown("""
        <div class="upload-area">
            <div style="font-size: 3rem;">📁</div>
            <div style="font-size: 1rem; font-weight: 500; margin: 10px 0;">请从左侧上传文件</div>
            <div style="color: #6c757d;">支持 .xlsx, .xls, .csv 格式</div>
        </div>
        """, unsafe_allow_html=True)

# ==================== 页面2：数据上传 ====================
elif current_menu == "数据上传":
    st.markdown("### 📂 数据上传")
    
    if df is not None and len(df) > 0:
        st.success(f"✅ 当前已上传并处理：{st.session_state.uploaded_file}")
        st.markdown("### 📄 数据预览")
        st.dataframe(df.head(10), use_container_width=True)
        st.info(f"📊 数据维度：{df.shape[0]} 行 × {df.shape[1]} 列")
    else:
        st.markdown("""
        <div class="upload-area">
            <div style="font-size: 2rem;">📂</div>
            <div style="font-weight: 500;">从左侧边栏选择文件上传</div>
            <div style="font-size: 0.8rem; color: #6c757d; margin-top: 10px;">支持 .xlsx, .xls, .csv 格式</div>
        </div>
        """, unsafe_allow_html=True)

# ==================== 页面3：特征分析 ====================
elif current_menu == "特征分析":
    st.markdown("### 📈 需求特征分析")
    
    if df is not None and len(df) > 0:
        numeric_cols = df.select_dtypes(include=[np.number]).columns.tolist()
        if len(numeric_cols) > 0:
            st.markdown("#### 📊 数据统计")
            st.dataframe(df[numeric_cols].describe(), use_container_width=True)
            
            if len(numeric_cols) >= 2:
                st.markdown("#### 🔗 相关性分析")
                corr_matrix = df[numeric_cols].corr()
                fig_corr = px.imshow(corr_matrix, text_auto=True, title="特征相关性矩阵", color_continuous_scale="Blues", aspect="auto")
                fig_corr.update_layout(height=500)
                st.plotly_chart(fig_corr, use_container_width=True)
        else:
            st.info("未找到数值型特征")
    else:
        st.info("请先上传数据文件")

# ==================== 页面4：需求预测 ====================
elif current_menu == "需求预测":
    st.markdown("### 📉 需求预测")
    
    if df is not None and len(df) > 0:
        month_pattern = re.compile(r'(\d{4})[年M](\d{1,2})')
        month_cols = [col for col in df.columns if month_pattern.search(str(col))]
        
        if len(month_cols) >= 3:
            st.markdown("#### 📈 历史销量趋势")
            top_materials = df.head(10)["物料编码"].tolist()
            trend_data = []
            for material in top_materials[:5]:
                row = df[df["物料编码"] == material]
                if len(row) > 0:
                    for col in month_cols:
                        val = pd.to_numeric(row[col].values[0], errors="coerce") if col in row.columns else 0
                        trend_data.append({"物料": material, "月份": col, "销量": val})
            
            if trend_data:
                df_trend = pd.DataFrame(trend_data)
                fig_trend = px.line(df_trend, x="月份", y="销量", color="物料", title="历史销量趋势", markers=True)
                fig_trend.update_layout(height=450)
                st.plotly_chart(fig_trend, use_container_width=True)
        else:
            st.info("未识别到月份销量数据，请确保列名包含年份月份（如：2025年M07）")
    else:
        st.info("请先上传数据文件")

# ==================== 页面5：库存优化 ====================
elif current_menu == "库存优化":
    st.markdown("### 🎯 智能库存优化")
    
    if df is not None and len(df) > 0:
        low_stock = df[df["实际库存"] < df["安全库存"]]
        if len(low_stock) > 0:
            st.info(f"📊 共有 **{len(low_stock)}** 种物料库存不足，需要补货")
            low_stock["建议补货量"] = (low_stock["安全库存"] - low_stock["实际库存"]).round(0).astype(int)
            st.dataframe(low_stock[["物料编码", "实际库存", "安全库存", "建议补货量"]].head(20), use_container_width=True)
        else:
            st.success("🎉 所有物料库存充足！")
    else:
        st.info("请先上传数据文件")

# ==================== 页面6：决策建议 ====================
elif current_menu == "决策建议":
    st.markdown("### 💡 决策建议")
    
    if df is not None and len(df) > 0:
        total = len(df)
        risk_count = len(df[df["实际库存"] < df["安全库存"]])
        risk_rate = risk_count / total * 100 if total > 0 else 0
        
        st.markdown(f"""
        <div class="kpi-card">
            <div class="metric-value">{risk_rate:.1f}%</div>
            <div class="metric-label">低库存物料占比</div>
        </div>
        """, unsafe_allow_html=True)
        
        if risk_rate > 30:
            st.warning("🔴 高风险：超过30%物料库存不足，建议立即安排补货并审查采购计划")
        elif risk_rate > 15:
            st.warning("🟡 中风险：部分物料库存不足，建议重点关注缺货风险高的物料")
        else:
            st.success("🟢 低风险：库存整体健康，建议保持当前管理水平")
        
        if risk_count > 0:
            low_stock = df[df["实际库存"] < df["安全库存"]].copy()
            low_stock["建议补货量"] = (low_stock["安全库存"] - low_stock["实际库存"]).round(0).astype(int)
            st.markdown("#### 📦 建议补货清单")
            st.dataframe(low_stock[["物料编码", "实际库存", "安全库存", "建议补货量"]].head(15), use_container_width=True)
    else:
        st.info("请先上传数据文件")
