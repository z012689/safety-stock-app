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
    initial_sidebar_state="collapsed"
)

# ==================== 丰富色彩CSS样式 ====================
st.markdown("""
<style>
    /* 隐藏默认元素 */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    
    /* 主容器 */
    .main .block-container {
        padding-top: 1rem;
        padding-bottom: 0rem;
    }
    
    /* 多彩渐变标题 */
    .gradient-title {
        font-size: 2.2rem;
        font-weight: 700;
        background: linear-gradient(135deg, #667eea 0%, #764ba2 25%, #f093fb 50%, #4facfe 75%, #00f2fe 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        margin-bottom: 0.3rem;
    }
    
    /* 副标题 */
    .subtitle {
        font-size: 0.9rem;
        color: #64748b;
        margin-bottom: 1.5rem;
        border-left: 3px solid #3b82f6;
        padding-left: 12px;
    }
    
    /* 左侧导航栏样式 */
    .nav-sidebar {
        background: linear-gradient(180deg, #1e1b4b 0%, #312e81 50%, #4338ca 100%);
        border-radius: 20px;
        padding: 20px 15px;
        margin-right: 15px;
        color: white;
    }
    
    .nav-item {
        padding: 12px 15px;
        margin: 5px 0;
        border-radius: 12px;
        cursor: pointer;
        transition: all 0.2s;
        font-weight: 500;
        color: #e0e7ff;
    }
    
    .nav-item:hover {
        background-color: rgba(255,255,255,0.15);
        color: white;
    }
    
    .nav-item-active {
        background: linear-gradient(90deg, #f59e0b, #ef4444);
        color: white;
        box-shadow: 0 4px 10px rgba(0,0,0,0.2);
    }
    
    /* 多彩KPI卡片 */
    .kpi-purple {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        border-radius: 20px;
        padding: 20px;
        color: white;
        text-align: center;
        transition: transform 0.2s;
    }
    .kpi-purple:hover { transform: translateY(-5px); }
    
    .kpi-blue {
        background: linear-gradient(135deg, #4facfe 0%, #00f2fe 100%);
        border-radius: 20px;
        padding: 20px;
        color: white;
        text-align: center;
        transition: transform 0.2s;
    }
    .kpi-blue:hover { transform: translateY(-5px); }
    
    .kpi-orange {
        background: linear-gradient(135deg, #f093fb 0%, #f5576c 100%);
        border-radius: 20px;
        padding: 20px;
        color: white;
        text-align: center;
        transition: transform 0.2s;
    }
    .kpi-orange:hover { transform: translateY(-5px); }
    
    .kpi-green {
        background: linear-gradient(135deg, #43e97b 0%, #38f9d7 100%);
        border-radius: 20px;
        padding: 20px;
        color: #1e3a5f;
        text-align: center;
        transition: transform 0.2s;
    }
    .kpi-green:hover { transform: translateY(-5px); }
    
    /* 功能卡片 */
    .feature-card {
        background: white;
        border-radius: 16px;
        padding: 20px;
        text-align: center;
        box-shadow: 0 4px 12px rgba(0,0,0,0.08);
        transition: all 0.3s;
        border: 1px solid #eef2ff;
    }
    .feature-card:hover {
        transform: translateY(-5px);
        box-shadow: 0 8px 25px rgba(0,0,0,0.12);
    }
    
    /* 上传区域美化 */
    .upload-area {
        border: 2px dashed #3b82f6;
        border-radius: 20px;
        padding: 30px;
        text-align: center;
        background: linear-gradient(135deg, #f0f9ff 0%, #e0f2fe 100%);
    }
    
    /* 表格样式 */
    .stDataFrame {
        border-radius: 16px;
        overflow: hidden;
    }
    
    /* 按钮样式 */
    .stButton > button {
        background: linear-gradient(90deg, #667eea, #764ba2);
        color: white;
        border-radius: 25px;
        border: none;
        padding: 8px 24px;
        font-weight: 500;
    }
    
    /* 标签页样式 */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
    }
    .stTabs [data-baseweb="tab"] {
        background-color: #f1f5f9;
        border-radius: 12px;
        padding: 8px 20px;
        font-weight: 500;
    }
    .stTabs [aria-selected="true"] {
        background: linear-gradient(90deg, #667eea, #764ba2);
        color: white;
    }
</style>
""", unsafe_allow_html=True)

# ==================== 页面头部 ====================
col_logo, col_title = st.columns([1, 5])
with col_title:
    st.markdown('<div class="gradient-title">📦 产品安全库存管理系统</div>', unsafe_allow_html=True)
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
            df = pd.read_excel(file, sheet_name="安全库存（202509月）", header=None)
            header_row = None
            for i, row in df.iterrows():
                row_str = row.astype(str).str.contains("物料编码")
                if row_str.any():
                    header_row = i
                    break
            
            if header_row is not None:
                df.columns = df.iloc[header_row]
                df = df.iloc[header_row + 1:].reset_index(drop=True)
            else:
                df = pd.read_excel(file, sheet_name="安全库存（202509月）", header=2)
            return df
    except Exception as e:
        st.error(f"读取数据失败：{e}")
        return None

@st.cache_data
def load_packaging_data(file):
    """加载包材安全库存数据"""
    try:
        df = pd.read_excel(file, sheet_name="包材安全库存", header=None)
        for i, row in df.iterrows():
            if "建议安全库存" in row.astype(str).values:
                df.columns = df.iloc[i]
                df = df.iloc[i+1:].reset_index(drop=True)
                break
        return df
    except:
        return None

# ==================== 侧边栏导航（参考样例图片） ====================
with st.sidebar:
    st.markdown("""
    <div style="background: linear-gradient(135deg, #667eea, #764ba2); border-radius: 20px; padding: 20px; text-align: center; margin-bottom: 20px;">
        <div style="font-size: 2rem;">📊</div>
        <div style="color: white; font-weight: 600; margin-top: 5px;">库存智能决策</div>
        <div style="color: rgba(255,255,255,0.7); font-size: 0.7rem;">版本 2.0</div>
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown("### 🧭 导航菜单")
    
    # 创建导航选项
    menu_options = ["📊 概览仪表盘", "📂 数据上传", "📈 特征分析", "📉 需求预测", "🎯 库存优化", "💡 决策建议"]
    
    # 使用radio作为导航（模拟侧边栏菜单）
    selected_menu = st.radio(
        "",
        menu_options,
        label_visibility="collapsed",
        index=0
    )
    
    st.markdown("---")
    st.markdown("### 🤖 AI 智能助手")
    st.info("💬 上传数据后，系统将自动分析库存状况并提供优化建议")

# ==================== 根据导航显示不同内容 ====================

# 文件上传区域（统一处理）
uploaded_file = st.file_uploader(
    "📁 点击或拖拽上传文件",
    type=["xlsx", "xls", "csv"],
    label_visibility="collapsed",
    help="支持 .xlsx, .xls, .csv 格式"
)

if selected_menu == "📊 概览仪表盘":
    # ========== 概览仪表盘页面 ==========
    
    if uploaded_file is not None:
        with st.spinner("🔄 正在加载数据..."):
            df_safety = load_safety_data(uploaded_file)
            df_packaging = load_packaging_data(uploaded_file) if uploaded_file.name.lower().endswith(('xlsx', 'xls')) else None
        
        if df_safety is not None and len(df_safety) > 0:
            # 数据清洗
            df_safety.columns = [str(col).replace('\n', '').strip() for col in df_safety.columns]
            
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
            
            # 重命名和计算
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
                st.success(f"✅ 自动识别到 {len(month_cols)} 个月份列，使用最新3个月（{', '.join(latest_months)}）计算月均用量")
            
            # ========== 彩色KPI卡片 ==========
            st.markdown("### 📊 核心指标")
            
            col1, col2, col3, col4 = st.columns(4)
            total = len(df_safety)
            risk = len(df_safety[df_safety["库存状态"] == "🔴 低于安全库存"])
            avg_safety = df_safety["安全库存"].mean()
            valid = df_safety[df_safety["安全库存"] > 0]
            coverage = (valid["实际库存"] / valid["安全库存"]).mean() if len(valid) > 0 else 0
            
            with col1:
                st.markdown(f"""
                <div class="kpi-purple">
                    <div style="font-size: 2rem;">📦</div>
                    <div style="font-size: 1.8rem; font-weight: 700;">{total}</div>
                    <div style="font-size: 0.85rem; opacity: 0.9;">物料总数</div>
                </div>
                """, unsafe_allow_html=True)
            
            with col2:
                st.markdown(f"""
                <div class="kpi-orange">
                    <div style="font-size: 2rem;">⚠️</div>
                    <div style="font-size: 1.8rem; font-weight: 700;">{risk}</div>
                    <div style="font-size: 0.85rem; opacity: 0.9;">低库存预警 | 占{round(risk/total*100)}%</div>
                </div>
                """, unsafe_allow_html=True)
            
            with col3:
                st.markdown(f"""
                <div class="kpi-blue">
                    <div style="font-size: 2rem;">📊</div>
                    <div style="font-size: 1.8rem; font-weight: 700;">{avg_safety:.0f}</div>
                    <div style="font-size: 0.85rem; opacity: 0.9;">平均安全库存(件)</div>
                </div>
                """, unsafe_allow_html=True)
            
            with col4:
                st.markdown(f"""
                <div class="kpi-green">
                    <div style="font-size: 2rem;">⏱️</div>
                    <div style="font-size: 1.8rem; font-weight: 700;">{coverage:.1f}</div>
                    <div style="font-size: 0.85rem; color: #1e3a5f;">平均库存覆盖倍数</div>
                </div>
                """, unsafe_allow_html=True)
            
            st.markdown("---")
            
            # ========== 功能卡片区 ==========
            st.markdown("### 🎯 核心功能")
            
            col_f1, col_f2, col_f3, col_f4 = st.columns(4)
            with col_f1:
                st.markdown("""
                <div class="feature-card">
                    <div style="font-size: 2rem;">📊</div>
                    <div style="font-weight: 600; margin: 10px 0;">需求特征分析</div>
                    <div style="font-size: 0.75rem; color: #6c757d;">自动计算CV、趋势、季节系数</div>
                </div>
                """, unsafe_allow_html=True)
            
            with col_f2:
                st.markdown("""
                <div class="feature-card">
                    <div style="font-size: 2rem;">📈</div>
                    <div style="font-weight: 600; margin: 10px 0;">需求预测</div>
                    <div style="font-size: 0.75rem; color: #6c757d;">多模型预测未来销量</div>
                </div>
                """, unsafe_allow_html=True)
            
            with col_f3:
                st.markdown("""
                <div class="feature-card">
                    <div style="font-size: 2rem;">🎯</div>
                    <div style="font-weight: 600; margin: 10px 0;">库存优化</div>
                    <div style="font-size: 0.75rem; color: #6c757d;">成本驱动模型智能求解</div>
                </div>
                """, unsafe_allow_html=True)
            
            with col_f4:
                st.markdown("""
                <div class="feature-card">
                    <div style="font-size: 2rem;">💡</div>
                    <div style="font-weight: 600; margin: 10px 0;">决策建议</div>
                    <div style="font-size: 0.75rem; color: #6c757d;">生成补货建议与行动清单</div>
                </div>
                """, unsafe_allow_html=True)
            
            st.markdown("---")
            
            # ========== 库存状态图表 ==========
            st.markdown("### 📈 库存监控看板")
            
            col_chart1, col_chart2 = st.columns(2)
            
            with col_chart1:
                status_counts = df_safety["库存状态"].value_counts()
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
                    fig_pie.update_layout(height=450)
                    st.plotly_chart(fig_pie, use_container_width=True)
            
            with col_chart2:
                # 安全库存分布
                safety_data = df_safety[df_safety["安全库存"] > 0]["安全库存"]
                if len(safety_data) > 0:
                    fig_hist = px.histogram(
                        safety_data, nbins=30,
                        title="安全库存值分布",
                        labels={"value": "安全库存 (件)", "count": "物料数量"},
                        color_discrete_sequence=["#3b82f6"],
                        opacity=0.7
                    )
                    fig_hist.update_layout(height=450)
                    st.plotly_chart(fig_hist, use_container_width=True)
            
            # 物料列表
            st.markdown("### 📋 物料安全库存列表")
            
            col_f1, col_f2 = st.columns(2)
            with col_f1:
                status_filter = st.multiselect(
                    "🔍 按状态筛选",
                    options=df_safety["库存状态"].unique(),
                    default=df_safety["库存状态"].unique(),
                    key="dashboard_filter"
                )
            with col_f2:
                search = st.text_input("🔎 搜索物料编码", placeholder="输入物料编码...", key="dashboard_search")
            
            filtered_df = df_safety[df_safety["库存状态"].isin(status_filter)]
            if search:
                filtered_df = filtered_df[filtered_df["物料编码"].astype(str).str.contains(search, na=False)]
            
            display_cols = ["物料编码", "安全库存", "实际库存", "库存状态", "缺货风险"]
            if "最近3个月月均用量" in filtered_df.columns:
                display_cols.append("最近3个月月均用量")
            available_cols = [c for c in display_cols if c in filtered_df.columns]
            
            def color_row(row):
                if row["库存状态"] == "🔴 低于安全库存":
                    return ["background-color: #fee2e2"] * len(row)
                elif row["库存状态"] == "🟡 临界状态":
                    return ["background-color: #fef3c7"] * len(row)
                return [""] * len(row)
            
            st.dataframe(
                filtered_df[available_cols].style.apply(color_row, axis=1),
                use_container_width=True,
                height=400
            )
            
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
            
            st.success("✅ 系统运行中 | 数据加载成功")
        else:
            st.warning("⚠️ 文件读取失败，请检查文件格式")
    else:
        # 未上传文件时的展示
        st.markdown("""
        <div class="upload-area">
            <div style="font-size: 3rem;">📁</div>
            <div style="font-size: 1.2rem; font-weight: 500; margin: 15px 0;">开始使用安全库存管理系统</div>
            <div style="color: #6c757d; margin-bottom: 15px;">请从左侧上传Excel或CSV文件</div>
            <div style="font-size: 0.8rem; color: #94a3b8;">支持 .xlsx, .xls, .csv 格式 | 自动识别月度数据</div>
        </div>
        """, unsafe_allow_html=True)
        
        # 系统亮点展示
        st.markdown("---")
        st.markdown("### ✨ 系统亮点")
        
        col_hl1, col_hl2, col_hl3 = st.columns(3)
        with col_hl1:
            st.markdown("""
            <div class="feature-card">
                <div style="font-size: 2rem;">🔄</div>
                <div style="font-weight: 600;">月度自动更新</div>
                <div style="font-size: 0.75rem;">每月更新数据，系统自动重新计算</div>
            </div>
            """, unsafe_allow_html=True)
        with col_hl2:
            st.markdown("""
            <div class="feature-card">
                <div style="font-size: 2rem;">🎯</div>
                <div style="font-weight: 600;">智能预警</div>
                <div style="font-size: 0.75rem;">自动识别低库存物料，提前预警</div>
            </div>
            """, unsafe_allow_html=True)
        with col_hl3:
            st.markdown("""
            <div class="feature-card">
                <div style="font-size: 2rem;">📊</div>
                <div style="font-weight: 600;">多维度分析</div>
                <div style="font-size: 0.75rem;">图表+表格，全面掌握库存状况</div>
            </div>
            """, unsafe_allow_html=True)

elif selected_menu == "📂 数据上传":
    st.markdown("### 📂 数据上传")
    st.markdown("请从左侧菜单点击上传文件，或直接拖拽文件到上传区域")
    
    if uploaded_file is not None:
        st.success(f"✅ 已上传文件：{uploaded_file.name}")
        st.info("💡 提示：上传成功后，请切换到「概览仪表盘」查看数据分析结果")

elif selected_menu == "📈 特征分析":
    st.markdown("### 📈 需求特征分析")
    st.markdown("自动计算CV、趋势、季节系数、IP影响系数等")
    if uploaded_file is not None:
        st.info("📊 特征分析功能开发中，请先上传数据")
    else:
        st.warning("⚠️ 请先上传数据文件")

elif selected_menu == "📉 需求预测":
    st.markdown("### 📉 需求预测")
    st.markdown("加权移动平均、霍尔特双参数、季节调整、IP缓冲、指数平滑+衰退、朴素批量")
    if uploaded_file is not None:
        st.info("📈 预测功能开发中，请先上传数据")
    else:
        st.warning("⚠️ 请先上传数据文件")

elif selected_menu == "🎯 库存优化":
    st.markdown("### 🎯 智能库存优化")
    st.markdown("成本驱动模型迭代求解、(R, Q)/(T, S)差异策略")
    if uploaded_file is not None:
        st.info("⚙️ 优化功能开发中，请先上传数据")
    else:
        st.warning("⚠️ 请先上传数据文件")

elif selected_menu == "💡 决策建议":
    st.markdown("### 💡 决策建议")
    if uploaded_file is not None:
        st.info("💡 根据库存状况生成补货建议...")
    else:
        st.warning("⚠️ 请先上传数据文件，系统将自动生成决策建议")
