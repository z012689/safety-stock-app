"""
产品安全库存管理系统
功能：支持多种文件格式（Excel/CSV），自动计算安全库存，月度更新
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO
import re

# ==================== 页面设置 ====================
st.set_page_config(
    page_title="产品安全库存管理系统", 
    layout="wide", 
    page_icon="📦",
    initial_sidebar_state="expanded"
)

# ==================== 蓝色主题CSS样式 ====================
st.markdown("""
<style>
    /* 隐藏默认元素 */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    
    /* 主标题 */
    .main-title {
        font-size: 2.5rem;
        font-weight: 700;
        background: linear-gradient(135deg, #1a56db 0%, #3b82f6 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        margin-bottom: 0.5rem;
    }
    
    /* 副标题 */
    .sub-title {
        font-size: 1rem;
        color: #6c757d;
        margin-bottom: 2rem;
    }
    
    /* 蓝色卡片 */
    .blue-card {
        background: linear-gradient(135deg, #1e3a8a 0%, #3b82f6 100%);
        padding: 20px;
        border-radius: 15px;
        color: white;
        box-shadow: 0 4px 15px rgba(59,130,246,0.3);
    }
    
    /* 浅蓝色信息卡片 */
    .info-card {
        background-color: #eff6ff;
        border-left: 4px solid #3b82f6;
        padding: 15px;
        border-radius: 10px;
        margin: 10px 0;
    }
    
    /* KPI卡片样式 */
    .kpi-card {
        background: white;
        padding: 15px;
        border-radius: 12px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.05);
        transition: transform 0.2s;
        border-top: 3px solid #3b82f6;
    }
    .kpi-card:hover {
        transform: translateY(-3px);
        box-shadow: 0 6px 12px rgba(0,0,0,0.1);
    }
    
    /* 侧边栏样式 */
    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #f0f9ff 0%, #e0f2fe 100%);
    }
    
    /* 按钮样式 */
    .stButton > button {
        background-color: #3b82f6;
        color: white;
        border-radius: 8px;
        border: none;
    }
    .stButton > button:hover {
        background-color: #2563eb;
    }
    
    /* 数据表格表头样式 */
    .stDataFrame thead th {
        background-color: #3b82f6;
        color: white;
        font-weight: 600;
    }
</style>
""", unsafe_allow_html=True)

# ==================== 标题区域 ====================
st.markdown('<div class="main-title">📦 产品安全库存管理系统</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">支持月度数据更新 | 自动计算安全库存 | 智能预警 | 支持Excel/CSV</div>', unsafe_allow_html=True)

# ==================== 侧边栏 ====================
with st.sidebar:
    st.markdown("### 📂 数据上传")
    
    # 支持多种文件类型
    uploaded_file = st.file_uploader(
        "上传文件",
        type=["xlsx", "xls", "csv"],
        label_visibility="collapsed",
        help="支持 .xlsx, .xls, .csv 格式"
    )
    
    st.markdown("---")
    st.markdown("### 🎯 系统功能")
    st.markdown("""
    - 🔄 **月度自动更新**
    - 📊 **实时库存监控**
    - ⚠️ **智能风险预警**
    - 📈 **多维度数据分析**
    - 💾 **一键导出报表**
    """)
    
    st.markdown("---")
    st.markdown("### 📋 文件要求")
    st.markdown("""
    文件需包含以下列：
    - **物料编码** / 物料代码
    - **安全库存** / 安全库存（最终结果）
    - **实际库存** / 6月末实际库存
    """)

# ==================== 数据加载函数（支持多种格式） ====================
@st.cache_data
def load_safety_data(file):
    """加载安全库存数据，支持Excel和CSV"""
    file_name = file.name.lower()
    
    try:
        if file_name.endswith('.csv'):
            # CSV文件
            df = pd.read_csv(file, encoding='utf-8')
            if df.columns.str.contains('Unnamed').any():
                df = pd.read_csv(file, encoding='gbk')
            return df
        else:
            # Excel文件
            df = pd.read_excel(file, sheet_name="安全库存（202509月）", header=None)
            
            # 寻找包含"物料编码"的行作为列名
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
                # 备用方案：直接读取
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
        # 寻找包含"建议安全库存"的行
        for i, row in df.iterrows():
            if "建议安全库存" in row.astype(str).values:
                df.columns = df.iloc[i]
                df = df.iloc[i+1:].reset_index(drop=True)
                break
        return df
    except:
        return None

# ==================== 主程序 ====================
if uploaded_file is not None:
    with st.spinner("🔄 正在加载数据，请稍候..."):
        df_safety = load_safety_data(uploaded_file)
        df_packaging = load_packaging_data(uploaded_file) if uploaded_file.name.lower().endswith(('xlsx', 'xls')) else None
    
    if df_safety is not None and len(df_safety) > 0:
        # 清理列名（去除换行符和空格）
        df_safety.columns = [str(col).replace('\n', '').strip() for col in df_safety.columns]
        
        # ========== 智能识别关键列 ==========
        # 物料编码列
        material_col = None
        for col in df_safety.columns:
            if '物料编码' in str(col) or '物料代码' in str(col):
                material_col = col
                break
        if material_col is None:
            material_col = df_safety.columns[0]
        
        # 安全库存列
        safety_col = None
        for col in df_safety.columns:
            if '安全库存' in str(col):
                safety_col = col
                break
        
        # 实际库存列
        actual_col = None
        for col in df_safety.columns:
            if '实际库存' in str(col) or '6月末实际库存' in str(col) or '月末库存' in str(col):
                actual_col = col
                break
        
        # 重命名列
        df_safety = df_safety.rename(columns={material_col: "物料编码"})
        if safety_col:
            df_safety = df_safety.rename(columns={safety_col: "安全库存"})
        if actual_col:
            df_safety = df_safety.rename(columns={actual_col: "实际库存"})
        
        # 转换数值列
        for col in ["安全库存", "实际库存"]:
            if col in df_safety.columns:
                df_safety[col] = pd.to_numeric(df_safety[col], errors="coerce").fillna(0)
            else:
                df_safety[col] = 0
        
        # 删除空行
        df_safety = df_safety.dropna(subset=["物料编码"], how="all")
        df_safety = df_safety[df_safety["物料编码"].notna()]
        df_safety = df_safety[df_safety["物料编码"].astype(str).str.len() > 0]
        
        # ========== 自动识别月份列并计算月均用量 ==========
        month_pattern = re.compile(r'(\d{4})[年M](\d{1,2})|M(\d{1,2})')
        month_cols = []
        for col in df_safety.columns:
            if month_pattern.search(str(col)):
                month_cols.append(col)
        
        if len(month_cols) >= 3:
            # 取最新的3个月
            latest_months = month_cols[-3:]
            df_safety["最近3个月月均用量"] = df_safety[latest_months].mean(axis=1).round(0)
            st.info(f"📊 自动识别到 {len(month_cols)} 个月份列，使用最新3个月计算月均用量")
        
        # ========== 计算库存状态 ==========
        df_safety["库存状态"] = df_safety.apply(
            lambda row: "🔴 低于安全库存" if row["实际库存"] < row["安全库存"] and row["安全库存"] > 0 
            else "🟡 临界状态" if row["实际库存"] <= row["安全库存"] * 1.2 and row["安全库存"] > 0
            else "🟢 库存充足" if row["安全库存"] > 0
            else "⚪ 未设置",
            axis=1
        )
        
        # 计算缺货风险百分比
        df_safety["缺货风险"] = ((df_safety["安全库存"] - df_safety["实际库存"]) / df_safety["安全库存"] * 100).clip(lower=0).round(1)
        df_safety["缺货风险"] = df_safety["缺货风险"].fillna(0).astype(int)
        
        # 计算库存健康度
        df_safety["库存健康度"] = (df_safety["实际库存"] / df_safety["安全库存"]).round(2)
        df_safety["库存健康度"] = df_safety["库存健康度"].replace([np.inf, -np.inf], 0).fillna(0)
        
        # ========== KPI 卡片（蓝色主题） ==========
        col1, col2, col3, col4 = st.columns(4)
        
        total = len(df_safety)
        risk = len(df_safety[df_safety["库存状态"] == "🔴 低于安全库存"])
        avg_safety = df_safety["安全库存"].mean() if "安全库存" in df_safety else 0
        valid = df_safety[df_safety["安全库存"] > 0]
        coverage = (valid["实际库存"] / valid["安全库存"]).mean() if len(valid) > 0 else 0
        
        with col1:
            st.metric("📦 物料总数", total, delta=None)
        with col2:
            st.metric("⚠️ 低库存预警", risk, delta=f"占{round(risk/total*100)}%" if total>0 else "0")
        with col3:
            st.metric("📊 平均安全库存", f"{avg_safety:.0f} 件")
        with col4:
            st.metric("⏱️ 库存覆盖倍数", f"{coverage:.1f} 倍")
        
        st.divider()
        
        # ========== 标签页 ==========
        tab1, tab2, tab3, tab4 = st.tabs(["📋 物料总览", "⚠️ 低库存预警", "📊 数据分析", "📦 包材安全库存"])
        
        # ---------- Tab 1: 物料总览 ----------
        with tab1:
            st.subheader("📋 所有物料安全库存一览")
            
            col_f1, col_f2 = st.columns(2)
            with col_f1:
                status_filter = st.multiselect(
                    "🔍 库存状态筛选",
                    options=df_safety["库存状态"].unique(),
                    default=df_safety["库存状态"].unique()
                )
            with col_f2:
                search = st.text_input("🔎 搜索物料编码", placeholder="输入物料编码...")
            
            filtered_df = df_safety[df_safety["库存状态"].isin(status_filter)]
            if search:
                filtered_df = filtered_df[filtered_df["物料编码"].astype(str).str.contains(search, na=False)]
            
            display_cols = ["物料编码", "安全库存", "实际库存", "库存状态", "缺货风险", "库存健康度"]
            if "最近3个月月均用量" in filtered_df.columns:
                display_cols.append("最近3个月月均用量")
            available_cols = [c for c in display_cols if c in filtered_df.columns]
            
            def highlight_risk(row):
                if row["库存状态"] == "🔴 低于安全库存":
                    return ["background-color: #fee2e2"] * len(row)
                elif row["库存状态"] == "🟡 临界状态":
                    return ["background-color: #fef3c7"] * len(row)
                return [""] * len(row)
            
            st.dataframe(
                filtered_df[available_cols].style.apply(highlight_risk, axis=1),
                use_container_width=True,
                height=500
            )
            
            # 导出功能
            output = BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                filtered_df.to_excel(writer, index=False, sheet_name="安全库存明细")
            st.download_button(
                label="📎 导出当前数据",
                data=output.getvalue(),
                file_name="安全库存明细.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        
        # ---------- Tab 2: 低库存预警 ----------
        with tab2:
            st.subheader("⚠️ 低于安全库存的物料清单")
            st.markdown("以下物料库存不足，建议尽快安排补货")
            
            alert_df = df_safety[df_safety["库存状态"] == "🔴 低于安全库存"].copy()
            if len(alert_df) > 0:
                alert_df = alert_df.sort_values("缺货风险", ascending=False)
                alert_cols = ["物料编码", "实际库存", "安全库存", "缺货风险", "库存健康度"]
                if "最近3个月月均用量" in alert_df.columns:
                    alert_cols.append("最近3个月月均用量")
                available_alert = [c for c in alert_cols if c in alert_df.columns]
                st.dataframe(alert_df[available_alert], use_container_width=True)
                
                if len(alert_df) > 0:
                    top10 = alert_df.head(10)
                    fig = px.bar(
                        top10, x="物料编码", y="缺货风险",
                        title="🔴 缺货风险最高的10种物料",
                        labels={"缺货风险": "缺货风险 (%)"},
                        color="缺货风险", color_continuous_scale="Reds",
                        text="缺货风险"
                    )
                    fig.update_traces(textposition="outside")
                    fig.update_layout(height=450, title_font_color="#1e3a8a")
                    st.plotly_chart(fig, use_container_width=True)
                    
                    # 建议补货量
                    alert_df["建议补货量"] = (alert_df["安全库存"] - alert_df["实际库存"]).round(0).astype(int)
                    st.subheader("📦 建议补货量")
                    st.dataframe(alert_df[["物料编码", "实际库存", "安全库存", "建议补货量"]].head(20), use_container_width=True)
            else:
                st.success("🎉 所有物料库存充足！")
        
        # ---------- Tab 3: 数据分析 ----------
        with tab3:
            st.subheader("📊 库存状态分布")
            
            col_ch1, col_ch2 = st.columns(2)
            
            with col_ch1:
                status_counts = df_safety["库存状态"].value_counts()
                if len(status_counts) > 0:
                    fig_pie = px.pie(
                        values=status_counts.values,
                        names=status_counts.index,
                        title="库存状态占比",
                        color_discrete_sequence=["#ef4444", "#f59e0b", "#10b981", "#94a3b8"],
                        hole=0.4
                    )
                    fig_pie.update_traces(textposition="inside", textinfo="percent+label")
                    fig_pie.update_layout(height=450)
                    st.plotly_chart(fig_pie, use_container_width=True)
            
            with col_ch2:
                if "安全库存" in df_safety.columns:
                    safety_data = df_safety[df_safety["安全库存"] > 0]["安全库存"]
                    if len(safety_data) > 0:
                        fig_hist = px.histogram(
                            safety_data, nbins=30,
                            title="安全库存值分布",
                            labels={"value": "安全库存", "count": "物料数量"},
                            color_discrete_sequence=["#3b82f6"],
                            opacity=0.7
                        )
                        fig_hist.update_layout(height=450)
                        st.plotly_chart(fig_hist, use_container_width=True)
            
            # 库存健康度分布
            st.subheader("💚 库存健康度分布")
            health_data = df_safety[df_safety["库存健康度"] > 0]["库存健康度"]
            if len(health_data) > 0:
                fig_health = px.box(
                    health_data,
                    title="库存覆盖倍数分布",
                    labels={"value": "库存覆盖倍数（实际/安全）"},
                    color_discrete_sequence=["#3b82f6"]
                )
                fig_health.update_layout(height=400)
                st.plotly_chart(fig_health, use_container_width=True)
            
            # 实际库存 vs 安全库存对比
            st.subheader("📊 实际库存 vs 安全库存对比（前30个物料）")
            compare_df = df_safety[["物料编码", "实际库存", "安全库存"]].dropna().head(30)
            if len(compare_df) > 0:
                fig_compare = go.Figure()
                fig_compare.add_trace(go.Bar(x=compare_df["物料编码"], y=compare_df["实际库存"], name="实际库存", marker_color="#10b981"))
                fig_compare.add_trace(go.Bar(x=compare_df["物料编码"], y=compare_df["安全库存"], name="安全库存", marker_color="#3b82f6"))
                fig_compare.update_layout(
                    barmode="group",
                    title="实际库存 vs 安全库存",
                    xaxis_title="物料编码",
                    yaxis_title="库存数量",
                    height=500,
                    title_font_color="#1e3a8a"
                )
                st.plotly_chart(fig_compare, use_container_width=True)
        
        # ---------- Tab 4: 包材安全库存 ----------
        with tab4:
            st.subheader("📦 包材安全库存列表")
            if df_packaging is not None and len(df_packaging) > 0:
                if "建议安全库存" in df_packaging.columns:
                    pack_cols = ["物料编码", "建议安全库存", "品类"] if "品类" in df_packaging.columns else ["物料编码", "建议安全库存"]
                    available_pack = [c for c in pack_cols if c in df_packaging.columns]
                    st.dataframe(df_packaging[available_pack], use_container_width=True, height=400)
                    
                    if "品类" in df_packaging.columns:
                        pack_cat = df_packaging["品类"].value_counts().reset_index()
                        pack_cat.columns = ["品类", "数量"]
                        fig_pack = px.bar(pack_cat, x="品类", y="数量", title="包材品类分布", color="品类", color_discrete_sequence=px.colors.qualitative.Set2)
                        fig_pack.update_layout(height=400)
                        st.plotly_chart(fig_pack, use_container_width=True)
                else:
                    st.info("包材数据格式已识别，但未找到'建议安全库存'列")
            else:
                st.info("当前文件未包含包材数据（仅Excel文件支持包材分析）")
        
        st.success("✅ 系统运行中 | 数据加载成功")
    else:
        st.error("❌ 数据处理失败，请检查文件格式")
else:
    # 空状态提示
    st.info("👈 请从左侧上传文件开始使用")
