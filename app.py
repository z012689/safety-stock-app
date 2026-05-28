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

# ==================== 页面设置 ====================
st.set_page_config(
    page_title="产品安全库存管理系统", 
    layout="wide", 
    page_icon="📦",
    initial_sidebar_state="expanded"
)

# ==================== 自定义CSS样式（丰富色彩） ====================
st.markdown("""
<style>
    /* 隐藏Streamlit默认的页脚和菜单 */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    
    /* 主标题样式 */
    .main-title {
        font-size: 2.5rem;
        font-weight: 700;
        background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        margin-bottom: 0.5rem;
    }
    
    /* 副标题样式 */
    .sub-title {
        font-size: 1rem;
        color: #6c757d;
        margin-bottom: 2rem;
    }
    
    /* 侧边栏样式 */
    .css-1d391kg {
        background: linear-gradient(180deg, #f8f9fa 0%, #e9ecef 100%);
    }
    
    /* 成功状态卡片 */
    .status-success {
        background-color: #d4edda;
        border-left: 4px solid #28a745;
        padding: 10px;
        border-radius: 8px;
    }
    
    /* 警告状态卡片 */
    .status-warning {
        background-color: #f8d7da;
        border-left: 4px solid #dc3545;
        padding: 10px;
        border-radius: 8px;
    }
    
    /* 信息卡片 */
    .info-card {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 20px;
        border-radius: 15px;
        color: white;
    }
    
    /* KPI卡片自定义 */
    .kpi-card {
        background: white;
        padding: 15px;
        border-radius: 12px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.05);
        transition: transform 0.2s;
    }
    .kpi-card:hover {
        transform: translateY(-3px);
        box-shadow: 0 6px 12px rgba(0,0,0,0.1);
    }
    
    /* 仪表盘装饰 */
    .dashboard-header {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 20px;
        border-radius: 15px;
        color: white;
        margin-bottom: 20px;
    }
</style>
""", unsafe_allow_html=True)

# ==================== 标题区域 ====================
col_title1, col_title2 = st.columns([3, 1])
with col_title1:
    st.markdown('<div class="main-title">📦 产品安全库存管理系统</div>', unsafe_allow_html=True)
    st.markdown('<div class="sub-title">支持月度数据更新 | 自动计算安全库存 | 智能预警 | 多维度分析</div>', unsafe_allow_html=True)

# ==================== 侧边栏 ====================
with st.sidebar:
    st.markdown("### 📂 数据上传")
    uploaded_file = st.file_uploader(
        "上传文件",
        type=["xlsx", "xls", "csv"],
        label_visibility="collapsed",
        help="支持 .xlsx, .xls, .csv 格式"
    )
    
    st.markdown("---")
    st.markdown("### 🎨 系统特色")
    st.markdown("""
    - 🔄 **月度自动更新**
    - 📊 **实时库存监控**
    - ⚠️ **智能风险预警**
    - 📈 **多维度数据分析**
    - 💾 **一键导出报表**
    """)
    
    st.markdown("---")
    st.markdown("### 📞 技术支持")
    st.markdown("如有问题，请联系系统管理员")

# ==================== 数据加载函数 ====================
@st.cache_data
def load_file_data(file):
    """加载上传的文件，自动识别格式"""
    file_name = file.name.lower()
    
    try:
        if file_name.endswith('.csv'):
            df = pd.read_csv(file, encoding='utf-8')
            if df.columns.str.contains('Unnamed').any():
                df = pd.read_csv(file, encoding='gbk')
        elif file_name.endswith(('.xlsx', '.xls')):
            # 尝试直接读取，跳过前2行
            df = pd.read_excel(file, sheet_name="安全库存（202509月）", header=2)
            
            # 检查是否成功读取到数据
            if df is not None and len(df) > 0:
                # 清理列名
                df.columns = [str(col).strip() for col in df.columns]
                return df
            
            # 备用方案：自动查找表头
            df_raw = pd.read_excel(file, sheet_name="安全库存（202509月）", header=None)
            for i, row in df_raw.iterrows():
                row_str = row.astype(str).str.contains("物料编码")
                if row_str.any():
                    df_raw.columns = df_raw.iloc[i]
                    df = df_raw.iloc[i+1:].reset_index(drop=True)
                    break
        else:
            return None
        return df
    except Exception as e:
        st.error(f"文件读取失败：{e}")
        return None

# ==================== 主程序 ====================
if uploaded_file is not None:
    with st.spinner("🔄 正在加载数据，请稍候..."):
        df = load_file_data(uploaded_file)
    
    if df is not None and len(df) > 0:
        # ========== 识别关键列 ==========
        # 物料编码列
        material_col = None
        for col in df.columns:
            if '物料编码' in str(col) or '物料代码' in str(col):
                material_col = col
                break
        if material_col is None:
            material_col = df.columns[0]
        
        # 安全库存列（支持多种列名）
        safety_col = None
        safety_keywords = ['安全库存', '安全库存（最终结果）', '安全库存(最终结果)']
        for col in df.columns:
            for kw in safety_keywords:
                if kw in str(col):
                    safety_col = col
                    break
            if safety_col:
                break
        if safety_col is None:
            # 尝试找包含"安全"和"库存"的列
            for col in df.columns:
                if '安全' in str(col) and '库存' in str(col):
                    safety_col = col
                    break
        
        # 实际库存列
        actual_col = None
        actual_keywords = ['实际库存', '6月末实际库存', '月末库存']
        for col in df.columns:
            for kw in actual_keywords:
                if kw in str(col):
                    actual_col = col
                    break
            if actual_col:
                break
        if actual_col is None:
            for col in df.columns:
                if '库存' in str(col) and col != safety_col:
                    actual_col = col
                    break
        
        # ========== 数据预处理 ==========
        if material_col:
            df = df.rename(columns={material_col: "物料编码"})
        
        if safety_col:
            df = df.rename(columns={safety_col: "安全库存"})
            df["安全库存"] = pd.to_numeric(df["安全库存"], errors="coerce").fillna(0)
        else:
            df["安全库存"] = 0
        
        if actual_col:
            df = df.rename(columns={actual_col: "实际库存"})
            df["实际库存"] = pd.to_numeric(df["实际库存"], errors="coerce").fillna(0)
        else:
            df["实际库存"] = 0
        
        # 删除空行
        df = df.dropna(subset=["物料编码"], how="all")
        df = df[df["物料编码"].notna()]
        df = df[df["物料编码"].astype(str).str.len() > 0]
        
        # 计算库存状态
        def get_status(row):
            if row["安全库存"] <= 0:
                return "⚪ 未设置"
            elif row["实际库存"] < row["安全库存"]:
                return "🔴 低于安全库存"
            elif row["实际库存"] <= row["安全库存"] * 1.2:
                return "🟡 临界状态"
            else:
                return "🟢 库存充足"
        
        df["库存状态"] = df.apply(get_status, axis=1)
        
        # 计算缺货风险百分比
        df["缺货风险"] = ((df["安全库存"] - df["实际库存"]) / df["安全库存"] * 100).clip(lower=0).round(1)
        df["缺货风险"] = df["缺货风险"].fillna(0).astype(int)
        
        # 计算库存健康度（实际/安全）
        df["库存健康度"] = (df["实际库存"] / df["安全库存"]).round(2)
        df["库存健康度"] = df["库存健康度"].replace([np.inf, -np.inf], 0).fillna(0)
        
        # ========== 统计摘要卡片（4列KPI） ==========
        col1, col2, col3, col4, col5 = st.columns(5)
        
        total = len(df)
        risk_count = len(df[df["库存状态"] == "🔴 低于安全库存"])
        warning_count = len(df[df["库存状态"] == "🟡 临界状态"])
        healthy_count = len(df[df["库存状态"] == "🟢 库存充足"])
        avg_safety = df["安全库存"].mean()
        
        with col1:
            st.markdown(f"""
            <div class="kpi-card" style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);">
                <div style="color: white; opacity: 0.8; font-size: 0.9rem;">📦 物料总数</div>
                <div style="color: white; font-size: 2rem; font-weight: 700;">{total}</div>
            </div>
            """, unsafe_allow_html=True)
        
        with col2:
            st.markdown(f"""
            <div class="kpi-card" style="background: linear-gradient(135deg, #f093fb 0%, #f5576c 100%);">
                <div style="color: white; opacity: 0.8; font-size: 0.9rem;">🔴 低库存预警</div>
                <div style="color: white; font-size: 2rem; font-weight: 700;">{risk_count}</div>
                <div style="color: white; font-size: 0.8rem;">占{round(risk_count/total*100)}%</div>
            </div>
            """, unsafe_allow_html=True)
        
        with col3:
            st.markdown(f"""
            <div class="kpi-card" style="background: linear-gradient(135deg, #fa709a 0%, #fee140 100%);">
                <div style="color: white; opacity: 0.8; font-size: 0.9rem;">🟡 临界状态</div>
                <div style="color: white; font-size: 2rem; font-weight: 700;">{warning_count}</div>
            </div>
            """, unsafe_allow_html=True)
        
        with col4:
            st.markdown(f"""
            <div class="kpi-card" style="background: linear-gradient(135deg, #4facfe 0%, #00f2fe 100%);">
                <div style="color: white; opacity: 0.8; font-size: 0.9rem;">🟢 库存充足</div>
                <div style="color: white; font-size: 2rem; font-weight: 700;">{healthy_count}</div>
            </div>
            """, unsafe_allow_html=True)
        
        with col5:
            st.markdown(f"""
            <div class="kpi-card" style="background: linear-gradient(135deg, #43e97b 0%, #38f9d7 100%);">
                <div style="color: white; opacity: 0.8; font-size: 0.9rem;">📊 平均安全库存</div>
                <div style="color: white; font-size: 1.5rem; font-weight: 700;">{avg_safety:.0f}</div>
                <div style="color: white; font-size: 0.7rem;">件</div>
            </div>
            """, unsafe_allow_html=True)
        
        st.markdown("---")
        
        # ========== 标签页 ==========
        tab1, tab2, tab3, tab4 = st.tabs(["📋 物料总览", "⚠️ 低库存预警", "📊 数据分析", "📈 趋势洞察"])
        
        # ---------- Tab 1: 物料总览 ----------
        with tab1:
            st.subheader("📋 物料安全库存一览")
            st.markdown("---")
            
            col_f1, col_f2, col_f3 = st.columns(3)
            with col_f1:
                status_filter = st.multiselect(
                    "🔍 按状态筛选",
                    options=df["库存状态"].unique(),
                    default=df["库存状态"].unique()
                )
            with col_f2:
                search = st.text_input("🔎 搜索物料编码", placeholder="输入物料编码...")
            with col_f3:
                sort_by = st.selectbox("📌 排序方式", ["缺货风险(高到低)", "安全库存(高到低)", "实际库存(高到低)"])
            
            filtered_df = df[df["库存状态"].isin(status_filter)]
            if search:
                filtered_df = filtered_df[filtered_df["物料编码"].astype(str).str.contains(search, na=False)]
            
            # 排序
            if sort_by == "缺货风险(高到低)":
                filtered_df = filtered_df.sort_values("缺货风险", ascending=False)
            elif sort_by == "安全库存(高到低)":
                filtered_df = filtered_df.sort_values("安全库存", ascending=False)
            elif sort_by == "实际库存(高到低)":
                filtered_df = filtered_df.sort_values("实际库存", ascending=False)
            
            display_cols = ["物料编码", "安全库存", "实际库存", "库存状态", "缺货风险", "库存健康度"]
            available_cols = [c for c in display_cols if c in filtered_df.columns]
            
            def color_status(val):
                if "🔴" in str(val):
                    return "background-color: #f8d7da; color: #721c24;"
                elif "🟡" in str(val):
                    return "background-color: #fff3cd; color: #856404;"
                elif "🟢" in str(val):
                    return "background-color: #d4edda; color: #155724;"
                return ""
            
            styled_df = filtered_df[available_cols].style.applymap(color_status, subset=["库存状态"])
            
            st.dataframe(styled_df, use_container_width=True, height=500)
            
            # 导出
            output = BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                filtered_df.to_excel(writer, index=False, sheet_name="安全库存明细")
            st.download_button(
                label="📎 导出数据",
                data=output.getvalue(),
                file_name="安全库存明细.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        
        # ---------- Tab 2: 低库存预警 ----------
        with tab2:
            st.subheader("⚠️ 低于安全库存的物料清单")
            st.markdown("以下物料库存不足，建议尽快安排补货")
            
            alert_df = df[df["库存状态"] == "🔴 低于安全库存"].copy()
            if len(alert_df) > 0:
                alert_df = alert_df.sort_values("缺货风险", ascending=False)
                alert_cols = ["物料编码", "实际库存", "安全库存", "缺货风险", "库存健康度"]
                available_alert = [c for c in alert_cols if c in alert_df.columns]
                
                st.dataframe(alert_df[available_alert], use_container_width=True)
                
                # 风险TOP10柱状图
                if len(alert_df) > 0:
                    top10 = alert_df.head(10)
                    fig = px.bar(
                        top10, x="物料编码", y="缺货风险",
                        title="🔴 缺货风险最高的10种物料",
                        labels={"缺货风险": "缺货风险 (%)"},
                        color="缺货风险", color_continuous_scale="reds",
                        text="缺货风险"
                    )
                    fig.update_traces(textposition="outside")
                    fig.update_layout(height=450)
                    st.plotly_chart(fig, use_container_width=True)
                    
                    # 建议补货量
                    st.subheader("📦 建议补货量")
                    alert_df["建议补货量"] = (alert_df["安全库存"] - alert_df["实际库存"]).round(0).astype(int)
                    st.dataframe(alert_df[["物料编码", "实际库存", "安全库存", "建议补货量"]].head(20), use_container_width=True)
            else:
                st.success("🎉 所有物料库存充足！")
        
        # ---------- Tab 3: 数据分析 ----------
        with tab3:
            st.subheader("📊 库存状态分布")
            
            col_ch1, col_ch2 = st.columns(2)
            
            with col_ch1:
                status_counts = df["库存状态"].value_counts()
                if len(status_counts) > 0:
                    colors = ["#e74c3c", "#f39c12", "#2ecc71", "#95a5a6"]
                    fig_pie = px.pie(
                        values=status_counts.values,
                        names=status_counts.index,
                        title="库存状态占比",
                        color_discrete_sequence=colors,
                        hole=0.4
                    )
                    fig_pie.update_traces(textposition="inside", textinfo="percent+label")
                    fig_pie.update_layout(height=450)
                    st.plotly_chart(fig_pie, use_container_width=True)
            
            with col_ch2:
                safety_data = df[df["安全库存"] > 0]["安全库存"]
                if len(safety_data) > 0:
                    fig_hist = px.histogram(
                        safety_data, nbins=30,
                        title="安全库存值分布",
                        labels={"value": "安全库存", "count": "物料数量"},
                        color_discrete_sequence=["#3498db"],
                        opacity=0.7
                    )
                    fig_hist.update_layout(height=450)
                    st.plotly_chart(fig_hist, use_container_width=True)
            
            # 库存健康度分布
            st.subheader("💚 库存健康度分布")
            health_data = df[df["库存健康度"] > 0]["库存健康度"]
            if len(health_data) > 0:
                fig_health = px.box(
                    health_data,
                    title="库存覆盖倍数分布",
                    labels={"value": "库存覆盖倍数（实际/安全）"},
                    color_discrete_sequence=["#27ae60"]
                )
                fig_health.update_layout(height=400)
                st.plotly_chart(fig_health, use_container_width=True)
            
            # 实际库存 vs 安全库存对比
            st.subheader("📊 实际库存 vs 安全库存对比（前30个物料）")
            compare_df = df[["物料编码", "实际库存", "安全库存"]].dropna().head(30)
            if len(compare_df) > 0:
                fig_compare = go.Figure()
                fig_compare.add_trace(go.Bar(x=compare_df["物料编码"], y=compare_df["实际库存"], name="实际库存", marker_color="#2ecc71"))
                fig_compare.add_trace(go.Bar(x=compare_df["物料编码"], y=compare_df["安全库存"], name="安全库存", marker_color="#e74c3c"))
                fig_compare.update_layout(
                    barmode="group",
                    title="实际库存 vs 安全库存",
                    xaxis_title="物料编码",
                    yaxis_title="库存数量",
                    height=500
                )
                st.plotly_chart(fig_compare, use_container_width=True)
        
        # ---------- Tab 4: 趋势洞察 ----------
        with tab4:
            st.subheader("📈 库存洞察分析")
            
            # 风险等级分布
            col_ins1, col_ins2 = st.columns(2)
            
            with col_ins1:
                # 按风险等级分组
                risk_levels = []
                for _, row in df.iterrows():
                    if row["缺货风险"] == 0:
                        risk_levels.append("无风险")
                    elif row["缺货风险"] < 30:
                        risk_levels.append("低风险")
                    elif row["缺货风险"] < 60:
                        risk_levels.append("中风险")
                    else:
                        risk_levels.append("高风险")
                df["风险等级"] = risk_levels
                
                risk_counts = df["风险等级"].value_counts()
                if len(risk_counts) > 0:
                    fig_risk = px.bar(
                        x=risk_counts.index, y=risk_counts.values,
                        title="风险等级分布",
                        labels={"x": "风险等级", "y": "物料数量"},
                        color=risk_counts.index,
                        color_discrete_sequence=["#2ecc71", "#f39c12", "#e67e22", "#e74c3c"]
                    )
                    fig_risk.update_layout(height=400)
                    st.plotly_chart(fig_risk, use_container_width=True)
            
            with col_ins2:
                # 库存集中度
                st.markdown("### 📊 库存集中度分析")
                top10_stock = df.nlargest(10, "安全库存")[["物料编码", "安全库存"]]
                fig_top = px.bar(
                    top10_stock, x="物料编码", y="安全库存",
                    title="安全库存 Top 10",
                    color="安全库存", color_continuous_scale="blues",
                    text="安全库存"
                )
                fig_top.update_traces(textposition="outside")
                fig_top.update_layout(height=400)
                st.plotly_chart(fig_top, use_container_width=True)
        
        st.success("✅ 系统运行中 | 数据加载成功")
    else:
        st.error("❌ 数据处理失败，请检查Excel文件格式")
else:
    st.info("👈 请从左侧上传文件开始使用")
