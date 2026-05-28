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

# 自定义CSS样式（美化界面）
st.markdown("""
<style>
    /* 隐藏Streamlit默认的页脚和菜单 */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    
    /* 卡片样式 */
    .stMetric {
        background-color: #f8f9fa;
        padding: 15px;
        border-radius: 10px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
    }
    
    /* 标题样式 */
    .main-title {
        font-size: 2.5rem;
        font-weight: 600;
        color: #1f3e6e;
        margin-bottom: 0.5rem;
    }
    
    /* 副标题样式 */
    .sub-title {
        font-size: 1rem;
        color: #6c757d;
        margin-bottom: 2rem;
    }
</style>
""", unsafe_allow_html=True)

# 主标题
st.markdown('<div class="main-title">📦 产品安全库存管理系统</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">支持月度数据更新 | 自动计算安全库存 | 智能预警</div>', unsafe_allow_html=True)

# ==================== 侧边栏 ====================
with st.sidebar:
    st.markdown("### 📂 数据上传")
    uploaded_file = st.file_uploader(
        "上传文件",
        type=["xlsx", "xls", "csv"],
        label_visibility="collapsed",
        help="支持 .xlsx, .xls, .csv 格式"
    )

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
            df = pd.read_excel(file, header=None)
            
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
                df = pd.read_excel(file, header=2)
        else:
            return None
        return df
    except Exception as e:
        st.error(f"文件读取失败：{e}")
        return None


@st.cache_data
def find_month_columns(df):
    """自动识别月份列"""
    month_patterns = [
        r'(\d{4})年M(\d{1,2})',
        r'(\d{4})M(\d{1,2})',
        r'(\d{4})-(\d{1,2})',
        r'(\d{4})/(\d{1,2})',
    ]
    
    month_columns = []
    for col in df.columns:
        col_str = str(col)
        for pattern in month_patterns:
            if re.search(pattern, col_str):
                month_columns.append(col)
                break
    return month_columns


# ==================== 主程序 ====================
if uploaded_file is not None:
    with st.spinner("正在加载数据..."):
        df = load_file_data(uploaded_file)
    
    if df is not None and len(df) > 0:
        # 识别关键列
        material_col = None
        for col in df.columns:
            if '物料编码' in str(col) or '物料代码' in str(col):
                material_col = col
                break
        
        safety_col = None
        for col in df.columns:
            if '安全库存' in str(col):
                safety_col = col
                break
        
        actual_col = None
        for col in df.columns:
            if '实际库存' in str(col) or '月末库存' in str(col):
                actual_col = col
                break
        
        if material_col and safety_col:
            df = df.rename(columns={material_col: "物料编码", safety_col: "安全库存"})
            if actual_col:
                df = df.rename(columns={actual_col: "实际库存"})
            
            for col in ["安全库存", "实际库存"]:
                if col in df.columns:
                    df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
            
            # 处理没有实际库存列的情况
            if "实际库存" not in df.columns:
                df["实际库存"] = 0
            
            # 计算库存状态
            df["库存状态"] = df.apply(
                lambda row: "⚠️ 低于安全库存" if row["实际库存"] < row["安全库存"] and row["安全库存"] > 0 else "✅ 库存充足",
                axis=1
            )
            
            df["缺货风险"] = ((df["安全库存"] - df["实际库存"]) / df["安全库存"] * 100).clip(lower=0).round(1)
            df["缺货风险"] = df["缺货风险"].fillna(0).astype(int)
            
            # 月份识别
            month_cols = find_month_columns(df)
            if len(month_cols) >= 3:
                latest_months = month_cols[-3:]
                df["最近3个月月均用量"] = df[latest_months].mean(axis=1).round(0)
            
            # ========== KPI 卡片 ==========
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("📦 物料总数", len(df))
            with col2:
                risk_count = len(df[df["库存状态"] == "⚠️ 低于安全库存"])
                st.metric("⚠️ 低库存预警", risk_count, delta=f"{round(risk_count/len(df)*100)}%")
            with col3:
                st.metric("📊 平均安全库存", f"{df['安全库存'].mean():.0f} 件")
            with col4:
                valid = df[df["安全库存"] > 0]
                coverage = (valid["实际库存"] / valid["安全库存"]).mean() if len(valid) > 0 else 0
                st.metric("⏱️ 库存覆盖倍数", f"{coverage:.1f} 倍")
            
            st.divider()
            
            # ========== 标签页 ==========
            tab1, tab2, tab3 = st.tabs(["📋 物料总览", "⚠️ 低库存预警", "📊 数据分析"])
            
            # Tab 1: 物料总览
            with tab1:
                col_f1, col_f2 = st.columns(2)
                with col_f1:
                    status_filter = st.multiselect(
                        "状态筛选",
                        options=df["库存状态"].unique(),
                        default=df["库存状态"].unique()
                    )
                with col_f2:
                    search = st.text_input("🔍 搜索物料", placeholder="输入物料编码...")
                
                filtered_df = df[df["库存状态"].isin(status_filter)]
                if search:
                    filtered_df = filtered_df[filtered_df["物料编码"].astype(str).str.contains(search, na=False)]
                
                display_cols = ["物料编码", "安全库存", "实际库存", "库存状态", "缺货风险"]
                if "最近3个月月均用量" in filtered_df.columns:
                    display_cols.append("最近3个月月均用量")
                available_cols = [c for c in display_cols if c in filtered_df.columns]
                
                def highlight_risk(row):
                    if row["库存状态"] == "⚠️ 低于安全库存":
                        return ["background-color: #ffe6e6"] * len(row)
                    return [""] * len(row)
                
                st.dataframe(
                    filtered_df[available_cols].style.apply(highlight_risk, axis=1),
                    use_container_width=True,
                    height=450
                )
                
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
            
            # Tab 2: 低库存预警
            with tab2:
                alert_df = df[df["库存状态"] == "⚠️ 低于安全库存"].copy()
                if len(alert_df) > 0:
                    alert_df = alert_df.sort_values("缺货风险", ascending=False)
                    alert_cols = ["物料编码", "实际库存", "安全库存", "缺货风险"]
                    if "最近3个月月均用量" in alert_df.columns:
                        alert_cols.append("最近3个月月均用量")
                    available_alert = [c for c in alert_cols if c in alert_df.columns]
                    st.dataframe(alert_df[available_alert], use_container_width=True)
                    
                    if len(alert_df) > 0:
                        top10 = alert_df.head(10)
                        fig = px.bar(
                            top10, x="物料编码", y="缺货风险",
                            title="缺货风险最高的10种物料",
                            labels={"缺货风险": "缺货风险 (%)"},
                            color="缺货风险", color_continuous_scale="reds"
                        )
                        fig.update_layout(height=400)
                        st.plotly_chart(fig, use_container_width=True)
                else:
                    st.success("🎉 所有物料库存充足！")
            
            # Tab 3: 数据分析
            with tab3:
                c1, c2 = st.columns(2)
                with c1:
                    status_counts = df["库存状态"].value_counts()
                    if len(status_counts) > 0:
                        fig_pie = px.pie(
                            values=status_counts.values,
                            names=status_counts.index,
                            title="库存状态分布",
                            color_discrete_sequence=["#2ecc71", "#e74c3c"]
                        )
                        fig_pie.update_layout(height=400)
                        st.plotly_chart(fig_pie, use_container_width=True)
                
                with c2:
                    safety_data = df[df["安全库存"] > 0]["安全库存"]
                    if len(safety_data) > 0:
                        fig_hist = px.histogram(
                            safety_data, nbins=30,
                            title="安全库存值分布",
                            labels={"value": "安全库存", "count": "物料数量"},
                            color_discrete_sequence=["#3498db"]
                        )
                        fig_hist.update_layout(height=400)
                        st.plotly_chart(fig_hist, use_container_width=True)
                
                # 对比图
                compare_df = df[["物料编码", "实际库存", "安全库存"]].dropna().head(30)
                if len(compare_df) > 0:
                    fig_compare = go.Figure()
                    fig_compare.add_trace(go.Bar(x=compare_df["物料编码"], y=compare_df["实际库存"], name="实际库存", marker_color="#2ecc71"))
                    fig_compare.add_trace(go.Bar(x=compare_df["物料编码"], y=compare_df["安全库存"], name="安全库存", marker_color="#e74c3c"))
                    fig_compare.update_layout(
                        barmode="group",
                        title="实际库存 vs 安全库存（前30项）",
                        xaxis_title="物料编码",
                        yaxis_title="库存数量",
                        height=450
                    )
                    st.plotly_chart(fig_compare, use_container_width=True)
            
            st.success("✅ 系统运行中 | 数据加载成功")
        else:
            st.warning("⚠️ 未能识别关键列，请确保文件包含：物料编码、安全库存")
    else:
        st.error("数据处理失败，请检查文件格式")
else:
    st.info("👈 请从左侧上传文件开始使用")
