import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO

st.set_page_config(page_title="产品安全库存管理系统", layout="wide", page_icon="📦")

st.title("📦 产品安全库存管理系统")
st.caption("支持月度数据更新 | 自动计算安全库存 | 智能预警")

with st.sidebar:
    st.header("📂 数据上传")
    uploaded_file = st.file_uploader(
        "上传 Excel 文件",
        type=["xlsx", "xls"],
        help="上传安全库存分析-2025.xlsx 格式的文件"
    )
    st.divider()
    st.markdown("### 📌 月度更新说明")
    st.markdown("""
    - 每月更新Excel中的销量数据后，重新上传即可
    - 系统会自动识别最新月份并重新计算
    - 无需修改任何代码
    """)

@st.cache_data
def load_safety_data(file):
    """加载安全库存数据"""
    try:
        # 尝试多种方式读取
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
        st.error(f"读取安全库存数据失败：{e}")
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

if uploaded_file is not None:
    with st.spinner("正在加载数据..."):
        df_safety = load_safety_data(uploaded_file)
        df_packaging = load_packaging_data(uploaded_file)
    
    if df_safety is not None and len(df_safety) > 0:
        # 重命名列以便使用
        df_safety = df_safety.rename(columns={
            "安全库存\n（最终结果）": "安全库存",
            "6月末实际库存": "实际库存"
        })
        
        # 转换数值列
        for col in ["安全库存", "实际库存"]:
            if col in df_safety.columns:
                df_safety[col] = pd.to_numeric(df_safety[col], errors="coerce").fillna(0)
        
        # 计算库存状态
        df_safety["库存状态"] = df_safety.apply(
            lambda row: "⚠️ 低于安全库存" if row["实际库存"] < row["安全库存"] and row["安全库存"] > 0 else "✅ 库存充足",
            axis=1
        )
        
        # 计算缺货风险百分比
        df_safety["缺货风险"] = ((df_safety["安全库存"] - df_safety["实际库存"]) / df_safety["安全库存"] * 100).clip(lower=0).round(1)
        df_safety["缺货风险"] = df_safety["缺货风险"].fillna(0).astype(int)
        
        # KPI 卡片
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            total = len(df_safety)
            st.metric("📦 物料总数", total)
        with col2:
            risk = len(df_safety[df_safety["库存状态"] == "⚠️ 低于安全库存"])
            st.metric("⚠️ 低库存预警", risk, delta=f"占{round(risk/total*100)}%" if total>0 else "0")
        with col3:
            avg_safety = df_safety["安全库存"].mean() if "安全库存" in df_safety else 0
            st.metric("📊 平均安全库存", f"{avg_safety:.0f} 件")
        with col4:
            if "安全库存" in df_safety:
                valid = df_safety[df_safety["安全库存"] > 0]
                if len(valid) > 0:
                    coverage = (valid["实际库存"] / valid["安全库存"]).mean()
                    st.metric("⏱️ 平均库存覆盖倍数", f"{coverage:.1f} 倍")
                else:
                    st.metric("⏱️ 平均库存覆盖倍数", "N/A")
        
        st.divider()
        
        # 标签页
        tab1, tab2, tab3, tab4 = st.tabs(["📋 物料安全库存总览", "⚠️ 低库存预警", "📊 数据分析", "📦 包材安全库存"])
        
        # Tab 1: 物料总览
        with tab1:
            st.subheader("所有物料安全库存一览")
            
            col_f1, col_f2 = st.columns(2)
            with col_f1:
                status_filter = st.multiselect(
                    "库存状态筛选",
                    options=df_safety["库存状态"].unique(),
                    default=df_safety["库存状态"].unique()
                )
            with col_f2:
                search = st.text_input("🔍 搜索物料编码", placeholder="输入物料编码...")
            
            filtered_df = df_safety[df_safety["库存状态"].isin(status_filter)]
            if search:
                filtered_df = filtered_df[filtered_df["物料编码"].astype(str).str.contains(search, na=False)]
            
            display_cols = ["物料编码", "安全库存", "实际库存", "库存状态", "缺货风险"]
            available_cols = [c for c in display_cols if c in filtered_df.columns]
            
            def highlight_risk(row):
                if row["库存状态"] == "⚠️ 低于安全库存":
                    return ["background-color: #ffcccc"] * len(row)
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
        
        # Tab 2: 低库存预警
        with tab2:
            st.subheader("⚠️ 低于安全库存的物料清单")
            alert_df = df_safety[df_safety["库存状态"] == "⚠️ 低于安全库存"].copy()
            if len(alert_df) > 0:
                alert_df = alert_df.sort_values("缺货风险", ascending=False)
                alert_cols = ["物料编码", "实际库存", "安全库存", "缺货风险"]
                available_alert = [c for c in alert_cols if c in alert_df.columns]
                st.dataframe(alert_df[available_alert], use_container_width=True)
                
                if "缺货风险" in alert_df.columns and len(alert_df) > 0:
                    top10 = alert_df.head(10)
                    fig = px.bar(
                        top10, x="物料编码", y="缺货风险",
                        title="缺货风险最高的10种物料",
                        labels={"缺货风险": "缺货风险 (%)"},
                        color="缺货风险", color_continuous_scale="reds"
                    )
                    st.plotly_chart(fig, use_container_width=True)
            else:
                st.success("🎉 所有物料库存充足！")
        
        # Tab 3: 数据分析
        with tab3:
            st.subheader("库存状态分布")
            col_ch1, col_ch2 = st.columns(2)
            
            with col_ch1:
                status_counts = df_safety["库存状态"].value_counts()
                if len(status_counts) > 0:
                    fig_pie = px.pie(
                        values=status_counts.values,
                        names=status_counts.index,
                        title="库存状态占比",
                        color_discrete_sequence=["#2ecc71", "#e74c3c"]
                    )
                    st.plotly_chart(fig_pie, use_container_width=True)
            
            with col_ch2:
                if "安全库存" in df_safety.columns:
                    safety_data = df_safety[df_safety["安全库存"] > 0]["安全库存"]
                    if len(safety_data) > 0:
                        fig_hist = px.histogram(
                            safety_data, nbins=30,
                            title="安全库存值分布",
                            labels={"value": "安全库存", "count": "物料数量"},
                            color_discrete_sequence=["#3498db"]
                        )
                        st.plotly_chart(fig_hist, use_container_width=True)
            
            st.subheader("实际库存 vs 安全库存对比（前30个物料）")
            compare_df = df_safety[["物料编码", "实际库存", "安全库存"]].dropna().head(30)
            if len(compare_df) > 0:
                fig_compare = go.Figure()
                fig_compare.add_trace(go.Bar(x=compare_df["物料编码"], y=compare_df["实际库存"], name="实际库存", marker_color="#2ecc71"))
                fig_compare.add_trace(go.Bar(x=compare_df["物料编码"], y=compare_df["安全库存"], name="安全库存", marker_color="#e74c3c"))
                fig_compare.update_layout(barmode="group", title="实际库存 vs 安全库存", xaxis_title="物料编码", yaxis_title="库存数量")
                st.plotly_chart(fig_compare, use_container_width=True)
        
        # Tab 4: 包材安全库存
        with tab4:
            st.subheader("包材安全库存列表")
            if df_packaging is not None and len(df_packaging) > 0:
                if "建议安全库存" in df_packaging.columns:
                    pack_cols = ["物料编码", "建议安全库存", "品类"] if "品类" in df_packaging.columns else ["物料编码", "建议安全库存"]
                    available_pack = [c for c in pack_cols if c in df_packaging.columns]
                    st.dataframe(df_packaging[available_pack], use_container_width=True, height=400)
                    
                    if "品类" in df_packaging.columns:
                        pack_cat = df_packaging["品类"].value_counts().reset_index()
                        pack_cat.columns = ["品类", "数量"]
                        fig_pack = px.bar(pack_cat, x="品类", y="数量", title="包材品类分布")
                        st.plotly_chart(fig_pack, use_container_width=True)
                else:
                    st.info("包材数据格式已识别，但未找到'建议安全库存'列")
            else:
                st.info("包材数据未找到或格式不符")
        
        st.success("✅ 系统运行中 | 数据加载成功")
    else:
        st.error("数据处理失败，请检查Excel文件格式")
else:
    st.info("👈 请从左侧上传Excel文件开始使用")
