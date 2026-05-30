
**采购周期系数：**
| 天数 | 系数 |
小时|------|------|
| 1-7天 | 0.2 |
| 8-15天 | 0.3 |
| 16-21天 | 0.6 |
| 22-30天 | 1.0 |
| 31-40天 | 1.2 |
| 41-45天 | 1.5 |
| 45-60天 | 2.0 |
| 60天以上 | 3.0 |
""")

# 按钮区域
col1, col2 = st.columns([1, 1])
with col1:
calc_btn = st.button("🚀 开始计算安全库存", type="primary", use_container_width=True)
with col2:
update_btn = st.button("🔄 执行月度更新", use_container_width=True)

st.divider()

# 处理上传文件
if uploaded_file is not None:
try:
with st.spinner("正在加载数据..."):
    excel_file = pd.ExcelFile(uploaded_file)
    sheets = {}
    for sheet_name in excel_file.sheet_names:
        sheets[sheet_name] = pd.read_excel(uploaded_file, sheet_name=sheet_name)

df_mat = sheets.get('安全库存（202509月）')
df_qual = sheets.get('原辅料质量等级风险')
df_cat = sheets.get('品类策略系数')

if df_mat is None:
    st.error("❌ 未找到工作表：安全库存（202509月）")
    st.stop()

calculator = SafetyStockCalculator()

if calc_btn:
    with st.spinner("正在计算安全库存..."):
        df_result = calculator.process(df_mat, df_qual, df_cat)
        st.session_state.result = df_result
    
    st.success("✅ 安全库存计算完成！")
    st.balloons()
    
    # ==================== KPI指标卡片 ====================
    total = len(df_result)
    low = len(df_result[df_result['库存覆盖倍数'] < 1.5])
    low_pct = round(low / total * 100, 1)
    avg_ss = df_result['安全库存'].mean()
    total_actual = df_result['实际库存'].sum()
    avg_cov = df_result['库存覆盖倍数'].mean()
    
    kpi_cols = st.columns(5)
    kpi_data = [
        ("📦 物料总览", f"{total:,}", "个物料"),
        ("⚠️ 低库存预警", f"{low:,}", f"占比 {low_pct}%"),
        ("📊 平均安全库存", f"{avg_ss:,.0f}", "件/物料"),
        ("💰 总实际库存", f"{total_actual:,.0f}", "件"),
        ("📈 平均库存覆盖", f"{avg_cov:.1f}倍", "充足" if avg_cov >= 1.5 else "偏低")
    ]
    
    for col, (label, value, unit) in zip(kpi_cols, kpi_data):
        with col:
            st.markdown(f"""
            <div class="metric-card">
                <div class="metric-value">{value}</div>
                <div class="metric-label">{label}</div>
                <div style="font-size: 11px; color: #94a3b8;">{unit}</div>
            </div>
            """, unsafe_allow_html=True)
    
    st.divider()
    
    # ==================== 四个并列卡片 ====================
    st.markdown("### 📋 功能模块")
    
    # 创建4个并列按钮
    col_alert, col_data, col_chart, col_export = st.columns(4)
    
    with col_alert:
        show_alert = st.button("⚠️ 物料预警", use_container_width=True, key="btn_alert")
    with col_data:
        show_data = st.button("📋 数据详情", use_container_width=True, key="btn_data")
    with col_chart:
        show_chart = st.button("📊 可视化分析", use_container_width=True, key="btn_chart")
    with col_export:
        show_export = st.button("💾 数据导出", use_container_width=True, key="btn_export")
    
    st.divider()
    
    # ==================== 高风险物料预警区域 ====================
    high_risk = df_result[df_result['库存覆盖倍数'] < 1].nlargest(10, '安全库存')
    
    if show_alert or ('active_tab' in st.session_state and st.session_state.active_tab == 'alert'):
        st.session_state.active_tab = 'alert'
        st.markdown("### ⚠️ 高风险物料预警")
        st.caption("以下物料库存严重不足，建议立即采购")
        
        if not high_risk.empty:
            alert_df = high_risk[['物料编码', '实际库存', '安全库存', '库存覆盖倍数', '预警等级']].copy()
            alert_df.columns = ['物料编码', '实际库存', '安全库存', '库存覆盖倍数', '预警等级']
            st.dataframe(
                alert_df.style.format({
                    '实际库存': '{:,.0f}',
                    '安全库存': '{:,.0f}',
                    '库存覆盖倍数': '{:.1f}'
                }),
                use_container_width=True,
                height=300
            )
        else:
            st.success("🎉 暂无高风险物料，所有物料库存充足！")
    
    # ==================== 数据详情区域 ====================
    if show_data or ('active_tab' in st.session_state and st.session_state.active_tab == 'data'):
        st.session_state.active_tab = 'data'
        st.markdown("### 📋 数据详情")
        st.caption(f"共 {len(df_result)} 条记录")
        
        detail_df = df_result[[
            '物料编码', '未来3个月月均用量', '过去6个月月均用量',
            '质量风险系数', '品类策略系数', '采购周期系数', 
            '安全库存', '实际库存', '库存覆盖倍数', '预警等级'
        ]].copy()
        
        detail_df.columns = [
            '物料编码', '未来3个月月均用量', '过去6个月月均用量',
            '质量风险系数', '品类策略系数', '采购周期系数',
            '安全库存', '实际库存', '库存覆盖倍数', '预警等级'
        ]
        
        st.dataframe(
            detail_df.head(100).style.format({
                '未来3个月月均用量': '{:,.0f}',
                '过去6个月月均用量': '{:,.0f}',
                '安全库存': '{:,.0f}',
                '实际库存': '{:,.0f}',
                '库存覆盖倍数': '{:.1f}'
            }),
            use_container_width=True,
            height=500
        )
    
    # ==================== 可视化分析区域 ====================
    if show_chart or ('active_tab' in st.session_state and st.session_state.active_tab == 'chart'):
        st.session_state.active_tab = 'chart'
        st.markdown("### 📊 可视化分析")
        
        chart_col1, chart_col2 = st.columns(2)
        
        with chart_col1:
            warning_counts = df_result['预警等级'].value_counts().reset_index()
            warning_counts.columns = ['预警等级', '数量']
            fig_pie = px.pie(
                warning_counts, values='数量', names='预警等级',
                title='📊 库存状态分布',
                color='预警等级',
                color_discrete_map={
                    '✅ 充足': '#10B981',
                    '🟢 正常': '#3B82F6',
                    '🟡 偏低': '#F59E0B',
                    '🔴 严重不足': '#EF4444'
                }
            )
            fig_pie.update_layout(height=400)
            st.plotly_chart(fig_pie, use_container_width=True)
        
        with chart_col2:
            top_ss = df_result.nlargest(15, '安全库存')[['物料编码', '安全库存']]
            fig_bar = px.bar(
                top_ss, x='物料编码', y='安全库存',
                title='📊 安全库存 TOP 15',
                color='安全库存',
                color_continuous_scale='Blues'
            )
            fig_bar.update_layout(height=400, xaxis_tickangle=-45)
            st.plotly_chart(fig_bar, use_container_width=True)
        
        fig_hist = px.histogram(
            df_result[df_result['库存覆盖倍数'] < 10],
            x='库存覆盖倍数', nbins=30,
            title='📊 库存覆盖倍数分布',
            color_discrete_sequence=['#3B82F6']
        )
        fig_hist.update_layout(height=400)
        st.plotly_chart(fig_hist, use_container_width=True)
    
    # ==================== 数据导出区域 ====================
    if show_export or ('active_tab' in st.session_state and st.session_state.active_tab == 'export'):
        st.session_state.active_tab = 'export'
        st.markdown("### 💾 数据导出")
        
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_result.to_excel(writer, sheet_name='安全库存计算结果', index=False)
            
            summary = pd.DataFrame({
                '指标': ['总物料数', '总安全库存', '平均安全库存', '总实际库存',
                       '低库存物料数', '低库存占比', '平均库存覆盖', '计算时间'],
                '数值': [
                    len(df_result),
                    f"{df_result['安全库存'].sum():,.0f}",
                    f"{df_result['安全库存'].mean():,.0f}",
                    f"{df_result['实际库存'].sum():,.0f}",
                    len(df_result[df_result['库存覆盖倍数'] < 1.5]),
                    f"{len(df_result[df_result['库存覆盖倍数'] < 1.5]) / len(df_result) * 100:.1f}%",
                    f"{df_result['库存覆盖倍数'].mean():.1f}倍",
                    datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                ]
            })
            summary.to_excel(writer, sheet_name='汇总统计', index=False)
            
            if not high_risk.empty:
                high_risk.to_excel(writer, sheet_name='高风险物料', index=False)
        
        output.seek(0)
        st.download_button(
            label="📎 下载Excel报告",
            data=output,
            file_name=f"安全库存报告_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
        st.info("💡 报告包含：安全库存计算结果、汇总统计、高风险物料")

elif st.session_state.result is not None:
    df_result = st.session_state.result
    total = len(df_result)
    low = len(df_result[df_result['库存覆盖倍数'] < 1.5])
    avg_ss = df_result['安全库存'].mean()
    avg_cov = df_result['库存覆盖倍数'].mean()
    
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("物料总览", f"{total:,}")
    with col2:
        st.metric("低库存预警", f"{low:,}")
    with col3:
        st.metric("平均安全库存", f"{avg_ss:,.0f}")
    with col4:
        st.metric("平均库存覆盖", f"{avg_cov:.1f}倍")
    
    st.info("💡 如需重新计算，请点击上方「开始计算安全库存」按钮")
else:
    st.info("👆 请点击上方「开始计算安全库存」按钮开始计算")
    
except Exception as e:
st.error(f"❌ 处理文件时出错: {str(e)}")
else:
st.info("👈 请在左侧上传Excel文件")

with st.expander("📖 使用说明"):
st.markdown("""
### 使用方法
1. 在左侧上传Excel文件
2. 点击「开始计算安全库存」按钮
3. 点击下方功能按钮查看详情

### 必需工作表
- `安全库存（202509月）`: 物料用量数据
""")

st.divider()
st.caption("安全库存管理系统 | 版本 3.0 | 支持月度数据更新")


if __name__ == "__main__":
main()
