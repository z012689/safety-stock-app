
**采购周期系数：**
- 1-7天: 0.2
- 8-15天: 0.3
- 16-21天: 0.6
- 22-30天: 1.0
- 31-40天: 1.2
- 41-45天: 1.5
- 45-60天: 2.0
- 60天以上: 3.0
""")

# 主内容区
if uploaded_file is not None:
try:
with st.spinner("正在加载数据..."):
    excel_file = pd.ExcelFile(uploaded_file)
    sheets = {}
    for sheet_name in excel_file.sheet_names:
        sheets[sheet_name] = pd.read_excel(uploaded_file, sheet_name=sheet_name)

df_materials = sheets.get('安全库存（202509月）')
df_quality = sheets.get('原辅料质量等级风险')
df_category = sheets.get('品类策略系数')

if df_materials is None:
    st.error("❌ 未找到「安全库存（202509月）」工作表")
    st.stop()

calculator = SafetyStockCalculator()

if calc_btn:
    with st.spinner("正在计算安全库存..."):
        df_result = calculator.process_data(df_materials, df_quality, df_category)
        st.session_state['df_result'] = df_result
    
    st.success("✅ 安全库存计算完成！")
    st.balloons()
    
    # ========== KPI 指标卡片 ==========
    total_materials = len(df_result)
    low_stock = len(df_result[df_result['库存覆盖倍数'] < 1.5])
    low_pct = round(low_stock / total_materials * 100, 1)
    avg_safety = df_result['安全库存'].mean()
    total_actual = df_result['实际库存'].sum()
    avg_coverage = df_result['库存覆盖倍数'].mean()
    
    col1, col2, col3, col4, col5 = st.columns(5)
    
    with col1:
        st.metric("📦 物料总览", f"{total_materials:,}", "个物料")
    with col2:
        st.metric("⚠️ 低库存预警", f"{low_stock:,}", f"占比 {low_pct}%")
    with col3:
        st.metric("📊 平均安全库存", f"{avg_safety:,.0f}", "件/物料")
    with col4:
        st.metric("💰 总实际库存", f"{total_actual:,.0f}", "件")
    with col5:
        color = "normal" if avg_coverage >= 2 else "inverse"
        st.metric("📈 平均库存覆盖", f"{avg_coverage:.1f}倍", "充足" if avg_coverage >= 1.5 else "不足")
    
    st.divider()
    
    # ========== 高风险物料预警 ==========
    high_risk = df_result[df_result['库存覆盖倍数'] < 1].nlargest(10, '安全库存')
    if not high_risk.empty:
        st.warning("⚠️ 以下物料库存严重不足，建议立即采购")
        st.dataframe(
            high_risk[['物料编码', 'SAP编码', '实际库存', '安全库存', '库存覆盖倍数', '预警等级']],
            use_container_width=True,
            height=200
        )
        st.divider()
    
    # ========== Tab 页面 ==========
    tab1, tab2, tab3 = st.tabs(["📋 数据详情", "📊 可视化分析", "💾 数据导出"])
    
    with tab1:
        # 数据显示
        display_cols = ['物料编码', 'SAP编码', '未来3个月月均用量', '过去6个月月均用量',
                       '质量风险系数', '品类策略系数', '采购周期系数', '安全库存', 
                       '实际库存', '库存覆盖倍数', '预警等级', '是否寄售', '备注']
        display_cols = [c for c in display_cols if c in df_result.columns]
        
        # 格式化显示
        df_display = df_result[display_cols].copy()
        for col in ['未来3个月月均用量', '过去6个月月均用量', '安全库存', '实际库存']:
            if col in df_display.columns:
                df_display[col] = df_display[col].apply(lambda x: f"{x:,.0f}" if pd.notna(x) else "")
        
        st.dataframe(df_display, use_container_width=True, height=500)
    
    with tab2:
        col1, col2 = st.columns(2)
        
        with col1:
            # 预警等级分布饼图
            warning_counts = df_result['预警等级'].value_counts().reset_index()
            warning_counts.columns = ['预警等级', '数量']
            fig_pie = px.pie(
                warning_counts, 
                values='数量', 
                names='预警等级',
                title='库存状态分布',
                color='预警等级',
                color_discrete_map={
                    '✅ 过量': '#00C853',
                    '🟢 充足': '#2196F3',
                    '🟡 偏低': '#FFC107',
                    '🔴 严重不足': '#D32F2F'
                }
            )
            fig_pie.update_layout(height=400)
            st.plotly_chart(fig_pie, use_container_width=True)
        
        with col2:
            # 安全库存Top15
            top_safety = df_result.nlargest(15, '安全库存')[['物料编码', '安全库存']]
            fig_bar = px.bar(
                top_safety,
                x='物料编码',
                y='安全库存',
                title='安全库存 TOP 15',
                color='安全库存',
                color_continuous_scale='Reds'
            )
            fig_bar.update_layout(height=400, xaxis_tickangle=-45)
            st.plotly_chart(fig_bar, use_container_width=True)
        
        # 库存覆盖分布直方图
        fig_hist = px.histogram(
            df_result[df_result['库存覆盖倍数'] < 10],
            x='库存覆盖倍数',
            nbins=30,
            title='库存覆盖倍数分布',
            color_discrete_sequence=['#FF6B35']
        )
        fig_hist.update_layout(height=400)
        st.plotly_chart(fig_hist, use_container_width=True)
    
    with tab3:
        st.subheader("📥 导出数据")
        
        # 创建Excel导出
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_result.to_excel(writer, sheet_name='安全库存计算结果', index=False)
            
            # 汇总统计
            summary = pd.DataFrame({
                '指标': ['总物料数', '总安全库存', '平均安全库存', '总实际库存', 
                       '低库存物料数', '低库存占比', '平均库存覆盖'],
                '数值': [
                    len(df_result),
                    f"{df_result['安全库存'].sum():,.0f}",
                    f"{df_result['安全库存'].mean():,.0f}",
                    f"{df_result['实际库存'].sum():,.0f}",
                    len(df_result[df_result['库存覆盖倍数'] < 1.5]),
                    f"{len(df_result[df_result['库存覆盖倍数'] < 1.5]) / len(df_result) * 100:.1f}%",
                    f"{df_result['库存覆盖倍数'].mean():.1f}倍"
                ]
            })
            summary.to_excel(writer, sheet_name='汇总统计', index=False)
            
            # 高风险物料
            high_risk.to_excel(writer, sheet_name='高风险物料', index=False)
        
        output.seek(0)
        st.download_button(
            label="📎 下载Excel报告",
            data=output,
            file_name=f"安全库存报告_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

elif 'df_result' in st.session_state:
    # 显示已缓存的结果
    df_result = st.session_state['df_result']
    
    total_materials = len(df_result)
    low_stock = len(df_result[df_result['库存覆盖倍数'] < 1.5])
    low_pct = round(low_stock / total_materials * 100, 1)
    avg_safety = df_result['安全库存'].mean()
    total_actual = df_result['实际库存'].sum()
    avg_coverage = df_result['库存覆盖倍数'].mean()
    
    col1, col2, col3, col4, col5 = st.columns(5)
    with col1:
        st.metric("📦 物料总览", f"{total_materials:,}")
    with col2:
        st.metric("⚠️ 低库存预警", f"{low_stock:,}", f"占比 {low_pct}%")
    with col3:
        st.metric("📊 平均安全库存", f"{avg_safety:,.0f}")
    with col4:
        st.metric("💰 总实际库存", f"{total_actual:,.0f}")
    with col5:
        st.metric("📈 平均库存覆盖", f"{avg_coverage:.1f}倍")
    
    st.info("💡 如需重新计算，请点击左侧「开始计算」按钮")
else:
    st.info("👈 请点击左侧「开始计算」按钮")
    
except Exception as e:
st.error(f"❌ 处理文件时出错: {str(e)}")
st.exception(e)
else:
# 未上传文件
st.info("👈 请先在左侧上传Excel文件")

# 示例说明
with st.expander("📖 使用说明"):
st.markdown("""
### 快速开始
1. 在左侧点击「上传Excel文件」
2. 选择包含数据的Excel文件
3. 点击「开始计算」

### 文件要求
Excel文件需包含以下工作表：
- `安全库存（202509月）`：物料用量数据（必需）
- `原辅料质量等级风险`：质量风险批次（可选）
- `品类策略系数`：品类风险评估（可选）

### 输出结果
- KPI指标卡片：核心数据一目了然
- 高风险预警：自动识别库存不足物料
- 数据详情：完整的安全库存计算结果
- 可视化图表：库存分布分析
- Excel导出：一键下载报告
""")

# 页脚
st.divider()
st.caption("安全库存管理系统 | 版本 2.0 | 支持月度数据更新")


if __name__ == "__main__":
main()
