
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

# ==================== 主内容区 ====================

# 上传区域
if uploaded_file is not None:
st.session_state.uploaded_file_name = uploaded_file.name

# 显示已上传文件
st.info(f"✅ 已上传: {uploaded_file.name}")

# 提取月份信息
if '202509' in uploaded_file.name or '2025年09月' in str(uploaded_file.name):
st.session_state.current_month = "2025年09月"

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
    st.error("❌ 未找到必需工作表「安全库存（202509月）」")
    st.stop()

calculator = SafetyStockCalculator()

# 开始计算
if calc_btn:
    with st.spinner("正在计算安全库存..."):
        df_result = calculator.process(df_mat, df_qual, df_cat)
        st.session_state.df_result = df_result
    
    st.success("✅ 安全库存计算完成！")
    st.balloons()
    
    # ==================== KPI 指标卡片 ====================
    total_materials = len(df_result)
    low_stock = len(df_result[df_result['库存覆盖倍数'] < 1.5])
    low_pct = round(low_stock / total_materials * 100, 1)
    avg_safety = df_result['安全库存'].mean()
    total_actual = df_result['实际库存'].sum()
    avg_coverage = df_result['库存覆盖倍数'].mean()
    
    # 第一行指标
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.markdown(f"""
        <div class="stat-card">
            <div class="stat-value">{total_materials}</div>
            <div class="stat-label">📦 物料总览</div>
            <div class="stat-unit">个物料</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        st.markdown(f"""
        <div class="stat-card">
            <div class="stat-value" style="color: #F59E0B;">{low_stock}</div>
            <div class="stat-label">⚠️ 低库存预警物料</div>
            <div class="stat-unit">占总数 {low_pct}%</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col3:
        st.markdown(f"""
        <div class="stat-card">
            <div class="stat-value" style="color: #10B981;">{avg_safety:,.0f}</div>
            <div class="stat-label">📊 平均计算安全库存</div>
            <div class="stat-unit">件/物料</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col4:
        status_color = "#10B981" if avg_coverage >= 1.5 else "#F59E0B"
        status_text = "库存充足" if avg_coverage >= 1.5 else "库存偏低"
        st.markdown(f"""
        <div class="stat-card">
            <div class="stat-value" style="color: {status_color}">{avg_coverage:.1f}倍</div>
            <div class="stat-label">📈 平均库存覆盖倍数</div>
            <div class="stat-unit">{status_text}</div>
        </div>
        """, unsafe_allow_html=True)
    
    # 第二行指标
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.markdown(f"""
        <div class="stat-card">
            <div class="stat-value" style="color: #8B5CF6;">{total_actual:,.0f}</div>
            <div class="stat-label">💰 总实际库存</div>
            <div class="stat-unit">所有物料合计</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        total_safety = df_result['安全库存'].sum()
        st.markdown(f"""
        <div class="stat-card">
            <div class="stat-value" style="color: #3B82F6;">{total_safety:,.0f}</div>
            <div class="stat-label">📋 建议库存总量</div>
            <div class="stat-unit">基于公式计算</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col3:
        # 包材库存（简化计算）
        pkg_status = "库存充足" if avg_coverage >= 1.5 else "库存偏低"
        st.markdown(f"""
        <div class="stat-card">
            <div class="stat-value" style="color: #06B6D4;">{avg_coverage:.1f}倍</div>
            <div class="stat-label">📦 包材库存</div>
            <div class="stat-unit">{pkg_status}</div>
        </div>
        """, unsafe_allow_html=True)
    
    st.markdown("---")
    
    # ==================== 高风险预警 ====================
    high_risk = df_result[df_result['库存覆盖倍数'] < 1].nlargest(10, '安全库存')
    if not high_risk.empty:
        st.markdown("""
        <div class="alert-card">
            <div class="alert-title">⚠️ 高风险物料预警</div>
            以下物料库存严重不足，建议立即采购
        </div>
        """, unsafe_allow_html=True)
        
        display_cols = [df_result.columns[0], '实际库存', '安全库存', '库存覆盖倍数', '预警等级']
        display_cols = [c for c in display_cols if c in high_risk.columns]
        st.dataframe(
            high_risk[display_cols].style.format({
                '实际库存': '{:,.0f}',
                '安全库存': '{:,.0f}',
                '库存覆盖倍数': '{:.1f}'
            }),
            use_container_width=True,
            height=250
        )
        st.markdown("---")
    
    # ==================== Tab 页面 ====================
    tab1, tab2, tab3 = st.tabs(["📋 数据详情", "📊 可视化分析", "💾 数据导出"])
    
    with tab1:
        st.caption(f"共 {len(df_result)} 条记录")
        display_cols = [df_result.columns[0], '未来3个月月均用量', '过去6个月月均用量',
                       '质量风险系数', '品类策略系数', '采购周期系数', '安全库存', 
                       '实际库存', '库存覆盖倍数', '预警等级']
        display_cols = [c for c in display_cols if c in df_result.columns]
        
        st.dataframe(
            df_result[display_cols].style.format({
                '未来3个月月均用量': '{:,.0f}',
                '过去6个月月均用量': '{:,.0f}',
                '安全库存': '{:,.0f}',
                '实际库存': '{:,.0f}',
                '库存覆盖倍数': '{:.1f}'
            }),
            use_container_width=True,
            height=500
        )
    
    with tab2:
        col1, col2 = st.columns(2)
        
        with col1:
            warning_counts = df_result['预警等级'].value_counts().reset_index()
            warning_counts.columns = ['预警等级', '数量']
            fig_pie = px.pie(
                warning_counts, values='数量', names='预警等级',
                title='库存状态分布',
                color='预警等级',
                color_discrete_map={
                    '✅ 充足': '#10B981',
                    '🟢 正常': '#3B82F6',
                    '🟡 偏低': '#F59E0B',
                    '🔴 严重不足': '#EF4444'
                }
            )
            fig_pie.update_layout(height=400, paper_bgcolor='rgba(0,0,0,0)')
            st.plotly_chart(fig_pie, use_container_width=True)
        
        with col2:
            top_ss = df_result.nlargest(15, '安全库存')[[df_result.columns[0], '安全库存']]
            fig_bar = px.bar(
                top_ss, x=df_result.columns[0], y='安全库存',
                title='安全库存 TOP 15',
                color='安全库存',
                color_continuous_scale='Blues'
            )
            fig_bar.update_layout(height=400, xaxis_tickangle=-45, paper_bgcolor='rgba(0,0,0,0)')
            st.plotly_chart(fig_bar, use_container_width=True)
        
        fig_hist = px.histogram(
            df_result[df_result['库存覆盖倍数'] < 10],
            x='库存覆盖倍数', nbins=30,
            title='库存覆盖倍数分布',
            color_discrete_sequence=['#3B82F6']
        )
        fig_hist.update_layout(height=400, paper_bgcolor='rgba(0,0,0,0)')
        st.plotly_chart(fig_hist, use_container_width=True)
    
    with tab3:
        st.subheader("📥 导出数据")
        
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

elif 'df_result' in st.session_state and st.session_state.df_result is not None:
    # 显示已缓存的结果
    df_result = st.session_state.df_result
    
    total_materials = len(df_result)
    low_stock = len(df_result[df_result['库存覆盖倍数'] < 1.5])
    low_pct = round(low_stock / total_materials * 100, 1)
    avg_safety = df_result['安全库存'].mean()
    total_actual = df_result['实际库存'].sum()
    avg_coverage = df_result['库存覆盖倍数'].mean()
    
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("📦 物料总览", f"{total_materials:,}")
    with col2:
        st.metric("⚠️ 低库存预警", f"{low_stock:,}", f"占比 {low_pct}%")
    with col3:
        st.metric("📊 平均安全库存", f"{avg_safety:,.0f}")
    with col4:
        st.metric("📈 平均库存覆盖", f"{avg_coverage:.1f}倍")
    
    st.info("💡 如需重新计算，请点击左侧「开始计算」按钮")
else:
    st.info("👈 请点击左侧「开始计算」按钮")
    
except Exception as e:
st.error(f"❌ 处理文件时出错: {str(e)}")

else:
# 未上传文件时的占位
st.markdown("""
<div class="upload-area">
<h3>📂 拖拽Excel文件到此处，或点击上传</h3>
<p>支持 .xlsx, .xls 格式 | 必需列：物料编码、未来3个月月均用量、过去6个月月均用量、质量风险系数、品类策略系数、采购周期系数、采购计划执行系数、6月末实际库存</p>
</div>
""", unsafe_allow_html=True)

# 页脚
st.markdown("---")
st.markdown(
"<div style='text-align: center; color: #94A3B8; padding: 20px;'>安全库存管理系统 | 支持月度数据自动更新 | 严格按照Excel公式计算</div>",
unsafe_allow_html=True
)


if __name__ == "__main__":
main()
