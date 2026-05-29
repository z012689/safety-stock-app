""")
with col2:
st.markdown("""
**采购周期系数对照表:**
| 交货周期 | 系数 |
|---------|------|
| 1-7天 | 0.2 |
| 8-15天 | 0.3 |
| 16-21天 | 0.6 |
| 22-30天 | 1.0 |
| 31-40天 | 1.2 |
| 41-45天 | 1.5 |
| 45-60天 | 2.0 |
| 60天以上 | 3.0 |
""")

# ==================== Tab 页面 ====================
tab1, tab2, tab3 = st.tabs(["📋 数据详情", "📊 可视化分析", "💾 数据导出"])

with tab1:
st.caption(f"共 {len(df_result)} 条记录，显示前100条")
display_cols = [df_result.columns[0], '未来3个月月均用量', '过去6个月月均用量',
           '质量风险系数', '品类策略系数', '采购周期系数', '安全库存', 
           '实际库存', '库存覆盖倍数', '预警等级']
display_cols = [c for c in display_cols if c in df_result.columns]

st.dataframe(
df_result[display_cols].head(100).style.format({
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
    title='📊 库存状态分布',
    color='预警等级',
    color_discrete_map={
        '✅ 充足': COLORS['success'],
        '🟢 正常': COLORS['info'],
        '🟡 偏低': COLORS['warning'],
        '🔴 严重不足': COLORS['danger']
    }
)
fig_pie.update_layout(height=400, paper_bgcolor='rgba(0,0,0,0)')
st.plotly_chart(fig_pie, use_container_width=True)

with col2:
top_ss = df_result.nlargest(15, '安全库存')[[df_result.columns[0], '安全库存']]
fig_bar = px.bar(
    top_ss, x=df_result.columns[0], y='安全库存',
    title='📊 安全库存 TOP 15',
    color='安全库存',
    color_continuous_scale='Blues'
)
fig_bar.update_layout(height=400, xaxis_tickangle=-45, paper_bgcolor='rgba(0,0,0,0)')
st.plotly_chart(fig_bar, use_container_width=True)

fig_hist = px.histogram(
df_result[df_result['库存覆盖倍数'] < 10],
x='库存覆盖倍数', nbins=30,
title='📊 库存覆盖倍数分布',
color_discrete_sequence=[COLORS['primary']]
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

high_risk.to_excel(writer, sheet_name='高风险物料', index=False)

output.seek(0)
st.download_button(
label="📎 下载Excel报告",
data=output,
file_name=f"安全库存报告_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

st.info("💡 报告包含3个工作表：安全库存计算结果、汇总统计、高风险物料")

except Exception as e:
st.error(f"❌ 处理文件时出错: {str(e)}")
st.exception(e)

else:
# 未上传文件时的占位
st.markdown("""
<div style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); 
border-radius: 20px; padding: 60px; text-align: center; color: white; margin: 20px 0;">
<h3>📂 拖拽Excel文件到此处，或点击上传</h3>
<p style="margin-top: 10px; opacity: 0.8;">支持 .xlsx, .xls 格式 | 必需列：物料编码、未来3个月月均用量、过去6个月月均用量、质量风险系数、品类策略系数、采购周期系数、采购计划执行系数、6月末实际库存</p>
</div>
""", unsafe_allow_html=True)

# ==================== 侧边栏（月度更新） ====================
with st.sidebar:
st.markdown(f"### 📅 月度数据更新")
st.markdown(f"<p style='color: #6C757D; font-size: 14px;'>更新新月份数据，系统自动滚动计算</p>", unsafe_allow_html=True)

new_month_name = st.text_input("新月份名称", value=f"{datetime.now().year}年{datetime.now().month}月")

update_file = st.file_uploader(
"上传新月份数据文件",
type=['xlsx', 'xls'],
key="monthly_update",
help="文件需包含「物料编码」和「用量」两列"
)

st.divider()

if st.session_state.df_result is not None:
st.markdown(f"### 📊 当前数据状态")
st.markdown(f"- 物料数量: {len(st.session_state.df_result)}")
st.markdown(f"- 低库存物料: {len(st.session_state.df_result[st.session_state.df_result['库存覆盖倍数'] < 1.5])}")
st.markdown(f"- 平均覆盖: {st.session_state.df_result['库存覆盖倍数'].mean():.1f}倍")

st.divider()

st.markdown("### 📞 技术支持")
st.caption("如有问题，请联系管理员")

# 页脚
st.divider()
st.markdown(
"<div style='text-align: center; color: #ADB5BD; padding: 20px;'>安全库存管理系统 | 支持月度数据自动更新 | 版本 3.0</div>",
unsafe_allow_html=True
)


if __name__ == "__main__":
main()
