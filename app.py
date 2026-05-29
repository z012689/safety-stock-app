
#### 3. 计算公式
安全库存 = (未来3个月月均 + 过去6个月月均) / 2 × 综合系数 × 采购周期系数

### 文件要求
Excel文件应包含以下sheet:
- `安全库存（202509月）`: 物料用量数据（必需）
- `原辅料质量等级风险`: 质量风险数据（可选）
- `品类策略系数`: 品类风险评估（可选）
""")

st.markdown("---")
st.markdown(
"<div style='text-align: center; color: gray;'>安全库存管理系统 | 版本 2.0（含月度更新）</div>",
unsafe_allow_html=True
)


def self_show_results(df_result):
"""显示计算结果"""
# 关键指标
col1, col2, col3, col4 = st.columns(4)
with col1:
st.metric("总物料数", len(df_result))
with col2:
total_ss = df_result['安全库存'].sum() if '安全库存' in df_result.columns else 0
st.metric("总安全库存", f"{total_ss:,.0f}")
with col3:
avg_ss = df_result['安全库存'].mean() if '安全库存' in df_result.columns else 0
st.metric("平均安全库存", f"{avg_ss:,.0f}")
with col4:
zero_ss = (df_result['安全库存'] == 0).sum() if '安全库存' in df_result.columns else 0
st.metric("零库存物料", zero_ss)

st.markdown("---")

# 标签页
tab1, tab2, tab3 = st.tabs(["📋 数据表格", "📊 统计分析", "💾 数据导出"])

with tab1:
display_cols = ['物料编码', 'SAP编码', '未来3个月月均用量', '过去6个月月均用量',
           '质量风险系数', '品类策略系数', '采购周期系数', '安全库存', '是否寄售', '备注']
display_cols = [c for c in display_cols if c in df_result.columns]

st.dataframe(
df_result[display_cols].head(100).style.format({
    '未来3个月月均用量': '{:,.0f}',
    '过去6个月月均用量': '{:,.0f}',
    '安全库存': '{:,.0f}'
}),
use_container_width=True,
height=400
)

with tab2:
col1, col2 = st.columns(2)

with col1:
if '安全库存' in df_result.columns:
    df_plot = df_result[df_result['安全库存'] > 0].copy()
    if len(df_plot) > 0:
        fig_hist = px.histogram(
            df_plot, x='安全库存', nbins=30,
            title='安全库存分布',
            labels={'安全库存': '安全库存数量', 'count': '物料数量'}
        )
        st.plotly_chart(fig_hist, use_container_width=True)

with col2:
if '是否寄售' in df_result.columns:
    consignment_data = df_result.groupby('是否寄售')['安全库存'].agg(['sum', 'count']).reset_index()
    consignment_data.columns = ['是否寄售', '总安全库存', '物料数量']
    fig_pie = px.pie(
        consignment_data, values='物料数量', names='是否寄售',
        title='寄售物料占比'
    )
    st.plotly_chart(fig_pie, use_container_width=True)

with tab3:
st.subheader("📥 数据导出")

output = BytesIO()
with pd.ExcelWriter(output, engine='openpyxl') as writer:
df_result.to_excel(writer, sheet_name='安全库存计算结果', index=False)

summary_data = {
    '指标': ['总物料数', '总安全库存', '平均安全库存', '最大安全库存', '最小安全库存', '零库存物料数'],
    '数值': [
        len(df_result),
        df_result['安全库存'].sum(),
        df_result['安全库存'].mean(),
        df_result['安全库存'].max(),
        df_result['安全库存'].min(),
        (df_result['安全库存'] == 0).sum()
    ]
}
df_summary = pd.DataFrame(summary_data)
df_summary.to_excel(writer, sheet_name='汇总统计', index=False)

output.seek(0)
st.download_button(
label="📎 下载Excel文件",
data=output,
file_name=f"安全库存计算结果_{datetime.now().strftime('%Y%m%d')}.xlsx",
mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)


if __name__ == "__main__":
main()
