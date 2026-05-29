import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO

# 页面配置
st.set_page_config(
    page_title="安全库存管理系统",
    page_icon="📊",
    layout="wide"
)

# 自定义CSS
st.markdown("""
<style>
    .stDataFrame { width: 100% }
    .metric-card {
        background-color: #f0f2f6;
        padding: 15px;
        border-radius: 10px;
        margin: 10px 0;
    }
</style>
""", unsafe_allow_html=True)


class SafetyStockCalculator:
    """安全库存计算器"""
    
    @staticmethod
    def calculate_lead_time_coefficient(days):
        """计算采购周期系数"""
        if pd.isna(days):
            return 1
        days = float(days)
        if days <= 7:
            return 0.2
        elif days <= 15:
            return 0.3
        elif days <= 21:
            return 0.6
        elif days <= 30:
            return 1
        elif days <= 40:
            return 1.2
        elif days <= 45:
            return 1.5
        elif days <= 60:
            return 2
        else:
            return 3
    
    @staticmethod
    def calculate_combined_coefficient(quality_risk, category_strategy):
        """计算综合系数"""
        if pd.isna(quality_risk):
            quality_risk = 1
        if pd.isna(category_strategy):
            category_strategy = 1
        return quality_risk * 0.4 + category_strategy * 0.6
    
    @staticmethod
    def calculate_safety_stock(avg_future, avg_past, combined_coef, lead_time_coef):
        """计算安全库存"""
        if pd.isna(avg_future) or pd.isna(avg_past):
            return 0
        if avg_future <= 0 and avg_past <= 0:
            return 0
        base_usage = (avg_future + avg_past) / 2
        return base_usage * combined_coef * lead_time_coef
    
    def process_raw_materials(self, df_materials, df_quality, df_category):
        """处理原料安全库存计算"""
        df = df_materials.copy()
        
        # 查找用量列
        future_cols = []
        past_cols = []
        
        for col in df.columns:
            col_str = str(col)
            if 'M07' in col_str or 'M08' in col_str or 'M09' in col_str:
                future_cols.append(col)
            elif 'M01' in col_str or 'M02' in col_str or 'M03' in col_str or 'M04' in col_str or 'M05' in col_str or 'M06' in col_str:
                past_cols.append(col)
        
        # 计算月均用量
        if future_cols:
            df['未来3个月月均用量'] = pd.to_numeric(df[future_cols].stack(), errors='coerce').unstack().mean(axis=1)
        elif '未来3个月月均用量' in df.columns:
            df['未来3个月月均用量'] = pd.to_numeric(df['未来3个月月均用量'], errors='coerce')
        
        if past_cols:
            df['过去6个月月均用量'] = pd.to_numeric(df[past_cols].stack(), errors='coerce').unstack().mean(axis=1)
        elif '月均量(半年)' in df.columns:
            df['过去6个月月均用量'] = pd.to_numeric(df['月均量(半年)'], errors='coerce')
        
        # 计算采购周期系数
        if '平均交货周期(天)' in df.columns:
            df['采购周期系数'] = df['平均交货周期(天)'].apply(self.calculate_lead_time_coefficient)
        
        # 获取质量风险系数
        quality_dict = {}
        if df_quality is not None and '物料编码' in df_quality.columns:
            for _, row in df_quality.iterrows():
                code = row.get('物料编码')
                score = row.get('质量风险等级得分')
                if pd.notna(code) and pd.notna(score):
                    quality_dict[str(code)] = float(score)
        
        # 获取品类策略系数
        category_dict = {}
        if df_category is not None:
            for col in df_category.columns:
                if '物料' in col:
                    code_col = col
                    break
            risk_col = None
            for col in df_category.columns:
                if '风险' in col:
                    risk_col = col
                    break
            
            if code_col and risk_col:
                for _, row in df_category.iterrows():
                    code = row.get(code_col)
                    risk = row.get(risk_col)
                    if pd.notna(code) and pd.notna(risk):
                        category_dict[str(code)] = float(risk)
        
        # 应用系数
        if '物料编码' in df.columns:
            df['质量风险系数'] = df['物料编码'].astype(str).map(quality_dict).fillna(1)
            df['品类策略系数'] = df['物料编码'].astype(str).map(category_dict).fillna(1)
        
        if '原辅料质量等级风险' in df.columns:
            df['质量风险系数'] = pd.to_numeric(df['原辅料质量等级风险'], errors='coerce').fillna(1)
        if '品类策略风险系数' in df.columns:
            df['品类策略系数'] = pd.to_numeric(df['品类策略风险系数'], errors='coerce').fillna(1)
        
        # 计算综合系数
        df['综合系数'] = df.apply(
            lambda x: self.calculate_combined_coefficient(
                x.get('质量风险系数', 1),
                x.get('品类策略系数', 1)
            ), axis=1
        )
        
        # 计算安全库存
        df['安全库存'] = df.apply(
            lambda x: self.calculate_safety_stock(
                x.get('未来3个月月均用量', 0),
                x.get('过去6个月月均用量', 0),
                x.get('综合系数', 1),
                x.get('采购周期系数', 1)
            ), axis=1
        )
        
        # 寄售物料安全库存为0
        if '是否寄售' in df.columns:
            df.loc[df['是否寄售'] == '寄售', '安全库存'] = 0
        
        return df


def main():
    st.title("📦 安全库存管理系统")
    st.markdown("---")
    
    # 侧边栏
    with st.sidebar:
        st.header("⚙️ 系统配置")
        
        uploaded_file = st.file_uploader(
            "上传安全库存Excel文件",
            type=['xlsx', 'xls'],
            help="请上传包含安全库存规则和数据的Excel文件"
        )
        
        st.markdown("---")
        st.subheader("📅 月度更新")
        st.info(f"当前时间: {datetime.now().strftime('%Y-%m-%d')}")
        
        st.markdown("---")
        calculate_btn = st.button("🚀 开始计算安全库存", type="primary", use_container_width=True)
    
    # 主要内容
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
            
            if df_materials is not None:
                calculator = SafetyStockCalculator()
                
                if calculate_btn:
                    with st.spinner("正在计算安全库存..."):
                        df_result = calculator.process_raw_materials(df_materials, df_quality, df_category)
                    
                    st.success("✅ 安全库存计算完成！")
                    
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
            else:
                st.error("未找到原料安全库存数据表")
        except Exception as e:
            st.error(f"处理文件时出错: {str(e)}")
    else:
        st.info("👈 请先上传安全库存Excel文件")
        
        with st.expander("📖 使用说明"):
            st.markdown("""
            ### 系统功能说明
            
            1. **上传文件**: 上传包含安全库存规则的Excel文件
            2. **自动计算**: 根据规则自动计算安全库存
            3. **数据导出**: 支持Excel格式导出
            
            ### 文件要求
            
            Excel文件应包含以下sheet:
            - `安全库存（202509月）`: 物料用量数据
            - `原辅料质量等级风险`: 质量风险批次数据（可选）
            - `品类策略系数`: 品类风险评估（可选）
            """)
    
    st.markdown("---")
    st.markdown(
        "<div style='text-align: center; color: gray;'>安全库存管理系统 | 版本 1.0</div>",
        unsafe_allow_html=True
    )


if __name__ == "__main__":
    main()
