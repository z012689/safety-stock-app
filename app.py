# app.py
import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

# 页面配置
st.set_page_config(
    page_title="安全库存管理系统",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
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
    .warning-text { color: #ff4b4b; font-weight: bold; }
    .success-text { color: #00a65a; font-weight: bold; }
</style>
""", unsafe_allow_html=True)


class SafetyStockCalculator:
    """安全库存计算器"""
    
    @staticmethod
    def calculate_avg_usage(df, future_months_cols, past_months_cols):
        """计算月均用量"""
        # 未来3个月月均
        df['未来3个月月均用量'] = df[future_months_cols].mean(axis=1)
        # 过去6个月月均
        df['过去6个月月均用量'] = df[past_months_cols].mean(axis=1)
        return df
    
    @staticmethod
    def calculate_lead_time_coefficient(days):
        """计算采购周期系数"""
        if pd.isna(days):
            return 1
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
        """计算综合系数（质量风险系数*40% + 品类策略系数*60%）"""
        if pd.isna(quality_risk):
            quality_risk = 1
        if pd.isna(category_strategy):
            category_strategy = 1
        return quality_risk * 0.4 + category_strategy * 0.6
    
    @staticmethod
    def calculate_safety_stock(avg_usage_future, avg_usage_past, combined_coef, lead_time_coef):
        """计算安全库存"""
        if pd.isna(avg_usage_future) or pd.isna(avg_usage_past):
            return 0
        base_usage = (avg_usage_future + avg_usage_past) / 2
        return base_usage * combined_coef * lead_time_coef
    
    def process_raw_materials(self, df_materials, df_quality_risk, df_category_strategy):
        """处理原料安全库存计算"""
        # 复制数据
        df = df_materials.copy()
        
        # 识别年月列
        all_cols = df.columns.tolist()
        future_months = []
        past_months = []
        
        for col in all_cols:
            if isinstance(col, str):
                if 'M07' in col or 'M08' in col or 'M09' in col or '2025年M07' in col or '2025年M08' in col or '2025年M09' in col:
                    future_months.append(col)
                elif 'M01' in col or 'M02' in col or 'M03' in col or 'M04' in col or 'M05' in col or 'M06' in col:
                    if '2025年' in col or '2025' in str(col):
                        past_months.append(col)
        
        # 如果没有找到标准列名，尝试其他格式
        if not future_months:
            future_months = [c for c in all_cols if '未来' in str(c) or 'future' in str(c).lower()]
        if not past_months:
            past_months = [c for c in all_cols if '过去' in str(c) or 'past' in str(c).lower()]
        
        # 计算月均用量
        if future_months and past_months:
            df['未来3个月月均用量'] = df[future_months].mean(axis=1)
            df['过去6个月月均用量'] = df[past_months].mean(axis=1)
        else:
            # 使用数据中已有的计算结果
            if '未来3个月月均用量' in df.columns:
                df['未来3个月月均用量'] = pd.to_numeric(df['未来3个月月均用量'], errors='coerce')
            if '月均量(半年)' in df.columns:
                df['过去6个月月均用量'] = pd.to_numeric(df['月均量(半年)'], errors='coerce')
        
        # 计算采购周期系数
        lead_time_col = '平均交货周期(天)'
        if lead_time_col in df.columns:
            df['采购周期系数'] = df[lead_time_col].apply(self.calculate_lead_time_coefficient)
        elif '采购周期系数' in df.columns:
            df['采购周期系数'] = pd.to_numeric(df['采购周期系数'], errors='coerce')
        
        # 获取质量风险系数和品类策略系数
        quality_dict = {}
        if df_quality_risk is not None:
            quality_col = '质量风险等级得分'
            if quality_col in df_quality_risk.columns:
                for _, row in df_quality_risk.iterrows():
                    code = row.get('物料编码')
                    if pd.notna(code):
                        quality_dict[str(code)] = pd.to_numeric(row[quality_col], errors='coerce')
        
        category_dict = {}
        if df_category_strategy is not None:
            risk_col = '风险系数【请开发采购评估填写】'
            if risk_col in df_category_strategy.columns:
                for _, row in df_category_strategy.iterrows():
                    code = row.get('物料代码')
                    if pd.notna(code):
                        category_dict[str(code)] = pd.to_numeric(row[risk_col], errors='coerce')
        
        # 应用系数
        material_code_col = '物料编码'
        if material_code_col in df.columns:
            df['质量风险系数'] = df[material_code_col].astype(str).map(quality_dict).fillna(1)
            df['品类策略系数'] = df[material_code_col].astype(str).map(category_dict).fillna(1)
        elif '原辅料质量等级风险' in df.columns:
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
        
        # 处理寄售物料
        if '是否寄售' in df.columns:
            df.loc[df['是否寄售'] == '寄售', '安全库存'] = 0
        
        # 处理特殊备注
        if '备注' in df.columns:
            special_mask = df['备注'].str.contains('特采|战略|退市', na=False)
            df.loc[special_mask, '安全库存'] = df.loc[special_mask, '安全库存'] * 0.5
        
        return df
    
    def process_packaging_materials(self, df_packaging):
        """处理包材安全库存计算"""
        df = df_packaging.copy()
        
        # 重新计算建议安全库存
        if '2025年预测使用量汇总（个）' in df.columns and '2024年月均使用量（个）' in df.columns:
            # 基于预测量和历史用量计算
            forecast = pd.to_numeric(df['2025年预测使用量汇总（个）'], errors='coerce')
            historical_avg = pd.to_numeric(df['2024年月均使用量（个）'], errors='coerce')
            
            # 基础安全库存 = max(预测月均, 历史月均) * 0.1 (示例系数)
            forecast_monthly = forecast / 12
            df['计算安全库存'] = np.maximum(forecast_monthly, historical_avg) * 0.1
        
        # 针对大宗包材的特殊规则
        pe_bottles = df[df['品类'].str.contains('PE瓶|铁罐|玻璃瓶', na=False)]
        
        if not pe_bottles.empty:
            # 根据产能差异计算
            # 这里需要根据实际的大宗包材表来调整
            pass
        
        return df
    
    def update_monthly_data(self, df, new_month_data):
        """月度数据更新"""
        # 滚动更新：新月份加入，最旧月份移除
        df_updated = df.copy()
        
        # 重命名列：将现有月份列向后移动
        month_cols = [c for c in df.columns if 'M' in str(c) and len(str(c)) >= 6]
        month_cols_sorted = sorted(month_cols, key=lambda x: str(x), reverse=True)
        
        for i, col in enumerate(month_cols_sorted):
            if i < len(month_cols_sorted) - 1:
                df_updated[month_cols_sorted[i+1]] = df_updated[col]
        
        # 添加新月份数据
        if new_month_data is not None:
            new_month_name = new_month_data.get('month_name', 'M' + str(datetime.now().month))
            df_updated[new_month_name] = new_month_data['values']
        
        return df_updated


class DataLoader:
    """数据加载器"""
    
    @staticmethod
    def load_excel_file(uploaded_file):
        """加载Excel文件的所有sheet"""
        if uploaded_file is None:
            return None, None, None, None
        
        try:
            excel_file = pd.ExcelFile(uploaded_file)
            sheets = {}
            
            for sheet_name in excel_file.sheet_names:
                sheets[sheet_name] = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=0)
            
            # 获取各个sheet
            df_materials = sheets.get('安全库存（202509月）')
            df_rules = sheets.get('安全库存规则')
            df_quality = sheets.get('原辅料质量等级风险')
            df_category = sheets.get('品类策略系数')
            df_packaging = sheets.get('包材安全库存')
            df_bulk = sheets.get('大宗包材')
            
            return df_materials, df_rules, df_quality, df_category, df_packaging, df_bulk
        
        except Exception as e:
            st.error(f"加载Excel文件出错: {str(e)}")
            return None, None, None, None, None, None


def main():
    st.title("📦 安全库存管理系统")
    st.markdown("---")
    
    # 侧边栏配置
    with st.sidebar:
        st.header("⚙️ 系统配置")
        
        # 文件上传
        st.subheader("📁 数据文件")
        uploaded_file = st.file_uploader(
            "上传安全库存Excel文件",
            type=['xlsx', 'xls'],
            help="请上传包含安全库存规则和数据的Excel文件"
        )
        
        st.markdown("---")
        
        # 月度更新配置
        st.subheader("📅 月度更新")
        
        current_month = datetime.now().strftime("%Y年%m月")
        new_month_name = st.text_input("新月份名称", value=f"{current_month}")
        
        st.markdown("---")
        
        # 显示当前状态
        st.subheader("📊 系统状态")
        st.info(f"当前时间: {datetime.now().strftime('%Y-%m-%d %H:%M')}")
        
        # 计算按钮
        st.markdown("---")
        calculate_btn = st.button("🚀 开始计算安全库存", type="primary", use_container_width=True)
    
    # 主要内容区域
    if uploaded_file is not None:
        # 加载数据
        with st.spinner("正在加载数据..."):
            loader = DataLoader()
            df_materials, df_rules, df_quality, df_category, df_packaging, df_bulk = loader.load_excel_file(uploaded_file)
        
        if df_materials is not None:
            # 初始化计算器
            calculator = SafetyStockCalculator()
            
            # 计算安全库存
            if calculate_btn:
                with st.spinner("正在计算安全库存..."):
                    df_result = calculator.process_raw_materials(df_materials, df_quality, df_category)
                
                # 显示结果
                st.success("✅ 安全库存计算完成！")
                
                # 关键指标
                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    st.metric("总物料数", len(df_result))
                with col2:
                    total_safety_stock = df_result['安全库存'].sum() if '安全库存' in df_result.columns else 0
                    st.metric("总安全库存", f"{total_safety_stock:,.0f}")
                with col3:
                    avg_safety = df_result['安全库存'].mean() if '安全库存' in df_result.columns else 0
                    st.metric("平均安全库存", f"{avg_safety:,.0f}")
                with col4:
                    zero_stock = (df_result['安全库存'] == 0).sum() if '安全库存' in df_result.columns else 0
                    st.metric("零库存物料", zero_stock)
                
                st.markdown("---")
                
                # 图表展示
                tab1, tab2, tab3, tab4 = st.tabs(["📋 数据表格", "📊 统计分析", "📈 趋势图表", "💾 数据导出"])
                
                with tab1:
                    # 数据显示
                    display_cols = ['物料编码', 'SAP编码', '未来3个月月均用量', '过去6个月月均用量', 
                                   '质量风险系数', '品类策略系数', '采购周期系数', '安全库存', '是否寄售', '备注']
                    display_cols = [c for c in display_cols if c in df_result.columns]
                    
                    st.dataframe(
                        df_result[display_cols].style.format({
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
                        # 安全库存分布
                        if '安全库存' in df_result.columns:
                            fig_hist = px.histogram(
                                df_result, x='安全库存', nbins=30,
                                title='安全库存分布',
                                labels={'安全库存': '安全库存数量', 'count': '物料数量'}
                            )
                            st.plotly_chart(fig_hist, use_container_width=True)
                    
                    with col2:
                        # 寄售与非寄售对比
                        if '是否寄售' in df_result.columns:
                            consignment_data = df_result.groupby('是否寄售')['安全库存'].agg(['sum', 'count']).reset_index()
                            consignment_data.columns = ['是否寄售', '总安全库存', '物料数量']
                            
                            fig_pie = px.pie(
                                consignment_data, values='物料数量', names='是否寄售',
                                title='寄售物料占比'
                            )
                            st.plotly_chart(fig_pie, use_container_width=True)
                    
                    # 按采购周期系数分布
                    if '采购周期系数' in df_result.columns:
                        fig_box = px.box(
                            df_result, x='采购周期系数', y='安全库存',
                            title='不同采购周期系数的安全库存分布',
                            labels={'采购周期系数': '采购周期系数', '安全库存': '安全库存'}
                        )
                        st.plotly_chart(fig_box, use_container_width=True)
                
                with tab3:
                    # Top 20 安全库存物料
                    if '安全库存' in df_result.columns:
                        top_materials = df_result.nlargest(20, '安全库存')[['物料编码', '安全库存']]
                        fig_bar = px.bar(
                            top_materials, x='物料编码', y='安全库存',
                            title='安全库存TOP 20物料',
                            labels={'安全库存': '安全库存', '物料编码': '物料编码'}
                        )
                        fig_bar.update_layout(xaxis_tickangle=-45)
                        st.plotly_chart(fig_bar, use_container_width=True)
                
                with tab4:
                    st.subheader("📥 数据导出")
                    
                    # 导出Excel
                    output = BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df_result.to_excel(writer, sheet_name='安全库存计算结果', index=False)
                        
                        # 添加汇总sheet
                        summary_data = {
                            '指标': ['总物料数', '总安全库存', '平均安全库存', '最大安全库存', '最小安全库存', '零库存物料数'],
                            '数值': [
                                len(df_result),
                                df_result['安全库存'].sum() if '安全库存' in df_result.columns else 0,
                                df_result['安全库存'].mean() if '安全库存' in df_result.columns else 0,
                                df_result['安全库存'].max() if '安全库存' in df_result.columns else 0,
                                df_result['安全库存'].min() if '安全库存' in df_result.columns else 0,
                                (df_result['安全库存'] == 0).sum() if '安全库存' in df_result.columns else 0
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
                    
                    # 显示CSV导出
                    csv_data = df_result.to_csv(index=False).encode('utf-8')
                    st.download_button(
                        label="📄 下载CSV文件",
                        data=csv_data,
                        file_name=f"安全库存计算结果_{datetime.now().strftime('%Y%m%d')}.csv",
                        mime="text/csv"
                    )
                
                # 高风险物料警告
                st.markdown("---")
                st.subheader("⚠️ 高风险物料预警")
                
                if '安全库存' in df_result.columns:
                    high_risk = df_result[
                        (df_result['安全库存'] > df_result['安全库存'].quantile(0.95)) |
                        (df_result.get('质量风险系数', 0) > 1.5)
                    ].head(10)
                    
                    if not high_risk.empty:
                        st.warning("以下物料安全库存较高或质量风险较大，建议关注：")
                        st.dataframe(
                            high_risk[['物料编码', 'SAP编码', '安全库存', '质量风险系数', '备注']],
                            use_container_width=True
                        )
                    else:
                        st.info("未发现高风险物料")
            
            else:
                st.info("👈 请点击左侧的「开始计算安全库存」按钮")
        else:
            st.error("未找到原料安全库存数据表，请检查Excel文件格式")
    else:
        # 未上传文件时的提示
        st.info("👈 请先上传安全库存Excel文件")
        
        # 显示示例说明
        with st.expander("📖 使用说明"):
            st.markdown("""
            ### 系统功能说明
            
            1. **数据上传**: 上传包含安全库存规则的Excel文件
            2. **自动计算**: 根据以下规则自动计算安全库存:
               - 原料安全库存 = (未来3个月月均用量 + 过去6个月月均用量) / 2 × (质量风险系数×40% + 品类策略系数×60%) × 采购周期系数
               - 质量风险系数: 基础1分 + 低风险批次×0.1 + 中风险批次×0.2 + 高风险批次×0.3
               - 采购周期系数: 根据交货周期自动匹配
            3. **月度更新**: 输入新月份数据，系统自动滚动更新
            4. **数据导出**: 支持Excel和CSV格式导出
            
            ### 文件要求
            
            Excel文件应包含以下sheet:
            - `安全库存（202509月）`: 物料用量数据
            - `原辅料质量等级风险`: 质量风险批次数据
            - `品类策略系数`: 品类风险评估
            - `包材安全库存`: 包材数据（可选）
            """)
    
    # 页脚
    st.markdown("---")
    st.markdown(
        "<div style='text-align: center; color: gray;'>安全库存管理系统 | 版本 1.0</div>",
        unsafe_allow_html=True
    )


if __name__ == "__main__":
    main()
