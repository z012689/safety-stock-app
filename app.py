import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO

st.set_page_config(
    page_title="安全库存管理系统",
    page_icon="package",
    layout="wide"
)


class SafetyStockCalculator:
    
    @staticmethod
    def get_lead_time_coef(days):
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
    def get_quality_score(df_quality, material_code):
        if df_quality is None or df_quality.empty:
            return 1.0
        try:
            matched = df_quality[df_quality.iloc[:, 0].astype(str) == str(material_code)]
            if matched.empty:
                return 1.0
            row = matched.iloc[0]
            low = float(row.iloc[1]) if len(row) > 1 and pd.notna(row.iloc[1]) else 0
            mid = float(row.iloc[2]) if len(row) > 2 and pd.notna(row.iloc[2]) else 0
            high = float(row.iloc[3]) if len(row) > 3 and pd.notna(row.iloc[3]) else 0
            return 1.0 + low * 0.1 + mid * 0.2 + high * 0.3
        except:
            return 1.0
    
    @staticmethod
    def get_category_risk(df_category, material_code):
        if df_category is None or df_category.empty:
            return 1.0
        try:
            code_col = df_category.columns[0]
            risk_col = df_category.columns[1] if len(df_category.columns) > 1 else df_category.columns[0]
            matched = df_category[df_category[code_col].astype(str) == str(material_code)]
            if not matched.empty:
                val = matched.iloc[0][risk_col]
                if pd.notna(val):
                    return float(val)
            return 1.0
        except:
            return 1.0
    
    @staticmethod
    def calc_safety_stock(future_avg, past_avg, quality, category, lead_coef):
        if pd.isna(future_avg) or pd.isna(past_avg):
            return 0
        if future_avg <= 0 and past_avg <= 0:
            return 0
        base = (float(future_avg) + float(past_avg)) / 2
        combined = quality * 0.4 + category * 0.6
        return base * combined * lead_coef
    
    def process(self, df_mat, df_qual, df_cat):
        df = df_mat.copy()
        
        future_cols = []
        past_cols = []
        for col in df.columns:
            col_str = str(col)
            if 'M07' in col_str or 'M08' in col_str or 'M09' in col_str:
                future_cols.append(col)
            elif 'M01' in col_str or 'M02' in col_str or 'M03' in col_str or 'M04' in col_str or 'M05' in col_str or 'M06' in col_str:
                past_cols.append(col)
        
        if future_cols:
            df['future_avg'] = df[future_cols].apply(pd.to_numeric, errors='coerce').mean(axis=1)
        elif '未来3个月月均用量' in df.columns:
            df['future_avg'] = pd.to_numeric(df['未来3个月月均用量'], errors='coerce')
        else:
            df['future_avg'] = 0
        
        if past_cols:
            df['past_avg'] = df[past_cols].apply(pd.to_numeric, errors='coerce').mean(axis=1)
        elif '月均量(半年)' in df.columns:
            df['past_avg'] = pd.to_numeric(df['月均量(半年)'], errors='coerce')
        else:
            df['past_avg'] = 0
        
        if '平均交货周期(天)' in df.columns:
            df['lead_coef'] = df['平均交货周期(天)'].apply(self.get_lead_time_coef)
        else:
            df['lead_coef'] = 1
        
        quality_list = []
        category_list = []
        code_col = df.columns[0]
        for idx, row in df.iterrows():
            code = row[code_col]
            quality_list.append(self.get_quality_score(df_qual, code))
            category_list.append(self.get_category_risk(df_cat, code))
        
        df['quality_risk'] = quality_list
        df['category_risk'] = category_list
        
        df['safety_stock'] = df.apply(
            lambda x: self.calc_safety_stock(
                x.get('future_avg', 0),
                x.get('past_avg', 0),
                x.get('quality_risk', 1),
                x.get('category_risk', 1),
                x.get('lead_coef', 1)
            ), axis=1
        )
        
        if '是否寄售' in df.columns:
            df.loc[df['是否寄售'] == '寄售', 'safety_stock'] = 0
        
        actual_col = None
        for col in df.columns:
            if '实际库存' in col or '6月末' in col:
                actual_col = col
                break
        if actual_col:
            df['actual_stock'] = pd.to_numeric(df[actual_col], errors='coerce').fillna(0)
        else:
            df['actual_stock'] = 0
        
        df['coverage'] = df.apply(
            lambda x: x['actual_stock'] / x['safety_stock'] if x['safety_stock'] > 0 else 999,
            axis=1
        )
        df['coverage'] = df['coverage'].replace([np.inf, -np.inf], 999)
        
        def get_warning(cov):
            if cov >= 3:
                return "充足"
            elif cov >= 1.5:
                return "正常"
            elif cov >= 0.5:
                return "偏低"
            else:
                return "严重不足"
        
        df['warning'] = df['coverage'].apply(get_warning)
        
        return df


def main():
    now = datetime.now()
    current_date = now.strftime("%Y年%m月%d日")
    
    st.title("安全库存管理系统")
    st.caption("自动计算 | 智能预警 | 月度更新")
    st.markdown(f"**当前日期：{current_date}**")
    st.divider()
    
    if 'result' not in st.session_state:
        st.session_state.result = None
    
    # 侧边栏
    with st.sidebar:
        st.header("系统配置")
        
        uploaded_file = st.file_uploader(
            "上传Excel文件",
            type=['xlsx', 'xls']
        )
        
        st.divider()
        
        st.subheader("月度数据更新")
        new_month_name = st.text_input("新月份名称", value=f"{now.year}年{now.month}月")
        update_file = st.file_uploader(
            "上传新月份数据文件",
            type=['xlsx', 'xls'],
            key="monthly"
        )
        
        st.divider()
        
        with st.expander("计算公式说明"):
            st.markdown("""
            安全库存 = (未来3个月月均 + 过去6个月月均) / 2 
            x (质量风险 x 0.4 + 品类策略 x 0.6) 
            x 采购周期系数
            
            质量风险 = 1 + 低风险x0.1 + 中风险x0.2 + 高风险x0.3
            
            采购周期系数:
            1-7天: 0.2
            8-15天: 0.3
            16-21天: 0.6
            22-30天: 1.0
            31-40天: 1.2
            41-45天: 1.5
            45-60天: 2.0
            60天以上: 3.0
            """)
    
    # 按钮区域
    col1, col2 = st.columns(2)
    with col1:
        calc_btn = st.button("开始计算安全库存", type="primary", use_container_width=True)
    with col2:
        update_btn = st.button("执行月度更新", use_container_width=True)
    
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
                st.error("未找到工作表：安全库存（202509月）")
                st.stop()
            
            calculator = SafetyStockCalculator()
            
            if calc_btn:
                with st.spinner("正在计算安全库存..."):
                    df_result = calculator.process(df_mat, df_qual, df_cat)
                    st.session_state.result = df_result
                
                st.success("安全库存计算完成！")
                st.balloons()
                
                # KPI指标
                total = len(df_result)
                low = len(df_result[df_result['coverage'] < 1.5])
                low_pct = round(low / total * 100, 1)
                avg_ss = df_result['safety_stock'].mean()
                total_actual = df_result['actual_stock'].sum()
                avg_cov = df_result['coverage'].mean()
                
                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    st.metric("物料总览", f"{total:,}", "个物料")
                with col2:
                    st.metric("低库存预警物料", f"{low:,}", f"占总数 {low_pct}%")
                with col3:
                    st.metric("平均安全库存", f"{avg_ss:,.0f}", "件/物料")
                with col4:
                    st.metric("平均库存覆盖", f"{avg_cov:.1f}倍", "充足" if avg_cov >= 1.5 else "偏低")
                
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.metric("总实际库存", f"{total_actual:,.0f}", "所有物料合计")
                with col2:
                    total_ss = df_result['safety_stock'].sum()
                    st.metric("建议库存总量", f"{total_ss:,.0f}", "基于公式计算")
                with col3:
                    st.metric("包材库存", f"{avg_cov:.1f}倍", "库存充足" if avg_cov >= 1.5 else "库存偏低")
                
                st.divider()
                
                # 高风险预警
                high_risk = df_result[df_result['coverage'] < 1].nlargest(10, 'safety_stock')
                if not high_risk.empty:
                    st.warning("高风险物料预警：以下物料库存严重不足，建议立即采购")
                    cols_to_show = [df_result.columns[0], 'actual_stock', 'safety_stock', 'coverage', 'warning']
                    cols_to_show = [c for c in cols_to_show if c in high_risk.columns]
                    st.dataframe(
                        high_risk[cols_to_show].style.format({
                            'actual_stock': '{:,.0f}',
                            'safety_stock': '{:,.0f}',
                            'coverage': '{:.1f}'
                        }),
                        use_container_width=True,
                        height=250
                    )
                    st.divider()
                
                # Tab页面
                tab1, tab2, tab3 = st.tabs(["数据详情", "可视化分析", "数据导出"])
                
                with tab1:
                    st.caption(f"共 {len(df_result)} 条记录")
                    display_cols = [df_result.columns[0], 'future_avg', 'past_avg',
                                   'quality_risk', 'category_risk', 'lead_coef',
                                   'safety_stock', 'actual_stock', 'coverage', 'warning']
                    display_cols = [c for c in display_cols if c in df_result.columns]
                    st.dataframe(df_result[display_cols].head(100), use_container_width=True, height=500)
                
                with tab2:
                    col1, col2 = st.columns(2)
                    with col1:
                        warning_counts = df_result['warning'].value_counts().reset_index()
                        warning_counts.columns = ['状态', '数量']
                        fig_pie = px.pie(warning_counts, values='数量', names='状态', title='库存状态分布')
                        fig_pie.update_layout(height=400)
                        st.plotly_chart(fig_pie, use_container_width=True)
                    with col2:
                        top_ss = df_result.nlargest(15, 'safety_stock')[[df_result.columns[0], 'safety_stock']]
                        fig_bar = px.bar(top_ss, x=df_result.columns[0], y='safety_stock', title='安全库存 TOP 15')
                        fig_bar.update_layout(height=400, xaxis_tickangle=-45)
                        st.plotly_chart(fig_bar, use_container_width=True)
                    
                    fig_hist = px.histogram(
                        df_result[df_result['coverage'] < 10],
                        x='coverage', nbins=30,
                        title='库存覆盖倍数分布'
                    )
                    fig_hist.update_layout(height=400)
                    st.plotly_chart(fig_hist, use_container_width=True)
                
                with tab3:
                    st.subheader("导出数据")
                    
                    output = BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df_result.to_excel(writer, sheet_name='安全库存计算结果', index=False)
                        
                        summary = pd.DataFrame({
                            '指标': ['总物料数', '总安全库存', '平均安全库存', '总实际库存',
                                   '低库存物料数', '低库存占比', '平均库存覆盖', '计算时间'],
                            '数值': [
                                len(df_result),
                                f"{df_result['safety_stock'].sum():,.0f}",
                                f"{df_result['safety_stock'].mean():,.0f}",
                                f"{df_result['actual_stock'].sum():,.0f}",
                                len(df_result[df_result['coverage'] < 1.5]),
                                f"{len(df_result[df_result['coverage'] < 1.5]) / len(df_result) * 100:.1f}%",
                                f"{df_result['coverage'].mean():.1f}倍",
                                datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                            ]
                        })
                        summary.to_excel(writer, sheet_name='汇总统计', index=False)
                        
                        if not high_risk.empty:
                            high_risk.to_excel(writer, sheet_name='高风险物料', index=False)
                    
                    output.seek(0)
                    st.download_button(
                        label="下载Excel报告",
                        data=output,
                        file_name=f"安全库存报告_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    
                    st.info("报告包含：安全库存计算结果、汇总统计、高风险物料")
            
            elif st.session_state.result is not None:
                df_result = st.session_state.result
                total = len(df_result)
                low = len(df_result[df_result['coverage'] < 1.5])
                avg_ss = df_result['safety_stock'].mean()
                avg_cov = df_result['coverage'].mean()
                
                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    st.metric("物料总览", f"{total:,}")
                with col2:
                    st.metric("低库存预警", f"{low:,}")
                with col3:
                    st.metric("平均安全库存", f"{avg_ss:,.0f}")
                with col4:
                    st.metric("平均库存覆盖", f"{avg_cov:.1f}倍")
                
                st.info("如需重新计算，请点击上方「开始计算安全库存」按钮")
            else:
                st.info("请点击上方「开始计算安全库存」按钮开始计算")
                
        except Exception as e:
            st.error(f"处理文件时出错: {str(e)}")
    else:
        st.info("请在左侧上传Excel文件")
        
        with st.expander("使用说明"):
            st.markdown("""
            ### 使用方法
            1. 在左侧上传Excel文件
            2. 点击「开始计算安全库存」按钮
            3. 查看结果并导出报告
            
            ### 必需工作表
            - `安全库存（202509月）`: 物料用量数据
            """)
    
    st.divider()
    st.caption("安全库存管理系统 | 版本 3.0 | 支持月度数据更新")


if __name__ == "__main__":
    main()
