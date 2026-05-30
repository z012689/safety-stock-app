import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO

st.set_page_config(
    page_title="Safety Stock Management",
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
                return "Sufficient"
            elif cov >= 1.5:
                return "Normal"
            elif cov >= 0.5:
                return "Low"
            else:
                return "Critical"
        
        df['warning'] = df['coverage'].apply(get_warning)
        
        return df


def main():
    now = datetime.now()
    
    st.title("Safety Stock Management System")
    st.caption("Auto Calculate | Smart Alert | Monthly Update")
    st.markdown(f"Date: {now.strftime('%Y-%m-%d')}")
    st.divider()
    
    if 'result' not in st.session_state:
        st.session_state.result = None
    if 'active_tab' not in st.session_state:
        st.session_state.active_tab = None
    
    with st.sidebar:
        st.header("Settings")
        
        uploaded_file = st.file_uploader(
            "Upload Excel File",
            type=['xlsx', 'xls']
        )
        
        st.divider()
        
        st.subheader("Monthly Update")
        new_month_name = st.text_input("New Month Name", value=f"{now.year}-{now.month}")
        update_file = st.file_uploader(
            "Upload New Month Data",
            type=['xlsx', 'xls'],
            key="monthly"
        )
        
        st.divider()
        
        with st.expander("Formula"):
            st.markdown("""
            Safety Stock = (Future 3M Avg + Past 6M Avg) / 2 
            x (Quality x 0.4 + Category x 0.6) 
            x Lead Time Coef
            
            Quality = 1 + Lowx0.1 + Midx0.2 + Highx0.3
            
            Lead Time Coef:
            1-7d: 0.2 | 8-15d: 0.3 | 16-21d: 0.6
            22-30d: 1.0 | 31-40d: 1.2 | 41-45d: 1.5
            45-60d: 2.0 | >60d: 3.0
            """)
    
    col1, col2 = st.columns(2)
    with col1:
        calc_btn = st.button("Calculate Safety Stock", type="primary", use_container_width=True)
    with col2:
        update_btn = st.button("Monthly Update", use_container_width=True)
    
    st.divider()
    
    if uploaded_file is not None:
        try:
            with st.spinner("Loading data..."):
                excel_file = pd.ExcelFile(uploaded_file)
                sheets = {}
                for sheet_name in excel_file.sheet_names:
                    sheets[sheet_name] = pd.read_excel(uploaded_file, sheet_name=sheet_name)
            
            df_mat = sheets.get('安全库存（202509月）')
            df_qual = sheets.get('原辅料质量等级风险')
            df_cat = sheets.get('品类策略系数')
            
            if df_mat is None:
                st.error("Sheet not found: 安全库存（202509月）")
                st.stop()
            
            calculator = SafetyStockCalculator()
            
            if calc_btn:
                with st.spinner("Calculating..."):
                    df_result = calculator.process(df_mat, df_qual, df_cat)
                    st.session_state.result = df_result
                
                st.success("Calculation completed!")
                st.balloons()
                
                total = len(df_result)
                low = len(df_result[df_result['coverage'] < 1.5])
                low_pct = round(low / total * 100, 1)
                avg_ss = df_result['safety_stock'].mean()
                total_actual = df_result['actual_stock'].sum()
                avg_cov = df_result['coverage'].mean()
                
                col1, col2, col3, col4, col5 = st.columns(5)
                metrics = [
                    ("Total Materials", f"{total:,}"),
                    ("Low Stock Alert", f"{low:,}", f"{low_pct}%"),
                    ("Avg Safety Stock", f"{avg_ss:,.0f}"),
                    ("Total Actual Stock", f"{total_actual:,.0f}"),
                    ("Avg Coverage", f"{avg_cov:.1f}x")
                ]
                
                for col, metric in zip([col1, col2, col3, col4, col5], metrics):
                    with col:
                        st.metric(metric[0], metric[1], metric[2] if len(metric) > 2 else None)
                
                st.divider()
                
                # Four function buttons
                col_a, col_b, col_c, col_d = st.columns(4)
                with col_a:
                    show_alert = st.button("Warning", use_container_width=True)
                with col_b:
                    show_data = st.button("Data Table", use_container_width=True)
                with col_c:
                    show_chart = st.button("Charts", use_container_width=True)
                with col_d:
                    show_export = st.button("Export", use_container_width=True)
                
                st.divider()
                
                high_risk = df_result[df_result['coverage'] < 1].nlargest(10, 'safety_stock')
                
                if show_alert or st.session_state.active_tab == 'alert':
                    st.session_state.active_tab = 'alert'
                    st.subheader("High Risk Materials")
                    st.caption("These materials have insufficient stock, recommend immediate purchase")
                    
                    if not high_risk.empty:
                        alert_df = high_risk[['物料编码', 'actual_stock', 'safety_stock', 'coverage', 'warning']].copy()
                        alert_df.columns = ['Material Code', 'Actual Stock', 'Safety Stock', 'Coverage', 'Status']
                        st.dataframe(
                            alert_df.style.format({
                                'Actual Stock': '{:,.0f}',
                                'Safety Stock': '{:,.0f}',
                                'Coverage': '{:.1f}'
                            }),
                            use_container_width=True,
                            height=300
                        )
                    else:
                        st.success("No high risk materials found!")
                
                if show_data or st.session_state.active_tab == 'data':
                    st.session_state.active_tab = 'data'
                    st.subheader("Data Details")
                    st.caption(f"Total {len(df_result)} records")
                    
                    detail_df = df_result[[
                        '物料编码', 'future_avg', 'past_avg',
                        'quality_risk', 'category_risk', 'lead_coef',
                        'safety_stock', 'actual_stock', 'coverage', 'warning'
                    ]].copy()
                    
                    detail_df.columns = [
                        'Material Code', 'Future 3M Avg', 'Past 6M Avg',
                        'Quality Risk', 'Category Risk', 'Lead Time Coef',
                        'Safety Stock', 'Actual Stock', 'Coverage', 'Status'
                    ]
                    
                    st.dataframe(
                        detail_df.head(100).style.format({
                            'Future 3M Avg': '{:,.0f}',
                            'Past 6M Avg': '{:,.0f}',
                            'Safety Stock': '{:,.0f}',
                            'Actual Stock': '{:,.0f}',
                            'Coverage': '{:.1f}'
                        }),
                        use_container_width=True,
                        height=500
                    )
                
                if show_chart or st.session_state.active_tab == 'chart':
                    st.session_state.active_tab = 'chart'
                    st.subheader("Charts")
                    
                    chart1, chart2 = st.columns(2)
                    
                    with chart1:
                        warning_counts = df_result['warning'].value_counts().reset_index()
                        warning_counts.columns = ['Status', 'Count']
                        fig_pie = px.pie(
                            warning_counts, values='Count', names='Status',
                            title='Stock Status Distribution'
                        )
                        fig_pie.update_layout(height=400)
                        st.plotly_chart(fig_pie, use_container_width=True)
                    
                    with chart2:
                        top_ss = df_result.nlargest(15, 'safety_stock')[['物料编码', 'safety_stock']]
                        fig_bar = px.bar(
                            top_ss, x='物料编码', y='safety_stock',
                            title='Top 15 Safety Stock'
                        )
                        fig_bar.update_layout(height=400, xaxis_tickangle=-45)
                        st.plotly_chart(fig_bar, use_container_width=True)
                    
                    fig_hist = px.histogram(
                        df_result[df_result['coverage'] < 10],
                        x='coverage', nbins=30,
                        title='Coverage Distribution'
                    )
                    fig_hist.update_layout(height=400)
                    st.plotly_chart(fig_hist, use_container_width=True)
                
                if show_export or st.session_state.active_tab == 'export':
                    st.session_state.active_tab = 'export'
                    st.subheader("Export Data")
                    
                    output = BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df_result.to_excel(writer, sheet_name='Safety_Stock_Result', index=False)
                        
                        summary = pd.DataFrame({
                            'Metric': ['Total Materials', 'Total Safety Stock', 'Avg Safety Stock',
                                      'Total Actual Stock', 'Low Stock Count', 'Low Stock Pct', 
                                      'Avg Coverage', 'Calculation Time'],
                            'Value': [
                                len(df_result),
                                f"{df_result['safety_stock'].sum():,.0f}",
                                f"{df_result['safety_stock'].mean():,.0f}",
                                f"{df_result['actual_stock'].sum():,.0f}",
                                len(df_result[df_result['coverage'] < 1.5]),
                                f"{len(df_result[df_result['coverage'] < 1.5]) / len(df_result) * 100:.1f}%",
                                f"{df_result['coverage'].mean():.1f}x",
                                datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                            ]
                        })
                        summary.to_excel(writer, sheet_name='Summary', index=False)
                        
                        if not high_risk.empty:
                            high_risk.to_excel(writer, sheet_name='High_Risk_Materials', index=False)
                    
                    output.seek(0)
                    st.download_button(
                        label="Download Excel Report",
                        data=output,
                        file_name=f"safety_stock_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    
                    st.info("Report includes: Safety Stock Result, Summary, High Risk Materials")
            
            elif st.session_state.result is not None:
                df_result = st.session_state.result
                total = len(df_result)
                low = len(df_result[df_result['coverage'] < 1.5])
                avg_ss = df_result['safety_stock'].mean()
                avg_cov = df_result['coverage'].mean()
                
                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    st.metric("Total Materials", f"{total:,}")
                with col2:
                    st.metric("Low Stock Alert", f"{low:,}")
                with col3:
                    st.metric("Avg Safety Stock", f"{avg_ss:,.0f}")
                with col4:
                    st.metric("Avg Coverage", f"{avg_cov:.1f}x")
                
                st.info("Click 'Calculate Safety Stock' to refresh")
            else:
                st.info("Click 'Calculate Safety Stock' to start")
                
        except Exception as e:
            st.error(f"Error: {str(e)}")
    else:
        st.info("Please upload an Excel file to begin")
        
        with st.expander("Instructions"):
            st.markdown("""
            ### How to Use
            1. Upload Excel file in the sidebar
            2. Click 'Calculate Safety Stock'
            3. Use buttons to view details
            
            ### Required Sheet
            - `安全库存（202509月）`: Material usage data
            """)
    
    st.divider()
    st.caption("Safety Stock Management System | Version 3.0")


if __name__ == "__main__":
    main()
