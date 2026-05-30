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
    st.title("Safety Stock Management System")
    st.caption("Auto Calculate | Smart Alert | Monthly Update")
    st.divider()
    
    if 'result' not in st.session_state:
        st.session_state.result = None
    
    with st.sidebar:
        st.header("Settings")
        
        uploaded_file = st.file_uploader(
            "Upload Excel File",
            type=['xlsx', 'xls']
        )
        
        st.divider()
        
        st.subheader("Monthly Update")
        new_month_name = st.text_input("New Month Name", value=f"{datetime.now().year}-{datetime.now().month}")
        update_file = st.file_uploader(
            "Upload New Month Data",
            type=['xlsx', 'xls'],
            key="monthly"
        )
        
        st.divider()
        
        calc_btn = st.button("Calculate", type="primary", use_container_width=True)
        update_btn = st.button("Update", use_container_width=True)
        
        st.divider()
        
        with st.expander("Formula"):
            st.markdown("""
            Safety Stock = (Future 3M Avg + Past 6M Avg) / 2 
            x (Quality x 0.4 + Category x 0.6) 
            x Lead Time Coef
            
            Quality = 1 + Lowx0.1 + Midx0.2 + Highx0.3
            
            Lead Time Coef:
            1-7d: 0.2 | 8-15d: 0.3 | 16-21d: 0.6 | 22-30d: 1.0
            31-40d: 1.2 | 41-45d: 1.5 | 45-60d: 2.0 | >60d: 3.0
            """)
    
    if uploaded_file is not None:
        try:
            with st.spinner("Loading..."):
                excel_file = pd.ExcelFile(uploaded_file)
                sheets = {}
                for sheet_name in excel_file.sheet_names:
                    sheets[sheet_name] = pd.read_excel(uploaded_file, sheet_name=sheet_name)
            
            df_mat = sheets.get('安全库存（202509月）')
            df_qual = sheets.get('原辅料质量等级风险')
            df_cat = sheets.get('品类策略系数')
            
            if df_mat is None:
                st.error("Sheet '安全库存（202509月）' not found")
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
                
                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    st.metric("Total Materials", f"{total:,}")
                with col2:
                    st.metric("Low Stock Alert", f"{low:,}", f"{low_pct}%")
                with col3:
                    st.metric("Avg Safety Stock", f"{avg_ss:,.0f}")
                with col4:
                    st.metric("Avg Coverage", f"{avg_cov:.1f}x")
                
                st.divider()
                
                high_risk = df_result[df_result['coverage'] < 1].nlargest(10, 'safety_stock')
                if not high_risk.empty:
                    st.warning("Critical: Materials with insufficient stock")
                    cols_to_show = [df_result.columns[0], 'actual_stock', 'safety_stock', 'coverage', 'warning']
                    cols_to_show = [c for c in cols_to_show if c in high_risk.columns]
                    st.dataframe(high_risk[cols_to_show], use_container_width=True, height=200)
                    st.divider()
                
                tab1, tab2, tab3 = st.tabs(["Data Table", "Charts", "Export"])
                
                with tab1:
                    display_cols = [df_result.columns[0], 'future_avg', 'past_avg', 
                                   'quality_risk', 'category_risk', 'lead_coef', 
                                   'safety_stock', 'actual_stock', 'coverage', 'warning']
                    display_cols = [c for c in display_cols if c in df_result.columns]
                    st.dataframe(df_result[display_cols].head(100), use_container_width=True, height=500)
                
                with tab2:
                    col1, col2 = st.columns(2)
                    with col1:
                        warning_counts = df_result['warning'].value_counts().reset_index()
                        warning_counts.columns = ['Status', 'Count']
                        fig_pie = px.pie(warning_counts, values='Count', names='Status', title='Stock Status')
                        fig_pie.update_layout(height=400)
                        st.plotly_chart(fig_pie, use_container_width=True)
                    with col2:
                        top_ss = df_result.nlargest(15, 'safety_stock')[[df_result.columns[0], 'safety_stock']]
                        fig_bar = px.bar(top_ss, x=df_result.columns[0], y='safety_stock', title='Top 15 Safety Stock')
                        fig_bar.update_layout(height=400, xaxis_tickangle=-45)
                        st.plotly_chart(fig_bar, use_container_width=True)
                
                with tab3:
                    output = BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df_result.to_excel(writer, sheet_name='Safety_Stock', index=False)
                        summary = pd.DataFrame({
                            'Metric': ['Total Materials', 'Total Safety Stock', 'Avg Safety Stock', 
                                      'Total Actual Stock', 'Low Stock Count', 'Low Stock Pct', 'Avg Coverage'],
                            'Value': [
                                len(df_result), f"{df_result['safety_stock'].sum():,.0f}",
                                f"{df_result['safety_stock'].mean():,.0f}", f"{df_result['actual_stock'].sum():,.0f}",
                                len(df_result[df_result['coverage'] < 1.5]),
                                f"{len(df_result[df_result['coverage'] < 1.5]) / len(df_result) * 100:.1f}%",
                                f"{df_result['coverage'].mean():.1f}x"
                            ]
                        })
                        summary.to_excel(writer, sheet_name='Summary', index=False)
                    output.seek(0)
                    st.download_button("Download Excel", data=output, file_name=f"safety_stock_{datetime.now().strftime('%Y%m%d')}.xlsx")
            
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
                st.info("Click 'Calculate' to refresh")
            else:
                st.info("Click 'Calculate' to start")
                
        except Exception as e:
            st.error(f"Error: {str(e)}")
    else:
        st.info("Please upload an Excel file to begin")
        
        with st.expander("Instructions"):
            st.markdown("""
            ### How to Use
            1. Upload Excel file in the sidebar
            2. Click 'Calculate'
            3. View results and export
            
            ### Required Sheet
            - `安全库存（202509月）`: Material usage data
            """)
    
    st.divider()
    st.caption("Safety Stock Management System | Version 2.0")


if __name__ == "__main__":
    main()
