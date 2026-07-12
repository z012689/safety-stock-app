# app.py - 适配您的Excel格式
import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import io
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

st.set_page_config(
    page_title="安全库存管理系统",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
    <style>
    .main-header {
        font-size: 2.2rem;
        font-weight: 700;
        color: #1f77b4;
        margin-bottom: 0.5rem;
        text-align: center;
    }
    .file-status {
        padding: 0.25rem 0.5rem;
        border-radius: 0.25rem;
        font-size: 0.9rem;
    }
    .file-uploaded {
        background-color: #d4edda;
        color: #155724;
    }
    .file-missing {
        background-color: #f8d7da;
        color: #721c24;
    }
    .file-optional {
        background-color: #fff3cd;
        color: #856404;
    }
    .monthly-update-badge {
        background-color: #28a745;
        color: white;
        padding: 0.2rem 0.8rem;
        border-radius: 1rem;
        font-size: 0.8rem;
        display: inline-block;
    }
    </style>
""", unsafe_allow_html=True)

@st.cache_data
def load_safety_stock(file):
    """加载安全库存模板 - 适配您的Excel格式"""
    try:
        df = pd.read_excel(file)
        
        # 统一物料编码为字符串
        if '物料编码' in df.columns:
            df['物料编码'] = df['物料编码'].astype(str)
        
        # 检查是否有M01~M09列（原始历史用量）
        month_cols = ['M01', 'M02', 'M03', 'M04', 'M05', 'M06', 'M07', 'M08', 'M09']
        has_month_cols = all(col in df.columns for col in month_cols)
        
        if not has_month_cols:
            # 如果Excel没有M01~M09，尝试使用已有的计算字段
            st.info("📌 检测到您的Excel已包含计算结果，直接使用现有数据")
            
            # 映射列名
            column_mapping = {}
            
            # 未来3个月月均用量
            if '未来3个月月均用量' in df.columns:
                pass  # 已经有这个列
            elif '未来3月月均用量' in df.columns:
                df.rename(columns={'未来3月月均用量': '未来3个月月均用量'}, inplace=True)
            
            # 过去6个月月均用量
            if '月均量(半年)' in df.columns:
                df.rename(columns={'月均量(半年)': '过去6个月月均用量'}, inplace=True)
            
            # 平均交货周期
            if '平均交货周期(天)' in df.columns:
                # 将天数转换为周数
                df['平均交货周期'] = df['平均交货周期(天)'].apply(
                    lambda x: f"{int(round(x/7, 0))}周" if pd.notna(x) else '2周'
                )
            
            # 质量风险系数
            if '原辅料质量等级风险' in df.columns:
                df.rename(columns={'原辅料质量等级风险': '质量风险系数'}, inplace=True)
            
            # 品类策略系数
            if '品类策略风险系数' in df.columns:
                df.rename(columns={'品类策略风险系数': '品类策略系数'}, inplace=True)
            
            # 采购周期系数
            if '采购周期系数' in df.columns:
                pass  # 已有
            
            # 如果没有安全库存量，计算它
            if '安全库存量' not in df.columns:
                # 使用现有数据计算
                quality_risk_map = {'高': 1.5, '中': 1.2, '低': 1.0}
                if '质量风险系数' in df.columns:
                    df['质量风险系数_数值'] = df['质量风险系数'].map(quality_risk_map).fillna(1.0)
                else:
                    df['质量风险系数_数值'] = 1.0
                
                category_map = {'A类': 1.3, 'B类': 1.1, 'C类': 0.9, 'D类': 0.7}
                if '品类策略系数' in df.columns:
                    df['品类策略系数_数值'] = df['品类策略系数'].map(category_map).fillna(1.0)
                else:
                    df['品类策略系数_数值'] = 1.0
                
                if '未来3个月月均用量' in df.columns and '平均交货周期' in df.columns:
                    lead_time_map = {'1周': 1.2, '2周': 1.4, '3周': 1.6, '4周': 1.8, '5周': 2.0, '6周': 2.2}
                    df['采购周期系数_数值'] = df['平均交货周期'].map(lead_time_map).fillna(1.5)
                    
                    df['安全库存量'] = (
                        df['未来3个月月均用量'] * 
                        df['采购周期系数_数值'] * 
                        df['质量风险系数_数值'] * 
                        df['品类策略系数_数值']
                    ).round(2)
        
        return df
        
    except Exception as e:
        st.error(f"加载安全库存文件失败: {e}")
        return None

@st.cache_data
def load_inventory(file):
    """加载6.1库存文件"""
    try:
        df = pd.read_excel(file, sheet_name='Data')
        df_qualified = df[df['存储地点'] == 3006].copy()
        inventory_summary = df_qualified.groupby('物料')['非限制使用的库存'].sum().reset_index()
        inventory_summary.columns = ['物料编码', '实际库存']
        inventory_summary['物料编码'] = inventory_summary['物料编码'].astype(str)
        return inventory_summary
    except Exception as e:
        st.error(f"加载库存文件失败: {e}")
        return None

@st.cache_data
def load_usage(file):
    """加载月度使用量文件"""
    try:
        df = pd.read_excel(file)
        if '物料编码' in df.columns and '日期' in df.columns:
            df['物料编码'] = df['物料编码'].astype(str)
            df['月份'] = pd.to_datetime(df['日期']).dt.month
            monthly_usage = df.groupby(['物料编码', '月份'])['消耗量'].sum().reset_index()
            days_in_month = {1:31, 2:28, 3:31, 4:30, 5:31, 6:30, 
                           7:31, 8:31, 9:30, 10:31, 11:30, 12:31}
            monthly_usage['日均消耗'] = monthly_usage.apply(
                lambda x: x['消耗量'] / days_in_month.get(x['月份'], 30), axis=1
            )
            latest_month = monthly_usage.groupby('物料编码')['月份'].max().reset_index()
            latest_month.columns = ['物料编码', '最新月份']
            result = latest_month.merge(monthly_usage, on=['物料编码', '月份'], how='left')
            return result[['物料编码', '日均消耗']]
        else:
            return None
    except Exception as e:
        st.error(f"加载使用量文件失败: {e}")
        return None

@st.cache_data
def load_forecast(file):
    """加载预测库存文件"""
    try:
        df = pd.read_excel(file)
        if '物料编码' in df.columns:
            df['物料编码'] = df['物料编码'].astype(str)
            forecast_col = [col for col in df.columns if '预测' in col or '库存' in col][0]
            forecast_summary = df[['物料编码', forecast_col]].copy()
            forecast_summary.columns = ['物料编码', '预测库存']
            return forecast_summary
        else:
            return None
    except Exception as e:
        st.error(f"加载预测文件失败: {e}")
        return None

def load_mock_data():
    """生成模拟数据"""
    np.random.seed(42)
    materials = [f'MAT-{i:04d}' for i in range(1, 51)]
    data = []
    for mat in materials:
        row = {
            '物料编码': mat,
            '物料描述': f'物料_{mat}',
            '工厂': np.random.choice(['广州工厂', '上海工厂', '北京工厂']),
            '存储地点': np.random.choice(['A区', 'B区', 'C区']),
            '单位': 'KG',
            '未来3个月月均用量': np.random.uniform(200, 600),
            '过去6个月月均用量': np.random.uniform(150, 500),
            '平均交货周期': np.random.choice(['1周', '2周', '3周', '4周']),
            '质量风险系数': np.random.choice(['高', '中', '低'], p=[0.2, 0.5, 0.3]),
            '品类策略系数': np.random.choice(['A类', 'B类', 'C类'], p=[0.2, 0.4, 0.4]),
        }
        # 计算安全库存量
        quality_risk_map = {'高': 1.5, '中': 1.2, '低': 1.0}
        category_map = {'A类': 1.3, 'B类': 1.1, 'C类': 0.9, 'D类': 0.7}
        lead_time_map = {'1周': 1.2, '2周': 1.4, '3周': 1.6, '4周': 1.8, '5周': 2.0, '6周': 2.2}
        
        row['安全库存量'] = (
            row['未来3个月月均用量'] * 
            lead_time_map.get(row['平均交货周期'], 1.5) *
            quality_risk_map.get(row['质量风险系数'], 1.0) *
            category_map.get(row['品类策略系数'], 1.0)
        )
        data.append(row)
    return pd.DataFrame(data)

def export_to_excel(df, template_columns=None):
    """导出Excel"""
    output = io.BytesIO()
    
    if template_columns and len(template_columns) > 0:
        export_cols = [col for col in template_columns if col in df.columns]
        if len(export_cols) == len(template_columns):
            export_df = df[export_cols].copy()
        else:
            export_df = df.copy()
    else:
        export_df = df.copy()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        export_df.to_excel(writer, sheet_name='计算结果', index=False)
        
        workbook = writer.book
        worksheet = writer.sheets['计算结果']
        
        for column in worksheet.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = min(max_length + 2, 30)
            worksheet.column_dimensions[column_letter].width = adjusted_width
        
        header_font = Font(bold=True, color="FFFFFF")
        header_fill = PatternFill(start_color="1f77b4", end_color="1f77b4", fill_type="solid")
        header_alignment = Alignment(horizontal="center", vertical="center")
        
        for cell in worksheet[1]:
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = header_alignment
    
    return output.getvalue()

def main():
    st.markdown('<div class="main-header">📦 安全库存管理系统</div>', unsafe_allow_html=True)
    
    current_month = datetime.now().strftime('%Y年%m月')
    st.markdown(f'<span class="monthly-update-badge">📅 {current_month} 月度报告</span>', unsafe_allow_html=True)
    st.markdown("---")
    
    if 'safety_data' not in st.session_state:
        st.session_state.safety_data = None
    if 'template_columns' not in st.session_state:
        st.session_state.template_columns = None
    if 'inventory_data' not in st.session_state:
        st.session_state.inventory_data = None
    if 'usage_data' not in st.session_state:
        st.session_state.usage_data = None
    if 'forecast_data' not in st.session_state:
        st.session_state.forecast_data = None
    if 'last_update' not in st.session_state:
        st.session_state.last_update = None
    
    with st.sidebar:
        st.markdown("## 📂 数据上传")
        
        if st.session_state.last_update:
            st.info(f"🕐 上次更新: {st.session_state.last_update}")
        
        if st.button("📊 加载示例数据"):
            st.session_state.safety_data = load_mock_data()
            st.session_state.template_columns = st.session_state.safety_data.columns.tolist()
            st.session_state.last_update = datetime.now().strftime('%Y-%m-%d %H:%M')
            st.rerun()
        
        st.markdown("---")
        
        st.markdown("### 🔴 必须上传")
        safety_stock_file = st.file_uploader(
            "① 安全库存（需计算）.xlsx",
            type=['xlsx', 'xls'],
            key='safety_stock_upload'
        )
        if safety_stock_file:
            st.markdown(f'<span class="file-status file-uploaded">✅ 已上传: {safety_stock_file.name}</span>', unsafe_allow_html=True)
        else:
            st.markdown('<span class="file-status file-missing">❌ 未上传</span>', unsafe_allow_html=True)
        
        st.markdown("---")
        
        st.markdown("### 🟡 推荐上传")
        inventory_file = st.file_uploader(
            "② X.1库存-去除名称.xlsx",
            type=['xlsx', 'xls'],
            key='inventory_upload'
        )
        if inventory_file:
            st.markdown(f'<span class="file-status file-uploaded">✅ 已上传: {inventory_file.name}</span>', unsafe_allow_html=True)
        else:
            st.markdown('<span class="file-status file-optional">⏭️ 将生成模拟数据</span>', unsafe_allow_html=True)
        
        st.markdown("---")
        
        st.markdown("### 🟢 可选上传")
        usage_file = st.file_uploader(
            "③ X月使用量.xlsx",
            type=['xlsx', 'xls'],
            key='usage_upload'
        )
        if usage_file:
            st.markdown(f'<span class="file-status file-uploaded">✅ 已上传: {usage_file.name}</span>', unsafe_allow_html=True)
        else:
            st.markdown('<span class="file-status file-optional">⏭️ 未上传</span>', unsafe_allow_html=True)
        
        st.markdown("---")
        
        st.markdown("### 🟢 可选上传")
        forecast_file = st.file_uploader(
            "④ 预测-X.1.xlsx",
            type=['xlsx', 'xls'],
            key='forecast_upload'
        )
        if forecast_file:
            st.markdown(f'<span class="file-status file-uploaded">✅ 已上传: {forecast_file.name}</span>', unsafe_allow_html=True)
        else:
            st.markdown('<span class="file-status file-optional">⏭️ 未上传</span>', unsafe_allow_html=True)
        
        st.markdown("---")
        
        if st.session_state.safety_data is not None:
            st.download_button(
                label="📥 导出安全库存计算结果",
                data=export_to_excel(
                    st.session_state.safety_data,
                    st.session_state.template_columns
                ),
                file_name=f"安全库存计算结果_{datetime.now().strftime('%Y%m')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
    
    if safety_stock_file is not None:
        try:
            df = load_safety_stock(safety_stock_file)
            if df is not None:
                st.session_state.template_columns = df.columns.tolist()
                st.session_state.safety_data = df
                st.session_state.last_update = datetime.now().strftime('%Y-%m-%d %H:%M')
                st.sidebar.success(f"✅ 已加载 {len(df)} 条物料数据")
        except Exception as e:
            st.error(f"数据加载失败: {e}")
            df = load_mock_data()
            st.session_state.template_columns = df.columns.tolist()
            st.session_state.safety_data = df
    else:
        if st.session_state.safety_data is None:
            df = load_mock_data()
            st.session_state.template_columns = df.columns.tolist()
            st.session_state.safety_data = df
            st.sidebar.info("ℹ️ 使用模拟数据")
    
    if inventory_file is not None:
        st.session_state.inventory_data = load_inventory(inventory_file)
    if usage_file is not None:
        st.session_state.usage_data = load_usage(usage_file)
    if forecast_file is not None:
        st.session_state.forecast_data = load_forecast(forecast_file)
    
    if st.session_state.safety_data is not None:
        df_display = st.session_state.safety_data.copy()
        if '物料编码' in df_display.columns:
            df_display['物料编码'] = df_display['物料编码'].astype(str)
        
        # 合并库存数据
        if st.session_state.inventory_data is not None:
            st.session_state.inventory_data['物料编码'] = st.session_state.inventory_data['物料编码'].astype(str)
            df_display = df_display.merge(st.session_state.inventory_data, on='物料编码', how='left')
            df_display['实际库存'] = df_display['实际库存'].fillna(0)
        else:
            np.random.seed(42)
            if '安全库存量' in df_display.columns:
                df_display['实际库存'] = df_display['安全库存量'] * np.random.uniform(0.5, 1.8, len(df_display))
            else:
                df_display['实际库存'] = 0
            df_display['实际库存'] = df_display['实际库存'].round(2)
        
        if st.session_state.forecast_data is not None:
            st.session_state.forecast_data['物料编码'] = st.session_state.forecast_data['物料编码'].astype(str)
            df_display = df_display.merge(st.session_state.forecast_data, on='物料编码', how='left')
            df_display['预测库存'] = df_display['预测库存'].fillna(0)
        else:
            np.random.seed(42)
            if '安全库存量' in df_display.columns:
                df_display['预测库存'] = df_display['安全库存量'] * np.random.uniform(0.6, 1.5, len(df_display))
            else:
                df_display['预测库存'] = 0
            df_display['预测库存'] = df_display['预测库存'].round(2)
        
        if '安全库存量' in df_display.columns and '实际库存' in df_display.columns:
            df_display['库存差异(实际-安全)'] = df_display['实际库存'] - df_display['安全库存量']
            df_display['库存状态'] = df_display['库存差异(实际-安全)'].apply(
                lambda x: '✅ 充足' if x >= 0 else '⚠️ 预警'
            )
        
        if st.session_state.usage_data is not None:
            st.session_state.usage_data['物料编码'] = st.session_state.usage_data['物料编码'].astype(str)
            df_display = df_display.merge(st.session_state.usage_data, on='物料编码', how='left')
            if '过去6个月月均用量' in df_display.columns:
                df_display['日均消耗'] = df_display['日均消耗'].fillna(df_display['过去6个月月均用量'] / 30)
            else:
                df_display['日均消耗'] = df_display['日均消耗'].fillna(10)
        else:
            if '未来3个月月均用量' in df_display.columns:
                df_display['日均消耗'] = df_display['未来3个月月均用量'] / 30
            else:
                df_display['日均消耗'] = 10
        
        tab1, tab2, tab3, tab4 = st.tabs([
            "📋 数据概览",
            "📊 库存状态分析",
            "⚠️ 预警列表",
            "📈 消耗趋势"
        ])
        
        with tab1:
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("📦 物料总数", len(df_display))
            with col2:
                if '安全库存量' in df_display.columns:
                    st.metric("📊 总安全库存", f"{df_display['安全库存量'].sum():,.0f}")
                else:
                    st.metric("📊 总安全库存", "N/A")
            with col3:
                if '未来3个月月均用量' in df_display.columns:
                    st.metric("📈 平均月用量", f"{df_display['未来3个月月均用量'].mean():,.0f}")
                else:
                    st.metric("📈 平均月用量", "N/A")
            with col4:
                if '库存状态' in df_display.columns:
                    warning_count = len(df_display[df_display['库存状态'] == '⚠️ 预警'])
                    st.metric("⚠️ 预警物料", warning_count, delta_color="inverse")
                else:
                    st.metric("⚠️ 预警物料", "N/A")
            
            st.markdown("---")
            
            display_cols = ['物料编码', '物料描述', '工厂', '存储地点', '单位',
                           '未来3个月月均用量', '过去6个月月均用量', '日均消耗',
                           '安全库存量', '实际库存', '预测库存',
                           '库存差异(实际-安全)', '库存状态']
            display_cols = [col for col in display_cols if col in df_display.columns]
            st.dataframe(df_display[display_cols], use_container_width=True, height=400)
        
        with tab2:
            st.markdown("### 📊 库存对比分析")
            top_n = st.slider("显示物料数量", 10, 50, 20)
            plot_df = df_display.head(top_n).copy()
            
            fig = go.Figure()
            if '安全库存量' in plot_df.columns:
                fig.add_trace(go.Bar(x=plot_df['物料编码'], y=plot_df['安全库存量'], name='安全库存', marker_color='#1f77b4'))
            if '实际库存' in plot_df.columns:
                fig.add_trace(go.Bar(x=plot_df['物料编码'], y=plot_df['实际库存'], name='实际库存', marker_color='#ff7f0e'))
            if '预测库存' in plot_df.columns:
                fig.add_trace(go.Bar(x=plot_df['物料编码'], y=plot_df['预测库存'], name='预测库存', marker_color='#2ca02c'))
            fig.update_layout(title="安全库存 vs 实际库存 vs 预测库存", xaxis_title="物料编码", yaxis_title="库存量", barmode='group', height=400)
            st.plotly_chart(fig, use_container_width=True)
            
            col1, col2 = st.columns(2)
            with col1:
                if '库存状态' in df_display.columns:
                    status_counts = df_display['库存状态'].value_counts()
                    fig_pie = px.pie(values=status_counts.values, names=status_counts.index, title="库存状态分布")
                    st.plotly_chart(fig_pie, use_container_width=True)
            with col2:
                if '日均消耗' in df_display.columns:
                    top_usage = df_display.nlargest(10, '日均消耗')
                    fig_bar = px.bar(top_usage, x='物料编码', y='日均消耗', title="日均消耗TOP10")
                    st.plotly_chart(fig_bar, use_container_width=True)
        
        with tab3:
            if '库存状态' in df_display.columns:
                warning_df = df_display[df_display['库存状态'] == '⚠️ 预警'].copy()
                if not warning_df.empty:
                    warning_df = warning_df.sort_values('库存差异(实际-安全)')
                    st.markdown(f"**共 {len(warning_df)} 个物料低于安全库存**")
                    display_cols = ['物料编码', '物料描述', '工厂', '存储地点', '安全库存量', '实际库存', '库存差异(实际-安全)']
                    display_cols = [col for col in display_cols if col in warning_df.columns]
                    st.dataframe(warning_df[display_cols], use_container_width=True, height=400)
                else:
                    st.success("🎉 所有物料库存充足，没有预警！")
        
        with tab4:
            view_type = st.radio("查看方式", ['所有物料总趋势', '单个物料趋势'], horizontal=True)
            st.info("💡 趋势图需要Excel中包含M01~M09的历史数据")

if __name__ == "__main__":
    main()
