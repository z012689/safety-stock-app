# app.py
import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import io
import re
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
    .main-header { font-size: 2.2rem; font-weight: 700; color: #1f77b4; text-align: center; }
    .file-status { padding: 0.25rem 0.5rem; border-radius: 0.25rem; font-size: 0.9rem; }
    .file-uploaded { background-color: #d4edda; color: #155724; }
    .file-missing { background-color: #f8d7da; color: #721c24; }
    .file-optional { background-color: #fff3cd; color: #856404; }
    .monthly-update-badge { background-color: #28a745; color: white; padding: 0.2rem 0.8rem; border-radius: 1rem; font-size: 0.8rem; display: inline-block; }
    </style>
""", unsafe_allow_html=True)

def extract_number_from_text(text):
    """从文本中提取数字"""
    if pd.isna(text):
        return 300
    try:
        # 如果是数字，直接返回
        if isinstance(text, (int, float)):
            return float(text)
        # 如果是字符串，提取数字
        text = str(text)
        # 查找所有数字（包括小数）
        numbers = re.findall(r'(\d+\.?\d*)', text)
        if numbers:
            return float(numbers[0])
        return 300
    except:
        return 300

# ============================================================================
# 数据加载函数
# ============================================================================

@st.cache_data
def load_safety_stock(file):
    """
    加载安全库存模板
    实际列名: '原料安全库存', '（未来3个月月均用量+过去6个月月均使用量）/2*(质量风险系数权重40%+品类策略系数权重60%) *采购周期系数'
    """
    try:
        df = pd.read_excel(file)
        
        # 显示所有列名
        col_list = df.columns.tolist()
        st.info(f"📄 安全库存列名: {col_list}")
        
        # ============================================================
        # 处理物料编码
        # ============================================================
        # 查找物料编码列
        material_col = None
        for col in col_list:
            col_clean = str(col).replace(' ', '').replace('（', '(').replace('）', ')')
            if '物料' in col_clean or '编码' in col_clean or 'SAP' in col_clean:
                material_col = col
                break
        
        if material_col is None:
            material_col = col_list[0]
            st.warning(f"⚠️ 未找到物料编码列，使用第一列: {material_col}")
        
        df.rename(columns={material_col: '物料编码'}, inplace=True)
        df['物料编码'] = df['物料编码'].astype(str)
        
        # ============================================================
        # 处理用量数据 - 提取数字
        # ============================================================
        # 查找用量列
        usage_col = None
        for col in df.columns:
            if col != '物料编码' and ('月均' in str(col) or '用量' in str(col) or '公式' in str(col) or '计算' in str(col)):
                usage_col = col
                break
        
        if usage_col is None:
            # 使用第二列
            for col in df.columns:
                if col != '物料编码':
                    usage_col = col
                    break
        
        if usage_col:
            st.info(f"📊 使用列 '{usage_col}' 作为用量数据")
            # 提取数字
            df['未来3个月月均用量'] = df[usage_col].apply(extract_number_from_text)
            # 删除原始列（可选）
            # df.drop(columns=[usage_col], inplace=True)
        else:
            st.warning("⚠️ 未找到用量列，使用默认值300")
            df['未来3个月月均用量'] = 300
        
        # ============================================================
        # 计算其他字段
        # ============================================================
        # 过去6个月月均用量（假设为用量的85%）
        df['过去6个月月均用量'] = df['未来3个月月均用量'] * 0.85
        
        # 默认值
        df['平均交货周期'] = '2周'
        df['采购周期系数'] = 1.5
        df['质量风险系数'] = 1.0
        df['品类策略系数'] = 1.0
        
        # 计算安全库存
        df['安全库存量'] = (df['未来3个月月均用量'] * df['采购周期系数'] * df['质量风险系数'] * df['品类策略系数']).round(2)
        
        # 添加默认列
        df['物料描述'] = df['物料编码']
        df['工厂'] = '默认工厂'
        df['存储地点'] = '默认仓库'
        df['单位'] = 'KG'
        df['实际库存'] = df['安全库存量'] * 0.8
        df['预测库存'] = df['安全库存量'] * 0.9
        
        st.success(f"✅ 成功加载 {len(df)} 条物料数据")
        st.info(f"📊 用量范围: {df['未来3个月月均用量'].min():.2f} ~ {df['未来3个月月均用量'].max():.2f}")
        return df
        
    except Exception as e:
        st.error(f"❌ 加载安全库存文件失败: {e}")
        import traceback
        st.error(traceback.format_exc())
        return None

@st.cache_data
def load_inventory(file):
    """加载6.1库存文件"""
    try:
        df = pd.read_excel(file, sheet_name='Data')
        st.info(f"📄 库存列名: {df.columns.tolist()}")
        
        if '存储地点' in df.columns:
            df_qualified = df[df['存储地点'] == 3006].copy()
            st.info(f"📊 合格仓数据: {len(df_qualified)} 条")
        else:
            df_qualified = df.copy()
        
        if '物料' in df_qualified.columns and '非限制使用的库存' in df_qualified.columns:
            result = df_qualified.groupby('物料')['非限制使用的库存'].sum().reset_index()
            result.columns = ['物料编码', '实际库存']
        else:
            st.warning("⚠️ 未找到'物料'或'非限制使用的库存'列")
            return None
        
        result['物料编码'] = result['物料编码'].astype(str)
        result['实际库存'] = result['实际库存'].round(2)
        
        st.success(f"✅ 成功加载 {len(result)} 条库存数据")
        return result
    except Exception as e:
        st.warning(f"⚠️ 加载库存文件失败: {e}")
        return None

@st.cache_data
def load_usage(file):
    """加载5月使用量文件"""
    try:
        df = pd.read_excel(file)
        st.info(f"📄 使用量列名: {df.columns.tolist()}")
        
        if '过账日期' not in df.columns or '以录入单位表示的数量' not in df.columns:
            st.warning("⚠️ 未找到'过账日期'或'以录入单位表示的数量'列")
            return None
        
        # 使用物料作为物料编码
        if '物料' in df.columns:
            result = df.groupby('物料')['以录入单位表示的数量'].sum().reset_index()
            result.columns = ['物料编码', '总消耗量']
        elif '物料凭证' in df.columns:
            result = df.groupby('物料凭证')['以录入单位表示的数量'].sum().reset_index()
            result.columns = ['物料编码', '总消耗量']
        else:
            result = df.groupby(df.index)['以录入单位表示的数量'].sum().reset_index()
            result.columns = ['物料编码', '总消耗量']
            result['物料编码'] = result['物料编码'].astype(str)
        
        result['日均消耗'] = (result['总消耗量'] / 31).round(2)
        result['物料编码'] = result['物料编码'].astype(str)
        
        st.success(f"✅ 成功加载 {len(result)} 条使用量数据")
        return result[['物料编码', '日均消耗']]
    except Exception as e:
        st.warning(f"⚠️ 加载使用量文件失败: {e}")
        return None

@st.cache_data
def load_forecast(file):
    """加载预测库存文件"""
    try:
        df = pd.read_excel(file)
        st.info(f"📄 预测列名: {df.columns.tolist()}")
        
        if '物料代码' not in df.columns or '物料理论数量（该订单）' not in df.columns:
            st.warning("⚠️ 未找到'物料代码'或'物料理论数量（该订单）'列")
            return None
        
        result = df.groupby('物料代码')['物料理论数量（该订单）'].sum().reset_index()
        result.columns = ['物料编码', '预测库存']
        result['物料编码'] = result['物料编码'].astype(str)
        result['预测库存'] = result['预测库存'].round(2)
        
        st.success(f"✅ 成功加载 {len(result)} 条预测数据")
        return result
    except Exception as e:
        st.warning(f"⚠️ 加载预测文件失败: {e}")
        return None

def load_mock_data():
    """生成模拟数据"""
    np.random.seed(42)
    data = []
    for i in range(1, 51):
        mat = f'MAT-{i:04d}'
        data.append({
            '物料编码': mat,
            '物料描述': f'物料_{mat}',
            '工厂': np.random.choice(['广州工厂', '上海工厂', '北京工厂']),
            '存储地点': np.random.choice(['A区', 'B区', 'C区']),
            '单位': 'KG',
            '未来3个月月均用量': round(np.random.uniform(200, 600), 2),
            '过去6个月月均用量': round(np.random.uniform(150, 500), 2),
            '平均交货周期': np.random.choice(['1周', '2周', '3周', '4周']),
            '安全库存量': round(np.random.uniform(200, 800), 2),
            '实际库存': round(np.random.uniform(100, 900), 2),
            '预测库存': round(np.random.uniform(150, 850), 2),
        })
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
    
    # 初始化session state
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
            "② 6.1库存-去除名称.xlsx",
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
            "③ 5月使用量.xlsx",
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
            "④ 预测-6.1.xlsx",
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
    
    # ========================================================================
    # 加载数据
    # ========================================================================
    
    if safety_stock_file is not None:
        df = load_safety_stock(safety_stock_file)
        if df is not None:
            st.session_state.safety_data = df
            st.session_state.template_columns = df.columns.tolist()
            st.session_state.last_update = datetime.now().strftime('%Y-%m-%d %H:%M')
            st.sidebar.success(f"✅ 已加载 {len(df)} 条物料数据")
        else:
            st.session_state.safety_data = load_mock_data()
            st.session_state.template_columns = st.session_state.safety_data.columns.tolist()
    else:
        if st.session_state.safety_data is None:
            st.session_state.safety_data = load_mock_data()
            st.session_state.template_columns = st.session_state.safety_data.columns.tolist()
            st.sidebar.info("ℹ️ 使用模拟数据")
    
    if inventory_file is not None:
        st.session_state.inventory_data = load_inventory(inventory_file)
    if usage_file is not None:
        st.session_state.usage_data = load_usage(usage_file)
    if forecast_file is not None:
        st.session_state.forecast_data = load_forecast(forecast_file)
    
    # ========================================================================
    # 显示结果
    # ========================================================================
    
    if st.session_state.safety_data is not None:
        df_display = st.session_state.safety_data.copy()
        df_display['物料编码'] = df_display['物料编码'].astype(str)
        
        # 确保数值列是数字类型
        numeric_cols = ['未来3个月月均用量', '过去6个月月均用量', '安全库存量', '实际库存', '预测库存']
        for col in numeric_cols:
            if col in df_display.columns:
                df_display[col] = pd.to_numeric(df_display[col], errors='coerce').fillna(0)
        
        # 确保必要的列存在
        if '安全库存量' not in df_display.columns or df_display['安全库存量'].sum() == 0:
            df_display['安全库存量'] = df_display['未来3个月月均用量'] * 1.5
            df_display['安全库存量'] = df_display['安全库存量'].round(2)
        
        if '实际库存' not in df_display.columns:
            df_display['实际库存'] = df_display['安全库存量'] * np.random.uniform(0.5, 1.8, len(df_display))
            df_display['实际库存'] = df_display['实际库存'].round(2)
        
        if '预测库存' not in df_display.columns:
            df_display['预测库存'] = df_display['安全库存量'] * np.random.uniform(0.6, 1.5, len(df_display))
            df_display['预测库存'] = df_display['预测库存'].round(2)
        
        # 合并库存数据
        if st.session_state.inventory_data is not None:
            st.session_state.inventory_data['物料编码'] = st.session_state.inventory_data['物料编码'].astype(str)
            df_display = df_display.merge(st.session_state.inventory_data, on='物料编码', how='left', suffixes=('', '_inv'))
            if '实际库存_inv' in df_display.columns:
                df_display['实际库存'] = df_display['实际库存_inv'].fillna(df_display['实际库存'])
                df_display.drop(columns=['实际库存_inv'], inplace=True)
            df_display['实际库存'] = pd.to_numeric(df_display['实际库存'], errors='coerce').fillna(0).round(2)
        
        # 合并预测数据
        if st.session_state.forecast_data is not None:
            st.session_state.forecast_data['物料编码'] = st.session_state.forecast_data['物料编码'].astype(str)
            df_display = df_display.merge(st.session_state.forecast_data, on='物料编码', how='left', suffixes=('', '_fore'))
            if '预测库存_fore' in df_display.columns:
                df_display['预测库存'] = df_display['预测库存_fore'].fillna(df_display['预测库存'])
                df_display.drop(columns=['预测库存_fore'], inplace=True)
            df_display['预测库存'] = pd.to_numeric(df_display['预测库存'], errors='coerce').fillna(0).round(2)
        
        # 计算库存差异
        df_display['库存差异(实际-安全)'] = df_display['实际库存'] - df_display['安全库存量']
        df_display['库存状态'] = df_display['库存差异(实际-安全)'].apply(
            lambda x: '✅ 充足' if x >= 0 else '⚠️ 预警'
        )
        
        # 合并使用量数据
        if st.session_state.usage_data is not None:
            st.session_state.usage_data['物料编码'] = st.session_state.usage_data['物料编码'].astype(str)
            df_display = df_display.merge(st.session_state.usage_data, on='物料编码', how='left')
            if '日均消耗' in df_display.columns:
                df_display['日均消耗'] = pd.to_numeric(df_display['日均消耗'], errors='coerce').fillna(
                    df_display['过去6个月月均用量'] / 30 if '过去6个月月均用量' in df_display.columns else 10
                )
            else:
                df_display['日均消耗'] = df_display['过去6个月月均用量'] / 30 if '过去6个月月均用量' in df_display.columns else 10
        else:
            df_display['日均消耗'] = df_display['过去6个月月均用量'] / 30 if '过去6个月月均用量' in df_display.columns else 10
        
        # ====================================================================
        # Tab布局
        # ====================================================================
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
                st.metric("📊 总安全库存", f"{df_display['安全库存量'].sum():,.0f}")
            with col3:
                if '未来3个月月均用量' in df_display.columns:
                    st.metric("📈 平均月用量", f"{df_display['未来3个月月均用量'].mean():,.0f}")
                else:
                    st.metric("📈 平均月用量", "N/A")
            with col4:
                warning_count = len(df_display[df_display['库存状态'] == '⚠️ 预警'])
                st.metric("⚠️ 预警物料", warning_count, delta_color="inverse")
            
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
            else:
                st.info("ℹ️ 暂无库存状态数据")
        
        with tab4:
            view_type = st.radio("查看方式", ['所有物料总趋势', '单个物料趋势'], horizontal=True)
            st.info("💡 趋势图需要Excel中包含M01~M09的历史数据，当前显示模拟趋势")
            mock_months = [f'M{i:02d}' for i in range(1, 10)]
            mock_values = [round(np.random.uniform(100, 500), 2) for _ in range(9)]
            fig = px.line(x=mock_months, y=mock_values, title="模拟消耗趋势", markers=True)
            fig.update_layout(height=400)
            st.plotly_chart(fig, use_container_width=True)

if __name__ == "__main__":
    main()
