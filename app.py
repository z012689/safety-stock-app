import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO

# 页面配置
st.set_page_config(
    page_title="安全库存管理系统",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ==================== 高饱和配色方案 ====================
COLORS = {
    "primary": "#FF4B4B",
    "secondary": "#FF6B35",
    "success": "#00C853",
    "warning": "#FFC107",
    "danger": "#D32F2F",
    "info": "#2196F3",
    "purple": "#9C27B0",
    "cyan": "#00BCD4",
    "dark": "#1E1E2E",
    "card_bg": "#2D2D3D",
    "bg": "#1A1A2E"
}

# 自定义CSS
st.markdown(f"""
<style>
    /* 主背景 */
    .stApp {{
        background: linear-gradient(135deg, #1A1A2E 0%, #16213E 100%);
    }}
    
    /* 卡片样式 */
    .metric-card {{
        background: linear-gradient(135deg, #2D2D3D 0%, #1E1E2E 100%);
        border-radius: 16px;
        padding: 20px;
        box-shadow: 0 8px 20px rgba(0,0,0,0.3);
        border: 1px solid rgba(255,255,255,0.1);
        transition: transform 0.3s;
    }}
    .metric-card:hover {{
        transform: translateY(-5px);
    }}
    
    /* 数值大字体 */
    .metric-value {{
        font-size: 32px;
        font-weight: 700;
        color: #FF6B35;
    }}
    .metric-label {{
        font-size: 14px;
        color: #A0A0B0;
        letter-spacing: 1px;
    }}
    
    /* 警告卡片 */
    .warning-card {{
        background: linear-gradient(135deg, #FF4B4B20 0%, #D32F2F20 100%);
        border-left: 4px solid #FF4B4B;
        border-radius: 12px;
        padding: 16px;
        margin: 10px 0;
    }}
    
    /* 成功卡片 */
    .success-card {{
        background: linear-gradient(135deg, #00C85320 0%, #00BCD420 100%);
        border-left: 4px solid #00C853;
        border-radius: 12px;
        padding: 16px;
        margin: 10px 0;
    }}
    
    /* 按钮样式 */
    .stButton > button {{
        background: linear-gradient(135deg, #FF4B4B, #FF6B35);
        color: white;
        border-radius: 25px;
        border: none;
        padding: 10px 24px;
        font-weight: 600;
        transition: all 0.3s;
    }}
    .stButton > button:hover {{
        transform: scale(1.02);
        box-shadow: 0 5px 15px rgba(255,75,75,0.4);
    }}
    
    /* 上传区域 */
    .upload-area {{
        border: 2px dashed #FF6B35;
        border-radius: 20px;
        background: rgba(255,107,53,0.1);
        padding: 40px;
        text-align: center;
    }}
    
    /* 数据表格 */
    .stDataFrame {{
        border-radius: 16px;
        overflow: hidden;
    }}
    
    /* 标题 */
    .main-title {{
        font-size: 42px;
        font-weight: 800;
        background: linear-gradient(135deg, #FF4B4B, #FF6B35, #FFC107);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        margin-bottom: 10px;
    }}
    .sub-title {{
        color: #A0A0B0;
        font-size: 16px;
        margin-bottom: 30px;
    }}
    
    /* 侧边栏 */
    .css-1d391kg {{
        background: #1E1E2E;
    }}
    
    /* Tab样式 */
    .stTabs [data-baseweb="tab-list"] {{
        gap: 8px;
    }}
    .stTabs [data-baseweb="tab"] {{
        background: #2D2D3D;
        border-radius: 12px 12px 0 0;
        padding: 10px 24px;
        color: #A0A0B0;
    }}
    .stTabs [aria-selected="true"] {{
        background: linear-gradient(135deg, #FF4B4B, #FF6B35);
        color: white;
    }}
    
    /* 进度条 */
    .progress-bar {{
        background: #2D2D3D;
        border-radius: 10px;
        height: 8px;
        overflow: hidden;
    }}
    .progress-fill {{
        background: linear-gradient(90deg, #FF4B4B, #FF6B35);
        height: 100%;
        border-radius: 10px;
        transition: width 0.5s;
    }}
</style>
""", unsafe_allow_html=True)


class SafetyStockCalculator:
    """安全库存计算器 - 严格按照Excel公式"""
    
    @staticmethod
    def calculate_lead_time_coefficient(days):
        """采购周期系数：严格按照规则"""
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
    def calculate_quality_risk_score(df_quality, material_code):
        """质量风险系数 = 1 + 低风险批次数*0.1 + 中风险批次数*0.2 + 高风险批次数*0.3"""
        if df_quality is None:
            return 1.0
        
        matched = df_quality[df_quality['物料编码'].astype(str) == str(material_code)]
        if matched.empty:
            return 1.0
        
        low = matched.iloc[0].get('低风险批次', 0) if '低风险批次' in matched.columns else 0
        mid = matched.iloc[0].get('中风险批次', 0) if '中风险批次' in matched.columns else 0
        high = matched.iloc[0].get('高风险批次', 0) if '高风险批次' in matched.columns else 0
        
        if pd.isna(low):
            low = 0
        if pd.isna(mid):
            mid = 0
        if pd.isna(high):
            high = 0
        
        return 1.0 + float(low) * 0.1 + float(mid) * 0.2 + float(high) * 0.3
    
    @staticmethod
    def get_category_risk(df_category, material_code):
        """获取品类策略系数"""
        if df_category is None:
            return 1.0
        
        code_col = None
        risk_col = None
        for col in df_category.columns:
            if '物料' in col:
                code_col = col
            if '风险' in col:
                risk_col = col
        
        if code_col and risk_col:
            matched = df_category[df_category[code_col].astype(str) == str(material_code)]
            if not matched.empty:
                val = matched.iloc[0][risk_col]
                if pd.notna(val):
                    return float(val)
        return 1.0
    
    @staticmethod
    def calculate_safety_stock(future_avg, past_avg, quality_score, category_risk, lead_coef):
        """安全库存 = (未来3个月月均 + 过去6个月月均) / 2 × (质量风险×40% + 品类策略×60%) × 采购周期系数"""
        if pd.isna(future_avg) or pd.isna(past_avg):
            return 0
        if future_avg <= 0 and past_avg <= 0:
            return 0
        
        base_avg = (float(future_avg) + float(past_avg)) / 2
        combined_coef = quality_score * 0.4 + category_risk * 0.6
        
        return base_avg * combined_coef * lead_coef
    
    def process_data(self, df_materials, df_quality, df_category):
        """完整数据处理"""
        df = df_materials.copy()
        
        # 识别月份列
        future_cols = []
        past_cols = []
        
        for col in df.columns:
            col_str = str(col)
            if any(x in col_str for x in ['M07', 'M08', 'M09', '7月', '8月', '9月']):
                future_cols.append(col)
            elif any(x in col_str for x in ['M01', 'M02', 'M03', 'M04', 'M05', 'M06', '1月', '2月', '3月', '4月', '5月', '6月']):
                past_cols.append(col)
        
        # 计算月均用量
        if future_cols:
            future_data = df[future_cols].apply(pd.to_numeric, errors='coerce')
            df['未来3个月月均用量'] = future_data.mean(axis=1)
        elif '未来3个月月均用量' in df.columns:
            df['未来3个月月均用量'] = pd.to_numeric(df['未来3个月月均用量'], errors='coerce')
        
        if past_cols:
            past_data = df[past_cols].apply(pd.to_numeric, errors='coerce')
            df['过去6个月月均用量'] = past_data.mean(axis=1)
        elif '月均量(半年)' in df.columns:
            df['过去6个月月均用量'] = pd.to_numeric(df['月均量(半年)'], errors='coerce')
        
        # 计算采购周期系数
        if '平均交货周期(天)' in df.columns:
            df['采购周期系数'] = df['平均交货周期(天)'].apply(self.calculate_lead_time_coefficient)
        
        # 计算各项系数并最终计算安全库存
        quality_scores = []
        category_risks = []
        
        for idx, row in df.iterrows():
            code = row.get('物料编码', '')
            quality = self.calculate_quality_risk_score(df_quality, code)
            category = self.get_category_risk(df_category, code)
            quality_scores.append(quality)
            category_risks.append(category)
        
        df['质量风险系数'] = quality_scores
        df['品类策略系数'] = category_risks
        
        # 计算安全库存
        df['安全库存'] = df.apply(
            lambda x: self.calculate_safety_stock(
                x.get('未来3个月月均用量', 0),
                x.get('过去6个月月均用量', 0),
                x.get('质量风险系数', 1),
                x.get('品类策略系数', 1),
                x.get('采购周期系数', 1)
            ), axis=1
        )
        
        # 寄售物料安全库存为0
        if '是否寄售' in df.columns:
            df.loc[df['是否寄售'] == '寄售', '安全库存'] = 0
        
        # 实际库存（6月末）
        actual_stock_col = None
        for col in df.columns:
            if '实际库存' in col or '6月末' in col:
                actual_stock_col = col
                break
        if actual_stock_col:
            df['实际库存'] = pd.to_numeric(df[actual_stock_col], errors='coerce').fillna(0)
        else:
            df['实际库存'] = 0
        
        # 计算库存覆盖倍数
        df['库存覆盖倍数'] = df.apply(
            lambda x: x['实际库存'] / x['安全库存'] if x['安全库存'] > 0 else 0,
            axis=1
        )
        
        # 低库存预警（覆盖倍数 < 1.5）
        df['预警等级'] = df['库存覆盖倍数'].apply(
            lambda x: '🔴 严重不足' if x < 0.5 else ('🟡 偏低' if x < 1.5 else ('🟢 充足' if x < 3 else '✅ 过量'))
        )
        
        return df


def show_kpi_cards(df):
    """显示KPI卡片"""
    total_materials = len(df)
    low_stock = len(df[df['库存覆盖倍数'] < 1.5])
    low_stock_pct = round(low_stock / total_materials * 100, 1) if total_materials > 0 else 0
    avg_safety = df['安全库存'].mean() if '安全库存' in df.columns else 0
    total_actual = df['实际库存'].sum() if '实际库存' in df.columns else 0
    avg_coverage = df['库存覆盖倍数'].mean() if '库存覆盖倍数' in df.columns else 0
    
    coverage_status = "✅ 充足" if avg_coverage >= 2 else ("⚠️ 一般" if avg_coverage >= 1 else "🔴 紧张")
    
    col1, col2, col3, col4, col5 = st.columns(5)
    
    cards = [
        (col1, "📦 物料总览", f"{total_materials:,}", "个物料"),
        (col2, "⚠️ 低库存预警物料", f"{low_stock:,}", f"占比 {low_stock_pct}%"),
        (col3, "📊 平均安全库存", f"{avg_safety:,.0f}", "件/物料"),
        (col4, "💰 总实际库存", f"{total_actual:,.0f}", "件"),
        (col5, "📈 平均库存覆盖", f"{avg_coverage:.1f}倍", coverage_status)
    ]
    
    for col, label, value, unit in cards:
        with col:
            st.markdown(f"""
            <div class="metric-card">
                <div class="metric-label">{label}</div>
                <div class="metric-value">{value}</div>
                <div class="metric-label" style="font-size:12px">{unit}</div>
            </div>
            """, unsafe_allow_html=True)


def show_top_warnings(df):
    """显示高风险物料预警"""
    warnings_df = df[df['库存覆盖倍数'] < 1.5].nlargest(5, '安全库存')[['物料编码', 'SAP编码', '实际库存', '安全库存', '库存覆盖倍数', '预警等级']]
    
    if not warnings_df.empty:
        st.markdown("""
        <div class="warning-card">
            <strong>⚠️ 高风险物料预警</strong><br>
            以下物料库存覆盖不足，建议优先采购
        </div>
        """, unsafe_allow_html=True)
        
        st.dataframe(
            warnings_df.style.format({
                '实际库存': '{:,.0f}',
                '安全库存': '{:,.0f}',
                '库存覆盖倍数': '{:.1f}'
            }).applymap(
                lambda x: 'color: #FF4B4B; font-weight: bold' if isinstance(x, str) and '严重' in x else '',
                subset=['预警等级']
            ),
            use_container_width=True,
            height=250
        )


def show_coverage_distribution(df):
    """库存覆盖分布图"""
    if '库存覆盖倍数' not in df.columns:
        return
    
    coverage_bins = ['<0.5', '0.5-1.0', '1.0-1.5', '1.5-2.0', '2.0-3.0', '>3.0']
    coverage_counts = [
        len(df[df['库存覆盖倍数'] < 0.5]),
        len(df[(df['库存覆盖倍数'] >= 0.5) & (df['库存覆盖倍数'] < 1.0)]),
        len(df[(df['库存覆盖倍数'] >= 1.0) & (df['库存覆盖倍数'] < 1.5)]),
        len(df[(df['库存覆盖倍数'] >= 1.5) & (df['库存覆盖倍数'] < 2.0)]),
        len(df[(df['库存覆盖倍数'] >= 2.0) & (df['库存覆盖倍数'] < 3.0)]),
        len(df[df['库存覆盖倍数'] >= 3.0])
    ]
    
    colors_bins = ['#D32F2F', '#FF6B35', '#FFC107', '#00C853', '#2196F3', '#9C27B0']
    
    fig = go.Figure(data=[
        go.Bar(
            x=coverage_bins,
            y=coverage_counts,
            marker_color=colors_bins,
            text=coverage_counts,
            textposition='auto',
            hovertemplate='库存覆盖: %{x}<br>物料数量: %{y}<extra></extra>'
        )
    ])
    
    fig.update_layout(
        title="📊 库存覆盖分布",
        xaxis_title="库存覆盖倍数",
        yaxis_title="物料数量",
        plot_bgcolor='rgba(45,45,61,0.5)',
        paper_bgcolor='rgba(0,0,0,0)',
        font_color='#A0A0B0',
        height=400,
        margin=dict(l=40, r=40, t=50, b=40)
    )
    
    fig.update_traces(marker_line_width=0)
    
    st.plotly_chart(fig, use_container_width=True)


def main():
    # 标题
    st.markdown('<div class="main-title">📦 安全库存管理系统</div>', unsafe_allow_html=True)
    st.markdown('<div class="sub-title">自动计算 · 智能预警 · 月度更新</div>', unsafe_allow_html=True)
    
    # 初始化session state
    if 'df_result' not in st.session_state:
        st.session_state.df_result = None
    if 'original_data' not in st.session_state:
        st.session_state.original_data = None
    if 'df_quality' not in st.session_state:
        st.session_state.df_quality = None
    if 'df_category' not in st.session_state:
        st.session_state.df_category = None
    
    # 侧边栏
    with st.sidebar:
        st.markdown("### ⚙️ 系统配置")
        st.markdown("---")
        
        # 文件上传
        uploaded_file = st.file_uploader(
            "📁 上传Excel文件",
            type=['xlsx', 'xls'],
            help="支持 .xlsx, .xls 格式"
        )
        
        st.markdown("---")
        
        # 月度更新区域
        st.markdown("### 📅 月度数据更新")
        
        current_date = datetime.now()
        new_month_name = st.text_input(
            "新月份名称",
            value=f"{current_date.year}年{current_date.month}月"
        )
        
        update_file = st.file_uploader(
            "上传新月份数据文件",
            type=['xlsx', 'xls'],
            key="monthly_update",
            help="文件需包含「物料编码」和「用量」列"
        )
        
        st.markdown("---")
        
        # 按钮
        calculate_btn = st.button("🚀 开始计算", type="primary", use_container_width=True)
        update_btn = st.button("🔄 月度更新", use_container_width=True)
        
        st.markdown("---")
        st.markdown("### 📖 计算公式")
        st.caption("安全库存 = (未来3个月月均 + 过去6个月月均) / 2 × (质量风险×40% + 品类策略×60%) × 采购周期系数")
    
    # 主内容区
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
                        df_result = calculator.process_data(df_materials, df_quality, df_category)
                        
                        st.session_state.df_result = df_result
                        st.session_state.original_data = df_materials
                        st.session_state.df_quality = df_quality
                        st.session_state.df_category = df_category
                    
                    st.success("✅ 安全库存计算完成！")
                    
                    # 显示KPI卡片
                    show_kpi_cards(df_result)
                    
                    st.markdown("---")
                    
                    # 高风险预警
                    show_top_warnings(df_result)
                    
                    st.markdown("---")
                    
                    # 图表和表格Tab
                    tab1, tab2, tab3 = st.tabs(["📊 数据表格", "📈 统计分析", "💾 数据导出"])
                    
                    with tab1:
                        display_cols = ['物料编码', 'SAP编码', '未来3个月月均用量', '过去6个月月均用量',
                                       '质量风险系数', '品类策略系数', '采购周期系数', '安全库存', '实际库存', 
                                       '库存覆盖倍数', '预警等级', '是否寄售', '备注']
                        display_cols = [c for c in display_cols if c in df_result.columns]
                        
                        st.dataframe(
                            df_result[display_cols].head(100).style.format({
                                '未来3个月月均用量': '{:,.0f}',
                                '过去6个月月均用量': '{:,.0f}',
                                '安全库存': '{:,.0f}',
                                '实际库存': '{:,.0f}',
                                '库存覆盖倍数': '{:.1f}'
                            }).applymap(
                                lambda x: 'background-color: rgba(255,75,75,0.2)' if isinstance(x, str) and '严重' in x else '',
                                subset=['预警等级']
                            ),
                            use_container_width=True,
                            height=400
                        )
                    
                    with tab2:
                        col1, col2 = st.columns(2)
                        
                        with col1:
                            show_coverage_distribution(df_result)
                        
                        with col2:
                            if '质量风险系数' in df_result.columns:
                                fig_scatter = px.scatter(
                                    df_result.head(200),
                                    x='安全库存',
                                    y='库存覆盖倍数',
                                    color='预警等级',
                                    color_discrete_map={
                                        '🔴 严重不足': '#D32F2F',
                                        '🟡 偏低': '#FFC107',
                                        '🟢 充足': '#00C853',
                                        '✅ 过量': '#2196F3'
                                    },
                                    title="📈 安全库存 vs 库存覆盖",
                                    labels={'安全库存': '安全库存', '库存覆盖倍数': '库存覆盖倍数'}
                                )
                                fig_scatter.update_layout(
                                    plot_bgcolor='rgba(45,45,61,0.5)',
                                    paper_bgcolor='rgba(0,0,0,0)',
                                    font_color='#A0A0B0',
                                    height=400
                                )
                                st.plotly_chart(fig_scatter, use_container_width=True)
                    
                    with tab3:
                        st.subheader("📥 数据导出")
                        
                        output = BytesIO()
                        with pd.ExcelWriter(output, engine='openpyxl') as writer:
                            df_result.to_excel(writer, sheet_name='安全库存计算结果', index=False)
                            
                            summary_data = {
                                '指标': ['总物料数', '总安全库存', '平均安全库存', '最大安全库存', '总实际库存', '低库存物料数', '低库存占比', '平均库存覆盖'],
                                '数值': [
                                    len(df_result),
                                    df_result['安全库存'].sum(),
                                    df_result['安全库存'].mean(),
                                    df_result['安全库存'].max(),
                                    df_result['实际库存'].sum(),
                                    len(df_result[df_result['库存覆盖倍数'] < 1.5]),
                                    f"{len(df_result[df_result['库存覆盖倍数'] < 1.5]) / len(df_result) * 100:.1f}%",
                                    f"{df_result['库存覆盖倍数'].mean():.1f}倍"
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
                
                elif st.session_state.df_result is not None:
                    show_kpi_cards(st.session_state.df_result)
                    st.markdown("---")
                    show_top_warnings(st.session_state.df_result)
                    
                    tab1, tab2, tab3 = st.tabs(["📊 数据表格", "📈 统计分析", "💾 数据导出"])
                    with tab1:
                        display_cols = ['物料编码', 'SAP编码', '安全库存', '实际库存', '库存覆盖倍数', '预警等级']
                        display_cols = [c for c in display_cols if c in st.session_state.df_result.columns]
                        st.dataframe(st.session_state.df_result[display_cols].head(100), use_container_width=True)
                
                else:
                    st.info("👈 请点击左侧「开始计算」按钮")
            else:
                st.error("❌ 未找到「安全库存（202509月）」工作表")
                
        except Exception as e:
            st.error(f"❌ 处理文件时出错: {str(e)}")
    else:
        # 未上传文件时的占位
        st.markdown("""
        <div class="upload-area">
            <h3>📂 拖拽Excel文件到此处，或点击上传</h3>
            <p>支持 .xlsx, .xls 格式 | 需包含「安全库存（202509月）」工作表</p>
        </div>
        """, unsafe_allow_html=True)
    
    # 页脚
    st.markdown("---")
    st.markdown(
        "<div style='text-align: center; color: #A0A0B0; padding: 20px;'>安全库存管理系统 | 版本 2.0 | 支持月度数据更新</div>",
        unsafe_allow_html=True
    )


if __name__ == "__main__":
    main()
