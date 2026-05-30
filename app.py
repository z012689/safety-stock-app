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
    page_icon="📦",
    layout="wide"
)

# 自定义CSS - 美观的现代风格
st.markdown("""
<style>
    /* 主背景 */
    .stApp {
        background: linear-gradient(135deg, #F0F4F8 0%, #E2E8F0 100%);
    }
    
    /* 主标题 */
    .main-title {
        font-size: 32px;
        font-weight: 700;
        background: linear-gradient(135deg, #1E293B, #3B82F6);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        margin-bottom: 5px;
    }
    .sub-title {
        font-size: 14px;
        color: #64748B;
        margin-bottom: 20px;
    }
    
    /* 统计卡片 */
    .stat-card {
        background: white;
        border-radius: 20px;
        padding: 20px;
        box-shadow: 0 4px 12px rgba(0,0,0,0.05);
        transition: all 0.3s;
        border: 1px solid #E2E8F0;
    }
    .stat-card:hover {
        transform: translateY(-3px);
        box-shadow: 0 8px 24px rgba(0,0,0,0.1);
    }
    .stat-value {
        font-size: 28px;
        font-weight: 700;
        color: #1E293B;
    }
    .stat-label {
        font-size: 13px;
        color: #64748B;
        margin-top: 8px;
    }
    .stat-unit {
        font-size: 11px;
        color: #94A3B8;
    }
    
    /* 预警卡片 */
    .alert-card {
        background: linear-gradient(135deg, #FEF2F2, #FEE2E2);
        border-left: 4px solid #EF4444;
        border-radius: 16px;
        padding: 16px 20px;
        margin: 20px 0;
    }
    .alert-title {
        font-weight: 700;
        color: #DC2626;
        font-size: 16px;
        margin-bottom: 8px;
    }
    
    /* 上传区域 */
    .upload-box {
        background: white;
        border: 2px dashed #CBD5E1;
        border-radius: 24px;
        padding: 40px;
        text-align: center;
        margin: 20px 0;
    }
    
    /* 顶部信息栏 */
    .top-bar {
        background: white;
        border-radius: 16px;
        padding: 12px 20px;
        margin-bottom: 24px;
        display: flex;
        justify-content: space-between;
        align-items: center;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
    }
    .current-date {
        color: #64748B;
        font-size: 14px;
    }
    .current-month {
        background: #3B82F6;
        color: white;
        padding: 6px 14px;
        border-radius: 20px;
        font-size: 13px;
        font-weight: 500;
    }
    
    /* 按钮区域 */
    .button-group {
        display: flex;
        gap: 12px;
        margin: 20px 0;
    }
    .stButton > button {
        border-radius: 12px;
        padding: 10px 24px;
        font-weight: 500;
        transition: all 0.2s;
    }
    .stButton > button:hover {
        transform: scale(1.02);
    }
    
    /* 侧边栏美化 */
    [data-testid="stSidebar"] {
        background: white;
        border-right: 1px solid #E2E8F0;
    }
    [data-testid="stSidebar"] .stMarkdown {
        color: #1E293B;
    }
    
    /* 分隔线 */
    hr {
        margin: 20px 0;
        border-color: #E2E8F0;
    }
</style>
""", unsafe_allow_html=True)


class SafetyStockCalculator:
    """安全库存计算器"""
    
    @staticmethod
    def get_lead_time_coef(days):
        """采购周期系数"""
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
        """质量风险系数"""
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
        """品类策略系数"""
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
        """安全库存计算公式"""
        if pd.isna(future_avg) or pd.isna(past_avg):
            return 0
        if future_avg <= 0 and past_avg <= 0:
            return 0
        base = (float(future_avg) + float(past_avg)) / 2
        combined = quality * 0.4 + category * 0.6
        return base * combined * lead_coef
    
    def process(self, df_mat, df_qual, df_cat):
        """主处理函数"""
        df = df_mat.copy()
        
        # 查找月份列
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
            df['未来3个月月均用量'] = df[future_cols].apply(pd.to_numeric, errors='coerce').mean(axis=1)
        elif '未来3个月月均用量' in df.columns:
            df['未来3个月月均用量'] = pd.to_numeric(df['未来3个月月均用量'], errors='coerce')
        else:
            df['未来3个月月均用量'] = 0
        
        if past_cols:
            df['过去6个月月均用量'] = df[past_cols].apply(pd.to_numeric, errors='coerce').mean(axis=1)
        elif '月均量(半年)' in df.columns:
            df['过去6个月月均用量'] = pd.to_numeric(df['月均量(半年)'], errors='coerce')
        else:
            df['过去6个月月均用量'] = 0
        
        # 采购周期系数
        if '平均交货周期(天)' in df.columns:
            df['采购周期系数'] = df['平均交货周期(天)'].apply(self.get_lead_time_coef)
        else:
            df['采购周期系数'] = 1
        
        # 系数计算
        quality_list = []
        category_list = []
        code_col = df.columns[0]
        for idx, row in df.iterrows():
            code = row[code_col]
            quality_list.append(self.get_quality_score(df_qual, code))
            category_list.append(self.get_category_risk(df_cat, code))
        
        df['质量风险系数'] = quality_list
        df['品类策略系数'] = category_list
        
        # 安全库存
        df['安全库存'] = df.apply(
            lambda x: self.calc_safety_stock(
                x.get('未来3个月月均用量', 0),
                x.get('过去6个月月均用量', 0),
                x.get('质量风险系数', 1),
                x.get('品类策略系数', 1),
                x.get('采购周期系数', 1)
            ), axis=1
        )
        
        # 寄售物料
        if '是否寄售' in df.columns:
            df.loc[df['是否寄售'] == '寄售', '安全库存'] = 0
        
        # 实际库存
        actual_col = None
        for col in df.columns:
            if '实际库存' in col or '6月末' in col:
                actual_col = col
                break
        if actual_col:
            df['实际库存'] = pd.to_numeric(df[actual_col], errors='coerce').fillna(0)
        else:
            df['实际库存'] = 0
        
        # 库存覆盖倍数
        df['库存覆盖倍数'] = df.apply(
            lambda x: x['实际库存'] / x['安全库存'] if x['安全库存'] > 0 else 999,
            axis=1
        )
        df['库存覆盖倍数'] = df['库存覆盖倍数'].replace([np.inf, -np.inf], 999)
        
        # 预警等级
        def get_warning(cov):
            if cov >= 3:
                return "✅ 充足"
            elif cov >= 1.5:
                return "🟢 正常"
            elif cov >= 0.5:
                return "🟡 偏低"
            else:
                return "🔴 严重不足"
        
        df['预警等级'] = df['库存覆盖倍数'].apply(get_warning)
        
        return df


def main():
    # 获取当前时间
    now = datetime.now()
    current_date = now.strftime("%Y年%m月%d日 %A")
    current_month = now.strftime("%Y年%m月")
    
    # 顶部信息栏
    col_title, col_date = st.columns([3, 1])
    with col_title:
        st.markdown('<div class="main-title">📦 安全库存管理系统</div>', unsafe_allow_html=True)
        st.markdown('<div class="sub-title">自动计算 · 智能预警 · 月度更新</div>', unsafe_allow_html=True)
    with col_date:
        st.markdown(f"""
        <div class="top-bar">
            <span class="current-date">📅 {current_date}</span>
            <span class="current-month">{current_month}</span>
        </div>
        """, unsafe_allow_html=True)
    
    # 初始化session
    if 'result' not in st.session_state:
        st.session_state.result = None
    if 'uploaded' not in st.session_state:
        st.session_state.uploaded = False
    
    # ==================== 侧边栏 ====================
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
        new_month_name = st.text_input("新月份名称", value=f"{now.year}年{now.month}月")
        
        update_file = st.file_uploader(
            "上传新月份数据文件",
            type=['xlsx', 'xls'],
            key="monthly",
            help="文件需包含「物料编码」和「用量」两列"
        )
        
        st.markdown("---")
        
        # 公式说明
        with st.expander("📖 计算公式说明"):
            st.markdown("""
            **安全库存公式：**
