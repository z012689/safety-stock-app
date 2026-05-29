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
    .update-success {
        background-color: #d4edda;
        padding: 10px;
        border-radius: 5px;
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
    
    def find_month_columns(self, df):
        """查找数据中的月份列"""
        future_cols = []
        past_cols = []
        all_months = []
        
        for col in df.columns:
            col_str = str(col)
            # 查找未来3个月（M07, M08, M09 或 7月,8月,9月）
            if any(x in col_str for x in ['M07', 'M08', 'M09', '7月', '8月', '9月', 'Jul', 'Aug', 'Sep']):
                if '2025' in col_str or '未来' not in col_str:
                    future_cols.append(col)
                    all_months.append(col)
            # 查找过去6个月（M01-M06 或 1月-6月）
            elif any(x in col_str for x in ['M01', 'M02', 'M03', 'M04', 'M05', 'M06', 
                                              '1月', '2月', '3月', '4月', '5月', '6月',
                                              'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun']):
                past_cols.append(col)
                all_months.append(col)
        
        return future_cols, past_cols, sorted(all_months)
    
    def update_monthly_data(self, df, new_month_name, new_month_data, month_columns):
        """月度数据滚动更新"""
        df_updated = df.copy()
        
        # 获取所有月份列（按时间排序）
        month_cols = [col for col in month_columns if col in df_updated.columns]
        month_cols_sorted = sorted(month_cols, key=lambda x: str(x))
        
        if len(month_cols_sorted) >= 6:
            # 删除最旧的月份列
            oldest_col = month_cols_sorted[0]
            df_updated = df_updated.drop(columns=[oldest_col])
            
            # 重命名剩余的月份列（可选：保持格式统一）
            for col in month_cols_sorted[1:]:
                pass  # 保留原列名
        
        # 添加新月份数据
        if new_month_data is not None:
            df_updated[new_month_name] = new_month_data
        
        return df_updated
    
    def process_raw_materials(self, df_materials, df_quality, df_category):
        """处理原料安全库存计算"""
        df = df_materials.copy()
        
        # 查找月份列
        future_cols, past_cols, all_month_cols = self.find_month_columns(df)
        
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
        
        # 获取质量风险系数
        quality_dict = {}
        if df_quality is not None and '物料编码' in df_quality.columns:
            for _, row in df_quality.iterrows():
                code = row.get('物料编码')
                score = row.get('质量风险等级得分')
                if pd.notna(code) and pd.notna(score):
                    try:
                        quality_dict[str(code)] = float(score)
                    except:
                        quality_dict[str(code)] = 1
        
        # 获取品类策略系数
        category_dict = {}
        if df_category is not None:
            code_col = None
            risk_col = None
            for col in df_category.columns:
                if '物料' in col:
                    code_col = col
                if '风险' in col:
                    risk_col = col
            if code_col and risk_col:
                for _, row in df_category.iterrows():
                    code = row.get(code_col)
                    risk = row.get(risk_col)
                    if pd.notna(code) and pd.notna(risk):
                        try:
                            category_dict[str(code)] = float(risk)
                        except:
                            category_dict[str(code)] = 1
        
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
        
        # 添加月份列信息到结果中（用于后续更新）
        df['_month_columns'] = str(all_month_cols)
        
        return df, future_cols, past_cols, all_month_cols


def main():
    st.title("📦 安全库存管理系统")
    st.markdown("---")
    
    # 初始化session state
    if 'df_result' not in st.session_state:
        st.session_state.df_result = None
    if 'future_cols' not in st.session_state:
        st.session_state.future_cols = None
    if 'past_cols' not in st.session_state:
        st.session_state.past_cols = None
    if 'all_month_cols' not in st.session_state:
        st.session_state.all_month_cols = None
    if 'original_data' not in st.session_state:
        st.session_state.original_data = None
    
    # 侧边栏
    with st.sidebar:
        st.header("⚙️ 系统配置")
        
        uploaded_file = st.file_uploader(
            "上传安全库存Excel文件",
            type=['xlsx', 'xls'],
            help="请上传包含安全库存规则和数据的Excel文件"
        )
        
        st.markdown("---")
        
        # ========== 月度更新区域 ==========
        st.subheader("📅 月度数据更新")
        
        # 选择更新方式
        update_mode = st.radio(
            "更新方式",
            ["📝 手动输入新月份数据", "📁 上传新月份Excel文件"],
            help="选择手动输入每个物料的数据，或上传包含新月份数据的文件"
        )
        
        # 新月份名称
        next_month = datetime.now().replace(day=1) + timedelta(days=32)
        default_month = next_month.strftime("2025年%m月")
        new_month_name = st.text_input("新月份名称", value=default_month)
        
        # 手动输入模式
        if update_mode == "📝 手动输入新月份数据":
            if st.session_state.df_result is not None:
                st.write(f"共 {len(st.session_state.df_result)} 个物料需要输入数据")
                
                # 创建输入区域（简化版：只显示前10个）
                manual_data = {}
                df_display = st.session_state.df_result.head(20)[['物料编码', 'SAP编码']]
                
                for idx, row in df_display.iterrows():
                    code = row['物料编码']
                    manual_data[code] = st.number_input(
                        f"{code} - {row.get('SAP编码', '')}",
                        value=0.0,
                        step=100.0,
                        key=f"manual_{idx}"
                    )
                
                if len(st.session_state.df_result) > 20:
                    st.info(f"显示前20个物料，共{len(st.session_state.df_result)}个")
            else:
                st.info("请先计算安全库存，再进行月度更新")
        
        # 文件上传模式
        else:
            new_month_file = st.file_uploader(
                "上传包含新月份数据的Excel文件",
                type=['xlsx', 'xls'],
                help="文件应包含物料编码和新月份用量数据"
            )
        
        st.markdown("---")
        
        # 执行月度更新按钮
        update_btn = st.button("🔄 执行月度更新", type="secondary", use_container_width=True)
        
        st.markdown("---")
        
        # 计算按钮
        calculate_btn = st.button("🚀 开始计算安全库存", type="primary", use_container_width=True)
    
    # 主要内容区域
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
                
                # 首次计算
                if calculate_btn:
                    with st.spinner("正在计算安全库存..."):
                        df_result, future_cols, past_cols, all_month_cols = calculator.process_raw_materials(
                            df_materials, df_quality, df_category
                        )
                        
                        st.session_state.df_result = df_result
                        st.session_state.future_cols = future_cols
                        st.session_state.past_cols = past_cols
                        st.session_state.all_month_cols = all_month_cols
                        st.session_state.original_data = df_materials.copy()
                        
                    st.success("✅ 安全库存计算完成！")
                    
                    # 显示结果摘要
                    self_show_results(st.session_state.df_result)
                
                # 执行月度更新
                elif update_btn:
                    if st.session_state.df_result is None:
                        st.warning("⚠️ 请先计算安全库存，再进行月度更新")
                    else:
                        with st.spinner("正在执行月度数据更新..."):
                            new_data_series = None
                            
                            if update_mode == "📝 手动输入新月份数据":
                                # 从手动输入收集数据
                                new_data_dict = {}
                                for idx, row in st.session_state.df_result.head(20).iterrows():
                                    code = row['物料编码']
                                    val = manual_data.get(code, 0)
                                    new_data_dict[code] = val
                                new_data_series = pd.Series(new_data_dict)
                            
                            elif update_mode == "📁 上传新月份Excel文件" and new_month_file:
                                # 从上传文件读取数据
                                new_month_df = pd.read_excel(new_month_file)
                                if '物料编码' in new_month_df.columns and '用量' in new_month_df.columns:
                                    new_data_series = new_month_df.set_index('物料编码')['用量']
                            
                            if new_data_series is not None:
                                # 执行滚动更新
                                df_updated = calculator.update_monthly_data(
                                    st.session_state.original_data,
                                    new_month_name,
                                    new_data_series,
                                    st.session_state.all_month_cols
                                )
                                
                                # 重新计算安全库存
                                df_result_updated, future_cols, past_cols, all_month_cols = calculator.process_raw_materials(
                                    df_updated, df_quality, df_category
                                )
                                
                                st.session_state.df_result = df_result_updated
                                st.session_state.original_data = df_updated
                                st.session_state.all_month_cols = all_month_cols
                                
                                st.success(f"✅ 月度更新完成！已添加 {new_month_name} 数据")
                                st.info(f"📊 已自动删除最旧的月份数据，保持6个月历史数据")
                                
                                # 显示更新后的结果
                                self_show_results(st.session_state.df_result)
                            else:
                                if update_mode == "📁 上传新月份Excel文件":
                                    st.error("请上传包含'物料编码'和'用量'列的Excel文件")
                
                elif st.session_state.df_result is not None:
                    # 显示已缓存的结果
                    self_show_results(st.session_state.df_result)
                    
            else:
                st.error("未找到原料安全库存数据表")
        except Exception as e:
            st.error(f"处理文件时出错: {str(e)}")
    else:
        st.info("👈 请先上传安全库存Excel文件")
        
        with st.expander("📖 使用说明"):
            st.markdown("""
            ### 系统功能说明
            
            #### 1. 安全库存计算
            - 上传包含物料用量数据的Excel文件
            - 系统自动计算安全库存
            
            #### 2. 月度数据更新（重要）
            月度更新功能会自动：
            - **添加**新月份的数据
            - **删除**最旧的月份数据
            - **重新计算**安全库存
            
            **更新方式：**
            - **手动输入**：逐个物料输入新月份用量
            - **文件上传**：上传包含新月份数据的Excel
            
            **更新效果：**
