import streamlit as st
import pandas as pd
import numpy as np
import re
import io
from datetime import datetime

# --- 页面配置 ---
st.set_page_config(page_title="返佣计算小工具", layout="wide")
st.title("🧮 月度回款返佣自动计算工具 (修正版)")
st.markdown("请依次上传以下 5 个文件，工具将自动完成计算并生成结果。")

# --- 核心辅助函数 ---

def safe_float(val):
    """安全转换浮点数"""
    if pd.isna(val): return 0.0
    s = str(val).strip()
    if s in ['无', 'None', 'nan', '']: return 0.0
    try: return float(s)
    except ValueError: return 0.0

def clean_order_id(order_id):
    """清洗订单号，去除 .0 后缀"""
    if pd.isna(order_id): return ''
    s = str(order_id).strip()
    if s.endswith('.0'): s = s[:-2]
    return s

def parse_xy_product(product_name):
    """解析产品名称中的 x+y 格式"""
    if pd.isna(product_name): return False, 0, 0
    name_str = str(product_name).strip()
    match = re.search(r'(\d+)\+(\d+)', name_str)
    if match: return True, int(match.group(1)), int(match.group(2))
    return False, 0, 0

def count_periods(period_str):
    """统计还款期次数量"""
    if pd.isna(period_str): return 1
    p_str = str(period_str)
    numbers = re.findall(r'\d+', p_str)
    return max(len(numbers), 1)

def calculate_commission(row, policy_map):
    """计算单笔返佣"""
    merchant = str(row.get('收款商户', '')).strip()
    product = str(row.get('产品名称', '')).strip()
    period_str = str(row.get('还款期次', '')).strip()
    
    # 基础数据获取
    is_xy, x_val, y_val = parse_xy_product(product)
    total_amount = safe_float(row.get('分期总金额', 0))
    service_fee = safe_float(row.get('服务费', 0))
    
    # 查找政策 (简化逻辑，实际需根据你的政策表结构调整)
    # 假设 policy_map 结构为 {商户名: {产品类型: 费率}}
    rate = 0.0
    if merchant in policy_map:
        # 这里需要根据你的具体业务逻辑完善匹配规则
        # 示例：默认取该商户的第一个费率，或者根据 x/y 值判断
        rates = list(policy_map[merchant].values())
        if rates: rate = rates[0] 

    commission = (total_amount + service_fee) * rate
    return round(commission, 2)

# --- 主程序逻辑 ---

def main():
    st.sidebar.header("📂 文件上传区")
    
    file_ledger = st.sidebar.file_uploader("1. 分账支付记录.xls", type=['xls', 'xlsx'])
    file_payment = st.sidebar.file_uploader("2. 代付记录.xls", type=['xls', 'xlsx'])
    file_order_main = st.sidebar.file_uploader("3. 订单.xls", type=['xls', 'xlsx'])
    file_order_detail = st.sidebar.file_uploader("4. 订单支付明细.xlsx", type=['xlsx', 'xls'])
    file_policy = st.sidebar.file_uploader("5. 返佣政策详情.xls", type=['xls', 'xlsx'])

    if st.button("🚀 开始计算"):
        if not all([file_ledger, file_payment, file_order_main, file_order_detail, file_policy]):
            st.error("❌ 请上传所有 5 个文件！")
            return

        try:
            # 1. 读取文件
            df_ledger = pd.read_excel(file_ledger)
            df_payment = pd.read_excel(file_payment)
            df_order_main = pd.read_excel(file_order_main)
            df_order_detail = pd.read_excel(file_order_detail)
            df_policy = pd.read_excel(file_policy)
            # ================= 【关键修复】自动对齐列名 =================
            # 这一步是为了防止 Excel 表头名字不一样导致 KeyError
            
            # 1. 检查并修复【订单主表】(解决你现在的报错)
            # 如果表里有 '业务订单号'，就把它改名为 '订单编号'
            if '业务订单号' in df_order_main.columns:
                df_order_main.rename(columns={'业务订单号': '订单编号'}, inplace=True)
            
            # 2. 检查并修复【订单支付明细】
            if '业务订单号' in df_order_detail.columns:
                df_order_detail.rename(columns={'业务订单号': '订单编号'}, inplace=True)
            
            # 3. 检查并修复【代付记录】
            if '订单编号' in df_payment.columns:
                df_payment.rename(columns={'订单编号': '业务订单号'}, inplace=True)
            # ==========================================================
            # ================= 【新增】万能列名修复补丁 =================
            # 这一步是为了防止 Excel 表头名字不一样导致 KeyError
            
            # 1. 统一【订单支付明细】的列名
            if '业务订单号' in df_order_detail.columns:
                df_order_detail.rename(columns={'业务订单号': '订单编号'}, inplace=True)
            
            # 2. 统一【代付记录】的列名
            if '订单编号' in df_payment.columns:
                df_payment.rename(columns={'订单编号': '业务订单号'}, inplace=True)

            # 3. 统一【订单主表】的列名
            if '业务订单号' in df_order_main.columns:
                df_order_main.rename(columns={'业务订单号': '订单编号'}, inplace=True)
            
            # ================= 【核心】生成序号与双键匹配逻辑 =================
            
            # A. 处理【订单支付明细】(作为查找字典)
            # 清洗订单号，去除 .0
            df_order_detail['订单编号'] = df_order_detail['订单编号'].astype(str).str.strip().str.replace('.0', '', regex=False)
            # 生成组内序号 (第1笔, 第2笔...)
            df_order_detail['组内序号'] = df_order_detail.groupby('订单编号').cumcount() + 1
            
            # 构建查找字典：key是(订单号, 序号)，value是还款类型
            detail_lookup = dict(zip(
                zip(df_order_detail['订单编号'], df_order_detail['组内序号']),
                df_order_detail['还款类型'] 
            ))

            # B. 处理【代付记录】(作为主表去匹配)
            df_payment['业务订单号'] = df_payment['业务订单号'].astype(str).str.strip().str.replace('.0', '', regex=False)
            # 同样生成组内序号
            df_payment['组内序号'] = df_payment.groupby('业务订单号').cumcount() + 1
            
            # C. 执行双键匹配
            def get_repayment_type(row):
                oid = row['业务订单号']
                seq = row['组内序号']
                return detail_lookup.get((oid, seq), '未知')

            # 将匹配到的结果写入新列
            df_payment['还款期次'] = df_payment.apply(get_repayment_type, axis=1)
            # ================= 核心修正逻辑开始 =================
            
            # A. 预处理：构建“订单支付明细”的查找字典
            # 目标：{(订单编号, 组内序号): 还款类型}
            df_order_detail['订单编号'] = df_order_detail['订单编号'].apply(clean_order_id)
            # 按订单号分组，生成 1, 2, 3... 的序号 (cumcount从0开始，所以+1)
            df_order_detail['组内序号'] = df_order_detail.groupby('订单编号').cumcount() + 1
            
            detail_lookup = dict(zip(
                zip(df_order_detail['订单编号'], df_order_detail['组内序号']),
                df_order_detail['还款类型'] 
            ))

            # B. 处理“代付记录”
            df_payment['业务订单号'] = df_payment['业务订单号'].apply(clean_order_id)
            # 同样按订单号分组，生成 1, 2, 3... 的序号
            df_payment['组内序号'] = df_payment.groupby('业务订单号').cumcount() + 1
            
            # C. 精准匹配：利用 (订单号 + 序号) 双键匹配
            def get_repayment_type(row):
                oid = row['业务订单号']
                seq = row['组内序号']
                # 在字典中查找，找不到则返回 '未知'
                return detail_lookup.get((oid, seq), '未知')

            df_payment['还款期次'] = df_payment.apply(get_repayment_type, axis=1)
            
            # ================= 核心修正逻辑结束 =================

            # 3. 合并主订单信息 (获取金额、产品等)
            # 注意：这里假设 df_order_main 中订单号是唯一的。如果不唯一，需要去重或特殊处理
            df_order_main['订单编号'] = df_order_main['订单编号'].apply(clean_order_id)
            
            final_df = pd.merge(
                df_payment, 
                df_order_main[['订单编号', '分期总金额', '服务费', '产品名称', '收款商户']], 
                left_on='业务订单号', 
                right_on='订单编号', 
                how='left'
            )

            # 4. 加载返佣政策并计算
            # 这里简单演示如何构建 policy_map，你需要根据实际 Excel 结构调整
            # 假设政策表有 '商户名称' 和 '返佣比例' 列
            policy_map = {}
            for _, row in df_policy.iterrows():
                m_name = str(row.get('商户名称', '')).strip()
                rate = safe_float(row.get('返佣比例', 0))
                if m_name:
                    policy_map[m_name] = {'default': rate}

            # 应用计算函数
            final_df['预估返佣'] = final_df.apply(lambda r: calculate_commission(r, policy_map), axis=1)

            st.success("✅ 数据处理完成！正在生成预览...")
            
            # 显示前 10 行预览
            st.dataframe(final_df.head(10))
            
            # 提供下载按钮
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                final_df.to_excel(writer, index=False, sheet_name='计算结果')
            
            st.download_button(
                label="📥 下载计算结果 Excel",
                data=output.getvalue(),
                file_name=f"返佣计算结果_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"计算过程中出现错误: {str(e)}")
            # 打印详细报错以便调试
            import traceback
            st.code(traceback.format_exc())

if __name__ == "__main__":
    main()
