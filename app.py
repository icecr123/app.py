import streamlit as st
import pandas as pd
import numpy as np
import re
from io import BytesIO

# ================= 页面配置 =================
st.set_page_config(page_title="返佣计算工具 V5-完整版", layout="wide")
st.title("🧮 月度回款返佣自动计算工具 (V5-完整版)")
st.markdown("""
**本次更新重点：**
1. **精准拆分**：同一订单多笔服务费不再合并，逐行提取。
2. **智能合并**：同一批次多笔罚息/逾期自动合并为一行。
3. **期次校准**：基于历史还款记录自动推算当前期次，解决“未匹配”问题。
4. **返佣计算**：完整补全政策匹配与返佣金额计算逻辑。
""")

# ================= 辅助函数 =================
def safe_float(val):
    """安全转换金额"""
    if pd.isna(val): return 0.0
    s = str(val).strip()
    if s in ['无', 'None', 'nan', '']: return 0.0
    try: return float(s.replace(',', ''))
    except ValueError: return 0.0

def clean_order_id(oid):
    """暴力清洗订单号：提取纯数字"""
    if pd.isna(oid): return ""
    s = str(oid).strip()
    nums = re.findall(r'\d+', s)
    return "".join(nums) if nums else s

def parse_xy_product(product_name):
    """解析 x+y 产品格式"""
    if pd.isna(product_name): return False, 0, 0
    name_str = str(product_name).strip()
    match = re.search(r'(\d+)\+(\d+)', name_str)
    if match: return True, int(match.group(1)), int(match.group(2))
    return False, 0, 0

def count_periods(period_str):
    """统计期数"""
    if pd.isna(period_str): return 1
    numbers = re.findall(r'\d+', str(period_str))
    return max(len(numbers), 1)

def determine_period(order_id, period_str, history_map):
    """智能判断还款期次"""
    if "延期" in str(period_str): return "" 
    
    match = re.search(r'(\d+)', str(period_str))
    current_target_period = int(match.group(1)) if match else None
    last_paid = history_map.get(order_id, 0)

    if current_target_period:
        if current_target_period > last_paid:
            return f"第{current_target_period}期"
        else:
            return f"第{current_target_period}期(复)" 
    else:
        next_period = last_paid + 1
        return f"第{next_period}期"

def calculate_commission(row, policy_map):
    """计算单笔返佣"""
    merchant = str(row.get('收款商户', '')).strip()
    product = str(row.get('产品名称', '')).strip()
    period_str = str(row.get('还款期次', '')).strip()
    amount = safe_float(row.get('分期金额', 0))
    
    key = f"{merchant}_{product}"
    policy = policy_map.get(key, {})
    if not policy: return pd.Series(['否', '0.0000', 0.0])
        
    is_xy, x_val, y_val = parse_xy_product(product)
    ratio = 0.0
    has_comm = '否'
    p_num = count_periods(period_str)
    
    if is_xy:
        last_period = 0
        numbers = re.findall(r'\d+', period_str)
        if numbers: last_period = int(numbers[-1])
            
        if isinstance(last_period, int) and 0 < last_period <= x_val:
            raw_ratio = policy.get('X-返佣', 0)
        else:
            raw_ratio = policy.get('Y-返佣', 0)
    else:
        raw_ratio = policy.get('等额-返佣', 0)
        
    ratio = safe_float(raw_ratio)
    if ratio > 0: has_comm = '是'
        
    comm_amount = 0.0
    if ratio > 0 and amount > 0:
        comm_amount = amount * ratio * p_num
        
    return pd.Series([has_comm, f"{ratio:.4f}", round(comm_amount, 2)])

# ================= 文件上传区 =================
st.sidebar.header("📁 数据文件上传")
file_ledger = st.sidebar.file_uploader("1. 分账支付记录 (线上)", type=['xls', 'xlsx'])
file_payment = st.sidebar.file_uploader("2. 代付记录 (线下)", type=['xls', 'xlsx'])
file_order = st.sidebar.file_uploader("3. 订单主表", type=['xls', 'xlsx'])
file_detail = st.sidebar.file_uploader("4. 订单支付明细 (核对期次)", type=['xls', 'xlsx'])
file_policy = st.sidebar.file_uploader("5. 返佣政策详情", type=['xls', 'xlsx'])

if file_ledger and file_payment and file_order and file_detail and file_policy:
    try:
        # 1. 读取文件 (强制指定订单相关列为字符串)
        df_ledger = pd.read_excel(file_ledger, dtype={'业务订单号': str, '订单编号': str})
        df_payment_raw = pd.read_excel(file_payment, dtype={'业务订单号': str, '订单编号': str})
        df_order = pd.read_excel(file_order, dtype={'订单号': str, '订单编号': str})
        df_detail = pd.read_excel(file_detail, dtype={'业务订单号': str, '订单编号': str})
        df_policy_raw = pd.read_excel(file_policy)

        # 2. 统一列名 & 清洗 ID
        col_map = {'订单号': '订单编号', '业务订单号': '订单编号'}
        for df in [df_ledger, df_payment_raw, df_order, df_detail]:
            df.rename(columns=col_map, inplace=True)
            if '订单编号' in df.columns:
                df['订单编号'] = df['订单编号'].apply(clean_order_id)

        # 3. 建立映射字典
        st.info("📊 正在构建订单与政策映射...")
        order_map = {}
        for _, row in df_order.iterrows():
            oid = row.get('订单编号', '')
            if oid and oid != "":
                order_map[oid] = {
                    '产品名称': row.get('产品名称', ''),
                    '下单时间': row.get('下单时间', ''),
                    '订单状态': row.get('订单状态', ''),
                    '维护商务': row.get('业务员', ''),
                    '付款人': row.get('客户姓名', ''),
                    '收款商户': row.get('机构简称', ''),
                    '分期金额': row.get('分期金额', 0)
                }

        # 构建明细表映射 (用于匹配还款类型和推算期次)
        history_map = {} 
        if '订单编号' in df_detail.columns:
            time_col = '支付时间' if '支付时间' in df_detail.columns else '完成时间'
            if time_col in df_detail.columns:
                df_detail = df_detail.sort_values(by=['订单编号', time_col])
                
            df_detail['还款期次_idx'] = df_detail.groupby('订单编号').cumcount() + 1
            for _, r in df_detail.iterrows():
                oid = r['订单编号']
                idx = r['还款期次_idx']
                if oid not in history_map or idx > history_map[oid]:
                    history_map[oid] = idx

        # 构建政策映射
        policy_map = {}
        for _, row in df_policy_raw.iterrows():
            inst = str(row.get('机构名称', '')).strip()
            prod = str(row.get('产品名称', '')).strip()
            if inst and prod:
                policy_map[f"{inst}_{prod}"] = {
                    '等额-返佣': row.get('等额-返佣', 0),
                    'X-返佣': row.get('X-返佣', 0),
                    'Y-返佣': row.get('Y-返佣', 0)
                }

        # 4. 处理线上分账 (Ledger)
        st.info("💻 正在处理线上分账数据...")
        res_list = []
        for _, row in df_ledger.iterrows():
            oid = row.get('订单编号', '')
            if not oid or oid == "": continue
            info = order_map.get(oid, {})
            new_row = {
                '业务订单号': oid,
                '产品名称': info.get('产品名称', ''),
                '收款商户': info.get('收款商户', ''),
                '付款人': info.get('付款人', ''),
                '分期金额': info.get('分期金额', 0),
                '还款期次': str(row.get('还款期次', '')),
                '支付时间': row.get('支付时间', ''),
                '服务费': row.get('服务费', 0),
                '逾期费用': row.get('逾期费', row.get('罚息', 0)),
                '还款方式': '线上还款',
                '下单时间': info.get('下单时间', ''),
                '订单状态': info.get('订单状态', ''),
                '维护商务': info.get('维护商务', ''),
                '备注': ''
            }
            res_list.append(new_row)

        # 5. 处理线下代付 (Payment)
        st.info("🏦 正在处理线下代付数据...")
        if not df_payment_raw.empty:
            if '支付批次号' not in df_payment_raw.columns:
                raise Exception("❌ 代付记录表中缺少【支付批次号】列！")
                
            grouped = df_payment_raw.groupby(['支付批次号', '订单编号'])
            
            for (batch_id, oid), group in grouped:
                if not oid or oid == "": continue
                info = order_map.get(oid, {}) 
                
                # 5.1 提取服务费 (逐行，不合并)
                service_rows = group[group['系统备注'].astype(str).str.contains('服务费', na=False)]
                for _, s_row in service_rows.iterrows():
                    note = str(s_row['系统备注']).strip()
                    amt = safe_float(s_row['清分金额'])
                    finish_time = s_row['完成时间']
                    
                    if '延期手续费' in note: note = '延期服务费'
                    
                    period = ""
                    if '延期服务费' not in note:
                        match = re.search(r'(\d+)', note)
                        if match:
                            period = f"第{match.group(1)}期"
                        else:
                            period = determine_period(oid, note, history_map)
                    
                    new_row = {
                        '业务订单号': oid, '产品名称': info.get('产品名称', ''),
                        '收款商户': info.get('收款商户', ''), '付款人': info.get('付款人', ''),
                        '分期金额': info.get('分期金额', 0), '支付时间': finish_time, 
                        '服务费': amt, '逾期费用': 0.0, '还款方式': '线下代付',
                        '下单时间': info.get('下单时间', ''), '订单状态': info.get('订单状态', ''),
                        '维护商务': info.get('维护商务', ''), '备注': note, '还款期次': period
                    }
                    res_list.append(new_row)

                # 5.2 提取罚息/逾期 (合并逻辑)
                penalty_keywords = ['罚息', '逾期', '违约金']
                penalty_mask = group['系统备注'].astype(str).apply(lambda x: any(k in str(x) for k in penalty_keywords))
                penalty_rows = group[penalty_mask]
                
                if not penalty_rows.empty:
                    total_penalty = penalty_rows['清分金额'].sum()
                    rep_time = penalty_rows['完成时间'].min()
                    period = determine_period(oid, "罚息", history_map)

                    new_row = {
                        '业务订单号': oid, '产品名称': info.get('产品名称', ''),
                        '收款商户': info.get('收款商户', ''), '付款人': info.get('付款人', ''),
                        '分期金额': info.get('分期金额', 0), '支付时间': rep_time, 
                        '服务费': 0.0, '逾期费用': total_penalty, '还款方式': '线下代付',
                        '下单时间': info.get('下单时间', ''), '订单状态': info.get('订单状态', ''),
                        '维护商务': info.get('维护商务', ''), '备注': '', '还款期次': period
                    }
                    res_list.append(new_row)

        # 6. 生成最终 DataFrame 并计算返佣
        st.info("📝 正在整合数据并计算返佣...")
        df_all = pd.DataFrame(res_list)
        
        if not df_all.empty:
            comm_results = df_all.apply(lambda row: calculate_commission(row, policy_map), axis=1)
            df_all['是否有返佣'] = comm_results[0]
            df_all['返佣比例'] = comm_results[1]
            df_all['返佣金额'] = comm_results[2]

            st.success(f"✅ 处理完成！共生成 {len(df_all)} 条有效记录。")
            st.dataframe(df_all.head(50))

            # 导出 Excel
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_all.to_excel(writer, index=False, sheet_name='返佣计算结果')
            output.seek(0)
            
            st.download_button(
                label="📥 下载处理后的 Excel 文件",
                data=output,
                file_name="月度回款返佣计算结果_完整版.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"❌ 处理出错: {e}")
        st.exception(e)
else:
    st.warning("⚠️ 请在左侧上传全部 5 个 Excel 文件以开始处理...")
