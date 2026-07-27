import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import datetime
import re

# ================= 页面配置 =================
st.set_page_config(page_title="月度回款返佣计算工具 V33-延期服务费专项修复版", layout="centered")
st.title("🧮 月度回款返佣自动计算工具 (V33)")
st.markdown("""
**V33 核心修复说明：**
1.  **延期服务费专项处理**：
    *   识别到“延期服务费”时，**不再匹配期次**（避免占用期次名额）。
    *   【还款期次】列强制留空。
    *   【备注】列自动填入“延期服务费”。
2.  **期次顺序匹配优化**：
    *   只有非延期的正常还款才会触发计数器+1。
    *   确保“平账第1期、第2期...”严格对应真实的还款顺序。
3.  **数据聚合增强**：
    *   同一订单+同批次下的多行罚息/服务费依然保持合并。
""")

# --- 核心逻辑函数 ---

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
    # 如果不是数字字符串（如"平账第1期"），视为1期
    if not p_str.isdigit(): return 1
    return max(int(p_str), 1)

def calculate_commission(row, policy_map):
    """计算单笔返佣"""
    merchant = str(row.get('收款商户', '')).strip()
    product = str(row.get('产品名称', '')).strip()
    period_str = str(row.get('还款期次', '')).strip()
    amount = safe_float(row.get('分期金额', 0))
    
    # 如果没有期次（如延期服务费），通常不计算返佣或按特定规则，这里默认不返佣
    if not period_str:
        return pd.Series(['否', '0.0000', 0.0])

    key = f"{merchant}_{product}"
    policy = policy_map.get(key, {})
    if not policy: return pd.Series(['否', '0.0000', 0.0])
    
    is_xy, x_val, y_val = parse_xy_product(product)
    ratio = 0.0
    has_comm = '否'
    
    # 提取期数数字
    p_num = 1
    numbers = re.findall(r'\d+', period_str)
    if numbers: 
        p_num = int(numbers[-1])
    
    if is_xy:
        last_period = p_num
        if 0 < last_period <= x_val: raw_ratio = policy.get('X-返佣', 0)
        else: raw_ratio = policy.get('Y-返佣', 0)
    else:
        raw_ratio = policy.get('等额-返佣', 0)
        
    ratio = safe_float(raw_ratio)
    if ratio > 0: has_comm = '是'
    
    comm_amount = 0.0
    if ratio > 0 and amount > 0:
        comm_amount = amount * ratio
        
    return pd.Series([has_comm, f"{ratio:.4f}", round(comm_amount, 2)])

# --- 主处理流程 ---

def process_data(ledger_file, payment_file, order_file, detail_file, policy_file):
    # 1. 读取所有文件
    df_ledger = pd.read_excel(ledger_file, dtype=str)
    df_payment_raw = pd.read_excel(payment_file, dtype=str)
    df_order = pd.read_excel(order_file, dtype=str)
    df_detail = pd.read_excel(detail_file, dtype=str)
    df_policy_raw = pd.read_excel(policy_file, dtype=str)

    # 2. 预处理：建立基础映射字典
    order_map = {}
    for _, row in df_order.iterrows():
        oid = clean_order_id(row.get('订单号'))
        if oid:
            order_map[oid] = {
                '产品名称': row.get('产品名称', ''),
                '下单时间': row.get('下单时间', ''),
                '订单状态': row.get('订单状态', ''),
                '维护商务': row.get('业务员', ''),
                '付款人': row.get('客户姓名', ''),
                '收款商户': row.get('机构简称', ''),
                '分期金额': row.get('分期金额', 0)
            }

    # --- 构建线下还款期次队列 ---
    detail_queue_map = {}
    if '支付方式' in df_detail.columns:
        mask_offline = df_detail['支付方式'].astype(str).str.contains('线下', na=False)
        df_detail_offline = df_detail[mask_offline].copy()
    else:
        df_detail_offline = df_detail.copy()
        
    if not df_detail_offline.empty:
        time_col = '支付时间'
        if '支付时间_dt' in df_detail_offline.columns:
            time_col = '支付时间_dt'
        elif '支付时间' in df_detail_offline.columns:
             df_detail_offline['支付时间_dt'] = pd.to_datetime(df_detail_offline['支付时间'], errors='coerce')
             time_col = '支付时间_dt'
        
        if time_col in df_detail_offline.columns:
            grouped_details = df_detail_offline.groupby('订单编号')
            for oid, group in grouped_details:
                try:
                    sorted_group = group.sort_values(by=time_col, ascending=True)
                except:
                    sorted_group = group
                # 仅保留还款类型列作为队列
                queue = sorted_group['还款类型'].tolist()
                detail_queue_map[clean_order_id(oid)] = queue

    # 政策映射
    policy_map = {}
    for _, row in df_policy_raw.iterrows():
        inst = str(row.get('机构名称', '')).strip()
        prod = str(row.get('产品名称', '')).strip()
        if inst and prod:
            key = f"{inst}_{prod}"
            policy_map[key] = {
                '等额-返佣': row.get('等额-返佣', 0),
                'X-返佣': row.get('X-返佣', 0),
                'Y-返佣': row.get('Y-返佣', 0),
                '返佣开始时间': str(row.get('返佣开始时间', '')).strip()
            }

    res_list = []

    # 3. 处理分账记录 (线上)
    for _, row in df_ledger.iterrows():
        oid = clean_order_id(row.get('业务订单号'))
        info = order_map.get(oid, {})
        period_str = str(row.get('还款期次', ''))
        remark_parts = []
        if '延期手续费' in period_str:
            remark_parts.append("延期服务费")
            
        new_row = {
            '业务订单号': oid,
            '产品名称': info.get('产品名称', row.get('产品名称', '')),
            '收款商户': info.get('收款商户', row.get('收款商户', '')),
            '付款人': info.get('付款人', row.get('付款人', '')),
            '分期金额': row.get('分期金额', 0),
            '还款期次': period_str,
            '支付时间': row.get('支付时间', ''),
            '服务费': row.get('服务费', 0),
            '逾期费用': row.get('逾期费', row.get('罚息', 0)),
            '还款方式': '线上还款',
            '下单时间': info.get('下单时间', ''),
            '订单状态': info.get('订单状态', ''),
            '维护商务': info.get('维护商务', ''),
            '备注': "，".join(remark_parts)
        }
        res_list.append(new_row)

    # 4. 处理代付记录 (线下 - 专项修复版)
    if not df_payment_raw.empty:
        offline_counter = {} # 记录每个订单当前已经消费到了第几期
        
        # 按 业务订单号 + 支付批次号 分组
        grouped = df_payment_raw.groupby(['业务订单号', '支付批次号'])
        
        for (oid, batch_id), group in grouped:
            oid_clean = clean_order_id(oid)
            info = order_map.get(oid_clean, {})
            
            total_service, total_overdue, service_time, is_delay_fee = 0.0, 0.0, None, False
            
            # 检查该批次是否包含“延期服务费”
            # 只要组内任意一行备注包含“延期服务费”，整组视为延期服务费处理
            for _, r in group.iterrows():
                note = str(r.get('系统备注', ''))
                amt = safe_float(r.get('清分金额', 0))
                finish_time = r.get('完成时间', '')
                
                if '延期服务费' in note or '延期手续费' in note:
                    is_delay_fee = True
                
                if '服务费' in note and '返服务费' not in note:
                    total_service += amt
                    if pd.notna(finish_time) and str(finish_time).strip() != '':
                        service_time = finish_time
                elif '罚息' in note or '逾期' in note:
                    total_overdue += amt

            final_pay_time = service_time if service_time else group.iloc[0].get('完成时间', '')
            
            # --- 核心分支判断 ---
            repayment_type = ""
            remark_content = ""
            
            if is_delay_fee:
                # 场景A：延期服务费 -> 不消耗期次，期次留空
                repayment_type = "" 
                remark_content = "延期服务费"
            else:
                # 场景B：正常还款 -> 消耗期次
                current_count = offline_counter.get(oid_clean, 0)
                queue = detail_queue_map.get(oid_clean, [])
                
                if current_count < len(queue):
                    repayment_type = queue[current_count]
                else:
                    repayment_type = f"超出明细范围({current_count+1})"
                
                # 计数器+1
                offline_counter[oid_clean] = current_count + 1
            
            new_row = {
                '业务订单号': oid_clean,
                '产品名称': info.get('产品名称', ''),
                '收款商户': info.get('收款商户', ''),
                '付款人': info.get('付款人', ''),
                '分期金额': info.get('分期金额', 0),
                '还款期次': repayment_type,
                '支付时间': final_pay_time,
                '服务费': total_service,
                '逾期费用': total_overdue,
                '还款方式': '线下代付',
                '下单时间': info.get('下单时间', ''),
                '订单状态': info.get('订单状态', ''),
                '维护商务': info.get('维护商务', ''),
                '备注': remark_content
            }
            res_list.append(new_row)

    df_all = pd.DataFrame(res_list)

    # 5. 计算返佣
    comm_results = df_all.apply(lambda row: calculate_commission(row, policy_map), axis=1)
    df_all['是否有返佣'] = comm_results[0]
    df_all['返佣比例'] = comm_results[1]
    df_all['返佣金额'] = comm_results[2]

    # 6. 补充日期备注与校验
    def check_date_and_adjust(row):
        order_time_str = str(row.get('下单时间', '')).strip()
        merchant = str(row.get('收款商户', '')).strip()
        product = str(row.get('产品名称', '')).strip()
        current_remarks = row['备注']
        if pd.isna(current_remarks): current_remarks = ""
        
        key = f"{merchant}_{product}"
        policy = policy_map.get(key, {})
        policy_start_str = str(policy.get('返佣开始时间', '')).strip()
        
        if order_time_str and order_time_str != 'nan' and policy_start_str and policy_start_str != 'nan':
            try:
                o_date = pd.to_datetime(order_time_str).date()
                p_date = pd.to_datetime(policy_start_str).date()
                if o_date < p_date:
                    current_remarks = ("，下单早于政策" if current_remarks else "下单早于政策")
                    df_all.at[row.name, '返佣金额'] = 0.0
                    df_all.at[row.name, '是否有返佣'] = '否'
            except Exception:
                pass
        return current_remarks

    df_all['备注'] = df_all.apply(check_date_and_adjust, axis=1)
    return df_all

# --- 网页界面部分 ---

uploaded_ledger = st.file_uploader("1. 上传《分账支付记录.xls》", type=['xls', 'xlsx'])
uploaded_payment = st.file_uploader("2. 上传《代付记录.xls》", type=['xls', 'xlsx'])
uploaded_order = st.file_uploader("3. 上传《订单.xls》", type=['xls', 'xlsx'])
uploaded_detail = st.file_uploader("4. 上传《订单支付明细.xlsx》", type=['xls', 'xlsx'])
uploaded_policy = st.file_uploader("5. 上传《返佣政策详情.xls》", type=['xls', 'xlsx'])

if all([uploaded_ledger, uploaded_payment, uploaded_order, uploaded_detail, uploaded_policy]):
    if st.button('🚀 开始计算', type='primary'):
        with st.spinner('数据正在飞速计算中，请稍候...'):
            try:
                result_df = process_data(uploaded_ledger, uploaded_payment, uploaded_order, uploaded_detail, uploaded_policy)
                
                FINAL_COLUMNS = [
                    '业务订单号', '产品名称', '收款商户', '付款人', '分期金额', '还款期次', '支付时间',
                    '服务费', '逾期费用', '还款方式', '下单时间', '订单状态', '维护商务', '是否有返佣',
                    '返佣比例', '返佣金额', '备注'
                ]
                result_df = result_df[FINAL_COLUMNS]
                
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    result_df.to_excel(writer, index=False, sheet_name='返佣计算结果')
                processed_data = output.getvalue()
                
                st.success("计算完成！")
                st.download_button(
                    label="💾 点击下载计算结果",
                    data=processed_data,
                    file_name="月度回款返佣计算结果_V33.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"计算过程中出现错误：{e}")
else:
    st.info("请先上传全部 5 个文件。")
