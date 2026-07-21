import streamlit as st导入 streamlit 作为 st导入 streamlit 作为 st导入 streamlit 作为 st导入 streamlit 作为 st
import pandas as pd导入 pandas 作为 pd导入 pandas 作为 pd导入 pandas 作为 pd
import numpy as np导入 numpy 作为 np
import re导入 re重新导入 re
import io导入 io

# --- 页面配置 ---
st.set_page_config(page_title="返佣计算小工具", layout="centered")
st.title("🧮 月度回款返佣自动计算工具")
st.markdown("请依次上传以下 5 个文件，工具将自动完成计算并生成结果。")

# --- 你的核心逻辑函数 (保持不变) ---
def safe_float(val):
    if pd.isna(val): return 0.0
    s = str(val).strip()
    if s in ['无', 'None', 'nan', '']: return 0.0
    try: return float(s)尝试: 返回 浮点数(s)尝试: 返回 浮点数(s)尝试: 返回 浮点数(s)
      except ValueError: return 0.0

def clean_order_id(order_id):定义 清理订单ID(订单ID):
    if pd.isna(order_id): return ''如果 pd.isna(order_id): 返回 ''
    s = str(order_id).strip()
    if s.endswith('.0'): s = s[:-2]如果s.以‘.0’结尾，则 s = s[:-2]
    return s

def parse_xy_product(product_name):定义 parse_xy_product(product_name):
    if pd.isna(product_name): return False, 0, 0
    name_str = str(product_name).strip()
    match = re.search(r'(\d+)\+(\d+)', name_str)
    if match: return True, int(match.group(1)), int(match.group(2))    如果匹配，则返回 True，以及 match.group(1) 和 match.group(2) 的整数值
    return False, 0, 0    返回 False, 0, 0

def count_periods(period_str):定义 count_periods(period_str):
    if pd.isna(period_str): return 1    如果 pd.isna(period_str)，则返回 1
    p_str = str(period_str)
    numbers = re.findall(r'\d+', p_str)
    return max(len(numbers), 1)

def calculate_commission(row, policy_map):
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
        if 0 < last_period <= x_val: raw_ratio = policy.get('X-返佣', 0)
        else: raw_ratio = policy.get('Y-返佣', 0)
    else:
        raw_ratio = policy.get('等额-返佣', 0)

    ratio = safe_float(raw_ratio)
    if ratio > 0: has_comm = '是'

    comm_amount = 0.0
    if ratio > 0 and amount > 0:
        comm_amount = amount * ratio * p_num
    return pd.Series([has_comm, f"{ratio:.4f}", round(comm_amount, 2)])

# --- 主处理函数 (稍作修改，接收文件对象) ---
def process_data(ledger_file, payment_file, order_file, detail_file, policy_file):
    # 1. 读取所有文件
    df_ledger = pd.read_excel(ledger_file, dtype=str)
    df_payment_raw = pd.read_excel(payment_file, dtype=str)
    df_order = pd.read_excel(order_file, dtype=str)
    df_detail = pd.read_excel(detail_file, dtype=str)
    df_policy_raw = pd.read_excel(policy_file, dtype=str)

    # 2. 预处理：建立映射字典
    order_map = {}
    for _, row in df_order.iterrows():
        oid = clean_order_id(row.get('订单号'))
        if oid:
            order_map[oid] = {
                '产品名称': row.get('产品名称', ''), '下单时间': row.get('下单时间', ''),
                '订单状态': row.get('订单状态', ''), '维护商务': row.get('业务员', ''),
                '付款人': row.get('客户姓名', ''), '收款商户': row.get('机构简称', ''),
                '分期金额': row.get('分期金额', 0)
            }

    detail_map = {}
    for _, row in df_detail.iterrows():
        oid = clean_order_id(row.get('订单编号'))
        if oid and oid not in detail_map:
            detail_map[oid] = row.get('还款类型', '')

    policy_map = {}
    for _, row in df_policy_raw.iterrows():
        inst = str(row.get('机构名称', '')).strip()
        prod = str(row.get('产品名称', '')).strip()
        if inst and prod:
            key = f"{inst}_{prod}"
            policy_map[key] = {
                '等额-返佣': row.get('等额-返佣', 0), 'X-返佣': row.get('X-返佣', 0),
                'Y-返佣': row.get('Y-返佣', 0), '返佣开始时间': str(row.get('返佣开始时间', '')).strip()
            }

    # 3. 处理分账记录 (线上)
    res_list = []
    for _, row in df_ledger.iterrows():
        oid = clean_order_id(row.get('业务订单号'))
        info = order_map.get(oid, {})
        period_str = str(row.get('还款期次', ''))
        remark_parts = []
        if '延期手续费' in period_str: remark_parts.append("延期服务费")
        new_row = {
            '业务订单号': oid, '产品名称': info.get('产品名称', row.get('产品名称', '')),
            '收款商户': info.get('收款商户', row.get('收款商户', '')),
            '付款人': info.get('付款人', row.get('付款人', '')),
            '分期金额': row.get('分期金额', 0), '还款期次': period_str,
            '支付时间': row.get('支付时间', ''), '服务费': row.get('服务费', 0),
            '逾期费用': row.get('逾期费', row.get('罚息', 0)), '还款方式': '线上还款',
            '下单时间': info.get('下单时间', ''), '订单状态': info.get('订单状态', ''),
            '维护商务': info.get('维护商务', ''), '备注': "，".join(remark_parts)
        }
        res_list.append(new_row)

    # 4. 处理代付记录 (线下 - 聚合版)
    if not df_payment_raw.empty:
        grouped = df_payment_raw.groupby(['支付批次号', '业务订单号'])
        for (batch_id, oid), group in grouped:
            oid_clean = clean_order_id(oid)
            info = order_map.get(oid_clean, {})
            total_service, total_overdue, service_time, has_delay_note = 0.0, 0.0, None, False
            for _, r in group.iterrows():
                note = str(r.get('系统备注', ''))
                amt = safe_float(r.get('清分金额', 0))
                finish_time = r.get('完成时间', '')
                if '服务费' in note and '返服务费' not in note:
                    total_service += amt
                    if pd.notna(finish_time) and str(finish_time).strip() != '':
                        service_time = finish_time
                elif '罚息' in note or '逾期' in note:
                    total_overdue += amt
                if '延期服务费' in note: has_delay_note = True

            final_pay_time = service_time if service_time else group.iloc[0].get('完成时间', '')
            remark_parts = ["延期服务费"] if has_delay_note else []
            new_row = {
                '业务订单号': oid_clean, '产品名称': info.get('产品名称', ''),
                '收款商户': info.get('收款商户', ''), '付款人': info.get('付款人', ''),
                '分期金额': info.get('分期金额', 0), '还款期次': detail_map.get(oid_clean, ''),
                '支付时间': final_pay_time, '服务费': total_service, '逾期费用': total_overdue,
                '还款方式': '线下代付', '下单时间': info.get('下单时间', ''),
                '订单状态': info.get('订单状态', ''), '维护商务': info.get('维护商务', ''),
                '备注': "，".join(remark_parts)
            }
            res_list.append(new_row)

    df_all = pd.DataFrame(res_list)

    # 5. 计算返佣
    comm_results = df_all.apply(lambda row: calculate_commission(row, policy_map), axis=1)
    df_all['是否有返佣'] = comm_results[0]
    df_all['返佣比例'] = comm_results[1]
    df_all['返佣金额'] = comm_results[2]

    # 6. 补充日期备注
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
            except Exception: pass
        return current_remarks
    df_all['备注'] = df_all.apply(check_date_and_adjust, axis=1)

    return df_all

# --- 网页界面部分 ---
# 创建文件上传器
uploaded_ledger = st.file_uploader("1. 上传《分账支付记录.xls》", type=['xls', 'xlsx'])
uploaded_payment = st.file_uploader("2. 上传《代付记录.xls》", type=['xls', 'xlsx'])
uploaded_order = st.file_uploader("3. 上传《订单.xls》", type=['xls', 'xlsx'])
uploaded_detail = st.file_uploader("4. 上传《订单支付明细.xlsx》", type=['xls', 'xlsx'])
uploaded_policy = st.file_uploader("5. 上传《返佣政策详情.xls》", type=['xls', 'xlsx'])

# 当所有文件都上传后，显示计算按钮
if all([uploaded_ledger, uploaded_payment, uploaded_order, uploaded_detail, uploaded_policy]):
    if st.button('🚀 开始计算', type='primary'):
        with st.spinner('数据正在飞速计算中，请稍候...'):
            try:
                # 调用主函数处理数据
                result_df = process_data(uploaded_ledger, uploaded_payment, uploaded_order, uploaded_detail, uploaded_policy)
                
                # 定义最终输出的标准列头
                FINAL_COLUMNS = [
                    '业务订单号', '产品名称', '收款商户', '付款人', '分期金额', '还款期次',
                    '支付时间', '服务费', '逾期费用', '还款方式', '下单时间', '订单状态',
                    '维护商务', '是否有返佣', '返佣比例', '返佣金额', '备注'
                ]
                result_df = result_df[FINAL_COLUMNS]

                # 将结果转换为 Excel 文件并放入内存
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    result_df.to_excel(writer, index=False, sheet_name='返佣计算结果')
                processed_data = output.getvalue()

                # 提供下载按钮
                st.success("计算完成！")
                st.download_button(
                    label="💾 点击下载计算结果",
                    data=processed_data,
                    file_name="月度回款返佣计算结果.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"计算过程中出现错误：{e}")
else:
    st.info("请先上传全部 5 个文件。")
