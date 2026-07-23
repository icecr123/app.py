import streamlit as st
import pandas as pd
import numpy as np
import re
import io

# --- 页面配置 ---
st.set_page_config(page_title="返佣计算小工具", layout="centered")
st.title("🧮 月度回款返佣自动计算工具")
st.markdown("请依次上传以下 5 个文件，工具将自动完成计算并生成结果。")

# --- 核心逻辑函数 ---
def safe_float(val):
    if pd.isna(val): return 0.0
    s = str(val).strip()
    if s in ['无', 'None', 'nan', '']: return 0.0
    try: return float(s)
    except ValueError: return 0.0

def clean_order_id(order_id):
    if pd.isna(order_id): return ''
    s = str(order_id).strip()
    if s.endswith('.0'): s = s[:-2]
    return s

def parse_xy_product(product_name):
    if pd.isna(product_name): return False, 0, 0
    name_str = str(product_name).strip()
    match = re.search(r'(\d+)\+(\d+)', name_str)
    if match: return True, int(match.group(1)), int(match.group(2))
    return False, 0, 0

def count_periods(period_str):
    if pd.isna(period_str): return 1
    numbers = re.findall(r'\d+', str(period_str))
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
    ratio, has_comm, p_num = 0.0, '否', count_periods(period_str)
    if is_xy:
        last_period = 0
        numbers = re.findall(r'\d+', period_str)
        if numbers: last_period = int(numbers[-1])
        raw_ratio = policy.get('X-返佣', 0) if 0 < last_period <= x_val else policy.get('Y-返佣', 0)
    else:
        raw_ratio = policy.get('等额-返佣', 0)
    ratio = safe_float(raw_ratio)
    if ratio > 0: has_comm = '是'
    comm_amount = amount * ratio * p_num if ratio > 0 and amount > 0 else 0.0
    return pd.Series([has_comm, f"{ratio:.4f}", round(comm_amount, 2)])

# --- 主处理流程 ---
def process_data(ledger_file, payment_file, order_file, detail_file, policy_file):
    df_ledger = pd.read_excel(ledger_file)
    df_payment_raw = pd.read_excel(payment_file)
    df_order = pd.read_excel(order_file)
    df_detail = pd.read_excel(detail_file)
    df_policy_raw = pd.read_excel(policy_file)

    # 统一列名
    rename_map = {'订单号': '订单编号', '业务订单号': '订单编号'}
    for df in [df_order, df_payment_raw, df_detail, df_ledger]:
        df.rename(columns={k: v for k, v in rename_map.items() if k in df.columns}, inplace=True)
    for df_temp in [df_order, df_payment_raw, df_detail, df_ledger]:
        if '订单编号' in df_temp.columns:
            df_temp['订单编号'] = df_temp['订单编号'].astype(str).str.strip()

    # 建立订单映射
    order_map = {}
    for _, row in df_order.iterrows():
        oid = clean_order_id(row.get('订单编号'))
        if oid:
            order_map[oid] = {
                '产品名称': row.get('产品名称', ''), '下单时间': row.get('下单时间', ''),
                '订单状态': row.get('订单状态', ''), '维护商务': row.get('业务员', ''),
                '付款人': row.get('客户姓名', ''), '收款商户': row.get('机构简称', ''),
                '分期金额': row.get('分期金额', 0)
            }

    # 建立明细双键映射
    detail_map = {}
    if '订单编号' in df_detail.columns:
        sort_col = '下单时间' if '下单时间' in df_detail.columns else df_detail.columns[0]
        df_detail_sorted = df_detail.sort_values(by=['订单编号', sort_col])
        df_detail_sorted['还款期次'] = df_detail_sorted.groupby('订单编号').cumcount() + 1
        for _, row in df_detail_sorted.iterrows():
            oid = clean_order_id(row.get('订单编号'))
            period = row.get('还款期次', 1)
            if oid: detail_map[f"{oid}_{period}"] = row.get('还款类型', '')

    # 建立政策映射
    policy_map = {}
    for _, row in df_policy_raw.iterrows():
        inst, prod = str(row.get('机构名称', '')).strip(), str(row.get('产品名称', '')).strip()
        if inst and prod:
            policy_map[f"{inst}_{prod}"] = {
                '等额-返佣': row.get('等额-返佣', 0), 'X-返佣': row.get('X-返佣', 0),
                'Y-返佣': row.get('Y-返佣', 0), '返佣开始时间': str(row.get('返佣开始时间', '')).strip()
            }

    # 4. 处理线上分账记录
    ledger_list = []
    for _, row in df_ledger.iterrows():
        oid = clean_order_id(row.get('订单编号'))
        info = order_map.get(oid, {})
        period_str = str(row.get('还款期次', ''))
        remark_parts = ["延期服务费"] if '延期手续费' in period_str else []
        ledger_list.append({
            '业务订单号': oid, '产品名称': info.get('产品名称', row.get('产品名称', '')),
            '收款商户': info.get('收款商户', row.get('收款商户', '')),
            '付款人': info.get('付款人', row.get('付款人', '')),
            '分期金额': row.get('分期金额', 0), '还款期次': period_str,
            '支付时间': row.get('支付时间', ''), '服务费': row.get('服务费', 0),
            '逾期费用': row.get('逾期费', row.get('罚息', 0)), '还款方式': '线上还款',
            '下单时间': info.get('下单时间', ''), '订单状态': info.get('订单状态', ''),
            '维护商务': info.get('维护商务', ''), '备注': "，".join(remark_parts)
        })
    df_ledger_res = pd.DataFrame(ledger_list)

    # 5. 处理线下代付记录 (彻底重写，修复空白和订单号问题)
    df_payment_res = pd.DataFrame()
    if not df_payment_raw.empty and '订单编号' in df_payment_raw.columns:
        # 5.1 前置过滤：只要包含“本金”或“返服务费”，整行丢弃
        df_payment_raw = df_payment_raw[
            ~df_payment_raw['系统备注'].astype(str).str.contains('本金|返服务费', na=False)
        ].copy()

        # 5.2 清洗订单号并标记费用类型
        df_payment_raw['订单编号_clean'] = df_payment_raw['订单编号'].apply(clean_order_id)
        def classify_type(row):
            note = str(row.get('系统备注', ''))
            amt = safe_float(row.get('清分金额', 0))
            if amt <= 0: return 'ignore'
            if '服务费' in note or '手续费' in note: return 'fee'
            elif '罚息' in note or '逾期' in note or '违约金' in note: return 'penalty'
            else: return 'ignore'
        df_payment_raw['_type'] = df_payment_raw.apply(classify_type, axis=1)

        # 5.3 按 支付批次号+订单编号 聚合
        grouped = df_payment_raw.groupby(['支付批次号', '订单编号_clean']).agg(
            支付时间=('完成时间', 'first'),
            服务费=('清分金额', lambda x: x[df_payment_raw.loc[x.index, '_type'] == 'fee'].sum()),
            罚息=('清分金额', lambda x: x[df_payment_raw.loc[x.index, '_type'] == 'penalty'].sum())
        ).reset_index()

        # 5.4 过滤掉服务费和罚息都为0的行
        df_payment_clean = grouped[(grouped['服务费'] > 0) | (grouped['罚息'] > 0)].copy()
        df_payment_clean['服务费'] = df_payment_clean['服务费'].round(2)
        df_payment_clean['罚息'] = df_payment_clean['罚息'].round(2)

        # 5.5 生成还款序号并匹配
        df_payment_clean = df_payment_clean.sort_values(by=['订单编号_clean', '支付时间'])
        df_payment_clean['还款期次_seq'] = df_payment_clean.groupby('订单编号_clean').cumcount() + 1
        
        payment_list = []
        for _, row in df_payment_clean.iterrows():
            oid = row['订单编号_clean']
            info = order_map.get(oid, {})
            period = row['还款期次_seq']
            repayment_type = detail_map.get(f"{oid}_{period}", '')
            payment_list.append({
                '业务订单号': oid, '产品名称': info.get('产品名称', ''),
                '收款商户': info.get('收款商户', ''), '付款人': info.get('付款人', ''),
                '分期金额': info.get('分期金额', 0), '还款期次': repayment_type,
                '支付时间': row['支付时间'], '服务费': row['服务费'],
                '逾期费用': row['罚息'], '还款方式': '线下代付',
                '下单时间': info.get('下单时间', ''), '订单状态': info.get('订单状态', ''),
                '维护商务': info.get('维护商务', ''), '备注': ""
            })
        df_payment_res = pd.DataFrame(payment_list)

    # 6. 安全合并线上和线下数据
    df_all = pd.concat([df_ledger_res, df_payment_res], ignore_index=True)

    # 7. 计算返佣
    comm_results = df_all.apply(lambda row: calculate_commission(row, policy_map), axis=1)
    df_all['是否有返佣'] = comm_results[0]
    df_all['返佣比例'] = comm_results[1]
    df_all['返佣金额'] = comm_results[2]

    # 8. 补充日期备注与校验
    def check_date_and_adjust(row):
        order_time_str = str(row.get('下单时间', '')).strip()
        merchant = str(row.get('收款商户', '')).strip()
        product = str(row.get('产品名称', '')).strip()
        current_remarks = row['备注'] if pd.notna(row['备注']) else ""
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
                    '业务订单号', '产品名称', '收款商户', '付款人', '分期金额', '还款期次', 
                    '支付时间', '服务费', '逾期费用', '还款方式', '下单时间', '订单状态', 
                    '维护商务', '是否有返佣', '返佣比例', '返佣金额', '备注'
                ]
                result_df = result_df[FINAL_COLUMNS]
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    result_df.to_excel(writer, index=False, sheet_name='返佣计算结果')
                processed_data = output.getvalue()
                st.success("计算完成！")
                st.download_button(
                    label="💾 点击下载计算结果", data=processed_data,
                    file_name="月度回款返佣计算结果.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"计算过程中出现错误：{e}")
else:
    st.info("请先上传全部 5 个文件。")
