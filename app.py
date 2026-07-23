import streamlit as st
import pandas as pd
import numpy as np
import re
import io

# --- 页面配置 ---
st.set_page_config(page_title="返佣计算小工具", layout="centered")
st.title("🧮 月度回款返佣自动计算工具")
st.markdown("请依次上传以下 5 个文件，工具将自动完成计算并生成结果。")

# --- 核心辅助函数 ---
def safe_float(val):
    """安全转换金额"""
    if pd.isna(val): return 0.0
    s = str(val).strip()
    if s in ['无', 'None', 'nan', '']: return 0.0
    try: return float(s)
    except ValueError: return 0.0

def clean_order_id(order_id):
    """【关键修复】清洗订单号，解决科学计数法和 .0 后缀"""
    if pd.isna(order_id): return ''
    s = str(order_id).strip()
    # 如果是科学计数法或带小数点的数字，转为整数再转字符串
    try:
        # 尝试转为浮点数再取整，消除 .0
        num_val = float(s)
        # 检查是否是极大的数字（防止把真正的短数字误伤，虽然这里主要是针对长ID）
        if abs(num_val) > 1e10: 
            s = str(int(num_val))
        else:
            # 普通数字直接去掉 .0
            if '.' in s:
                s = s.rstrip('0').rstrip('.')
    except ValueError:
        pass
    
    # 最后兜底：如果还有 .0 手动去除
    if s.endswith('.0'):
        s = s[:-2]
    return s

def parse_xy_product(product_name):
    """解析 x+y 产品格式"""
    if pd.isna(product_name): return False, 0, 0
    name_str = str(product_name).strip()
    match = re.search(r'(\d+)\+(\d+)', name_str)
    if match: return True, int(match.group(1)), int(match.group(2))
    return False, 0, 0

def count_periods(period_str):
    """统计期次数量"""
    if pd.isna(period_str): return 1
    p_str = str(period_str)
    numbers = re.findall(r'\d+', p_str)
    return max(len(numbers), 1)

def calculate_commission(row, policy_map):
    """计算单笔返佣"""
    merchant = str(row.get('收款商户', '')).strip()
    product = str(row.get('产品名称', '')).strip()
    period_str = str(row.get('还款期次', '')).strip()
    amount = safe_float(row.get('分期金额', 0))
    
    key = f"{merchant}_{product}"
    policy = policy_map.get(key, {})
    
    if not policy:
        return pd.Series(['否', '0.0000', 0.0])
        
    is_xy, x_val, y_val = parse_xy_product(product)
    ratio = 0.0
    has_comm = '否'
    p_num = count_periods(period_str)
    
    if is_xy:
        last_period = 0
        numbers = re.findall(r'\d+', period_str)
        if numbers:
            try: last_period = int(numbers[-1])
            except: pass
            
        if 0 < last_period <= x_val:
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

# --- 主处理流程 ---
def process_data(ledger_file, payment_file, order_file, detail_file, policy_file):
    # 【关键修复】读取文件时，强制指定订单相关列为字符串，防止科学计数法
    # 假设各表的订单列名可能不同，这里统一处理
    
    # 1. 读取数据
    df_ledger = pd.read_excel(ledger_file, dtype={'业务订单号': str, '订单编号': str})
    df_payment_raw = pd.read_excel(payment_file, dtype={'业务订单号': str, '订单编号': str})
    df_order = pd.read_excel(order_file, dtype={'订单号': str, '订单编号': str})
    df_detail = pd.read_excel(detail_file, dtype={'业务订单号': str, '订单编号': str})
    df_policy_raw = pd.read_excel(policy_file)

    # 2. 预处理：统一列名 & 清洗
    # ---------------------------------------------------------
    rename_map = {
        '订单号': '订单编号',
        '业务订单号': '订单编号'
    }
    
    for df in [df_ledger, df_payment_raw, df_order, df_detail]:
        # 统一列名
        df.rename(columns={k: v for k, v in rename_map.items() if k in df.columns}, inplace=True)
        # 统一清洗订单号 (再次确保去除空格和潜在的格式问题)
        if '订单编号' in df.columns:
            df['订单编号'] = df['订单编号'].apply(clean_order_id).str.strip()

    # 为明细表生成期次序号 (用于后续匹配)
    if '订单编号' in df_detail.columns:
        if '下单时间' in df_detail.columns:
            df_detail.sort_values(by=['订单编号', '下单时间'], inplace=True)
        df_detail['还款期次_seq'] = df_detail.groupby('订单编号').cumcount() + 1
        
    # ---------------------------------------------------------
    # 3. 建立映射字典 (Key 必须是清洗后的纯数字字符串)
    # ---------------------------------------------------------
    order_map = {}
    for _, row in df_order.iterrows():
        oid = row.get('订单编号')
        if oid and oid != 'nan':
            order_map[oid] = {
                '产品名称': row.get('产品名称', ''),
                '下单时间': row.get('下单时间', ''),
                '订单状态': row.get('订单状态', ''),
                '维护商务': row.get('业务员', ''),
                '付款人': row.get('客户姓名', ''),
                '收款商户': row.get('机构简称', ''),
                '分期金额': row.get('分期金额', 0)
            }

    detail_map = {}
    for _, row in df_detail.iterrows():
        oid = row.get('订单编号')
        seq = row.get('还款期次_seq', 1)
        if oid and oid != 'nan':
            detail_map[f"{oid}_{seq}"] = row.get('还款类型', '')

    policy_map = {}
    for _, row in df_policy_raw.iterrows():
        inst = str(row.get('机构名称', '')).strip()
        prod = str(row.get('产品名称', '')).strip()
        if inst and prod:
            policy_map[f"{inst}_{prod}"] = {
                '等额-返佣': row.get('等额-返佣', 0),
                'X-返佣': row.get('X-返佣', 0),
                'Y-返佣': row.get('Y-返佣', 0),
                '返佣开始时间': str(row.get('返佣开始时间', '')).strip()
            }

    # 4. 处理线上分账记录
    res_list_online = []
    for _, row in df_ledger.iterrows():
        oid = row.get('订单编号')
        info = order_map.get(oid, {})
        
        new_row = {
            '业务订单号': oid,
            '产品名称': info.get('产品名称', ''),
            '收款商户': info.get('收款商户', ''),
            '付款人': info.get('付款人', ''),
            '分期金额': info.get('分期金额', row.get('分期金额', 0)),
            '还款期次': str(row.get('还款期次', '')),
            '支付时间': row.get('支付时间', ''),
            '服务费': row.get('服务费', 0),
            '逾期费用': row.get('逾期费', row.get('罚息', 0)),
            '还款方式': '线上还款',
            '下单时间': info.get('下单时间', ''),
            '订单状态': info.get('订单状态', ''),
            '维护商务': info.get('维护商务', ''),
            '备注': "延期服务费" if '延期' in str(row.get('还款期次', '')) else ""
        }
        res_list_online.append(new_row)

    # 5. 处理线下代付记录 (彻底重构版)
    res_list_offline = []
    
    if not df_payment_raw.empty:
        # A. 过滤脏数据：只要备注含“本金”或“返服务费”，直接丢弃
        df_clean = df_payment_raw.copy()
        note_col = '系统备注' if '系统备注' in df_clean.columns else '备注'
        if note_col in df_clean.columns:
            mask_bad = df_clean[note_col].astype(str).str.contains('本金|返服务费', na=False)
            df_clean = df_clean[~mask_bad]

        # B. 分组聚合
        # 注意：此时 df_clean['订单编号'] 已经是干净的字符串了
        grouped = df_clean.groupby(['订单编号']) 
        
        # 为了计算期次，我们需要知道每个订单出现了几次有效的代付
        # 这里先收集所有聚合后的数据，再排序算期次
        temp_agg_list = []
        
        for oid, group in grouped:
            total_service = 0.0
            total_overdue = 0.0
            pay_time = None
            
            # 遍历组内行寻找时间和累加金额
            for _, r in group.iterrows():
                note = str(r.get(note_col, ''))
                amt = safe_float(r.get('清分金额', 0))
                
                if '服务费' in note or '手续费' in note:
                    total_service += amt
                if '罚息' in note or '逾期' in note or '违约金' in note:
                    total_overdue += amt
                    
                # 抓取时间：只要有时间就更新，通常最后一行或服务费行的时间较准
                t = r.get('完成时间', r.get('支付时间', ''))
                if pd.notna(t) and str(t).strip() != '':
                    pay_time = t
            
            # 只有当有钱（服务费或罚息）时才保留
            if total_service > 0 or total_overdue > 0:
                temp_agg_list.append({
                    '_oid': oid,
                    '服务费': total_service,
                    '逾期费用': total_overdue,
                    '支付时间': pay_time
                })

        # C. 组装线下结果并匹配期次
        if temp_agg_list:
            df_pay_temp = pd.DataFrame(temp_agg_list)
            # 按订单号排序，保证期次计算顺序一致
            df_pay_temp.sort_values(by='_oid', inplace=True)
            
            # 计算这是该订单的第几次代付 (1, 2, 3...)
            df_pay_temp['pay_seq'] = df_pay_temp.groupby('_oid').cumcount() + 1
            
            for _, row in df_pay_temp.iterrows():
                oid = row['_oid']
                seq = row['pay_seq']
                info = order_map.get(oid, {})
                
                # 匹配还款类型 (例如：提前结清、正常还款)
                repay_type = detail_map.get(f"{oid}_{seq}", "")
                
                new_row = {
                    '业务订单号': oid,
                    '产品名称': info.get('产品名称', ''),
                    '收款商户': info.get('收款商户', ''),
                    '付款人': info.get('付款人', ''),
                    '分期金额': info.get('分期金额', 0),
                    '还款期次': repay_type, # 填入匹配到的类型
                    '支付时间': row['支付时间'],
                    '服务费': row['服务费'],
                    '逾期费用': row['逾期费用'],
                    '还款方式': '线下代付',
                    '下单时间': info.get('下单时间', ''),
                    '订单状态': info.get('订单状态', ''),
                    '维护商务': info.get('维护商务', ''),
                    '备注': ""
                }
                res_list_offline.append(new_row)

    # 6. 合并数据
    df_all = pd.DataFrame(res_list_online + res_list_offline)
    
    # 7. 计算返佣
    if not df_all.empty:
        comm_results = df_all.apply(lambda row: calculate_commission(row, policy_map), axis=1)
        df_all['是否有返佣'] = comm_results[0]
        df_all['返佣比例'] = comm_results[1]
        df_all['返佣金额'] = comm_results[2]
        
        # 8. 政策时间校验
        def check_date(row):
            o_time = str(row.get('下单时间', ''))
            m = str(row.get('收款商户', ''))
            p = str(row.get('产品名称', ''))
            key = f"{m}_{p}"
            pol = policy_map.get(key, {})
            p_start = str(pol.get('返佣开始时间', ''))
            
            if o_time != 'nan' and p_start != 'nan' and o_time and p_start:
                try:
                    if pd.to_datetime(o_time) < pd.to_datetime(p_start):
                        return "下单早于政策", 0.0, '否'
                except: pass
            return row.get('备注', ''), row.get('返佣金额', 0), row.get('是否有返佣', '否')
            
        adjustments = df_all.apply(check_date, axis=1)
        df_all['备注'] = adjustments.apply(lambda x: x[0])
        df_all['返佣金额'] = adjustments.apply(lambda x: x[1])
        df_all['是否有返佣'] = adjustments.apply(lambda x: x[2])

    return df_all

# --- Streamlit 界面 ---
st.sidebar.header("文件上传区")
f_ledger = st.sidebar.file_uploader("1. 分账记录 (线上)", type=['xlsx', 'xls'])
f_payment = st.sidebar.file_uploader("2. 代付记录 (线下)", type=['xlsx', 'xls'])
f_order = st.sidebar.file_uploader("3. 订单主表", type=['xlsx', 'xls'])
f_detail = st.sidebar.file_uploader("4. 订单支付明细", type=['xlsx', 'xls'])
f_policy = st.sidebar.file_uploader("5. 返佣政策表", type=['xlsx', 'xls'])

if st.sidebar.button("开始计算"):
    if all([f_ledger, f_payment, f_order, f_detail, f_policy]):
        with st.spinner("正在处理数据，请稍候..."):
            try:
                result_df = process_data(f_ledger, f_payment, f_order, f_detail, f_policy)
                st.success("计算完成！")
                st.dataframe(result_df)
                
                # 导出 Excel
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    result_df.to_excel(writer, index=False, sheet_name='返佣计算结果')
                st.download_button(
                    label="下载结果 Excel",
                    data=output.getvalue(),
                    file_name="返佣计算结果.xlsx",
                    mime="application/vnd.ms-excel"
                )
            except Exception as e:
                st.error(f"发生错误: {str(e)}")
                st.exception(e)
    else:
        st.warning("请上传所有 5 个文件后再点击计算。")
