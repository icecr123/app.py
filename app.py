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
    """
    强力清洗订单号：
    1. 去除 .0 后缀
    2. 提取纯数字（防止 Excel 变成科学计数法或带括号）
    """
    if pd.isna(order_id): return ''
    s = str(order_id).strip()
    # 使用正则提取第一个连续的数字串，忽略前面的括号或后面的小数点
    match = re.search(r'(\d+)', s)
    if match:
        return match.group(1)
    return s.replace('.0', '')

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
            
        # 【修复报错】确保比较时类型一致
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

# --- 主处理流程 ---
def process_data(ledger_file, payment_file, order_file, detail_file, policy_file):
    # 【关键修复】读取文件时，强制指定订单列为字符串，防止 Excel 自动转为科学计数法
    converters_dict = {'业务订单号': str, '订单号': str} 
    
    try:
        df_ledger = pd.read_excel(ledger_file, converters=converters_dict)
    except: df_ledger = pd.read_excel(ledger_file)
    
    try:
        df_payment_raw = pd.read_excel(payment_file, converters=converters_dict)
    except: df_payment_raw = pd.read_excel(payment_file)
        
    try:
        df_order = pd.read_excel(order_file, converters=converters_dict)
    except: df_order = pd.read_excel(order_file)
        
    try:
        df_detail = pd.read_excel(detail_file, converters=converters_dict)
    except: df_detail = pd.read_excel(detail_file)
        
    try:
        df_policy_raw = pd.read_excel(policy_file)
    except: df_policy_raw = pd.read_excel(policy_file)

    # 1. 统一列名
    rename_map = {
        '订单号': '订单编号', '业务订单号': '订单编号'
    }
    for df in [df_order, df_payment_raw, df_detail, df_ledger]:
        df.rename(columns=rename_map, inplace=True)
        if '订单编号' in df.columns:
            # 再次清洗，确保没有空格
            df['订单编号'] = df['订单编号'].astype(str).str.strip()

    # 2. 预处理明细表期次
    if '订单编号' in df_detail.columns:
        if '下单时间' in df_detail.columns:
            df_detail = df_detail.sort_values(by=['订单编号', '下单时间'])
        df_detail['还款期次'] = df_detail.groupby('订单编号').cumcount() + 1
        
    # 3. 建立映射字典 (Key 全部使用清洗后的纯数字字符串)
    order_map = {}
    for _, row in df_order.iterrows():
        oid = clean_order_id(row.get('订单编号')) 
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

    detail_map = {}
    for _, row in df_detail.iterrows():
        oid = clean_order_id(row.get('订单编号'))
        period = row.get('还款期次', 1) 
        if oid:
            map_key = f"{oid}_{period}"
            detail_map[map_key] = row.get('还款类型', '')

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

    # 4. 处理分账记录 (线上) - 保持不变
    res_list_online = []
    for _, row in df_ledger.iterrows():
        oid = clean_order_id(row.get('订单编号'))
        info = order_map.get(oid, {})
        period_str = str(row.get('还款期次', ''))
        remark_parts = []
        if '延期手续费' in period_str: remark_parts.append("延期服务费")
            
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
        res_list_online.append(new_row)

    # 5. 处理代付记录 (线下) - 【重点修复区域】
    res_list_offline = []
    
    if not df_payment_raw.empty:
        # A. 标记类型
        def classify_row(row):
            note = str(row.get('系统备注', '')).strip()
            amt = safe_float(row.get('清分金额', 0))
            if amt <= 0: return 'ignore'
            if '服务费' in note or '手续费' in note: return 'fee'
            if '罚息' in note or '逾期' in note or '违约金' in note: return 'penalty'
            return 'principal'

        df_payment_raw['_type'] = df_payment_raw.apply(classify_row, axis=1)

        # B. 分组聚合
        grouped = df_payment_raw.groupby(['支付批次号', '订单编号'])
        
        for (batch_id, oid), group in grouped:
            # 【关键修复】在这里立刻清洗 ID，确保后续匹配正确
            oid_clean = clean_order_id(oid)
            info = order_map.get(oid_clean, {}) # 现在能匹配上了
            
            total_service = 0.0
            total_overdue = 0.0
            valid_time = None # 用于抓取任意有效时间
            has_delay_note = False
            
            for _, r in group.iterrows():
                note = str(r.get('系统备注', ''))
                amt = safe_float(r.get('清分金额', 0))
                finish_time = r.get('完成时间', '')
                
                # 抓取时间：只要组内有非空时间，就记录下来
                if pd.notna(finish_time) and str(finish_time).strip() != '' and valid_time is None:
                    valid_time = finish_time
                
                if '服务费' in note and '返服务费' not in note:
                    total_service += amt
                elif '罚息' in note or '逾期' in note:
                    total_overdue += amt
                    
                if '延期服务费' in note: has_delay_note = True
            
            # C. 过滤纯本金 (保留有费用的行)
            if total_service == 0 and total_overdue == 0: continue

            remark_parts = ["延期服务费"] if has_delay_note else []
            
            new_row = {
                '业务订单号': oid_clean, # 存入清洗后的 ID
                '产品名称': info.get('产品名称', ''),
                '收款商户': info.get('收款商户', ''),
                '付款人': info.get('付款人', ''),
                '分期金额': info.get('分期金额', 0),
                '支付时间': valid_time, # 使用抓取到的时间
                '服务费': total_service,
                '逾期费用': total_overdue,
                '还款方式': '线下代付',
                '下单时间': info.get('下单时间', ''),
                '订单状态': info.get('订单状态', ''),
                '维护商务': info.get('维护商务', ''),
                '备注': "，".join(remark_parts),
                '_temp_oid': oid_clean 
            }
            res_list_offline.append(new_row)

    # 6. 合并与排序
    df_all = pd.DataFrame(res_list_online + res_list_offline)
    
    # 针对线下数据进行期次匹配
    if not df_all.empty:
        # 筛选出线下数据进行处理
        mask_offline = df_all['还款方式'] == '线下代付'
        df_offline_part = df_all[mask_offline].copy()
        
        if not df_offline_part.empty:
            df_offline_part = df_offline_part.sort_values(by=['_temp_oid', '支付时间'])
            df_offline_part['还款期次_seq'] = df_offline_part.groupby('_temp_oid').cumcount() + 1
            
            matched_types = []
            for _, row in df_offline_part.iterrows():
                oid = row['_temp_oid']
                seq = row['还款期次_seq']
                key = f"{oid}_{seq}"
                r_type = detail_map.get(key, '未匹配')
                matched_types.append(r_type)
                
            df_offline_part['还款期次'] = matched_types
            
            # 更新回总表
            df_all.loc[mask_offline, '还款期次'] = df_offline_part['还款期次']

    # 7. 计算返佣
    comm_results = df_all.apply(lambda row: calculate_commission(row, policy_map), axis=1)
    df_all['是否有返佣'] = comm_results[0]
    df_all['返佣比例'] = comm_results[1]
    df_all['返佣金额'] = comm_results[2]

    # 8. 日期校验
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
ledger_file = st.file_uploader("1. 上传分账记录表 (线上)", type=["xlsx", "xls"])
payment_file = st.file_uploader("2. 上传代付记录表 (线下)", type=["xlsx", "xls"])
order_file = st.file_uploader("3. 上传订单主表", type=["xlsx", "xls"])
detail_file = st.file_uploader("4. 上传订单支付明细表", type=["xlsx", "xls"])
policy_file = st.file_uploader("5. 上传返佣政策表", type=["xlsx", "xls"])

if st.button("开始计算"):
    if all([ledger_file, payment_file, order_file, detail_file, policy_file]):
        with st.spinner("正在处理数据..."):
            try:
                result_df = process_data(ledger_file, payment_file, order_file, detail_file, policy_file)
                st.success("计算完成！")
                st.dataframe(result_df)
                
                # 导出 Excel
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    result_df.to_excel(writer, index=False, sheet_name='返佣计算结果')
                st.download_button(
                    label="下载结果 Excel",
                    data=output.getvalue(),
                    file_name="返佣计算结果.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"发生错误: {str(e)}")
    else:
        st.warning("请上传所有 5 个文件后再点击开始计算。")
