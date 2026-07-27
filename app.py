import streamlit as st
import pandas as pd
import numpy as np
import re
import io

# --- 页面配置与时间控制器 ---
st.set_page_config(page_title="返佣计算小工具 V34", layout="centered")
st.title("🧮 月度回款返佣自动计算工具 (V34 - 修复版)")

st.markdown("""
<div style='background-color:#e8f5e9; padding:10px; border-radius:5px; border:1px solid #c8e6c9;'>
<b>✅ V34 核心修复：</b><br>
1. <b>修复 KeyError：</b>已更正“订单支付明细”中的时间列为 <code>支付成功时间</code>。<br>
2. <b>双重模糊匹配：</b>采用 <code>业务订单号</code> + <code>支付成功时间(±60s)</code> 联合匹配，解决分钟/秒数不一致问题。<br>
3. <b>精准定位：</b>匹配成功后自动回填“还款期次”与“还款类型”。
</div>
""", unsafe_allow_html=True)

# 时间控制器
col_year, col_month = st.columns(2)
with col_year:
    selected_year = st.selectbox("选择年份", [2023, 2024, 2025, 2026], index=2)
with col_month:
    selected_month = st.selectbox("选择月份", list(range(1, 13)), index=6) # 默认7月

# --- 核心逻辑函数 ---
def safe_float(val):
    """安全转换浮点数"""
    if pd.isna(val):
        return 0.0
    s = str(val).strip()
    if s in ['无', 'None', 'nan', '']:
        return 0.0
    try:
        return float(s)
    except ValueError:
        return 0.0

def clean_order_id(order_id):
    """清洗订单号，去除 .0 后缀"""
    if pd.isna(order_id):
        return ''
    s = str(order_id).strip()
    if s.endswith('.0'):
        s = s[:-2]
    return s

def parse_xy_product(product_name):
    """解析产品名称中的 x+y 格式"""
    if pd.isna(product_name):
        return False, 0, 0
    name_str = str(product_name).strip()
    match = re.search(r'(\d+)\+(\d+)', name_str)
    if match:
        return True, int(match.group(1)), int(match.group(2))
    return False, 0, 0

def count_periods(period_str):
    """统计还款期次数量"""
    if pd.isna(period_str):
        return 1
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

    if not period_str:
        return pd.Series(['否', '0.0000', 0.0])

    p_num = count_periods(period_str)

    if is_xy:
        last_period = 0
        numbers = re.findall(r'\d+', period_str)
        if numbers:
            last_period = int(numbers[-1])
        if 0 < last_period <= x_val:
            raw_ratio = policy.get('X-返佣', 0)
        else:
            raw_ratio = policy.get('Y-返佣', 0)
    else:
        raw_ratio = policy.get('等额-返佣', 0)

    ratio = safe_float(raw_ratio)
    if ratio > 0:
        has_comm = '是'
    
    comm_amount = 0.0
    if ratio > 0 and amount > 0:
        comm_amount = amount * ratio * p_num
        
    return pd.Series([has_comm, f"{ratio:.4f}", round(comm_amount, 2)])

# --- 主处理流程 ---
def process_data(ledger_file, payment_file, order_file, detail_file, policy_file, year, month):
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
                '产品名称': row.get('产品名称', ''),
                '下单时间': row.get('下单时间', ''),
                '订单状态': row.get('订单状态', ''),
                '维护商务': row.get('业务员', ''),
                '付款人': row.get('客户姓名', ''),
                '收款商户': row.get('机构简称', ''),
                '分期金额': row.get('分期金额', 0)
            }

    # --- V34 核心：构建明细队列 (Detail Queue) ---
    detail_queue_map = {}
    detail_time_map = {} # 新增：用于存储每个订单的支付时间列表，用于模糊匹配
    
    if not df_detail.empty:
        # 【修复点1】：使用正确的列名 '支付成功时间'
        time_col = '支付成功时间'
        order_col = '订单编号'
        
        if time_col in df_detail.columns and order_col in df_detail.columns:
            # 转换时间格式
            df_detail['sort_time'] = pd.to_datetime(df_detail[time_col], errors='coerce')
            df_detail = df_detail.sort_values(by=[order_col, 'sort_time'])
            
            grouped_details = df_detail.groupby(order_col)
            for oid, group in grouped_details:
                clean_oid = clean_order_id(oid)
                # 存储还款类型列表
                type_list = group['还款类型'].fillna('').astype(str).tolist()
                detail_queue_map[clean_oid] = type_list
                # 存储支付时间列表
                time_list = group['sort_time'].tolist()
                detail_time_map[clean_oid] = time_list
        else:
            st.error(f"错误：在《订单支付明细》中找不到 '{order_col}' 或 '{time_col}' 列，请检查表头。")
            return pd.DataFrame()

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

    # 3. 处理分账记录 (线上)
    res_list = []
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

    # 4. 处理代付记录 (线下 - 模糊匹配版)
    if not df_payment_raw.empty:
        offline_counter = {}
        group_cols = ['业务订单号', '支付批次号']
        if not all(col in df_payment_raw.columns for col in group_cols):
            st.error("代付记录表中缺少 '业务订单号' 或 '支付批次号' 列，请检查表头。")
            return pd.DataFrame()
            
        grouped = df_payment_raw.groupby(group_cols)
        for (oid, batch_id), group in grouped:
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
                if '延期服务费' in note:
                    has_delay_note = True

            final_pay_time = service_time if service_time else group.iloc[0].get('完成时间', '')
            remark_parts = ["延期服务费"] if has_delay_note else []
            repayment_type = ""
            should_increment_counter = True

            if has_delay_note:
                repayment_type = ""
                should_increment_counter = False
            else:
                # --- V34 核心：模糊匹配逻辑 ---
                queue = detail_queue_map.get(oid_clean, [])
                time_list = detail_time_map.get(oid_clean, [])
                
                if queue:
                    # 尝试进行时间模糊匹配
                    if final_pay_time:
                        pay_time_dt = pd.to_datetime(final_pay_time, errors='coerce')
                        matched_idx = -1
                        min_diff = pd.Timedelta(seconds=60) # 默认容差60秒
                        
                        for i, ref_time in enumerate(time_list):
                            if pd.isna(ref_time) or pd.isna(pay_time_dt):
                                continue
                            diff = abs(ref_time - pay_time_dt)
                            if diff <= min_diff:
                                min_diff = diff
                                matched_idx = i
                                break # 找到第一个匹配的就停止，或者可以继续找最近的
                                
                        if matched_idx != -1:
                            repayment_type = queue[matched_idx]
                            offline_counter[oid_clean] = matched_idx + 1 # 更新计数器，避免重复匹配
                        else:
                            # 时间没匹配上，按顺序取
                            current_idx = offline_counter.get(oid_clean, 0)
                            if current_idx < len(queue):
                                repayment_type = queue[current_idx]
                                offline_counter[oid_clean] = current_idx + 1
                            else:
                                repayment_type = f"超出明细范围({current_idx+1})"
                    else:
                        # 没有支付时间，按顺序取
                        current_idx = offline_counter.get(oid_clean, 0)
                        if current_idx < len(queue):
                            repayment_type = queue[current_idx]
                            offline_counter[oid_clean] = current_idx + 1
                        else:
                            repayment_type = f"超出明细范围({current_idx+1})"
                else:
                    repayment_type = "无明细数据"

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
                '备注': "，".join(remark_parts)
            }
            res_list.append(new_row)

    df_all = pd.DataFrame(res_list)

    # ==========================================
    # V34 数据清洗与合并工序
    # ==========================================
    if not df_all.empty:
        # 1. 删除无效代付记录
        condition_invalid = (
            (df_all['还款方式'] == '线下代付') & 
            (df_all['服务费'] == 0) & 
            (df_all['逾期费用'] == 0)
        )
        df_all = df_all[~condition_invalid].reset_index(drop=True)

        # 2. 合并同单号特殊费用行
        rows_to_drop = []
        for i in range(len(df_all) - 1, -1, -1):
            row = df_all.iloc[i]
            oid = row['业务订单号']
            svc = safe_float(row['服务费'])
            ovd = safe_float(row['逾期费用'])
            if svc == 0 and ovd > 0:
                for j in range(i - 1, -1, -1):
                    prev_row = df_all.iloc[j]
                    if prev_row['业务订单号'] == oid:
                        prev_ovd = safe_float(prev_row['逾期费用'])
                        df_all.at[j, '逾期费用'] = prev_ovd + ovd
                        prev_remark = str(df_all.at[j, '备注'])
                        curr_remark = str(row['备注'])
                        if curr_remark and curr_remark != 'nan':
                            df_all.at[j, '备注'] = prev_remark + "，含合并逾期费"
                        rows_to_drop.append(i)
                        break
        if rows_to_drop:
            df_all = df_all.drop(rows_to_drop).reset_index(drop=True)

    # ==========================================
    # 5. 时间过滤 (按选择的年月过滤支付时间)
    # ==========================================
    if not df_all.empty:
        # 将支付时间转换为 datetime 格式
        df_all['支付时间_dt'] = pd.to_datetime(df_all['支付时间'], errors='coerce')
        # 提取年月，过滤数据
        df_all['year_month'] = df_all['支付时间_dt'].dt.to_period('M')
        target_period = pd.Period(f"{year}-{month:02d}", freq='M')
        # 仅保留目标月份的数据
        df_all = df_all[df_all['year_month'] == target_period].reset_index(drop=True)
        # 删除辅助列
        df_all.drop(columns=['支付时间_dt', 'year_month'], inplace=True)

        # 6. 计算返佣
        comm_results = df_all.apply(lambda row: calculate_commission(row, policy_map), axis=1)
        df_all['是否有返佣'] = comm_results[0]
        df_all['返佣比例'] = comm_results[1]
        df_all['返佣金额'] = comm_results[2]

        # 7. 补充日期备注与校验
        for idx, row in df_all.iterrows():
            order_time_str = str(row.get('下单时间', '')).strip()
            merchant = str(row.get('收款商户', '')).strip()
            product = str(row.get('产品名称', '')).strip()
            key = f"{merchant}_{product}"
            policy = policy_map.get(key, {})
            policy_start_str = str(policy.get('返佣开始时间', '')).strip()
            
            if order_time_str and order_time_str != 'nan' and policy_start_str and policy_start_str != 'nan':
                try:
                    o_date = pd.to_datetime(order_time_str).date()
                    p_date = pd.to_datetime(policy_start_str).date()
                    if o_date < p_date:
                        df_all.at[idx, '备注'] = str(df_all.at[idx, '备注']) + "下单早于政策"
                        df_all.at[idx, '返佣金额'] = 0.0
                        df_all.at[idx, '是否有返佣'] = '否'
                except Exception:
                    pass
        return df_all
    return pd.DataFrame()

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
                result_df = process_data(
                    uploaded_ledger,
                    uploaded_payment,
                    uploaded_order,
                    uploaded_detail,
                    uploaded_policy,
                    selected_year,
                    selected_month
                )
                FINAL_COLUMNS = [
                    '业务订单号', '产品名称', '收款商户', '付款人', '分期金额', '还款期次',
                    '支付时间', '服务费', '逾期费用', '还款方式', '下单时间', '订单状态',
                    '维护商务', '是否有返佣', '返佣比例', '返佣金额', '备注'
                ]
                for col in FINAL_COLUMNS:
                    if col not in result_df.columns:
                        result_df[col] = ""
                result_df = result_df[FINAL_COLUMNS]
                
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    result_df.to_excel(writer, index=False, sheet_name='返佣计算结果')
                processed_data = output.getvalue()
                
                st.success(f"计算完成！当前输出范围：{selected_year}年{selected_month}月")
                st.download_button(
                    label="💾 点击下载计算结果",
                    data=processed_data,
                    file_name=f"月度回款返佣计算结果_{selected_year}{selected_month:02d}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                st.dataframe(result_df.head(10))
            except Exception as e:
                st.error(f"计算过程中出现错误：{e}")
                st.exception(e)
else:
    st.info("请先上传全部 5 个文件。")
