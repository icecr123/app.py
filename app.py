import streamlit as st
import pandas as pd
import numpy as np
import re
import io

# --- 页面配置 ---
st.set_page_config(page_title="返佣计算小工具 V33-顺序匹配修正版", layout="centered")
st.title("🧮 月度回款返佣自动计算工具 (V33-顺序匹配修正版)")
st.markdown("""
**V33 核心修复说明：**
1.  **顺序匹配逻辑（防重）**：
    *   **前置处理**：基于【订单支付明细】构建有序的“还款类型队列”。
    *   **动态指针**：在匹配【代付记录】时，根据订单号出现的次数（第1次、第2次...），依次从队列中取出对应的还款类型。
    *   **效果**：彻底解决重复订单号重复匹配首行的问题。
2.  **延期服务费特殊处理**：
    *   **不占期次**：识别到“延期服务费”时，不消耗顺序计数器。
    *   **置空与备注**：直接将【还款期次】置空，并在备注中标记。
3.  **数据清洗与合并**：
    *   **无效删除**：自动删除【线下代付】中服务费和逾期费用均为 0 的无效行。
    *   **费用归集（新增）**：若同一订单出现“仅含逾期费”的行，自动将其合并至该订单的上一行记录中。
""")

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
    
    # 如果期次为空（如延期服务费），不参与返佣计算或按特定逻辑
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
        
    return pd.Series(['否', '0.0000', 0.0])

# --- 主处理流程 ---

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
                '产品名称': row.get('产品名称', ''),
                '下单时间': row.get('下单时间', ''),
                '订单状态': row.get('订单状态', ''),
                '维护商务': row.get('业务员', ''),
                '付款人': row.get('客户姓名', ''),
                '收款商户': row.get('机构简称', ''),
                '分期金额': row.get('分期金额', 0)
            }
            
    # --- V33 核心：构建明细队列 (Detail Queue) ---
    # 结构: { '订单号': ['正常还款', '正常还款', '结清', ...] }
    # 按支付时间排序，确保顺序一致
    detail_queue_map = {}
    
    if not df_detail.empty:
        time_col = '支付时间'
        if '支付时间' in df_detail.columns:
            df_detail['sort_time'] = pd.to_datetime(df_detail['支付时间'], errors='coerce')
            df_detail = df_detail.sort_values(by=['订单编号', 'sort_time'])
            
        grouped_details = df_detail.groupby('订单编号')
        for oid, group in grouped_details:
            clean_oid = clean_order_id(oid)
            type_list = group['还款类型'].fillna('').astype(str).tolist()
            detail_queue_map[clean_oid] = type_list

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
    
    # --- 线上逻辑 ---
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

    # 4. 处理代付记录 (线下 - 顺序匹配版)
    # --- V33 核心：线下代付顺序匹配 ---
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
            
            # --- V33 核心匹配逻辑 ---
            repayment_type = ""
            should_increment_counter = True 
            
            if has_delay_note:
                repayment_type = "" 
                should_increment_counter = False 
            else:
                current_idx = offline_counter.get(oid_clean, 0)
                queue = detail_queue_map.get(oid_clean, [])
                
                if current_idx < len(queue):
                    repayment_type = queue[current_idx]
                else:
                    repayment_type = f"超出明细范围({current_idx+1})"
                
                should_increment_counter = True
            
            if should_increment_counter:
                offline_counter[oid_clean] = offline_counter.get(oid_clean, 0) + 1
            
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
    # V33 新增：数据清洗与合并工序
    # ==========================================
    if not df_all.empty:
        # 1. 删除无效代付记录 (服务费=0 且 逾期费用=0)
        condition_invalid = (
            (df_all['还款方式'] == '线下代付') & 
            (df_all['服务费'] == 0) & 
            (df_all['逾期费用'] == 0)
        )
        df_all = df_all[~condition_invalid].reset_index(drop=True)

        # 2. 合并同单号特殊费用行
        # 逻辑：若某行服务费=0 但 逾期费用>0，尝试将其合并到该订单号的上一行
        rows_to_drop = []
        
        # 倒序遍历，避免索引变动影响后续处理
        for i in range(len(df_all) - 1, -1, -1):
            row = df_all.iloc[i]
            oid = row['业务订单号']
            svc = safe_float(row['服务费'])
            ovd = safe_float(row['逾期费用'])
            
            # 判定条件：服务费为0 且 逾期费用大于0
            if svc == 0 and ovd > 0:
                # 向前查找同一订单号的行
                merged = False
                for j in range(i - 1, -1, -1):
                    prev_row = df_all.iloc[j]
                    if prev_row['业务订单号'] == oid:
                        # 执行合并：将当前行的逾期费加到上一行
                        prev_ovd = safe_float(prev_row['逾期费用'])
                        df_all.at[j, '逾期费用'] = prev_ovd + ovd
                        
                        # 可选：更新备注，标记已合并
                        prev_remark = str(df_all.at[j, '备注'])
                        curr_remark = str(row['备注'])
                        if curr_remark and curr_remark != 'nan':
                            df_all.at[j, '备注'] = prev_remark + "，含合并逾期费"
                        
                        # 标记当前行待删除
                        rows_to_drop.append(i)
                        merged = True
                        break
                
        # 执行删除操作
        if rows_to_drop:
            df_all = df_all.drop(rows_to_drop).reset_index(drop=True)

    # 5. 计算返佣
    if not df_all.empty:
        comm_results = df_all.apply(lambda row: calculate_commission(row, policy_map), axis=1)
        df_all['是否有返佣'] = comm_results[0]
        df_all['返佣比例'] = comm_results[1]
        df_all['返佣金额'] = comm_results[2]
        
        # 6. 补充日期备注与校验
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
                    '业务订单号', '产品名称', '收款商户', '付款人', '分期金额', 
                    '还款期次', '支付时间', '服务费', '逾期费用', '还款方式', 
                    '下单时间', '订单状态', '维护商务', '是否有返佣', '返佣比例', 
                    '返佣金额', '备注'
                ]
                
                # 确保所有列都存在
                for col in FINAL_COLUMNS:
                    if col not in result_df.columns:
                        result_df[col] = ""
                        
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
                    file_name="月度回款返佣计算结果_V33.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
                # 预览
                st.dataframe(result_df.head(10))
                
            except Exception as e:
                st.error(f"计算过程中出现错误：{e}")
                st.exception(e)
else:
    st.info("请先上传全部 5 个文件。")
