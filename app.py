import streamlit as st
import pandas as pd
import numpy as np
import re
import io

# ================= 页面配置 =================
st.set_page_config(page_title="返佣计算小工具", layout="centered")
st.title("🧮 月度回款返佣自动计算工具 (V3-终极修复版)")
st.markdown("请依次上传以下 5 个文件，工具将自动完成计算并生成结果。")

# ================= 核心辅助函数 =================
def safe_float(val):
    """安全转换金额"""
    if pd.isna(val): return 0.0
    s = str(val).strip()
    if s in ['无', 'None', 'nan', '']: return 0.0
    try: return float(s)
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

# ================= 主处理流程 =================
def process_data(ledger_file, payment_file, order_file, detail_file, policy_file):
    # 1. 读取文件 (强制指定订单相关列为字符串)
    df_ledger = pd.read_excel(ledger_file, dtype={'业务订单号': str, '订单编号': str})
    df_payment_raw = pd.read_excel(payment_file, dtype={'业务订单号': str, '订单编号': str})
    df_order = pd.read_excel(order_file, dtype={'订单号': str, '订单编号': str})
    df_detail = pd.read_excel(detail_file, dtype={'业务订单号': str, '订单编号': str})
    df_policy_raw = pd.read_excel(policy_file)

    # 2. 统一列名 & 清洗 ID
    col_map = {'订单号': '订单编号', '业务订单号': '订单编号'}
    for df in [df_ledger, df_payment_raw, df_order, df_detail]:
        df.rename(columns=col_map, inplace=True)
        if '订单编号' in df.columns:
            df['订单编号'] = df['订单编号'].apply(clean_order_id)

    # 3. 建立映射字典
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

    # 构建明细表映射 (用于匹配还款类型)
    detail_map = {}
    if '订单编号' in df_detail.columns:
        # 关键：明细表也要按时间排序，确保期次索引准确
        if '支付时间' in df_detail.columns or '完成时间' in df_detail.columns:
            time_col = '支付时间' if '支付时间' in df_detail.columns else '完成时间'
            df_detail = df_detail.sort_values(by=['订单编号', time_col])
            
        df_detail['还款期次_idx'] = df_detail.groupby('订单编号').cumcount() + 1
        for _, r in df_detail.iterrows():
            k = f"{r['订单编号']}_{r['还款期次_idx']}"
            detail_map[k] = r.get('还款类型', '')

    # 构建政策映射
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

    # 4. 处理线上分账 (Ledger)
    res_list = []
    for _, row in df_ledger.iterrows():
        oid = row.get('订单编号', '')
        if not oid or oid == "": continue # 跳过空ID
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

    # 5. 处理线下代付 (Payment) - 【深度重构】
    if not df_payment_raw.empty:
        if '支付批次号' not in df_payment_raw.columns:
            st.error("❌ 代付记录表中缺少【支付批次号】列！")
            st.stop()
            
        grouped = df_payment_raw.groupby(['支付批次号', '订单编号'])
        
        for (batch_id, oid), group in grouped:
            if not oid or oid == "": continue # 跳过空ID
            
            info = order_map.get(oid, {}) 
            service_fee = 0.0
            overdue_fee = 0.0
            pay_time = ''
            remark_parts = []
            
            for _, r in group.iterrows():
                note = str(r.get('系统备注', '')).strip()
                amt = safe_float(r.get('清分金额', 0))
                finish_time = r.get('完成时间', '')
                
                # 1. 抓取支付时间：只要该行有时间且当前还没抓到，就记录
                if pd.notna(finish_time) and str(finish_time).strip() != '' and (not pay_time):
                    pay_time = finish_time
                
                # 2. 精准匹配服务费
                if note == '服务费':
                    service_fee += amt
                # 3. 匹配延期服务费（备注里通常会有）
                elif '延期服务费' in note or '延滞费' in note:
                    service_fee += amt
                    remark_parts.append("含延期服务费")
                # 4. 匹配罚息/逾期
                elif '罚息' in note or '逾期' in note or '违约金' in note:
                    overdue_fee += amt
                    remark_parts.append("含罚息/逾期")

            new_row = {
                '业务订单号': oid, 
                '产品名称': info.get('产品名称', ''),
                '收款商户': info.get('收款商户', ''),
                '付款人': info.get('付款人', ''),
                '分期金额': info.get('分期金额', 0),
                '支付时间': pay_time, 
                '服务费': service_fee,
                '逾期费用': overdue_fee,
                '还款方式': '线下代付',
                '下单时间': info.get('下单时间', ''),
                '订单状态': info.get('订单状态', ''),
                '维护商务': info.get('维护商务', ''),
                '备注': "; ".join(remark_parts) if remark_parts else "",
                '_temp_oid': oid,
                '_temp_batch': batch_id,
                '_temp_time': pay_time # 用于排序
            }
            res_list.append(new_row)

    # 6. 生成最终 DataFrame 并匹配还款类型
    df_all = pd.DataFrame(res_list)
    
    # 单独处理线下数据的期次匹配
    df_offline = df_all[df_all['还款方式'] == '线下代付'].copy()
    if not df_offline.empty:
        # 【关键修复】按订单号和支付时间排序，确保生成的 seq 与明细表顺序一致
        df_offline = df_offline.sort_values(by=['_temp_oid', '_temp_time'])
        df_offline['seq'] = df_offline.groupby('_temp_oid').cumcount() + 1
        
        types_list = []
        for _, row in df_offline.iterrows():
            k = f"{row['_temp_oid']}_{row['seq']}"
            types_list.append(detail_map.get(k, '未匹配'))
            
        df_offline['还款期次'] = types_list
        df_all.loc[df_all['还款方式'] == '线下代付', '还款期次'] = df_offline['还款期次'].values

    # 7. 计算返佣
    if not df_all.empty:
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

    # 9. 强制输出固定 17 列
    final_columns = [
        '业务订单号', '产品名称', '收款商户', '付款人', '分期金额', '还款期次', 
        '支付时间', '服务费', '逾期费用', '还款方式', '下单时间', '订单状态', 
        '维护商务', '是否有返佣', '返佣比例', '返佣金额', '备注'
    ]
    
    for col in final_columns:
        if col not in df_all.columns:
            df_all[col] = ""
            
    df_result = df_all[final_columns].copy()
    return df_result

# ================= 网页界面部分 =================
col1, col2 = st.columns(2)
with col1:
    f_ledger = st.file_uploader("1. 上传分账记录表 (线上)", type=['xlsx', 'xls'])
    f_detail = st.file_uploader("3. 上传订单支付明细表", type=['xlsx', 'xls'])
    f_policy = st.file_uploader("5. 上传返佣政策表", type=['xlsx', 'xls'])
with col2:
    f_order = st.file_uploader("2. 上传订单主表", type=['xlsx', 'xls'])
    f_payment = st.file_uploader("4. 上传代付记录表 (线下)", type=['xlsx', 'xls'])

if st.button("开始计算", type="primary"):
    if all([f_ledger, f_payment, f_order, f_detail, f_policy]):
        with st.spinner("正在处理数据，请稍候..."):
            try:
                result_df = process_data(f_ledger, f_payment, f_order, f_detail, f_policy)
                st.success(f"✅ 计算完成！共处理 {len(result_df)} 条有效记录。")
                st.dataframe(result_df)
                
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    result_df.to_excel(writer, index=False, sheet_name='返佣计算结果')
                
                st.download_button(
                    label="📥 下载最终结果表格 (17列标准版)",
                    data=output.getvalue(),
                    file_name="返佣计算结果_标准版.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"❌ 发生错误: {str(e)}")
                import traceback
                st.code(traceback.format_exc())
    else:
        st.warning("⚠️ 请上传所有 5 个文件后再点击开始计算。")
