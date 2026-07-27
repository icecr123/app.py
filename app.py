import pandas as pd
import numpy as np
import warnings
import re
import os

# 忽略 Pandas 的一些样式警告
warnings.filterwarnings('ignore')

# ================= 配置区域 =================
FILE_LEDGER = '分账支付记录.xls'      # 来源1：线上还款
FILE_PAYMENT = '代付记录.xls'         # 来源2：线下还款
FILE_ORDER_MAIN = '订单.xls'          # 主订单信息表 (含分期金额、下单时间)
FILE_ORDER_DETAIL = '订单支付明细.xlsx' # 用于匹配代付的还款期次
FILE_POLICY = '返佣政策详情.xls'      # 返佣政策表
OUTPUT_FILE = '月度回款返佣计算结果_精准匹配版.xlsx'

# 定义最终输出的标准列头
FINAL_COLUMNS = [
    '业务订单号', '产品名称', '收款商户', '付款人', '分期金额', '还款期次', '支付时间',
    '服务费', '逾期费用', '还款方式', '下单时间', '订单状态', '维护商务',
    '是否有返佣', '返佣比例', '返佣金额', '备注'
]

# ================= 工具函数 =================
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
    """解析 X+Y 格式的产品名称"""
    if pd.isna(product_name): return False, 0, 0
    name_str = str(product_name).strip()
    match = re.search(r'(\d+)\+(\d+)', name_str)
    if match: return True, int(match.group(1)), int(match.group(2))
    return False, 0, 0

def count_periods(period_str):
    """使用正则通杀逻辑统计期数"""
    if pd.isna(period_str): return 1
    p_str = str(period_str)
    numbers = re.findall(r'\d+', p_str)
    return max(len(numbers), 1)

def calculate_commission(row, policy_map):
    """核心返佣计算逻辑"""
    merchant = str(row.get('收款商户', '')).strip()
    product = str(row.get('产品名称', '')).strip()
    period_str = str(row.get('还款期次', '')).strip()
    amount = safe_float(row.get('分期金额', 0))

    # 1. 获取策略字典
    key = f"{merchant}_{product}"
    policy = policy_map.get(key, {})
    if not policy: return pd.Series(['否', '0.0000', 0.0])

    # 2. 判断产品类型 (X+Y 还是 等额)
    is_xy, x_val, y_val = parse_xy_product(product)
    ratio = 0.0
    has_comm = '否'

    # 3. 统计还款期次
    p_num = count_periods(period_str)

    if is_xy:
        # 取最后一期来判断是X段还是Y段
        last_period = 0
        numbers = re.findall(r'\d+', period_str)
        if numbers: last_period = int(numbers[-1])
        if 0 < last_period <= x_val:
            raw_ratio = policy.get('X-返佣', 0)
        else:
            raw_ratio = policy.get('Y-返佣', 0)
    else:
        # 等额本息逻辑
        raw_ratio = policy.get('等额-返佣', 0)

    # 4. 安全转换比例并计算金额
    ratio = safe_float(raw_ratio)
    if ratio > 0:
        has_comm = '是'
        comm_amount = amount * ratio * p_num
        return pd.Series([has_comm, f"{ratio:.4f}", round(comm_amount, 2)])
    
    return pd.Series([has_comm, f"{ratio:.4f}", 0.0])

# ================= 主流程 =================
def process_data():
    print("开始执行数据清洗与计算...")
    print("="*40)

    # --- 1. 读取所有文件 ---
    try:
        df_ledger = pd.read_excel(FILE_LEDGER, dtype=str)
        print(f"[OK] 读取分账记录: {len(df_ledger)} 条")
    except Exception as e: print(f"[ERR] 分账记录读取失败: {e}"); return

    try:
        df_payment_raw = pd.read_excel(FILE_PAYMENT, dtype=str)
        print(f"[OK] 读取代付记录(原始): {len(df_payment_raw)} 条")
    except Exception as e: print(f"[ERR] 代付记录读取失败: {e}"); return

    try:
        df_order = pd.read_excel(FILE_ORDER_MAIN, dtype=str)
        print(f"[OK] 读取订单主表: {len(df_order)} 条")
    except Exception as e: print(f"[ERR] 订单主表读取失败: {e}"); return

    try:
        df_detail = pd.read_excel(FILE_ORDER_DETAIL, dtype=str)
        print(f"[OK] 读取订单支付明细: {len(df_detail)} 条")
    except Exception as e: print(f"[ERR] 订单支付明细读取失败: {e}"); return

    try:
        df_policy_raw = pd.read_excel(FILE_POLICY, dtype=str)
        print(f"[OK] 读取返佣政策: {len(df_policy_raw)} 条")
    except Exception as e: print(f"[ERR] 返佣政策读取失败: {e}"); return

    # --- 2. 预处理：建立映射字典 ---
    
    # A. 订单信息映射 (主表)
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

    # B. 返佣政策映射
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

    # C. 【关键修改】为订单支付明细生成还款序号
    # 按订单编号分组，为每笔还款记录打上序号 (1, 2, 3...)
    df_detail['还款序号'] = df_detail.groupby('订单编号').cumcount() + 1
    # 创建一个以 (订单编号, 还款序号) 为索引的 Series，方便快速查找还款类型
    detail_type_map = df_detail.set_index(['订单编号', '还款序号'])['还款类型']
    print(f"[OK] 已为订单支付明细生成还款序号，共 {len(detail_type_map)} 个唯一期次记录。")

    # --- 3. 处理分账记录 (线上) ---
    print("正在处理分账记录...")
    res_list = []
    for _, row in df_ledger.iterrows():
        oid = clean_order_id(row.get('业务订单号'))
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
        res_list.append(new_row)

    # --- 4. 处理代付记录 (线下 - 精准匹配版) ---
    print("正在处理代付记录 (执行精准匹配)...")
    if not df_payment_raw.empty:
        # 【关键修改】为代付记录也生成还款序号
        df_payment_raw['还款序号'] = df_payment_raw.groupby('业务订单号').cumcount() + 1
        
        # 按批次和订单号分组进行聚合
        grouped = df_payment_raw.groupby(['支付批次号', '业务订单号', '还款序号'])
        
        for (batch_id, oid, seq_num), group in grouped:
            oid_clean = clean_order_id(oid)
            info = order_map.get(oid_clean, {})
            
            total_service = 0.0
            total_overdue = 0.0
            service_time = None
            has_delay_note = False
            
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

            # 【关键修改】使用 (业务订单号, 还款序号) 双键匹配还款类型
            period_type = ''
            if oid_clean in df_detail['订单编号'].values:
                # 尝试从映射中获取还款类型
                period_type = detail_type_map.get((oid_clean, seq_num), '')
            
            final_pay_time = service_time if service_time else group.iloc[0].get('完成时间', '')
            remark_parts = []
            if has_delay_note: remark_parts.append("延期服务费")

            new_row = {
                '业务订单号': oid_clean,
                '产品名称': info.get('产品名称', ''),
                '收款商户': info.get('收款商户', ''),
                '付款人': info.get('付款人', ''),
                '分期金额': info.get('分期金额', 0),
                '还款期次': period_type, # 使用精准匹配到的期次
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

    # --- 5. 计算返佣 ---
    print("正在合并数据并计算返佣...")
    comm_results = df_all.apply(lambda row: calculate_commission(row, policy_map), axis=1)
    df_all['是否有返佣'] = comm_results[0]
    df_all['返佣比例'] = comm_results[1]
    df_all['返佣金额'] = comm_results[2]

    # --- 6. 补充日期备注 ---
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
                    if current_remarks: current_remarks += "，下单早于政策"
                    else: current_remarks = "下单早于政策"
                    # 更新 DataFrame 中的值
                    idx = row.name
                    df_all.at[idx, '返佣金额'] = 0.0
                    df_all.at[idx, '是否有返佣'] = '否'
            except Exception: pass
        return current_remarks

    df_all['备注'] = df_all.apply(check_date_and_adjust, axis=1)

    # --- 7. 导出 ---
    df_out = df_all[FINAL_COLUMNS]
    with pd.ExcelWriter(OUTPUT_FILE, engine='openpyxl') as writer:
        df_out.to_excel(writer, index=False, sheet_name='返佣计算结果')
    print(f"\n[成功] 计算完成！结果已保存至: {OUTPUT_FILE}")
    print(f" 总行数: {len(df_out)}")

if __name__ == '__main__':
    process_data()
