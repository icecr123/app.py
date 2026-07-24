import streamlit as st
import pandas as pd
import numpy as np
import re
from io import BytesIO
import datetime

# ================= 页面配置 =================
st.set_page_config(page_title="月度回款返佣计算工具 V22-精准期次与罚息合并版", layout="wide")
st.title("🧮 月度回款返佣自动计算工具 (V22-精准期次与罚息合并版)")
st.markdown("""
**V22 核心修复说明：**
1. **代付金额列精准锁定**：严格使用 `清分金额` 列提取代付金额。
2. **代付备注过滤逻辑重写**：仅抓取“服务费/罚息/逾期/违约金”，剔除“本金/返服务费”。
3. **罚息合并逻辑**：同一订单+同批次下，若同时有服务费和罚息，罚息金额合并至服务费行；若仅有罚息，则单独成行。
4. **期次精准递增（核心修复）**：不再全局+1，而是按订单在【订单支付明细】中的历史还款时间排序，依次消费期次，完美解决期次错乱问题。
5. **多表列名适配**：订单主表识别`订单号`，支付明细识别`订单编号`，最终输出统一为`业务订单号`。
""")

# ================= 辅助函数 =================

def safe_float(val):
    """安全地将任意值转换为浮点数"""
    if isinstance(val, pd.Series):
        val = val.iloc[0] if not val.empty else 0
    if pd.isna(val): return 0.0
    s = str(val).strip()
    if s in ['无', 'None', 'nan', '', '-']: return 0.0
    try:
        return float(s.replace(',', ''))
    except ValueError:
        return 0.0

def clean_remark(remark):
    """清洗备注字段，仅保留延期服务费相关描述"""
    if pd.isna(remark): return ""
    s = str(remark).strip()
    if "延期手续费" in s or "延期服务费" in s:
        period_match = re.search(r'\d+期', s)
        period_str = period_match.group(0) if period_match else ""
        return f"延期服务费{period_str}"
    return s

def parse_xy_product(product_name):
    """解析 x+y 产品格式"""
    if pd.isna(product_name): return False, 0, 0
    name_str = str(product_name).strip()
    match = re.search(r'(\d+)\+(\d+)', name_str)
    if match: return True, int(match.group(1)), int(match.group(2))
    return False, 0, 0

def count_periods(period_str):
    """统计还款期次字符串中的数字个数"""
    if pd.isna(period_str): return 1
    numbers = re.findall(r'\d+', str(period_str))
    return max(len(numbers), 1)

def calculate_commission(row, policy_map):
    """根据订单信息和返佣政策，计算该笔还款的返佣情况"""
    merchant = str(row.get('收款商户', '')).strip()
    product = str(row.get('产品名称', '')).strip()
    period_str = str(row.get('还款期次', '')).strip()
    amount = safe_float(row.get('分期金额', 0))
    order_time = row.get('下单时间', '')
    
    key = f"{merchant}_{product}"
    policy = policy_map.get(key, {})
    no_comm_reason = "" 
    
    if not policy: 
        return pd.Series(['否', '0.0000', 0.0, no_comm_reason])
    
    policy_start_time = policy.get('开始时间', None)
    if policy_start_time and pd.notna(order_time):
        try:
            if pd.to_datetime(order_time) < pd.to_datetime(policy_start_time):
                no_comm_reason = "下单早于返佣政策开始时间"
                return pd.Series(['否', '0.0000', 0.0, no_comm_reason])
        except Exception:
            pass 

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
        
    return pd.Series([has_comm, f"{ratio:.4f}", round(comm_amount, 2), no_comm_reason])

# ================= 主程序入口 =================

def main():
    st.sidebar.header("📂 文件上传区")
    file_ledger = st.sidebar.file_uploader("1. 上传【分账支付记录】(线上)", type=['xls', 'xlsx'])
    file_payment = st.sidebar.file_uploader("2. 上传【代付记录】(线下)", type=['xls', 'xlsx'])
    file_order = st.sidebar.file_uploader("3. 上传【订单主表】", type=['xls', 'xlsx'])
    file_detail = st.sidebar.file_uploader("4. 上传【订单支付明细】(核对期次)", type=['xls', 'xlsx'])
    file_policy = st.sidebar.file_uploader("5. 上传【返佣政策详情】", type=['xls', 'xlsx'])

    if st.sidebar.button("🚀 开始计算"):
        if not all([file_ledger, file_payment, file_order, file_detail, file_policy]):
            st.error("请上传所有 5 个文件！")
        else:
            try:
                # 1. 读取文件 (保持原始列名)
                df_ledger = pd.read_excel(file_ledger, dtype=str)
                df_payment = pd.read_excel(file_payment, dtype=str)
                df_order = pd.read_excel(file_order, dtype=str)
                df_detail = pd.read_excel(file_detail, dtype=str)
                df_policy_raw = pd.read_excel(file_policy, dtype=str)

                st.success("文件读取成功，正在建立映射...")

                # 2. 构建基础映射字典
                
                # --- 订单主表映射 (识别: 订单号) ---
                order_map = {}
                if '订单号' in df_order.columns:
                    for _, row in df_order.iterrows():
                        oid = str(row.get('订单号', '')).strip()
                        if oid:
                            order_map[oid] = {
                                '产品名称': str(row.get('产品名称', '')).strip(),
                                '下单时间': row.get('下单时间', ''),
                                '订单状态': str(row.get('订单状态', '')).strip(),
                                '维护商务': str(row.get('维护商务', '')).strip(),
                                '付款人': str(row.get('付款人', '')).strip(),
                                '收款商户': str(row.get('收款商户', '')).strip(),
                                '分期金额': safe_float(row.get('分期金额', 0))
                            }

                # --- 政策表映射 ---
                policy_map = {}
                for _, row in df_policy_raw.iterrows():
                    inst = str(row.get('机构名称', '')).strip()
                    prod = str(row.get('产品名称', '')).strip()
                    if inst and prod:
                        policy_map[f"{inst}_{prod}"] = {
                            '等额-返佣': row.get('等额-返佣', 0),
                            'X-返佣': row.get('X-返佣', 0),
                            'Y-返佣': row.get('Y-返佣', 0),
                            '开始时间': row.get('开始时间', None)
                        }

                # --- 【核心修复】历史还款期次队列映射 (来自支付明细, 识别: 订单编号) ---
                # 按订单号分组，按支付时间正序排列，将还款类型存入列表，用于后续按顺序消费
                history_queue_map = {}
                if '订单编号' in df_detail.columns:
                    df_detail_sorted = df_detail.copy()
                    df_detail_sorted['_clean_oid'] = df_detail_sorted['订单编号'].astype(str).str.strip()
                    # 假设支付明细表中有支付时间或类似时间列，这里用通用名，如果没有则按原顺序
                    time_col = None
                    for c in ['支付时间', '还款时间', '交易时间']:
                        if c in df_detail_sorted.columns:
                            time_col = c
                            break
                    if time_col:
                        df_detail_sorted = df_detail_sorted.sort_values(['_clean_oid', time_col])
                    
                    # 提取还款类型（假设列名为 还款类型 或 期次）
                    period_col = None
                    for c in ['还款类型', '期次', '还款期次']:
                        if c in df_detail_sorted.columns:
                            period_col = c
                            break
                            
                    if period_col:
                        for oid, group in df_detail_sorted.groupby('_clean_oid'):
                            history_queue_map[oid] = list(group[period_col].astype(str).values)

                results = []

                # ================= 模块一：线上分账处理 =================
                st.info("正在处理线上分账数据...")
                ledger_id_col = None
                for c in ['业务订单号', '订单号', '订单编号']:
                    if c in df_ledger.columns:
                        ledger_id_col = c
                        break

                if ledger_id_col:
                    for _, row in df_ledger.iterrows():
                        oid = str(row.get(ledger_id_col, '')).strip()
                        if not oid: continue
                        info = order_map.get(oid, {})
                        
                        results.append({
                            '业务订单号': oid,
                            '产品名称': info.get('产品名称', ''),
                            '收款商户': info.get('收款商户', ''),
                            '付款人': info.get('付款人', ''),
                            '分期金额': info.get('分期金额', 0),
                            '还款期次': str(row.get('还款期次', '')),
                            '支付时间': row.get('支付时间', ''),
                            '服务费': safe_float(row.get('服务费', 0)),
                            '逾期费用': safe_float(row.get('逾期费用', 0)),
                            '还款方式': '线上还款',
                            '下单时间': info.get('下单时间', ''),
                            '订单状态': info.get('订单状态', ''),
                            '维护商务': info.get('维护商务', ''),
                            '备注': clean_remark(row.get('备注', ''))
                        })

                # ================= 模块二：线下代付处理 =================
                st.info("正在处理线下代付数据...")
                
                payment_id_col = None
                for c in ['业务订单号', '订单号', '订单编号']:
                    if c in df_payment.columns:
                        payment_id_col = c
                        break

                if payment_id_col:
                    df_payment['_clean_oid'] = df_payment[payment_id_col].astype(str).str.strip()
                    df_payment = df_payment[df_payment['_clean_oid'] != '']
                    
                    # 严格使用“清分金额”列
                    if '清分金额' in df_payment.columns:
                        df_payment['_amount'] = df_payment['清分金额'].apply(safe_float)
                    else:
                        st.warning("⚠️ 代付记录中未找到‘清分金额’列，代付金额将默认记为0。")
                        df_payment['_amount'] = 0.0
                    
                    if '备注' not in df_payment.columns:
                        df_payment['备注'] = ''
                    else:
                        df_payment['备注'] = df_payment['备注'].fillna('')

                    # 过滤逻辑：包含服务费/罚息，且排除本金/返服务费
                    include_mask = df_payment['备注'].astype(str).str.contains('服务费|罚息|逾期|违约金', na=False)
                    exclude_mask = df_payment['备注'].astype(str).str.contains('本金|返服务费', na=False)
                    valid_mask = include_mask & (~exclude_mask)
                    df_payment_filtered = df_payment[valid_mask]
                    
                    filtered_out_count = len(df_payment) - len(df_payment_filtered)
                    if filtered_out_count > 0:
                        st.info(f"💡 代付记录中已自动过滤掉 {filtered_out_count} 条不符合要求的数据。")

                    # 【核心修复】按 订单号 + 支付批次号 分组
                    batch_id_col = '支付批次号' if '支付批次号' in df_payment_filtered.columns else None
                    group_cols = ['_clean_oid']
                    if batch_id_col: group_cols.append(batch_id_col)
                    
                    grouped = df_payment_filtered.groupby(group_cols)
                    
                    # 记录当前每个订单已经消费了多少次历史期次
                    runtime_period_idx = {oid: 0 for oid in history_queue_map.keys()}

                    for name, group in grouped:
                        oid = name[0]
                        info = order_map.get(oid, {})
                        
                        # 区分当前批次下的服务费和罚息
                        service_rows = group[group['备注'].astype(str).str.contains('服务费', na=False)]
                        penalty_rows = group[group['备注'].astype(str).str.contains('罚息|逾期|违约金', na=False, regex=True)]
                        
                        total_penalty = penalty_rows['_amount'].sum()
                        
                        # 获取当前订单应该匹配的还款期次
                        current_period = ""
                        if oid in history_queue_map:
                            idx = runtime_period_idx.get(oid, 0)
                            queue = history_queue_map[oid]
                            if idx < len(queue):
                                current_period = queue[idx]
                                # 只有当本批次有服务费时，才认为消费了一个新的期次
                                if not service_rows.empty:
                                    runtime_period_idx[oid] = idx + 1
                            else:
                                # 超出历史记录，按递增生成
                                base = len(queue)
                                actual_p = base + runtime_period_idx.get(oid, 0) + 1
                                current_period = f"第{actual_p}期"
                                if not service_rows.empty:
                                    runtime_period_idx[oid] = runtime_period_idx.get(oid, 0) + 1
                        else:
                            # 没有历史记录的订单，从第1期开始递增
                            actual_p = runtime_period_idx.get(oid, 0) + 1
                            current_period = f"第{actual_p}期"
                            if not service_rows.empty:
                                runtime_period_idx[oid] = actual_p

                        # 场景1：本批次有服务费（可能伴随罚息）
                        if not service_rows.empty:
                            for _, s_row in service_rows.iterrows():
                                note = clean_remark(s_row['备注'])
                                results.append({
                                    '业务订单号': oid, 
                                    '产品名称': info.get('产品名称', ''),
                                    '收款商户': info.get('收款商户', ''), 
                                    '付款人': info.get('付款人', ''),
                                    '分期金额': info.get('分期金额', 0), 
                                    '还款期次': current_period,
                                    '支付时间': s_row.get('支付时间', ''), 
                                    '服务费': safe_float(s_row['_amount']),
                                    '逾期费用': total_penalty,  # 罚息合并到服务费行
                                    '还款方式': '线下代付',
                                    '下单时间': info.get('下单时间', ''), 
                                    '订单状态': info.get('订单状态', ''),
                                    '维护商务': info.get('维护商务', ''),
                                    '备注': note
                                })
                        # 场景2：本批次只有罚息，没有服务费
                        elif total_penalty > 0:
                            results.append({
                                '业务订单号': oid, 
                                '产品名称': info.get('产品名称', ''),
                                '收款商户': info.get('收款商户', ''), 
                                '付款人': info.get('付款人', ''),
                                '分期金额': info.get('分期金额', 0), 
                                '还款期次': current_period,
                                '支付时间': penalty_rows.iloc[0].get('支付时间', ''), 
                                '服务费': 0.0,
                                '逾期费用': total_penalty, 
                                '还款方式': '线下代付',
                                '下单时间': info.get('下单时间', ''), 
                                '订单状态': info.get('订单状态', ''),
                                '维护商务': info.get('维护商务', ''),
                                '备注': '补缴罚息/逾期'
                            })

                # ================= 模块三：汇总与返佣计算 =================
                st.info("正在计算返佣并生成结果...")
                df_result = pd.DataFrame(results)
                
                if not df_result.empty:
                    comm_results = df_result.apply(lambda row: calculate_commission(row, policy_map), axis=1)
                    df_result['是否有返佣'] = comm_results[0]
                    df_result['返佣比例'] = comm_results[1]
                    df_result['返佣金额'] = comm_results[2]
                    
                    df_result['备注'] = df_result['备注'].astype(str) + comm_results[3].apply(
                        lambda x: f"；{x}" if x else ""
                    )
                    
                    # 严格固定17列表头
                    target_columns = [
                        '业务订单号', '产品名称', '收款商户', '付款人', '分期金额', 
                        '还款期次', '支付时间', '服务费', '逾期费用', '还款方式', 
                        '下单时间', '订单状态', '维护商务', '是否有返佣', '返佣比例', 
                        '返佣金额', '备注'
                    ]
                    
                    for col in target_columns:
                        if col not in df_result.columns:
                            df_result[col] = ""
                            
                    df_output = df_result[target_columns]

                    st.success(f"✅ 处理完成！共生成 {len(df_output)} 条有效记录。")
                    st.dataframe(df_output.head(50), use_container_width=True)
                    
                    output = BytesIO()
                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                        df_output.to_excel(writer, index=False, sheet_name='返佣计算结果')
                    output.seek(0)
                    
                    st.download_button(
                        label="📥 下载处理后的 Excel 文件",
                        data=output,
                        file_name=f"月度回款返佣计算结果_{datetime.date.today()}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

            except Exception as e:
                st.error(f"❌ 处理出错: {e}")
                import traceback
                st.code(traceback.format_exc())
    else:
        st.warning("⚠️ 请在左侧上传全部 5 个 Excel 文件以开始处理...")

if __name__ == "__main__":
    main()
