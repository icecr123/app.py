import streamlit as st
import pandas as pd
import numpy as np
import re
from io import BytesIO
import datetime

# ================= 页面配置 =================
st.set_page_config(page_title="月度回款返佣计算工具 V18-终极完整版", layout="wide")
st.title("🧮 月度回款返佣自动计算工具 (V18-终极完整版)")
st.markdown("""
**V18 终极修复说明：**
1. **精准适配多表列名**：订单主表识别`订单号`，支付明细表识别`订单编号`，最终输出统一为`业务订单号`。
2. **代付过滤逻辑**：代付记录中，若备注包含“本金”或“返服务费”，直接剔除，不参与计算。
3. **修复报错**：彻底解决线下代付处理时的 `The truth value of a Series is ambiguous` 错误。
4. **代码绝对完整**：包含所有辅助函数及主程序入口，无截断，可直接运行。
""")

# ================= 辅助函数 =================

def safe_float(val):
    """安全转换金额，防止 Series 类型错误"""
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
    """清洗备注：仅保留延期服务费相关，其余原样返回"""
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
    """统计期数"""
    if pd.isna(period_str): return 1
    numbers = re.findall(r'\d+', str(period_str))
    return max(len(numbers), 1)

def calculate_commission(row, policy_map):
    """计算返佣逻辑"""
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

                # --- 历史已还期数映射 (来自支付明细, 识别: 订单编号) ---
                history_map = {}
                if '订单编号' in df_detail.columns:
                    df_detail['_clean_oid'] = df_detail['订单编号'].astype(str).str.strip()
                    history_map = df_detail.groupby('_clean_oid').size().to_dict()

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
                    df_payment['_amount'] = df_payment['服务费'].apply(safe_float)
                    
                    if '备注' not in df_payment.columns:
                        df_payment['备注'] = ''
                    else:
                        df_payment['备注'] = df_payment['备注'].fillna('')

                    # 【过滤逻辑】剔除备注包含“本金”或“返服务费”的行
                    mask_exclude = df_payment['备注'].astype(str).str.contains('本金|返服务费', na=False)
                    df_payment_filtered = df_payment[~mask_exclude]
                    
                    grouped = df_payment_filtered.groupby(['_clean_oid', '支付批次号'])
                    runtime_counters = {} 

                    for (oid, batch_id), group in grouped:
                        info = order_map.get(oid, {})
                        base_paid = history_map.get(oid, 0)
                        curr_runtime = runtime_counters.get(oid, 0)
                        
                        service_rows = group[group['备注'].astype(str).str.contains('服务费', na=False)]
                        penalty_rows = group[group['备注'].astype(str).str.contains('罚息|逾期|违约金', na=False, regex=True)]
                        
                        total_penalty = penalty_rows['_amount'].sum()
                        penalty_added = False
                        
                        for _, s_row in service_rows.iterrows():
                            note = clean_remark(s_row['备注'])
                            is_deferred = '延期服务费' in note
                            
                            period = ""
                            if not is_deferred:
                                actual_period = base_paid + curr_runtime + 1
                                period = f"第{actual_period}期"
                                curr_runtime += 1
                            
                            p_to_add = total_penalty if not penalty_added else 0.0
                            if not penalty_added and total_penalty > 0: penalty_added = True
                            
                            results.append({
                                '业务订单号': oid, 
                                '产品名称': info.get('产品名称', ''),
                                '收款商户': info.get('收款商户', ''), 
                                '付款人': info.get('付款人', ''),
                                '分期金额': info.get('分期金额', 0), 
                                '还款期次': period,
                                '支付时间': s_row.get('支付时间', ''), 
                                '服务费': safe_float(s_row['_amount']),
                                '逾期费用': p_to_add, 
                                '还款方式': '线下代付',
                                '下单时间': info.get('下单时间', ''), 
                                '订单状态': info.get('订单状态', ''),
                                '维护商务': info.get('维护商务', ''),
                                '备注': note
                            })
                            
                        if service_rows.empty and total_penalty > 0:
                            actual_period = base_paid + curr_runtime + 1
                            results.append({
                                '业务订单号': oid, 
                                '产品名称': info.get('产品名称', ''),
                                '收款商户': info.get('收款商户', ''), 
                                '付款人': info.get('付款人', ''),
                                '分期金额': info.get('分期金额', 0), 
                                '还款期次': f"第{actual_period}期",
                                '支付时间': penalty_rows.iloc[0].get('支付时间', ''), 
                                '服务费': 0.0,
                                '逾期费用': total_penalty, 
                                '还款方式': '线下代付',
                                '下单时间': info.get('下单时间', ''), 
                                '订单状态': info.get('订单状态', ''),
                                '维护商务': info.get('维护商务', ''),
                                '备注': '补缴罚息/逾期'
                            })
                            curr_runtime += 1
                            
                    runtime_counters.update({k: v for k, v in runtime_counters.items()})

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
                    
                    # 【严格固定17列表头】
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
