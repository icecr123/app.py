import streamlit as st
import pandas as pd
import numpy as np
import re
from io import BytesIO

# ================= 页面配置 =================
st.set_page_config(page_title="返佣计算工具 V10-备注逻辑完整版", layout="wide")
st.title("🧮 月度回款计算工具 (V10-备注逻辑完整版)")
st.markdown("""
**V10 备注逻辑强化说明：**
1. **延期服务费清洗**：自动将“延期手续费”替换为“延期服务费”，且期次强制留空。
2. **剔除冗余字符**：自动删除备注中多余的“含罚息/逾期”字样。
3. **全局备注校验**：线上与线下数据在合并后，统一进行备注清洗，确保格式绝对一致。
4. **核心业务保留**：期次动态锚定、多罚息合并、多服务费拆分逻辑完整保留。
""")

# ================= 辅助函数 =================
def safe_float(val):
    """安全转换金额"""
    if pd.isna(val): return 0.0
    s = str(val).strip()
    if s in ['无', 'None', 'nan', '']: return 0.0
    try: return float(s.replace(',', ''))
    except ValueError: return 0.0

def clean_order_id(oid):
    """暴力清洗订单号：提取纯数字"""
    if pd.isna(oid): return ""
    s = str(oid).strip()
    nums = re.findall(r'\d+', s)
    return "".join(nums) if nums else s

def normalize_columns(df):
    """标准化列名：去空格，防止 KeyError"""
    df.columns = [str(c).strip() for c in df.columns]
    return df

def clean_remark(remark_str):
    """
    全局备注清洗函数
    """
    if pd.isna(remark_str): return ""
    s = str(remark_str).strip()
    # 1. 替换“延期手续费”为“延期服务费”
    if "延期手续费" in s:
        s = s.replace("延期手续费", "延期服务费")
    # 2. 剔除“含罚息/逾期”或类似冗余字符
    s = re.sub(r'含罚息.*', '', s).strip()
    s = re.sub(r'含逾期.*', '', s).strip()
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

# ================= 侧边栏文件上传 =================
st.sidebar.header("📂 文件上传区")
file_ledger = st.sidebar.file_uploader("1. 上传【分账支付记录】(线上)", type=['xls', 'xlsx'])
file_payment = st.sidebar.file_uploader("2. 上传【代付记录】(线下)", type=['xls', 'xlsx'])
file_order = st.sidebar.file_uploader("3. 上传【订单主表】", type=['xls', 'xlsx'])
file_detail = st.sidebar.file_uploader("4. 上传【订单支付明细】(用于核对期次)", type=['xls', 'xlsx'])
file_policy = st.sidebar.file_uploader("5. 上传【返佣政策详情】", type=['xls', 'xlsx'])

if st.sidebar.button("🚀 开始计算"):
    if not all([file_ledger, file_payment, file_order, file_detail, file_policy]):
        st.error("请上传所有 5 个文件！")
    else:
        try:
            # 1. 读取并清洗列名
            df_ledger = normalize_columns(pd.read_excel(file_ledger, dtype=str))
            df_payment = normalize_columns(pd.read_excel(file_payment, dtype=str))
            df_order = normalize_columns(pd.read_excel(file_order, dtype=str))
            df_detail = normalize_columns(pd.read_excel(file_detail, dtype=str))
            df_policy_raw = normalize_columns(pd.read_excel(file_policy))

            st.success("文件读取成功，正在处理数据...")

            # 2. 构建基础映射字典
            order_map = {}
            for _, row in df_order.iterrows():
                oid = clean_order_id(row.get('订单编号', row.get('业务订单号', '')))
                if oid:
                    order_map[oid] = {
                        '产品名称': row.get('产品名称', ''),
                        '下单时间': row.get('下单时间', ''),
                        '订单状态': row.get('订单状态', ''),
                        '维护商务': row.get('业务员', row.get('维护商务', '')),
                        '付款人': row.get('客户姓名', row.get('付款人', '')),
                        '收款商户': row.get('机构简称', row.get('收款商户', '')),
                        '分期金额': safe_float(row.get('分期金额', 0))
                    }

            policy_map = {}
            for _, row in df_policy_raw.iterrows():
                inst = str(row.get('机构名称', '')).strip()
                prod = str(row.get('产品名称', '')).strip()
                if inst and prod:
                    policy_map[f"{inst}_{prod}"] = {
                        '等额-返佣': row.get('等额-返佣', 0),
                        'X-返佣': row.get('X-返佣', 0),
                        'Y-返佣': row.get('Y-返佣', 0)
                    }

            # 3. 计算历史已还期数 (解决期次错乱核心)
            history_map = {}
            detail_order_col = '订单编号' if '订单编号' in df_detail.columns else '业务订单号'
            if detail_order_col in df_detail.columns:
                df_detail['_clean_oid'] = df_detail[detail_order_col].apply(clean_order_id)
                history_map = df_detail.groupby('_clean_oid').size().to_dict()

            results = []

            # ================= 模块一：线上分账处理 =================
            st.info("正在处理线上分账数据...")
            ledger_oid_col = '业务订单号' if '业务订单号' in df_ledger.columns else '订单编号'
            for _, row in df_ledger.iterrows():
                oid = clean_order_id(row.get(ledger_oid_col, ''))
                if not oid: continue
                info = order_map.get(oid, {})
                
                results.append({
                    '业务订单号': oid,
                    '产品名称': info.get('产品名称', ''),
                    '收款商户': info.get('收款商户', ''),
                    '付款人': info.get('付款人', ''),
                    '分期金额': info.get('分期金额', 0),
                    '还款期次': str(row.get('还款期次', '')),
                    '支付时间': row.get('支付时间', row.get('完成时间', '')),
                    '服务费': safe_float(row.get('服务费', row.get('清分金额', 0))),
                    '逾期费用': safe_float(row.get('逾期费', row.get('罚息', 0))),
                    '还款方式': '线上还款',
                    '下单时间': info.get('下单时间', ''),
                    '订单状态': info.get('订单状态', ''),
                    '维护商务': info.get('维护商务', ''),
                    '备注': str(row.get('系统备注', '')) # 原始备注，后续统一清洗
                })

            # ================= 模块二：线下代付处理 =================
            st.info("正在处理线下代付数据...")
            pay_oid_col = '业务订单号' if '业务订单号' in df_payment.columns else '订单编号'
            
            df_payment['_clean_oid'] = df_payment[pay_oid_col].apply(clean_order_id)
            df_payment = df_payment[df_payment['_clean_oid'] != '']
            df_payment['_amount'] = df_payment['清分金额'].apply(safe_float)
            
            grouped = df_payment.groupby(['_clean_oid', '支付批次号'])
            runtime_counters = {} 

            for (oid, batch_id), group in grouped:
                info = order_map.get(oid, {})
                base_paid = history_map.get(oid, 0)
                curr_runtime = runtime_counters.get(oid, 0)
                
                service_rows = group[group['系统备注'].astype(str).str.contains('服务费', na=False)]
                penalty_rows = group[group['系统备注'].astype(str).str.contains('罚息|逾期|违约金', na=False, regex=True)]
                
                total_penalty = penalty_rows['_amount'].sum()
                penalty_added = False
                
                for _, s_row in service_rows.iterrows():
                    note = str(s_row['系统备注']).strip()
                    is_deferred = '延期' in note
                    if '延期手续费' in note: note = '延期服务费'
                    
                    period = ""
                    if not is_deferred:
                        actual_period = base_paid + curr_runtime + 1
                        period = f"第{actual_period}期"
                        curr_runtime += 1
                    
                    p_to_add = total_penalty if not penalty_added else 0.0
                    if not penalty_added and total_penalty > 0: penalty_added = True
                    
                    results.append({
                        '业务订单号': oid, '产品名称': info.get('产品名称', ''),
                        '收款商户': info.get('收款商户', ''), '付款人': info.get('付款人', ''),
                        '分期金额': info.get('分期金额', 0), '还款期次': period,
                        '支付时间': s_row.get('完成时间', ''), '服务费': safe_float(s_row['_amount']),
                        '逾期费用': p_to_add, '还款方式': '线下代付',
                        '下单时间': info.get('下单时间', ''), '订单状态': info.get('订单状态', ''),
                        '维护商务': info.get('维护商务', ''), '备注': note
                    })
                    
                if service_rows.empty and total_penalty > 0:
                    actual_period = base_paid + curr_runtime + 1
                    results.append({
                        '业务订单号': oid, '产品名称': info.get('产品名称', ''),
                        '收款商户': info.get('收款商户', ''), '付款人': info.get('付款人', ''),
                        '分期金额': info.get('分期金额', 0), '还款期次': f"第{actual_period}期",
                        '支付时间': penalty_rows.iloc[0].get('完成时间', ''), '服务费': 0.0,
                        '逾期费用': total_penalty, '还款方式': '线下代付',
                        '下单时间': info.get('下单时间', ''), '订单状态': info.get('订单状态', ''),
                        '维护商务': info.get('维护商务', ''), '备注': '补缴罚息/逾期'
                    })
                    curr_runtime += 1
                    
                runtime_counters[oid] = curr_runtime

            # ================= 模块三：汇总、备注清洗与返佣计算 =================
            st.info("正在清洗备注、计算返佣并生成结果...")
            df_result = pd.DataFrame(results)
            
            if not df_result.empty:
                # 【核心备注清洗逻辑】：全局统一处理
                df_result['备注'] = df_result['备注'].apply(clean_remark)
                
                # 【期次与备注联动】：如果备注是延期服务费，期次强制留空
                df_result.loc[df_result['备注'] == '延期服务费', '还款期次'] = ""

                # 计算返佣
                comm_results = df_result.apply(lambda row: calculate_commission(row, policy_map), axis=1)
                df_result['是否有返佣'] = comm_results[0]
                df_result['返佣比例'] = comm_results[1]
                df_result['返佣金额'] = comm_results[2]
                
                st.success(f"✅ 处理完成！共生成 {len(df_result)} 条有效记录。")
                st.dataframe(df_result.head(50))
                
                output = BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df_result.to_excel(writer, index=False, sheet_name='返佣计算结果')
                output.seek(0)
                
                st.download_button(
                    label="📥 下载处理后的 Excel 文件",
                    data=output,
                    file_name="月度回款返佣计算结果_V10备注完整版.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        except Exception as e:
            st.error(f"❌ 处理出错: {e}")
            import traceback
            st.code(traceback.format_exc())
else:
    st.warning("⚠️ 请在左侧上传全部 5 个 Excel 文件以开始处理...")
