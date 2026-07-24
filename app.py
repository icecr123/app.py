import streamlit as st
import pandas as pd
import numpy as np
import re
from io import BytesIO
import datetime

# ================= 页面配置 =================
st.set_page_config(page_title="月度回款返佣计算工具 V16-最终版", layout="wide")
st.title("🧮 月度回款返佣自动计算工具 (V16-最终版)")
st.markdown("""
**V16 最终修复说明：**
1. **字段名修正**：彻底删除“合同号”映射，严格保留“业务订单号/订单编号”。
2. **严格固定17列表头**：完全按照业务标准顺序输出，彻底删除多余的“是否有过户”列。
3. **返佣与备注还原**：恢复原有政策表校验逻辑，当“下单早于政策开始时间”时，返佣为0并在备注追加原因。
""")

# ================= 辅助函数 =================

def safe_float(val):
    """安全转换金额"""
    if pd.isna(val): return 0.0
    s = str(val).strip()
    if s in ['无', 'None', 'nan', '', '-']: return 0.0
    try:
        return float(s.replace(',', ''))
    except ValueError:
        return 0.0

def clean_columns(df):
    """清洗列名并建立标准映射"""
    df.columns = [str(c).strip() for c in df.columns]
    col_mapping = {
        '业务订单号': '订单编号', '订单编号': '订单编号',
        # 【修正】彻底删除 '合同号' 映射
        '产品名称': '产品名称', '产品': '产品名称',
        '收款商户': '收款商户', '商户': '收款商户',
        '付款人': '付款人', '客户姓名': '付款人',
        '分期金额': '分期金额', '贷款金额': '分期金额', '本金': '分期金额',
        '还款期次': '还款期次', '期数': '还款期次', '当前期数': '还款期次',
        '支付时间': '支付时间', '交易时间': '支付时间', '还款日期': '支付时间',
        '服务费': '服务费', '手续费': '服务费', '利息': '服务费',
        '逾期费用': '逾期费用', '罚息': '逾期费用', '违约金': '逾期费用',
        '还款方式': '还款方式', '渠道': '还款方式', '来源': '还款方式',
        '下单时间': '下单时间', '创建时间': '下单时间',
        '订单状态': '订单状态', '状态': '订单状态',
        '维护商务': '维护商务', '商务': '维护商务', '业务员': '维护商务',
        '备注': '备注', '说明': '备注'
    }
    rename_dict = {col: col_mapping[col] for col in df.columns if col in col_mapping}
    df.rename(columns=rename_dict, inplace=True)
    return df

def clean_remark(remark):
    """【严格还原】备注清洗逻辑：仅处理延期手续费"""
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
    """【严格还原】原有的返佣计算逻辑（含下单时间校验及备注追加）"""
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
    
    # 校验下单时间是否早于政策返佣开始时间
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
                # 1. 读取并清洗列名
                df_ledger = clean_columns(pd.read_excel(file_ledger, dtype=str))
                df_payment = clean_columns(pd.read_excel(file_payment, dtype=str))
                df_order = clean_columns(pd.read_excel(file_order, dtype=str))
                df_detail = clean_columns(pd.read_excel(file_detail, dtype=str))
                df_policy_raw = clean_columns(pd.read_excel(file_policy, dtype=str))

                st.success("文件读取成功，正在处理数据...")

                # 2. 构建基础映射字典
                order_map = {}
                for _, row in df_order.iterrows():
                    oid = str(row.get('订单编号', '')).strip()
                    if oid:
                        order_map[oid] = {
                            '产品名称': row.get('产品名称', ''),
                            '下单时间': row.get('下单时间', ''),
                            '订单状态': row.get('订单状态', ''),
                            '维护商务': row.get('维护商务', ''),
                            '付款人': row.get('付款人', ''),
                            '收款商户': row.get('收款商户', ''),
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
                            'Y-返佣': row.get('Y-返佣', 0),
                            '开始时间': row.get('开始时间', None)
                        }

                # 3. 计算历史已还期数
                history_map = {}
                if '订单编号' in df_detail.columns:
                    df_detail['_clean_oid'] = df_detail['订单编号'].astype(str).str.strip()
                    history_map = df_detail.groupby('_clean_oid').size().to_dict()

                results = []

                # ================= 模块一：线上分账处理 =================
                st.info("正在处理线上分账数据...")
                for _, row in df_ledger.iterrows():
                    oid = str(row.get('订单编号', '')).strip()
                    if not oid: continue
                    info = order_map.get(oid, {})
                    
                    results.append({
                        '订单编号': oid,
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
                df_payment['_clean_oid'] = df_payment['订单编号'].astype(str).str.strip()
                df_payment = df_payment[df_payment['_clean_oid'] != '']
                df_payment['_amount'] = df_payment['服务费'].apply(safe_float)
                
                grouped = df_payment.groupby(['_clean_oid', '支付批次号'])
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
                            '订单编号': oid, '产品名称': info.get('产品名称', ''),
                            '收款商户': info.get('收款商户', ''), '付款人': info.get('付款人', ''),
                            '分期金额': info.get('分期金额', 0), '还款期次': period,
                            '支付时间': s_row.get('支付时间', ''), '服务费': safe_float(s_row['_amount']),
                            '逾期费用': p_to_add, '还款方式': '线下代付',
                            '下单时间': info.get('下单时间', ''), '订单状态': info.get('订单状态', ''),
                            '维护商务': info.get('维护商务', ''),
                            '备注': note
                        })
                        
                    if service_rows.empty and total_penalty > 0:
                        actual_period = base_paid + curr_runtime + 1
                        results.append({
                            '订单编号': oid, '产品名称': info.get('产品名称', ''),
                            '收款商户': info.get('收款商户', ''), '付款人': info.get('付款人', ''),
                            '分期金额': info.get('分期金额', 0), '还款期次': f"第{actual_period}期",
                            '支付时间': penalty_rows.iloc[0].get('支付时间', ''), '服务费': 0.0,
                            '逾期费用': total_penalty, '还款方式': '线下代付',
                            '下单时间': info.get('下单时间', ''), '订单状态': info.get('订单状态', ''),
                            '维护商务': info.get('维护商务', ''),
                            '备注': '补缴罚息/逾期'
                        })
                        curr_runtime += 1
                    runtime_counters[oid] = curr_runtime

                # ================= 模块三：汇总与返佣计算 =================
                st.info("正在计算返佣并生成结果...")
                df_result = pd.DataFrame(results)
                
                if not df_result.empty:
                    comm_results = df_result.apply(lambda row: calculate_commission(row, policy_map), axis=1)
                    df_result['是否有返佣'] = comm_results[0]
                    df_result['返佣比例'] = comm_results[1]
                    df_result['返佣金额'] = comm_results[2]
                    
                    # 将不返佣原因追加到备注中
                    df_result['备注'] = df_result['备注'].astype(str) + comm_results[3].apply(
                        lambda x: f"；{x}" if x else ""
                    )
                    
                    # 【严格还原】固定输出17列表头逻辑（彻底删除“是否有过户”）
                    target_columns = [
                        '订单编号', '产品名称', '收款商户', '付款人', '分期金额', 
                        '还款期次', '支付时间', '服务费', '逾期费用', '还款方式', 
                        '下单时间', '订单状态', '维护商务', '是否有返佣', '返佣比例', 
                        '返佣金额', '备注'
                    ]
                    for col in target_columns:
                        if col not in df_result.columns:
                            df_result[col] = ""
                    
                    # 将内部使用的 '订单编号' 映射回 '业务订单号'
                    if '订单编号' in df_result.columns:
                        df_result.rename(columns={'订单编号': '业务订单号'}, inplace=True)
                        target_columns[0] = '业务订单号' # 同步更新导出表头
                        
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
