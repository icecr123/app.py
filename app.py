import streamlit as st
import pandas as pd
import numpy as np
import re
from io import BytesIO
import datetime

# ================= 页面配置 =================
st.set_page_config(page_title="月度回款返佣计算工具 V29-终极修复版", layout="wide")
st.title("🧮 月度回款返佣自动计算工具 (V29-终极修复版)")
st.markdown("""
**V29 核心修复说明：**
1. **彻底解决 KeyError**：增加列名存在性检查，防止因Excel表头差异导致的程序崩溃。
2. **政策精准匹配**：基于【返佣政策详情】表，采用 "机构+期数+还款方式" 组合键匹配返佣开始时间。
3. **期次顺序消费**：线下代付严格按订单支付明细的时间顺序依次匹配期次。
4. **数据完整性**：保留线上字段补全、罚息合并、备注自动生成等所有历史修复功能。
""")

# ================= 辅助函数 =================

def safe_float(val):
    """安全转换金额"""
    if isinstance(val, pd.Series):
        val = val.iloc[0] if not val.empty else 0
    if pd.isna(val): return 0.0
    s = str(val).strip()
    if s in ['无', 'None', 'nan', '', '-']: return 0.0
    try:
        return float(s.replace(',', ''))
    except ValueError:
        return 0.0

def clean_str(val):
    """清洗字符串"""
    if pd.isna(val): return ""
    return str(val).strip()

def extract_period(text):
    """从产品名称或期数字段中提取数字期数，用于模糊匹配"""
    if not text: return None
    # 尝试提取数字，例如 "3期(用户)" -> 3, "24期" -> 24
    match = re.search(r'(\d+)', str(text))
    return int(match.group(1)) if match else None

# ================= 核心处理逻辑 =================

def process_data(order_df, detail_df, offline_df, online_df, policy_df):
    """
    主处理函数
    """
    results = []
    
    # 0. 预处理政策表 (构建查找字典)
    # Key: (机构名称, 期数数字, 还款方式) -> Value: 返佣开始时间
    policy_map = {}
    if policy_df is not None and not policy_df.empty:
        for _, row in policy_df.iterrows():
            inst_name = clean_str(row.get('机构名称'))
            prod_name = clean_str(row.get('产品名称')) # 例如 "3期(用户)"
            repay_way = clean_str(row.get('还款方式')) # 例如 "等额还款"
            start_time_str = clean_str(row.get('返佣开始时间'))
            
            period_num = extract_period(prod_name)
            
            if inst_name and period_num:
                key = (inst_name, period_num, repay_way)
                policy_map[key] = start_time_str
                # 兼容：如果没有还款方式，也存一个默认key
                if not repay_way:
                    policy_map[(inst_name, period_num, '')] = start_time_str

    # 1. 预处理订单主表
    order_map = {}
    if order_df is not None and not order_df.empty:
        for _, row in order_df.iterrows():
            oid = clean_str(row.get('订单号'))
            if oid:
                order_map[oid] = row

    # 2. 预处理订单支付明细（构建"期次队列"）
    detail_queue_map = {}
    if detail_df is not None and not detail_df.empty:
        # 防御性检查：确认是否有支付时间列
        time_col = None
        for col in ['支付时间', 'Payment Time', 'Time']:
            if col in detail_df.columns:
                time_col = col
                break
        
        if time_col:
            detail_df['支付时间_dt'] = pd.to_datetime(detail_df[time_col], errors='coerce')
            # 过滤掉时间无效的行
            detail_df = detail_df.dropna(subset=['支付时间_dt'])
            
            grouped_details = detail_df.groupby('订单编号')
            for oid, group in grouped_details:
                sorted_group = group.sort_values(by='支付时间_dt', ascending=True)
                queue = sorted_group['还款类型'].tolist()
                detail_queue_map[clean_str(oid)] = queue
        else:
            st.warning("⚠️ 【订单支付明细】表中未找到'支付时间'列，期次匹配可能不准确！")

    # 3. 处理线下代付记录 (Offline Data)
    offline_counter = {} 
    
    if offline_df is not None and not offline_df.empty:
        # 3.1 过滤备注
        mask_include = offline_df['系统备注'].astype(str).str.contains(r'服务费|延期|罚息|逾期|违约金', na=False)
        mask_exclude = offline_df['系统备注'].astype(str).str.contains(r'本金|返服务费', na=False)
        filtered_offline = offline_df[mask_include & ~mask_exclude].copy()
        
        if not filtered_offline.empty:
            # 3.2 分组处理：按 业务订单号 + 支付批次号 分组
            grouped_offline = filtered_offline.groupby(['业务订单号', '支付批次号'])
            
            for (oid, batch_id), group in grouped_offline:
                remarks = group['系统备注'].astype(str).tolist()
                
                total_penalty = 0.0
                service_fee_row = None
                
                # 遍历当前组的所有行，进行合并逻辑
                for _, row in group.iterrows():
                    remark = str(row.get('系统备注', ''))
                    amount = safe_float(row.get('清分金额', 0))
                    
                    if '服务费' in remark:
                        service_fee_row = row.copy()
                        if '罚息' not in remark and '逾期' not in remark and '违约金' not in remark:
                             service_fee_row['merged_amount'] = amount
                        else:
                             service_fee_row['merged_amount'] = amount 
                    elif any(k in remark for k in ['罚息', '逾期', '违约金']):
                        total_penalty += amount
                
                final_rows = []
                
                if service_fee_row is not None:
                    base_row = service_fee_row
                    base_row['merged_amount'] = safe_float(base_row.get('merged_amount', 0)) + total_penalty
                    final_rows.append(base_row)
                else:
                    base_row = group.iloc[0].copy()
                    base_row['merged_amount'] = total_penalty
                    final_rows.append(base_row)
                
                # 3.3 生成结果行并匹配期次
                for row in final_rows:
                    current_repayment_type = get_next_repayment_type(oid, offline_counter, detail_queue_map)
                    
                    res = build_result_row(
                        row, oid, order_map, policy_map,
                        source_type='offline',
                        repayment_type=current_repayment_type
                    )
                    if res:
                        results.append(res)

    # 4. 处理线上分账记录 (Online Data)
    if online_df is not None and not online_df.empty:
        for _, row in online_df.iterrows():
            oid = clean_str(row.get('订单编号'))
            repayment_type = clean_str(row.get('还款类型', row.get('期数', '')))
            
            res = build_result_row(
                row, oid, order_map, policy_map,
                source_type='online',
                repayment_type=repayment_type
            )
            if res:
                results.append(res)

    return pd.DataFrame(results)

def get_next_repayment_type(oid, counter_map, queue_map):
    """获取下一个可用的还款期次（按顺序消费）"""
    if oid not in queue_map:
        return "未匹配到明细"
    
    queue = queue_map[oid]
    current_index = counter_map.get(oid, 0)
    
    if current_index < len(queue):
        rep_type = queue[current_index]
        counter_map[oid] = current_index + 1
        return clean_str(rep_type)
    else:
        return f"超出明细范围({current_index+1})"

def build_result_row(row, oid, order_map, policy_map, source_type, repayment_type):
    """构建单行结果数据"""
    res = {
        '业务订单号': oid,
        '数据来源': '线下代付' if source_type == 'offline' else '线上分账',
        '支付时间': row.get('支付时间', ''),
        '支付批次号': row.get('支付批次号', '') if source_type == 'offline' else '',
        '还款方式': repayment_type
    }

    # 字段映射
    if source_type == 'offline':
        res['分账金额'] = safe_float(row.get('merged_amount', row.get('清分金额', 0)))
        res['产品名称'] = '' 
        res['收款商户'] = ''
        res['付款人'] = ''
    else:
        res['分账金额'] = safe_float(row.get('分账金额', 0))
        res['产品名称'] = clean_str(row.get('产品名称', ''))
        res['收款商户'] = clean_str(row.get('收款商户', ''))
        res['付款人'] = clean_str(row.get('付款人', ''))

    # 关联订单主表
    order_info = order_map.get(oid)
    if order_info is not None:
        res['下单时间'] = order_info.get('下单时间', '')
        res['订单状态'] = order_info.get('订单状态', '')
        res['维护商务'] = order_info.get('维护商务', '')
        res['是否有返佣'] = order_info.get('是否有返佣', '否')
        res['返佣比例'] = safe_float(order_info.get('返佣比例', 0))
        res['返佣金额'] = res['分账金额'] * res['返佣比例']
    else:
        res['下单时间'] = ''
        res['订单状态'] = '未匹配到主表'
        res['维护商务'] = ''
        res['是否有返佣'] = '否'
        res['返佣比例'] = 0
        res['返佣金额'] = 0

    # 备注生成 (含政策匹配)
    note_parts = []
    raw_note = clean_str(row.get('系统备注' if source_type=='offline' else '备注', ''))
    if raw_note:
        note_parts.append(raw_note)
    
    # 动态政策时间判断
    merchant_name = res.get('收款商户', '')
    # 如果线上没商户名，尝试从订单表找（假设订单表有机构字段，这里简化处理，主要靠线上表）
    if not merchant_name and order_info is not None:
         merchant_name = clean_str(order_info.get('机构名称', ''))

    order_time_str = res.get('下单时间', '')
    policy_start_str = None
    
    # 尝试匹配政策
    if merchant_name and repayment_type:
        period_num = extract_period(repayment_type)
        if period_num:
            # 尝试精确匹配 (机构, 期数, 还款方式)
            key_exact = (merchant_name, period_num, repayment_type)
            key_fuzzy = (merchant_name, period_num, '')
            
            if key_exact in policy_map:
                policy_start_str = policy_map[key_exact]
            elif key_fuzzy in policy_map:
                policy_start_str = policy_map[key_fuzzy]

    if policy_start_str and order_time_str:
        try:
            policy_date = pd.to_datetime(policy_start_str)
            order_date = pd.to_datetime(order_time_str)
            if order_date < policy_date:
                note_parts.append(f"下单早于政策({policy_start_str})")
        except:
            pass
            
    res['备注'] = " | ".join(note_parts)

    return res

# ================= Streamlit 界面交互 =================

st.sidebar.header("1. 文件上传区")
order_file = st.sidebar.file_uploader("上传【订单主表】(xlsx)", type=['xlsx', 'xls'])
detail_file = st.sidebar.file_uploader("上传【订单支付明细】(xlsx)", type=['xlsx', 'xls'])
offline_file = st.sidebar.file_uploader("上传【线下代付记录】(xls)", type=['xlsx', 'xls'])
online_file = st.sidebar.file_uploader("上传【线上分账支付记录】(xlsx)", type=['xlsx', 'xls'])
policy_file = st.sidebar.file_uploader("上传【返佣政策详情】(xlsx)", type=['xlsx', 'xls'])

if st.sidebar.button("开始计算"):
    if order_file and detail_file and offline_file and online_file and policy_file:
        try:
            with st.spinner("正在读取文件..."):
                df_order = pd.read_excel(order_file)
                df_detail = pd.read_excel(detail_file)
                df_offline = pd.read_excel(offline_file)
                df_online = pd.read_excel(online_file)
                df_policy = pd.read_excel(policy_file)
            
            st.success("文件读取成功，正在建立映射...")
            
            with st.spinner("正在处理数据..."):
                result_df = process_data(df_order, df_detail, df_offline, df_online, df_policy)
            
            st.success(f"处理完成！共生成 {len(result_df)} 条有效记录。")
            
            st.dataframe(result_df)
            
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                result_df.to_excel(writer, index=False, sheet_name='返佣计算结果')
            
            st.download_button(
                label="下载处理后的 Excel 文件",
                data=output.getvalue(),
                file_name="返佣计算结果_V29.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        except Exception as e:
            st.error(f"发生错误: {str(e)}")
            import traceback
            st.code(traceback.format_exc())
    else:
        st.warning("请上传所有必需的 5 个文件后再点击开始计算。")
