import streamlit as st
import pandas as pd
import numpy as np
import re
from io import BytesIO
import datetime

# ================= 页面配置 =================
st.set_page_config(page_title="月度回款返佣计算工具 V26-含政策匹配版", layout="wide")
st.title("🧮 月度回款返佣自动计算工具 (V26-含政策匹配版)")
st.markdown("""
**V26 核心修复说明：**
1. **新增政策匹配**：引入【返佣政策详情】表，根据"收款商户/机构名称"动态匹配返佣开始时间。
2. **智能备注**：对比下单时间与对应机构的政策时间，自动标记"下单早于政策"。
3. **期次顺序消费**：线下还款严格按明细表顺序分配期次，不受时间误差影响。
4. **数据完整性**：修复线上分账字段丢失问题；修复线下罚息合并逻辑。
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

def parse_date(val):
    """安全解析日期"""
    if pd.isna(val): return None
    try:
        return pd.to_datetime(val)
    except:
        return None

# ================= 核心处理逻辑 =================

def process_data(order_df, detail_df, offline_df, online_df, policy_df):
    """
    主处理函数
    """
    results = []
    
    # 0. 预处理返佣政策表 (Policy Map)
    # 假设列名为：机构名称, 返佣开始时间
    policy_map = {}
    if policy_df is not None and not policy_df.empty:
        for _, row in policy_df.iterrows():
            inst_name = clean_str(row.get('机构名称'))
            start_time = parse_date(row.get('返佣开始时间'))
            if inst_name:
                policy_map[inst_name] = start_time

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
        if '支付时间' in detail_df.columns:
            detail_df['支付时间_dt'] = pd.to_datetime(detail_df['支付时间'], errors='coerce')
        
        grouped_details = detail_df.groupby('订单编号')
        for oid, group in grouped_details:
            sorted_group = group.sort_values(by='支付时间_dt', ascending=True)
            queue = sorted_group['还款类型'].tolist()
            detail_queue_map[clean_str(oid)] = queue

    # 3. 处理线下代付记录 (Offline Data)
    offline_counter = {} 
    
    if offline_df is not None and not offline_df.empty:
        mask_include = offline_df['系统备注'].astype(str).str.contains(r'服务费|延期|罚息|逾期|违约金', na=False)
        mask_exclude = offline_df['系统备注'].astype(str).str.contains(r'本金|返服务费', na=False)
        filtered_offline = offline_df[mask_include & ~mask_exclude].copy()
        
        if not filtered_offline.empty:
            grouped_offline = filtered_offline.groupby(['业务订单号', '支付批次号'])
            
            for (oid, batch_id), group in grouped_offline:
                remarks = group['系统备注'].astype(str).tolist()
                
                total_penalty = 0.0
                service_fee_row = None
                
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
                
                for row in final_rows:
                    current_repayment_type = get_next_repayment_type(oid, offline_counter, detail_queue_map)
                    res = build_result_row(
                        row, oid, order_map, policy_map,
                        source_type='offline',
                        repayment_type=current_repayment_type
                    )
                    if res: results.append(res)

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
            if res: results.append(res)

    return pd.DataFrame(results)

def get_next_repayment_type(oid, counter_map, queue_map):
    """获取下一个可用的还款期次（按顺序消费）"""
    if oid not in queue_map: return "未匹配到明细"
    queue = queue_map[oid]
    current_index = counter_map.get(oid, 0)
    
    if current_index < len(queue):
        rep_type = queue[current_index]
        counter_map[oid] = current_index + 1
        return clean_str(rep_type)
    else:
        return f"超出明细范围({current_index+1})"

def check_policy_note(order_time_str, merchant_name, policy_map):
    """检查是否需要添加'下单早于政策'备注"""
    if not order_time_str or not merchant_name: return ""
    
    order_time = parse_date(order_time_str)
    policy_start_time = policy_map.get(merchant_name)
    
    if order_time and policy_start_time:
        if order_time < policy_start_time:
            return "下单早于政策"
    return ""

def build_result_row(row, oid, order_map, policy_map, source_type, repayment_type):
    """构建单行结果数据"""
    res = {
        '业务订单号': oid,
        '数据来源': '线下代付' if source_type == 'offline' else '线上分账',
        '支付时间': row.get('支付时间', ''),
        '支付批次号': row.get('支付批次号', '') if source_type == 'offline' else '',
        '还款方式': repayment_type
    }

    # 字段映射与机构获取
    merchant_name = ""
    if source_type == 'offline':
        res['分账金额'] = safe_float(row.get('merged_amount', row.get('清分金额', 0)))
        res['产品名称'] = '' 
        res['收款商户'] = ''
        res['付款人'] = ''
        # 线下代付通常没有直接的商户名，可能需要从订单表或其他地方找，这里暂置空
        # 如果线下表里有"机构"字段，可以改为 row.get('机构', '')
    else:
        res['分账金额'] = safe_float(row.get('分账金额', 0))
        res['产品名称'] = clean_str(row.get('产品名称', ''))
        res['收款商户'] = clean_str(row.get('收款商户', ''))
        res['付款人'] = clean_str(row.get('付款人', ''))
        merchant_name = res['收款商户'] # 线上直接用收款商户作为机构名去查政策

    # 关联订单主表
    order_info = order_map.get(oid)
    if order_info is not None:
        res['下单时间'] = order_info.get('下单时间', '')
        res['订单状态'] = order_info.get('订单状态', '')
        res['维护商务'] = order_info.get('维护商务', '')
        res['是否有返佣'] = order_info.get('是否有返佣', '否')
        res['返佣比例'] = safe_float(order_info.get('返佣比例', 0))
        res['返佣金额'] = res['分账金额'] * res['返佣比例']
        
        # 如果线下没拿到商户名，尝试从订单表拿（假设订单表有"机构名称"或类似字段）
        if source_type == 'offline' and not merchant_name:
             merchant_name = clean_str(order_info.get('机构名称', order_info.get('收款商户', '')))
    else:
        res['下单时间'] = ''
        res['订单状态'] = '未匹配到主表'
        res['维护商务'] = ''
        res['是否有返佣'] = '否'
        res['返佣比例'] = 0
        res['返佣金额'] = 0

    # 备注生成
    note_parts = []
    raw_note = clean_str(row.get('系统备注' if source_type=='offline' else '备注', ''))
    if raw_note: note_parts.append(raw_note)
    
    # 动态政策检查
    policy_note = check_policy_note(res['下单时间'], merchant_name, policy_map)
    if policy_note: note_parts.append(policy_note)
        
    res['备注'] = " | ".join(note_parts)

    return res

# ================= Streamlit 界面交互 =================

st.sidebar.header("1. 文件上传区")
order_file = st.sidebar.file_uploader("上传【订单主表】(xlsx)", type=['xlsx', 'xls'])
detail_file = st.sidebar.file_uploader("上传【订单支付明细】(xlsx)", type=['xlsx', 'xls'])
offline_file = st.sidebar.file_uploader("上传【线下代付记录】(xls)", type=['xlsx', 'xls'])
online_file = st.sidebar.file_uploader("上传【线上分账支付记录】(xlsx)", type=['xlsx', 'xls'])
policy_file = st.sidebar.file_uploader("上传【返佣政策详情】(xlsx)", type=['xlsx', 'xls']) # 新增

if st.sidebar.button("开始计算"):
    # 校验5个文件是否齐全
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
                file_name="返佣计算结果_V26.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        except Exception as e:
            st.error(f"发生错误: {str(e)}")
            import traceback
            st.code(traceback.format_exc())
    else:
        st.warning("请上传所有必需的 5 个文件后再点击开始计算。")
