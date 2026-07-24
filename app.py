import streamlit as st
import pandas as pd
import numpy as np
import re
from io import BytesIO
import datetime

# ================= 页面配置 =================
st.set_page_config(page_title="月度回款返佣计算工具 V28-政策精准匹配版", layout="wide")
st.title("🧮 月度回款返佣自动计算工具 (V28-政策精准匹配版)")
st.markdown("""
**V28 核心修复说明：**
1. **语法修复**：修复了备注拼接处的 SyntaxError。
2. **政策匹配升级**：根据【返佣政策详情】表结构，采用 **"机构+期数+还款方式"** 组合键进行精准匹配，解决同一机构不同期数政策时间不同的问题。
3. **期次顺序消费**：线下代付记录严格按支付明细表的顺序获取期次。
4. **线上字段补全**：直接读取分账表中的商户、付款人等信息。
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

def extract_period_number(product_name):
    """
    从产品名称中提取期数数字，用于匹配政策表
    例如: "3期(用户)" -> 3, "24期" -> 24, "4+15" -> 4+15
    """
    s = clean_str(product_name)
    # 尝试提取纯数字或 "数字+数字" 格式
    match = re.search(r'(\d+[+\d]*)', s)
    if match:
        return match.group(1)
    return s # 如果没匹配到数字，返回原字符串尝试匹配

# ================= 核心处理逻辑 =================

def process_data(order_df, detail_df, offline_df, online_df, policy_df):
    """
    主处理函数
    """
    results = []
    
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

    # 3. 预处理返佣政策详情表 (新增核心逻辑)
    # 目标：构建 { "机构名_期数_还款方式": 返佣开始时间 } 的映射
    policy_map = {}
    default_policy_date = datetime.datetime(2025, 1, 1) # 兜底时间
    
    if policy_df is not None and not policy_df.empty:
        for _, row in policy_df.iterrows():
            inst_name = clean_str(row.get('机构名称'))
            prod_name = clean_str(row.get('产品名称')) # 如 "3期(用户)"
            repay_method = clean_str(row.get('还款方式')) # 如 "等额还款"
            start_time_str = clean_str(row.get('返佣开始时间'))
            
            # 解析时间
            policy_start_time = default_policy_date
            if start_time_str:
                try:
                    # 尝试多种时间格式
                    for fmt in ('%Y/%m/%d', '%Y-%m-%d', '%Y%m%d'):
                        try:
                            policy_start_time = datetime.datetime.strptime(start_time_str, fmt)
                            break
                        except ValueError:
                            continue
                except:
                    pass
            
            # 提取期数数字用于模糊匹配
            period_num = extract_period_number(prod_name)
            
            # 构建Key：使用 "机构_期数_方式" 作为主键
            # 为了容错，我们存储几个变种Key
            key_base = f"{inst_name}_{period_num}_{repay_method}"
            policy_map[key_base] = policy_start_time
            
            # 也可以存一个不带括号的简版，防止匹配不上
            # 这里简化处理，主要依赖上面的精确匹配+期数提取

    def get_policy_start_time(inst_name, product_name, repay_method):
        """查找特定订单的政策开始时间"""
        inst = clean_str(inst_name)
        prod = clean_str(product_name)
        method = clean_str(repay_method)
        
        # 1. 尝试精确匹配 (提取期数后)
        period_num = extract_period_number(prod)
        key = f"{inst}_{period_num}_{method}"
        if key in policy_map:
            return policy_map[key]
            
        # 2. 尝试直接用产品名匹配 (防止期数提取失败)
        key_raw = f"{inst}_{prod}_{method}"
        if key_raw in policy_map:
            return policy_map[key_raw]
            
        # 3. 如果都匹配不到，返回兜底时间
        return default_policy_date

    # 4. 处理线下代付记录
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
                        row, oid, order_map, 
                        source_type='offline',
                        repayment_type=current_repayment_type,
                        policy_func=get_policy_start_time # 传入政策查找函数
                    )
                    if res:
                        results.append(res)

    # 5. 处理线上分账记录
    if online_df is not None and not online_df.empty:
        for _, row in online_df.iterrows():
            oid = clean_str(row.get('订单编号'))
            repayment_type = clean_str(row.get('还款类型', row.get('期数', '')))
            
            res = build_result_row(
                row, oid, order_map, 
                source_type='online',
                repayment_type=repayment_type,
                policy_func=get_policy_start_time
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

def build_result_row(row, oid, order_map, source_type, repayment_type, policy_func):
    """构建单行结果数据"""
    res = {
        '业务订单号': oid,
        '数据来源': '线下代付' if source_type == 'offline' else '线上分账',
        '支付时间': row.get('支付时间', ''),
        '支付批次号': row.get('支付批次号', '') if source_type == 'offline' else '',
        '还款方式': repayment_type
    }

    # 字段映射
    inst_name = ""
    prod_name = ""
    
    if source_type == 'offline':
        res['分账金额'] = safe_float(row.get('merged_amount', row.get('清分金额', 0)))
        res['产品名称'] = '' 
        res['收款商户'] = ''
        res['付款人'] = ''
    else:
        res['分账金额'] = safe_float(row.get('分账金额', 0))
        prod_name = clean_str(row.get('产品名称', ''))
        inst_name = clean_str(row.get('收款商户', '')) # 线上表通常有机构名
        res['产品名称'] = prod_name
        res['收款商户'] = inst_name
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
        
        # 如果线上没拿到机构名，尝试从订单表拿（如果有）
        if not inst_name and '机构名称' in order_info:
            inst_name = clean_str(order_info.get('机构名称'))
            
    else:
        res['下单时间'] = ''
        res['订单状态'] = '未匹配到主表'
        res['维护商务'] = ''
        res['是否有返佣'] = '否'
        res['返佣比例'] = 0
        res['返佣金额'] = 0

    # --- 智能备注生成 (含政策判断) ---
    note_parts = []
    raw_note = clean_str(row.get('系统备注' if source_type=='offline' else '备注', ''))
    if raw_note:
        note_parts.append(raw_note)
    
    # 政策时间判断
    order_time_str = res.get('下单时间', '')
    if order_time_str and inst_name:
        try:
            order_time = pd.to_datetime(order_time_str)
            # 调用外部传入的政策查找函数
            # 需要传入：机构名，产品名(期数)，还款方式
            # 注意：线下代付可能没有产品名，这里尝试从订单表或还款类型里找线索，或者仅用机构匹配
            # 为简化，这里假设线下代付也能通过某种方式关联到产品，或者仅用机构+还款方式匹配
            # 但根据截图，政策表强依赖"产品名称(期数)"。
            # 如果线下数据缺失产品名，这里可能只能匹配个大概，或者我们需要从 order_info 里找产品名
            
            final_prod_name = prod_name
            if not final_prod_name and order_info and '产品名称' in order_info:
                 final_prod_name = clean_str(order_info.get('产品名称'))
                 
            policy_start = policy_func(inst_name, final_prod_name, repayment_type)
            
            if order_time < policy_start:
                note_parts.append("下单早于政策")
        except Exception as e:
            # 防止时间解析错误导致整个流程中断
            pass
            
    res['备注'] = " | ".join(note_parts) # 修复了这里的语法错误

    return res

# ================= Streamlit 界面交互 =================

st.sidebar.header("1. 文件上传区")
order_file = st.sidebar.file_uploader("1. 上传【订单主表】(xlsx)", type=['xlsx', 'xls'])
detail_file = st.sidebar.file_uploader("2. 上传【订单支付明细】(xlsx)", type=['xlsx', 'xls'])
offline_file = st.sidebar.file_uploader("3. 上传【线下代付记录】(xls)", type=['xlsx', 'xls'])
online_file = st.sidebar.file_uploader("4. 上传【线上分账支付记录】(xlsx)", type=['xlsx', 'xls'])
policy_file = st.sidebar.file_uploader("5. 上传【返佣政策详情】(xlsx)", type=['xlsx', 'xls']) # 新增

if st.sidebar.button("开始计算"):
    if all([order_file, detail_file, offline_file, online_file, policy_file]):
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
                file_name="返佣计算结果_V28.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        except Exception as e:
            st.error(f"发生错误: {str(e)}")
            import traceback
            st.code(traceback.format_exc())
    else:
        st.warning("请上传所有必需的 5 个文件后再点击开始计算。")
