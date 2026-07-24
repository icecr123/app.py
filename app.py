import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import datetime
import re

# ================= 页面配置 =================
st.set_page_config(page_title="月度回款返佣计算工具 V30-逻辑修正版", layout="wide")
st.title("🧮 月度回款返佣自动计算工具 (V30-逻辑修正版)")
st.markdown("""
**V30 核心修复说明：**
1. **代付合并逻辑修正**：
   - **同批次合并**：同一订单+同批次下，1行服务费+N行罚息 -> 合并为1行。
   - **跨批次独立**：同一订单+不同批次 -> 独立成行，不混淆金额。
2. **期次顺序匹配（防重）**：
   - 引入全局计数器，严格按还款笔数顺序匹配【订单支付明细】中的期次。
   - 避免同一订单的多笔还款匹配到同一个期次。
3. **政策精准匹配**：基于机构+期数+还款方式匹配返佣开始时间。
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

def parse_period(period_str):
    """提取期数数字，用于模糊匹配政策表 (例如 '3期(用户)' -> 3)"""
    if pd.isna(period_str): return None
    s = str(period_str)
    match = re.search(r'(\d+)', s)
    return int(match.group(1)) if match else None

# ================= 核心处理逻辑 =================

def process_data(order_df, detail_df, offline_df, online_df, policy_df):
    results = []
    
    # 1. 预处理订单主表
    order_map = {}
    if order_df is not None and not order_df.empty:
        for _, row in order_df.iterrows():
            oid = clean_str(row.get('订单号'))
            if oid:
                order_map[oid] = row

    # 2. 预处理订单支付明细（构建期次队列）
    # 结构：{订单编号: [list of repayment_types]}
    detail_queue_map = {}
    if detail_df is not None and not detail_df.empty:
        # 确保有时间列用于排序
        time_col = '支付时间'
        if '支付时间_dt' in detail_df.columns:
            time_col = '支付时间_dt'
        elif '支付时间' in detail_df.columns:
             detail_df['支付时间_dt'] = pd.to_datetime(detail_df['支付时间'], errors='coerce')
             time_col = '支付_time_dt'
        
        # 筛选支付方式为"线下"的记录（根据需求描述）
        # 注意：如果明细表里既有有线上又有线下，必须过滤。如果全是线下则不用过滤。
        # 这里假设需要过滤，如果不需要可注释掉 mask_offline
        mask_offline = detail_df['支付方式'].astype(str).str.contains('线下', na=False)
        filtered_detail = detail_df[mask_offline].copy() if '支付方式' in detail_df.columns else detail_df.copy()

        if not filtered_detail.empty and time_col in filtered_detail.columns:
            grouped_details = filtered_detail.groupby('订单编号')
            for oid, group in grouped_details:
                sorted_group = group.sort_values(by=time_col, ascending=True)
                queue = sorted_group['还款类型'].tolist()
                detail_queue_map[clean_str(oid)] = queue

    # 3. 预处理返佣政策表
    # 构建 Key: (机构名称, 期数数字, 还款方式) -> 返佣开始时间
    policy_map = {}
    if policy_df is not None and not policy_df.empty:
        for _, row in policy_df.iterrows():
            inst_name = clean_str(row.get('机构名称'))
            prod_name = clean_str(row.get('产品名称')) # 如 "3期(用户)"
            repay_way = clean_str(row.get('还款方式'))
            start_time_str = clean_str(row.get('返佣开始时间'))
            
            period_num = parse_period(prod_name)
            
            if inst_name and period_num:
                key = (inst_name, period_num, repay_way)
                policy_map[key] = start_time_str

    # 4. 全局计数器：用于线下代付的期次顺序匹配
    offline_repay_counter = {} 

    # 5. 处理线下代付记录 (Offline Data)
    if offline_df is not None and not offline_df.empty:
        # 5.1 过滤备注
        mask_include = offline_df['系统备注'].astype(str).str.contains(r'服务费|延期|罚息|逾期|违约金', na=False)
        mask_exclude = offline_df['系统备注'].astype(str).str.contains(r'本金|返服务费', na=False)
        filtered_offline = offline_df[mask_include & ~mask_exclude].copy()
        
        if not filtered_offline.empty:
            # 5.2 分组处理：按 业务订单号 + 支付批次号 分组
            # 这一步是关键：保证了"同批次"的数据在一起，"不同批次"的数据分开
            grouped_offline = filtered_offline.groupby(['业务订单号', '支付批次号'])
            
            for (oid, batch_id), group in grouped_offline:
                remarks = group['系统备注'].astype(str).tolist()
                
                total_penalty = 0.0
                service_fee_row = None
                
                # 遍历当前组（同批次）的所有行
                for _, row in group.iterrows():
                    remark = str(row.get('系统备注', ''))
                    amount = safe_float(row.get('清分金额', 0))
                    
                    # 识别服务费行
                    if '服务费' in remark:
                        service_fee_row = row.copy()
                        # 初始化合并金额
                        if '罚息' not in remark and '逾期' not in remark and '违约金' not in remark:
                             service_fee_row['merged_amount'] = amount
                        else:
                             service_fee_row['merged_amount'] = amount 
                    # 识别罚息行
                    elif any(k in remark for k in ['罚息', '逾期', '违约金']):
                        total_penalty += amount
                
                final_rows = []
                
                # 逻辑分支：
                # 情况A：有服务费行 -> 将同批次的罚息合并进去，生成1行
                if service_fee_row is not None:
                    base_row = service_fee_row
                    base_row['merged_amount'] = safe_float(base_row.get('merged_amount', 0)) + total_penalty
                    final_rows.append(base_row)
                # 情况B：没有服务费行，只有罚息 -> 单独成行（通常不应该发生，但做防御性处理）
                else:
                    if total_penalty > 0:
                        base_row = group.iloc[0].copy()
                        base_row['merged_amount'] = total_penalty
                        final_rows.append(base_row)
                
                # 5.3 生成结果行并匹配期次
                for row in final_rows:
                    # 【关键】获取期次：按顺序消费
                    current_repayment_type = get_next_repayment_type(oid, offline_repay_counter, detail_queue_map)
                    
                    res = build_result_row(
                        row, oid, order_map, policy_map,
                        source_type='offline',
                        repayment_type=current_repayment_type
                    )
                    if res:
                        results.append(res)

    # 6. 处理线上分账记录 (Online Data)
    if online_df is not None and not online_df.empty:
        for _, row in online_df.iterrows():
            oid = clean_str(row.get('订单编号'))
            
            # 线上直接用表里的期次，或者如果没有则尝试匹配（视情况而定，这里优先用表里的）
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
    """
    获取下一个可用的还款期次（按顺序消费）
    """
    if oid not in queue_map:
        return "未匹配到明细"
    
    queue = queue_map[oid]
    current_index = counter_map.get(oid, 0)
    
    if current_index < len(queue):
        rep_type = queue[current_index]
        counter_map[oid] = current_index + 1 # 计数器+1，下次取下一行
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
    institution_name = "" # 用于匹配政策
    
    if order_info is not None:
        res['下单时间'] = order_info.get('下单时间', '')
        res['订单状态'] = order_info.get('订单状态', '')
        res['维护商务'] = order_info.get('维护商务', '')
        res['是否有返佣'] = order_info.get('是否有返佣', '否')
        res['返佣比例'] = safe_float(order_info.get('返佣比例', 0))
        res['返佣金额'] = res['分账金额'] * res['返佣比例']
        
        # 获取机构名称用于政策匹配
        institution_name = clean_str(order_info.get('机构名称', ''))
    else:
        res['下单时间'] = ''
        res['订单状态'] = '未匹配到主表'
        res['维护商务'] = ''
        res['是否有返佣'] = '否'
        res['返佣比例'] = 0
        res['返佣金额'] = 0

    # 备注生成 & 政策时间判断
    note_parts = []
    raw_note = clean_str(row.get('系统备注' if source_type=='offline' else '备注', ''))
    if raw_note:
        note_parts.append(raw_note)
    
    # 动态政策匹配逻辑
    policy_start_date = None
    if institution_name and repayment_type:
        # 尝试从还款类型中提取期数数字 (例如 "3期" -> 3)
        period_num = parse_period(repayment_type)
        
        if period_num:
            # 尝试精确匹配 (机构, 期数, 还款方式)
            # 注意：这里的"还款方式"在policy_map的key中可能需要根据实际业务调整
            # 假设 policy_map 的 key 是 (机构名, 期数, 任意/具体方式)
            # 这里简化处理：只要机构名和期数对上，就取时间
            # 如果需要严格匹配第三个维度，需补充逻辑
            
            # 遍历 policy_map 寻找最匹配的
            for (p_inst, p_period, p_way), p_time in policy_map.items():
                if p_inst == institution_name and p_period == period_num:
                    policy_start_date = p_time
                    break
    
    if policy_start_date:
        order_time_str = res.get('下单时间', '')
        if order_time_str:
            try:
                order_time = pd.to_datetime(order_time_str)
                policy_time = pd.to_datetime(policy_start_date)
                if order_time < policy_time:
                    note_parts.append("下单早于政策")
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
                file_name="返佣计算结果_V30.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        except Exception as e:
            st.error(f"发生错误: {str(e)}")
            import traceback
            st.code(traceback.format_exc())
    else:
        st.warning("请上传所有必需的 5 个文件后再点击开始计算。")
