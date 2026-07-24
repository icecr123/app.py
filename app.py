import streamlit as st
import pandas as pd
import numpy as np
import re
from io import BytesIO
import datetime

# ================= 页面配置 =================
st.set_page_config(page_title="月度回款返佣计算工具 V27-排序修复版", layout="wide")
st.title("🧮 月度回款返佣自动计算工具 (V27-排序修复版)")
st.markdown("""
**V27 核心修复说明：**
1. **修复排序报错**：解决因列名缺失或转换时机不对导致的 `KeyError: '支付时间_dt'`。
2. **政策动态匹配**：根据【返佣政策详情】表中的机构名称，动态匹配返佣开始时间。
3. **期次顺序消费**：线下代付严格按订单支付明细的时间顺序依次匹配期次。
4. **线上字段补全**：修复线上分账记录中关键字段为空的问题。
5. **罚息智能合并**：同批次下多行罚息自动合并至服务费行。
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

# ================= 核心处理逻辑 =================

def process_data(order_df, detail_df, offline_df, online_df, policy_df):
    """
    主处理函数：整合所有数据源
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
    # 结构：{订单编号: [list of repayment_types sorted by time]}
    detail_queue_map = {}
    if detail_df is not None and not detail_df.empty:
        # 【关键修复】先转换时间列，再分组
        time_col = '支付时间'
        if time_col in detail_df.columns:
            detail_df['支付时间_dt'] = pd.to_datetime(detail_df[time_col], errors='coerce')
            
            grouped_details = detail_df.groupby('订单编号')
            for oid, group in grouped_details:
                # 再次检查分组内是否有时间列
                if '支付时间_dt' in group.columns:
                    sorted_group = group.sort_values(by='支付时间_dt', ascending=True)
                    queue = sorted_group['还款类型'].tolist()
                    detail_queue_map[clean_str(oid)] = queue

    # 3. 预处理返佣政策详情表
    # 结构：{机构名称/收款商户: 返佣开始时间}
    policy_map = {}
    if policy_df is not None and not policy_df.empty:
        # 假设政策表包含 '机构名称' 和 '返佣开始时间' 列
        # 如果列名不同，请根据实际情况修改下方的列名
        inst_col = '机构名称' 
        date_col = '返佣开始时间'
        
        # 尝试兼容不同的列名
        if inst_col not in policy_df.columns:
            # 寻找包含"机构"或"商户"的列
            for c in policy_df.columns:
                if '机构' in c or '商户' in c:
                    inst_col = c
                    break
        
        if date_col not in policy_df.columns:
             for c in policy_df.columns:
                if '开始' in c or '生效' in c:
                    date_col = c
                    break

        for _, row in policy_df.iterrows():
            inst_name = clean_str(row.get(inst_col))
            start_date = row.get(date_col)
            if inst_name and pd.notna(start_date):
                try:
                    policy_map[inst_name] = pd.to_datetime(start_date)
                except:
                    pass

    # 4. 处理线下代付记录 (Offline Data)
    offline_counter = {} 
    
    if offline_df is not None and not offline_df.empty:
        # 4.1 过滤备注
        mask_include = offline_df['系统备注'].astype(str).str.contains(r'服务费|延期|罚息|逾期|违约金', na=False)
        mask_exclude = offline_df['系统备注'].astype(str).str.contains(r'本金|返服务费', na=False)
        filtered_offline = offline_df[mask_include & ~mask_exclude].copy()
        
        if not filtered_offline.empty:
            # 4.2 分组处理：按 业务订单号 + 支付批次号 分组
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
                
                # 4.3 生成结果行并匹配期次
                for row in final_rows:
                    current_repayment_type = get_next_repayment_type(oid, offline_counter, detail_queue_map)
                    
                    res = build_result_row(
                        row, oid, order_map, policy_map,
                        source_type='offline',
                        repayment_type=current_repayment_type
                    )
                    if res:
                        results.append(res)

    # 5. 处理线上分账记录 (Online Data)
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
    """
    获取下一个可用的还款期次（按顺序消费）
    """
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
    """
    构建单行结果数据
    """
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

    # 备注生成（含政策匹配）
    note_parts = []
    raw_note = clean_str(row.get('系统备注' if source_type=='offline' else '备注', ''))
    if raw_note:
        note_parts.append(raw_note)
    
    # 动态政策时间判断
    policy_start_date = None
    merchant_name = res.get('收款商户', '')
    
    # 如果线上有商户名，直接查；如果是线下，可能需要从订单表或其他地方找机构名
    # 这里假设线下代付的机构名可能藏在订单表的某个字段，或者暂时无法匹配
    # 如果线下也需要匹配政策，需要确保 order_info 里有机构名字段
    if not merchant_name and order_info is not None:
        # 尝试从订单表找机构名（假设列名包含'机构'或'商户'）
        for k, v in order_info.items():
            if ('机构' in k or '商户' in k) and pd.notna(v):
                merchant_name = clean_str(v)
                break

    if merchant_name and merchant_name in policy_map:
        policy_start_date = policy_map[merchant_name]
    
    # 如果没有匹配到特定机构，可以使用全局默认值（可选）
    # if policy_start_date is None:
    #     policy_start_date = datetime.datetime(2025, 1, 1)

    if policy_start_date:
        order_time_str = res.get('下单时间', '')
        if order_time_str:
            try:
                order_time = pd.to_datetime(order_time_str)
                # 确保比较的是日期部分，忽略时分秒差异
                if order_time.date() < policy_start_date.date():
                    note_parts.append("下单早于政策")
            except:
                pass
            
    res['备注'] | ".join(note_parts)

    return res

# ================= Streamlit 界面交互 =================

st.sidebar.header("1. 文件上传区")
order_file = st.sidebar.file_uploader("上传【订单主表】(xlsx)", type=['xlsx', 'xls'])
detail_file = st.sidebar.file_uploader("上传【订单支付明细】(xlsx)", type=['xlsx', 'xls'])
offline_file = st.sidebar.file_uploader("上传【线下代付记录】(xls)", type=['xlsx', 'xls'])
online_file = st.sidebar.file_uploader("上传【线上分账支付记录】(xlsx)", type=['xlsx', 'xls'])
policy_file = st.sidebar.file_uploader("上传【返佣政策详情】(xlsx)", type=['xlsx', 'xls'])

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
                file_name="返佣计算结果_V27.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        except Exception as e:
            st.error(f"发生错误: {str(e)}")
            import traceback
            st.code(traceback.format_exc())
    else:
        st.warning("请上传所有必需的 5 个文件后再点击开始计算。")
