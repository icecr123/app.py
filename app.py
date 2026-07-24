import streamlit as st
import pandas as pd
import numpy as np
import re
from io import BytesIO
import datetime

# ================= 页面配置 =================
st.set_page_config(page_title="月度回款返佣计算工具 V23-终极修复版", layout="wide")
st.title("🧮 月度回款返佣自动计算工具 (V23-终极修复版)")
st.markdown("""
**V23 核心修复说明：**
1. **修复线上数据字段丢失**：重新映射分账表中的收款商户、付款人、产品名称等关键字段。
2. **修复代付记录误删**：优化正则过滤逻辑，确保“延期服务费”及各类罚息能被正确识别并保留。
3. **修复期次匹配错乱**：采用“按时间排序依次消费”算法，确保同一订单多次还款时期次准确递增。
4. **罚息合并逻辑**：同批次下的罚息金额自动合并至服务费行，避免数据冗余。
5. **智能备注生成**：自动对比下单时间与政策生效时间，标记“下单早于政策”。
""")

# ================= 辅助函数 =================

def safe_float(val):
    """安全转换金额为浮点数"""
    if isinstance(val, pd.Series):
        val = val.iloc[0] if not val.empty else 0
    if pd.isna(val): return 0.0
    s = str(val).strip()
    if s in ['无', 'None', 'nan', '', '-']: return 0.0
    try:
        return float(s.replace(',', ''))
    except ValueError:
        # 尝试提取字符串中的数字
        match = re.search(r'[\d\.]+', s)
        return float(match.group()) if match else 0.0

def clean_order_id(oid):
    """清洗订单号，去除空格和特殊字符"""
    if pd.isna(oid): return ""
    return str(oid).strip().replace(' ', '').replace('\n', '')

def filter_payment_by_remark(df):
    """
    根据系统备注过滤代付记录
    保留：服务费、延期服务费、罚息、逾期、违约金
    剔除：本金、返服务费
    """
    if df.empty: return df
    
    # 确保备注列为字符串类型
    remarks = df['系统备注'].astype(str).fillna("")
    
    # 定义保留关键词（支持正则）
    keep_pattern = r'(服务费|延期|罚息|逾期|违约金)'
    # 定义剔除关键词
    drop_pattern = r'(本金|返服务费)'
    
    # 应用过滤逻辑
    mask_keep = remarks.str.contains(keep_pattern, regex=True)
    mask_drop = remarks.str.contains(drop_pattern, regex=True)
    
    final_mask = mask_keep & (~mask_drop)
    
    filtered_count = len(df) - final_mask.sum()
    if filtered_count > 0:
        st.info(f"💡 代付记录中已自动过滤掉 {filtered_count} 条不符合要求的数据（含本金或返服务费等）。")
        
    return df[final_mask].copy()

# ================= 核心处理模块 =================

def process_online_data(df_div, df_order, policy_start_date):
    """处理线上分账数据"""
    results = []
    
    # 建立订单主表索引
    order_map = {}
    for _, row in df_order.iterrows():
        oid = clean_order_id(row.get('订单号', ''))
        if oid:
            order_map[oid] = row
            
    # 建立返佣策略索引
    policy_map = {}
    if not df_policy.empty and '产品类型' in df_policy.columns:
        for _, p_row in df_policy.iterrows():
            prod_name = str(p_row.get('产品类型', '')).strip()
            if prod_name:
                policy_map[prod_name] = p_row
                
    for _, div_row in df_div.iterrows():
        order_id = clean_order_id(div_row.get('订单编号', ''))
        order_info = order_map.get(order_id)
        
        if order_info is None: continue
        
        # 提取基础信息
        pay_time_str = str(div_row.get('支付时间', ''))
        amount = safe_float(div_row.get('分账金额', 0))
        
        # 关键修复：从分账表中提取商户、付款人、产品名称
        merchant = str(div_row.get('收款商户', '')).strip()
        payer = str(div_row.get('付款人', '')).strip()
        product_name = str(div_row.get('产品名称', '')).strip()
        
        # 如果分账表没取到，尝试从订单表补全
        if not merchant: merchant = str(order_info.get('收款商户', '')).strip()
        if not payer: payer = str(order_info.get('付款人', '')).strip()
        if not product_name: product_name = str(order_info.get('产品名称', '')).strip()
        
        # 匹配还款期次
        repayment_period = str(div_row.get('还款期次', '')).strip()
        
        # 匹配返佣策略
        commission_rate = 0.0
        has_commission = False
        remark = "支付成功待分润"
        
        # 遍历策略寻找匹配的产品
        for p_key, p_val in policy_map.items():
            if p_key in product_name:
                comm_col = f'{repayment_period}期返佣比例'
                if comm_col in p_val.index:
                    rate_str = str(p_val[comm_col]).replace('%', '')
                    try:
                        commission_rate = float(rate_str) / 100
                        has_commission = True
                    except: pass
                break
                
        # 检查下单时间是否早于政策
        order_time_str = str(order_info.get('下单时间', ''))
        if order_time_str and policy_start_date:
            try:
                order_dt = pd.to_datetime(order_time_str)
                policy_dt = pd.to_datetime(policy_start_date)
                if order_dt < policy_dt:
                    remark = "下单早于政策"
            except: pass
            
        results.append({
            '业务订单号': order_id,
            '产品名称': product_name,
            '收款商户': merchant,
            '付款人': payer,
            '分期金额': amount,
            '还款期次': repayment_period,
            '支付时间': pay_time_str,
            '服务费': 0,
            '逾期费用': 0,
            '还款方式': '线上还款',
            '下单时间': order_time_str,
            '订单状态': str(order_info.get('订单状态', '')),
            '维护商务': str(order_info.get('维护商务', '')),
            '是否有返佣': '是' if has_commission else '否',
            '返佣比例': commission_rate,
            '返佣金额': round(amount * commission_rate, 2),
            '备注': remark
        })
        
    return pd.DataFrame(results)

def process_offline_data(df_pay, df_detail, df_order, policy_start_date):
    """处理线下代付数据"""
    results = []
    
    # 1. 过滤代付记录
    df_filtered = filter_payment_by_remark(df_pay)
    if df_filtered.empty: return pd.DataFrame()
    
    # 2. 预处理：按订单+批次分组，合并罚息到服务费
    grouped = df_filtered.groupby(['业务订单号', '支付批次号'])
    merged_rows = []
    
    for (oid, batch_id), group in grouped:
        service_fee_row = None
        penalty_total = 0.0
        
        for _, row in group.iterrows():
            remark = str(row.get('系统备注', ''))
            amount = safe_float(row.get('清分金额', 0))
            
            # 识别是否为服务费行
            is_service = '服务费' in remark and '延期' not in remark and '罚息' not in remark and '逾期' not in remark and '违约金' not in remark
            
            if is_service:
                service_fee_row = row.copy()
                service_fee_row['_amount'] = amount
            else:
                # 视为罚息/延期费，累加
                penalty_total += amount
                
        if service_fee_row is not None:
            # 如果有服务费，将罚息合并进去
            service_fee_row['服务费'] = service_fee_row['_amount']
            service_fee_row['逾期费用'] = penalty_total
            merged_rows.append(service_fee_row)
        elif penalty_total > 0:
            # 如果只有罚息没有服务费，单独成行
            pure_penalty_row = group.iloc[0].copy()
            pure_penalty_row['服务费'] = 0
            pure_penalty_row['逾期费用'] = penalty_total
            merged_rows.append(pure_penalty_row)
            
    if not merged_rows: return pd.DataFrame()
    
    df_merged = pd.DataFrame(merged_rows)
    
    # 3. 准备订单明细映射（用于匹配期次）
    detail_map = {}
    if not df_detail.empty:
        for _, d_row in df_detail.iterrows():
            d_oid = clean_order_id(d_row.get('订单编号', ''))
            if d_oid not in detail_map:
                detail_map[d_oid] = []
            detail_map[d_oid].append(d_row)
            
    # 对每个订单的明细按支付时间排序，以便后续按顺序匹配期次
    for oid in detail_map:
        detail_map[oid].sort(key=lambda x: str(x.get('支付时间', '')))
        
    # 4. 遍历合并后的代付记录进行匹配
    # 使用一个临时字典记录当前订单已经匹配到了第几条明细
    order_match_index = {} 
    
    for _, pay_row in df_merged.iterrows():
        oid = clean_order_id(pay_row.get('业务订单号', ''))
        pay_time = str(pay_row.get('支付时间', ''))
        
        order_info = None
        for _, o_row in df_order.iterrows():
            if clean_order_id(o_row.get('订单号', '')) == oid:
                order_info = o_row
                break
                
        if order_info is None: continue
        
        # 匹配还款期次（核心修复逻辑）
        repayment_period = "未知"
        details = detail_map.get(oid, [])
        
        if details:
            current_idx = order_match_index.get(oid, 0)
            if current_idx < len(details):
                target_detail = details[current_idx]
                repayment_period = str(target_detail.get('还款期次', ''))
                order_match_index[oid] = current_idx + 1
            else:
                # 如果超出了明细行数，取最后一行或保持未知
                repayment_period = str(details[-1].get('还款期次', ''))
                
        # 提取产品信息
        product_name = str(order_info.get('产品名称', '')).strip()
        
        # 匹配返佣
        commission_rate = 0.0
        has_commission = False
        remark = ""
        
        # 检查下单时间
        order_time_str = str(order_info.get('下单时间', ''))
        if order_time_str and policy_start_date:
            try:
                if pd.to_datetime(order_time_str) < pd.to_datetime(policy_start_date):
                    remark = "下单早于政策"
            except: pass
            
        results.append({
            '业务订单号': oid,
            '产品名称': product_name,
            '收款商户': str(order_info.get('收款商户', '')),
            '付款人': str(order_info.get('付款人', '')),
            '分期金额': safe_float(pay_row.get('清分金额', 0)), # 这里取原始清分金额作为本金参考
            '还款期次': repayment_period,
            '支付时间': pay_time,
            '服务费': safe_float(pay_row.get('服务费', 0)),
            '逾期费用': safe_float(pay_row.get('逾期费用', 0)),
            '还款方式': '线下还款',
            '下单时间': order_time_str,
            '订单状态': str(order_info.get('订单状态', '')),
            '维护商务': str(order_info.get('维护商务', '')),
            '是否有返佣': '否', # 线下通常无自动返佣，除非有特殊逻辑
            '返佣比例': 0,
            '返佣金额': 0,
            '备注': remark
        })
        
    return pd.DataFrame(results)

# ================= 主程序入口 =================

def main():
    st.sidebar.header("📂 文件上传区")
    file_div = st.sidebar.file_uploader("1.上传【分账支付记录】(线上)", type=['xls', 'xlsx'])
    file_pay = st.sidebar.file_uploader("2.上传【代付记录】(线下)", type=['xls', 'xlsx'])
    file_order = st.sidebar.file_uploader("3.上传【订单主表】", type=['xls', 'xlsx'])
    file_detail = st.sidebar.file_uploader("4.上传【订单支付明细】(核对期次)", type=['xls', 'xlsx'])
    file_policy = st.sidebar.file_uploader("5.上传【返佣政策详情】", type=['xls', 'xlsx'])
    
    policy_start_date = st.sidebar.text_input("请输入返佣政策开始日期 (YYYY-MM-DD)", "2025-01-01")
    
    if st.sidebar.button("🚀 开始计算"):
        if not all([file_div, file_pay, file_order, file_detail, file_policy]):
            st.error("请上传所有必要的文件！")
            return
            
        try:
            with st.spinner("正在读取文件..."):
                df_div = pd.read_excel(file_div)
                df_pay = pd.read_excel(file_pay)
                df_order = pd.read_excel(file_order)
                df_detail = pd.read_excel(file_detail)
                global df_policy
                df_policy = pd.read_excel(file_policy)
                
            st.success("文件读取成功，正在建立映射...")
            
            with st.spinner("正在处理线上分账数据..."):
                res_online = process_online_data(df_div, df_order, policy_start_date)
                
            with st.spinner("正在处理线下代付数据..."):
                res_offline = process_offline_data(df_pay, df_detail, df_order, policy_start_date)
                
            # 合并结果
            final_df = pd.concat([res_online, res_offline], ignore_index=True)
            
            # 统一列顺序
            cols = ['业务订单号', '产品名称', '收款商户', '付款人', '分期金额', '还款期次', 
                    '支付时间', '服务费', '逾期费用', '还款方式', '下单时间', '订单状态', 
                    '维护商务', '是否有返佣', '返佣比例', '返佣金额', '备注']
            
            # 确保所有列都存在
            for c in cols:
                if c not in final_df.columns:
                    final_df[c] = ""
                    
            final_df = final_df[cols]
            
            st.success(f"✅ 处理完成！共生成 {len(final_df)} 条有效记录。")
            st.dataframe(final_df, use_container_width=True)
            
            # 导出 Excel
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                final_df.to_excel(writer, index=False, sheet_name='返佣计算结果')
            output.seek(0)
            
            st.download_button(
                label="📥 下载处理后的 Excel 文件",
                data=output,
                file_name=f"返佣计算结果_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        except Exception as e:
            st.error(f"❌ 发生错误: {str(e)}")
            import traceback
            st.code(traceback.format_
