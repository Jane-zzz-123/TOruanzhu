# -*- coding: utf-8 -*-
""" 三表勾稽核板（完整版 + 升级：全指标卡片 + 三表联动 + Excel导出下载）
功能：
   1. 上传Excel（三sheet：资产负债表/利润表/现金流量表）
   2. 全套财务KPI卡片：偿债、盈利、营运指标（红涨绿跌）
   3. 资产负债率滑块（步长5%），按现有负债结构比例自动缩放负债，未分配利润平衡报表
   4. 可编辑三表：input行可改，calc行只读，改完自动联动重算
   5. 【新增】开关：净利润自动同步资产负债表未分配利润（强勾稽）
   6. 【新增】一键下载调整后三表Excel，和上传模板表头一致
   7. 【保留】HTML负债模拟面板、JSON复制+回写解析
运行：streamlit run sanbiaohedui.py
"""
import pandas as pd
import streamlit as st
import json
import numpy as np
import math
from io import BytesIO

st.set_page_config(page_title="三表勾稽核板", layout="wide", page_icon="📊")

# ====================== 字段定义（严格对齐原表） ======================
BS_ROWS = [
    ("流动资产", "header"),
    ("货币资金", "input"),
    ("短期投资", "input"),
    ("应收票据", "input"),
    ("应收账款", "input"),
    ("预付款项", "input"),
    ("应收股利", "input"),
    ("应收利息", "input"),
    ("其他应收款", "input"),
    ("存货", "calc"),
    ("  原材料", "input"),
    ("  在产品", "input"),
    ("  库存商品", "input"),
    ("  周转材料", "input"),
    ("其他流动资产", "input"),
    ("流动资产合计", "calc"),
    ("非流动资产", "header"),
    ("长期债券投资", "input"),
    ("长期股权投资", "input"),
    ("固定资产原价", "input"),
    ("减：累计折旧", "input"),
    ("固定资产账面价值", "calc"),
    ("在建工程", "input"),
    ("工程物资", "input"),
    ("固定资产清理", "input"),
    ("生产性生物资产", "input"),
    ("无形资产", "input"),
    ("开发支出", "input"),
    ("长期待摊费用", "input"),
    ("其他非流动资产", "input"),
    ("非流动资产合计", "calc"),
    ("资产总计", "calc"),
    ("流动负债", "header"),
    ("短期借款", "input"),
    ("应付票据", "input"),
    ("应付账款", "input"),
    ("预收账款", "input"),
    ("应付职工薪酬", "input"),
    ("应交税费", "input"),
    ("应付利息", "input"),
    ("应付利润", "input"),
    ("其他应付款", "input"),
    ("其他流动负债", "input"),
    ("流动负债合计", "calc"),
    ("非流动负债", "header"),
    ("长期借款", "input"),
    ("长期应付款", "input"),
    ("递延收益", "input"),
    ("其他非流动负债", "input"),
    ("非流动负债合计", "calc"),
    ("负债合计", "calc"),
    ("所有者权益（或股东权益）", "header"),
    ("实收资本（或股本）", "input"),
    ("资本公积", "input"),
    ("盈余公积", "input"),
    ("未分配利润", "input"),
    ("所有者权益（或股东权益）合计", "calc"),
    ("负债和所有者权益（或股东权益）总计", "calc"),
]

PL_ROWS = [
    ("一、营业收入", "input"),
    ("减：营业成本", "input"),
    ("税金及附加", "input"),
    ("  其中：消费税", "input"),
    ("  营业税", "input"),
    ("  城市维护建设税", "input"),
    ("  资源税", "input"),
    ("  土地增值税", "input"),
    ("  城镇土地使用税、房产税、车船税、印花税", "input"),
    ("  教育费附加、矿产资源补偿费、排污费", "input"),
    ("销售费用", "input"),
    ("  其中：商品维修费", "input"),
    ("  广告费和业务宣传费", "input"),
    ("管理费用", "input"),
    ("  其中：开办费", "input"),
    ("  业务招待费", "input"),
    ("  研究费用", "input"),
    ("财务费用", "input"),
    ("  其中：利息费用（收入以\"-\"号填列）", "input"),
    ("加：投资收益（损失以\"-\"号填列）", "input"),
    ("二、营业利润（亏损以\"-\"号填列）", "calc"),
    ("加：营业外收入", "input"),
    ("  其中：政府补助", "input"),
    ("减：营业外支出", "input"),
    ("  其中：坏账损失", "input"),
    ("  无法收回的长期债券投资损失", "input"),
    ("  无法收回的长期股权投资损失", "input"),
    ("  自然灾害等不可抗力因素造成的损失", "input"),
    ("  税收滞纳金", "input"),
    ("三、利润总额（亏损总额以\"-\"号填列）", "calc"),
    ("减：所得税费用", "input"),
    ("四、净利润（净亏损以\"-\"号填列）", "calc"),
]

CF_ROWS = [
    ("一、经营活动产生的现金流量：", "header"),
    ("销售产成品、商品、提供劳务收到的现金", "input"),
    ("收到其他与经营活动有关的现金", "input"),
    ("购买原材料、商品、接受劳务支付的现金", "input"),
    ("支付的职工薪酬", "input"),
    ("支付的税费", "input"),
    ("支付其他与经营活动有关的现金", "input"),
    ("经营活动产生的现金流量净额", "calc"),
    ("二、投资活动产生的现金流量：", "header"),
    ("收回短期投资、长期债券投资和长期股权投资收到的现金", "input"),
    ("取得投资收益收到的现金", "input"),
    ("处置固定资产、无形资产和其他非流动资产收回的现金净额", "input"),
    ("短期投资、长期债券投资和长期股权投资支付的现金", "input"),
    ("购建固定资产、无形资产和其他非流动资产支付的现金", "input"),
    ("投资活动产生的现金流量净额", "calc"),
    ("三、筹资活动所产生的现金流量：", "header"),
    ("取得借款收到的现金", "input"),
    ("吸收投资者投资收到的现金", "input"),
    ("偿还借款本金支付的现金", "input"),
    ("偿还借款利息支付的现金", "input"),
    ("分配利润支付的现金", "input"),
    ("筹资活动产生的现金流量净额", "calc"),
    ("四、现金净增加额", "calc"),
    ("加：期初现金余额", "input"),
    ("五、期末现金余额", "calc"),
]

# 负债类input科目（用于按比例缩放）
LIABILITY_INPUT_ITEMS = [
    "短期借款", "应付票据", "应付账款", "预收账款", "应付职工薪酬",
    "应交税费", "应付利息", "应付利润", "其他应付款", "其他流动负债",
    "长期借款", "长期应付款", "递延收益", "其他非流动负债",
]

# ====================== 初始化 ======================
def empty_df(rows):
    return pd.DataFrame({"项目名称": [r[0] for r in rows],
                         "类型": [r[1] for r in rows],
                         "期初金额(元)": 0.0,
                         "期末金额(元)": 0.0})

if "bs" not in st.session_state:
    st.session_state.bs = empty_df(BS_ROWS)
if "pl" not in st.session_state:
    st.session_state.pl = empty_df(PL_ROWS)
if "cf" not in st.session_state:
    st.session_state.cf = empty_df(CF_ROWS)
if "target_debt_ratio" not in st.session_state:
    st.session_state.target_debt_ratio = 0.75
# ==========【新增状态：净利润联动未分配利润开关】==========
if "auto_sync_retained" not in st.session_state:
    st.session_state.auto_sync_retained = True

# ====================== 工具函数【加固防报错版本】 ======================
def getv(df, name, col="期末金额(元)"):
    try:
        if df is None or not isinstance(df, pd.DataFrame):
            return 0.0
        row = df.loc[df["项目名称"] == name, col]
        if len(row) == 0:
            return 0.0
        val = float(row.values[0])
        # 过滤inf / nan
        if np.isnan(val) or np.isinf(val):
            return 0.0
        return val
    except (ValueError, TypeError, KeyError):
        return 0.0

def setv(df, name, val, col="期末金额(元)"):
    try:
        val = float(val)
        if np.isnan(val) or np.isinf(val):
            val = 0.0
        df.loc[df["项目名称"] == name, col] = val
    except Exception:
        df.loc[df["项目名称"] == name, col] = 0.0

# ====================== 核心重算 ======================
def recalc():
    bs = st.session_state.bs.copy()
    pl = st.session_state.pl.copy()
    cf = st.session_state.cf.copy()

    for col in ["期初金额(元)", "期末金额(元)"]:
        # 资产负债表
        setv(bs, "存货",
             getv(bs, "  原材料", col)+getv(bs, "  在产品", col)+getv(bs, "  库存商品", col)+getv(bs, "  周转材料", col), col)
        setv(bs, "流动资产合计",
             getv(bs, "货币资金", col)+getv(bs, "短期投资", col)+getv(bs, "应收票据", col)+getv(bs, "应收账款", col)
             +getv(bs, "预付款项", col)+getv(bs, "应收股利", col)+getv(bs, "应收利息", col)+getv(bs, "其他应收款", col)
             +getv(bs, "存货", col)+getv(bs, "其他流动资产", col), col)
        setv(bs, "固定资产账面价值",
             getv(bs, "固定资产原价", col)-getv(bs, "减：累计折旧", col), col)
        setv(bs, "非流动资产合计",
             getv(bs, "长期债券投资", col)+getv(bs, "长期股权投资", col)+getv(bs, "固定资产账面价值", col)
             +getv(bs, "在建工程", col)+getv(bs, "工程物资", col)+getv(bs, "固定资产清理", col)
             +getv(bs, "生产性生物资产", col)+getv(bs, "无形资产", col)+getv(bs, "开发支出", col)
             +getv(bs, "长期待摊费用", col)+getv(bs, "其他非流动资产", col), col)
        setv(bs, "资产总计",
             getv(bs, "流动资产合计", col)+getv(bs, "非流动资产合计", col), col)
        setv(bs, "流动负债合计",
             getv(bs, "短期借款", col)+getv(bs, "应付票据", col)+getv(bs, "应付账款", col)+getv(bs, "预收账款", col)
             +getv(bs, "应付职工薪酬", col)+getv(bs, "应交税费", col)+getv(bs, "应付利息", col)
             +getv(bs, "应付利润", col)+getv(bs, "其他应付款", col)+getv(bs, "其他流动负债", col), col)
        setv(bs, "非流动负债合计",
             getv(bs, "长期借款", col)+getv(bs, "长期应付款", col)+getv(bs, "递延收益", col)+getv(bs, "其他非流动负债", col), col)
        setv(bs, "负债合计",
             getv(bs, "流动负债合计", col)+getv(bs, "非流动负债合计", col), col)
        setv(bs, "所有者权益（或股东权益）合计",
             getv(bs, "实收资本（或股本）", col)+getv(bs, "资本公积", col)+getv(bs, "盈余公积", col)+getv(bs, "未分配利润", col), col)
        setv(bs, "负债和所有者权益（或股东权益）总计",
             getv(bs, "负债合计", col)+getv(bs, "所有者权益（或股东权益）合计", col), col)

        # 利润表
        setv(pl, "二、营业利润（亏损以\"-\"号填列）",
             getv(pl, "一、营业收入", col)-getv(pl, "减：营业成本", col)-getv(pl, "税金及附加", col)
             -getv(pl, "销售费用", col)-getv(pl, "管理费用", col)-getv(pl, "财务费用", col)
             +getv(pl, "加：投资收益（损失以\"-\"号填列）", col), col)
        setv(pl, "三、利润总额（亏损总额以\"-\"号填列）",
             getv(pl, "二、营业利润（亏损以\"-\"号填列）", col)+getv(pl, "加：营业外收入", col)-getv(pl, "减：营业外支出", col), col)
        setv(pl, "四、净利润（净亏损以\"-\"号填列）",
             getv(pl, "三、利润总额（亏损总额以\"-\"号填列）", col)-getv(pl, "减：所得税费用", col), col)

        # =========【新增联动逻辑：净利润同步未分配利润，开关控制】=========
        if col == "期末金额(元)" and st.session_state.auto_sync_retained:
            net_p = getv(pl, "四、净利润（净亏损以\"-\"号填列）", col)
            setv(bs, "未分配利润", net_p, col)
            # 同步重算权益合计
            setv(bs, "所有者权益（或股东权益）合计",
                 getv(bs, "实收资本（或股本）", col)+getv(bs, "资本公积", col)+getv(bs, "盈余公积", col)+getv(bs, "未分配利润", col), col)
            setv(bs, "负债和所有者权益（或股东权益）总计",
                 getv(bs, "负债合计", col)+getv(bs, "所有者权益（或股东权益）合计", col), col)

        # 现金流量表
        setv(cf, "经营活动产生的现金流量净额",
             getv(cf, "销售产成品、商品、提供劳务收到的现金", col)+getv(cf, "收到其他与经营活动有关的现金", col)
             -getv(cf, "购买原材料、商品、接受劳务支付的现金", col)-getv(cf, "支付的职工薪酬", col)
             -getv(cf, "支付的税费", col)-getv(cf, "支付其他与经营活动有关的现金", col), col)
        setv(cf, "投资活动产生的现金流量净额",
             getv(cf, "收回短期投资、长期债券投资和长期股权投资收到的现金", col)
             +getv(cf, "取得投资收益收到的现金", col)
             +getv(cf, "处置固定资产、无形资产和其他非流动资产收回的现金净额", col)
             -getv(cf, "短期投资、长期债券投资和长期股权投资支付的现金", col)
             -getv(cf, "购建固定资产、无形资产和其他非流动资产支付的现金", col), col)
        setv(cf, "筹资活动产生的现金流量净额",
             getv(cf, "取得借款收到的现金", col)+getv(cf, "吸收投资者投资收到的现金", col)
             -getv(cf, "偿还借款本金支付的现金", col)-getv(cf, "偿还借款利息支付的现金", col)
             -getv(cf, "分配利润支付的现金", col), col)
        setv(cf, "四、现金净增加额",
             getv(cf, "经营活动产生的现金流量净额", col)+getv(cf, "投资活动产生的现金流量净额", col)
             +getv(cf, "筹资活动产生的现金流量净额", col), col)
        setv(cf, "五、期末现金余额",
             getv(cf, "加：期初现金余额", col)+getv(cf, "四、现金净增加额", col), col)

    st.session_state.bs = bs
    st.session_state.pl = pl
    st.session_state.cf = cf

# ====================== 按目标负债率缩放负债 ======================
def scale_liabilities_to_target():
    """保持资产总计不变，按现有负债结构比例缩放负债到目标负债率，差额用未分配利润平衡"""
    recalc()
    bs = st.session_state.bs.copy()
    col = "期末金额(元)"
    total_asset = getv(bs, "资产总计", col)
    if total_asset <= 0:
        st.warning("资产总计为0，请先填写资产数据。")
        return
    target_liab = total_asset * st.session_state.target_debt_ratio
    current_liab = getv(bs, "负债合计", col)
    if current_liab <= 0:
        st.warning("当前负债合计为0，无法按比例缩放。")
        return
    k = target_liab / current_liab
    for item in LIABILITY_INPUT_ITEMS:
        old = getv(bs, item, col)
        setv(bs, item, old * k, col)
    st.session_state.bs = bs
    recalc()
    # 平衡权益
    bs = st.session_state.bs.copy()
    new_liab = getv(bs, "负债合计", col)
    other_equity = getv(bs, "实收资本（或股本）", col)+getv(bs, "资本公积", col)+getv(bs, "盈余公积", col)
    setv(bs, "未分配利润", total_asset - new_liab - other_equity, col)
    st.session_state.bs = bs
    recalc()

# ====================== 上传解析 ======================
def parse_uploaded(file):
    try:
        xls = pd.ExcelFile(file)
        need = ["资产负债表", "利润表", "现金流量表"]
        miss = [s for s in need if s not in xls.sheet_names]
        if miss:
            st.error(f"缺少sheet：{miss}")
            return False
        bs_up = xls.parse("资产负债表")
        pl_up = xls.parse("利润表")
        cf_up = xls.parse("现金流量表")

        def map_up(up_df, rows):
            col_name = "项目名称" if "项目名称" in up_df.columns else up_df.columns[0]
            col_b = "期初金额(元)" if "期初金额(元)" in up_df.columns else up_df.columns[1]
            col_c = "期末金额(元)" if "期末金额(元)" in up_df.columns else up_df.columns[2]
            b_map = dict(zip(up_df[col_name].astype(str).str.strip(),
                             pd.to_numeric(up_df[col_b], errors="coerce").fillna(0.0)))
            c_map = dict(zip(up_df[col_name].astype(str).str.strip(),
                             pd.to_numeric(up_df[col_c], errors="coerce").fillna(0.0)))
            df = pd.DataFrame({"项目名称": [r[0] for r in rows],
                               "类型": [r[1] for r in rows]})
            df["期初金额(元)"] = df["项目名称"].map(b_map).fillna(0.0)
            df["期末金额(元)"] = df["项目名称"].map(c_map).fillna(0.0)
            return df

        st.session_state.bs = map_up(bs_up, BS_ROWS)
        st.session_state.pl = map_up(pl_up, PL_ROWS)
        st.session_state.cf = map_up(cf_up, CF_ROWS)
        recalc()
        st.success("✅ 上传成功，已自动重算全部勾稽项！")
        return True
    except Exception as e:
        st.error(f"解析失败：{e}")
        return False

# ======================【新增：导出三表Excel函数】======================
def export_three_sheet_excel():
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        st.session_state.bs.to_excel(writer, sheet_name="资产负债表", index=False)
        st.session_state.pl.to_excel(writer, sheet_name="利润表", index=False)
        st.session_state.cf.to_excel(writer, sheet_name="现金流量表", index=False)
    output.seek(0)
    return output

# ====================== 页面 ======================
st.title("📊 三表勾稽核板")

with st.sidebar:
    st.header("📤 数据上传")
    up = st.file_uploader("上传填好的Excel", type=["xlsx"])
    if up:
        parse_uploaded(up)

    st.divider()
    st.header("🎛️ 全局联动设置")
    st.session_state.auto_sync_retained = st.checkbox("利润表净利润自动同步未分配利润", value=st.session_state.auto_sync_retained, help="开启：修改利润表，净利润自动写入资产负债表未分配利润，严格勾稽；关闭：两张表独立，只做校验提示")

    st.divider()
    st.header("🎛️ 资产负债率参数")
    st.session_state.target_debt_ratio = st.slider(
        "目标资产负债率（步长5%）",
        min_value=0.10, max_value=0.95,
        value=st.session_state.target_debt_ratio,
        step=0.05, format="%.0f%%"
    )
    if st.button("🔧 按目标负债率自动调整（按现有负债比例缩放）", use_container_width=True):
        scale_liabilities_to_target()
        st.success(f"已按负债率={st.session_state.target_debt_ratio:.0%}调整完成！")

    st.divider()
    # =========【新增侧边栏下载按钮】=========
    excel_buf = export_three_sheet_excel()
    st.download_button(
        label="📥 下载调整后三表Excel",
        data=excel_buf,
        file_name="三表模拟结果.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

recalc()
bs = st.session_state.bs
pl = st.session_state.pl
cf = st.session_state.cf

col = "期末金额(元)"
total_asset = getv(bs, "资产总计", col)
total_liab = getv(bs, "负债合计", col)
total_equity = getv(bs, "所有者权益（或股东权益）合计", col)
curr_asset = getv(bs, "流动资产合计", col)
curr_liab = getv(bs, "流动负债合计", col)
inv = getv(bs, "存货", col)
receivable = getv(bs, "应收账款", col)
revenue = getv(pl, "一、营业收入", col)
cost = getv(pl, "减：营业成本", col)
net_profit = getv(pl, "四、净利润（净亏损以\"-\"号填列）", col)
interest_exp = getv(pl, "  其中：利息费用（收入以\"-\"号填列）", col)
ebit = getv(pl, "三、利润总额（亏损总额以\"-\"号填列）", col) + interest_exp

# 基础指标
debt_ratio = total_liab/total_asset if total_asset else 0
current_ratio = curr_asset/curr_liab if curr_liab else 0
quick_ratio = (curr_asset-inv)/curr_liab if curr_liab else 0
gross_margin = (revenue-cost)/revenue if revenue else 0
net_margin = net_profit/revenue if revenue else 0

# =========【新增全套拓展财务指标】=========
roe = net_profit / total_equity if total_equity else 0
roa = net_profit / total_asset if total_asset else 0
equity_ratio = total_liab / total_equity if total_equity else 0  #产权比率
inv_turn = cost / inv if inv else 0
receivable_turn = revenue / receivable if receivable else 0
asset_turn = revenue / total_asset if total_asset else 0
interest_cover = ebit / interest_exp if interest_exp !=0 else np.inf

st.subheader("📌 核心指标（期末口径）【偿债｜盈利｜营运】")
# 第一行 8卡
k1,k2,k3,k4 = st.columns(4)
k1.metric("资产总计", f"{total_asset:,.0f}")
k2.metric("负债合计", f"{total_liab:,.0f}")
k3.metric("资产负债率", f"{debt_ratio:.1%}", f"目标 {st.session_state.target_debt_ratio:.0%}")
k4.metric("净利润", f"{net_profit:,.0f}", "盈利" if net_profit>0 else "亏损")

k5,k6,k7,k8 = st.columns(4)
k5.metric("流动比率", f"{current_ratio:.2f}")
k6.metric("速动比率", f"{quick_ratio:.2f}")
k7.metric("毛利率", f"{gross_margin:.1%}")
k8.metric("净利率", f"{net_margin:.1%}")

# 第二行 新增指标卡片
k9,k10,k11,k12 = st.columns(4)
k9.metric("ROE", f"{roe:.1%}")
k10.metric("ROA", f"{roa:.1%}")
k11.metric("产权比率", f"{equity_ratio:.2f}")
k12.metric("利息保障倍数", f"{interest_cover:.2f}" if np.isfinite(interest_cover) else "∞")

k13,k14,k15,k16 = st.columns(4)
k13.metric("存货周转率", f"{inv_turn:.2f}")
k14.metric("应收周转率", f"{receivable_turn:.2f}")
k15.metric("总资产周转率", f"{asset_turn:.2f}")
k16.metric("所有者权益合计", f"{total_equity:,.0f}")

st.divider()

eq_diff = total_asset - getv(bs, "负债和所有者权益（或股东权益）总计", col)
cash_bs = getv(bs, "货币资金", col)
cash_cf = getv(cf, "五、期末现金余额", col)
c1,c2,c3 = st.columns(3)
with c1:
    if abs(eq_diff)<0.01:
        st.success("✅ 资产=负债+权益")
    else:
        st.error(f"❌ 资产≠负债+权益，差额{eq_diff:,.2f}")
with c2:
    if abs(cash_bs-cash_cf)<0.01:
        st.success("✅ 货币资金=期末现金余额")
    else:
        st.warning(f"⚠️ 货币资金{cash_bs:,.0f} vs 期末现金{cash_cf:,.0f}")
with c3:
    if net_profit>0:
        st.success(f"✅ 净利润>0（{net_profit:,.0f}）")
    else:
        st.error(f"❌ 净利润≤0（{net_profit:,.0f}）")

st.divider()

tab1, tab2, tab3 = st.tabs(["📋 资产负债表", "📈 利润表", "💰 现金流量表"])

def render_tab(tab, df, key):
    with tab:
        st.caption("白色行可编辑，黄色行自动计算（只读）")
        show = df[["项目名称","类型","期初金额(元)","期末金额(元)"]].copy()
        disabled_rows = (show["类型"].isin(["calc","header"])).tolist()
        edited = st.data_editor(
            show, key=key, hide_index=True, use_container_width=True,
            column_config={
                "项目名称": st.column_config.Column(disabled=True),
                "类型": st.column_config.Column(disabled=True),
                "期初金额(元)": st.column_config.NumberColumn(format="%.2f", step=1000.0),
                "期末金额(元)": st.column_config.NumberColumn(format="%.2f", step=1000.0),
            },
            disabled=disabled_rows,
        )
        for i, row in edited.iterrows():
            if row["类型"] == "input":
                df.loc[df["项目名称"]==row["项目名称"], "期初金额(元)"] = float(row["期初金额(元)"])
                df.loc[df["项目名称"]==row["项目名称"], "期末金额(元)"] = float(row["期末金额(元)"])
        recalc()

render_tab(tab1, st.session_state.bs, "editor_bs")
render_tab(tab2, st.session_state.pl, "editor_pl")
render_tab(tab3, st.session_state.cf, "editor_cf")

# ===================== 新增：资产负债率HTML模拟模块（复刻TACOS交互范式） =====================
st.markdown("## 📊 资产负债率自定义模拟（HTML前端实时测算）")
# session状态隔离，和原来scale函数互不干扰
if "debt_html_json_rows" not in st.session_state:
    st.session_state.debt_html_json_rows = None
if "html_target_ratio" not in st.session_state:
    st.session_state.html_target_ratio = None
if "debt_sim_result" not in st.session_state:
    st.session_state.debt_sim_result = None
if "debt_sim_summary" not in st.session_state:
    st.session_state.debt_sim_summary = None

# 读取当前最新资产负债表数据
bs_data = st.session_state.bs.copy()
# 取出资产总计
total_asset_val = getv(bs_data, "资产总计", col="期末金额(元)")

# 取出所有负债input科目（和LIABILITY_INPUT_ITEMS一致）
liab_items = LIABILITY_INPUT_ITEMS
liab_raw_list = []
sum_old_liab = 0.0
for item_name in liab_items:
    val = getv(bs_data, item_name, col="期末金额(元)")
    liab_raw_list.append({
        "item_name": item_name,
        "original_value": val
    })
    sum_old_liab += val

# 权益固定项：实收资本、资本公积、盈余公积（不变，只有未分配利润做平衡项）【✅这里修复了多余参数】
equity_fixed = [
    {"item_name":"实收资本（或股本）","value": getv(bs_data,"实收资本（或股本）","期末金额(元)")},
    {"item_name":"资本公积","value": getv(bs_data,"资本公积","期末金额(元)")},
    {"item_name":"盈余公积","value": getv(bs_data,"盈余公积","期末金额(元)")},
]
sum_equity_fixed = sum([x["value"] for x in equity_fixed])

st.subheader("⚙️ 第一步：设置目标资产负债率")
st.markdown("""
> 📘**模拟规则（资产总额保持固定不变）**
> 1. 目标总负债 = 资产总计 × 目标资产负债率
> 2. 负债基准分配：【按原有各个负债科目占原有总负债比例分摊目标总负债】
> 3. 权益：实收资本/资本公积/盈余公积固定不变，差额全部由【未分配利润】自动平衡
> 4. 支持手动修改任意负债科目金额，前端实时重算总负债、资产负债率、校验资产恒等式
> 💡微小差额属于计算机浮点精度误差。
""")
col1, col2 = st.columns([4,1])
with col1:
    target_ratio_input = st.number_input(
        "目标资产负债率(%)", min_value=0.0, max_value=95.0,
        value=75.0, step=5.0,
        help="资产总额保持不变；基准负债按原有负债结构比例分摊，可单独修改每个负债科目金额"
    )
with col2:
    reset_debt_sim = st.button("🔄 重置负债模拟", type="secondary")
    if reset_debt_sim:
        st.session_state.debt_html_json_rows = None
        st.session_state.html_target_ratio = None
        st.session_state.debt_sim_result = None
        st.session_state.debt_sim_summary = None
        st.rerun()

start_debt_calc = st.button("✅ 确认参数，生成负债模拟表格", type="primary")
if start_debt_calc:
    json_liab_rows = json.dumps(liab_raw_list, ensure_ascii=False)
    st.session_state.debt_html_json_rows = json_liab_rows
    st.session_state.html_target_ratio = str(target_ratio_input)

if st.session_state.debt_html_json_rows is not None:
    liab_json_str = st.session_state.debt_html_json_rows
    target_ratio_val = float(st.session_state.html_target_ratio)

    debt_html_template = '''
    <div style="font-size:13px;">
    <div style="display:flex;gap:15px;flex-wrap:wrap;margin-bottom:15px;padding:10px;background:#f6f8fa;border-radius:6px;">
      <div><b>设定目标负债率：</b><span id="targetRatioSpan" style="color:#165DFF;font-weight:bold;">__TR__%</span></div>
      <div><b>固定资产总额：</b>$<span id="assetTotalSpan" style="color:#00B42A;font-weight:bold;">0</span></div>
      <div><b>基准总负债：</b>$<span id="baseLiabTotal" style="color:#722ED1;font-weight:bold;">0</span></div>
      <div><b>修改后总负债：</b>$<span id="editLiabTotal" style="color:#722ED1;font-weight:bold;">0</span></div>
      <div><b>修改后实际负债率：</b><span id="realRatioSpan" style="color:#F53F3F;font-weight:bold;">0%</span></div>
      <div><b>资产恒等式差额：</b><span id="balanceDiffSpan" style="font-weight:bold;">0</span></div>
    </div>
    <div style="max-height:400px;overflow:auto;">
    <table border="1" cellpadding="4" cellspacing="0" style="width:100%;border-collapse:collapse;">
    <thead>
    <tr style="background:#e5e6eb;">
    <th>负债科目名称</th>
    <th>原始期末金额</th>
    <th>基准分摊金额【只读】</th>
    <th>修改金额【可编辑】</th>
    <th>金额差值(修改-基准)</th>
    </tr>
    </thead>
    <tbody id="liabTb"></tbody>
    </table>
    </div>
    <br />
    <button onclick="saveDebtData()" style="background:#165DFF;color:white;padding:6px 14px;border-radius:4px;border:none;">✅ 回传负债模拟结果（一键复制JSON）</button>
    <div id="debtTip" style="color:#00875a;display:none;margin:8px 0;">✅JSON已复制到剪贴板！粘贴到下方输入框，点击解析按钮即可</div>
    </div>
    <script>
    const liabRawData = JSON.parse(`__LIAB_JSON__`);
    const targetRatio = Number("__TR__") / 100;
    const assetTotal = Number("__A_TOTAL__");
    const sumEquityFixed = Number("__EQ_FIXED__");

    // 第一步：原始负债合计
    const originalSumLiab = liabRawData.reduce((s, r)=> s + Number(r.original_value), 0);
    // 目标负债总额
    const targetTotalLiab = assetTotal * targetRatio;

    // 基准分摊：按原始占比分配
    liabRawData.forEach(row=>{
        const orig = Number(row.original_value);
        let baseVal;
        if(originalSumLiab > 0){
            baseVal = targetTotalLiab * (orig / originalSumLiab);
        }else{
            baseVal = targetTotalLiab / liabRawData.length;
        }
        row.baseVal = baseVal;
        row.editVal = baseVal;
    })

    function recalcLiabTable(){
        const tbDom = document.getElementById("liabTb");
        tbDom.innerHTML = "";
        let sumEditLiab = 0;
        let sumBaseLiab = 0;

        liabRawData.forEach((row, idx)=>{
            const tr = document.createElement("tr");
            const origVal = Number(row.original_value);
            const baseVal = Number(row.baseVal);
            const editVal = Number(row.editVal);
            const diff = editVal - baseVal;
            sumBaseLiab += baseVal;
            sumEditLiab += editVal;

            let diffColor = "#000000";
            if(diff > 1) diffColor = "#c41e3a";
            if(diff < -1) diffColor = "#00875a";

            tr.innerHTML = `
                <td>${row.item_name}</td>
                <td>$${origVal.toFixed(2)}</td>
                <td>$${baseVal.toFixed(2)}</td>
                <td><input type="number" min=0 step=1 style="width:140px" value="${editVal.toFixed(2)}" onchange="updateLiabVal(${idx},this.value)"></td>
                <td style="color:${diffColor};font-weight:bold">${diff.toFixed(2)}</td>
            `
            tbDom.appendChild(tr);
        })
        // 平衡项：未分配利润
        const retainedProfit = assetTotal - sumEditLiab - sumEquityFixed;
        const balanceDiff = (sumEditLiab + sumEquityFixed + retainedProfit) - assetTotal;
        const realRatio = assetTotal>0 ? sumEditLiab / assetTotal : 0;

        document.getElementById("assetTotalSpan").innerText = assetTotal.toFixed(2);
        document.getElementById("baseLiabTotal").innerText = sumBaseLiab.toFixed(2);
        document.getElementById("editLiabTotal").innerText = sumEditLiab.toFixed(2);
        document.getElementById("realRatioSpan").innerText = (realRatio*100).toFixed(2)+"%";
        document.getElementById("balanceDiffSpan").innerText = balanceDiff.toFixed(2);
        if(Math.abs(balanceDiff)<0.01){
            document.getElementById("balanceDiffSpan").style.color="#00875a";
        }else{
            document.getElementById("balanceDiffSpan").style.color="#c41e3a";
        }
    }

    function updateLiabVal(idx, newVal){
        liabRawData[idx].editVal = Number(newVal);
        recalcLiabTable();
    }

    function round2(num){
        return Math.round(num*100)/100;
    }

    async function saveDebtData(){
        const finalSumEdit = liabRawData.reduce((s,r)=>s+Number(r.editVal),0);
        const retainedProfit = assetTotal - finalSumEdit - sumEquityFixed;
        const payload = {
            target_ratio_pct: Number("__TR__"),
            asset_total: assetTotal,
            sum_equity_fixed: sumEquityFixed,
            new_retained_profit: retainedProfit,
            liab_list: liabRawData.map(x=>{
                return {
                    item_name: x.item_name,
                    original_value: round2(x.original_value),
                    baseVal: round2(x.baseVal),
                    editVal: round2(x.editVal),
                    diff: round2(x.editVal - x.baseVal)
                }
            })
        }
        const jsonStr = JSON.stringify(payload);
        await navigator.clipboard.writeText(jsonStr);
        document.getElementById("debtTip").style.display="block";
        setTimeout(()=>{document.getElementById("debtTip").style.display="none"},3000)
    }
    recalcLiabTable();
    </script>
    '''
    # 替换占位符
    # 处理资产负债率，格式化+异常兜底
    if target_ratio_val is not None and isinstance(target_ratio_val, (int, float)) and not (
            math.isnan(target_ratio_val) or math.isinf(target_ratio_val)):
        tr_str = f"{target_ratio_val:.2f}"
    else:
        tr_str = "-"

    final_html = debt_html_template \
        .replace("__TR__", tr_str) \
        .replace("__LIAB_JSON__", liab_json_str) \
        .replace("__A_TOTAL__", str(total_asset_val)) \
        .replace("__EQ_FIXED__", str(sum_equity_fixed))

    st.components.v1.html(final_html, height=520, scrolling=True)

    st.info("👉 修改负债金额后，点击【回传负债模拟结果（一键复制JSON）】，复制JSON粘贴到下方输入框，点击解析回写资产负债表")
    debt_json_input = st.text_area("粘贴负债模拟JSON", height=120)
    parse_debt_btn = st.button("📥 解析负债模拟结果，写入资产负债表", type="primary")
    if parse_debt_btn:
        try:
            payload = json.loads(debt_json_input.strip())
            bs_update = st.session_state.bs.copy()
            # 回写所有负债科目
            for liab_row in payload["liab_list"]:
                setv(bs_update, liab_row["item_name"], liab_row["editVal"], col="期末金额(元)")
            # 回写平衡项：未分配利润
            setv(bs_update, "未分配利润", payload["new_retained_profit"], col="期末金额(元)")
            st.session_state.bs = bs_update
            recalc()
            st.success("✅ 解析成功！负债、未分配利润已更新到资产负债表，三表自动重算勾稽关系")
            st.rerun()
        except Exception as e:
            st.error(f"解析失败：{e}")

    if st.session_state.debt_sim_result is not None:
        st.divider()
        st.subheader("📊 本次负债率模拟汇总")
        st.dataframe(st.session_state.debt_sim_result, use_container_width=True)
