# -*- coding: utf-8 -*-
"""
三表勾稽核板（完整版）
功能：
  1. 上传Excel（三sheet：资产负债表/利润表/现金流量表）
  2. KPI卡片：资产总计、负债合计、资产负债率、净利润、毛利率、净利率、流动比率、速动比率
  3. 资产负债率滑块（步长5%），按现有负债结构比例自动缩放负债，未分配利润平衡报表
  4. 可编辑三表：input行可改，calc行只读，改完自动联动重算
运行：streamlit run app.py
"""
import pandas as pd
import streamlit as st

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

# ====================== 工具函数 ======================
def getv(df, name, col="期末金额(元)"):
    row = df.loc[df["项目名称"] == name, col]
    return float(row.values[0]) if len(row) else 0.0

def setv(df, name, val, col="期末金额(元)"):
    df.loc[df["项目名称"] == name, col] = float(val)

# ====================== 核心重算 ======================
def recalc():
    bs = st.session_state.bs.copy()
    pl = st.session_state.pl.copy()
    cf = st.session_state.cf.copy()

    for col in ["期初金额(元)", "期末金额(元)"]:
        # 资产负债表
        setv(bs, "存货",
             getv(bs,"  原材料",col)+getv(bs,"  在产品",col)+getv(bs,"  库存商品",col)+getv(bs,"  周转材料",col), col)
        setv(bs, "流动资产合计",
             getv(bs,"货币资金",col)+getv(bs,"短期投资",col)+getv(bs,"应收票据",col)+getv(bs,"应收账款",col)
             +getv(bs,"预付款项",col)+getv(bs,"应收股利",col)+getv(bs,"应收利息",col)+getv(bs,"其他应收款",col)
             +getv(bs,"存货",col)+getv(bs,"其他流动资产",col), col)
        setv(bs, "固定资产账面价值",
             getv(bs,"固定资产原价",col)-getv(bs,"减：累计折旧",col), col)
        setv(bs, "非流动资产合计",
             getv(bs,"长期债券投资",col)+getv(bs,"长期股权投资",col)+getv(bs,"固定资产账面价值",col)
             +getv(bs,"在建工程",col)+getv(bs,"工程物资",col)+getv(bs,"固定资产清理",col)
             +getv(bs,"生产性生物资产",col)+getv(bs,"无形资产",col)+getv(bs,"开发支出",col)
             +getv(bs,"长期待摊费用",col)+getv(bs,"其他非流动资产",col), col)
        setv(bs, "资产总计",
             getv(bs,"流动资产合计",col)+getv(bs,"非流动资产合计",col), col)
        setv(bs, "流动负债合计",
             getv(bs,"短期借款",col)+getv(bs,"应付票据",col)+getv(bs,"应付账款",col)+getv(bs,"预收账款",col)
             +getv(bs,"应付职工薪酬",col)+getv(bs,"应交税费",col)+getv(bs,"应付利息",col)
             +getv(bs,"应付利润",col)+getv(bs,"其他应付款",col)+getv(bs,"其他流动负债",col), col)
        setv(bs, "非流动负债合计",
             getv(bs,"长期借款",col)+getv(bs,"长期应付款",col)+getv(bs,"递延收益",col)+getv(bs,"其他非流动负债",col), col)
        setv(bs, "负债合计",
             getv(bs,"流动负债合计",col)+getv(bs,"非流动负债合计",col), col)
        setv(bs, "所有者权益（或股东权益）合计",
             getv(bs,"实收资本（或股本）",col)+getv(bs,"资本公积",col)+getv(bs,"盈余公积",col)+getv(bs,"未分配利润",col), col)
        setv(bs, "负债和所有者权益（或股东权益）总计",
             getv(bs,"负债合计",col)+getv(bs,"所有者权益（或股东权益）合计",col), col)

        # 利润表
        setv(pl, "二、营业利润（亏损以\"-\"号填列）",
             getv(pl,"一、营业收入",col)-getv(pl,"减：营业成本",col)-getv(pl,"税金及附加",col)
             -getv(pl,"销售费用",col)-getv(pl,"管理费用",col)-getv(pl,"财务费用",col)
             +getv(pl,"加：投资收益（损失以\"-\"号填列）",col), col)
        setv(pl, "三、利润总额（亏损总额以\"-\"号填列）",
             getv(pl,"二、营业利润（亏损以\"-\"号填列）",col)+getv(pl,"加：营业外收入",col)-getv(pl,"减：营业外支出",col), col)
        setv(pl, "四、净利润（净亏损以\"-\"号填列）",
             getv(pl,"三、利润总额（亏损总额以\"-\"号填列）",col)-getv(pl,"减：所得税费用",col), col)

        # 现金流量表
        setv(cf, "经营活动产生的现金流量净额",
             getv(cf,"销售产成品、商品、提供劳务收到的现金",col)+getv(cf,"收到其他与经营活动有关的现金",col)
             -getv(cf,"购买原材料、商品、接受劳务支付的现金",col)-getv(cf,"支付的职工薪酬",col)
             -getv(cf,"支付的税费",col)-getv(cf,"支付其他与经营活动有关的现金",col), col)
        setv(cf, "投资活动产生的现金流量净额",
             getv(cf,"收回短期投资、长期债券投资和长期股权投资收到的现金",col)
             +getv(cf,"取得投资收益收到的现金",col)
             +getv(cf,"处置固定资产、无形资产和其他非流动资产收回的现金净额",col)
             -getv(cf,"短期投资、长期债券投资和长期股权投资支付的现金",col)
             -getv(cf,"购建固定资产、无形资产和其他非流动资产支付的现金",col), col)
        setv(cf, "筹资活动产生的现金流量净额",
             getv(cf,"取得借款收到的现金",col)+getv(cf,"吸收投资者投资收到的现金",col)
             -getv(cf,"偿还借款本金支付的现金",col)-getv(cf,"偿还借款利息支付的现金",col)
             -getv(cf,"分配利润支付的现金",col), col)
        setv(cf, "四、现金净增加额",
             getv(cf,"经营活动产生的现金流量净额",col)+getv(cf,"投资活动产生的现金流量净额",col)
             +getv(cf,"筹资活动产生的现金流量净额",col), col)
        setv(cf, "五、期末现金余额",
             getv(cf,"加：期初现金余额",col)+getv(cf,"四、现金净增加额",col), col)

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
    other_equity = getv(bs,"实收资本（或股本）",col)+getv(bs,"资本公积",col)+getv(bs,"盈余公积",col)
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

# ====================== 页面 ======================
st.title("📊 三表勾稽核板")

with st.sidebar:
    st.header("📤 数据上传")
    up = st.file_uploader("上传填好的Excel", type=["xlsx"])
    if up:
        parse_uploaded(up)

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

recalc()
bs = st.session_state.bs
pl = st.session_state.pl
cf = st.session_state.cf

col = "期末金额(元)"
total_asset = getv(bs,"资产总计",col)
total_liab = getv(bs,"负债合计",col)
curr_asset = getv(bs,"流动资产合计",col)
curr_liab = getv(bs,"流动负债合计",col)
inv = getv(bs,"存货",col)
revenue = getv(pl,"一、营业收入",col)
cost = getv(pl,"减：营业成本",col)
net_profit = getv(pl,"四、净利润（净亏损以\"-\"号填列）",col)

debt_ratio = total_liab/total_asset if total_asset else 0
current_ratio = curr_asset/curr_liab if curr_liab else 0
quick_ratio = (curr_asset-inv)/curr_liab if curr_liab else 0
gross_margin = (revenue-cost)/revenue if revenue else 0
net_margin = net_profit/revenue if revenue else 0

st.subheader("📌 核心指标（期末口径）")
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

st.divider()

eq_diff = total_asset - getv(bs,"负债和所有者权益（或股东权益）总计",col)
cash_bs = getv(bs,"货币资金",col)
cash_cf = getv(cf,"五、期末现金余额",col)
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
