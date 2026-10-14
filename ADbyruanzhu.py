import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import requests
from io import BytesIO
import numpy as np

# ===================== 用户权限配置 =====================
USER_AUTH = {
    # 超级管理员（可查看全部店铺）
    "admin": {
        "pwd": "@admin1234",
        "shops": ["全部"]
    },
    "黄怡": {
        "pwd": "syc-huangyi123",
        "shops": ["思业成-US","创奇-US黄怡D"]
    },
    "镱娜": {
        "pwd": "dx-yyyna123",
        "shops": ["定行-US","创奇-US镱娜H"]
    },
    "小娇": {
        "pwd": "pt and ys-xiaojiao",
        "shops": ["拼途-US","艺胜-US","创奇-US小娇C"]
    },
    "楷纯": {
        "pwd": "zy and cr-kaichun",
        "shops": ["争艳-US","辰瑞-US","创奇-US楷纯E"]
    },
    "淑谊": {
        "pwd": "sx and jy-shuyi",
        "shops": ["势兴-US","进益-US","创奇-US淑谊F"]
    },
    "佰英": {
        "pwd": "cq-baiying123",
        "shops": ["创奇-US"]
    },
    "李珊": {
        "pwd": "dm-lishan123",
        "shops": ["大卖-US","创奇-US李珊G"]
    },
    "雨宜": {
        "pwd": "cq-uuuyi123",
        "shops": ["创奇-US雨宜B"]
    },
}

# session初始化
if "login_ok" not in st.session_state:
    st.session_state.login_ok = False
if "login_username" not in st.session_state:
    st.session_state.login_username = ""
if "allow_shops" not in st.session_state:
    st.session_state.allow_shops = []


def login_page():
    st.markdown("# 🔐 广告分析看板 - 用户登录")
    st.divider()
    user_name_list = list(USER_AUTH.keys())

    with st.form("login_form"):
        # ✅ 用户名下拉单选
        select_user = st.selectbox("请选择用户名", options=user_name_list)
        input_pwd = st.text_input("请输入密码", type="password")
        submit_login = st.form_submit_button("登录", type="primary")

        if submit_login:
            user_info = USER_AUTH[select_user]
            if input_pwd == user_info["pwd"]:
                st.session_state.login_ok = True
                st.session_state.login_username = select_user
                st.session_state.allow_shops = user_info["shops"]
                st.success("✅登录成功，正在加载看板...")
                st.rerun()
            else:
                st.error("❌密码错误，请重新输入")
    st.stop()


# 未登录展示登录页，终止后续执行
if not st.session_state.login_ok:
    login_page()

# -------------------------- 页面基础配置 --------------------------
st.set_page_config(
    page_title="广告分析看板",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ==========页面最顶部：当前用户 + 右上角退出登录按钮（不在侧边栏）==========
top_col1, top_col2 = st.columns([0.92, 0.08])
with top_col1:
    st.markdown(f"👤 当前登录用户：**{st.session_state.login_username}**")
with top_col2:
    logout_btn = st.button("🚪退出登录")
    if logout_btn:
        st.session_state.login_ok = False
        st.session_state.login_username = ""
        st.session_state.allow_shops = []
        st.rerun()

st.divider()
st.title("📊 广告分析看板")

# 全局TACOS模拟参数初始化
if "global_target_tacos_pct" not in st.session_state:
    st.session_state.global_target_tacos_pct = 15.0

# 侧边栏：只保留预算模拟参数，无登录相关
with st.sidebar:
    st.markdown("### ⚙️预算模拟参数")
    st.session_state.global_target_tacos_pct = st.number_input(
        "店铺全局目标TACOS(%)", min_value=0.0, max_value=50.0,
        value=st.session_state.global_target_tacos_pct, step=0.5,
        help="基准测算使用该值；SKU填写自定义后模拟测算会优先使用单品值"
    )
GLOBAL_TARGET_TACOS = st.session_state.global_target_tacos_pct / 100




# -------------------------- 缓存加载原始数据 --------------------------
@st.cache_data
def load_raw_data():
    url = "https://github.com/Jane-zzz-123/---/raw/main/ADdata_all.xlsx"
    resp = requests.get(url)
    df = pd.read_excel(BytesIO(resp.content), sheet_name="源数据")

    # 标准化时间
    df["时间"] = pd.to_datetime(df["时间"])
    df["年月"] = df["时间"].dt.to_period("M").astype(str)
    df["年月日期"] = pd.to_datetime(df["年月"] + "-01")

    # 数值清洗
    num_cols = [
        "展示", "点击", "广告花费", "SP广告费", "SB广告费", "SBV广告费",
        "广告销售额", "SP广告销售额", "SB广告销售额", "SBV广告销售额",
        "广告订单量", "SP广告订单量", "SB广告订单量", "SBV广告订单量",
        "销量", "销售额", "订单量"
    ]
    for col in num_cols:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    # 上架时间转换
    df["开售时间"] = pd.to_datetime(df["开售时间"], errors="coerce")
    # 校验产品类型是否存在，不存在抛提示
    if "产品类型" not in df.columns:
        st.error("❌ 远程Excel未包含【产品类型】字段，请先上传更新后的文件到GitHub！")
    return df


df_raw = load_raw_data()

# ===================== 页面顶部筛选区：仅2个单选控件【店铺、单月年月】 =====================
st.markdown("### 🔍 数据筛选条件（仅单月数据）")
filter_shop_col, filter_month_col = st.columns([1, 2])

# 1、店铺单选下拉【权限过滤】
with filter_shop_col:
    all_shop_list = sorted(df_raw["店铺"].unique())
    user_allow_shops = st.session_state.allow_shops
    if "全部" in user_allow_shops:
        filter_shop_list = all_shop_list
    else:
        filter_shop_list = [s for s in all_shop_list if s in user_allow_shops]

    if len(filter_shop_list) == 0:
        st.error("⚠️该账号没有可访问店铺，请联系管理员配置权限！")
        st.stop()
    select_shop = st.selectbox("选择分析店铺", filter_shop_list)

# 2、年月单选下拉，默认最新月份
with filter_month_col:
    month_list = sorted(df_raw["年月"].unique())
    latest_month = month_list[-1]
    select_month = st.selectbox("选择分析单月", month_list, index=month_list.index(latest_month))

st.divider()

# ===================== 新增：指标释义&计算公式折叠面板（放在筛选下方、当月卡片上方） =====================
with st.expander("📖 全部指标释义 & 计算公式（点击展开查看）", expanded=False):
    st.markdown("""
### 一、基础金额/订单指标
1. **当月总广告花费**
    - 含义：当月SP+SB+SBV全部广告扣费总和
    - 公式：广告花费 = SP广告费 + SB广告费 + SBV广告费
2. **当月广告销售额**
    - 含义：当月所有广告渠道点击成交的订单总金额（仅广告单营收）
    - 公式：广告销售额 = SP广告销售额 + SB广告销售额 + SBV广告销售额
3. **当月全店总销售额**
    - 含义：店铺当月所有订单（广告单+自然单）总成交金额
    - 公式：全店销售额 = 全部订单对应销售额总和
4. **当月广告订单总数**
    - 含义：当月通过广告点击成交的订单数量
    - 公式：广告订单数 = 所有广告渠道订单数量求和
5. **当月全店总订单**
    - 含义：店铺当月全部成交订单（广告+自然流量）
    - 公式：总订单 = 广告订单 + 自然流量订单

### 二、广告成本核心指标
6. **TACOS 广告花费占比**
    - 含义：广告花费占全店总销售额的比例，衡量整体广告投入力度
    - 公式：TACOS = 当月广告花费 ÷ 当月全店总销售额
7. **ACOS 广告销售成本**
    - 含义：广告花费占广告自身销售额的比例，广告投放内部转化成本
    - 公式：ACOS = 当月广告花费 ÷ 当月广告销售额
8. **ROAS 广告投产比**
    - 含义：每1美金广告花费带来多少广告销售额，反向ACOS
    - 公式：ROAS = 当月广告销售额 ÷ 当月广告花费
9. **ASoAS 广告销售依赖度**
    - 含义：广告成交销售额占店铺全店总销售额的比重，店铺流量依赖度
    - 公式：ASoAS = 当月广告销售额 ÷ 当月全店总销售额

### 三、流量效率指标
10. **CPC 单次点击成本**
    - 含义：广告平均单次点击扣费
    - 公式：CPC = 当月广告总花费 ÷ 当月广告总点击量
11. **CTR 点击率**
    - 含义：广告曝光后产生点击的比例，主图/标题吸引力
    - 公式：CTR = 当月广告点击量 ÷ 当月广告曝光（展示）量
12. **CVR 广告转化率**
    - 含义：广告点击后下单成交比例，落地页/价格/评价转化能力
    - 公式：CVR = 当月广告订单数 ÷ 当月广告点击量

### 补充说明
- 环比差值：当月指标 − 上月同期指标
- 环比百分比：(当月 − 上月) ÷ 上月绝对值，正数上涨（红色），负数下降（绿色）
- 新品判定：MSKU上架≤90天=新品；无开售时间/上架超90天=老品
""")
st.divider()

# -------------------------- 筛选当月数据集 --------------------------
df_filter_single_month = df_raw[
    (df_raw["店铺"] == select_shop) &
    (df_raw["年月"] == select_month)
    ].copy()

if df_filter_single_month.empty:
    st.warning(f"⚠️ {select_shop}店铺 {select_month} 无任何数据，请更换店铺或月份！")
    st.stop()

# -------------------------- 当月单月聚合 --------------------------
df_month_single = df_filter_single_month.groupby("年月").agg({
    "展示": "sum",
    "点击": "sum",
    "广告花费": "sum",
    "SP广告费": "sum",
    "SB广告费": "sum",
    "SBV广告费": "sum",
    "广告销售额": "sum",
    "广告订单量": "sum",
    "销售额": "sum",
    "订单量": "sum"
}).reset_index()

# 当月衍生指标
df_month_single["CTR"] = np.where(df_month_single["展示"] == 0, 0, df_month_single["点击"] / df_month_single["展示"])
df_month_single["CPC"] = np.where(df_month_single["点击"] == 0, 0,
                                  df_month_single["广告花费"] / df_month_single["点击"])
df_month_single["CVR广告转化率"] = np.where(df_month_single["点击"] == 0, 0,
                                            df_month_single["广告订单量"] / df_month_single["点击"])
df_month_single["ACOS"] = np.where(df_month_single["广告销售额"] == 0, 0,
                                   df_month_single["广告花费"] / df_month_single["广告销售额"])
df_month_single["ROAS"] = np.where(df_month_single["广告花费"] == 0, 0,
                                   df_month_single["广告销售额"] / df_month_single["广告花费"])
df_month_single["TACOS广告花费占比"] = np.where(df_month_single["销售额"] == 0, 0,
                                                df_month_single["广告花费"] / df_month_single["销售额"])
df_month_single["ASoAS广告销售依赖度"] = np.where(df_month_single["销售额"] == 0, 0,
                                                  df_month_single["广告销售额"] / df_month_single["销售额"])
df_month_single["SP广告花费占比"] = np.where(df_month_single["广告花费"] == 0, 0,
                                             df_month_single["SP广告费"] / df_month_single["广告花费"])
df_month_single["SB广告花费占比"] = np.where(df_month_single["广告花费"] == 0, 0,
                                             df_month_single["SB广告费"] / df_month_single["广告花费"])
df_month_single["SBV广告花费占比"] = np.where(df_month_single["广告花费"] == 0, 0,
                                              df_month_single["SBV广告费"] / df_month_single["广告花费"])

# -------------------------- 当月数值提取（新增广告销售额）
curr_ad_spend = df_month_single["广告花费"].iloc[0]
curr_sales = df_month_single["销售额"].iloc[0]
curr_ad_sales = df_month_single["广告销售额"].iloc[0] # 新增当月广告销售额
curr_ad_order = df_month_single["广告订单量"].iloc[0]
curr_all_order = df_month_single["订单量"].iloc[0]
curr_tacos = df_month_single["TACOS广告花费占比"].iloc[0]
curr_acos = df_month_single["ACOS"].iloc[0]
curr_roas = df_month_single["ROAS"].iloc[0]
curr_asoas = df_month_single["ASoAS广告销售依赖度"].iloc[0]
curr_cpc = df_month_single["CPC"].iloc[0]
curr_ctr = df_month_single["CTR"].iloc[0]
curr_cvr = df_month_single["CVR广告转化率"].iloc[0]

# -------------------------- 计算上月环比数据（补充广告销售额环比）
curr_period = pd.Period(select_month, freq="M")
last_period = curr_period - 1
last_month_str = str(last_period)
df_last_raw = df_raw[(df_raw["店铺"] == select_shop) & (df_raw["年月"] == last_month_str)]
has_last_month = not df_last_raw.empty

# 上月默认0 新增广告销售额变量
last_ad_spend = last_sales = last_ad_sales = last_ad_order = last_all_order = 0
last_tacos = last_acos = last_roas = last_asoas = last_cpc = last_ctr = last_cvr = 0
delta_ad_spend = delta_sales = delta_ad_sales = delta_ad_order = delta_all_order = 0
delta_tacos = delta_acos = delta_roas = delta_asoas = delta_cpc = delta_ctr = delta_cvr = 0
pct_ad_spend = pct_sales = pct_ad_sales = pct_ad_order = pct_all_order = 0
pct_tacos = pct_acos = pct_roas = pct_asoas = pct_cpc = pct_ctr = pct_cvr = 0

if has_last_month:
    df_last_agg = df_last_raw.groupby("年月").agg({
        "广告花费":"sum","销售额":"sum","广告订单量":"sum","订单量":"sum",
        "展示":"sum","点击":"sum","广告销售额":"sum"
    }).reset_index()
    last_ad_spend = df_last_agg["广告花费"].iloc[0]
    last_sales = df_last_agg["销售额"].iloc[0]
    last_ad_sales = df_last_agg["广告销售额"].iloc[0] # 上月广告销售额
    last_ad_order = df_last_agg["广告订单量"].iloc[0]
    last_all_order = df_last_agg["订单量"].iloc[0]
    last_imp = df_last_agg["展示"].iloc[0]
    last_click = df_last_agg["点击"].iloc[0]
    last_ad_sales_base = df_last_agg["广告销售额"].iloc[0]

    last_ctr = last_click / last_imp if last_imp !=0 else 0
    last_cpc = last_ad_spend / last_click if last_click !=0 else 0
    last_cvr = last_ad_order / last_click if last_click !=0 else 0
    last_acos = last_ad_spend / last_ad_sales_base if last_ad_sales_base !=0 else 0
    last_roas = last_ad_sales_base / last_ad_spend if last_ad_spend !=0 else 0
    last_tacos = last_ad_spend / last_sales if last_sales !=0 else 0
    last_asoas = last_ad_sales_base / last_sales if last_sales !=0 else 0

    # 广告销售额差值&环比
    delta_ad_spend = curr_ad_spend - last_ad_spend
    delta_sales = curr_sales - last_sales
    delta_ad_sales = curr_ad_sales - last_ad_sales
    delta_ad_order = curr_ad_order - last_ad_order
    delta_all_order = curr_all_order - last_all_order
    delta_tacos = curr_tacos - last_tacos
    delta_acos = curr_acos - last_acos
    delta_roas = curr_roas - last_roas
    delta_asoas = curr_asoas - last_asoas
    delta_cpc = curr_cpc - last_cpc
    delta_ctr = curr_ctr - last_ctr
    delta_cvr = curr_cvr - last_cvr

    pct_ad_spend = delta_ad_spend / last_ad_spend if last_ad_spend !=0 else 0
    pct_sales = delta_sales / last_sales if last_sales !=0 else 0
    pct_ad_sales = delta_ad_sales / last_ad_sales if last_ad_sales !=0 else 0
    pct_ad_order = delta_ad_order / last_ad_order if last_ad_order !=0 else 0
    pct_all_order = delta_all_order / last_all_order if last_all_order !=0 else 0
    pct_tacos = delta_tacos / last_tacos if last_tacos !=0 else 0
    pct_acos = delta_acos / last_acos if last_acos !=0 else 0
    pct_roas = delta_roas / last_roas if last_roas !=0 else 0
    pct_asoas = delta_asoas / last_asoas if last_asoas !=0 else 0
    pct_cpc = delta_cpc / last_cpc if last_cpc !=0 else 0
    pct_ctr = delta_ctr / last_ctr if last_ctr !=0 else 0
    pct_cvr = delta_cvr / last_cvr if last_cvr !=0 else 0

# ===================== 一、当月指标卡片（第一行新增广告销售额，5列布局） =====================
st.markdown(f"## 🎯 一、{select_month} 当月店铺整体概况（单月快照）")
if not has_last_month:
    st.info("当前为最早统计月份，无上月对比数据")

# 第一行：5列 广告花费、总销售额、广告销售额、广告订单、总订单
row1_col1, row1_col2, row1_col3, row1_col4, row1_col5 = st.columns(5)
with row1_col1:
    st.metric(label="当月总广告花费", value=f"${curr_ad_spend:,.2f}", delta=f"{delta_ad_spend:,.2f}")
    if has_last_month:
        st.caption(f"上月：${last_ad_spend:,.2f} | 环比：{pct_ad_spend:.1%}")
with row1_col2:
    st.metric(label="当月全店总销售额", value=f"${curr_sales:,.2f}", delta=f"{delta_sales:,.2f}")
    if has_last_month:
        st.caption(f"上月：${last_sales:,.2f} | 环比：{pct_sales:.1%}")
with row1_col3:
    st.metric(label="当月广告销售额", value=f"${curr_ad_sales:,.2f}", delta=f"{delta_ad_sales:,.2f}")
    if has_last_month:
        st.caption(f"上月：${last_ad_sales:,.2f} | 环比：{pct_ad_sales:.1%}")
with row1_col4:
    st.metric(label="当月广告订单总数", value=f"{curr_ad_order:,.0f}", delta=f"{delta_ad_order:,.0f}")
    if has_last_month:
        st.caption(f"上月：{last_ad_order:,.0f} | 环比：{pct_ad_order:.1%}")
with row1_col5:
    st.metric(label="当月全店总订单", value=f"{curr_all_order:,.0f}", delta=f"{delta_all_order:,.0f}")
    if has_last_month:
        st.caption(f"上月：{last_all_order:,.0f} | 环比：{pct_all_order:.1%}")

# 第二行4卡：TACOS、ACOS、ROAS、ASoAS（不变）
row2_col1, row2_col2, row2_col3, row2_col4 = st.columns(4)
with row2_col1:
    st.metric(label="当月TACOS广告花费占比", value=f"{curr_tacos:.2%}", delta=f"{delta_tacos:.2%}")
    if has_last_month:
        st.caption(f"上月：{last_tacos:.2%} | 环比：{pct_tacos:.1%}")
with row2_col2:
    st.metric(label="当月广告ACOS", value=f"{curr_acos:.2%}", delta=f"{delta_acos:.2%}")
    if has_last_month:
        st.caption(f"上月：{last_acos:.2%} | 环比：{pct_acos:.1%}")
with row2_col3:
    st.metric(label="当月广告ROAS投产比", value=f"{curr_roas:.2f}", delta=f"{delta_roas:.2f}")
    if has_last_month:
        st.caption(f"上月：{last_roas:.2f} | 环比：{pct_roas:.1%}")
with row2_col4:
    st.metric(label="当月广告销售依赖度ASoAS", value=f"{curr_asoas:.2%}", delta=f"{delta_asoas:.2%}")
    if has_last_month:
        st.caption(f"上月：{last_asoas:.2%} | 环比：{pct_asoas:.1%}")

# 第三行3卡：CPC、CTR、CVR（不变）
row3_col1, row3_col2, row3_col3 = st.columns(3)
with row3_col1:
    st.metric(label="当月平均单次点击CPC", value=f"${curr_cpc:.2f}", delta=f"{delta_cpc:.2f}")
    if has_last_month:
        st.caption(f"上月：${last_cpc:.2f} | 环比：{pct_cpc:.1%}")
with row3_col2:
    st.metric(label="当月平均点击率CTR", value=f"{curr_ctr:.2%}", delta=f"{delta_ctr:.2%}")
    if has_last_month:
        st.caption(f"上月：{last_ctr:.2%} | 环比：{pct_ctr:.1%}")
with row3_col3:
    st.metric(label="当月平均广告转化率CVR", value=f"{curr_cvr:.2%}", delta=f"{delta_cvr:.2%}")
    if has_last_month:
        st.caption(f"上月：{last_cvr:.2%} | 环比：{pct_cvr:.1%}")

# 当月文字说明同步补充广告销售额
st.info(f"""
【{select_month}单月概况说明】
分析店铺：{select_shop} | 分析月份：{select_month}
1. 当月广告投放总额 ${curr_ad_spend:,.2f}，广告成交销售额 ${curr_ad_sales:,.2f}，店铺全部总销售额 ${curr_sales:,.2f}，广告花费占营收比重TACOS {curr_tacos:.2%}
2. 当月广告自身转化成本ACOS {curr_acos:.2%}，每投入1美金广告带来 {curr_roas:.2f} 美金广告营收
3. 当月店铺 {curr_asoas:.2%} 的营收来自付费广告，剩余为自然免费流量订单
4. 当月单次点击成本CPC ${curr_cpc:.2f}，广告曝光点击率CTR {curr_ctr:.2%}，点击下单转化率 {curr_cvr:.2%}
""")
# -------------------------- 全店铺所有月份完整指标明细表（修改后） --------------------------
st.subheader(f"📋 {select_shop}店铺 全月份完整指标明细表")
# 取当前选中店铺全部月份数据
df_all_month_table = df_raw[df_raw["店铺"] == select_shop].groupby("年月").agg({
    "展示": "sum",
    "点击": "sum",
    "广告花费": "sum",
    "SP广告费": "sum",
    "SB广告费": "sum",
    "SBV广告费": "sum",
    "SP广告销售额": "sum",
    "SB广告销售额": "sum",
    "SBV广告销售额": "sum",
    "广告销售额": "sum",
    "广告订单量": "sum",
    "销售额": "sum",
    "订单量": "sum"
}).reset_index()

# 批量计算所有月份衍生指标
df_all_month_table["CTR"] = np.where(df_all_month_table["展示"] == 0, 0, df_all_month_table["点击"] / df_all_month_table["展示"])
df_all_month_table["CPC"] = np.where(df_all_month_table["点击"] == 0, 0, df_all_month_table["广告花费"] / df_all_month_table["点击"])
df_all_month_table["CVR广告转化率"] = np.where(df_all_month_table["点击"] == 0, 0, df_all_month_table["广告订单量"] / df_all_month_table["点击"])
df_all_month_table["ACOS"] = np.where(df_all_month_table["广告销售额"] == 0, 0, df_all_month_table["广告花费"] / df_all_month_table["广告销售额"])
df_all_month_table["ROAS"] = np.where(df_all_month_table["广告花费"] == 0, 0, df_all_month_table["广告销售额"] / df_all_month_table["广告花费"])
df_all_month_table["TACOS广告花费占比"] = np.where(df_all_month_table["销售额"] == 0, 0, df_all_month_table["广告花费"] / df_all_month_table["销售额"])
df_all_month_table["ASoAS广告销售依赖度"] = np.where(df_all_month_table["销售额"] == 0, 0, df_all_month_table["广告销售额"] / df_all_month_table["销售额"])
df_all_month_table["SP广告花费占比"] = np.where(df_all_month_table["广告花费"] == 0, 0, df_all_month_table["SP广告费"] / df_all_month_table["广告花费"])
df_all_month_table["SB广告花费占比"] = np.where(df_all_month_table["广告花费"] == 0, 0, df_all_month_table["SB广告费"] / df_all_month_table["广告花费"])
df_all_month_table["SBV广告花费占比"] = np.where(df_all_month_table["广告花费"] == 0, 0, df_all_month_table["SBV广告费"] / df_all_month_table["广告花费"])

# 固定展示列不变
show_month_cols = [
    "年月", "展示", "点击", "广告花费", "销售额", "广告销售额",
    "TACOS广告花费占比", "ACOS", "ROAS", "CPC", "CTR", "CVR广告转化率",
    "ASoAS广告销售依赖度", "SP广告花费占比", "SB广告花费占比", "SBV广告花费占比"
]
# 按年月倒序，最新月份排在最上方
df_all_month_table = df_all_month_table.sort_values("年月", ascending=False).reset_index(drop=True)
df_table_show = df_all_month_table[show_month_cols].copy()

# ---------------------- 格式化配置 ----------------------
# 所有需要百分比展示的列
percent_cols = [
    "TACOS广告花费占比", "ACOS", "CTR", "CVR广告转化率",
    "ASoAS广告销售依赖度", "SP广告花费占比", "SB广告花费占比", "SBV广告花费占比"
]
# 金额、单价、ROAS 保留两位小数
float_2_cols = ["广告花费", "销售额", "广告销售额", "CPC", "ROAS"]
# 展示、点击 整数，无小数
int_cols = ["展示", "点击"]

# 表格样式渲染
styled_table = df_table_show.style\
    .format(formatter="{:.2%}", subset=percent_cols)\
    .format(formatter="{:.2f}", subset=float_2_cols)\
    .format(formatter="{:.0f}", subset=int_cols)

st.dataframe(styled_table, use_container_width=True, height=320)
st.caption("表格按月份倒序排列，顶部为最新月份；比率类字段以百分比展示，金额/单价保留两位小数，展示点击为整数，可横向滑动查看全部广告指标历史数据")

# ===================== 新增：TACOS走势与归因分析 =====================
st.markdown(f"## 📈 二、{select_shop}店铺 TACOS走势与上涨归因分析")

# 复用全月数据并按时间正序排列（画图用）
df_trend = df_all_month_table.sort_values("年月", ascending=True).reset_index(drop=True)
# 生成中文年月格式：2026年6月
df_trend["年月中文"] = df_trend["年月"].apply(lambda x: f"{x.split('-')[0]}年{int(x.split('-')[1])}月")

# 同一行两列布局
trend_col1, trend_col2 = st.columns([1.3, 1])

# 左图：销售额柱形 + 广告花费折线 组合图（双Y轴）
with trend_col1:
    fig_combo = go.Figure()
    # 柱形：全店总销售额
    fig_combo.add_trace(go.Bar(
        x=df_trend["年月中文"],
        y=df_trend["销售额"],
        name="全店总销售额",
        yaxis="y",
        marker_color="#4C78A8",
        opacity=0.7,
        hovertemplate="%{x}<br>全店销售额：$%{y:,.2f}<extra></extra>"
    ))
    # 折线：广告花费
    fig_combo.add_trace(go.Scatter(
        x=df_trend["年月中文"],
        y=df_trend["广告花费"],
        name="广告总花费",
        yaxis="y2",
        mode="lines+markers",
        line=dict(color="#E45756", width=3),
        marker=dict(size=8),
        hovertemplate="%{x}<br>广告花费：$%{y:,.2f}<extra></extra>"
    ))
    fig_combo.update_layout(
        title="销售额（柱形）vs 广告花费（折线）",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0),
        yaxis=dict(title="全店总销售额 ($)", side="left"),
        yaxis2=dict(title="广告总花费 ($)", side="right", overlaying="y"),
        hovermode="x unified",
        height=420,
        margin=dict(l=10, r=10, t=50, b=10)
    )
    st.plotly_chart(fig_combo, use_container_width=True)

# 右图：TACOS 折线图
with trend_col2:
    fig_tacos = go.Figure()
    fig_tacos.add_trace(go.Scatter(
        x=df_trend["年月中文"],
        y=df_trend["TACOS广告花费占比"],
        name="TACOS",
        mode="lines+markers+text",
        line=dict(color="#F58518", width=3),
        marker=dict(size=9),
        text=[f"{v:.1%}" for v in df_trend["TACOS广告花费占比"]],
        textposition="top center",
        hovertemplate="%{x}<br>TACOS：%{y:.2%}<extra></extra>"
    ))
    fig_tacos.update_layout(
        title="TACOS 月度走势",
        yaxis=dict(title="TACOS（广告花费/总销售额）", tickformat=".1%"),
        height=420,
        margin=dict(l=10, r=10, t=50, b=10),
        hovermode="x unified"
    )
    st.plotly_chart(fig_tacos, use_container_width=True)

# ---- 自动生成归因分析文本 ----
# 取最新月与上月对比
latest_idx = df_trend.index[-1]
curr_tacos_val = df_trend["TACOS广告花费占比"].iloc[latest_idx]
curr_spend_val = df_trend["广告花费"].iloc[latest_idx]
curr_sales_val = df_trend["销售额"].iloc[latest_idx]

if latest_idx >= 1:
    prev_tacos_val = df_trend["TACOS广告花费占比"].iloc[latest_idx - 1]
    prev_spend_val = df_trend["广告花费"].iloc[latest_idx - 1]
    prev_sales_val = df_trend["销售额"].iloc[latest_idx - 1]
    tacos_change = curr_tacos_val - prev_tacos_val
    spend_pct = (curr_spend_val - prev_spend_val) / prev_spend_val if prev_spend_val != 0 else 0
    sales_pct = (curr_sales_val - prev_sales_val) / prev_sales_val if prev_sales_val != 0 else 0

    # 判断归因方向
    if tacos_change > 0:
        direction = "上涨"
        direction_color = "🔴"
    elif tacos_change < 0:
        direction = "下降"
        direction_color = "🟢"
    else:
        direction = "持平"
        direction_color = "⚪"

    # 拆解原因
    reason_list = []
    if spend_pct > 0 and sales_pct <= 0:
        reason_list.append(f"广告花费环比**上升 {spend_pct:.1%}**，而全店销售额反而**下降 {sales_pct:.1%}**，一增一减共同推高TACOS")
    elif spend_pct > sales_pct and sales_pct > 0:
        reason_list.append(f"广告花费增速（{spend_pct:.1%}）**快于**销售额增速（{sales_pct:.1%}），投入扩张但营收未同步跟上")
    elif spend_pct >= 0 and sales_pct < 0:
        reason_list.append(f"销售额环比下滑 {sales_pct:.1%}，分母收缩导致TACOS被动抬升")
    elif spend_pct < 0 and sales_pct < 0 and abs(sales_pct) > abs(spend_pct):
        reason_list.append(f"广告花费虽下降 {abs(spend_pct):.1%}，但销售额下滑更快（{abs(sales_pct):.1%}），整体TACOS仍走高")
    else:
        reason_list.append("广告花费与销售额变动幅度接近，TACOS基本稳定")

    # 补充建议
    suggestions = []
    if tacos_change > 0.02:  # 上涨超2个百分点
        suggestions.append("建议排查广告投放效率：检查ACOS是否同步恶化，重点优化高花费低产出的SP广告活动")
        suggestions.append("关注自然流量与自然单占比：若ASoAS同步上升，说明店铺对广告依赖度在增加，需布局自然流量")
    elif tacos_change < -0.01:
        suggestions.append("TACOS改善明显，可评估是否有进一步加预算扩大规模的空间")
    suggestions.append("可结合SP/SB/SBV分渠道花费占比，定位是哪类广告拉动了整体花费上涨")

    analysis_text = f"""
### {direction_color} {select_month} TACOS {direction}归因分析

**核心数据对比（环比上月）：**
- TACOS：{prev_tacos_val:.2%} → {curr_tacos_val:.2%}，**{direction} {abs(tacos_change):.2%}**
- 广告花费：${prev_spend_val:,.2f} → ${curr_spend_val:,.2f}，环比 {spend_pct:+.1%}
- 全店销售额：${prev_sales_val:,.2f} → ${curr_sales_val:,.2f}，环比 {sales_pct:+.1%}

**主要原因判断：**
"""
    for i, r in enumerate(reason_list, 1):
        analysis_text += f"{i}. {r}\n"

    analysis_text += "\n**优化建议：**\n"
    for i, s in enumerate(suggestions, 1):
        analysis_text += f"{i}. {s}\n"

    st.markdown(analysis_text)
else:
    st.info("仅有单月数据，暂无法进行环比归因分析，请等待更多月份数据积累。")

st.divider()

# ===================== 新增：渠道拆分 SP/SB/SBV 深度拆解（修复字段缺失） =====================
st.markdown("## 📊 三、分广告渠道(SP/SB/SBV)TACOS&ACOS拆解分析")
# 复用按月汇总数据集，时间正序绘图
df_channel = df_all_month_table.sort_values("年月", ascending=True).reset_index(drop=True)
df_channel["年月中文"] = df_channel["年月"].apply(lambda x: f"{x.split('-')[0]}年{int(x.split('-')[1])}月")

# 双列布局：左=渠道花费堆叠柱状，右=分渠道TACOS折线
chan_col1, chan_col2 = st.columns([1.2, 1])

# 左图：SP/SB/SBV月度花费堆叠柱状图
# 左图：SP/SB/SBV月度花费堆叠柱状图（新增占比悬浮提示）
with chan_col1:
    # 先计算每月总花费，用于算占比
    df_channel["总广告花费"] = df_channel["SP广告费"] + df_channel["SB广告费"] + df_channel["SBV广告费"]
    df_channel["SP占比"] = (df_channel["SP广告费"] / df_channel["总广告花费"]).apply(lambda x: f"{x:.1%}")
    df_channel["SB占比"] = (df_channel["SB广告费"] / df_channel["总广告花费"]).apply(lambda x: f"{x:.1%}")
    df_channel["SBV占比"] = (df_channel["SBV广告费"] / df_channel["总广告花费"]).apply(lambda x: f"{x:.1%}")

    fig_spend_stack = go.Figure()
    fig_spend_stack.add_trace(go.Bar(
        x=df_channel["年月中文"],
        y=df_channel["SP广告费"],
        name="SP搜索广告",
        marker_color="#E45756",
        hovertemplate="%{x}<br>SP花费：$%{y:,.2f}<br>当月广告占比：%{customdata}<extra></extra>",
        customdata=df_channel["SP占比"]
    ))
    fig_spend_stack.add_trace(go.Bar(
        x=df_channel["年月中文"],
        y=df_channel["SB广告费"],
        name="SB品牌广告",
        marker_color="#4C78A8",
        hovertemplate="%{x}<br>SB花费：$%{y:,.2f}<br>当月广告占比：%{customdata}<extra></extra>",
        customdata=df_channel["SB占比"]
    ))
    fig_spend_stack.add_trace(go.Bar(
        x=df_channel["年月中文"],
        y=df_channel["SBV广告费"],
        name="SBV视频展示广告",
        marker_color="#59A14F",
        hovertemplate="%{x}<br>SBV花费：$%{y:,.2f}<br>当月广告占比：%{customdata}<extra></extra>",
        customdata=df_channel["SBV占比"]
    ))
    fig_spend_stack.update_layout(
        title="各渠道月度广告花费堆叠（hover查看渠道当月花费占比）",
        barmode="stack",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0),
        yaxis_title="广告花费 ($)",
        height=420,
        margin=dict(l=10, r=10, t=50, b=10),
        hovermode="x unified"
    )
    st.plotly_chart(fig_spend_stack, use_container_width=True)

# 计算指标：渠道TACOS、渠道ACOS
# 渠道TACOS = 渠道广告费 / 全店总销售额
df_channel["SP_TACOS"] = np.where(df_channel["销售额"] == 0, 0, df_channel["SP广告费"] / df_channel["销售额"])
df_channel["SB_TACOS"] = np.where(df_channel["销售额"] == 0, 0, df_channel["SB广告费"] / df_channel["销售额"])
df_channel["SBV_TACOS"] = np.where(df_channel["销售额"] == 0, 0, df_channel["SBV广告费"] / df_channel["销售额"])

# 渠道ACOS = 渠道广告费 / 渠道自身广告销售额
df_channel["SP_ACOS"] = np.where(df_channel["SP广告销售额"] == 0, 0, df_channel["SP广告费"] / df_channel["SP广告销售额"])
df_channel["SB_ACOS"] = np.where(df_channel["SB广告销售额"] == 0, 0, df_channel["SB广告费"] / df_channel["SB广告销售额"])
df_channel["SBV_ACOS"] = np.where(df_channel["SBV广告销售额"] == 0, 0, df_channel["SBV广告费"] / df_channel["SBV广告销售额"])

# 右图：分渠道独立TACOS折线对比
with chan_col2:
    fig_channel_tacos = go.Figure()
    # 整体总TACOS基准虚线
    fig_channel_tacos.add_trace(go.Scatter(
        x=df_channel["年月中文"],
        y=df_channel["TACOS广告花费占比"],
        name="整体总TACOS",
        mode="lines+markers",
        line=dict(color="#F58518", width=4, dash="dash"),
        marker=dict(size=7),
        hovertemplate="%{x}<br>整体TACOS：%{y:.2%}<extra></extra>"
    ))
    fig_channel_tacos.add_trace(go.Scatter(
        x=df_channel["年月中文"],
        y=df_channel["SP_TACOS"],
        name="SP渠道TACOS",
        mode="lines+markers",
        line=dict(color="#E45756", width=2),
        marker=dict(size=6),
        hovertemplate="%{x}<br>SP-TACOS：%{y:.2%}<extra></extra>"
    ))
    fig_channel_tacos.add_trace(go.Scatter(
        x=df_channel["年月中文"],
        y=df_channel["SB_TACOS"],
        name="SB渠道TACOS",
        mode="lines+markers",
        line=dict(color="#4C78A8", width=2),
        marker=dict(size=6),
        hovertemplate="%{x}<br>SB-TACOS：%{y:.2%}<extra></extra>"
    ))
    fig_channel_tacos.add_trace(go.Scatter(
        x=df_channel["年月中文"],
        y=df_channel["SBV_TACOS"],
        name="SBV渠道TACOS",
        mode="lines+markers",
        line=dict(color="#59A14F", width=2),
        marker=dict(size=6),
        hovertemplate="%{x}<br>SBV-TACOS：%{y:.2%}<extra></extra>"
    ))
    fig_channel_tacos.update_layout(
        title="分渠道TACOS vs 店铺整体TACOS",
        yaxis_title="TACOS占比",
        yaxis_tickformat=".1%",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0),
        height=420,
        margin=dict(l=10, r=10, t=50, b=10),
        hovermode="x unified"
    )
    st.plotly_chart(fig_channel_tacos, use_container_width=True)

# ===================== 新增：四、广告底层投放效率深度拆解 =====================
st.markdown("## 📉 四、广告底层投放效率（CPC/CTR/CVR/整体ACOS）根源分析")
df_eff = df_all_month_table.sort_values("年月", ascending=True).reset_index(drop=True)
df_eff["年月中文"] = df_eff["年月"].apply(lambda x: f"{x.split('-')[0]}年{int(x.split('-')[1])}月")

# 预先计算四大核心效率指标（兜底除0防报错）
# CTR 点击率 = 点击 / 展示
df_eff["CTR"] = np.where(df_eff["展示"] == 0, 0, df_eff["点击"] / df_eff["展示"])
# CPC 单次点击成本 = 广告花费 / 点击
df_eff["CPC"] = np.where(df_eff["点击"] == 0, 0, df_eff["广告花费"] / df_eff["点击"])
# CVR 广告转化率 = 广告订单量 / 点击
df_eff["CVR"] = np.where(df_eff["点击"] == 0, 0, df_eff["广告订单量"] / df_eff["点击"])
# ACOS 整体广告投产 = 广告花费 / 广告销售额
df_eff["整体ACOS"] = np.where(df_eff["广告销售额"] == 0, 0, df_eff["广告花费"] / df_eff["广告销售额"])

# 布局：上下两组图表，第一行双列，第二行单列全宽
eff_row1_col1, eff_row1_col2 = st.columns([1, 1])

# 左图1：CPC & CTR 双折线（流量成本+点击精准度）
with eff_row1_col1:
    fig_cost_click = go.Figure()
    # CPC 折线
    fig_cost_click.add_trace(go.Scatter(
        x=df_eff["年月中文"],
        y=df_eff["CPC"],
        name="CPC单次点击成本($)",
        mode="lines+markers",
        line=dict(color="#E45756", width=3),
        hovertemplate="%{x}<br>CPC：$%{y:.2f}<extra></extra>"
    ))
    # CTR 次坐标轴
    fig_cost_click.add_trace(go.Scatter(
        x=df_eff["年月中文"],
        y=df_eff["CTR"],
        name="CTR点击率",
        mode="lines+markers",
        line=dict(color="#4C78A8", width=3),
        yaxis="y2",
        hovertemplate="%{x}<br>CTR：%{y:.2%}<extra></extra>"
    ))
    fig_cost_click.update_layout(
        title="CPC单次点击成本 & CTR点击率走势",
        yaxis=dict(title="CPC ($)", side="left"),
        yaxis2=dict(title="CTR 占比", side="right", overlaying="y", tickformat=".1%"),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0),
        height=400,
        margin=dict(l=10, r=10, t=45, b=10),
        hovermode="x unified"
    )
    st.plotly_chart(fig_cost_click, use_container_width=True)

# 右图2：CVR转化率 & 整体ACOS双折线（转化&广告投产）
with eff_row1_col2:
    fig_conv_acos = go.Figure()
    # CVR 转化率
    fig_conv_acos.add_trace(go.Scatter(
        x=df_eff["年月中文"],
        y=df_eff["CVR"],
        name="CVR广告转化率",
        mode="lines+markers",
        line=dict(color="#59A14F", width=3),
        hovertemplate="%{x}<br>CVR：%{y:.2%}<extra></extra>"
    ))
    # ACOS 次坐标轴
    fig_conv_acos.add_trace(go.Scatter(
        x=df_eff["年月中文"],
        y=df_eff["整体ACOS"],
        name="店铺整体ACOS",
        mode="lines+markers",
        line=dict(color="#F58518", width=3),
        yaxis="y2",
        hovertemplate="%{x}<br>ACOS：%{y:.2%}<extra></extra>"
    ))
    fig_conv_acos.update_layout(
        title="CVR广告转化率 & 整体ACOS走势",
        yaxis=dict(title="CVR 转化率", side="left", tickformat=".1%"),
        yaxis2=dict(title="ACOS广告成本占比", side="right", overlaying="y", tickformat=".1%"),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0),
        height=400,
        margin=dict(l=10, r=10, t=45, b=10),
        hovermode="x unified"
    )
    st.plotly_chart(fig_conv_acos, use_container_width=True)

# 下方：自动底层效率归因文本
st.subheader("🔍 底层投放效率自动归因分析")
if len(df_eff) >= 2:
    curr_eff = df_eff.iloc[-1]
    prev_eff = df_eff.iloc[-2]
    curr_month = curr_eff["年月中文"]

    # 环比差值计算
    cpc_diff = curr_eff["CPC"] - prev_eff["CPC"]
    ctr_diff = curr_eff["CTR"] - prev_eff["CTR"]
    cvr_diff = curr_eff["CVR"] - prev_eff["CVR"]
    acos_diff = curr_eff["整体ACOS"] - prev_eff["整体ACOS"]

    # 三列并排展示核心指标环比
    eff_col1, eff_col2, eff_col3 = st.columns(3)
    with eff_col1:
        st.markdown("#### 流量成本指标")
        st.markdown(f"""
- CPC环比变动：${cpc_diff:+.2f}
- 当月CPC：${curr_eff["CPC"]:.2f}
- CTR环比变动：{ctr_diff:+.2%}
- 当月CTR：{curr_eff["CTR"]:.2%}
""")
    with eff_col2:
        st.markdown("#### 转化效率指标")
        st.markdown(f"""
- CVR环比变动：{cvr_diff:+.2%}
- 当月CVR：{curr_eff["CVR"]:.2%}
""")
    with eff_col3:
        st.markdown("#### 广告投产指标")
        st.markdown(f"""
- ACOS环比变动：{acos_diff:+.2%}
- 当月整体ACOS：{curr_eff["整体ACOS"]:.2%}
""")
else:
    st.info("效率环比对比至少需要2个月历史数据，当前数据不足无法拆解底层变动。")

st.divider()


# ===================== 五、新品老品分层投放拆解（修复TACOS归零+hover明细） =====================
# ===================== 五、新品老品分层投放拆解（适配Excel完整产品类型标签） =====================
st.markdown("## 🧩 五、新品老品分层投放拆解（定位拖垮TACOS的商品层级）")

# 1、筛选当前店铺全量明细，按年月+产品类型聚合分层数据
df_shop_raw = df_raw[df_raw["店铺"] == select_shop].copy()
df_type_group = df_shop_raw.groupby(["年月", "产品类型"]).agg({
    "SP广告费":"sum",
    "SB广告费":"sum",
    "SBV广告费":"sum",
    "广告花费":"sum",
    "广告销售额":"sum",
    "销售额":"sum",
    "广告订单量":"sum",
    "订单量":"sum",
    "展示":"sum",
    "点击":"sum"
}).reset_index()

# 生成中文年月
df_type_group["年月中文"] = df_type_group["年月"].apply(lambda x: f"{x.split('-')[0]}年{int(x.split('-')[1])}月")

# ========== 修复1：重写分层TACOS/ACOS计算，销售额为0不强制归零 ==========
df_type_group["分层TACOS"] = df_type_group.apply(
    lambda row: row["广告花费"] / row["销售额"] if row["销售额"] != 0 else None, axis=1
)
df_type_group["分层ACOS"] = df_type_group.apply(
    lambda row: row["广告花费"] / row["广告销售额"] if row["广告销售额"] != 0 else None, axis=1
)

# 排序绘图数据
df_type_sort = df_type_group.sort_values("年月", ascending=True)
st.subheader("📈 各分层月度花费 & 分层TACOS全周期走势")
t1, t2 = st.columns([1.2, 1])

# 左图：三层商品广告花费堆叠柱状
# 左图：三层商品广告花费堆叠柱状
with t1:
    fig_type_spend = go.Figure()
    # 【重点修改】key和Excel完整标签一字不差
    color_map = {
        "新品 (开售天数小于等于90)": "#ff7f0e",
        "老品 (开售天数大于90天)": "#2ca02c",
        "新品未出单": "#d62728"
    }
    type_list = [
        "新品 (开售天数小于等于90)",
        "老品 (开售天数大于90天)",
        "新品未出单"
    ]

    # 1. 先计算每个月份店铺整体总额，用来算占比
    df_month_total_all = df_type_group.groupby("年月").agg(
        当月总广告花费=("广告花费", "sum"),
        当月总销售额=("销售额", "sum")
    ).reset_index()
    # 关联回分层表
    df_type_join = df_type_group.merge(df_month_total_all, on="年月", how="left")
    # 计算分层占比
    df_type_join["花费占比"] = df_type_join["广告花费"] / df_type_join["当月总广告花费"]
    df_type_join["销售额占比"] = df_type_join["销售额"] / df_type_join["当月总销售额"]

    df_type_sort = df_type_join.sort_values("年月", ascending=True)

    for tp in type_list:
        sub = df_type_sort[df_type_sort["产品类型"] == tp]
        fig_type_spend.add_trace(go.Bar(
            x=sub["年月中文"],
            y=sub["广告花费"],
            name=tp,
            marker_color=color_map[tp],
            customdata=sub[
                ["销售额","广告销售额","分层TACOS",
                 "花费占比","销售额占比","当月总广告花费","当月总销售额"]
            ].values,
            hovertemplate="""%{x}<br>分层：%{name}
广告花费:$%{y:,.2f}（占当月店铺总广告花费 %{customdata[3]:.2%}）
总销售额:$%{customdata[0]:,.2f}（占当月店铺总销售额 %{customdata[4]:.2%}）
广告销售额:$%{customdata[1]:,.2f}
分层TACOS:%{customdata[2]:.2%}<extra></extra>"""
        ))
    fig_type_spend.update_layout(
        title="各商品分层月度广告花费堆叠",
        barmode="stack",
        height=420,
        yaxis_title="广告花费($)",
        legend=dict(orientation="h", y=1.02, x=0),
        hovermode="x unified"
    )
    st.plotly_chart(fig_type_spend, use_container_width=True)

# 右图：三层商品分层TACOS折线对比
with t2:
    fig_type_tacos = go.Figure()
    for tp in type_list:
        sub = df_type_sort[df_type_sort["产品类型"] == tp]
        fig_type_tacos.add_trace(go.Scatter(
            x=sub["年月中文"],
            y=sub["分层TACOS"],
            name=f"{tp} TACOS",
            mode="lines+markers",
            line=dict(color=color_map[tp], width=2),
            marker=dict(size=6),
            customdata=sub[["广告花费","分层ACOS"]].values,
            hovertemplate="%{x}<br>分层：%{name}<br>TACOS:%{y:.2%}<br>广告花费:$%{customdata[0]:,.2f}<br>分层ACOS:%{customdata[1]:.2%}<extra></extra>"
        ))
    fig_type_tacos.update_layout(
        title="各商品分层TACOS走势对比",
        height=420,
        yaxis_title="分层TACOS",
        yaxis_tickformat=".1%",
        legend=dict(orientation="h", y=1.02, x=0),
        hovermode="x unified"
    )
    st.plotly_chart(fig_type_tacos, use_container_width=True)

st.divider()

# ---------------------- 单品维度：当月全部单品明细表（TACOS高于店铺均值标红 + 金额2位小数+比率百分比格式化） ----------------------
# ---------------------- 单品维度：当月全部单品明细表（TACOS高于店铺均值标红 + 新增CTR/CPC/CVR） ----------------------
st.subheader(f"📋 {select_month} 全单品广告花费明细表（单品TACOS高于店铺整体TACOS标红）")
# 筛选当前月份单品明细
df_single_item = df_shop_raw[df_shop_raw["年月"] == select_month].copy()

# 计算店铺当月整体TACOS（对比基准）
df_month_total = df_shop_raw[df_shop_raw["年月"] == select_month].agg({
    "广告花费":"sum",
    "销售额":"sum"
})
shop_total_tacos = df_month_total["广告花费"] / df_month_total["销售额"] if df_month_total["销售额"] != 0 else 0

# 计算单品核心指标
df_single_item["单品ACOS"] = df_single_item.apply(
    lambda r: r["广告花费"]/r["广告销售额"] if r["广告销售额"] !=0 else None, axis=1
)
df_single_item["单品TACOS"] = df_single_item.apply(
    lambda r: r["广告花费"]/r["销售额"] if r["销售额"] !=0 else None, axis=1
)
# ========== 新增：CTR点击率、CPC单次点击成本、CVR广告转化率 ==========
df_single_item["CTR"] = df_single_item.apply(
    lambda r: r["点击"]/r["展示"] if r["展示"] !=0 else None, axis=1
)
df_single_item["CPC"] = df_single_item.apply(
    lambda r: r["广告花费"]/r["点击"] if r["点击"] !=0 else None, axis=1
)
df_single_item["CVR"] = df_single_item.apply(
    lambda r: r["广告订单量"]/r["点击"] if r["点击"] !=0 else None, axis=1
)

# 按广告花费倒序（全部数据）
df_all_item = df_single_item.sort_values("广告花费", ascending=False)

# 展示字段（新增CTR/CPC/CVR，按流量→成本→转化→投产逻辑排序）
item_show_cols = [
    "MSKU","品名","产品类型","分层", "开售时间",
    "展示", "点击", "CTR", "CPC", "CVR",
    "广告花费", "广告销售额", "销售额",
    "单品ACOS", "单品TACOS", "广告订单量"
]
df_table = df_all_item[item_show_cols].copy()

# 条件标红函数：单品TACOS > 店铺整体TACOS 文字红色
def highlight_high_tacos(val):
    if pd.isna(val):
        return ""
    if val > shop_total_tacos:
        color = "red"
    else:
        color = "black"
    return f"color: {color}"

# 表格格式化：
# 1. 百分比列：CTR、CVR、ACOS、TACOS
pct_cols = ["CTR", "CVR", "单品ACOS", "单品TACOS"]
# 2. 金额列：CPC、广告花费、广告销售额、销售额（保留2位小数）
money_cols = ["CPC", "广告花费", "广告销售额", "销售额"]
# 3. 整数列：展示、点击、广告订单量
int_cols = ["展示", "点击", "广告订单量"]

styled_df = df_table.style\
    .map(highlight_high_tacos, subset=["单品TACOS"])\
    .format(formatter="{:.2%}", subset=pct_cols)\
    .format(formatter="{:.2f}", subset=money_cols)\
    .format(formatter="{:.0f}", subset=int_cols)

st.dataframe(styled_df, use_container_width=True, height=400)

st.caption(f"本月店铺整体TACOS：{shop_total_tacos:.2%}；红色=单品TACOS高于店铺平均水平；CTR/CVR为百分比，CPC/花费/销售额保留两位小数")

# ===================== 新增：分层TACOS超标数量统计分析 =====================
st.subheader("📊 三类商品单品TACOS超标数量对比分析")

# 筛选有效TACOS数据（排除空值无销售额商品）
df_valid = df_all_item.dropna(subset=["单品TACOS"])

# 初始化统计字典
stat_result = {}
type_name_list = [
    "新品 (开售天数小于等于90)",
    "老品 (开售天数大于90天)",
    "新品未出单"
]

for tag in type_name_list:
    df_type = df_valid[df_valid["产品类型"] == tag]
    total_cnt = len(df_type)
    over_cnt = len(df_type[df_type["单品TACOS"] > shop_total_tacos])
    over_rate = over_cnt / total_cnt if total_cnt > 0 else 0
    stat_result[tag] = {
        "总单品数": total_cnt,
        "TACOS超标单品数": over_cnt,
        "超标占比": over_rate
    }

# 三列展示统计数字
col_a, col_b, col_c = st.columns(3)
col_list = [col_a, col_b, col_c]
for idx, tag in enumerate(type_name_list):
    data = stat_result[tag]
    with col_list[idx]:
        st.markdown(f"**{tag}**")
        st.metric("总SKU数量", value=data["总单品数"])
        st.metric("TACOS超标SKU", value=data["TACOS超标单品数"])
        st.caption(f"超标占比：{data['超标占比']:.2%}")

# 自动生成综合分析文案
st.markdown("### 分层投放成本综合解读")
text_lines = []
text_lines.append(f"本月店铺整体基准TACOS为：**{shop_total_tacos:.2%}**，各分层单品投放超标情况如下：")

for tag in type_name_list:
    d = stat_result[tag]
    if d["总单品数"] == 0:
        line = f"- {tag}：本月无投放SKU"
    else:
        line = f"- {tag}：共{d['总单品数']}个SKU，其中{d['TACOS超标单品数']}个单品TACOS高于店铺均值，超标占比{d['超标占比']:.2%}"
    text_lines.append(line)

# 风险总结
max_over_tag = max(stat_result.items(), key=lambda x: x[1]["超标占比"])
max_tag, max_data = max_over_tag
text_lines.append(f"""
#### 核心风险提示
本月**{max_tag}**分层商品超标占比最高，达到{max_data["超标占比"]:.2%}，是拉高店铺整体TACOS的主要商品来源，建议优先针对该分层内高花费、高TACOS单品削减广告预算，优化关键词与Listing转化。
""")

st.markdown("\n".join(text_lines))
st.divider()

# ===================== ✨新增模块：自然流量、告警、漏斗、帕累托、四象限散点 =====================
st.markdown("## 🔎 七、自然流量与多维度深度诊断")
# -------------------------- ①自然流量指标拆解（当月） --------------------------
# 计算自然流量衍生指标
curr_natural_sales = curr_sales - curr_ad_sales
curr_natural_order = curr_all_order - curr_ad_order

# 占比（防除0）
curr_natural_sales_pct = curr_natural_sales / curr_sales if curr_sales > 0 else 0
curr_natural_order_pct = curr_natural_order / curr_all_order if curr_all_order > 0 else 0
curr_ad_order_pct = curr_ad_order / curr_all_order if curr_all_order > 0 else 0

# 客单价
curr_ad_avg_price = curr_ad_sales / curr_ad_order if curr_ad_order > 0 else 0
curr_natural_avg_price = curr_natural_sales / curr_natural_order if curr_natural_order > 0 else 0
curr_total_avg_price = curr_sales / curr_all_order if curr_all_order > 0 else 0

# 上月自然流量环比
last_natural_sales = last_sales - last_ad_sales
last_natural_order = last_all_order - last_ad_order
last_natural_sales_pct = last_natural_sales / last_sales if last_sales > 0 else 0
last_natural_order_pct = last_natural_order / last_all_order if last_all_order > 0 else 0

delta_natural_sales = curr_natural_sales - last_natural_sales
delta_natural_order = curr_natural_order - last_natural_order
pct_natural_sales = delta_natural_sales / last_natural_sales if last_natural_sales !=0 else 0
pct_natural_order = delta_natural_order / last_natural_order if last_natural_order !=0 else 0

st.subheader("7.1 当月自然流量指标拆解")
nat_col1, nat_col2, nat_col3, nat_col4 = st.columns(4)
with nat_col1:
    st.metric("自然销售额", value=f"${curr_natural_sales:,.2f}", delta=f"{delta_natural_sales:,.2f}")
    if has_last_month:
        st.caption(f"上月：${last_natural_sales:,.2f}｜环比 {pct_natural_sales:.1%}")
with nat_col2:
    st.metric("自然订单量", value=f"{curr_natural_order:,.0f}", delta=f"{delta_natural_order:,.0f}")
    if has_last_month:
        st.caption(f"上月：{last_natural_order:,.0f}｜环比 {pct_natural_order:.1%}")
with nat_col3:
    st.metric("自然销售额占比", value=f"{curr_natural_sales_pct:.2%}")
    st.caption(f"广告销售额占比ASoAS：{curr_asoas:.2%}")
with nat_col4:
    st.metric("自然订单占比", value=f"{curr_natural_order_pct:.2%}")
    st.caption(f"广告订单占比：{curr_ad_order_pct:.2%}")

nat_p1, nat_p2, nat_p3 = st.columns(3)
with nat_p1:
    st.metric("全店客单价", value=f"${curr_total_avg_price:.2f}")
with nat_p2:
    st.metric("广告订单客单价", value=f"${curr_ad_avg_price:.2f}")
with nat_p3:
    st.metric("自然订单客单价", value=f"${curr_natural_avg_price:.2f}")

st.info(f"""
📌自然流量解读：
- {select_month}自然销售额 ${curr_natural_sales:,.2f}，占总销售额 {curr_natural_sales_pct:.2%}；自然订单 {curr_natural_order:,.0f}，占全部订单 {curr_natural_order_pct:.2%}
- 广告客单 ${curr_ad_avg_price:.2f}，自然客单 ${curr_natural_avg_price:.2f}；
> 提示：如果总销售额下滑，优先看是【广告单下跌】还是【自然单下跌】，区分TACOS上涨是广告花费拉高，还是自然流量崩塌分母收缩。
""")

# -------------------------- ②广告订单占比 & 自然订单占比趋势 --------------------------
st.subheader("7.2 广告/自然订单占比历史趋势")
df_trend_order = df_all_month_table.sort_values("年月", ascending=True).reset_index(drop=True)
df_trend_order["年月中文"] = df_trend_order["年月"].apply(lambda x: f"{x.split('-')[0]}年{int(x.split('-')[1])}月")
# 按月计算自然订单、广告订单占比
df_trend_order["自然订单"] = df_trend_order["订单量"] - df_trend_order["广告订单量"]
df_trend_order["广告订单占比"] = np.where(df_trend_order["订单量"]==0,0, df_trend_order["广告订单量"] / df_trend_order["订单量"])
df_trend_order["自然订单占比"] = np.where(df_trend_order["订单量"]==0,0, df_trend_order["自然订单"] / df_trend_order["订单量"])

fig_order_ratio = go.Figure()
fig_order_ratio.add_trace(go.Scatter(
    x=df_trend_order["年月中文"], y=df_trend_order["广告订单占比"],
    name="广告订单占比", mode="lines+markers", line=dict(color="#E45756",width=3),
    hovertemplate="%{x}<br>广告订单占比：%{y:.2%}<extra></extra>"
))
fig_order_ratio.add_trace(go.Scatter(
    x=df_trend_order["年月中文"], y=df_trend_order["自然订单占比"],
    name="自然订单占比", mode="lines+markers", line=dict(color="#59A14F",width=3),
    hovertemplate="%{x}<br>自然订单占比：%{y:.2%}<extra></extra>"
))
fig_order_ratio.update_layout(
    title="月度广告订单占比 vs 自然订单占比",
    yaxis_tickformat=".1%", yaxis_title="占比",
    legend=dict(orientation="h",yanchor="bottom",y=1.02,xanchor="left",x=0),
    height=380, hovermode="x unified"
)
st.plotly_chart(fig_order_ratio, use_container_width=True)

# -------------------------- ③简单异常告警模块 --------------------------
st.subheader("7.3 当月广告风险自动告警")
warn_list = []
# 告警规则
if curr_acos > 0.35:
    warn_list.append(f"🔴 ACOS过高：当月ACOS {curr_acos:.2%}，大于35%，广告投产较差")
if has_last_month and last_cpc>0 and (curr_cpc - last_cpc)/last_cpc > 0.30:
    warn_list.append(f"🔴 CPC大幅上涨：CPC环比上涨 {((curr_cpc-last_cpc)/last_cpc):.1%}，点击成本抬升明显")
if has_last_month and last_cvr>0 and ((curr_cvr-last_cvr)/last_cvr) < -0.20:
    warn_list.append(f"🟡 CVR大幅下滑：转化率环比下降 {abs((curr_cvr-last_cvr)/last_cvr):.1%}，广告落地页转化能力下降")
if has_last_month and last_ad_spend>0 and (curr_ad_spend-last_ad_spend)/last_ad_spend>0.50 and curr_ad_sales <= last_ad_sales*1.05:
    warn_list.append(f"🔴 广告花费大幅增加，但广告销售额几乎无增长，存在无效消耗风险")
# 新品未出单花费告警
df_warn_newzero = df_single_item[ (df_single_item["产品类型"]=="新品未出单") & (df_single_item["广告花费"]>5) ]
if len(df_warn_newzero)>0:
    sum_zero_spend = df_warn_newzero["广告花费"].sum()
    warn_list.append(f"🟡新品未出单SKU合计广告花费 ${sum_zero_spend:,.2f}，多个新品花钱无广告出单，请排查")
if curr_natural_sales_pct < 0.30:
    warn_list.append(f"🟢⚠️店铺流量高度依赖广告，自然销售额占比仅 {curr_natural_sales_pct:.2%}，需布局自然流量")

if len(warn_list)==0:
    st.success("✅ 本月未检测到明显广告风险，投放整体健康")
else:
    for w in warn_list:
        st.warning(w)

# -------------------------- ④广告流量漏斗分析 --------------------------
st.subheader("7.4 广告流量漏斗（展示→点击→广告订单）")
funnel_data = [
    {"name":"广告展示曝光","value": int(df_month_single["展示"].iloc[0])},
    {"name":"广告点击","value": int(df_month_single["点击"].iloc[0])},
    {"name":"广告成交订单","value": int(df_month_single["广告订单量"].iloc[0])}
]
fig_funnel = go.Figure(go.Funnel(
    y=[i["name"] for i in funnel_data],
    x=[i["value"] for i in funnel_data],
    textinfo = "value+percent initial",
    marker={"color":["#4C78A8","#F58518","#59A14F"]}
))
fig_funnel.update_layout(title="广告完整转化漏斗", height=360)
st.plotly_chart(fig_funnel, use_container_width=True)

# -------------------------- ⑤80/20帕累托分析（SKU维度） --------------------------
st.subheader("7.5 SKU帕累托80‑20分析")
df_pareto = df_single_item[df_single_item["销售额"]>0].copy()
sum_pareto_sales = df_pareto["销售额"].sum()
st.write(f"df_pareto子集内销售额合计：${sum_pareto_sales:,.2f}")
st.write(f"店铺全局总销售额curr_sales：${curr_sales:,.2f}")
df_pareto = df_pareto.sort_values("销售额", ascending=False).reset_index(drop=True)
# 计算累计占比
df_pareto["销售额累计"] = df_pareto["销售额"].cumsum()
total_s_pareto = df_pareto["销售额"].sum()
df_pareto["销售额累计占比"] = df_pareto["销售额累计"] / total_s_pareto

df_pareto["广告花费累计"] = df_pareto["广告花费"].cumsum()
total_spend_pareto = df_pareto["广告花费"].sum()
df_pareto["广告花费累计占比"] = df_pareto["广告花费累计"] / total_spend_pareto

# 帕累托曲线
fig_pareto = go.Figure()
fig_pareto.add_trace(go.Scatter(
    x=list(range(len(df_pareto))), y=df_pareto["销售额累计占比"],
    name="销售额累计占比", mode="lines+markers", line=dict(color="#E45756",width=2),
    hovertemplate="SKU序号 %{x}<br>销售额累计占比：%{y:.2%}<extra></extra>"
))
fig_pareto.add_trace(go.Scatter(
    x=list(range(len(df_pareto))), y=df_pareto["广告花费累计占比"],
    name="广告花费累计占比", mode="lines+markers", line=dict(color="#4C78A8",width=2),
    hovertemplate="SKU序号 %{x}<br>广告花费累计占比：%{y:.2%}<extra></extra>"
))
fig_pareto.add_hline(y=0.8, line_dash="dash", line_color="gray", annotation_text="80%参考线")
fig_pareto.update_layout(
    title="SKU帕累托曲线（按销售额降序）X轴：SKU序号",
    yaxis_tickformat=".1%", yaxis_title="累计占比",
    xaxis_title="SKU数量(排序后)",
    legend=dict(orientation="h",yanchor="bottom",y=1.02,xanchor="left",x=0),
    height=400
)
st.plotly_chart(fig_pareto, use_container_width=True)

# 统计多少SKU贡献80%销售额
idx_80_sales = (df_pareto["销售额累计占比"]>=0.8).idxmin() if (df_pareto["销售额累计占比"]>=0.8).any() else len(df_pareto)-1
cnt_80sales_sku = idx_80_sales + 1
pct_80sales_sku = cnt_80sales_sku / len(df_pareto)
st.caption(f"📊统计：{cnt_80sales_sku}个SKU贡献本店铺80%销售额，占全部有效SKU数量 {pct_80sales_sku:.2%}")

# -------------------------- ⑥ASoAS & TACOS 联合矩阵四象限散点图 --------------------------
st.subheader("7.6 ASoAS & TACOS 四象限矩阵分析")
quad_col1, quad_col2 = st.columns([1,1])
# 左：历史月度四象限
with quad_col1:
    df_quad_month = df_all_month_table.copy()
    fig_quad_month = go.Figure()
    fig_quad_month.add_trace(go.Scatter(
        x=df_quad_month["ASoAS广告销售依赖度"],
        y=df_quad_month["TACOS广告花费占比"],
        text=df_quad_month["年月"],
        mode="markers+text",
        marker=dict(size=10,color="#165DFF"),
        hovertemplate="月份:%{text}<br>ASoAS:%{x:.2%}<br>TACOS:%{y:.2%}<extra></extra>"
    ))
    fig_quad_month.add_vline(x=0.5, line_dash="dash", line_color="#aaaaaa")
    fig_quad_month.add_hline(y=GLOBAL_TARGET_TACOS, line_dash="dash", line_color="#ff7700",annotation_text=f"目标TACOS {GLOBAL_TARGET_TACOS:.1%}")
    fig_quad_month.update_layout(
        title="历史月份四象限",
        xaxis_title="ASoAS广告销售依赖度", xaxis_tickformat=".1%",
        yaxis_title="TACOS广告花费占比", yaxis_tickformat=".1%",
        height=420
    )
    st.plotly_chart(fig_quad_month, use_container_width=True)

# 右：当月SKU维度四象限（修复KeyError：不要引用不存在的列ASoAS广告销售依赖度）
with quad_col2:
    df_quad_sku = df_single_item.copy()
    # SKU维度ASoAS：单品广告销售额 /单品总销售额
    df_quad_sku["sku_ASoAS"] = np.where(df_quad_sku["销售额"]>0, df_quad_sku["广告销售额"]/df_quad_sku["销售额"], np.nan)
    # ✅只过滤真实存在的两列：sku_ASoAS（现场生成）、单品TACOS
    df_quad_sku = df_quad_sku.dropna(subset=["sku_ASoAS","单品TACOS"])

    fig_quad_sku = go.Figure()
    fig_quad_sku.add_trace(go.Scatter(
        x=df_quad_sku["sku_ASoAS"],
        y=df_quad_sku["单品TACOS"],
        text=df_quad_sku["MSKU"],
        mode="markers",
        marker=dict(size=8,color="#E45756"),
        hovertemplate="MSKU:%{text}<br>单品ASoAS:%{x:.2%}<br>单品TACOS:%{y:.2%}<extra></extra>"
    ))
    fig_quad_sku.add_vline(x=0.5, line_dash="dash", line_color="#aaaaaa")
    fig_quad_sku.add_hline(y=shop_total_tacos, line_dash="dash", line_color="#ff7700",annotation_text=f"店铺TACOS {shop_total_tacos:.1%}")
    fig_quad_sku.update_layout(
        title=f"{select_month} SKU四象限散点",
        xaxis_title="单品ASoAS(广告销售占该SKU总销售)", xaxis_tickformat=".1%",
        yaxis_title="单品TACOS", yaxis_tickformat=".1%",
        height=420
    )
    st.plotly_chart(fig_quad_sku, use_container_width=True)

st.markdown("""
> 📍四象限解读
> 1️⃣右上【高ASoAS+高TACOS】风险区：高度依赖广告，广告成本还很高，重点优化
> 2️⃣左上【高ASoAS+低TACOS】：广告效率好，但过度依赖广告，需要做自然流量
> 3️⃣右下【低ASoAS+高TACOS】：广告花钱，但广告贡献销售占比很低，广告浪费严重
> 4️⃣左下【低ASoAS+低TACOS】健康区：自然流量强，广告成本可控
""")


st.divider()

# ===================== 八、自定义TACOS预算模拟（HTML前端实时计算，按历史广告花费占比分配老品预算） =====================
st.markdown("## 📈 八、自定义TACOS预算模拟")
shop_total_tacos = df_month_single["TACOS广告花费占比"].iloc[0]
# ========== session状态存储 ==========
if "sim_df_result" not in st.session_state:
    st.session_state.sim_df_result = None
if "sim_summary" not in st.session_state:
    st.session_state.sim_summary = None
if "html_json_rows" not in st.session_state:
    st.session_state.html_json_rows = None
if "html_global_t" not in st.session_state:
    st.session_state.html_global_t = None

df_8020_raw = df_all_item.dropna(subset=["销售额"]).copy()
df_8020_raw = df_8020_raw[df_8020_raw["销售额"] > 0]
if df_8020_raw.empty:
    st.warning("本月无有效销售额SKU，无法计算")
else:
    df_8020_sort = df_8020_raw.sort_values("销售额", ascending=False).reset_index(drop=True)

    def get_flow_tag(row):
        sales = row["销售额"]
        ad_sales = row["广告销售额"]
        ad_spend = row["广告花费"]
        prod_type = row["产品类型"]
        if "新品" in prod_type:
            return "新品推广款（刚性预算）"
        if ad_sales > sales and sales > 0:
            return "数据异常"
        if ad_spend > 0 and ad_sales == 0:
            return "纯自然出单"
        if sales > 0 and (ad_sales / sales) >= 0.95:
            return "重度广告依赖"
        return "正常老品"

    df_8020_sort["商品流量标签"] = df_8020_sort.apply(get_flow_tag, axis=1)

    st.subheader("⚙️ 第一步：设置全局目标TACOS")
    st.markdown("""
> 📘**模拟规则（销售额固定，新品预算固定不参与分配）**
> 1. 全店允许总广告花费 = 全店总销售额 ×全局目标TACOS
> 2. 老品预算池 = max(0,全店允许总广告花费 −新品刚性广告费)
> 3. **老品基准分配规则：按照【老品历史实际广告花费占老品总广告花费比例】分配老品预算池**
>    - 单品基准广告花费 = 老品预算池 × (单品老品实际广告花费 ÷全部老品实际广告花费)
>    - 单品基准TACOS = 单品基准广告花费 ÷SKU销售额 ×100，每个老品SKU基准TACOS各不相同
> 4. 修改后广告花费 = SKU销售额 ×手动输入的修改TACOS
> 5. 花费差值 = 修改后广告花费 − 基准广告花费（多花为正）
> 💡微小±0.01属于计算机浮点精度误差。
""")
    col1, col2 = st.columns([4,1])
    with col1:
        global_t = st.number_input(
            "全局目标TACOS(%)", min_value=0.0, max_value=50.0,
            value=15.0, step=0.5,
            help="全局目标TACOS，新品广告花费保持不变；基准预算老品按历史广告花费占比分配。运营修改TACOS为独立测算，用于对比"
        )
    with col2:
        reset_sim = st.button("🔄 重置模拟", type="secondary")
        if reset_sim:
            st.session_state.html_json_rows = None
            st.session_state.html_global_t = None
            st.session_state.sim_df_result = None
            st.session_state.sim_summary = None
            st.rerun()

    start_calc = st.button("✅ 确认参数，生成模拟表格", type="primary")

    if start_calc:
        import json
        df_html = df_8020_sort.copy()
        df_html["单品当前TACOS"] = df_html.apply(
            lambda r: round(r["广告花费"]/r["销售额"]*100,2) if r["销售额"]>0 else 0, axis=1
        )
        json_rows = df_html[["MSKU","品名","产品类型","分层","商品流量标签","销售额","单品当前TACOS","广告花费"]].to_json(orient="records", force_ascii=False)
        st.session_state.html_json_rows = json_rows
        st.session_state.html_global_t = str(global_t)

    if st.session_state.html_json_rows is not None:
        import json
        json_rows = st.session_state.html_json_rows
        gt_val = st.session_state.html_global_t

        html_tpl = '''
        <div style="font-size:13px;">
        <div style="display:flex;gap:15px;flex-wrap:wrap;margin-bottom:15px;padding:10px;background:#f6f8fa;border-radius:6px;">
          <div><b>当前选定全局TACOS：</b><span id="sumT" style="color:#165DFF;font-weight:bold;">__GT__%</span></div>
          <div><b>全店允许总广告花费：</b>$<span id="sumTotalAllow" style="color:#00B42A;font-weight:bold;">0</span></div>
          <div><b>扣新品后老品预算池上限：</b>$<span id="sumOldAllow" style="color:#00B42A;font-weight:bold;">0</span></div>
          <div><b>基准老品总花费：</b>$<span id="baseOldTotal" style="color:#722ED1;font-weight:bold;">0</span></div>
          <div><b>修改后模拟老品花费：</b>$<span id="editOldTotal" style="color:#722ED1;font-weight:bold;">0</span></div>
          <div><b>修改后模拟总花费(含新品)：</b>$<span id="editTotalSpend" style="color:#722ED1;font-weight:bold;">0</span></div>
          <div><b>修改后整体TACOS：</b><span id="editWholeTacos" style="color:#F53F3F;font-weight:bold;">0%</span></div>
        </div>
        <div style="max-height:400px;overflow:auto;">
        <table border="1" cellpadding="4" cellspacing="0" style="width:100%;border-collapse:collapse;">
        <thead>
        <tr style="background:#e5e6eb;">
        <th>MSKU</th><th>品名</th><th>产品类型</th><th>分层</th><th>商品流量标签</th>
        <th>销售额</th><th>当前实际TACOS(%)</th><th>当前实际广告花费</th>
        <th>单品目标TACOS(%)<br />【基准只读】</th>
        <th>单品目标TACOS广告花费<br />【基准只读】</th>
        <th>修改的TACOS(%)<br />【可编辑】</th>
        <th>修改TACOS的广告花费</th>
        <th>花费差值(修改-基准)<br />【多花为正】</th>
        </tr>
        </thead>
        <tbody id="tb"></tbody>
        </table>
        </div>
        <br />
        <button onclick="saveData()" style="background:#165DFF;color:white;padding:6px 14px;border-radius:4px;border:none;">✅ 回传结果到看板（一键复制JSON）</button>
        <div id="tip" style="color:#00875a;display:none;margin:8px 0;">✅JSON已复制到剪贴板！粘贴到下方输入框，点击解析按钮即可</div>
        </div>
        <script>
        const rows = JSON.parse(`__ROWS__`);
        const gT = Number("__GT__");
        const totalSales = rows.reduce((s,r)=>s+Number(r["销售额"]),0);
        const tb = document.getElementById("tb");

        let sumNewAd = 0;
        let oldRows = [];
        rows.forEach((row)=>{
            const isNew = row["商品流量标签"].includes("新品");
            if(isNew){
                sumNewAd += Number(row["广告花费"]);
                row.baseSpend = Number(row["广告花费"]);
                row.baseTacos = null;
            }else{
                oldRows.push(row);
            }
        });
        const totalAllow = totalSales * gT / 100;
        const oldBudgetPool = Math.max(0, totalAllow - sumNewAd);

        // ============ 修改：按老品历史实际广告花费占比分配预算池 ============
        const sumOldRealAdSpend = oldRows.reduce((s, r)=> s + Number(r["广告花费"]),0);

        oldRows.forEach(r=>{
            const sales = Number(r["销售额"]);
            const realAdSpend = Number(r["广告花费"]);
            let baseSpend;
            // 兜底：老品全部历史广告花费=0，则老品预算池平均分摊
            if(sumOldRealAdSpend > 0){
                baseSpend = oldBudgetPool * (realAdSpend / sumOldRealAdSpend);
            }else{
                baseSpend = oldBudgetPool / oldRows.length;
            }
            // 反算单品基准TACOS（每个老品SKU各自独立）
            const baseTacos = sales > 0 ? (baseSpend / sales)*100 : 0;

            r.baseTacos = baseTacos;
            r.baseSpend = baseSpend;
            r.editTacos = r.baseTacos;
            r.editSpend = r.baseSpend;
        })

        function recalc(){
          tb.innerHTML = "";
          let sumEditTotal = sumNewAd;
          let sumBaseOld = 0;
          let sumEditOld = 0;
          rows.forEach((row,idx)=>{
            const tr = document.createElement("tr");
            const isNew = row["商品流量标签"].includes("新品");
            const sales = Number(row["销售额"]);
            const curAd = Number(row["广告花费"]);
            const curTacos = Number(row["单品当前TACOS"]);
            let baseTacos = row.baseTacos;
            let baseSpend = row.baseSpend;
            let editTacos = row.editTacos;
            let editSpend = 0;
            let diff = 0;

            if(isNew){
                editTacos = null;
                editSpend = baseSpend;
                diff = 0;
            }else{
                editTacos = Number(row.editTacos);
                editSpend = sales * editTacos / 100;
                diff = editSpend - baseSpend;

                editSpend = Math.round(editSpend *100)/100;
                baseSpend = Math.round(baseSpend *100)/100;
                diff = Math.round(diff *100)/100;

                row.editSpend = editSpend;
                sumEditTotal += editSpend;
                sumBaseOld += baseSpend;
                sumEditOld += editSpend;
            }

            let diffColor = "#000000";
            if(diff > 1) diffColor = "#c41e3a";
            if(diff < -1) diffColor = "#00875a";
            tr.innerHTML = `
        <td>${row.MSKU||""}</td>
        <td>${(row.品名||"").replace(/[<>]/g,"")}</td>
        <td>${row.产品类型||""}</td>
        <td>${row.分层||""}</td>
        <td>${row.商品流量标签||""}</td>
        <td>$${sales.toFixed(2)}</td>
        <td>${curTacos.toFixed(2)}%</td>
        <td>$${curAd.toFixed(2)}</td>
        <td>${isNew ? "-" : baseTacos.toFixed(2)+"%"}</td>
        <td>${isNew ? "-" : "$"+baseSpend.toFixed(2)}</td>
        <td>${isNew ? "-" : `<input type="number" min=0 max=50 step=0.5 style="width:70px" value="${editTacos.toFixed(2)}" onchange="updateTacos(${idx},this.value)">`}</td>
        <td>${isNew ? "-" : "$"+editSpend.toFixed(2)}</td>
        <td style="color:${diffColor};font-weight:bold">${isNew ? "-" : diff.toFixed(2)}</td>
        `;
            tb.appendChild(tr);
          })
          let editWholeTacos = totalSales>0 ? (sumEditTotal / totalSales *100) :0;
          document.getElementById("sumTotalAllow").innerText = totalAllow.toFixed(2);
          document.getElementById("sumOldAllow").innerText = oldBudgetPool.toFixed(2);
          document.getElementById("baseOldTotal").innerText = sumBaseOld.toFixed(2);
          document.getElementById("editOldTotal").innerText = sumEditOld.toFixed(2);
          document.getElementById("editTotalSpend").innerText = sumEditTotal.toFixed(2);
          document.getElementById("editWholeTacos").innerText = editWholeTacos.toFixed(2)+"%";
        }

        function updateTacos(idx,val){
          rows[idx].editTacos = Number(val);
          recalc();
        }

        function roundObj2(obj) {
            const newObj = {...obj};
            for(let k in newObj) {
                const v = newObj[k];
                if(typeof v === 'number'){
                    newObj[k] = Math.round(v * 100) / 100;
                }
            }
            return newObj;
        }

        async function saveData(){
          const payloadRaw = {
            global_t:gT,
            total_sales:totalSales,
            sum_new_ad: sumNewAd,
            rows: rows.map(roundObj2)
          };
          const payload = JSON.stringify(payloadRaw);
          await navigator.clipboard.writeText(payload);
          document.getElementById("tip").style.display="block";
          setTimeout(()=>{document.getElementById("tip").style.display="none"},3000)
        }
        recalc();
        </script>
        '''

        html_code = html_tpl.replace("__GT__", gt_val).replace("__ROWS__", json_rows)
        # 调试：打印前300字符，确认占位符被替换
        # st.code(html_code[:300])
        st.components.v1.html(html_code, height=600, scrolling=True)

        st.info("👉调整完单品TACOS，点击【回传结果到看板（一键复制JSON）】，复制完成后粘贴到下方输入框，再点【解析回传结果】")
        json_input = st.text_area("粘贴回传JSON", height=120)
        parse_btn = st.button("📥 解析回传结果", type="primary")
        if parse_btn:
            try:
                payload = json.loads(json_input.strip())
                df_res = pd.DataFrame(payload["rows"])

                # ========= Python直接复用JS回传结果，禁止二次重算花费 =========
                df_res["base_spend"] = df_res["baseSpend"]
                df_res["base_tacos"] = df_res["baseTacos"]
                df_res["edit_tacos"] = df_res["editTacos"]
                df_res["edit_spend"] = df_res["editSpend"]  # 使用JS已经计算完成的值

                df_res["diff_spend"] = (df_res["edit_spend"] - df_res["base_spend"]).round(2)

                numeric_cols = df_res.select_dtypes(include=["float", "int"]).columns
                df_res[numeric_cols] = df_res[numeric_cols].round(2)

                st.session_state.sim_df_result = df_res
                st.session_state.sim_summary = {
                    "global_t": payload["global_t"],
                    "total_sales": payload["total_sales"],
                    "sum_new_ad": payload["sum_new_ad"]
                }
                st.success("✅解析成功！上方编辑表格保留，下方展示模拟结果，可继续修改TACOS重新测算")
                st.rerun()
            except Exception as e:
                st.error(f"解析失败：{e}")

        if st.session_state.sim_df_result is not None:
            st.divider()
            st.subheader("📊 本次TACOS模拟回传结果")
            sm = st.session_state.sim_summary
            df_res = st.session_state.sim_df_result.copy()

            total_sales = sm["total_sales"]
            sum_new_ad = sm["sum_new_ad"]
            base_global_t = sm["global_t"]
            base_total_budget = total_sales * base_global_t / 100
            base_old_budget = base_total_budget - sum_new_ad

            df_res["is_new_flag"] = df_res["edit_tacos"].isna()
            sum_edit_old_spend = df_res[~df_res["is_new_flag"]]["edit_spend"].sum()

            sim_total_budget = sum_edit_old_spend + sum_new_ad
            sim_global_tacos = sim_total_budget / total_sales * 100 if total_sales > 0 else 0

            st.markdown("##### 基准参数")
            bc1, bc2, bc3, bc4, bc5 = st.columns(5)
            with bc1:
                st.metric("全局目标TACOS", f"{base_global_t}%")
            with bc2:
                st.metric("全店总销售额", f"${total_sales:,.2f}")
            with bc3:
                st.metric("全店总预算广告花费", f"${base_total_budget:,.2f}")
            with bc4:
                st.metric("新品刚性广告费", f"${sum_new_ad:,.2f}")
            with bc5:
                st.metric("老品预算广告花费", f"${base_old_budget:,.2f}")

            st.markdown("##### 模拟后参数")
            sc1, sc2, sc3, sc4, sc5 = st.columns(5)
            with sc1:
                st.metric("全局目标TACOS", f"{sim_global_tacos:.2f}%")
            with sc2:
                st.metric("全店总销售额", f"${total_sales:,.2f}")
            with sc3:
                st.metric("全店总预算广告花费", f"${sim_total_budget:,.2f}",
                          delta=f"{sim_total_budget - base_total_budget:.2f}")
            with sc4:
                st.metric("新品刚性广告费", f"${sum_new_ad:,.2f}")
            with sc5:
                st.metric("老品预算广告花费", f"${sum_edit_old_spend:,.2f}",
                          delta=f"{sum_edit_old_spend - base_old_budget:.2f}")

            df_show = st.session_state.sim_df_result.copy()

            df_show = df_show.rename(columns={
                "单品当前TACOS": "当前实际TACOS(%)",
                "base_tacos": "单品目标TACOS(%)【基准只读】",
                "base_spend": "单品目标TACOS广告花费【基准只读】",
                "edit_tacos": "修改的TACOS(%)【可编辑】",
                "edit_spend": "修改TACOS的广告花费",
                "diff_spend": "花费差值(修改-基准)【多花为正】",
                "广告花费": "当前实际广告花费"
            })

            for col in df_show.columns:
                df_show[col] = pd.to_numeric(df_show[col], errors='coerce').fillna(df_show[col])
            df_show = df_show.round(2)

            format_dict = {}
            for col in df_show.columns:
                if pd.api.types.is_numeric_dtype(df_show[col]):
                    format_dict[col] = "{:.2f}"

            show_cols = [
                "MSKU", "品名", "产品类型", "分层", "商品流量标签",
                "销售额", "当前实际TACOS(%)", "当前实际广告花费",
                "单品目标TACOS(%)【基准只读】", "单品目标TACOS广告花费【基准只读】",
                "修改的TACOS(%)【可编辑】", "修改TACOS的广告花费",
                "花费差值(修改-基准)【多花为正】"
            ]
            show_cols = [c for c in show_cols if c in df_show.columns]
            df_show = df_show[show_cols].sort_values("销售额", ascending=False)

            def color_diff(s):
                colors = []
                for v in s:
                    if v == "-" or pd.isna(v):
                        colors.append("")
                    elif float(v) > 1:
                        colors.append("background-color: #ffcccc; color: #c41e3a; font-weight:bold")
                    elif float(v) < -1:
                        colors.append("background-color: #d4edda; color: #00875a; font-weight:bold")
                    else:
                        colors.append("")
                return colors

            if "花费差值(修改-基准)【多花为正】" in df_show.columns:
                styled_df = df_show.style.format(format_dict).apply(color_diff,
                                                                    subset=["花费差值(修改-基准)【多花为正】"])
            else:
                styled_df = df_show.style

            st.dataframe(styled_df, use_container_width=True, height=450)
        st.divider()