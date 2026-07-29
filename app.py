"""
网红团购 · 价格链路自动化工具
网红团购一站式平台 - 第一个模块
"""

import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import re

# ─────────────────────────────────────────────
# 页面配置
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="价格链路自动化 | 网红团购平台",
    page_icon="💰",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────
# 自定义样式
# ─────────────────────────────────────────────
st.markdown("""
<style>
    .main-header {
        font-size: 1.6rem;
        font-weight: 700;
        margin-bottom: 0.2rem;
    }
    .sub-header {
        color: #666;
        font-size: 0.9rem;
        margin-bottom: 1.5rem;
    }
    .metric-card {
        background: #f8f9fa;
        border-radius: 10px;
        padding: 1rem;
        text-align: center;
    }
    .error-cell {
        background-color: #ffe0e0 !important;
    }
    .warn-cell {
        background-color: #fff3cd !important;
    }
    .ok-cell {
        background-color: #d4edda !important;
    }
    div[data-testid="stSidebar"] {
        background: #fafbfc;
    }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# 侧边栏：全局参数配置
# ─────────────────────────────────────────────
with st.sidebar:
    st.markdown("### ⚙️ 全局参数")
    st.caption("规则暂定，可随时调整")

    exchange_rate = st.number_input(
        "汇率 (KRW → USD)",
        min_value=1000.0,
        max_value=2000.0,
        value=1550.0,
        step=10.0,
        help="韩元÷此汇率=美金。当前约1550",
    )

    st.markdown("---")
    st.markdown("**补贴上限规则**")

    brand_plus_pct = st.slider(
        "百补比例 (Brand+)",
        min_value=0,
        max_value=15,
        value=5,
        step=1,
        format="%d%%",
        help="全托管 / POP半托 brand+ 商品的百补补贴",
    )
    brand_plus_rate = brand_plus_pct / 100.0

    code_cap_pct = st.slider(
        "Code上限 (占原价)",
        min_value=10,
        max_value=35,
        value=20,
        step=1,
        format="%d%%",
        help="网红code补贴占报名原价的上限",
    )
    code_cap_rate = code_cap_pct / 100.0

    total_cap_pct = st.slider(
        "总补贴上限",
        min_value=15,
        max_value=40,
        value=25,
        step=1,
        format="%d%%",
        help="百补+code+店铺券 叠加后不能超过此比例",
    )
    total_cap_rate = total_cap_pct / 100.0

    st.markdown("---")
    st.markdown("**Code自动计算**")

    auto_code = st.checkbox(
        "自动计算Code金额",
        value=True,
        help="根据目标总补贴比例自动倒推code金额（向下取到.5）。关闭则使用表格中已有的code。",
    )

    target_subsidy_pct = st.slider(
        "目标总补贴比例",
        min_value=15,
        max_value=25,
        value=24,
        step=1,
        format="%d%%",
        help="工具按此比例倒推code。设24%留1个点余量，需要时可手动调到25%",
        disabled=not auto_code,
    )
    target_subsidy_rate = target_subsidy_pct / 100.0

    st.markdown("---")
    st.markdown("**校验规则**")
    check_code_round = st.checkbox(
        "Code金额须为整数或.5",
        value=True,
        help="例如 5, 5.5, 10 合法；5.3 不合法",
    )
    check_price_vs_external = st.checkbox(
        "比价：AE最终价 vs 站外最低价",
        value=True,
    )

# ─────────────────────────────────────────────
# 主区域
# ─────────────────────────────────────────────
st.markdown('<div class="main-header">💰 价格链路自动化</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">网红团购一站式平台 · 模块一：上传商品表 → 自动计算 → 校验 → 导出</div>', unsafe_allow_html=True)

# ─────────────────────────────────────────────
# Step 1: 上传文件
# ─────────────────────────────────────────────
uploaded_file = st.file_uploader(
    "上传价格链路表 (.xlsx)",
    type=["xlsx", "xls"],
    help="支持你现有的价格链路表格式，系统会自动识别列",
)

if uploaded_file is None:
    st.info("👆 上传你的价格链路 Excel 表，系统会自动完成所有计算和校验。")
    st.markdown("""
    **工具会自动完成：**
    - 百补金额、页面价、最终价格、GMV、ROI 等全部公式计算
    - Code金额校验（整数/.5）
    - 叠加补贴是否超上限
    - 站外比价（韩元→美金换算 + 高低判断）
    - 异常标红提醒
    - 一键导出完整表格
    """)
    st.stop()

# ─────────────────────────────────────────────
# Step 2: 读取并识别列
# ─────────────────────────────────────────────
@st.cache_data
def load_excel(file_bytes):
    """读取Excel，尝试识别表头行"""
    df_raw = pd.read_excel(BytesIO(file_bytes), header=None)
    # 找到表头行：包含"商品报名原价"或"노미네이션가"的行
    header_row = None
    for i in range(min(5, len(df_raw))):
        row_text = " ".join([str(x) for x in df_raw.iloc[i].tolist() if pd.notna(x)])
        if "报名原价" in row_text or "노미네이션가" in row_text or "商品ID" in row_text:
            header_row = i
            break

    if header_row is None:
        # 默认第2行(index 1)是表头
        header_row = 1

    df = pd.read_excel(BytesIO(file_bytes), header=header_row)
    return df, header_row


df, header_row = load_excel(uploaded_file.getvalue())

# 列名映射：支持多种可能的列名
COLUMN_ALIASES = {
    "负责人": ["负责人", "메인"],
    "对接人": ["对接人", "서브"],
    "频道名": ["频道名", "채널명"],
    "组": ["组", "조"],
    "团购时间": ["团购时间", "영상 업로드"],
    "是否brand+": ["是否brand+商品", "brand+"],
    "付费商品": ["付费商品", "추천리스트"],
    "供给类型": ["供给类型", "공급 유형"],
    "商品ID": ["商品ID", "상품ID"],
    "商品名": ["商品名", "상품명"],
    "承接SKU_ID": ["承接SKU ID", "진행 상품 SKU"],
    "SKU": ["SKU", "옵션"],
    "数量": ["数量", "진행 수량"],
    "行业反馈库存": ["行业反馈库存"],
    "报名原价": ["商品报名原价", "노미네이션가"],
    "百补金额": ["百补补贴金额", "빅세이브 할인금액"],
    "百补力度": ["百补补贴力度", "brand+ 할인율"],
    "叠加补贴力度": ["叠加补贴力度", "중복 할인율"],
    "页面价": ["页面价", "상세페이지가격"],
    "店铺券": ["店铺券", "스토어 쿠폰"],
    "店铺券CODE": ["定向发放店铺券CODE", "스토어 쿠폰 네임"],
    "门槛美金": ["门槛美金", "허들"],
    "code金额": ["code补贴美金", "코드금액"],
    "code预算": ["code预算", "코드총예산"],
    "折扣率": ["折扣率", "할인율"],
    "最终价格": ["最终价格", "최종할인가"],
    "GMV": ["GMV"],
    "ROI": ["ROI"],
    "站外美金": ["站外价格（美金）", "국내 최저가(달러)"],
    "站外韩元": ["站外价格（韩元）", "국내 최저가(원)"],
    "站外链接": ["站外比价链接", "국내 최저가 링크"],
    "比价结果": ["比价结果", "가격 비교 결과"],
}


def find_column(df, key):
    """根据别名找到实际列名"""
    aliases = COLUMN_ALIASES.get(key, [key])
    for col in df.columns:
        col_str = str(col).replace("\n", " ").strip()
        for alias in aliases:
            if alias in col_str:
                return col
    return None


# 建立列映射
col_map = {}
for key in COLUMN_ALIASES:
    found = find_column(df, key)
    if found:
        col_map[key] = found

# 显示识别结果
with st.expander("📋 列识别结果（点击展开检查）", expanded=False):
    identified = len(col_map)
    total = len(COLUMN_ALIASES)
    st.write(f"已识别 {identified}/{total} 列")
    for key, col in col_map.items():
        st.write(f"  ✅ {key} → `{col}`")
    missing = set(COLUMN_ALIASES.keys()) - set(col_map.keys())
    if missing:
        st.write(f"  ⚠️ 未识别: {', '.join(missing)}")

# ─────────────────────────────────────────────
# Step 3: 自动计算
# ─────────────────────────────────────────────
# 确保数值列为数字类型
numeric_cols = ["报名原价", "数量", "店铺券", "code金额", "站外韩元", "站外美金"]
for key in numeric_cols:
    if key in col_map:
        col = col_map[key]
        df[col] = pd.to_numeric(df[col], errors="coerce")

# 获取关键列名
col_price = col_map.get("报名原价")
col_qty = col_map.get("数量")
col_coupon = col_map.get("店铺券")
col_code = col_map.get("code金额")
col_krw = col_map.get("站外韩元")
col_usd_ext = col_map.get("站外美金")
col_brand = col_map.get("是否brand+")
col_supply = col_map.get("供给类型")
col_name = col_map.get("商品名")
col_result = col_map.get("比价结果")

if not col_price:
    st.error("❌ 未找到「报名原价」列，请检查表格格式。")
    st.stop()

# 计算各列
# 百补力度：根据供给类型和brand+判断
if col_brand and col_supply:
    def get_brand_rate(row):
        supply = str(row.get(col_supply, "")).strip()
        brand = str(row.get(col_brand, "")).strip().upper()
        # POP/半托管 且 非brand+ → 无百补
        if ("POP" in supply or "半托" in supply) and brand != "Y":
            return 0.0
        return brand_plus_rate
    df["_百补力度"] = df.apply(get_brand_rate, axis=1)
else:
    df["_百补力度"] = brand_plus_rate

# 百补金额 = 原价 × 百补力度
df["_百补金额"] = df[col_price] * df["_百补力度"]

# 页面价 = 原价 - 百补金额
df["_页面价"] = df[col_price] - df["_百补金额"]

# 店铺券（默认为0）
if col_coupon:
    coupon_values = df[col_coupon].fillna(0)
else:
    coupon_values = pd.Series(0, index=df.index)

# Code金额：自动计算 或 使用表格已有值
if auto_code:
    # 倒推公式: code = 目标比例 × (原价 - 店铺券) - 百补金额
    raw_code = target_subsidy_rate * (df[col_price] - coupon_values) - df["_百补金额"]
    # 向下取到最近的0.5（例如 5.75 → 5.5, 2.3 → 2.0, 4.99 → 4.5）
    code_values = (np.floor(raw_code * 2) / 2).clip(lower=0)
else:
    if col_code:
        code_values = df[col_code].fillna(0)
    else:
        code_values = pd.Series(0, index=df.index)

# 最终价格 = 页面价 - 店铺券 - code
df["_最终价格"] = df["_页面价"] - coupon_values - code_values

# Code预算 = 数量 × code金额
if col_qty:
    df["_code预算"] = df[col_qty].fillna(0) * code_values
else:
    df["_code预算"] = code_values

# GMV = 数量 × 页面价
if col_qty:
    df["_GMV"] = df[col_qty].fillna(0) * df["_页面价"]
else:
    df["_GMV"] = df["_页面价"]

# ROI = GMV / code预算
df["_ROI"] = np.where(df["_code预算"] > 0, df["_GMV"] / df["_code预算"], np.nan)

# 折扣率 = code / 页面价
df["_折扣率"] = np.where(df["_页面价"] > 0, code_values / df["_页面价"], 0)

# 叠加补贴力度 = (百补金额 + code) / (原价 - 店铺券)
denominator = df[col_price] - coupon_values
df["_叠加补贴"] = np.where(
    denominator > 0,
    (df["_百补金额"] + code_values) / denominator,
    0,
)

# 站外美金价（如果只有韩元，自动换算）
if col_krw and col_usd_ext:
    # 优先用已有的美金价，没有则从韩元换算
    df["_站外美金"] = df[col_usd_ext].copy()
    mask_need_convert = df["_站外美金"].isna() & df[col_krw].notna()
    df.loc[mask_need_convert, "_站外美金"] = df.loc[mask_need_convert, col_krw] / exchange_rate
elif col_krw:
    df["_站外美金"] = df[col_krw] / exchange_rate
elif col_usd_ext:
    df["_站外美金"] = df[col_usd_ext]
else:
    df["_站外美金"] = np.nan

# 比价结果（站外价为0或空视为未填写，不做比价）
_ext_invalid = df["_站外美金"].isna() | (df["_站外美金"] == 0)
df["_比价结果"] = np.where(
    _ext_invalid | df["_最终价格"].isna(),
    "",
    np.where(df["_最终价格"] > df["_站外美金"], "AE价高 ⚠️",
             np.where(df["_最终价格"] < df["_站外美金"], "AE价低 ✅", "同价")),
)

# ─────────────────────────────────────────────
# Step 4: 校验
# ─────────────────────────────────────────────
errors = []
warnings = []

for idx, row in df.iterrows():
    row_num = idx + 2 + header_row  # Excel行号
    name = str(row.get(col_name, f"第{row_num}行"))[:20] if col_name else f"第{row_num}行"

    # 校验1: code金额必须为整数或.5
    if check_code_round:
        if auto_code:
            code_val = code_values.get(idx, 0)
        elif col_code:
            code_val = row.get(col_code)
        else:
            code_val = 0
        if pd.notna(code_val) and code_val != 0:
            remainder = code_val % 0.5
            if abs(remainder) > 0.001 and abs(remainder - 0.5) > 0.001:
                errors.append(f"行{row_num} [{name}]: code金额 {code_val} 不是整数或.5")

    # 校验2: 叠加补贴不能超上限
    subsidy = row.get("_叠加补贴", 0)
    if pd.notna(subsidy) and subsidy > total_cap_rate + 0.001:
        errors.append(
            f"行{row_num} [{name}]: 叠加补贴 {subsidy:.1%} 超过上限 {total_cap_rate:.0%}"
        )

    # 校验3: 比价 - AE价高
    if check_price_vs_external:
        result = row.get("_比价结果", "")
        if "AE价高" in str(result):
            warnings.append(
                f"行{row_num} [{name}]: AE最终价 ${row['_最终价格']:.2f} > 站外 ${row['_站外美金']:.2f}"
            )

    # 校验4: 报名原价异常（非数字或为0）
    price_val = row.get(col_price)
    if pd.isna(price_val) or price_val == 0:
        errors.append(f"行{row_num} [{name}]: 报名原价为空或为0")

    # 校验5: 最终价格为负
    final_price = row.get("_最终价格")
    if pd.notna(final_price) and final_price < 0:
        errors.append(f"行{row_num} [{name}]: 最终价格为负 ({final_price:.2f})")

    # 校验6: 报名原价异常高（疑似韩币误填入美金列）
    if pd.notna(price_val) and price_val > 2000:
        corrected = price_val / exchange_rate
        errors.append(
            f"行{row_num} [{name}]: 报名原价 {price_val:,.0f} 异常高，"
            f"疑似韩币误填入美金列（若为韩币则约 ${corrected:.2f}）"
        )
    elif pd.notna(price_val) and price_val > 500:
        corrected = price_val / exchange_rate
        warnings.append(
            f"行{row_num} [{name}]: 报名原价 ${price_val:,.0f} 较高，"
            f"请确认是美金（若为韩币则约 ${corrected:.2f}）"
        )

# ─────────────────────────────────────────────
# Step 5: 展示结果
# ─────────────────────────────────────────────
st.markdown("---")

# 概览指标
col1, col2, col3, col4, col5 = st.columns(5)
total_items = len(df)
total_gmv = df["_GMV"].sum()
total_budget = df["_code预算"].sum()
avg_roi = df["_ROI"].mean() if df["_ROI"].notna().any() else 0
price_high_count = df["_比价结果"].str.contains("AE价高", na=False).sum()

col1.metric("商品数", f"{total_items}")
col2.metric("预估总GMV", f"${total_gmv:,.0f}")
col3.metric("Code总预算", f"${total_budget:,.0f}")
col4.metric("平均ROI", f"{avg_roi:.1f}x")
col5.metric("AE价高商品", f"{price_high_count} 个", delta=f"-{price_high_count}" if price_high_count > 0 else "0", delta_color="inverse")

# 校验结果
st.markdown("---")
if errors:
    st.error(f"🚨 发现 {len(errors)} 个错误（必须修正）")
    for e in errors[:10]:
        st.write(f"  ❌ {e}")
    if len(errors) > 10:
        st.write(f"  ... 还有 {len(errors)-10} 个错误")

if warnings:
    st.warning(f"⚠️ {len(warnings)} 个比价警告（AE价格高于站外）")
    for w in warnings[:10]:
        st.write(f"  ⚠️ {w}")
    if len(warnings) > 10:
        st.write(f"  ... 还有 {len(warnings)-10} 个警告")

if not errors and not warnings:
    st.success("✅ 全部校验通过，无异常！")

# 按网红分组查看
st.markdown("---")
st.markdown("### 📊 按网红分组")

col_channel = col_map.get("频道名")
col_owner = col_map.get("负责人")

if col_channel:
    channels = df[col_channel].dropna().unique().tolist()
    # 向前填充频道名（因为可能只在第一行出现）
    df["_频道名_filled"] = df[col_channel].ffill()
    channels = df["_频道名_filled"].dropna().unique().tolist()

    if channels:
        selected_channel = st.selectbox("选择网红频道", ["全部"] + channels)
        if selected_channel != "全部":
            df_view = df[df["_频道名_filled"] == selected_channel].copy()
        else:
            df_view = df.copy()
    else:
        df_view = df.copy()
else:
    df_view = df.copy()

# 重合商品检测
col_pid = col_map.get("商品ID")
if col_pid and col_channel:
    product_counts = df.groupby(col_pid)[col_pid].transform("count")
    df["_重合"] = product_counts > 1
    overlap_count = df["_重合"].sum()
    if overlap_count > 0:
        st.info(f"📌 有 {overlap_count} 条记录涉及重合商品（同一商品被多个网红选中）")

# 展示计算结果表
st.markdown("### 📋 计算结果")

# 构建展示用的DataFrame
display_cols = {}
if col_name:
    display_cols["商品名"] = df_view[col_name]
if col_pid:
    display_cols["商品ID"] = df_view[col_pid]
if col_map.get("SKU"):
    display_cols["SKU"] = df_view[col_map["SKU"]]
if col_qty:
    display_cols["数量"] = df_view[col_qty]
if col_supply:
    display_cols["供给类型"] = df_view[col_supply]
if col_brand:
    display_cols["Brand+"] = df_view[col_brand]

display_cols["报名原价($)"] = df_view[col_price]
display_cols["百补金额($)"] = df_view["_百补金额"].round(2)
display_cols["页面价($)"] = df_view["_页面价"].round(2)
if col_coupon:
    display_cols["店铺券($)"] = df_view[col_coupon]
display_cols["Code金额($)"] = code_values.loc[df_view.index]
display_cols["最终价格($)"] = df_view["_最终价格"].round(2)
display_cols["Code预算($)"] = df_view["_code预算"].round(0)
display_cols["GMV($)"] = df_view["_GMV"].round(0)
display_cols["ROI"] = df_view["_ROI"].round(2)
display_cols["叠加补贴"] = (df_view["_叠加补贴"] * 100).round(1).astype(str) + "%"
display_cols["站外价($)"] = df_view["_站外美金"].round(2)
display_cols["比价结果"] = df_view["_比价结果"]

df_display = pd.DataFrame(display_cols)

# 使用data_editor展示（可编辑）
edited_df = st.data_editor(
    df_display,
    width="stretch",
    hide_index=True,
    num_rows="fixed",
    column_config={
        "报名原价($)": st.column_config.NumberColumn(format="%.2f"),
        "百补金额($)": st.column_config.NumberColumn(format="%.2f"),
        "页面价($)": st.column_config.NumberColumn(format="%.2f"),
        "最终价格($)": st.column_config.NumberColumn(format="%.2f"),
        "Code预算($)": st.column_config.NumberColumn(format="%.0f"),
        "GMV($)": st.column_config.NumberColumn(format="%.0f"),
        "ROI": st.column_config.NumberColumn(format="%.2f"),
        "站外价($)": st.column_config.NumberColumn(format="%.2f"),
    },
)

# ─────────────────────────────────────────────
# Step 6: 导出
# ─────────────────────────────────────────────
st.markdown("---")
st.markdown("### 📥 导出")

col_a, col_b = st.columns(2)

with col_a:
    # 导出完整计算结果
    export_df = df.copy()
    # 把计算列写回
    if col_map.get("百补金额"):
        export_df[col_map["百补金额"]] = export_df["_百补金额"]
    if col_map.get("页面价"):
        export_df[col_map["页面价"]] = export_df["_页面价"]
    if col_map.get("最终价格"):
        export_df[col_map["最终价格"]] = export_df["_最终价格"]
    if col_map.get("code预算"):
        export_df[col_map["code预算"]] = export_df["_code预算"]
    if col_map.get("GMV"):
        export_df[col_map["GMV"]] = export_df["_GMV"]
    if col_map.get("ROI"):
        export_df[col_map["ROI"]] = export_df["_ROI"]
    if col_map.get("叠加补贴力度"):
        export_df[col_map["叠加补贴力度"]] = export_df["_叠加补贴"]
    if col_map.get("折扣率"):
        export_df[col_map["折扣率"]] = export_df["_折扣率"]
    if col_result:
        export_df[col_result] = export_df["_比价结果"]

    # 去掉内部辅助列
    internal_cols = [c for c in export_df.columns if c.startswith("_")]
    export_df = export_df.drop(columns=internal_cols)

    buffer = BytesIO()
    export_df.to_excel(buffer, index=False, engine="openpyxl")
    buffer.seek(0)

    st.download_button(
        label="📥 下载完整价格链路表（含计算结果）",
        data=buffer,
        file_name="价格链路_计算完成.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        width="stretch",
    )

with col_b:
    # 导出异常商品清单
    if errors or warnings:
        issues_data = []
        for e in errors:
            issues_data.append({"类型": "❌ 错误", "详情": e})
        for w in warnings:
            issues_data.append({"类型": "⚠️ 警告", "详情": w})
        df_issues = pd.DataFrame(issues_data)

        buffer2 = BytesIO()
        df_issues.to_excel(buffer2, index=False, engine="openpyxl")
        buffer2.seek(0)

        st.download_button(
            label="📥 下载异常清单",
            data=buffer2,
            file_name="价格链路_异常清单.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            width="stretch",
        )

# ─────────────────────────────────────────────
# 过往价格对比（从商品名中提取）
# ─────────────────────────────────────────────
if col_name:
    st.markdown("---")
    st.markdown("### 📈 过往价格对比")
    st.caption("自动从商品名中提取「X月最终价」记录，与本次最终价对比")

    history_records = []
    for idx, row in df.iterrows():
        name = str(row.get(col_name, ""))
        # 匹配 "3月最终价：30.98" 或 "3月最终价30.98" 或 "6月最终价：73.33"
        match = re.search(r"(\d+)月最终价[：:]?\s*([\d.]+)", name)
        if match:
            month = match.group(1)
            old_price = float(match.group(2))
            new_price = row.get("_最终价格", np.nan)
            if pd.notna(new_price):
                diff = new_price - old_price
                diff_pct = (diff / old_price * 100) if old_price > 0 else 0
                history_records.append({
                    "商品名": name[:30],
                    f"{month}月最终价": old_price,
                    "本次最终价": round(new_price, 2),
                    "差额": round(diff, 2),
                    "涨幅": f"{diff_pct:+.1f}%",
                    "状态": "⚠️ 涨价" if diff > 0.5 else ("✅ 降价" if diff < -0.5 else "≈ 持平"),
                })

    if history_records:
        df_history = pd.DataFrame(history_records)
        st.dataframe(df_history, width="stretch", hide_index=True)
        price_up = sum(1 for r in history_records if "涨价" in r["状态"])
        if price_up > 0:
            st.warning(f"有 {price_up} 个商品本次价格高于过往，建议关注")
    else:
        st.info("未在商品名中发现过往价格记录。")

# 页脚
st.markdown("---")
st.caption("网红团购一站式平台 · 价格链路模块 v1.0 | 数据仅在本地浏览器处理，不上传任何服务器")
