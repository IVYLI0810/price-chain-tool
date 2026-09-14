"""
独立网站 · 大促历史到手价查询
──────────────────────────────────────────────
Streamlit Cloud 独立入口：新建 app 时 Main file path 选 history_app.py，
即可得到独立网址，只含「历史数据查询」全部功能（SKU批量查价 / 总览 /
商品档案 / 品牌档案 / 网红档案 / 维度交叉 / 口径说明）。
与主站 app.py 共用同一套代码与底表，改一处两边同步生效。
"""

import streamlit as st

st.set_page_config(
    page_title="大促历史到手价查询",
    page_icon="🗂️",
    layout="wide",
    initial_sidebar_state="collapsed",
)

import theme        # 全局主题样式（薄荷绿 · 游戏像素 · 黑描边 · 硬阴影）
import history_hub  # 历史数据查询全部功能

theme.inject()

# ── 顶栏大标题（与主站历史页一致） ──
st.markdown(
    '<div class="title-bar"><span class="coin">🗂️</span>'
    '<span class="title-en">DATA</span>'
    '<span class="title-cn">历史数据查询</span>'
    '<span class="cursor-blink">▮</span></div>',
    unsafe_allow_html=True,
)
st.markdown(
    '<p class="subtitle">4 期大促数据 · 品牌 / 商品 / 网红 / 期次 · 定价的历史依据</p>',
    unsafe_allow_html=True,
)

history_hub.render()
