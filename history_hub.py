"""
历史数据模块 (history_hub)
─────────────────────────────────────────────
网红团购 · 历史大促数据中枢（数据依据）
维度：品牌 / 商品 / 网红 / 期次，任意 2 维或 3 维交叉查询
数据来源：4 期大促底表（11月/3月/6月/8月），详见口径说明
"""

import sqlite3
from pathlib import Path

import streamlit as st
import pandas as pd

PERIODS = ['11月', '3月', '6月', '8月']

_HERE = Path(__file__).parent
_DB_CANDIDATES = [_HERE / 'history_data.db', Path('history_data.db')]
DB_PATH = next((str(p) for p in _DB_CANDIDATES if p.exists()), None)


# ─────────────────────────────────────────────
# 数据访问层
# ─────────────────────────────────────────────
def _query(sql, params=None):
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    try:
        rows = conn.execute(sql, params or ()).fetchall()
        return pd.DataFrame([dict(r) for r in rows])
    finally:
        conn.close()


def _fmt_money(v):
    if v is None or pd.isna(v):
        return '-'
    return f'${v:,.0f}'


def _fmt_pct(v):
    if v is None or pd.isna(v):
        return '-'
    return f'{v:.1f}%'


def _fmt_num(v):
    if v is None or pd.isna(v):
        return '-'
    return f'{v:,.0f}'


@st.cache_data
def load_brands():
    return _query("SELECT * FROM brands ORDER BY total_gmv DESC")


@st.cache_data
def load_products():
    return _query("""
        SELECT p.gkey, p.std_name, p.band, p.product_ids,
               b.brand_name
        FROM products p LEFT JOIN brands b ON p.brand_id=b.brand_id
    """)


@st.cache_data
def load_price():
    return _query("SELECT gkey, period, price FROM price_history")


@st.cache_data
def load_channels():
    return _query("SELECT * FROM channels")


@st.cache_data
def load_channel_perf():
    return _query("SELECT * FROM channel_perf")


@st.cache_data
def load_fact():
    """网红×商品×期次 扁平事实表（含品牌、标准名）"""
    return _query("""
        SELECT cp.channel_name, cp.gkey, cp.period, cp.gmv, cp.redeem,
               cp.page_price, cp.qty, cp.band,
               p.std_name, b.brand_name
        FROM channel_product cp
        LEFT JOIN products p ON cp.gkey=p.gkey
        LEFT JOIN brands b ON p.brand_id=b.brand_id
    """)


def page_product():
    """商品档案：搜索 + 价格轨迹 + 各期表现 + 卖过它的网红"""
    prods = load_products()
    price = load_price()
    fact = load_fact()

    col1, col2 = st.columns([2, 1])
    with col1:
        kw = st.text_input('搜索商品（名称模糊匹配）', key='prod_kw')
    with col2:
        band_sel = st.multiselect('价格带筛选', sorted(prods['band'].dropna().unique()), key='prod_band')

    df = prods
    if kw.strip():
        k = kw.strip().lower()
        df = df[df['std_name'].str.lower().str.contains(k, na=False)]
    if band_sel:
        df = df[df['band'].isin(band_sel)]

    st.caption(f'共 {len(df)} 个商品')
    st.dataframe(df[['std_name', 'brand_name', 'band', 'product_ids']].rename(columns={
        'std_name': '商品名', 'brand_name': '品牌', 'band': '价格带', 'product_ids': '历次商品ID'}),
        use_container_width=True, height=300)

    if len(df) == 0:
        return

    # 选择具体商品看详情
    sel = st.selectbox('选择商品查看详情', df['std_name'].tolist(), key='prod_sel')
    if not sel:
        return

    info = df[df['std_name'] == sel].iloc[0]
    gkey = info['gkey']

    st.markdown(f"#### {sel}")
    m1, m2, m3 = st.columns(3)
    m1.metric('品牌', info['brand_name'] or '-')
    m2.metric('价格带', info['band'] or '-')
    m3.metric('历次商品ID', str(info['product_ids']).count('/') + 1 if info['product_ids'] else 1)

    # 价格轨迹
    pp = price[price['gkey'] == gkey].set_index('period').reindex(PERIODS)
    traj = pd.DataFrame({'到手价($)': pp['price']})

    if traj['到手价($)'].notna().any():
        st.line_chart(traj)
    st.dataframe(traj.T.rename(columns={'11月': '25年11月', '3月': '26年3月', '6月': '26年6月', '8月': '26年8月'}),
                 use_container_width=True)

    # 各期表现（来自事实表）
    fp = fact[fact['gkey'] == gkey]
    if len(fp):
        st.markdown('**各期销售表现**')
        agg = fp.groupby('period').agg(
            GMV=('gmv', 'sum'), 核销数=('redeem', 'sum'),
            渠道数=('channel_name', 'nunique')).reindex(PERIODS)
        st.dataframe(agg.rename(columns={}), use_container_width=True)

        st.markdown('**卖过这个品的网红（按GMV）**')
        ch = fp.groupby('channel_name')['gmv'].sum().sort_values(ascending=False)
        st.dataframe(pd.DataFrame({'GMV($)': ch}).reset_index().rename(
            columns={'channel_name': '网红', 'GMV($)': 'GMV($)'}), use_container_width=True)

def page_brand():
    """品牌档案：品牌排行 + 品牌下商品 + 品牌×网红"""
    brands = load_brands()
    prods = load_products()
    fact = load_fact()

    st.caption(f'共 {len(brands)} 个品牌（含无品牌白牌商品）')

    # 品牌排行
    top_n = st.slider('显示 Top N 品牌', 5, 50, 20, key='brand_topn')
    top = brands[brands['brand_name'] != '未识别'].head(top_n)
    st.dataframe(top[['brand_name', 'product_count', 'total_gmv']].rename(columns={
        'brand_name': '品牌', 'product_count': '商品数', 'total_gmv': '累计GMV($)'}),
        use_container_width=True)

    st.bar_chart(top.set_index('brand_name')['total_gmv'].head(top_n))

    # 选择品牌看详情
    brand_options = brands['brand_name'].tolist()
    sel = st.selectbox('选择品牌查看详情', brand_options, key='brand_sel')
    if not sel:
        return

    st.markdown(f"#### 品牌：{sel}")
    bprods = prods[prods['brand_name'] == sel]
    bf = fact[fact['brand_name'] == sel]

    m1, m2, m3, m4 = st.columns(4)
    m1.metric('商品数', len(bprods))
    m2.metric('累计GMV', _fmt_money(bf['gmv'].sum()))
    m3.metric('总核销数', _fmt_num(bf['redeem'].sum()))
    m4.metric('覆盖渠道', bf['channel_name'].nunique())

    # 品牌各期表现
    per = bf.groupby('period')['gmv'].sum().reindex(PERIODS)
    st.markdown('**各期GMV**')
    st.bar_chart(per.fillna(0))

    # 品牌×网红（该品牌被哪些网红卖）
    st.markdown('**卖过该品牌的网红（按GMV）**')
    ch = bf.groupby('channel_name')['gmv'].sum().sort_values(ascending=False).head(30)
    st.dataframe(pd.DataFrame({'GMV($)': ch}).reset_index().rename(
        columns={'channel_name': '网红'}), use_container_width=True)

    st.markdown('**该品牌商品清单**')
    st.dataframe(bprods[['std_name', 'band', 'product_ids']].rename(columns={
        'std_name': '商品名', 'band': '价格带', 'product_ids': '历次商品ID'}),
        use_container_width=True)

def page_channel():
    """网红档案：网红排行 + 四期表现 + 卖过的品"""
    chans = load_channels()
    perf = load_channel_perf()
    fact = load_fact()

    st.caption(f'共 {len(chans)} 个网红渠道（含 8月活跃 + 历史参与）')

    # 筛选：参与期数
    col1, col2 = st.columns([1, 2])
    with col1:
        min_periods = st.slider('最少参与期数', 1, 4, 1, key='ch_minp')
    with col2:
        only_aug = st.checkbox('仅看 8月活跃', key='ch_onlyaug')

    df = chans[chans['periods_n'] >= min_periods]
    if only_aug:
        aug_chs = set(perf[perf['period'] == '8月']['channel_name'])
        df = df[df['channel_name'].isin(aug_chs)]

    # 合并历史合计排序
    df = df.sort_values('hist_total', ascending=False)
    st.dataframe(df[['channel_name', 'periods_n', 'timeline', 'hist_total']].rename(columns={
        'channel_name': '网红', 'periods_n': '参与期数', 'timeline': '参与时间线',
        'hist_total': '历史合计GMV($)'}), use_container_width=True, height=300)

    if len(df) == 0:
        return

    sel = st.selectbox('选择网红查看详情', df['channel_name'].tolist(), key='ch_sel')
    if not sel:
        return

    st.markdown(f"#### 网红：{sel}")
    cp = perf[perf['channel_name'] == sel]
    cf = fact[fact['channel_name'] == sel]

    m1, m2, m3, m4 = st.columns(4)
    m1.metric('累计GMV', _fmt_money(cp['gmv'].sum()))
    m2.metric('参与期数', chans[chans['channel_name'] == sel]['periods_n'].iloc[0])
    aug_perf = cp[cp['period'] == '8月']
    m3.metric('8月核销率', _fmt_pct(aug_perf['redeem_rate'].iloc[0]) if len(aug_perf) and aug_perf['redeem_rate'].notna().any() else '-')
    m4.metric('8月播放量', _fmt_num(aug_perf['views'].iloc[0]) if len(aug_perf) and aug_perf['views'].notna().any() else '-')

    # 四期GMV趋势
    per = cp.set_index('period')['gmv'].reindex(PERIODS)
    st.markdown('**各期GMV趋势**')
    st.bar_chart(per.fillna(0))

    # 卖过的品
    st.markdown('**卖过的商品（按GMV）**')
    prods = cf.groupby(['gkey', 'std_name'])['gmv'].sum().sort_values(ascending=False).reset_index()
    st.dataframe(prods.rename(columns={'std_name': '商品名', 'gmv': 'GMV($)'}),
                 use_container_width=True)

def page_explorer():
    """维度交叉查询：任意 2 维或 3 维组合
    维度：品牌 / 商品 / 网红 / 期次 / 价格带
    """
    fact = load_fact()

    st.markdown("""
    选择 **2 个或 3 个维度**作为交叉分析轴，系统会自动聚合出对应结果。
    例如：品牌×期次、网红×品牌×期次、商品×价格带 等。
    """)

    DIMS = ['品牌', '商品', '网红', '期次', '价格带']
    DIM_COL = {'品牌': 'brand_name', '商品': 'std_name', '网红': 'channel_name',
               '期次': 'period', '价格带': 'band'}

    ndim = st.radio('交叉维度数', [2, 3], horizontal=True, key='exp_ndim',
                    format_func=lambda x: f'{x} 维交叉')
    selected = st.multiselect('选择维度（按顺序：行→列→筛选）', DIMS, default=['品牌', '期次'],
                              max_selections=ndim, key='exp_dims')

    if len(selected) < 2:
        st.info('请至少选择 2 个维度')
        return

    # 期次排序辅助
    def _sort_period(df, col):
        if col in df.columns:
            order = {p: i for i, p in enumerate(PERIODS)}
            df['_po'] = df[col].map(order)
            df = df.sort_values('_po').drop(columns='_po')
        return df

    agg_metrics = st.multiselect('聚合指标', ['GMV', '核销数', '计划数量', '商品数', '渠道数', '平均页面价'],
                                 default=['GMV'], key='exp_metrics')
    if not agg_metrics:
        agg_metrics = ['GMV']

    METRIC_FN = {
        'GMV': ('gmv', 'sum'), '核销数': ('redeem', 'sum'),
        '计划数量': ('qty', 'sum'), '商品数': ('gkey', 'nunique'),
        '渠道数': ('channel_name', 'nunique'), '平均页面价': ('page_price', 'mean'),
    }
    agg_dict = {METRIC_FN[m][0]: METRIC_FN[m][1] for m in agg_metrics}

    group_cols = [DIM_COL[d] for d in selected]
    res = fact.groupby(group_cols).agg(agg_dict).reset_index()

    # 重命名列
    rename = {DIM_COL[d]: d for d in selected}
    metric_names = list(agg_dict.keys())
    for i, m in enumerate(agg_metrics):
        if i < len(metric_names):
            rename[metric_names[i]] = m
    res = res.rename(columns=rename)

    # 透视：第1维=行，第2维=列，第3维=筛选
    row_dim, col_dim = selected[0], selected[1]

    if len(selected) == 3:
        filt_dim = selected[2]
        st.markdown(f'**筛选维度：{filt_dim}**')
        filt_vals = sorted(res[filt_dim].dropna().unique().tolist())
        sel_filt = st.multiselect(f'选择 {filt_dim}', filt_vals, key='exp_filt')
        if sel_filt:
            res = res[res[filt_dim].isin(sel_filt)]
            # 重新聚合掉筛选维
            res = res.groupby([DIM_COL[row_dim], DIM_COL[col_dim]]).agg(
                {METRIC_FN[m][0]: METRIC_FN[m][1] for m in agg_metrics}).reset_index()
            res = res.rename(columns={DIM_COL[row_dim]: row_dim, DIM_COL[col_dim]: col_dim,
                                      **{METRIC_FN[m][0]: m for i, m in enumerate(agg_metrics)}})

    if len(res) == 0:
        st.warning('该组合下无数据')
        return

    # 主指标透视表
    primary = agg_metrics[0]
    pivot = res.pivot_table(index=row_dim, columns=col_dim, values=primary,
                            aggfunc='sum').fillna(0)
    # 期次列排序
    period_cols = [p for p in PERIODS if p in pivot.columns]
    other_cols = [c for c in pivot.columns if c not in PERIODS]
    pivot = pivot[period_cols + other_cols]

    st.markdown(f'**{primary} · {row_dim} × {col_dim}**')
    st.dataframe(pivot.style.format(lambda v: f'{v:,.0f}'), use_container_width=True)

    # 图表
    if pivot.shape[1] <= 12:
        chart_df = pivot
        if st.checkbox('显示图表', value=True, key='exp_chart'):
            st.bar_chart(chart_df)

    # 明细表
    with st.expander('查看完整明细'):
        st.dataframe(_sort_period(res, DIM_COL[col_dim]) if col_dim == '期次' else res,
                     use_container_width=True)

    # 下载
    csv = res.to_csv(index=False).encode('utf-8-sig')
    st.download_button('下载交叉结果 CSV', csv, file_name='cross_query.csv',
                       mime='text/csv', key='exp_dl')

def page_notes():
    """口径与数据说明"""
    st.markdown("""
### 数据来源
4 期大促底表 + 「网红团购核销数据汇总」表：

| 期次 | 口径 | 说明 |
|---|---|---|
| 25年11月 | 底表 BU列「gmv」(实际成交) | 覆盖率96%，与页面价×领取张数交叉验证吻合89% |
| 26年3月 | 页面价(Q列) × 领取张数/核销数量(CD列) | 覆盖率92% |
| 26年6月 | 页面价(T列) × 核销张数(CF列) | 覆盖率97%，含3行韩币已按汇率折算 |
| 26年8月 | 核销数据表实际核销值 | 覆盖率100%，计划数量按活动ID关联 |

### 数据对账（入库校验）
- 11月：$9,424,564（含4行退款负值，如实保留）
- 3月：$7,279,961 ｜ 6月：$5,813,708 ｜ 8月：$6,344,204
- 与渠道汇总表逐分对平，无遗漏

### 维度与关联
- **标准商品**：以品牌+型号为统一主键，商品ID每次大促会变（已归并）
- **品牌**：从商品名提取，识别率约74%，无品牌白牌商品归「未识别」
- **价格带**：A<$20引流 / B$20-50 / C$50-100 / D$100-200 / E≥$200高客单
- **网红**：含8月活跃 + 历史参与，共159渠道

### 已知数据质量标记
- 11月有4行退款负值（톡써니테크/슈퍼테슬라/알리박스/겜용이），保留不剔除
- 더더마 11月PID与8月差1位，待确认
- 랑스형님 与 랑스알리 名称相似，未合并
- 核销率>100% 为追加预算后核销超预期，保留标注

### 样本量提醒
单网红×单品多数只有1-2期样本，单点结论仅供参考；跨品聚合的价格带/品牌结论更可靠。
""")


def render():
    st.markdown('<div class="main-header">📊 历史大促数据中枢</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="sub-header">4 期大促（11月 / 3月 / 6月 / 8月）· 品牌 / 商品 / 网红 / 期次 任意维度交叉查询 · 数据依据</div>',
        unsafe_allow_html=True)

    if DB_PATH is None:
        st.error('未找到 history_data.db 数据库文件。请将其放在与 app.py 相同的目录下。')
        return

    tab_p, tab_b, tab_c, tab_x, tab_s = st.tabs(
        ['🔍 商品档案', '🏷️ 品牌档案', '👤 网红档案', '🧩 维度交叉查询', '📖 口径与数据说明'])

    with tab_p:
        page_product()
    with tab_b:
        page_brand()
    with tab_c:
        page_channel()
    with tab_x:
        page_explorer()
    with tab_s:
        page_notes()
