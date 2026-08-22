"""
历史数据中枢 (history_hub) · 小白友好版
─────────────────────────────────────────────
网红团购 · 历史大促数据中枢（数据依据）
轻盈毛玻璃风 · 图表优先 · 大白话注释 · 一目了然
"""

import sqlite3
from pathlib import Path

import streamlit as st
import pandas as pd

PERIODS = ['11月', '3月', '6月', '8月']
PERIOD_LABEL = {'11月': '25年11月', '3月': '26年3月', '6月': '26年6月', '8月': '26年8月'}

_HERE = Path(__file__).parent
_DB_CANDIDATES = [_HERE / 'history_data.db', Path('history_data.db')]
DB_PATH = next((str(p) for p in _DB_CANDIDATES if p.exists()), None)


# ─────────────────────────────────────────────
# 毛玻璃风样式
# ─────────────────────────────────────────────
def inject_css():
    st.markdown("""
<style>
    .hub-app { margin-top:-1rem; }
    .stApp { background: linear-gradient(160deg,#fdf2f8 0%,#f3e8ff 45%,#eef2ff 100%); }
    footer {visibility:hidden;} #MainMenu {visibility:hidden;}

    h1,h2,h3,h4{font-family:ui-rounded,'PingFang SC','Microsoft YaHei',sans-serif;color:#1f2430;font-weight:800;}

    /* 页头 */
    .hub-hero{padding:18px 26px;margin:4px 0 18px;border-radius:22px;
      background:linear-gradient(135deg,rgba(244,114,182,.85),rgba(167,139,250,.85));
      color:#fff;box-shadow:0 10px 30px rgba(167,139,250,.25);}
    .hub-hero .t{font-size:26px;font-weight:800;letter-spacing:.5px;}
    .hub-hero .s{font-size:13px;opacity:.95;margin-top:4px;font-weight:600;}

    /* 毛玻璃卡片 */
    .glass{background:rgba(255,255,255,.62);backdrop-filter:blur(14px);
      border:1px solid rgba(255,255,255,.85);border-radius:18px;
      box-shadow:0 8px 24px rgba(120,90,200,.08);padding:18px 20px;margin-bottom:14px;}

    /* KPI 卡 */
    .kpi{background:rgba(255,255,255,.7);backdrop-filter:blur(14px);
      border:1px solid rgba(255,255,255,.9);border-radius:16px;padding:16px 18px;
      box-shadow:0 6px 18px rgba(120,90,200,.08);height:100%;}
    .kpi .k{font-size:12px;color:#8b8fa3;font-weight:700;margin-bottom:6px;}
    .kpi .v{font-size:26px;font-weight:800;color:#1f2430;line-height:1;}
    .kpi .d{font-size:12px;margin-top:8px;font-weight:600;}
    .kpi .hint{font-size:11px;color:#a5a8ba;margin-top:6px;line-height:1.4;}
    .up{color:#16a34a;} .down{color:#dc2626;}

    /* 章节标题 */
    .sec{font-size:16px;font-weight:800;color:#1f2430;margin:20px 0 8px;
      padding-left:12px;border-left:4px solid #a78bfa;}
    .sec-sub{font-size:12px;color:#8b8fa3;margin:0 0 10px 16px;font-weight:600;}

    /* 场景快捷卡片 */
    .scn{background:rgba(255,255,255,.66);backdrop-filter:blur(12px);
      border:1px solid rgba(255,255,255,.9);border-radius:14px;padding:14px 16px;
      cursor:pointer;transition:.18s;height:100%;}
    .scn:hover{transform:translateY(-3px);box-shadow:0 10px 24px rgba(167,139,250,.22);}
    .scn .e{font-size:22px;} .scn .t{font-size:14px;font-weight:800;color:#1f2430;margin-top:6px;}
    .scn .d{font-size:11px;color:#8b8fa3;margin-top:4px;line-height:1.4;}

    /* 大白话注释 */
    .explain{font-size:12px;color:#8b8fa3;line-height:1.6;background:rgba(244,114,182,.06);
      border-left:3px solid #f9a8d4;padding:8px 12px;border-radius:0 10px 10px 0;margin:6px 0 10px;}

    /* 涨跌徽章 */
    .badge{display:inline-block;padding:2px 10px;border-radius:999px;font-size:11px;font-weight:700;}
    .b-up{background:#dcfce7;color:#16a34a;} .b-down{background:#fee2e2;color:#dc2626;}
    .b-flat{background:#eef2ff;color:#6366f1;} .b-info{background:#ede9fe;color:#7c3aed;}

    /* Streamlit 组件微调 */
    div[data-testid="stSidebar"]{background:rgba(255,255,255,.5);backdrop-filter:blur(10px);}
    .stTabs [data-baseweb="tab-list"]{gap:6px;background:rgba(255,255,255,.5);
      padding:6px;border-radius:14px;}
    .stTabs [data-baseweb="tab"]{border-radius:10px;padding:8px 16px;font-weight:700;}
    div[data-testid="stMetric"]{background:rgba(255,255,255,.6);border-radius:14px;padding:12px;}
</style>
""", unsafe_allow_html=True)


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
    if v is None or pd.isna(v): return '-'
    return f'${v:,.0f}'

def _fmt_pct(v):
    if v is None or pd.isna(v): return '-'
    return f'{v:.1f}%'

def _fmt_num(v):
    if v is None or pd.isna(v): return '-'
    return f'{v:,.0f}'


@st.cache_data
def load_brands():
    return _query("SELECT * FROM brands ORDER BY total_gmv DESC")

@st.cache_data
def load_products():
    return _query("""
        SELECT p.gkey,p.std_name,p.band,p.product_ids,b.brand_name
        FROM products p LEFT JOIN brands b ON p.brand_id=b.brand_id
    """)

@st.cache_data
def load_price():
    return _query("SELECT gkey,period,price FROM price_history")

@st.cache_data
def load_channels():
    return _query("SELECT * FROM channels")

@st.cache_data
def load_channel_perf():
    return _query("SELECT * FROM channel_perf")

@st.cache_data
def load_fact():
    return _query("""
        SELECT cp.channel_name,cp.gkey,cp.period,cp.gmv,cp.redeem,
               cp.page_price,cp.qty,cp.band,p.std_name,b.brand_name
        FROM channel_product cp
        LEFT JOIN products p ON cp.gkey=p.gkey
        LEFT JOIN brands b ON p.brand_id=b.brand_id
    """)


def page_dashboard():
    """总览仪表盘：一眼看懂全局 + 场景快捷入口"""
    prods = load_products()
    fact = load_fact()
    chans = load_channels()
    price = load_price()

    # ── 顶部 KPI 卡 ──
    total_gmv = fact['gmv'].sum()
    n_products = prods['gkey'].nunique()
    n_channels = chans['channel_name'].nunique()
    # 4期全勤商品数
    period_cnt = price.groupby('gkey')['period'].nunique()
    full4 = int((period_cnt == 4).sum())

    st.markdown('<div class="sec">📌 先看这几个关键数字</div>', unsafe_allow_html=True)
    st.markdown('<div class="sec-sub">4 期大促加起来的总盘子</div>', unsafe_allow_html=True)

    c1, c2, c3, c4 = st.columns(4)
    c1.markdown(f"""<div class="kpi"><div class="k">4期总成交额</div>
      <div class="v">{_fmt_money(total_gmv)}</div>
      <div class="hint">11月+3月+6月+8月 四场大促加起来卖了多少钱</div></div>""", unsafe_allow_html=True)
    c2.markdown(f"""<div class="kpi"><div class="k">累计商品数</div>
      <div class="v">{n_products:,}</div>
      <div class="hint">按"品牌+型号"去重后的标准商品（同款不同ID已合并）</div></div>""", unsafe_allow_html=True)
    c3.markdown(f"""<div class="kpi"><div class="k">合作网红数</div>
      <div class="v">{n_channels}</div>
      <div class="hint">4期累计合作过的渠道/网红总数</div></div>""", unsafe_allow_html=True)
    c4.markdown(f"""<div class="kpi"><div class="k">4期全勤商品</div>
      <div class="v">{full4}</div>
      <div class="hint">每一期都参加的"常青品"，价格轨迹最完整</div></div>""", unsafe_allow_html=True)

    # ── 各期大盘趋势 ──
    st.markdown('<div class="sec">📈 各期成交走势</div>', unsafe_allow_html=True)
    st.markdown('<div class="explain">每一期大促的总成交额。看大盘是涨是跌、哪一期最旺。</div>',
                unsafe_allow_html=True)
    per_gmv = fact.groupby('period')['gmv'].sum().reindex(PERIODS)
    trend_df = pd.DataFrame({'成交额($)': per_gmv.values},
                            index=[PERIOD_LABEL[p] for p in PERIODS])
    st.bar_chart(trend_df)

    # 环比
    if len(per_gmv) >= 2 and per_gmv.iloc[-2] and per_gmv.iloc[-2] > 0:
        chg = (per_gmv.iloc[-1] - per_gmv.iloc[-2]) / per_gmv.iloc[-2] * 100
        cls = 'up' if chg >= 0 else 'down'
        arrow = '↑' if chg >= 0 else '↓'
        st.markdown(f'<span class="badge b-{"up" if chg>=0 else "down"}">8月比6月 {arrow} {abs(chg):.1f}%</span>',
                    unsafe_allow_html=True)

    # ── 场景快捷入口 ──
    st.markdown('<div class="sec">🎯 你想看什么？点下面直达</div>', unsafe_allow_html=True)
    st.markdown('<div class="sec-sub">按常见决策场景一键跳转，不用自己想怎么查</div>', unsafe_allow_html=True)

    sc1, sc2, sc3, sc4 = st.columns(4)
    sc1.markdown("""<div class="scn"><div class="e">💰</div><div class="t">哪些品降价了？</div>
      <div class="d">看商品档案，对比各期到手价，红涨绿跌一目了然</div></div>""", unsafe_allow_html=True)
    sc2.markdown("""<div class="scn"><div class="e">🌟</div><div class="t">哪个网红最能卖？</div>
      <div class="d">看网红档案，按成交额和核销率排座次</div></div>""", unsafe_allow_html=True)
    sc3.markdown("""<div class="scn"><div class="e">🏷️</div><div class="t">哪个品牌是主力？</div>
      <div class="d">看品牌档案，品牌成交额排行+旗下商品</div></div>""", unsafe_allow_html=True)
    sc4.markdown("""<div class="scn"><div class="e">🧩</div><div class="t">自由交叉查数</div>
      <div class="d">任意2-3个维度组合，透视出你想看的角度</div></div>""", unsafe_allow_html=True)
    st.caption('💡 提示：切换到对应页签即可深入查看。每个指标旁边都有大白话解释。')

    # ── Top 速览 ──
    st.markdown('<div class="sec">🏆 榜单速览</div>', unsafe_allow_html=True)
    lc, rp = st.columns(2)
    with lc:
        st.markdown('**成交额 Top5 网红**')
        top_ch = fact.groupby('channel_name')['gmv'].sum().sort_values(ascending=False).head(5)
        st.dataframe(pd.DataFrame({'成交额($)': top_ch}).reset_index().rename(
            columns={'channel_name': '网红'}), use_container_width=True, hide_index=True)
    with rp:
        st.markdown('**成交额 Top5 商品**')
        top_p = fact.groupby('std_name')['gmv'].sum().sort_values(ascending=False).head(5)
        st.dataframe(pd.DataFrame({'成交额($)': top_p}).reset_index().rename(
            columns={'std_name': '商品'}), use_container_width=True, hide_index=True)

def page_product():
    """商品档案（小白友好版）：搜索 + 价格轨迹图 + 红涨绿跌"""
    prods = load_products()
    price = load_price()
    fact = load_fact()

    st.markdown('<div class="sec">🔍 查商品</div>', unsafe_allow_html=True)
    st.markdown('<div class="explain">想看哪个品的历史价格，就搜哪个。'
                '同一个商品在不同大促可能用不同的商品ID，这里已经自动合并成一个，不会重复也不会漏。</div>',
                unsafe_allow_html=True)

    col1, col2 = st.columns([2, 1])
    with col1:
        kw = st.text_input('输入商品关键词（如 edifier / y700 / 键盘）', key='prod_kw')
    with col2:
        band_opts = sorted([b for b in prods['band'].dropna().unique()])
        band_sel = st.multiselect('按价格带筛选（可不选）', band_opts, key='prod_band')

    df = prods
    if kw.strip():
        k = kw.strip().lower()
        df = df[df['std_name'].str.lower().str.contains(k, na=False)]
    if band_sel:
        df = df[df['band'].isin(band_sel)]

    if len(df) == 0:
        st.info('没有匹配的商品，换个关键词试试～')
        return

    st.markdown(f'找到 **{len(df)}** 个商品，下面选一个看详情：')
    sel = st.selectbox('选择商品', df['std_name'].tolist(), key='prod_sel')
    if not sel:
        return

    info = df[df['std_name'] == sel].iloc[0]
    gkey = info['gkey']

    # ── 商品头部信息 ──
    h1, h2, h3 = st.columns(3)
    h1.markdown(f"""<div class="kpi"><div class="k">品牌</div>
      <div class="v" style="font-size:18px">{info['brand_name'] or '-'}</div></div>""", unsafe_allow_html=True)
    h2.markdown(f"""<div class="kpi"><div class="k">价格带</div>
      <div class="v" style="font-size:18px">{info['band'] or '-'}</div></div>""", unsafe_allow_html=True)
    n_ids = str(info['product_ids']).count('/') + 1 if info['product_ids'] else 1
    h3.markdown(f"""<div class="kpi"><div class="k">用过的商品ID数</div>
      <div class="v" style="font-size:18px">{n_ids}</div>
      <div class="hint">换过ID也没关系，都已合并</div></div>""", unsafe_allow_html=True)

    # ── 价格轨迹 ──
    st.markdown('<div class="sec">💹 到手价走势</div>', unsafe_allow_html=True)
    st.markdown('<div class="explain">到手价 = 顾客实际付的钱。'
                '这条线越高说明卖得越贵，往下走说明在降价。</div>', unsafe_allow_html=True)

    pp = price[price['gkey'] == gkey].set_index('period').reindex(PERIODS)
    traj = pd.DataFrame({'到手价($)': pp['price']})

    if traj['到手价($)'].notna().any():
        st.line_chart(traj)

        # 红涨绿跌对比
        vals = traj['到手价($)'].dropna()
        if len(vals) >= 2:
            first, last = vals.iloc[0], vals.iloc[-1]
            chg = (last - first) / first * 100
            if chg > 0.5:
                badge, cls = f'↑ 涨价 {chg:.1f}%', 'b-up'
            elif chg < -0.5:
                badge, cls = f'↓ 降价 {abs(chg):.1f}%', 'b-down'
            else:
                badge, cls = '→ 价格持平', 'b-flat'
            st.markdown(f'从首期到最近一期：<span class="badge {cls}">{badge}</span>　'
                        f'历史最低 <b>${vals.min():.2f}</b>　历史最高 <b>${vals.max():.2f}</b>',
                        unsafe_allow_html=True)

    # 各期明细
    with st.expander('📋 各期到手价明细'):
        t = traj.T
        t.columns = [PERIOD_LABEL[p] for p in PERIODS]
        st.dataframe(t.rename(index={'到手价($)': '到手价($)'}), use_container_width=True)

    # ── 销售表现 ──
    fp = fact[fact['gkey'] == gkey]
    if len(fp):
        st.markdown('<div class="sec">📦 每期卖了多少</div>', unsafe_allow_html=True)
        st.markdown('<div class="explain">同一个品每期有多少网红在卖、总共卖了多少钱。</div>',
                    unsafe_allow_html=True)
        agg = fp.groupby('period').agg(
            成交额=('gmv', 'sum'), 核销数=('redeem', 'sum'),
            在卖网红数=('channel_name', 'nunique')).reindex(PERIODS)
        agg.index = [PERIOD_LABEL[p] for p in agg.index]
        st.dataframe(agg, use_container_width=True)

        st.markdown('<div class="sec">👥 谁在卖这个品</div>', unsafe_allow_html=True)
        st.markdown('<div class="explain">按成交额从高到低排，越靠前说明这个网红卖这个品越厉害。</div>',
                    unsafe_allow_html=True)
        ch = fp.groupby('channel_name')['gmv'].sum().sort_values(ascending=False)
        st.dataframe(pd.DataFrame({'成交额($)': ch}).reset_index().rename(
            columns={'channel_name': '网红'}), use_container_width=True, hide_index=True)

def page_brand():
    """品牌档案（小白友好版）"""
    brands = load_brands()
    prods = load_products()
    fact = load_fact()

    st.markdown('<div class="sec">🏷️ 品牌大盘</div>', unsafe_allow_html=True)
    st.markdown('<div class="explain">哪个品牌卖得最好、旗下有哪些品、被哪些网红带过货。'
                '"未识别"是没有明确品牌的白牌商品（多为户外用品）。</div>', unsafe_allow_html=True)

    top_n = st.slider('显示 Top N 品牌', 5, 50, 15, key='brand_topn')
    top = brands[brands['brand_name'] != '未识别'].head(top_n)

    st.bar_chart(top.set_index('brand_name')['total_gmv'])

    with st.expander('📋 品牌排行明细'):
        st.dataframe(top[['brand_name', 'product_count', 'total_gmv']].rename(columns={
            'brand_name': '品牌', 'product_count': '商品数', 'total_gmv': '累计成交额($)'}),
            use_container_width=True, hide_index=True)

    # 品牌详情
    st.markdown('<div class="sec">🔎 看某个品牌的详情</div>', unsafe_allow_html=True)
    brand_options = brands['brand_name'].tolist()
    sel = st.selectbox('选择品牌', brand_options, key='brand_sel')
    if not sel:
        return

    bprods = prods[prods['brand_name'] == sel]
    bf = fact[fact['brand_name'] == sel]

    m1, m2, m3, m4 = st.columns(4)
    m1.markdown(f"""<div class="kpi"><div class="k">旗下商品数</div>
      <div class="v" style="font-size:20px">{len(bprods)}</div></div>""", unsafe_allow_html=True)
    m2.markdown(f"""<div class="kpi"><div class="k">累计成交额</div>
      <div class="v" style="font-size:20px">{_fmt_money(bf['gmv'].sum())}</div></div>""", unsafe_allow_html=True)
    m3.markdown(f"""<div class="kpi"><div class="k">总核销数</div>
      <div class="v" style="font-size:20px">{_fmt_num(bf['redeem'].sum())}</div>
      <div class="hint">卖出去多少件</div></div>""", unsafe_allow_html=True)
    m4.markdown(f"""<div class="kpi"><div class="k">合作网红数</div>
      <div class="v" style="font-size:20px">{bf['channel_name'].nunique()}</div></div>""", unsafe_allow_html=True)

    per = bf.groupby('period')['gmv'].sum().reindex(PERIODS)
    st.markdown('**各期成交额**')
    st.bar_chart(pd.DataFrame({'成交额($)': per.fillna(0).values},
                              index=[PERIOD_LABEL[p] for p in PERIODS]))

    st.markdown('<div class="sec">👥 谁在卖这个品牌</div>', unsafe_allow_html=True)
    ch = bf.groupby('channel_name')['gmv'].sum().sort_values(ascending=False).head(30)
    st.dataframe(pd.DataFrame({'成交额($)': ch}).reset_index().rename(
        columns={'channel_name': '网红'}), use_container_width=True, hide_index=True)

    with st.expander('📋 该品牌全部商品清单'):
        st.dataframe(bprods[['std_name', 'band', 'product_ids']].rename(columns={
            'std_name': '商品名', 'band': '价格带', 'product_ids': '历次商品ID'}),
            use_container_width=True, hide_index=True)

def page_channel():
    """网红档案（小白友好版）"""
    chans = load_channels()
    perf = load_channel_perf()
    fact = load_fact()

    st.markdown('<div class="sec">👤 网红大盘</div>', unsafe_allow_html=True)
    st.markdown('<div class="explain">哪个网红最能卖、参加了哪几期、卖过哪些品。'
                '"参与期数"越多说明合作越稳定。</div>', unsafe_allow_html=True)

    col1, col2 = st.columns([1, 2])
    with col1:
        min_periods = st.slider('最少参与期数', 1, 4, 1, key='ch_minp')
    with col2:
        only_aug = st.checkbox('只看 8月还在合作的', key='ch_onlyaug')

    df = chans[chans['periods_n'] >= min_periods]
    if only_aug:
        aug_chs = set(perf[perf['period'] == '8月']['channel_name'])
        df = df[df['channel_name'].isin(aug_chs)]

    df = df.sort_values('hist_total', ascending=False)

    if len(df) == 0:
        st.info('没有符合条件的网红')
        return

    st.markdown(f'共 **{len(df)}** 个网红，按历史总成交额排序，选一个看详情：')
    sel = st.selectbox('选择网红', df['channel_name'].tolist(), key='ch_sel')
    if not sel:
        return

    cp = perf[perf['channel_name'] == sel]
    cf = fact[fact['channel_name'] == sel]
    aug_perf = cp[cp['period'] == '8月']

    m1, m2, m3, m4 = st.columns(4)
    m1.markdown(f"""<div class="kpi"><div class="k">累计成交额</div>
      <div class="v" style="font-size:20px">{_fmt_money(cp['gmv'].sum())}</div></div>""", unsafe_allow_html=True)
    pn = chans[chans['channel_name'] == sel]['periods_n'].iloc[0]
    m2.markdown(f"""<div class="kpi"><div class="k">参与期数</div>
      <div class="v" style="font-size:20px">{pn}/4</div>
      <div class="hint">{pn>=3 and '老搭档，合作稳定' or '合作期数较少'}</div></div>""", unsafe_allow_html=True)
    rr = aug_perf['redeem_rate'].iloc[0] if len(aug_perf) and aug_perf['redeem_rate'].notna().any() else None
    m3.markdown(f"""<div class="kpi"><div class="k">8月核销率</div>
      <div class="v" style="font-size:20px">{_fmt_pct(rr)}</div>
      <div class="hint">优惠券被真正用掉的比例，越高说明带货越实</div></div>""", unsafe_allow_html=True)
    vv = aug_perf['views'].iloc[0] if len(aug_perf) and aug_perf['views'].notna().any() else None
    m4.markdown(f"""<div class="kpi"><div class="k">8月播放量</div>
      <div class="v" style="font-size:20px">{_fmt_num(vv)}</div>
      <div class="hint">视频被看了多少次</div></div>""", unsafe_allow_html=True)

    # 各期趋势
    st.markdown('<div class="sec">📈 各期成交额走势</div>', unsafe_allow_html=True)
    per = cp.set_index('period')['gmv'].reindex(PERIODS)
    st.bar_chart(pd.DataFrame({'成交额($)': per.fillna(0).values},
                              index=[PERIOD_LABEL[p] for p in PERIODS]))

    # 卖过的品
    st.markdown('<div class="sec">📦 卖过的商品</div>', unsafe_allow_html=True)
    st.markdown('<div class="explain">这个网红带过哪些货、各卖了多少钱，按成交额排序。</div>',
                unsafe_allow_html=True)
    prods = cf.groupby(['gkey', 'std_name'])['gmv'].sum().sort_values(ascending=False).reset_index()
    st.dataframe(prods[['std_name', 'gmv']].rename(columns={
        'std_name': '商品', 'gmv': '成交额($)'}), use_container_width=True, hide_index=True)

def page_explorer():
    """维度交叉查询（小白友好版）"""
    fact = load_fact()

    st.markdown('<div class="sec">🧩 自由交叉查数</div>', unsafe_allow_html=True)
    st.markdown('<div class="explain">'
                '想从什么角度看数据，就选什么维度。比如选"品牌+期次"，就能看每个品牌每期卖多少；'
                '选"网红+品牌+期次"，就能看某个网红卖某个品牌每期卖多少。'
                '最多选 3 个维度：第 1 个当行、第 2 个当列、第 3 个当筛选条件。</div>',
                unsafe_allow_html=True)

    DIMS = ['品牌', '商品', '网红', '期次', '价格带']
    DIM_COL = {'品牌': 'brand_name', '商品': 'std_name', '网红': 'channel_name',
               '期次': 'period', '价格带': 'band'}

    ndim = st.radio('想看几个维度的交叉？', [2, 3], horizontal=True, key='exp_ndim',
                    format_func=lambda x: f'{x} 个维度')
    selected = st.multiselect(f'选择 {ndim} 个维度（按顺序：第1个=行，第2个=列，第3个=筛选）',
                              DIMS, default=['品牌', '期次'],
                              max_selections=ndim, key='exp_dims')

    if len(selected) < 2:
        st.info('请至少选择 2 个维度')
        return

    agg_metrics = st.multiselect('想看什么指标？', ['成交额', '核销数', '商品数', '网红数', '平均页面价'],
                                 default=['成交额'], key='exp_metrics')
    if not agg_metrics:
        agg_metrics = ['成交额']

    METRIC_FN = {
        '成交额': ('gmv', 'sum'), '核销数': ('redeem', 'sum'),
        '商品数': ('gkey', 'nunique'), '网红数': ('channel_name', 'nunique'),
        '平均页面价': ('page_price', 'mean'),
    }

    def _run():
        agg_dict = {METRIC_FN[m][0]: METRIC_FN[m][1] for m in agg_metrics}
        group_cols = [DIM_COL[d] for d in selected]
        res = fact.groupby(group_cols).agg(agg_dict).reset_index()
        rename = {DIM_COL[d]: d for d in selected}
        metric_cols = list(agg_dict.keys())
        for i, m in enumerate(agg_metrics):
            if i < len(metric_cols):
                rename[metric_cols[i]] = m
        return res.rename(columns=rename)

    res = _run()
    row_dim, col_dim = selected[0], selected[1]

    if len(selected) == 3:
        filt_dim = selected[2]
        filt_vals = sorted(res[filt_dim].dropna().unique().tolist())
        sel_filt = st.multiselect(f'🔻 第3个维度【{filt_dim}】：选要看的值（不选=全部）',
                                  filt_vals, key='exp_filt')
        if sel_filt:
            res = res[res[filt_dim].isin(sel_filt)]
            agg_dict = {METRIC_FN[m][0]: METRIC_FN[m][1] for m in agg_metrics}
            res = res.groupby([DIM_COL[row_dim], DIM_COL[col_dim]]).agg(agg_dict).reset_index()
            rename = {DIM_COL[row_dim]: row_dim, DIM_COL[col_dim]: col_dim}
            metric_cols = list(agg_dict.keys())
            for i, m in enumerate(agg_metrics):
                if i < len(metric_cols):
                    rename[metric_cols[i]] = m
            res = res.rename(columns=rename)

    if len(res) == 0:
        st.warning('这个组合下没有数据，换个维度试试～')
        return

    # 主指标透视
    primary = agg_metrics[0]
    pivot = res.pivot_table(index=row_dim, columns=col_dim, values=primary,
                            aggfunc='sum').fillna(0)
    period_cols = [p for p in PERIODS if p in pivot.columns]
    other_cols = [c for c in pivot.columns if c not in PERIODS]
    pivot = pivot[period_cols + other_cols]

    st.markdown(f'<div class="sec">📊 {primary} · {row_dim} × {col_dim}</div>', unsafe_allow_html=True)
    st.markdown(f'<div class="explain">行是【{row_dim}】，列是【{col_dim}】，'
                f'格子里的数字是{primary}。</div>', unsafe_allow_html=True)

    st.dataframe(pivot.style.format(lambda v: f'{v:,.0f}'), use_container_width=True)

    if pivot.shape[0] <= 25 and pivot.shape[1] <= 12:
        if st.checkbox('📈 显示图表', value=True, key='exp_chart'):
            st.bar_chart(pivot)

    with st.expander('📋 查看完整明细（可下载）'):
        st.dataframe(res, use_container_width=True, hide_index=True)
        csv = res.to_csv(index=False).encode('utf-8-sig')
        st.download_button('⬇️ 下载交叉结果 CSV', csv, file_name='cross_query.csv',
                           mime='text/csv', key='exp_dl')

def page_notes():
    """口径与数据说明（大白话版）"""
    st.markdown('<div class="sec">📖 这些数据是怎么来的？</div>', unsafe_allow_html=True)
    st.markdown('<div class="explain">这里是"数据说明书"，告诉你每个数字是怎么算出来的、'
                '哪些地方要注意。看懂了这个，你就知道每个结论靠不靠谱。</div>', unsafe_allow_html=True)

    st.markdown('### 🗂️ 数据覆盖范围')
    st.dataframe(pd.DataFrame({
        '期次': ['25年11月', '26年3月', '26年6月', '26年8月'],
        '成交额口径': ['实际成交（底表gmv列）', '页面价×核销数量', '页面价×核销数量', '实际核销值'],
        '数据完整度': ['96%', '92%', '97%', '100%'],
        '说明': ['与页面价×领取张数交叉验证吻合89%', '少量行缺核销数', '含3行韩币已折算', '计划数量已关联'],
    }), use_container_width=True, hide_index=True)

    st.markdown('### 📐 几个名词的大白话解释')
    terms = pd.DataFrame({
        '名词': ['到手价', '成交额', '核销数', '核销率', '页面价', '价格带'],
        '大白话解释': [
            '顾客实际付的钱（扣完所有优惠）',
            '这个品/网红/品牌一共卖了多少钱',
            '卖出去多少件',
            '优惠券被真正用掉的比例，越高说明带货越实在；超过100%说明中途追加了预算、卖得比预期还好',
            '商品详情页上标的价格',
            '按价格分档：A<$20引流 / B$20-50 / C$50-100 / D$100-200 / E≥$200高客单',
        ],
    })
    st.dataframe(terms, use_container_width=True, hide_index=True)

    st.markdown('### ✅ 数据对账（入库时逐分核对过）')
    st.markdown('<div class="explain">每期的总成交额都和原始表逐分对平，确保没有漏数据、没有算错。</div>',
                unsafe_allow_html=True)
    st.dataframe(pd.DataFrame({
        '期次': ['11月', '3月', '6月', '8月'],
        '总成交额': ['$9,424,564', '$7,279,961', '$5,813,708', '$6,344,204'],
        '备注': ['含4行退款负值，如实保留', '与渠道汇总逐分对平', '含3行韩币已折算', '实际核销值'],
    }), use_container_width=True, hide_index=True)

    st.markdown('### ⚠️ 需要你知道的数据质量提醒')
    st.markdown("""
- 11月有 4 行是**退款/取消**的负数（톡써니테크 / 슈퍼테슬라 / 알리박스 / 겜용이），我们如实保留、没有删掉。
- **더더마**：11月的PID和8月差1位，是不是同一个网红还需要你确认。
- **랑스형님** 和 **랑스알리**：名字很像，但目前没有合并，当成两个网红处理。
- 单个网红×单个商品大多只有 1-2 期数据，**单点结论仅供参考**；跨品聚合的品牌/价格带结论更可靠。
""")


def render():
    inject_css()
    st.markdown('<div class="hub-app">', unsafe_allow_html=True)
    st.markdown("""
    <div class="hub-hero">
      <div class="t">📊 历史大促数据中枢</div>
      <div class="s">4 期大促数据 · 品牌 / 商品 / 网红 / 期次 · 一目了然的决策依据</div>
    </div>""", unsafe_allow_html=True)

    if DB_PATH is None:
        st.error('未找到 history_data.db 数据库文件。请将其放在与 app.py 相同的目录下。')
        return

    tab_d, tab_p, tab_b, tab_c, tab_x, tab_s = st.tabs(
        ['🏠 总览', '🔍 商品档案', '🏷️ 品牌档案', '👤 网红档案', '🧩 维度交叉查询', '📖 口径说明'])

    with tab_d:
        page_dashboard()
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
    st.markdown('</div>', unsafe_allow_html=True)
