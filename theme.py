"""
共享主题样式（薄荷绿 · 游戏像素 · 黑描边 · 硬阴影）
──────────────────────────────────────────────
从 app.py 抽出的全局 CSS，供主站与独立历史价站共用。
用法：import theme; theme.inject()
"""

import streamlit as st

THEME_CSS = """
<style>
    @import url('https://cdn.jsdelivr.net/npm/@fontsource/press-start-2p@5.3.0/index.css');
    @import url('https://cdn.jsdelivr.net/npm/@fontsource/zcool-qingke-huangyou@5.2.6/index.css');

    /* ---------- 全局：薄荷绿棋盘格底色 ---------- */
    .stApp {
        font-family: 'PingFang SC', 'Microsoft YaHei', sans-serif;
        background-color: #9defc4;
        background-image:
            linear-gradient(45deg, #8fe6b8 25%, transparent 25%, transparent 75%, #8fe6b8 75%),
            linear-gradient(45deg, #8fe6b8 25%, transparent 25%, transparent 75%, #8fe6b8 75%);
        background-size: 32px 32px;
        background-position: 0 0, 16px 16px;
        background-attachment: fixed;
    }
    footer { visibility: hidden; }
    #MainMenu { visibility: hidden; }
    header[data-testid="stHeader"] { background: transparent; }

    .block-container { padding-top: 2.5rem; max-width: 1150px; }

    h1, h2, h3, h4 {
        font-family: 'ZCOOL QingKe HuangYou', 'PingFang SC', sans-serif;
        color: #1c1c1e; letter-spacing: 1px;
    }

    /* ---------- 侧边栏 ---------- */
    [data-testid="stSidebar"] {
        background-color: #d8fbe9;
        border-right: 4px solid #1c1c1e;
    }
    [data-testid="stSidebar"] label {
        font-family: 'ZCOOL QingKe HuangYou', 'PingFang SC', sans-serif;
        font-size: 15px !important;
        color: #1c1c1e !important;
    }
    [data-testid="stSidebar"] hr {
        border: none !important;
        border-top: 3px dashed #9fd8bc !important;
        margin: 18px 0 !important;
    }
    [data-testid="stSidebar"] .stButton button {
        background: #1c1c1e !important;
        color: #9defc4 !important;
        box-shadow: 0 6px 0 #4a4a4e !important;
        font-size: 16px !important;
        height: 46px !important;
        min-height: 46px !important;
    }
    [data-testid="stSidebar"] .stButton button:hover { filter: brightness(1.25); }
    [data-testid="stSidebar"] .stButton button:active {
        transform: translateY(6px) !important;
        box-shadow: 0 0 0 transparent !important;
    }

    /* ---------- 标题横幅（游戏机顶栏） ---------- */
    .title-bar {
        background: #1c1c1e;
        border: 4px solid #1c1c1e;
        border-radius: 12px;
        box-shadow: 8px 8px 0 rgba(20, 114, 74, .35);
        padding: 18px 26px;
        display: flex;
        align-items: center;
        gap: 16px;
        flex-wrap: wrap;
    }
    .coin { font-size: 26px; display: inline-block; animation: px-bounce .7s steps(2, jump-none) infinite alternate; }
    .title-en {
        font-family: 'Press Start 2P', monospace;
        font-size: 20px;
        color: #9defc4;
        text-shadow: 3px 3px 0 #14724a;
    }
    .title-cn {
        font-family: 'ZCOOL QingKe HuangYou', 'PingFang SC', sans-serif;
        font-size: 26px;
        color: #fff;
        letter-spacing: 4px;
    }
    .cursor-blink {
        font-family: 'Press Start 2P', monospace;
        color: #ffd93d;
        font-size: 18px;
        animation: px-blink 1s steps(1) infinite;
    }
    .subtitle { font-size: 14px; color: #2c6e4f; font-weight: 600; margin-top: 14px; margin-bottom: 0; }
    .home-hero { justify-content: center; padding: 30px 26px; }
    .home-hero .coin { font-size: 34px; }
    .home-hero .title-en { font-size: 26px; }
    .home-hero .title-cn { font-size: 34px; }
    .home-hint { text-align: center; font-size: 15px; margin-top: 20px; }
    @keyframes px-bounce { from { transform: translateY(0); } to { transform: translateY(-7px); } }
    @keyframes px-blink { 50% { opacity: 0; } }

    /* ---------- 首页模块卡：整卡可点（透明按钮覆盖整卡） ---------- */
    .homecard-marker { display: none; }
    [data-testid="stVerticalBlock"]:has(> .element-container:first-child .homecard-marker) {
        background: #fff;
        border: 4px solid #1c1c1e;
        border-radius: 8px;
        box-shadow: 8px 8px 0 #1c1c1e;
        padding: 32px 30px;
        min-height: 330px;
        position: relative;
        cursor: pointer;
        transition: transform .1s, box-shadow .1s;
    }
    [data-testid="stVerticalBlock"]:has(> .element-container:first-child .homecard-marker):hover {
        transform: translate(-3px, -3px);
        box-shadow: 11px 11px 0 #1c1c1e;
    }
    [data-testid="stVerticalBlock"]:has(> .element-container:first-child .homecard-marker) > .element-container {
        margin: 0 !important;
        /* element-container 自带 position:relative，会劫持透明按钮的定位（塌成一条线），
           必须强制 static，让按钮相对整张卡片定位 */
        position: static !important;
    }
    [data-testid="stVerticalBlock"]:has(> .element-container:first-child .homecard-marker) .stButton {
        margin: 0 !important;
    }
    [data-testid="stVerticalBlock"]:has(> .element-container:first-child .homecard-marker) .stButton button {
        position: absolute;
        top: 0; left: 0; right: 0; bottom: 0;
        width: 100% !important;
        height: 100% !important;
        min-height: 0 !important;
        opacity: 0;
        z-index: 6;
        border-radius: 8px;
    }
    .hc-icon {
        width: 78px; height: 78px;
        border-radius: 14px;
        background: #1c1c1e;
        display: flex; align-items: center; justify-content: center;
        font-size: 38px;
        box-shadow: 4px 4px 0 rgba(20, 114, 74, .35);
        margin-bottom: 18px;
    }
    .hc-title {
        font-family: 'ZCOOL QingKe HuangYou', 'PingFang SC', sans-serif;
        font-size: 25px;
        letter-spacing: 2px;
        color: #1c1c1e;
        margin-bottom: 10px;
    }
    .hc-desc {
        font-size: 13px;
        color: #4c8a6b;
        line-height: 1.9;
        font-weight: 600;
        margin-bottom: 20px;
    }
    .hc-enter {
        display: inline-flex;
        align-items: center;
        font-family: 'ZCOOL QingKe HuangYou', 'PingFang SC', sans-serif;
        font-size: 18px;
        letter-spacing: 3px;
        color: #fff;
        background: #2fbf7f;
        border: 4px solid #1c1c1e;
        border-radius: 12px;
        padding: 8px 30px;
        box-shadow: 0 5px 0 #14724a;
    }

    /* ---------- 返回首页按钮（小号白键帽） ---------- */
    .backbtn-marker { display: none; }
    [data-testid="stVerticalBlock"]:has(> .element-container:first-child .backbtn-marker) > .element-container {
        margin: 0 !important;
    }
    [data-testid="stVerticalBlock"]:has(> .element-container:first-child .backbtn-marker) .stButton button {
        height: 42px !important;
        min-height: 42px !important;
        font-size: 15px !important;
        letter-spacing: 1px;
        padding: 0 20px !important;
        border-radius: 10px !important;
        background: #fff !important;
        color: #1c1c1e !important;
        box-shadow: 0 4px 0 #9dbfae !important;
        width: auto;
        display: inline-flex;
        align-items: center;
    }
    [data-testid="stVerticalBlock"]:has(> .element-container:first-child .backbtn-marker) .stButton button:hover {
        background: #d8fbe9 !important;
    }
    [data-testid="stVerticalBlock"]:has(> .element-container:first-child .backbtn-marker) .stButton button:active {
        transform: translateY(4px) !important;
        box-shadow: 0 0 0 transparent !important;
    }

    /* ---------- 轮次切换按钮组：绿键帽=选中 · 白键帽=未选中 ----------
       注意：1.60 里 button 被 tooltip span 包了几层，必须用后代选择器！ ---------- */
    .modswitch-marker { display: none; }
    [data-testid="stVerticalBlock"]:has(> .element-container .modswitch-marker) .stButton button {
        width: 100% !important;
        height: 52px !important;
        min-height: 52px !important;
        border-radius: 14px !important;
    }
    [data-testid="stVerticalBlock"]:has(> .element-container .modswitch-marker) .stButton button[data-testid="stBaseButton-primary"] {
        background: #2fbf7f !important;
        color: #fff !important;
        box-shadow: 0 6px 0 #14724a !important;
    }
    [data-testid="stVerticalBlock"]:has(> .element-container .modswitch-marker) .stButton button:not([data-testid="stBaseButton-primary"]) {
        background: #fff !important;
        color: #1c1c1e !important;
        box-shadow: 0 6px 0 #9dbfae !important;
    }
    [data-testid="stVerticalBlock"]:has(> .element-container .modswitch-marker) .stButton button:not([data-testid="stBaseButton-primary"]):hover {
        background: #d8fbe9 !important;
    }

    /* ---------- 游戏机键帽按钮（全局） ---------- */
    .stButton button, [data-testid="stDownloadButton"] button {
        font-family: 'ZCOOL QingKe HuangYou', 'PingFang SC', sans-serif !important;
        font-size: 18px !important;
        letter-spacing: 3px;
        border: 4px solid #1c1c1e !important;
        border-radius: 14px !important;
        height: 52px !important;
        min-height: 52px !important;
        transition: transform .06s, box-shadow .06s, filter .15s;
    }
    .stButton button p, [data-testid="stDownloadButton"] button p {
        color: inherit !important;
        font-family: inherit !important;
    }
    button[data-testid="stBaseButton-primary"] {
        background: #2fbf7f !important;
        color: #fff !important;
        box-shadow: 0 6px 0 #14724a !important;
    }
    button[data-testid="stBaseButton-primary"]:hover { filter: brightness(1.06); }
    button[data-testid="stBaseButton-primary"]:active {
        transform: translateY(6px) !important;
        box-shadow: 0 0 0 transparent !important;
    }
    button[data-testid="stBaseButton-secondary"] {
        background: #fff !important;
        color: #1c1c1e !important;
        box-shadow: 0 6px 0 #9dbfae !important;
    }
    button[data-testid="stBaseButton-secondary"]:hover { background: #d8fbe9 !important; }
    button[data-testid="stBaseButton-secondary"]:active {
        transform: translateY(6px) !important;
        box-shadow: 0 0 0 transparent !important;
    }
    [data-testid="stDownloadButton"] button {
        background: #ffd93d !important;
        color: #1c1c1e !important;
        box-shadow: 0 6px 0 #b8930a !important;
    }
    [data-testid="stDownloadButton"] button:hover { filter: brightness(1.05); }
    [data-testid="stDownloadButton"] button:active {
        transform: translateY(6px) !important;
        box-shadow: 0 0 0 transparent !important;
    }

    /* ---------- Tabs：白键帽=未选中 · 黑键帽=选中 ---------- */
    .stTabs [role="tablist"] {
        gap: 12px;
        border-bottom: none;
        background: transparent;
        padding: 0;
        display: flex;
    }
    .stTabs [role="tab"] {
        flex: 1;
        border-radius: 12px !important;
        padding: 12px 0 !important;
        font-family: 'ZCOOL QingKe HuangYou', 'PingFang SC', sans-serif;
        font-size: 16px;
        letter-spacing: 2px;
        background: #fff !important;
        color: #1c1c1e !important;
        border: 4px solid #1c1c1e !important;
        justify-content: center;
        box-shadow: 4px 4px 0 #1c1c1e;
        transition: all .12s;
    }
    .stTabs [role="tab"]:hover { background: #d8fbe9 !important; color: #1c1c1e !important; }
    .stTabs [role="tab"][aria-selected="true"] {
        background: #1c1c1e !important;
        color: #9defc4 !important;
        box-shadow: 4px 4px 0 rgba(20, 114, 74, .35);
    }
    .stTabs [role="tab"] p { color: inherit; }
    /* 隐藏默认下划线指示器（1.60 新结构） */
    .stTabs .react-aria-SelectionIndicator { display: none !important; }

    /* ---------- 指标卡 ---------- */
    div[data-testid="stMetric"] {
        background: #f0fdf7;
        border: 4px solid #1c1c1e;
        border-radius: 10px;
        padding: 16px 20px;
        box-shadow: 4px 4px 0 #1c1c1e;
    }
    div[data-testid="stMetricLabel"] { color: #4c8a6b; font-weight: 700; }
    /* 指标大数字：用易读的粗黑体，不用装饰像素字体（数字笔画花，久看费劲） */
    div[data-testid="stMetricValue"] {
        color: #14724a;
        font-family: -apple-system, 'PingFang SC', 'Microsoft YaHei', sans-serif;
        font-weight: 800;
        letter-spacing: 0;
    }

    /* ---------- 数据表 ---------- */
    div[data-testid="stDataFrame"] {
        background: #fff;
        border: 4px solid #1c1c1e;
        border-radius: 8px;
        overflow: hidden;
        box-shadow: 4px 4px 0 rgba(28, 28, 30, .2);
    }

    /* ---------- 上传框（1.60 外层是 div 不是 section） ---------- */
    [data-testid="stFileUploader"] > div {
        border: 4px dashed #1c1c1e !important;
        border-radius: 10px !important;
        background: #f0fdf7 !important;
    }
    [data-testid="stFileUploader"] button {
        font-family: 'ZCOOL QingKe HuangYou', 'PingFang SC', sans-serif !important;
        border: 3px solid #1c1c1e !important;
        border-radius: 10px !important;
        background: #fff !important;
        color: #1c1c1e !important;
        box-shadow: 0 4px 0 #9dbfae !important;
    }

    /* ---------- 输入框 / 数字框 ---------- */
    .stTextInput input, .stNumberInput input {
        border: 4px solid #1c1c1e !important;
        border-radius: 10px !important;
        background: #fff !important;
        height: 44px;
        font-weight: 700;
    }
    .stTextInput input:focus, .stNumberInput input:focus, .stTextArea textarea:focus {
        border-color: #2fbf7f !important;
        box-shadow: 4px 4px 0 #2fbf7f !important;
    }
    .stTextArea textarea {
        border: 4px solid #1c1c1e !important;
        border-radius: 10px !important;
        background: #fff !important;
    }
    /* 数字框：隐藏 −/+ 按钮（1.60 改名 stNumberInputStepUp/Down） */
    [data-testid="stNumberInputStepUp"], [data-testid="stNumberInputStepDown"] {
        display: none !important;
    }
    .stNumberInput input { padding-right: 20px !important; }
    /* 密码框（Naver Secret）：描边挪到外层容器 */
    [data-testid="stTextInputRootElement"]:has(input[type="password"]) {
        border: 4px solid #1c1c1e !important;
        border-radius: 10px !important;
        background: #fff !important;
        height: 44px;
        align-items: center;
    }
    [data-testid="stTextInputRootElement"]:has(input[type="password"]) input {
        border: none !important;
        box-shadow: none !important;
        background: transparent !important;
        border-radius: 0 !important;
        height: 38px !important;
    }
    [data-testid="stTextInputRootElement"]:has(input[type="password"]):focus-within {
        border-color: #2fbf7f !important;
        box-shadow: 4px 4px 0 #2fbf7f !important;
    }

    /* ---------- 下拉框 / 多选框（1.60 React Aria 结构） ---------- */
    [data-testid="stSelectbox"] [role="group"],
    [data-testid="stMultiSelect"] [role="group"],
    div[data-baseweb="select"] > div {
        border-radius: 10px !important;
        min-height: 44px;
        border: 4px solid #1c1c1e !important;
        background: #fff !important;
        font-weight: 700;
    }
    div[data-baseweb="popover"] > ul {
        border-radius: 10px;
        border: 3px solid #1c1c1e;
    }

    /* ---------- 提示条 / 展开器 ---------- */
    div[data-testid="stAlert"], .stAlert {
        border-radius: 10px !important;
        border: 3px solid #1c1c1e !important;
        box-shadow: 4px 4px 0 rgba(20, 114, 74, .25);
    }
    [data-testid="stExpander"] {
        border-radius: 8px !important;
        border: 3px solid #1c1c1e !important;
        background: #f0fdf7 !important;
    }
    [data-testid="stExpander"] summary {
        font-family: 'ZCOOL QingKe HuangYou', 'PingFang SC', sans-serif;
        font-size: 15px;
    }
    pre {
        border-radius: 8px !important;
        border: 3px solid #1c1c1e !important;
        background: #d8fbe9 !important;
    }

    /* ---------- 进度条（XP 条纹） ---------- */
    [data-testid="stProgress"] .react-aria-ProgressBar {
        border: 4px solid #1c1c1e;
        border-radius: 10px;
        background: #e8fff3;
        height: 36px;
        padding: 3px;
    }
    [data-testid="stProgress"] .react-aria-ProgressBar > div:last-child {
        background: repeating-linear-gradient(45deg, #2fbf7f 0 12px, #26a86e 12px 24px) !important;
        border-radius: 5px;
        animation: px-stripes .8s linear infinite;
    }
    @keyframes px-stripes { to { background-position: 34px 0; } }

    hr { border: none; border-top: 3px dashed #9fd8bc; }
</style>
"""


def inject():
    st.markdown(THEME_CSS, unsafe_allow_html=True)
