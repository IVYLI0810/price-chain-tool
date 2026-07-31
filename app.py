"""
Naver 底价查询工具 - 团队版 v2
薄荷绿 · 游戏像素风 Streamlit 应用
升级：品牌+型号锚点搜索 / 匹配度验证 / Excel导出
"""

import streamlit as st
import requests
import csv
import io
import time
import re
import html as html_lib
from urllib.parse import quote
import pandas as pd

# ==================== 页面基础配置 ====================
st.set_page_config(
    page_title="Naver 底价查询",
    page_icon="🎮",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ==================== 薄荷绿 · 游戏像素风样式 ====================
GAME_CSS = """
<style>
    @import url('https://cdn.jsdelivr.net/npm/@fontsource/press-start-2p@5.3.0/index.css');
    @import url('https://cdn.jsdelivr.net/npm/@fontsource/zcool-qingke-huangyou@5.2.6/index.css');

    /* ---------- 全局：薄荷绿棋盘格背景 ---------- */
    .stApp {
        font-family: 'PingFang SC', 'Microsoft YaHei', sans-serif;
        background-color: #9defc4;
        background-image:
            linear-gradient(45deg, #8fe6b8 25%, transparent 25%, transparent 75%, #8fe6b8 75%),
            linear-gradient(45deg, #8fe6b8 25%, transparent 25%, transparent 75%, #8fe6b8 75%);
        background-size: 32px 32px;
        background-position: 0 0, 16px 16px;
    }
    footer {visibility: hidden;}
    #MainMenu {visibility: hidden;}
    header[data-testid="stHeader"] { background: transparent; }

    .block-container {
        padding-top: 2.5rem;
        max-width: 960px;
    }

    /* ---------- 侧边栏 ---------- */
    [data-testid="stSidebar"] {
        background-color: #d8fbe9;
        border-right: 4px solid #1c1c1e;
    }
    .side-title {
        font-family: 'ZCOOL QingKe HuangYou', 'PingFang SC', sans-serif;
        font-size: 22px;
        letter-spacing: 2px;
        color: #1c1c1e;
        margin-bottom: 4px;
    }
    .side-note { font-size: 12px; color: #4c8a6b; line-height: 1.7; }
    .side-help { font-size: 12px; color: #4c8a6b; line-height: 1.9; }
    .side-help b { color: #1c1c1e; }
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

    /* ---------- 输入框（像素描边） ---------- */
    [data-testid="stTextInput"] input {
        border: 4px solid #1c1c1e !important;
        border-radius: 10px !important;
        background: #fff !important;
        font-size: 13px;
    }
    [data-testid="stTextInput"] input:focus {
        border-color: #2fbf7f !important;
        box-shadow: 4px 4px 0 #2fbf7f !important;
    }
    [data-testid="stTextInputRootElement"]:has(input[type="password"]) {
        border: 4px solid #1c1c1e !important;
        border-radius: 10px !important;
        background: #fff !important;
    }
    [data-testid="stTextInputRootElement"]:has(input[type="password"]) input {
        border: none !important;
        box-shadow: none !important;
        background: transparent !important;
    }
    [data-testid="stTextInputRootElement"]:has(input[type="password"]):focus-within {
        border-color: #2fbf7f !important;
        box-shadow: 4px 4px 0 #2fbf7f !important;
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
    .coin {
        font-size: 26px;
        display: inline-block;
        animation: px-bounce .7s steps(2, jump-none) infinite alternate;
    }
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
    .subtitle {
        font-size: 14px;
        color: #2c6e4f;
        font-weight: 600;
        margin-top: 14px;
        margin-bottom: 0;
    }
    @keyframes px-bounce { from { transform: translateY(0); } to { transform: translateY(-7px); } }
    @keyframes px-blink { 50% { opacity: 0; } }

    /* ---------- 像素卡片 ---------- */
    .px-marker { display: none; }
    [data-testid="stVerticalBlock"]:has(> .element-container:first-child .px-marker) {
        background: #fff;
        border: 4px solid #1c1c1e;
        border-radius: 8px;
        box-shadow: 8px 8px 0 #1c1c1e;
        padding: 26px 30px;
        margin-top: 34px;
        gap: 22px;
    }
    [data-testid="stVerticalBlock"]:has(> .element-container:first-child .px-marker) > .element-container {
        margin: 0 !important;
    }
    [data-testid="stVerticalBlock"]:has(> .element-container:first-child .px-marker) [data-testid="stHorizontalBlock"] {
        margin: 0 !important;
        gap: 14px;
    }

    /* ---------- 步骤标题 ---------- */
    .step { display: flex; align-items: center; gap: 12px; flex-wrap: wrap; }
    .step-num {
        font-family: 'Press Start 2P', monospace;
        font-size: 13px;
        color: #fff;
        background: #2fbf7f;
        border: 3px solid #1c1c1e;
        border-radius: 8px;
        width: 36px;
        height: 36px;
        display: inline-flex;
        align-items: center;
        justify-content: center;
        box-shadow: 3px 3px 0 #1c1c1e;
        flex-shrink: 0;
    }
    .step-title {
        font-family: 'ZCOOL QingKe HuangYou', 'PingFang SC', sans-serif;
        font-size: 22px;
        letter-spacing: 2px;
        color: #1c1c1e;
    }
    .step-sub { font-size: 12px; color: #7aa88f; font-weight: 600; }
    .count-badge {
        font-family: 'Press Start 2P', monospace;
        font-size: 12px;
        background: #1c1c1e;
        color: #9defc4;
        border-radius: 6px;
        padding: 6px 10px;
    }

    /* ---------- 上传区 ---------- */
    .upload-tip { font-size: 14px; font-weight: 600; color: #1c1c1e; }
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
    .file-chip {
        display: inline-flex;
        align-items: center;
        gap: 8px;
        background: #ffd93d;
        border: 3px solid #1c1c1e;
        border-radius: 8px;
        box-shadow: 3px 3px 0 #1c1c1e;
        padding: 6px 14px;
        font-size: 13px;
        font-weight: 700;
        color: #1c1c1e;
    }

    /* ---------- 游戏机键帽按钮 ---------- */
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
    button[data-testid="stBaseButton-primary"]:hover {
        background: #2fbf7f !important;
        color: #fff !important;
        filter: brightness(1.06);
    }
    button[data-testid="stBaseButton-primary"]:active {
        transform: translateY(6px) !important;
        box-shadow: 0 0 0 transparent !important;
    }
    [data-testid="stDownloadButton"] button {
        background: #ffd93d !important;
        color: #1c1c1e !important;
        box-shadow: 0 6px 0 #b8930a !important;
    }
    [data-testid="stDownloadButton"] button:hover {
        background: #ffd93d !important;
        color: #1c1c1e !important;
        filter: brightness(1.05);
    }
    [data-testid="stDownloadButton"] button:active {
        transform: translateY(6px) !important;
        box-shadow: 0 0 0 transparent !important;
    }
    [data-testid="stSidebar"] .stButton button {
        background: #1c1c1e !important;
        color: #9defc4 !important;
        box-shadow: 0 6px 0 #4a4a4e !important;
        font-size: 16px !important;
        height: 46px !important;
        min-height: 46px !important;
    }
    [data-testid="stSidebar"] .stButton button:hover {
        background: #1c1c1e !important;
        color: #9defc4 !important;
        filter: brightness(1.25);
    }
    [data-testid="stSidebar"] .stButton button:active {
        transform: translateY(6px) !important;
        box-shadow: 0 0 0 transparent !important;
    }

    /* ---------- 经验值进度条 ---------- */
    .xp-label {
        font-family: 'Press Start 2P', monospace;
        font-size: 11px;
        color: #14724a;
        display: flex;
        justify-content: space-between;
        margin-bottom: 8px;
    }
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

    /* ---------- 统计像素块 ---------- */
    .stat {
        border: 3px solid #1c1c1e;
        border-radius: 10px;
        box-shadow: 4px 4px 0 #1c1c1e;
        background: #f0fdf7;
        padding: 14px 10px;
        text-align: center;
    }
    .stat-icon { font-size: 18px; }
    .stat-num {
        font-family: 'Press Start 2P', monospace;
        font-size: 22px;
        margin: 8px 0 6px;
        color: #1c1c1e;
    }
    .stat-green .stat-num { color: #14724a; }
    .stat-coral .stat-num { color: #ff7b6b; }
    .stat-label { font-size: 12px; color: #5c8a72; font-weight: 700; }

    /* ---------- 结果表格 ---------- */
    .px-table-wrap { overflow-x: auto; }
    .px-table {
        width: 100%;
        border-collapse: collapse;
        border: 4px solid #1c1c1e;
        font-size: 13px;
    }
    .px-table th {
        background: #1c1c1e;
        color: #9defc4;
        font-family: 'ZCOOL QingKe HuangYou', 'PingFang SC', sans-serif;
        font-weight: 400;
        letter-spacing: 1px;
        padding: 10px 12px;
        text-align: left;
        border: 2px solid #1c1c1e;
        white-space: nowrap;
    }
    .px-table td {
        padding: 10px 12px;
        border: 2px solid #d5efe2;
        color: #1c1c1e;
    }
    .px-table tbody tr:nth-child(even) { background: #f0fdf7; }
    .px-tag {
        display: inline-block;
        border: 2px solid #1c1c1e;
        border-radius: 6px;
        padding: 2px 8px;
        font-size: 11px;
        font-weight: 700;
        background: #d8fbe9;
        white-space: nowrap;
    }
    .px-price {
        font-family: 'Press Start 2P', monospace;
        font-size: 11px;
        color: #14724a;
        white-space: nowrap;
    }
    .px-err { color: #ff7b6b; font-weight: 700; }
    .px-note { font-size: 12px; color: #7aa88f; }
    .match-high { color: #14724a; font-weight: 700; }
    .match-mid { color: #b8860b; font-weight: 700; }
    .match-low { color: #ff7b6b; font-weight: 700; }

    /* ---------- 展开框 / 提示框 ---------- */
    [data-testid="stExpander"] {
        border: 3px solid #1c1c1e !important;
        border-radius: 8px !important;
        background: #f0fdf7 !important;
    }
    [data-testid="stExpander"] summary {
        font-family: 'ZCOOL QingKe HuangYou', 'PingFang SC', sans-serif;
        font-size: 15px;
    }
    [data-testid="stAlert"] {
        border: 3px solid #1c1c1e !important;
        border-radius: 10px !important;
        box-shadow: 4px 4px 0 #1c1c1e;
    }
</style>
"""
st.markdown(GAME_CSS, unsafe_allow_html=True)

# ==================== 屏蔽词（二手/翻新） ====================
BLOCK_WORDS = ["중고", "리퍼", "박스훼손", "렌탈", "중고나라", "당근", "번개장터"]

# ==================== 泛用词（搜索时排除） ====================
GENERIC_WORDS = {
    "야외", "캠핑", "피크닉", "여행", "휴대용", "다기능",
    "스테인리스", "스틸", "대용량", "경량", "방수", "미니",
    "인증", "정품", "무료", "당일", "특가", "세일",
    "새로운", "할인", "인기", "추천", "베스트", "신상",
    "고품질", "프리미엄", "스마트", "자동", "초경량",
}

# 品类词（保留，帮助Naver定位品类）
CATEGORY_WORDS = {
    "태블릿", "스피커", "이어폰", "헤드폰", "모니터", "키보드", "마우스",
    "청소기", "공기청정기", "가습기", "선풍기", "노트북", "데스크탑",
    "충전기", "보조배터리", "파워뱅크", "케이블", "어댑터", "허브",
    "카메라", "액션캠", "드론", "자전거", "트레이너", "펌프",
    "tablet", "speaker", "earbuds", "monitor", "keyboard",
}


# ==================== 品牌+型号提取引擎 ====================

# 韩文品牌名 → 英文搜索名映射（Naver搜英文品牌名命中率更高）
KR_BRAND_MAP = {
    "벤션": "Vention", "베이스어스": "Baseus", "로지텍": "Logitech",
    "레노버": "Lenovo", "에디파이어": "Edifier", "올독큐브": "ALLDOCUBE",
    "샤오미": "Xiaomi", "미지아": "MIJIA", "제이오엔알": "JONR",
    "유팡": "Uwant", "지엠텍": "GMKtec", "파이어뱃": "Firebat",
    "넷택": "Netac", "에이엠디": "AMD", "시마노": "SHIMANO",
    "에텐울프": "ETENWOLF", "이텐울프": "ETENWOLF",
    "큐사이클": "CYCPLUS", "라이드나우": "RIDENOW",
    "투키": "TOOCKI", "엘디니오": "LDNIO",
    "유 perfect": "Uperfect", "유퍼펙트": "Uperfect",
    "다이브디어": "DIVEDEER", "빅미": "BIGME",
    "오르티잔": "ORTIZAN", "무브스피드": "MOVESPEED",
    "킹스펙": "KingSpec", "주호": "JUHOR",
    "에프오에스": "Fosi Audio", "포시오디오": "Fosi Audio",
    "레이네오": "RAYNEO", "이피지": "EPZ",
    "바토카": "VATOKA", "테일리": "TAILI",
    "웨스턴디지털": "WD", "유그린": "UGREEN",
    "큐씨와이": "QCY", "세븐티마이": "70mai",
}

# 技术术语/通用缩写 — 绝不能当品牌或型号
TECH_TERMS = {
    "gan", "usb", "pd", "pd3", "qc", "qc4", "ccc", "kc", "ce", "fcc", "rohs",
    "ble", "ant", "wifi", "bt", "nfc", "gps", "led", "lcd", "oled",
    "ips", "hdmi", "dp", "vga", "ssd", "hdd", "ram", "rom",
    "type", "mini", "pro", "max", "plus", "ultra", "gen",
    "cn", "kr", "eu", "us", "uk", "au", "qi2",
    # 手机型号（兼容性描述里出现，不是产品本身型号）
    "s25", "s24", "s23", "s22", "note",
    # 规格描述词
    "in", "pin", "port", "way", "core", "thread",
    # 接口/协议/插头类型
    "pps", "rj45", "usb-a", "usb-c", "type-c", "type-a", "cn-eu", "cn-us", "cn-kr",
    "eu", "uk", "au", "ai", "anc", "enc", "ipx", "ipx7", "ip68",
    "mah", "w", "v", "a", "hz", "ghz", "mhz",
}

# 已知英文品牌名（直接匹配，优先级最高）
KNOWN_BRANDS = {
    "lenovo", "edifier", "baseus", "vention", "logitech", "allodcube",
    "gmktec", "firebat", "jonr", "uwant", "mijia", "xiaomi",
    "netac", "amd", "shimano", "etenwolf", "cycplus", "ridenow",
    "toocki", "ldnio", "uperfect", "divedeer", "bigme", "ortizan",
    "movespeed", "kingspec", "juhor", "fosi", "rayneo", "epz",
    "vatoka", "taili", "ugreen", "qcy", "70mai", "asometech",
    "ryet", "allodcube", "westdigital",
}


def extract_brand_model(product_name):
    """
    从商品名中提取品牌名和型号作为搜索锚点。
    返回 (brand, model, category_hint)
    """
    if not product_name or str(product_name).strip() in ("nan", ""):
        return "", "", ""

    name = str(product_name).strip()
    # 去掉内部标注
    name = re.sub(r"\d+月最终价[：:]?\s*[\d.]+", "", name)
    name = re.sub(r"[*·|]+", " ", name)
    name = re.sub(r"【.*?】", " ", name)
    name = re.sub(r"\[.*?\]", " ", name)
    name = re.sub(r"\(Coming Soon\)", "", name, flags=re.IGNORECASE)
    name = re.sub(r"\s+", " ", name).strip()

    words = name.split()

    # --- 提取品牌 ---
    brand = ""

    # 策略0: 韩文品牌名（在前3个词里找，因为可能有"2026년형"等前缀）
    for w in words[:3]:
        w_clean = w.strip(",.!?()")
        if w_clean in KR_BRAND_MAP:
            brand = KR_BRAND_MAP[w_clean]
            break

    # 策略1: 已知英文品牌名（在前5个词里找）
    if not brand:
        for w in words[:5]:
            w_clean = w.strip(",.!?()").lower()
            if w_clean in KNOWN_BRANDS:
                brand = w.strip(",.!?()")
                break

    # 策略2: 前3个词里找第一个"像品牌"的英文词（排除技术术语）
    if not brand:
        for w in words[:3]:
            w_clean = w.strip(",.!?()")
            if re.match(r'^[A-Za-z][A-Za-z0-9\-\.]+$', w_clean) and len(w_clean) >= 3:
                if (w_clean.lower() not in GENERIC_WORDS
                        and w_clean.lower() not in CATEGORY_WORDS
                        and w_clean.lower() not in TECH_TERMS):
                    brand = w_clean
                    break

    # 策略3: 更宽范围（前6个词），要求>=4字符
    if not brand:
        for w in words[:6]:
            w_clean = w.strip(",.!?()")
            if re.match(r'^[A-Za-z][A-Za-z0-9\-\.]+$', w_clean) and len(w_clean) >= 4:
                if (w_clean.lower() not in GENERIC_WORDS
                        and w_clean.lower() not in CATEGORY_WORDS
                        and w_clean.lower() not in TECH_TERMS):
                    brand = w_clean
                    break

    # --- 提取型号 ---
    model = ""

    def _is_real_model(candidate):
        """排除瓦数/电压/规格描述被误认为型号"""
        c = candidate.lower().strip()
        if c in TECH_TERMS:
            return False
        if c == brand.lower():
            return False
        if len(c) < 2 or c.isdigit():
            return False
        # 排除纯瓦数: 70W, 2500W, 100W, 22.5w, 65W
        if re.match(r'^\d+[\.\d]*\s*[wW]$', candidate):
            return False
        # 排除 "in-1" / "in-10" 碎片 (来自 "10-in-1")
        if re.match(r'^in[\-]?\d*$', c):
            return False
        # 排除 mAh 规格
        if re.match(r'^\d+mah$', c):
            return False
        # 排除纯数字+单位 (2500W, 10000mAh)
        if re.match(r'^\d+[a-z]+$', c) and not re.search(r'[a-z]\d', c):
            return False
        return True

    # 型号模式: 字母+数字 (M90, Y700, K12, MR4, GT13, F1, S4, T7, V3, iPlay80, TB376FC)
    model_pattern = re.compile(
        r'\b([A-Za-z]{1,5}[\-]?\d{1,4}[A-Za-z0-9]{0,6}(?:\s*(?:Gen\d?|세대|Pro|Max|Plus|Ultra|MKII|Mini|mini))?)\b',
        re.IGNORECASE
    )
    for w in words:
        m = model_pattern.search(w)
        if m:
            candidate = m.group(1)
            if _is_real_model(candidate):
                model = candidate
                break

    # 策略2: 品牌后面紧跟的alphanumeric token（如 "JONR P20 Pro"）
    if not model and brand:
        found_brand = False
        for w in words:
            w_clean = w.strip(",.!?()")
            if w_clean.lower() == brand.lower():
                found_brand = True
                continue
            if found_brand:
                if re.match(r'^[A-Za-z0-9\-]+$', w_clean) and len(w_clean) >= 2:
                    if (w_clean.lower() not in GENERIC_WORDS
                            and w_clean.lower() not in CATEGORY_WORDS
                            and _is_real_model(w_clean)):
                        model = w_clean
                        break

    # 去重: 如果model和brand一样，清空model
    if model and brand and model.lower() == brand.lower():
        model = ""

    # --- 品类提示 ---
    category_hint = ""
    name_lower = name.lower()
    for cw in CATEGORY_WORDS:
        if cw in name_lower:
            category_hint = cw
            break

    return brand, model, category_hint


def build_query(product_name, sku_option=""):
    """
    构建精准搜索词：品牌 + 型号 + 品类 + SKU核心属性
    """
    brand, model, category_hint = extract_brand_model(product_name)

    parts = []
    if brand:
        parts.append(brand)
    if model:
        parts.append(model)
    if category_hint and category_hint not in parts:
        parts.append(category_hint)

    # SKU属性（颜色/容量/尺寸）
    sku_attrs = []
    if sku_option and str(sku_option).strip() not in ("nan", "", "单一sku", "단일sku", "全部sku"):
        sku = str(sku_option).strip()
        for seg in re.split(r'[/,|]', sku):
            seg = seg.strip()
            if len(seg) > 1 and seg.lower() not in GENERIC_WORDS:
                if re.search(r'\d', seg) or (seg[0].isupper() and seg.isalpha()):
                    sku_attrs.append(seg)
                elif len(seg) <= 12:
                    sku_attrs.append(seg)
    for attr in sku_attrs[:2]:
        if attr not in parts:
            parts.append(attr)

    # 如果品牌型号都没提取到，退回核心词模式
    if not brand and not model:
        if product_name and str(product_name).strip() not in ("nan", ""):
            name_clean = re.sub(r"\d+月最终价[：:]?\s*[\d.]+", "", str(product_name))
            name_clean = re.sub(r"[*·|【】\[\]]+", " ", name_clean)
            words = name_clean.strip().split()
            core = [w for w in words if w.lower() not in GENERIC_WORDS and len(w) > 1]
            parts = core[:5]
            for attr in sku_attrs[:2]:
                if attr not in parts:
                    parts.append(attr)

    query = " ".join(parts)
    return query[:80] if len(query) > 80 else query


def compute_match_score(query_info, result_title):
    """
    计算搜索结果与目标商品的匹配度分数。
    query_info: dict with keys: brand, model, category, sku_attrs
    返回 (score_pct: 0-100, label: str)
    """
    title_lower = result_title.lower().replace("-", "").replace(" ", "")
    score = 0
    max_score = 0

    brand = query_info.get("brand", "")
    model = query_info.get("model", "")
    category = query_info.get("category", "")
    sku_attrs = query_info.get("sku_attrs", [])

    # 品牌匹配 (权重40)
    if brand:
        max_score += 40
        b = brand.lower().replace("-", "").replace(" ", "")
        if b in title_lower:
            score += 40
        elif b[:4] in title_lower:
            score += 20

    # 型号匹配 (权重40)
    if model:
        max_score += 40
        m = model.lower().replace("-", "").replace(" ", "")
        if m in title_lower:
            score += 40
        elif m[:3] in title_lower:
            score += 25

    # 品类匹配 (权重10)
    if category:
        max_score += 10
        if category.lower() in result_title.lower():
            score += 10

    # SKU属性匹配 (权重10)
    if sku_attrs:
        max_score += 10
        hit = sum(1 for a in sku_attrs if a.lower() in result_title.lower())
        score += int(10 * hit / len(sku_attrs))

    if max_score == 0:
        return 50, "⚪ 无法验证"

    pct = int(score / max_score * 100)
    if pct >= 70:
        return pct, "✅ 高匹配"
    elif pct >= 40:
        return pct, "⚠️ 需确认"
    else:
        return pct, "❌ 低匹配"


# ==================== 核心查询逻辑 ====================
def clean_html(raw_html):
    cleanr = re.compile("<.*?>")
    return re.sub(cleanr, "", raw_html)


def extract_shipping_tag(title):
    title_lower = title.lower()
    if "로켓배송" in title_lower:
        return "🚀 火箭配送"
    elif "무료배송" in title_lower or "배송비포함" in title_lower:
        return "📦 包邮"
    elif "배송비별도" in title_lower or "유료배송" in title_lower:
        return "⚠️ 运费另算"
    else:
        return "❓ 未标明"


def get_lowest_price(keyword, client_id, client_secret, query_info=None):
    """查询 Naver Shopping 最低价，含匹配度验证"""
    encoded_query = quote(keyword)
    url = f"https://openapi.naver.com/v1/search/shop.json?query={encoded_query}&display=20&sort=asc"
    headers = {"X-Naver-Client-Id": client_id, "X-Naver-Client-Secret": client_secret}

    try:
        response = requests.get(url, headers=headers, timeout=5)
        if response.status_code != 200:
            return {"status": "error", "note": f"API错误({response.status_code})"}

        items = response.json().get("items", [])
        valid = []
        for i in items:
            title_clean = clean_html(i["title"])
            if any(x in title_clean.lower() for x in BLOCK_WORDS):
                continue
            link = i.get("link", "")
            if "aliexpress" in link or "aliexp" in link:
                continue
            valid.append({"data": i, "clean_title": title_clean})

        if not valid:
            return {"status": "empty", "note": "无匹配商品(或全是二手)"}

        valid.sort(key=lambda x: int(x["data"]["lprice"]))
        best = valid[0]["data"]
        best_title = valid[0]["clean_title"]

        # 匹配度验证
        match_pct, match_label = 50, "⚪ 无法验证"
        if query_info:
            match_pct, match_label = compute_match_score(query_info, best_title)

        return {
            "status": "ok",
            "mall": best["mallName"],
            "price": int(best["lprice"]),
            "shipping": extract_shipping_tag(best_title),
            "title": best_title,
            "reviews": best.get("reviewCount", "0"),
            "link": best["link"],
            "match_pct": match_pct,
            "match_label": match_label,
        }
    except requests.exceptions.Timeout:
        return {"status": "error", "note": "查询超时"}
    except Exception as e:
        return {"status": "error", "note": f"出错: {str(e)[:30]}"}


def make_result_excel(rows, rate=1550.0):
    """生成带颜色的 Excel 结果文件"""
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

    wb = Workbook()
    ws = wb.active
    ws.title = "Naver最低价"

    headers = ["搜索词", "匹配度", "最低价商城", "最低价(KRW)", "≈USD", "物流标签", "商品原标题", "评价数", "购买链接"]
    header_fill = PatternFill("solid", fgColor="1C1C1E")
    header_font = Font(bold=True, color="9DEFC4", size=10)
    green_fill = PatternFill("solid", fgColor="E8F5E9")
    yellow_fill = PatternFill("solid", fgColor="FFF9C4")
    red_fill = PatternFill("solid", fgColor="FFE3E3")
    thin_border = Border(
        left=Side("thin"), right=Side("thin"),
        top=Side("thin"), bottom=Side("thin")
    )

    for col, h in enumerate(headers, 1):
        cell = ws.cell(1, col, h)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border

    for row_idx, r in enumerate(rows, 2):
        sku, match_label, mall, price, ship, title, rev, link = r
        usd = f"{int(price) / rate:.2f}" if str(price).isdigit() else ""
        data = [sku, match_label, mall, price, usd, ship, title, rev, link]
        for col, val in enumerate(data, 1):
            cell = ws.cell(row_idx, col, val)
            cell.border = thin_border

        # 按匹配度着色
        if "高匹配" in str(match_label):
            fill = green_fill
        elif "需确认" in str(match_label):
            fill = yellow_fill
        elif "低匹配" in str(match_label) or "API" in str(mall) or "无匹配" in str(mall):
            fill = red_fill
        else:
            fill = None
        if fill:
            for col in range(1, len(headers) + 1):
                ws.cell(row_idx, col).fill = fill

    # 列宽
    widths = [35, 12, 14, 12, 10, 12, 50, 8, 40]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[chr(64 + i)].width = w

    ws.freeze_panes = "A2"

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def esc(text):
    """HTML 转义"""
    return html_lib.escape(str(text))


# ==================== 侧边栏：API 设置 ====================
with st.sidebar:
    st.markdown('<div class="side-title">⚙️ API 设置</div>', unsafe_allow_html=True)
    st.markdown('<div class="side-note">密钥只存在浏览器会话中，刷新页面后需重新填写。</div>', unsafe_allow_html=True)

    client_id = st.text_input(
        "Client ID",
        value=st.session_state.get("naver_id", ""),
        placeholder="Naver 开发者平台的 Client ID",
        key="input_id",
    )
    client_secret = st.text_input(
        "Client Secret",
        value=st.session_state.get("naver_secret", ""),
        type="password",
        placeholder="Naver 开发者平台的 Client Secret",
        key="input_secret",
    )

    if st.button("💾 保存密钥", width="stretch"):
        st.session_state["naver_id"] = client_id.strip()
        st.session_state["naver_secret"] = client_secret.strip()
        st.success("已保存 ✓")

    st.divider()
    exchange_rate = st.number_input(
        "汇率 (KRW → USD)",
        min_value=1000.0,
        max_value=2000.0,
        value=1550.0,
        step=10.0,
        help="韩元÷此汇率=美金，用于和AE价格对比",
    )

    st.divider()
    st.markdown(
        '<div class="side-help"><b>使用说明</b><br>'
        '1. 填入 Naver API 密钥并保存<br>'
        '2. 上传商品表格（自动识别商品名+SKU列）<br>'
        '3. 确认搜索词后开始查询<br>'
        '4. 下载带匹配度标注的 Excel 结果</div>',
        unsafe_allow_html=True,
    )

# ==================== 主页面 ====================
st.markdown(
    '<div class="title-bar"><span class="coin">🪙</span><span class="title-en">NAVER</span>'
    '<span class="title-cn">底价查询</span><span class="cursor-blink">▮</span></div>'
    '<p class="subtitle">▸ 品牌+型号锚点搜索 · 匹配度验证 · 韩国全网最低价</p>',
    unsafe_allow_html=True,
)

# ---------- 第一步：上传文件 ----------
with st.container():
    st.markdown(
        '<div class="px-marker"></div>'
        '<div class="step"><span class="step-num">1</span><span class="step-title">上传商品表格</span>'
        '<span class="step-sub">自动提取「品牌+型号+SKU」构建精准搜索词</span></div>',
        unsafe_allow_html=True,
    )
    st.markdown('<div class="upload-tip">🎮 支持价格链路表 / 报名价表 / 任意含商品名的表格</div>', unsafe_allow_html=True)
    uploaded_file = st.file_uploader(
        "支持 Excel (.xlsx) 或 CSV",
        type=["csv", "xlsx", "xls"],
        label_visibility="collapsed",
    )
    if uploaded_file is not None:
        st.markdown(f'<div class="file-chip">📄 {esc(uploaded_file.name)} ✓ 已放入</div>', unsafe_allow_html=True)


# ---------- 列识别 ----------
def detect_columns(df):
    """自动识别商品名列和SKU列"""
    name_col = None
    sku_col = None
    for c in df.columns:
        cs = str(c).replace("\n", " ").strip()
        if name_col is None and ("商品名" in cs or "상품명" in cs):
            name_col = c
        if sku_col is None and ("SKU" in cs.upper() or "옵션" in cs) and "ID" not in cs.upper():
            sku_col = c
    return name_col, sku_col


# ---------- 解析上传文件 ----------
sku_list = []         # 搜索词
query_detail = []     # (商品名, SKU, 搜索词, query_info_dict)
if uploaded_file is not None:
    fname = uploaded_file.name.lower()
    if fname.endswith(".csv"):
        content = uploaded_file.read().decode("utf-8-sig", errors="replace")
        reader = csv.reader(io.StringIO(content))
        for row in reader:
            if row and row[0].strip():
                q = row[0].strip()
                sku_list.append(q)
                query_detail.append((q[:30], "", q, {"brand": "", "model": "", "category": "", "sku_attrs": []}))
    else:
        # Excel: 尝试 header=0 和 header=1，选能识别到商品名列的那个
        df_up = pd.read_excel(uploaded_file, dtype=str)
        name_col, sku_col = detect_columns(df_up)
        if not name_col:
            # 尝试 header=1
            uploaded_file.seek(0)
            df_up2 = pd.read_excel(uploaded_file, header=1, dtype=str)
            name_col2, sku_col2 = detect_columns(df_up2)
            if name_col2:
                df_up = df_up2
                name_col, sku_col = name_col2, sku_col2

        if name_col:
            for _, row in df_up.iterrows():
                pname = str(row.get(name_col, "")).strip()
                psku = str(row.get(sku_col, "")).strip() if sku_col else ""
                if not pname or pname == "nan":
                    continue
                q = build_query(pname, psku)
                if q:
                    brand, model, cat = extract_brand_model(pname)
                    sku_attrs = []
                    if psku and psku not in ("nan", "", "单一sku", "단일sku", "全部sku"):
                        sku_attrs = [s.strip() for s in re.split(r'[/,|]', psku) if len(s.strip()) > 1][:2]
                    info = {"brand": brand, "model": model, "category": cat, "sku_attrs": sku_attrs}
                    sku_list.append(q)
                    query_detail.append((pname[:30], psku[:20], q, info))
        else:
            first_col = df_up.columns[0]
            for val in df_up[first_col].dropna().astype(str).str.strip():
                if val and val != "nan":
                    sku_list.append(val)
                    query_detail.append((val[:30], "", val, {"brand": "", "model": "", "category": "", "sku_attrs": []}))

    if sku_list:
        # ---------- 第二步：确认并查询 ----------
        with st.container():
            st.markdown(
                f'<div class="px-marker"></div>'
                f'<div class="step"><span class="step-num">2</span><span class="step-title">确认查询</span>'
                f'<span class="count-badge">{len(sku_list)} SKU</span></div>',
                unsafe_allow_html=True,
            )

            with st.expander("预览搜索词（前 10 个）"):
                for pname, psku, q, info in query_detail[:10]:
                    anchor = ""
                    if info["brand"] or info["model"]:
                        anchor = f" [锚:{info['brand']} {info['model']}]".strip()
                    if psku:
                        st.text(f"{pname} | {psku} → {q}{anchor}")
                    else:
                        st.text(f"{pname} → {q}{anchor}")
                if len(query_detail) > 10:
                    st.caption(f"... 还有 {len(query_detail) - 10} 个")

            api_ready = bool(st.session_state.get("naver_id")) and bool(st.session_state.get("naver_secret"))

            if not api_ready:
                st.warning("请先在左侧边栏填写 Naver API 密钥 👈")
            else:
                if st.button("▶ 开始查询", type="primary", width="stretch"):
                    st.session_state["results"] = []
                    st.session_state["query_done"] = False

                    status_text = st.empty()
                    progress_bar = st.progress(0)

                    results = []
                    ok_count = 0
                    high_match = 0
                    total = len(sku_list)

                    for idx, (sku, info) in enumerate(zip(sku_list, [d[3] for d in query_detail])):
                        status_text.markdown(
                            f'<div class="xp-label"><span>QUERYING...</span><span>{idx + 1} / {total}</span></div>',
                            unsafe_allow_html=True,
                        )

                        r = get_lowest_price(sku, st.session_state["naver_id"], st.session_state["naver_secret"], query_info=info)

                        if r["status"] == "ok":
                            ok_count += 1
                            if r["match_pct"] >= 70:
                                high_match += 1
                            results.append([sku, r["match_label"], r["mall"], r["price"], r["shipping"], r["title"], r["reviews"], r["link"]])
                        else:
                            results.append([sku, r.get("note", "错误"), r.get("note", ""), "", "", "", "", ""])

                        progress_bar.progress((idx + 1) / total)
                        time.sleep(0.25)

                    status_text.markdown("")
                    st.session_state["results"] = results
                    st.session_state["ok_count"] = ok_count
                    st.session_state["high_match"] = high_match
                    st.session_state["query_done"] = True

        # ---------- 第三步：结果展示与下载 ----------
        if st.session_state.get("query_done") and st.session_state.get("results"):
            results = st.session_state["results"]
            ok_count = st.session_state.get("ok_count", 0)
            high_match = st.session_state.get("high_match", 0)

            with st.container():
                st.markdown(
                    '<div class="px-marker"></div>'
                    '<div class="step"><span class="step-num">3</span><span class="step-title">查询结果</span></div>',
                    unsafe_allow_html=True,
                )

                # 统计
                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    st.markdown(
                        f'<div class="stat"><div class="stat-icon">🎯</div>'
                        f'<div class="stat-num">{len(results)}</div><div class="stat-label">总查询</div></div>',
                        unsafe_allow_html=True,
                    )
                with col2:
                    st.markdown(
                        f'<div class="stat stat-green"><div class="stat-icon">⭐</div>'
                        f'<div class="stat-num">{ok_count}</div><div class="stat-label">有结果</div></div>',
                        unsafe_allow_html=True,
                    )
                with col3:
                    st.markdown(
                        f'<div class="stat stat-green"><div class="stat-icon">✅</div>'
                        f'<div class="stat-num">{high_match}</div><div class="stat-label">高匹配</div></div>',
                        unsafe_allow_html=True,
                    )
                with col4:
                    st.markdown(
                        f'<div class="stat stat-coral"><div class="stat-icon">🔧</div>'
                        f'<div class="stat-num">{len(results) - ok_count}</div><div class="stat-label">需人工</div></div>',
                        unsafe_allow_html=True,
                    )

                # 结果表格
                display_rows = results[:100]
                thead = ('<tr><th>搜索词</th><th>匹配度</th><th>最低价商城</th><th>最低价(KRW)</th>'
                         '<th>≈USD</th><th>物流</th><th>商品原标题</th><th>评价</th></tr>')
                tbody = ""
                for r in display_rows:
                    sku, match_label, mall, price, ship, title, rev = r[0], r[1], r[2], r[3], r[4], r[5], r[6]
                    if str(price) == "":
                        tbody += f'<tr><td>{esc(sku)}</td><td colspan="7"><span class="px-err">{esc(mall)}</span></td></tr>'
                    else:
                        usd = int(price) / exchange_rate
                        # 匹配度颜色
                        if "高匹配" in str(match_label):
                            mcls = "match-high"
                        elif "需确认" in str(match_label):
                            mcls = "match-mid"
                        else:
                            mcls = "match-low"
                        tbody += (f'<tr><td>{esc(sku)}</td>'
                                  f'<td><span class="{mcls}">{esc(match_label)}</span></td>'
                                  f'<td>{esc(mall)}</td>'
                                  f'<td><span class="px-price">{int(price):,}</span></td>'
                                  f'<td><span class="px-price">${usd:.2f}</span></td>'
                                  f'<td><span class="px-tag">{esc(ship)}</span></td>'
                                  f'<td>{esc(title)}</td><td>{esc(rev)}</td></tr>')
                st.markdown(
                    f'<div class="px-table-wrap"><table class="px-table">'
                    f'<thead>{thead}</thead><tbody>{tbody}</tbody></table></div>',
                    unsafe_allow_html=True,
                )
                if len(results) > 100:
                    st.markdown('<div class="px-note">页面仅展示前 100 条，完整数据请下载</div>', unsafe_allow_html=True)

                # 下载 Excel
                xlsx_bytes = make_result_excel(results, rate=exchange_rate)
                st.download_button(
                    label="⬇ 下载结果 Excel（带匹配度颜色）",
                    data=xlsx_bytes,
                    file_name="Naver最低价_结果.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    width="stretch",
                )
