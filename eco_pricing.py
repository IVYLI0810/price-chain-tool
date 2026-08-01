# -*- coding: utf-8 -*-
"""二轮定价 · 生态表批量定价 核心逻辑（独立可测试，验证后接入 app.py）

输入：行业表(报名价+券) / 空白生态表(待填P-AE) / 到手价表(红线，可选)
逻辑（用户最终确认）：
  - 百补固定 brand_rate(默认5%)：Q=P*rate, T=P-Q
  - code 统一"倒推红线"：code = ceil0.5(页面价T - 券W - 一轮到手价V)，<0 给 0
  - 20%/30% 是"超code"标记阈值（非上限）；25%/35% 是"超帽"标记阈值
  - 唯一硬约束：二轮到手价 AC <= 一轮到手价 V
  - 无红线新品：按 target_rate 倒推 code = floor0.5(target*(P-W)-Q)
匹配：双键(商品ID+承接SKU) -> 单键唯一价 -> 多价留空标黄+候选 -> 行业无价标"未报名"
"""
import re
import math
from io import BytesIO

import numpy as np
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment


# ----------------------------- 解析工具 -----------------------------
def parse_price(v):
    """把报名价解析成 float。支持 '594.15 USD' / '$668' / '195.2美金' / 数字。"""
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v) if pd.notna(v) else None
    s = str(v).strip().replace(",", "")
    if not s or s.lower() in ("nan", "none", "未报名"):
        return None
    m = re.search(r"\d+\.?\d*", s)
    if not m:
        return None
    try:
        return float(m.group())
    except ValueError:
        return None


def parse_coupon(v):
    """把店铺券解析成 float（美金）。优先级见 MEMORY 券解析规则。"""
    if v is None:
        return 0.0
    if isinstance(v, (int, float)):
        return float(v) if pd.notna(v) else 0.0
    s = str(v).strip().replace(",", "")
    if not s or s.lower() in ("nan", "none", "无", "-"):
        return 0.0
    # 纯数字
    if re.fullmatch(r"\$?\d+\.?\d*", s):
        return float(re.search(r"\d+\.?\d*", s).group())
    # $X满减 / $X店铺code
    m = re.search(r"\$\s*(\d+\.?\d*)\s*(满减|店铺code|店铺码)", s)
    if m:
        return float(m.group(1))
    # Code-$A-B / 满立减-$A-B / 满$A-B -> 取 B（减免额）
    m = re.search(r"(Code|满立减|满)\s*\$?\s*(\d+\.?\d*)\s*[-减]\s*(\d+\.?\d*)", s, re.IGNORECASE)
    if m:
        return float(m.group(3))
    # X美金 / XUSD
    m = re.search(r"(\d+\.?\d*)\s*(美金|USD|美元)", s, re.IGNORECASE)
    if m:
        return float(m.group(1))
    # 兜底：第一个数字
    m = re.search(r"\d+\.?\d*", s)
    return float(m.group()) if m else 0.0


def ceil05(x):
    return math.ceil(x * 2) / 2.0


def floor05(x):
    return math.floor(x * 2) / 2.0


def norm_id(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    s = str(v).strip()
    if s.lower() in ("nan", "none"):
        return ""
    if s.endswith(".0"):
        s = s[:-2]
    return s


# ----------------------------- 倾斜钩子（留口子）-----------------------------
def determine_code_rate(商品, 参数, 表现数据=None):
    """过往表现驱动 code 档位倾斜的钩子。当前未启用，返回 None 表示走统一倒推红线逻辑。

    未来：根据 表现数据（如过往GMV/核销率）返回一个建议 code 率或档位偏移。
    """
    return None


# ----------------------------- 列定位 -----------------------------
def _find_col(df, subs, exclude=None, known_idx=None):
    """按子串找列；exclude 用于排除误匹配；找不到回退 known_idx。"""
    for i, c in enumerate(df.columns):
        cs = str(c).replace("\n", " ").strip()
        if exclude and any(e in cs for e in exclude):
            continue
        if any(sub in cs for sub in subs):
            return i
    if known_idx is not None and known_idx < len(df.columns):
        return known_idx
    return None


def build_hy_lookups(df_hy):
    """从行业表构建：双键价/券、单键唯一价、多价候选。"""
    id_c = _find_col(df_hy, ["承接ID", "연계ID"], known_idx=30)
    # 承接SKU：承接ID 之后第一个含 SKU 的列
    sku_c = None
    if id_c is not None:
        for i in range(id_c + 1, len(df_hy.columns)):
            if "SKU" in str(df_hy.columns[i]).upper():
                sku_c = i
                break
    if sku_c is None:
        sku_c = _find_col(df_hy, ["承接SKU"], known_idx=31)
    price_c = _find_col(df_hy, ["报名价"], exclude=["截图", "询价"], known_idx=35)
    coupon_c = _find_col(df_hy, ["券金额"], known_idx=39)

    dual_price, dual_coupon, single_price, multi_price = {}, {}, {}, {}
    coupon_single = {}
    by_id_prices = {}

    for _, r in df_hy.iterrows():
        pid = norm_id(r.iloc[id_c]) if id_c is not None else ""
        if not pid:
            continue
        psku = norm_id(r.iloc[sku_c]) if sku_c is not None else ""
        price = parse_price(r.iloc[price_c]) if price_c is not None else None
        coupon = parse_coupon(r.iloc[coupon_c]) if coupon_c is not None else 0.0
        if price is None:
            continue
        dual_price[(pid, psku)] = price
        dual_coupon[(pid, psku)] = coupon
        by_id_prices.setdefault(pid, []).append((price, coupon))

    for pid, lst in by_id_prices.items():
        prices = sorted({round(p, 4) for p, _ in lst})
        if len(prices) == 1:
            single_price[pid] = lst[0][0]
            coupon_single[pid] = lst[0][1]
        else:
            multi_price[pid] = prices

    return {
        "dual_price": dual_price,
        "dual_coupon": dual_coupon,
        "single_price": single_price,
        "coupon_single": coupon_single,
        "multi_price": multi_price,
    }


def build_redline_lookups(df_dj):
    """从到手价表构建红线（一轮到手价）：双键 + 单键。红线=最终价格，回退第一轮到手价。"""
    if df_dj is None:
        return {"dual": {}, "single": {}, "coupon_dual": {}, "coupon_single": {}}
    id_c = _find_col(df_dj, ["商品ID"], known_idx=9)
    sku_c = _find_col(df_dj, ["承接SKU"], known_idx=11)
    fp_c = _find_col(df_dj, ["最终价格", "최종할인가"], known_idx=28)
    r1_c = _find_col(df_dj, ["第一轮到手价"], known_idx=21)
    cp_c = _find_col(df_dj, ["店铺券", "满立减", "스토어 쿠폰"], exclude=["CODE"], known_idx=22)

    dual, single, coupon_dual, coupon_single = {}, {}, {}, {}
    for _, r in df_dj.iterrows():
        pid = norm_id(r.iloc[id_c]) if id_c is not None else ""
        if not pid:
            continue
        psku = norm_id(r.iloc[sku_c]) if sku_c is not None else ""
        rl = parse_price(r.iloc[fp_c]) if fp_c is not None else None
        if rl is None and r1_c is not None:
            rl = parse_price(r.iloc[r1_c])
        cp = parse_coupon(r.iloc[cp_c]) if cp_c is not None else 0.0
        if rl is None:
            continue
        # 双键取最低（最保守红线）
        k = (pid, psku)
        if k not in dual or rl < dual[k]:
            dual[k] = rl
            coupon_dual[k] = cp
        if pid not in single or rl < single[pid]:
            single[pid] = rl
            coupon_single[pid] = cp
    return {"dual": dual, "single": single, "coupon_dual": coupon_dual, "coupon_single": coupon_single}


# ----------------------------- 主流程 -----------------------------
# 生态表 openpyxl 列号（1-indexed）
ECO_IN = {"供给类型": 9, "商品ID": 10, "商品名": 11, "承接SKU": 12, "SKU": 13, "数量": 14}
ECO_OUT = {
    "P": 16, "Q": 17, "R": 18, "S": 19, "T": 20,
    "V": 22, "W": 23, "Z": 26, "AA": 27, "AB": 28,
    "AC": 29, "AD": 30, "AE": 31,
}

# 颜色
FILL_RED = PatternFill("solid", fgColor="FFFFE3E3")      # 超价
FILL_ORANGE = PatternFill("solid", fgColor="FFFFE8CC")    # 超code/超帽
FILL_YELLOW = PatternFill("solid", fgColor="FFFFF7CC")    # 待人工/未报名
FILL_GRAY = PatternFill("solid", fgColor="FFF2F2F2")      # 新品无红线


def is_full_supply(s):
    s = str(s)
    return ("全托" in s) or ("海托" in s)


def run_eco_pricing(hy_bytes, eco_bytes, dj_bytes=None, params=None, perf_df=None):
    params = params or {}
    brand_rate = params.get("brand_rate", 0.05)
    pop_code_flag = params.get("pop_code_flag", 0.20)
    full_code_flag = params.get("full_code_flag", 0.30)
    pop_cap = params.get("pop_cap", 0.25)
    full_cap = params.get("full_cap", 0.35)
    target_rate = params.get("target_rate", 0.24)

    df_hy, _ = _read(hy_bytes)
    df_dj = None
    if dj_bytes is not None:
        df_dj, _ = _read(dj_bytes)

    hy = build_hy_lookups(df_hy)
    rl = build_redline_lookups(df_dj)

    wb = load_workbook(BytesIO(eco_bytes))
    ws = wb.active

    # 找表头行（含"商品报名原价"或"商品ID"）
    header_row = None
    for ridx in range(1, min(6, ws.max_row + 1)):
        rowtext = " ".join(str(ws.cell(ridx, c).value or "") for c in range(1, min(20, ws.max_column + 1)))
        if "商品ID" in rowtext or "报名原价" in rowtext:
            header_row = ridx
            break
    if header_row is None:
        header_row = 2
    data_start = header_row + 1

    preview_rows = []
    stats = {"filled": 0, "multi": 0, "unreg": 0, "over_code": 0, "over_cap": 0,
             "over_price": 0, "new": 0, "total": 0}

    for ridx in range(data_start, ws.max_row + 1):
        eid = norm_id(ws.cell(ridx, ECO_IN["商品ID"]).value)
        name = ws.cell(ridx, ECO_IN["商品名"]).value
        esku = norm_id(ws.cell(ridx, ECO_IN["承接SKU"]).value)
        supply = ws.cell(ridx, ECO_IN["供给类型"]).value
        qty = parse_price(ws.cell(ridx, ECO_IN["数量"]).value) or 0.0

        # 跳过完全空行
        if not eid and not name:
            continue
        stats["total"] += 1

        full = is_full_supply(supply)
        code_flag_th = full_code_flag if full else pop_code_flag
        cap_th = full_cap if full else pop_cap

        # ---- 匹配报名价 ----
        note = ""
        status = ""
        P = hy["dual_price"].get((eid, esku))
        W = hy["dual_coupon"].get((eid, esku))
        if P is None and eid in hy["single_price"]:
            P = hy["single_price"][eid]
            W = hy["coupon_single"].get(eid, 0.0)
        if P is None and eid in hy["multi_price"]:
            cands = hy["multi_price"][eid]
            note = "多价待人工，候选价: " + " / ".join(f"${c:g}" for c in cands)
            status = "🟡 多价待人工"
            stats["multi"] += 1
            _fill_row(ws, ridx, None, fill=FILL_YELLOW, note=note)
            preview_rows.append(_preview(eid, name, supply, None, None, None, None, None, None, status, note))
            continue
        if P is None:
            status = "🟡 未报名"
            note = "行业表无此商品报名价"
            stats["unreg"] += 1
            _fill_row(ws, ridx, None, fill=FILL_YELLOW, note=note)
            preview_rows.append(_preview(eid, name, supply, None, None, None, None, None, None, status, note))
            continue

        W = W or 0.0
        # 红线
        V = rl["dual"].get((eid, esku))
        if V is None:
            V = rl["single"].get(eid)
        # 券兜底：行业表无券则用到手价表券
        if (W is None or W == 0.0):
            wfb = rl["coupon_dual"].get((eid, esku)) or rl["coupon_single"].get(eid) or 0.0
            W = wfb

        # ---- 计算 ----
        Q = P * brand_rate
        T = P - Q
        tilt = determine_code_rate({"id": eid, "sku": esku, "supply": supply},
                                   {"target_rate": target_rate}, perf_df)
        if V is not None:
            need = T - W - V
            code = ceil05(need) if need > 0 else 0.0
            is_new = False
        else:
            code = max(floor05(target_rate * (P - W) - Q), 0.0)
            is_new = True

        if tilt is not None:
            code = tilt  # 预留：倾斜钩子若返回值则覆盖

        S = (Q + code) / (P - W) if (P - W) > 0 else 0.0
        AB = code / T if T > 0 else 0.0
        AC = T - W - code
        AA = qty * code
        AD = qty * T
        AE = AD / AA if AA > 0 else None

        # ---- 标记 ----
        flags = []
        over_code = AB > code_flag_th + 1e-9
        over_cap = S > cap_th + 1e-9
        over_price = (V is not None) and (AC > V + 1e-6)
        if over_code:
            flags.append("超code")
        if over_cap:
            flags.append("超帽")
        if over_price:
            flags.append("超价")

        if is_new:
            status = "⚪ 新品无红线"
            stats["new"] += 1
            fill = FILL_GRAY
        elif over_price:
            status = "🔴 超价放行"
            stats["over_price"] += 1
            fill = FILL_RED
        elif flags:
            status = "⚠️ " + "+".join(flags)
            fill = FILL_ORANGE
            if over_code:
                stats["over_code"] += 1
            if over_cap:
                stats["over_cap"] += 1
        else:
            status = "✅ 压价成功"
            fill = None
            stats["filled"] += 1

        vals = {"P": P, "Q": Q, "R": brand_rate, "S": S, "T": T,
                "V": V, "W": W, "Z": code, "AA": AA, "AB": AB,
                "AC": AC, "AD": AD, "AE": AE}
        _fill_row(ws, ridx, vals, fill=fill, note=note)
        preview_rows.append(_preview(eid, name, supply, P, V, T, W, code, AC, status, note, S, AB))

    # 备注列表头
    note_col = ws.max_column + 1 if "匹配备注" not in [str(ws.cell(header_row, c).value or "") for c in range(1, ws.max_column + 1)] else None
    # 简化：固定写到第 61 列（BI）若不存在备注列
    _ensure_note_header(ws, header_row)

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)

    preview = pd.DataFrame(preview_rows)
    return buf, preview, stats


def _read(b):
    df_raw = pd.read_excel(BytesIO(b), header=None)
    header_row = 1
    for i in range(min(5, len(df_raw))):
        rt = " ".join(str(x) for x in df_raw.iloc[i].tolist() if pd.notna(x))
        if "商品ID" in rt or "报名原价" in rt or "报名价" in rt or "承接ID" in rt:
            header_row = i
            break
    df = pd.read_excel(BytesIO(b), header=header_row)
    return df, header_row


def _preview(eid, name, supply, P, V, T, W, code, AC, status, note="", S=None, AB=None):
    return {
        "商品ID": eid,
        "商品名": (str(name)[:26] if name is not None else ""),
        "供给类型": (str(supply) if supply is not None else ""),
        "报名价P": round(P, 2) if P is not None else None,
        "红线V": round(V, 2) if V is not None else None,
        "页面价T": round(T, 2) if T is not None else None,
        "券W": round(W, 2) if W is not None else None,
        "code": round(code, 1) if code is not None else None,
        "到手价AC": round(AC, 2) if AC is not None else None,
        "叠加率S": S,
        "code率AB": AB,
        "状态": status,
        "备注": note,
    }


def _fill_row(ws, ridx, vals, fill=None, note=""):
    """把计算值写入 P-AE；vals=None 表示待人工行（只上色+备注）。"""
    money_fmt = "0.00"
    rate_fmt = "0.00%"
    code_fmt = "0.0"
    if vals:
        ws.cell(ridx, ECO_OUT["P"], round(vals["P"], 2)).number_format = money_fmt
        ws.cell(ridx, ECO_OUT["Q"], round(vals["Q"], 2)).number_format = money_fmt
        ws.cell(ridx, ECO_OUT["R"], vals["R"]).number_format = rate_fmt
        ws.cell(ridx, ECO_OUT["S"], vals["S"]).number_format = rate_fmt
        ws.cell(ridx, ECO_OUT["T"], round(vals["T"], 2)).number_format = money_fmt
        if vals["V"] is not None:
            ws.cell(ridx, ECO_OUT["V"], round(vals["V"], 2)).number_format = money_fmt
        ws.cell(ridx, ECO_OUT["W"], round(vals["W"], 2)).number_format = money_fmt
        ws.cell(ridx, ECO_OUT["Z"], vals["Z"]).number_format = code_fmt
        ws.cell(ridx, ECO_OUT["AA"], round(vals["AA"], 2)).number_format = money_fmt
        ws.cell(ridx, ECO_OUT["AB"], vals["AB"]).number_format = rate_fmt
        ws.cell(ridx, ECO_OUT["AC"], round(vals["AC"], 2)).number_format = money_fmt
        ws.cell(ridx, ECO_OUT["AD"], round(vals["AD"], 2)).number_format = money_fmt
        if vals["AE"] is not None:
            ws.cell(ridx, ECO_OUT["AE"], round(vals["AE"], 2)).number_format = "0.00"
    # 上色（整行 P-AC 区间）
    if fill is not None:
        for c in range(ECO_OUT["P"], ECO_OUT["AC"] + 1):
            ws.cell(ridx, c).fill = fill
    # 备注
    nc = _note_col(ws)
    if note:
        ws.cell(ridx, nc, note)


_NOTE_COL_IDX = 61  # BI 列


def _note_col(ws):
    return _NOTE_COL_IDX


def _ensure_note_header(ws, header_row):
    ws.cell(header_row, _NOTE_COL_IDX, "匹配备注")
