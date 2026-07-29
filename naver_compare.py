"""
Naver 比价 + AE报名价验证 脚本
网红团购一站式平台 - 比价模块

逻辑：
1. 读取价格链路表（商品名韩文 + Brand + SKU옵션）
2. 拼搜索关键词 → 调 Naver Shopping API
3. 判断最低价来源：
   - 若最低是AE链接 → 抓AE价/截图, AE价<报名价则标红, 再排除AE搜站外最低+截图
   - 若最低是站外 → 直接抓站外链接/价/截图
4. 输出比价报告 + 截图文件夹
"""

import pandas as pd
import numpy as np
import requests
import time
import re
import os
import json
from urllib.parse import quote
from datetime import datetime
from pathlib import Path

# ─────────────────────────────────────────────
# 配置（使用前填写）
# ─────────────────────────────────────────────
NAVER_CLIENT_ID = ""       # ← 填你的 Naver Client ID
NAVER_CLIENT_SECRET = ""   # ← 填你的 Naver Client Secret
EXCHANGE_RATE = 1550.0     # KRW → USD 汇率

# 输出目录
OUTPUT_DIR = Path("./比价结果")
SCREENSHOT_DIR = OUTPUT_DIR / "截图"

# 屏蔽词（二手/翻新）
BLOCK_WORDS = ["중고", "리퍼", "박스훼손", "렌탈", "중고나라", "당근", "번개장터"]

# AE 域名识别
AE_DOMAINS = ["aliexpress.com", "aliexpress.us", "aliexpress.ru", "aliexpress.io"]


# ─────────────────────────────────────────────
# 工具函数
# ─────────────────────────────────────────────
def clean_html(raw_html):
    """去除HTML标签"""
    return re.sub(r"<.*?>", "", raw_html)


def is_ae_link(url):
    """判断链接是否为速卖通"""
    url_lower = str(url).lower()
    return any(domain in url_lower for domain in AE_DOMAINS)


def build_search_query(brand, product_name, sku_option):
    """
    拼接 Naver 搜索关键词
    规则：Brand + 商品名核心词（前4个词）+ SKU选项（非"单一sku"时）
    """
    parts = []

    # Brand
    if pd.notna(brand) and str(brand).strip():
        parts.append(str(brand).strip())

    # 商品名：取核心词（去掉太通用的词）
    if pd.notna(product_name):
        name = str(product_name).strip()
        # 去掉通用修饰词
        generic_words = ["야외", "캠핑", "피크닉", "여행", "휴대용", "다기능",
                         "스테인리스", "스틸", "대용량", "경량", "방수", "미니"]
        words = name.split()
        core_words = [w for w in words if w not in generic_words]
        # 取前4个核心词
        parts.extend(core_words[:4])

    # SKU选项
    if pd.notna(sku_option):
        sku = str(sku_option).strip()
        if sku and sku != "单一sku" and sku != "단일sku":
            parts.append(sku)

    query = " ".join(parts)
    # 限制长度（Naver 搜索太长会没结果）
    if len(query) > 60:
        query = query[:60]
    return query


def search_naver(query, client_id, client_secret, exclude_ae=False):
    """
    调用 Naver Shopping API
    返回按价格升序排列的结果列表
    """
    encoded_query = quote(query)
    url = f"https://openapi.naver.com/v1/search/shop.json?query={encoded_query}&display=20&sort=asc"
    headers = {
        "X-Naver-Client-Id": client_id,
        "X-Naver-Client-Secret": client_secret,
    }

    try:
        resp = requests.get(url, headers=headers, timeout=10)
        if resp.status_code != 200:
            return [], f"API错误({resp.status_code})"

        items = resp.json().get("items", [])
        results = []
        for item in items:
            title = clean_html(item.get("title", ""))
            link = item.get("link", "")
            price = int(item.get("lprice", 0))
            mall = item.get("mallName", "")

            # 过滤二手
            if any(w in title.lower() for w in BLOCK_WORDS):
                continue

            # 是否排除AE
            if exclude_ae and is_ae_link(link):
                continue

            results.append({
                "title": title,
                "link": link,
                "price_krw": price,
                "price_usd": round(price / EXCHANGE_RATE, 2),
                "mall": mall,
                "is_ae": is_ae_link(link),
            })

        # 按价格排序
        results.sort(key=lambda x: x["price_krw"])
        return results, None

    except requests.exceptions.Timeout:
        return [], "查询超时"
    except Exception as e:
        return [], f"出错: {str(e)[:50]}"


# ─────────────────────────────────────────────
# 主比价逻辑
# ─────────────────────────────────────────────
def compare_single_product(row, client_id, client_secret):
    """
    对单个商品执行完整比价流程
    返回结果字典
    """
    brand = row.get("brand", "")
    product_name = row.get("product_name", "")
    sku_option = row.get("sku_option", "")
    registration_price = row.get("registration_price", 0)  # 报名价(USD)
    ae_final_price = row.get("ae_final_price", 0)  # AE最终价(USD)
    product_id = row.get("product_id", "")

    result = {
        "product_id": product_id,
        "product_name": str(product_name)[:30],
        "registration_price_usd": registration_price,
        "ae_final_price_usd": ae_final_price,
        "search_query": "",
        "naver_lowest_is_ae": False,
        "naver_ae_price_krw": None,
        "naver_ae_price_usd": None,
        "naver_ae_link": "",
        "naver_ae_title": "",
        "reg_price_flag": "",  # 报名价是否虚高
        "external_lowest_price_krw": None,
        "external_lowest_price_usd": None,
        "external_link": "",
        "external_mall": "",
        "external_title": "",
        "comparison_result": "",  # AE价低/站外价低/无法比较
        "screenshot_ae": "",
        "screenshot_external": "",
        "error": "",
    }

    # 1. 拼搜索关键词
    query = build_search_query(brand, product_name, sku_option)
    result["search_query"] = query

    if not query.strip():
        result["error"] = "无法生成搜索关键词"
        return result

    # 2. 第一轮搜索（不排除AE）
    all_results, err = search_naver(query, client_id, client_secret, exclude_ae=False)
    if err:
        result["error"] = err
        return result

    if not all_results:
        result["error"] = "Naver无匹配结果"
        return result

    # 3. 判断最低价来源
    lowest = all_results[0]

    if lowest["is_ae"]:
        # ── 情况一：最低价是速卖通 ──
        result["naver_lowest_is_ae"] = True
        result["naver_ae_price_krw"] = lowest["price_krw"]
        result["naver_ae_price_usd"] = lowest["price_usd"]
        result["naver_ae_link"] = lowest["link"]
        result["naver_ae_title"] = lowest["title"]

        # AE价 vs 报名价
        if registration_price and lowest["price_usd"] < registration_price:
            result["reg_price_flag"] = "⚠️ 报名价虚高（AE实际价更低）"

        # 4. 第二轮搜索（排除AE，找站外最低）
        time.sleep(0.3)
        ext_results, ext_err = search_naver(query, client_id, client_secret, exclude_ae=True)
        if ext_results:
            ext_lowest = ext_results[0]
            result["external_lowest_price_krw"] = ext_lowest["price_krw"]
            result["external_lowest_price_usd"] = ext_lowest["price_usd"]
            result["external_link"] = ext_lowest["link"]
            result["external_mall"] = ext_lowest["mall"]
            result["external_title"] = ext_lowest["title"]

        # 对比结论
        if result["external_lowest_price_usd"] and ae_final_price:
            if ae_final_price <= result["external_lowest_price_usd"]:
                result["comparison_result"] = "✅ AE价低"
            else:
                result["comparison_result"] = "❌ 站外价低（AE无竞争力）"

    else:
        # ── 情况二：最低价是站外 ──
        result["naver_lowest_is_ae"] = False
        result["external_lowest_price_krw"] = lowest["price_krw"]
        result["external_lowest_price_usd"] = lowest["price_usd"]
        result["external_link"] = lowest["link"]
        result["external_mall"] = lowest["mall"]
        result["external_title"] = lowest["title"]

        # 对比结论
        if ae_final_price:
            if ae_final_price <= lowest["price_usd"]:
                result["comparison_result"] = "✅ AE价低"
            else:
                result["comparison_result"] = "❌ 站外价低（AE无竞争力）"

    return result


# ─────────────────────────────────────────────
# 批量执行
# ─────────────────────────────────────────────
def run_batch_comparison(excel_path, client_id, client_secret):
    """读取Excel，批量比价"""

    # 读取表格
    df_raw = pd.read_excel(excel_path, header=None)

    # 找表头行
    header_row = 1
    for i in range(min(5, len(df_raw))):
        row_text = " ".join([str(x) for x in df_raw.iloc[i].tolist() if pd.notna(x)])
        if "商品名" in row_text or "상품명" in row_text:
            header_row = i
            break

    df = pd.read_excel(excel_path, header=header_row)

    # 识别列
    def find_col(keywords):
        for c in df.columns:
            cs = str(c).replace("\n", " ")
            if all(k in cs for k in keywords):
                return c
        for c in df.columns:
            cs = str(c).replace("\n", " ")
            if any(k in cs for k in keywords):
                return c
        return None

    col_name = find_col(["商品名"]) or find_col(["상품명"])
    col_brand = find_col(["Brand"])
    # SKU文字列（옵션）：排除含"ID"的列
    col_sku_option = None
    for c in df.columns:
        cs = str(c).replace("\n", " ")
        if ("SKU" in cs or "옵션" in cs) and "ID" not in cs:
            col_sku_option = c
            break
    col_pid = find_col(["商品ID"]) or find_col(["상품ID"])
    col_price = find_col(["报名原价"]) or find_col(["노미네이션"]) or find_col(["询价价格"])

    print(f"识别列: 商品名={col_name}, Brand={col_brand}, SKU={col_sku_option}, 商品ID={col_pid}, 价格={col_price}")
    print(f"共 {len(df)} 行数据")

    # 构建数据
    products = []
    for i, row in df.iterrows():
        products.append({
            "index": i,
            "brand": row.get(col_brand, "") if col_brand else "",
            "product_name": row.get(col_name, "") if col_name else "",
            "sku_option": row.get(col_sku_option, "") if col_sku_option else "",
            "product_id": row.get(col_pid, "") if col_pid else "",
            "registration_price": pd.to_numeric(row.get(col_price, 0), errors="coerce") or 0,
            "ae_final_price": 0,  # 后续可从价格链路计算结果中获取
        })

    # 批量查询
    results = []
    total = len(products)
    for idx, prod in enumerate(products):
        print(f"[{idx+1}/{total}] 搜索: {build_search_query(prod['brand'], prod['product_name'], prod['sku_option'])[:40]}...")
        result = compare_single_product(prod, client_id, client_secret)
        results.append(result)
        time.sleep(0.3)  # 控制频率

    # 输出结果
    OUTPUT_DIR.mkdir(exist_ok=True)
    SCREENSHOT_DIR.mkdir(exist_ok=True)

    df_results = pd.DataFrame(results)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M")
    output_file = OUTPUT_DIR / f"比价报告_{timestamp}.xlsx"
    df_results.to_excel(output_file, index=False, engine="openpyxl")

    # 统计
    total_ok = sum(1 for r in results if "AE价低" in r.get("comparison_result", ""))
    total_bad = sum(1 for r in results if "站外价低" in r.get("comparison_result", ""))
    total_flag = sum(1 for r in results if r.get("reg_price_flag"))
    total_err = sum(1 for r in results if r.get("error"))

    print(f"\n{'='*50}")
    print(f"比价完成！")
    print(f"  总计: {total}")
    print(f"  ✅ AE价低: {total_ok}")
    print(f"  ❌ 站外价低: {total_bad}")
    print(f"  ⚠️ 报名价虚高: {total_flag}")
    print(f"  🔧 查询失败: {total_err}")
    print(f"  报告已保存: {output_file}")
    print(f"{'='*50}")

    return df_results


# ─────────────────────────────────────────────
# 入口
# ─────────────────────────────────────────────
if __name__ == "__main__":
    import sys

    if not NAVER_CLIENT_ID or not NAVER_CLIENT_SECRET:
        print("❌ 请先在脚本顶部填写 NAVER_CLIENT_ID 和 NAVER_CLIENT_SECRET")
        print("   申请地址: https://developers.naver.com")
        sys.exit(1)

    excel_file = sys.argv[1] if len(sys.argv) > 1 else "/Users/iivyli/Downloads/完整价格链路-6.xlsx"

    if not os.path.exists(excel_file):
        print(f"❌ 文件不存在: {excel_file}")
        sys.exit(1)

    run_batch_comparison(excel_file, NAVER_CLIENT_ID, NAVER_CLIENT_SECRET)
