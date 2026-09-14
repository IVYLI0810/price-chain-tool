"""
SKU 历史价 · 查询层
──────────────────────────────────────────────
把 sku_detail 底表加载成"按标准商品名聚合"的商品对象，
对输入行 (商品ID / 商品名称 / sku选项) 做：
  1) 候选粗筛（ID命中 + 名称模糊Top）
  2) 三维打分选出最匹配的商品
  3) 在该商品内，按期次挑出"型号最贴合输入SKU"的那一行 → 取到手价
  4) 输出 4 期价 + 匹配度 + 等级 + 人话依据
"""

import sqlite3
from functools import lru_cache
from pathlib import Path

import pandas as pd
from rapidfuzz import process, fuzz

import sku_matcher as M

_HERE = Path(__file__).parent
DB_PATH = next((str(p) for p in [_HERE / 'history_data.db', Path('history_data.db')] if p.exists()), None)
# 纯文本底表：GitHub 网页端不接受二进制 .db，改上传 sku_detail.csv（内容完全一致）。
CSV_PATH = next((str(p) for p in [_HERE / 'sku_detail.csv', Path('sku_detail.csv')] if p.exists()), None)

_NAME_CANDIDATES = 15  # 名称粗筛保留的候选数


def _read_detail_frame() -> pd.DataFrame:
    """优先读 sku_detail.csv（纯文本、网页可上传）；没有再退回 history_data.db。"""
    cols = ['std_name', 'product_id', 'period', 'sku_raw', 'price', 'raw_val']
    if CSV_PATH:
        df = pd.read_csv(CSV_PATH, dtype={'product_id': str, 'std_name': str,
                                          'period': str, 'sku_raw': str, 'raw_val': str})
        for c in cols:
            if c not in df.columns:
                df[c] = ''
        df = df[cols]
    elif DB_PATH:
        conn = sqlite3.connect(DB_PATH)
        df = pd.read_sql_query(
            'SELECT std_name,product_id,period,sku_raw,price,raw_val FROM sku_detail', conn)
        conn.close()
    else:
        raise FileNotFoundError('找不到底表：请确认 sku_detail.csv 或 history_data.db 已上传到程序目录')
    df['price'] = pd.to_numeric(df['price'], errors='coerce')
    df['product_id'] = df['product_id'].apply(M.norm_id)
    df['sku_raw'] = df['sku_raw'].fillna('')
    return df


def load_detail(db_path=None) -> pd.DataFrame:
    """加载 SKU 底表。默认走 CSV；显式传入 db_path 时从该 SQLite 读取。"""
    if db_path:
        conn = sqlite3.connect(db_path)
        df = pd.read_sql_query(
            'SELECT std_name,product_id,period,sku_raw,price,raw_val FROM sku_detail', conn)
        conn.close()
        df['price'] = pd.to_numeric(df['price'], errors='coerce')
        df['product_id'] = df['product_id'].apply(M.norm_id)
        df['sku_raw'] = df['sku_raw'].fillna('')
    else:
        df = _read_detail_frame()
    # 预拆选项，避免重复计算
    df['options'] = df['sku_raw'].apply(M.split_sku_options)
    df['name_norm'] = df['std_name'].apply(M.norm_text)
    return df


def build_products(df: pd.DataFrame):
    """按标准商品名聚合成商品列表。"""
    products = []
    for std_name, g in df.groupby('std_name', sort=False):
        products.append({
            'std_name': std_name,
            'name_norm': M.norm_text(std_name),
            'ids': set(x for x in g['product_id'].tolist() if x),
            'rows': g,
        })
    return products


@lru_cache(maxsize=1)
def get_cached_products():
    """进程级缓存：底表只解析一次，全站所有会话共享（只读，不修改）。
    避免每次点「开始匹配」都重读 CSV + 重拆 SKU 选项，多人并发时显著降 CPU。"""
    detail = load_detail()
    products = build_products(detail)
    name_pool = [p['name_norm'] for p in products]
    return detail, products, name_pool


def _id_score(input_id, ids) -> float:
    if not input_id:
        return None
    if not ids:
        return 0.0
    return 100.0 if str(input_id).strip() in ids else 0.0


def _row_dim_verdict(in_dims, row_options, strict_mod=False):
    """输入维度 vs 底表某行各选项：取差异最小的一组 (hard, soft)。"""
    if in_dims is None or not row_options:
        return [], []
    best = None
    for co in row_options:
        hard, soft = M.dim_diffs(in_dims, M.extract_dims(co), strict_mod=strict_mod)
        key = (len(hard), len(soft))
        if best is None or key < best[0]:
            best = (key, hard, soft)
    return best[1], best[2]


def _ambiguous_dim(pr):
    """输入没写规格时：该期各行之间某影响单价的维度取值不一 → 无法代运营选择，判参照。"""
    opts_all = [o for opts in pr['options'] for o in (opts or [])]
    # mod 特殊：一方有 Pro、一方无，也算取值不一（None 会被通用循环漏掉）
    if 'mod' in M.AMBIGUOUS_DIMS and M.mod_varies(opts_all):
        return 'mod'
    for dim in M.AMBIGUOUS_DIMS:
        if dim == 'mod':
            continue
        vals = set()
        for opts in pr['options']:
            for o in opts:
                v = M.extract_dims(o).get(dim)
                if v:
                    vals.add(v)
        if len(vals) > 1:
            return dim
    return None


def _dim_varies(prod_rows, dim) -> bool:
    """该商品的选项里某个 NOMATCH 维度是否有>1种取值（"未标注"也算一种取值）。"""
    vals = set()
    for opts in prod_rows['options']:
        for o in (opts or []):
            v = M.extract_dims(o).get(dim)
            if dim == 'mod':
                vals.add(frozenset(v) if v else frozenset())
            else:
                vals.add(str(v).lower() if v is not None else '__none__')
    return len(vals) > 1


def _dim_present(prod_rows, dim, iv) -> bool:
    """该商品是否至少有一行选项的该维度值 == 输入值 iv。"""
    for opts in prod_rows['options']:
        for o in (opts or []):
            rv = M.extract_dims(o).get(dim)
            if dim == 'mod':
                if (frozenset(rv) if rv else frozenset()) == (frozenset(iv) if iv else frozenset()):
                    return True
            elif rv is not None and str(rv).lower() == str(iv).lower():
                return True
    return False


def _pick_period_price(prod_rows, input_sku):
    """
    在选中商品内，按期次挑价（规格维度分类）：
      - SOFT 差异（颜色/尺码/同底型号后缀）→ 照常取价，verdict=take
      - REF 差异（数量/容量/瓦数/电压）→ 同款不同规格，不取价，verdict=ref（给参照）
      - NOMATCH 差异（Pro/Plus修饰词、长度、底型号）→ 明显不是同商品，verdict=nomatch（不取价、不给参照）
      - 无输入SKU：取当期最低价；但若该期影响单价的维度取值不一 → verdict=ref
    返回 {period: {price, verdict, hard, soft, sku_raw, ref_row, sku_score}}
    """
    out = {}
    in_opts = M.split_sku_options(input_sku)
    have_sku = bool(in_opts)
    in_dims = M.extract_dims(' '.join(in_opts)) if have_sku else None
    all_opts = [o for opts in prod_rows['options'] for o in (opts or [])]
    strict_mod = M.mod_varies(all_opts)   # 该商品是否真的区分 Pro/非Pro
    # 产品级闸门：输入指定的 Pro/长度/底型号，商品确实区分该维度、却没有任何一行匹配
    # → 明显不是同款，整行判型号不符（避免某期空白选项误取价）
    gate_dim = None
    if have_sku:
        for dim in M.NOMATCH_DIMS:
            iv = in_dims.get(dim)
            if not iv:
                continue
            if _dim_varies(prod_rows, dim) and not _dim_present(prod_rows, dim, iv):
                gate_dim = dim
                break
    for period in M.PERIODS:
        pr = prod_rows[prod_rows['period'] == period]
        if pr.empty:
            continue
        pr2 = pr.dropna(subset=['price'])
        low_row = (pr2 if not pr2.empty else pr).sort_values('price').iloc[0]
        if gate_dim:
            out[period] = {'price': None, 'verdict': 'nomatch', 'hard': [gate_dim], 'soft': [],
                           'sku_raw': low_row['sku_raw'], 'ref_row': None, 'sku_score': None}
            continue
        if not have_sku:
            amb = _ambiguous_dim(pr)
            if amb:
                out[period] = {'price': None, 'verdict': 'ref', 'hard': [amb], 'soft': [],
                               'sku_raw': low_row['sku_raw'], 'ref_row': low_row, 'sku_score': None}
            else:
                out[period] = {'price': low_row['price'], 'verdict': 'take', 'hard': [], 'soft': [],
                               'sku_raw': low_row['sku_raw'], 'ref_row': None, 'sku_score': None}
            continue
        take_best, ref_best, nomatch_best = None, None, None
        for _, row in pr.iterrows():
            hard, soft = _row_dim_verdict(in_dims, row['options'], strict_mod=strict_mod)
            sim = M.sku_similarity(input_sku, row['options'])
            sc = sim['score'] if sim['score'] is not None else -1
            if not hard:
                if take_best is None or sc > take_best['sc']:
                    take_best = {'sc': sc, 'row': row, 'soft': soft}
            elif any(k in M.NOMATCH_DIMS for k in hard):
                # 含 Pro/长度/底型号 差异 → 明显不是同商品，不取价也不给参照
                if nomatch_best is None or sc > nomatch_best['sc']:
                    nomatch_best = {'sc': sc, 'row': row, 'hard': hard}
            else:
                # 仅 数量/容量/瓦数/电压 差异 → 同款不同规格，给参照
                if ref_best is None or sc > ref_best['sc']:
                    ref_best = {'sc': sc, 'row': row, 'hard': hard}
        if take_best:
            r = take_best['row']
            out[period] = {'price': r['price'], 'verdict': 'take', 'hard': [],
                           'soft': take_best['soft'], 'sku_raw': r['sku_raw'],
                           'ref_row': None, 'sku_score': take_best['sc']}
        elif ref_best:
            r = ref_best['row']
            out[period] = {'price': None, 'verdict': 'ref', 'hard': ref_best['hard'],
                           'soft': [], 'sku_raw': r['sku_raw'], 'ref_row': r,
                           'sku_score': ref_best['sc']}
        elif nomatch_best:
            r = nomatch_best['row']
            out[period] = {'price': None, 'verdict': 'nomatch', 'hard': nomatch_best['hard'],
                           'soft': [], 'sku_raw': r['sku_raw'], 'ref_row': None,
                           'sku_score': nomatch_best['sc']}
    return out


def match_one(input_id, input_name, input_sku, products, name_pool):
    """匹配单行，返回结果字典。"""
    input_id = M.norm_id(input_id)
    input_name = '' if input_name is None else str(input_name).strip()

    # ── 候选粗筛 ──
    cand_idx = set()
    if input_id:
        for i, p in enumerate(products):
            if input_id in p['ids']:
                cand_idx.add(i)
    if input_name and name_pool:
        hits = process.extract(M.norm_text(input_name), name_pool,
                               scorer=fuzz.token_set_ratio,
                               limit=_NAME_CANDIDATES)
        for _, score, idx in hits:
            if score >= 45:
                cand_idx.add(idx)
    if not cand_idx:
        return {'matched': False}

    # ── 选同款商品：只认【商品ID】+【商品名称】，SKU 不参与选品 ──
    #    名称是门槛：ID没命中时，商品名相似度必须≥NAME_GATE 才算同款；
    #    名称不够像 → 直接淘汰，SKU再一致也不采纳（避免本末倒置）。
    best = None
    for i in cand_idx:
        p = products[i]
        id_hit = bool(input_id) and input_id in p['ids']
        nmsc = M.name_similarity(input_name, p['std_name']) if input_name else 0.0
        # 门槛：ID命中 或 名称高相似，二者都不满足则跳过
        if not id_hit and nmsc < M.NAME_GATE:
            continue
        # 排序键：ID命中优先，其次名称相似度
        sel = (1 if id_hit else 0, nmsc)
        if best is None or sel > best['sel']:
            best = {'idx': i, 'product': p, 'sel': sel, 'id_hit': id_hit, 'name_score': nmsc}

    if best is None:
        return {'matched': False}

    p = best['product']
    id_hit = best['id_hit']

    # ── 按期取价（SOFT差异取价 / REF差异给参照 / NOMATCH差异判型号不符） ──
    period_price = _pick_period_price(p['rows'], input_sku)

    verdicts = [v.get('verdict') for v in period_price.values()]
    if any(v == 'take' for v in verdicts):
        conclusion = '取到价'
    elif any(v == 'ref' for v in verdicts):
        conclusion = '仅参照'
    elif any(v == 'nomatch' for v in verdicts):
        conclusion = '型号不符'
    else:
        conclusion = '无价'

    score = M.match_score(id_hit, best['name_score'])
    g = M.grade(id_hit, best['name_score'])

    # ── 依据（人话） ──
    reasons = []
    if id_hit:
        reasons.append('商品ID完全一致')
    elif input_id:
        reasons.append(f"商品ID未命中，靠商品名匹配到同款（底表用过ID：{'/'.join(sorted(p['ids'])) or '无'}）")
    else:
        reasons.append('无商品ID，靠商品名匹配')
    reasons.append(f"商品名相似 {best['name_score']:.0f}%")

    soft_dims, ref_dims, nomatch_dims, ref_parts = [], [], [], []
    for period in M.PERIODS:
        v = period_price.get(period)
        if not v:
            continue
        if v['verdict'] == 'take':
            for k in v['soft']:
                if M.DIM_LABEL[k] not in soft_dims:
                    soft_dims.append(M.DIM_LABEL[k])
        elif v['verdict'] == 'ref':
            for k in v['hard']:
                if M.DIM_LABEL[k] not in ref_dims:
                    ref_dims.append(M.DIM_LABEL[k])
            rp = v['ref_row']
            pv = rp['price'] if rp is not None else None
            ref_parts.append(f"{period}→底表选项『{v['sku_raw']}』"
                             + (f" ${float(pv):.2f}" if pv is not None and pd.notna(pv) else ''))
        else:  # nomatch：明显不是同商品，不取价、不给参照
            for k in v['hard']:
                if M.DIM_LABEL[k] not in nomatch_dims:
                    nomatch_dims.append(M.DIM_LABEL[k])
    if soft_dims:
        reasons.append(f"{'/'.join(soft_dims)}不同（不影响单价），照常取价")
    if ref_dims:
        reasons.append(f"⚠ {'/'.join(ref_dims)}不同（影响单价），不取价，见参照信息供运营判断")
    if nomatch_dims:
        reasons.append(f"⛔ {'/'.join(nomatch_dims)}不同——判定不是同款，不取价也不给参照")
    ref_info = '；'.join(ref_parts)

    return {
        'matched': g != '未匹配',
        'grade': g,
        'score': score,
        'conclusion': conclusion,
        'ref_info': ref_info,
        'std_name': p['std_name'],
        'matched_ids': '/'.join(sorted(p['ids'])),
        'period_price': period_price,
        'reason': '；'.join(reasons) if reasons else '—',
        'id_hit': id_hit, 'name_score': best['name_score'],
    }



def match_batch(df_input: pd.DataFrame, col_id, col_name, col_sku, db_path=None):
    """
    批量匹配。df_input 至少含 col_id/col_name/col_sku 三列（列名由调用方传入）。
    返回结果 DataFrame（原列 + 4期价 + 匹配度/等级/依据/命中信息）。
    """
    if db_path:
        detail = load_detail(db_path)
        products = build_products(detail)
        name_pool = [p['name_norm'] for p in products]
    else:
        detail, products, name_pool = get_cached_products()

    def _clean(v):
        if v is None:
            return ''
        s = str(v).strip()
        return '' if s.lower() in ('nan', 'none', 'null') else s

    has_id_col = col_id is not None
    recs = []
    for _, row in df_input.iterrows():
        iid = row.get(col_id) if col_id else ''
        iname = _clean(row.get(col_name) if col_name else '')
        isku = _clean(row.get(col_sku) if col_sku else '')
        rec = dict(row)

        # ── 新品：ID为空(且表里有ID列) 或 名称/SKU 标了「新品/待生成链接」→ 不查历史价 ──
        is_new, new_reason = M.is_new_product(iid, iname, isku, has_id_col)
        if is_new:
            rec.update({'匹配等级': '新品·无历史价', '匹配度%': None, '依据': new_reason,
                        '命中标准商品名': '', '命中商品ID': '',
                        '取价结论': '新品·无历史价', '参照信息(运营判断)': ''})
            for p_ in M.PERIODS:
                rec[f'{p_}到手价($)'] = None
            recs.append(rec)
            continue

        res = match_one(iid, iname, isku, products, name_pool)
        if not res.get('matched'):
            rec.update({'匹配等级': '未匹配', '匹配度%': res.get('score'),
                        '依据': '底表中未找到足够相似的商品',
                        '命中标准商品名': '', '命中商品ID': '',
                        '取价结论': '无价', '参照信息(运营判断)': ''})
            for p_ in M.PERIODS:
                rec[f'{p_}到手价($)'] = None
            recs.append(rec)
            continue
        pp = res['period_price']
        for p_ in M.PERIODS:
            v = pp.get(p_, {})
            price = v.get('price') if v.get('verdict') == 'take' else None
            rec[f'{p_}到手价($)'] = round(float(price), 2) if price is not None and pd.notna(price) else None
        rec.update({
            '匹配等级': res['grade'],
            '匹配度%': res['score'],
            '依据': res['reason'],
            '命中标准商品名': res['std_name'],
            '命中商品ID': res['matched_ids'],
            '取价结论': res['conclusion'],
            '参照信息(运营判断)': res['ref_info'],
        })
        recs.append(rec)
    return pd.DataFrame(recs)
