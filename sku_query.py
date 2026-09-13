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
from pathlib import Path

import pandas as pd
from rapidfuzz import process, fuzz

import sku_matcher as M

_HERE = Path(__file__).parent
DB_PATH = next((str(p) for p in [_HERE / 'history_data.db', Path('history_data.db')] if p.exists()), None)

_NAME_CANDIDATES = 15  # 名称粗筛保留的候选数


def load_detail(db_path=None) -> pd.DataFrame:
    db_path = db_path or DB_PATH
    conn = sqlite3.connect(db_path)
    df = pd.read_sql_query('SELECT std_name,product_id,period,sku_raw,price,raw_val FROM sku_detail', conn)
    conn.close()
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


def _id_score(input_id, ids) -> float:
    if not input_id:
        return None
    if not ids:
        return 0.0
    return 100.0 if str(input_id).strip() in ids else 0.0


def _pick_period_price(prod_rows, input_sku):
    """
    在选中商品内，按期次挑价：
      - 有输入SKU：挑型号最贴合的行（优先 model_ok!=False），取该行到手价
      - 无输入SKU：取当期最低价（与原口径一致）
    返回 {period: (price, sku_raw, sku_score, model_ok)}
    """
    out = {}
    have_sku = bool(M.split_sku_options(input_sku))
    for period in M.PERIODS:
        pr = prod_rows[prod_rows['period'] == period]
        if pr.empty:
            continue
        if not have_sku:
            # 取当期最低价那行
            pr2 = pr.dropna(subset=['price'])
            row = (pr2 if not pr2.empty else pr).sort_values('price').iloc[0]
            out[period] = {'price': row['price'], 'sku_raw': row['sku_raw'],
                           'sku_score': None, 'model_ok': None}
            continue
        best = None
        for _, row in pr.iterrows():
            sim = M.sku_similarity(input_sku, row['options'])
            sc = sim['score'] if sim['score'] is not None else -1
            # model_ok False 的降权，避免选错型号
            eff = sc if sim['model_ok'] is not False else sc * 0.3
            if best is None or eff > best['eff']:
                best = {'eff': eff, 'price': row['price'], 'sku_raw': row['sku_raw'],
                        'sku_score': sim['score'], 'model_ok': sim['model_ok'],
                        'best_option': sim['best_option']}
        if best:
            # 护栏：该期只有"另一个型号"的行 → 不给错型号的价，留空并标记
            if best['model_ok'] is False:
                out[period] = {'price': None, 'sku_raw': best['sku_raw'],
                               'sku_score': best['sku_score'], 'model_ok': False,
                               'model_conflict': True}
            else:
                out[period] = best
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

    # ── SKU/型号：只在选中的同款商品内部用来挑价，不决定是不是同款 ──
    all_opts = [o for opts in p['rows']['options'].tolist() for o in opts]
    sksim = M.sku_similarity(input_sku, all_opts)
    model_ok = sksim['model_ok']

    score = M.match_score(id_hit, best['name_score'])
    g = M.grade(id_hit, best['name_score'], model_ok)

    # ── 按期取价（SKU 在这里发挥作用：同名商品下按型号/选项挑对应价） ──
    period_price = _pick_period_price(p['rows'], input_sku)

    # ── 依据（人话） ──
    reasons = []
    if id_hit:
        reasons.append('商品ID完全一致')
    elif input_id:
        reasons.append(f"商品ID未命中，靠商品名匹配到同款（底表用过ID：{'/'.join(sorted(p['ids'])) or '无'}）")
    else:
        reasons.append('无商品ID，靠商品名匹配')
    reasons.append(f"商品名相似 {best['name_score']:.0f}%")
    if sksim['score'] is not None:
        reasons.append('同款内按SKU挑价：' + sksim['detail'])
    if g == '型号不符':
        reasons.append('⚠ 商品名对上了，但输入的型号(如Pro款)在底表该商品里没有，未取价')

    return {
        'matched': g not in ('未匹配', '型号不符'),
        'grade': g,
        'score': score,
        'std_name': p['std_name'],
        'matched_ids': '/'.join(sorted(p['ids'])),
        'period_price': period_price,
        'reason': '；'.join(reasons) if reasons else '—',
        'id_hit': id_hit, 'name_score': best['name_score'], 'sku_score': sksim['score'],
    }



def match_batch(df_input: pd.DataFrame, col_id, col_name, col_sku, db_path=None):
    """
    批量匹配。df_input 至少含 col_id/col_name/col_sku 三列（列名由调用方传入）。
    返回结果 DataFrame（原列 + 4期价 + 匹配度/等级/依据/命中信息）。
    """
    detail = load_detail(db_path)
    products = build_products(detail)
    name_pool = [p['name_norm'] for p in products]

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
                        '命中标准商品名': '', '命中商品ID': ''})
            for p_ in M.PERIODS:
                rec[f'{p_}到手价($)'] = None
            recs.append(rec)
            continue

        res = match_one(iid, iname, isku, products, name_pool)
        if not res.get('matched'):
            # 区分"完全没找到"与"找到了但型号不符"
            if res.get('grade') == '型号不符':
                reason = res.get('reason', '型号不符')
                hit_name = res.get('std_name', '')
                hit_ids = res.get('matched_ids', '')
                g = '型号不符'
            else:
                reason = '底表中未找到足够相似的商品'
                hit_name, hit_ids, g = '', '', '未匹配'
            rec.update({'匹配等级': g, '匹配度%': res.get('score'), '依据': reason,
                        '命中标准商品名': hit_name, '命中商品ID': hit_ids})
            for p_ in M.PERIODS:
                rec[f'{p_}到手价($)'] = None
            recs.append(rec)
            continue
        pp = res['period_price']
        for p_ in M.PERIODS:
            v = pp.get(p_, {}).get('price')
            rec[f'{p_}到手价($)'] = round(float(v), 2) if v is not None and pd.notna(v) else None
        rec.update({
            '匹配等级': res['grade'],
            '匹配度%': res['score'],
            '依据': res['reason'],
            '命中标准商品名': res['std_name'],
            '命中商品ID': res['matched_ids'],
        })
        recs.append(rec)
    return pd.DataFrame(recs)
