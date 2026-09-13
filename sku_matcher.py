"""
SKU 历史价 · 模糊匹配引擎
──────────────────────────────────────────────
输入一行 (商品ID / 商品名称 / sku选项)，在 SKU明细底表里找到"同一个商品、同一个型号"，
返回 11/3/6/8 月的历史到手价 + 匹配度 + 人话依据。

核心难点（也是本引擎的设计重点）：
  同一个商品ID、同一期，底表里可能有多行，SKU 写法不同、价格也不同。
  例：AULA F108 —— "F108 Pro" 系列 $32.6，"F108(非Pro)" $27.38，是两个型号。
  所以不能无脑取当期最低价，必须先把【型号】拆出来，按输入 SKU 锁定对应型号的价。

打分三维度：
  ID 分   —— 商品ID 是否命中（含"换ID同款"）
  名称分  —— 商品名清洗后的相似度（品牌+型号是最强信号）
  SKU 分  —— 权重最高；以【型号签名】为硬门槛，再看颜色/规格/插头
"""

import re
import unicodedata

from rapidfuzz import fuzz

# ─────────────────────────────────────────────
# 期次映射（Excel 写法 → 内部简写）
# ─────────────────────────────────────────────
PERIOD_MAP = {
    '25年11月': '11月', '11月': '11月',
    '26年3月': '3月', '3月': '3月',
    '26年6月': '6月', '6月': '6月',
    '26年8月': '8月', '8月': '8月',
}
PERIODS = ['11月', '3月', '6月', '8月']

# ─────────────────────────────────────────────
# 词表
# ─────────────────────────────────────────────
# 型号修饰词：改变型号身份（Pro 与 非Pro 是不同型号）
MODEL_MODIFIERS = {'pro', 'plus', 'max', 'mini', 'ultra', 'lite', 'gen', 'se', 'air'}

# 颜色同义归一（Gray=Grey、BK=Black …）
COLOR_ALIAS = {
    'grey': 'gray', 'bk': 'black', 'bl': 'blue', 'wt': 'white', 'wh': 'white',
    'wht': 'white', 'blu': 'blue', 'blk': 'black', 'slv': 'silver', 'gld': 'gold',
    'nv': 'navy', 'pk': 'pink', 'pp': 'purple', 'gn': 'green', 'rd': 'red',
    'darkgray': 'dark gray', 'smoky pink': 'pink', 'ivory': 'white',
    'sea salt blue': 'blue', 'dark storm': 'black', 'star trek': 'gray',
}
COLOR_WORDS = {
    'black', 'white', 'gray', 'blue', 'red', 'green', 'yellow', 'pink', 'purple',
    'navy', 'silver', 'gold', 'orange', 'brown', 'beige', 'ivory', 'cream', 'khaki',
    'transparent', 'clear', 'mix', 'dark', 'light', 'sky', 'rose', 'mint', 'cyan',
}

# 规格单位（这些数字是"规格"不是"型号"，比如 256GB / 45W / 110mm）
SPEC_UNITS = r'(?:gb|mb|tb|mah|wh|w|v|a|mm|cm|inch|in|k|hz|pcs|pc|g|l|ml|kg|core|pin)'

# 需要丢掉的噪声行（库存、占位）
NOISE_LINE_RE = re.compile(r'可售|库存|stock|单一\s*sku|单sku|^sku$|^-$|^/$', re.I)


# ─────────────────────────────────────────────
# 基础清洗
# ─────────────────────────────────────────────
def _fullwidth_to_half(s: str) -> str:
    return unicodedata.normalize('NFKC', s)


def norm_text(s) -> str:
    """通用清洗：全角转半角、小写、去多余空白与常见标点。用于名称/选项比对。"""
    if s is None:
        return ''
    s = _fullwidth_to_half(str(s)).lower()
    s = s.replace('\u00a0', ' ')
    # 统一分隔标点为空格
    s = re.sub(r'[,，;；:：/\\\|\(\)\[\]\{\}\'"“”‘’·、。，]', ' ', s)
    s = re.sub(r'\s+', ' ', s).strip()
    return s


# ─────────────────────────────────────────────
# SKU 规格 → 拆成一个个"干净选项"
# ─────────────────────────────────────────────
def _strip_qty(opt: str) -> str:
    """剥掉选项尾巴上的数量/库存数字，但保留规格数字(256GB/110mm/4K)。"""
    o = opt.strip()
    # 1) 明确的分隔符+数字尾巴： ":200" ",300" "，100"
    o = re.sub(r'[,，:：]\s*\d+\s*$', '', o)
    # 2) 结尾独立裸数字（前面是空格），且不是规格单位： "White 50" -> "White"
    m = re.search(r'\s(\d{1,5})\s*$', o)
    if m:
        tail = m.group(1)
        # 若这个数字紧跟单位（在原文里）则不剥；这里 tail 是纯数字无单位 → 视为数量
        o = o[:m.start()].strip()
    return o.strip()


def split_sku_options(raw) -> list:
    """把一格 SKU规格 文本拆成多个干净选项字符串（已去噪、去数量尾巴）。"""
    if raw is None:
        return []
    s = _fullwidth_to_half(str(raw))
    # 分隔：换行 与 斜杠 都当作选项分隔符
    parts = re.split(r'[\r\n]+|/', s)
    opts = []
    for p in parts:
        p = p.strip()
        if not p:
            continue
        if NOISE_LINE_RE.search(p):
            continue
        p = _strip_qty(p)
        p = p.strip(' ,，:：-')
        if not p or NOISE_LINE_RE.search(p):
            continue
        opts.append(p)
    # 去重（保序）
    seen, out = set(), []
    for o in opts:
        k = norm_text(o)
        if k and k not in seen:
            seen.add(k)
            out.append(o)
    return out


# ─────────────────────────────────────────────
# 型号签名 / 规格 / 颜色 抽取
# ─────────────────────────────────────────────
_SPEC_RE = re.compile(r'^\d+(?:\.\d+)?\s*' + SPEC_UNITS + r'$', re.I)
_MODELCODE_RE = re.compile(r'(?=[a-z0-9]*[a-z])(?=[a-z0-9]*\d)[a-z0-9][a-z0-9\-\./]*')


def extract_signature(text) -> dict:
    """从一段文本（选项或商品名）抽出：型号码 / 型号修饰词 / 规格 / 颜色。"""
    t = norm_text(text)
    toks = t.split()
    model_codes, mods, specs, colors, other = set(), set(), set(), set(), set()
    for tok in toks:
        tok = tok.strip('.')
        if not tok:
            continue
        low = tok.lower()
        if low in MODEL_MODIFIERS:
            mods.add(low)
            continue
        if low in COLOR_WORDS or low in COLOR_ALIAS:
            colors.add(COLOR_ALIAS.get(low, low))
            continue
        # 规格：数字(+单位)，如 256gb / 45w / 110mm / 4k / 8
        if _SPEC_RE.match(low) or re.match(r'^\d+$', low):
            specs.add(low.replace(' ', ''))
            continue
        # 型号码：字母+数字混合，如 f108 / wx242 / m60 / qi2.2 / 100p
        if _MODELCODE_RE.fullmatch(low):
            model_codes.add(low)
            continue
        # 多词型号里可能带数字的复合词已在上面命中；其余归 other
        # 再尝试从复合词里抠型号码
        found = _MODELCODE_RE.findall(low)
        if found and any(c.isdigit() for c in low) and any(c.isalpha() for c in low):
            for f in found:
                model_codes.add(f)
        else:
            other.add(low)
    # 颜色别名再归一
    colors = {COLOR_ALIAS.get(c, c) for c in colors}
    return {'model': model_codes, 'mod': mods, 'spec': specs, 'color': colors, 'other': other}


def _model_signature(sig: dict) -> set:
    """型号身份 = 型号码 + 修饰词(pro/plus…)。这是硬门槛。"""
    return set(sig['model']) | set(sig['mod'])


# ─────────────────────────────────────────────
# 相似度打分
# ─────────────────────────────────────────────
def _jaccard(a: set, b: set) -> float:
    if not a and not b:
        return 1.0
    if not a or not b:
        return 0.0
    return len(a & b) / len(a | b)


def name_similarity(a: str, b: str) -> float:
    """商品名相似度 0~100。综合 token 重合 + 序列相似 + 型号命中。"""
    na, nb = norm_text(a), norm_text(b)
    if not na or not nb:
        return 0.0
    base = max(fuzz.token_set_ratio(na, nb), fuzz.token_sort_ratio(na, nb))
    partial = fuzz.partial_ratio(na, nb)
    seq = fuzz.ratio(na, nb)
    score = 0.5 * base + 0.3 * partial + 0.2 * seq
    # 型号签名一致性做加成/惩罚
    sa, sb = extract_signature(na), extract_signature(nb)
    ma, mb = _model_signature(sa), _model_signature(sb)
    if ma and mb:
        mj = _jaccard(ma, mb)
        score = score * (0.6 + 0.4 * mj)
    return round(min(100.0, score), 1)


def sku_similarity(input_sku: str, cand_options: list) -> dict:
    """
    输入 SKU 选项 vs 底表某行的选项集合。
    返回 {score:0~100, best_option, model_ok:bool/None, detail}
    以【型号】为硬门槛：型号矛盾(Pro vs 非Pro / 型号码不同) → 大幅扣分。
    """
    in_opts = split_sku_options(input_sku)
    if not in_opts:
        # 输入没写 SKU → SKU 维度不表态，交给 ID/名称
        return {'score': None, 'best_option': None, 'model_ok': None,
                'detail': '输入无SKU选项，按ID/名称判断'}
    if not cand_options:
        return {'score': 0.0, 'best_option': None, 'model_ok': None,
                'detail': '底表该行无SKU规格'}

    in_sig = extract_signature(' '.join(in_opts))
    in_model = _model_signature(in_sig)

    best = {'score': -1, 'opt': None, 'model_ok': None}
    for co in cand_options:
        co_sig = extract_signature(co)
        co_model = _model_signature(co_sig)
        # 文本层面相似度
        txt = max(fuzz.token_set_ratio(norm_text(' '.join(in_opts)), norm_text(co)),
                  fuzz.partial_ratio(norm_text(' '.join(in_opts)), norm_text(co)),
                  fuzz.ratio(norm_text(in_opts[0]), norm_text(co)))
        # 型号门槛
        model_ok = None
        model_mult = 1.0
        if in_model and co_model:
            mj = _jaccard(in_model, co_model)
            model_ok = mj >= 0.999
            # 修饰词(pro/plus…)冲突 → 硬惩罚
            if in_sig['mod'] != co_sig['mod'] and (in_sig['mod'] or co_sig['mod']):
                # 一方有 pro 一方没有 = 不同型号
                if not (in_sig['mod'] & co_sig['mod']):
                    model_mult = 0.25
                    model_ok = False
            else:
                model_mult = 0.35 + 0.65 * mj
        # 颜色/规格作为次要加成
        cs = _jaccard(in_sig['color'], co_sig['color'])
        sp = _jaccard(in_sig['spec'], co_sig['spec'])
        adj = txt * model_mult + 8 * cs + 6 * sp
        adj = min(100.0, adj)
        if adj > best['score']:
            best = {'score': adj, 'opt': co, 'model_ok': model_ok}

    score = round(best['score'], 1)
    if best['model_ok'] is True:
        detail = f"SKU选项命中底表『{best['opt']}』，型号一致"
    elif best['model_ok'] is False:
        detail = f"型号不一致（输入={sorted(in_model)} vs 底表『{best['opt']}』），判为不同型号"
    else:
        detail = f"SKU选项近似命中『{best['opt']}』"
    return {'score': score, 'best_option': best['opt'], 'model_ok': best['model_ok'], 'detail': detail}


# ─────────────────────────────────────────────
# 同款判定（名称优先）+ 匹配度分级
# ─────────────────────────────────────────────
# 逻辑：SKU 不参与"是不是同款"的判定，只在选中同款后用来挑对应价。
#   1) 商品ID命中 → 认同款
#   2) ID没命中 → 商品名相似度必须 ≥ NAME_GATE 才认同款；不够像直接淘汰
# 名称不够像时，SKU 再一致也没意义（同名不同SKU才用SKU区分价格）。
NAME_GATE = 80.0


def grade(id_hit: bool, name_score: float, model_ok) -> str:
    """同款判定 + 分级。name_score 0~100；model_ok 来自选中商品内的 SKU 型号比对。"""
    name_score = name_score or 0.0
    # 根本不算同款：ID没命中且名称不够像
    if not id_hit and name_score < NAME_GATE:
        return '未匹配'
    # 名称/ID对上了，但输入的型号在底表该商品里找不到（如只有非Pro、没有Pro）
    if model_ok is False:
        return '型号不符'
    if id_hit and name_score >= 75:
        return '高'
    if id_hit:
        return '中'          # ID一致但名称差很多（可能改写/换名），仍认同款
    if name_score >= 90:
        return '高'
    if name_score >= NAME_GATE:
        return '中'
    return '低'


def match_score(id_hit: bool, name_score: float) -> float:
    """对外展示的匹配度%：只看 ID + 名称，SKU 不计入（SKU 只用于同款内挑价）。"""
    name_score = name_score or 0.0
    if id_hit:
        return round(0.4 * 100 + 0.6 * name_score, 1)
    return round(name_score, 1)


# ─────────────────────────────────────────────
# 商品ID 归一化 + 新品识别
# ─────────────────────────────────────────────
def norm_id(v) -> str:
    """把商品ID统一成干净字符串：去掉浮点尾巴 .0、空值(nan/none)归为''。"""
    if v is None:
        return ''
    s = str(v).strip()
    if s.lower() in ('nan', 'none', 'null', ''):
        return ''
    if s.endswith('.0'):
        s = s[:-2]
    return s


# 新品标记词（用户手写的标注；不用裸"new"/"新"以免误伤真实品名如"New Jet Fan"）
NEW_PRODUCT_KEYWORDS = ['新品', '待生成链接', '待生成', '待上链接', '待生成id', '未上架', '未生成链接']


def is_new_product(input_id, input_name, input_sku, has_id_col=True):
    """
    判断是否新品（新品此前没出现过，不该有历史价 → 跳过匹配）。
      1) 表格有【商品ID】列、但该行ID为空 → 新品（尚未上架/待生成链接）
      2) 商品名称或SKU里写了「新品/待生成链接」等标记 → 新品
    返回 (是否新品, 原因)
    """
    pid = norm_id(input_id)
    if has_id_col and pid == '':
        return True, '商品ID为空 → 判为新品（尚未上架/待生成链接），无历史价'
    text = f'{input_name or ""} {input_sku or ""}'
    for kw in NEW_PRODUCT_KEYWORDS:
        if kw in text:
            return True, f'标注了「{kw}」→ 判为新品，无历史价'
    return False, ''


