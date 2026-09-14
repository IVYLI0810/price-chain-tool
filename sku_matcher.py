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

# 多语言颜色词（中/韩/英文变体）→ 归一到英文规范色，供规格维度比对
# 输入表单可能写中文/韩文，底表多为英文，需归一后才能判断"是否同色"
COLOR_MULTILINGUAL = {
    # 中文
    '黑色': 'black', '黑': 'black', '白色': 'white', '白': 'white', '灰色': 'gray', '灰': 'gray',
    '蓝色': 'blue', '蓝': 'blue', '红色': 'red', '红': 'red', '绿色': 'green', '绿': 'green',
    '黄色': 'yellow', '黄': 'yellow', '粉色': 'pink', '粉红': 'pink', '粉': 'pink',
    '紫色': 'purple', '紫': 'purple', '橙色': 'orange', '橙': 'orange', '棕色': 'brown', '棕': 'brown',
    '咖啡色': 'brown', '米色': 'beige', '米白': 'beige', '银色': 'silver', '银': 'silver',
    '金色': 'gold', '金': 'gold', '藏青': 'navy', '卡其': 'khaki', '透明': 'clear', '奶油': 'cream',
    # 韩文
    '블랙': 'black', '검정': 'black', '화이트': 'white', '흰색': 'white', '그레이': 'gray',
    '회색': 'gray', '블루': 'blue', '파랑': 'blue', '레드': 'red', '빨강': 'red',
    '그린': 'green', '초록': 'green', '옐로우': 'yellow', '노랑': 'yellow', '핑크': 'pink',
    '퍼플': 'purple', '보라': 'purple', '오렌지': 'orange', '브라운': 'brown', '베이지': 'beige',
    '실버': 'silver', '골드': 'gold', '네이비': 'navy', '아이보리': 'ivory',
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
    # 分隔：换行、竖线|、斜杠/ 都当作选项分隔符
    # （CSV底表把单元格内换行统一存成 |，避免多行记录，故这里也认 |）
    parts = re.split(r'[\r\n|]+|/', s)
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


# ─────────────────────────────────────────────
# 规格维度抽取与分类比对（2026-09 新增）
#   同品判定靠 ID+名称；取价判定靠规格维度分类：
#   HARD（影响单价）：数量/容量/瓦数/长度/电压/Pro等修饰词/底型号 → 差异时不取价、只给参照
#   SOFT（不影响单价）：颜色/尺码/同底型号后缀(色码) → 差异时照常取价、依据里注明
# ─────────────────────────────────────────────
QTY_RE = re.compile(r'(\d+)\s*(?:pcs|pc|pieces|sets?|packs?|pairs?|개입|个装)|:\s*(\d{2,4})$|(\d+)\s*in\s*1', re.I)
CAP_RE = re.compile(r'(\d+(?:\.\d+)?)\s*(mah|wh|gb|tb|mb)\b', re.I)
WATT_RE = re.compile(r'(\d+(?:\.\d+)?)\s*(kw|w)\b', re.I)
VOLT_RE = re.compile(r'(\d+(?:\.\d+)?)\s*v\b', re.I)
LEN_RE = re.compile(r'(\d+(?:\.\d+)?)\s*(inch|inches|cm|m|ft)\b', re.I)
SIZE_RE = re.compile(r'(?:eu|us|uk)\s?(\d+(?:\.\d+)?)\b|(\d{3})\s?mm\b', re.I)
CODEBASE_RE = re.compile(r'^([A-Za-z]{2,6}\d+)(?:[-–]([A-Za-z0-9]+))?$', re.I)

HARD_DIMS = ('qty', 'cap', 'watt', 'len', 'volt', 'mod', 'base')
SOFT_DIMS = ('color', 'size', 'suffix')
AMBIGUOUS_DIMS = ('mod', 'base', 'qty', 'cap', 'watt', 'len', 'volt')
DIM_LABEL = {'qty': '数量', 'cap': '容量', 'watt': '瓦数', 'len': '长度', 'volt': '电压',
             'mod': '型号(Pro/Plus等)', 'base': '底型号', 'color': '颜色', 'size': '尺码',
             'suffix': '型号后缀'}


def extract_dims(text) -> dict:
    """从一条 SKU 选项文本抽出各规格维度值（没有则为 None）。"""
    t = _fullwidth_to_half(str(text or '')).strip()
    d = {k: None for k in ('qty', 'cap', 'watt', 'len', 'volt', 'size', 'color', 'mod', 'base', 'suffix')}
    m = QTY_RE.search(t)
    if m:
        d['qty'] = next((g for g in m.groups() if g), None)
    m = CAP_RE.search(t)
    if m:
        d['cap'] = f'{float(m.group(1)):g}{m.group(2).lower()}'
    m = WATT_RE.search(t)
    if m:
        d['watt'] = f'{float(m.group(1)):g}{m.group(2).lower()}'
    m = VOLT_RE.search(t)
    if m:
        d['volt'] = f'{float(m.group(1)):g}v'
    m = LEN_RE.search(t)
    if m:
        d['len'] = f'{float(m.group(1)):g}{m.group(2).lower()}'
    m = SIZE_RE.search(t)
    if m:
        d['size'] = next((g for g in m.groups() if g), None)
    low = t.lower()
    cols = {COLOR_ALIAS.get(w, w) for w in COLOR_WORDS if w in low}
    for term, canon in COLOR_MULTILINGUAL.items():   # 中/韩文颜色
        if term in t:
            cols.add(COLOR_ALIAS.get(canon, canon))
    if cols:
        d['color'] = frozenset(cols)
    mods = {w for w in MODEL_MODIFIERS if re.search(r'\b' + w + r'\b', low)}
    if mods:
        d['mod'] = frozenset(mods)
    cm = CODEBASE_RE.match(t)
    if cm:
        d['base'] = cm.group(1).upper()
        d['suffix'] = (cm.group(2) or '').upper() or None
    return d


def mod_varies(option_texts) -> bool:
    """该商品的选项里 mod 维度（Pro/Plus/Max…）是否有>1种取值（含"无修饰词"这一取值）。
    True 说明商品确实区分 Pro/非Pro，比对时须严格。"""
    vals = set()
    for t in option_texts:
        m = extract_dims(t).get('mod')
        vals.add(frozenset(m) if m else frozenset())
    return len(vals) > 1


def dim_diffs(a: dict, b: dict, strict_mod: bool = False):
    """比较两侧规格维度，返回 (hard差异列表, soft差异列表)。
    默认只比双方都写了的维度；strict_mod=True 时，mod 维度（Pro/Plus等）
    一方有一方无也算 hard 冲突（该商品确实区分 Pro/非Pro）。"""
    hard, soft = [], []
    for k in HARD_DIMS:
        va, vb = a.get(k), b.get(k)
        if k == 'mod' and strict_mod:
            sa = frozenset(va) if va else frozenset()
            sb = frozenset(vb) if vb else frozenset()
            if sa != sb:
                hard.append(k)
            continue
        if va is None or vb is None:
            continue
        if va != vb:
            hard.append(k)
    for k in SOFT_DIMS:
        va, vb = a.get(k), b.get(k)
        if va is None or vb is None:
            continue
        if k == 'color':
            if not (va & vb):          # 颜色集合完全不交集才算差异
                soft.append(k)
        elif k == 'suffix':
            if a.get('base') and a.get('base') == b.get('base') and va != vb:
                soft.append(k)         # 同底型号的后缀(色码)差异 → 软
        elif va != vb:
            soft.append(k)
    return hard, soft


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


def grade(id_hit: bool, name_score: float) -> str:
    """同款判定 + 分级（只看 ID/名称；规格差异由取价结论单独表达）。"""
    name_score = name_score or 0.0
    # 根本不算同款：ID没命中且名称不够像
    if not id_hit and name_score < NAME_GATE:
        return '未匹配'
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
    """把商品ID统一成干净字符串：兼容浮点(…609.0)/科学计数(1.005e15)，空值归为''。"""
    if v is None:
        return ''
    s = str(v).strip()
    if s.lower() in ('nan', 'none', 'null', ''):
        return ''
    # 浮点/科学计数 → 还原成整数字符串（AE 商品ID 16位，在 float64 精确范围内）
    try:
        f = float(s)
        if f.is_integer():
            return str(int(f))
    except ValueError:
        pass
    if s.endswith('.0'):
        s = s[:-2]
    return s


def is_valid_product_id(pid: str) -> bool:
    """有效的 AE 商品ID = 纯数字且足够长。文字（新品/待生成链接）、空、乱填都不算。"""
    return bool(pid) and pid.isdigit() and len(pid) >= 6


# 新品标记词（用户手写的标注；不用裸"new"/"新"以免误伤真实品名如"New Jet Fan"）
NEW_PRODUCT_KEYWORDS = ['新品', '待生成链接', '待生成', '待上链接', '待生成id', '未上架', '未生成链接', '待发链接']


def is_new_product(input_id, input_name, input_sku, has_id_col=True):
    """
    判断是否新品（新品此前没出现过，不该有历史价 → 跳过匹配）。
      1) 表格有【商品ID】列，但该行 ID 不是有效数字ID：
         - 为空 → 新品
         - 写了「新品 / 待生成链接」等文字 → 新品（这类标注常直接写在ID列里）
      2) 商品ID/名称/SKU 任一含「新品/待生成链接」等标记词 → 新品
    返回 (是否新品, 原因)
    """
    pid = norm_id(input_id)
    if has_id_col and not is_valid_product_id(pid):
        if pid == '':
            return True, '商品ID为空 → 判为新品（尚未上架/待生成链接），无历史价'
        return True, f'商品ID不是有效ID（填的是「{pid}」）→ 判为新品/待生成链接，无历史价'
    # ID 有效时，再扫 ID/名称/SKU 三处文字里的新品标记
    text = f'{input_id or ""} {input_name or ""} {input_sku or ""}'
    for kw in NEW_PRODUCT_KEYWORDS:
        if kw in text:
            return True, f'标注了「{kw}」→ 判为新品，无历史价'
    return False, ''


