#!/usr/bin/env python3
"""
逆向 Zotero 域代码插入器（Reverse Zotero Field-Code Inserter）
=============================================================
把 .docx 里的纯文字参考文献转成 Zotero 可识别的域代码（无需 Word 插件）。

依赖: python-docx, lxml  (pip install python-docx lxml)

需要事先准备 item_mapping.json:
  { "<ref_key>": {"itemKey":"ABCD1234","uri":"http://zotero.org/users/<uid>/items/ABCD1234",
                  "csl":<CSL-JSON>,"doi":"10.xxxx/...","first_author":"Smith","year":"2007"}, ... }
其中 csl 是 CSL-JSON（type/title/author/issued/container-title/volume/issue/page/DOI...）。
用 CrossRef(https://api.crossref.org/works/<DOI>) 解析 DOI 得到；Zotero MCP(http://127.0.0.1:23120/mcp)
的 write_item(action=create) 建 item 得 itemKey；userID 从 ~/Zotero/zotero.sqlite 的 users 表读。

用法见文件末尾示例 / SKILL.md。
"""
import zipfile, json, re, unicodedata, copy, random, string
from lxml import etree

try:
    import config
except Exception:
    config = None

def _resolve_uri_prefix(uri_prefix):
    if uri_prefix is not None:
        return uri_prefix
    if config is not None:
        return config.get_uri_prefix()
    return "http://zotero.org/users/<YOUR_USER_ID>/items/"

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
SCHEMA = "https://github.com/citation-style-language/schema/raw/master/csl-citation.json"
XMLSPACE = '{http://www.w3.org/XML/1998/namespace}space'


# ---------- 基础工具 ----------
def _norm(s):
    if not s: return ""
    # 保留 CJK（中文文献），NFKD 折叠拉丁重音，去标点/空白
    s = unicodedata.normalize('NFKD', s).lower()
    return re.sub(r'[^a-z0-9一-鿿]', '', s)

def _is_cjk(s):
    """是否含 CJK 统一表意文字（中文/日文汉字）。"""
    return any('一' <= ch <= '鿿' for ch in s)

def _first_author_lastname(csl):
    au = csl.get('author', [])
    if au and au[0].get('family'): return au[0]['family']
    if au and au[0].get('literal'): return au[0]['literal']
    return ""

def _year(csl):
    dp = csl.get('issued', {}).get('date-parts', [[None]])
    return str(dp[0][0]) if dp and dp[0] and dp[0][0] else ""


# ---------- 引文匹配器 ----------
def build_matcher(item_mapping):
    """返回 (by_year: {year:[(lastname, ref_key)]}, nih_key)。需 item_mapping 每项含 'csl'。"""
    by_year = {}; nih = None
    for k, v in item_mapping.items():
        csl = v['csl']
        if v.get('is_nih'): nih = k; continue
        ln = _first_author_lastname(csl); y = _year(csl)
        if ln and y: by_year.setdefault(y, []).append((ln, k))
    return by_year, nih

def _match_cjk_surname(part, var, offset):
    """在 part 里匹配中文姓。中文姓后通常直接跟名（继续 CJK）或标准边界；
    姓前须是名字起始（段首/空格/逗号等分隔符或 '和/及/与/或/等' 连接词）。返回位置或 -1。
    """
    for m in re.finditer(re.escape(var), part):
        i = m.start()
        if i != offset:
            prev = part[i - 1]
            if prev and not (prev.isspace() or prev in ',(;(&和及与或等、'):
                continue
        e = m.end()
        after = part[e] if e < len(part) else ''
        if after:
            if _is_cjk(after):
                return i          # 姓后接名（"张" + "三"）
            if not re.match(r'\s*(?:et al\.?|and\b|&\b|,|\)|$)', after):
                continue
        return i
    return -1


def match_part(part, year, by_year, nih_key):
    """在 part 文本里找年份=year、第一作者姓出现的条目。大小写不敏感，要求姓后跟边界。"""
    if ('NIH' in part or 'NOT-OD' in part) and nih_key: return nih_key
    part_l = part.lstrip(); offset = len(part) - len(part_l); cands = []
    for (ln, k) in by_year.get(year, []):
        variants = [ln]
        if ' ' in ln: variants.append(ln.split()[-1])  # 多词姓也试末词(De Vries->Vries)
        for var in variants:
            if not var: continue
            if _is_cjk(var):
                pos = _match_cjk_surname(part, var, offset)
                if pos >= 0:
                    cands.append((pos, k)); break
            else:
                found = False
                for m in re.finditer(re.escape(var), part, re.IGNORECASE):
                    is_start = (m.start() == offset)
                    bm = re.match(r'\s*(et al\.?|and\b|&\b|,|\)|$)', part[m.end():])
                    ok = bool(bm) and not (bm.group(1) == ',' and not is_start)
                    if ok:
                        cands.append((m.start(), k)); found = True; break
                if found: break
    if not cands: return None
    cands.sort()
    return cands[0][1]

NARR = re.compile(r'([A-Z][\wÀ-ÿ\'\-]+(?:\s+(?:and|&)\s+[\wÀ-ÿ\'\-]+(?:\s+[\wÀ-ÿ\'\-]+){0,2})?(?:\s+et al\.)?)\s*\((\d{4})\)')
NARR_CJK = re.compile(r'([一-鿿]{2,10}(?:、[一-鿿]{2,10})*(?:\s*等)?)\s*[（(](\d{4})[）)]')
PAREN = re.compile(r'[（(]([^（）()]*?\d{4}[^（）()]*?)[）)]')

def find_citations(text, by_year, nih_key):
    """返回 [(start, end, display_text, [ref_key])]。括号式 + 叙述式（拉丁 + 中文，兼容全角标点）。"""
    spans = []
    for m in NARR.finditer(text):
        k = match_part(m.group(1), m.group(2), by_year, nih_key)
        if k: spans.append((m.start(), m.end(), m.group(0), [k]))
    for m in NARR_CJK.finditer(text):
        k = match_part(m.group(1), m.group(2), by_year, nih_key)
        if k: spans.append((m.start(), m.end(), m.group(0), [k]))
    for m in PAREN.finditer(text):
        g = m.group(1)
        if not re.search(r'\d{4}', g): continue
        if re.match(r'^\d{4}', g.strip()): continue                  # 跳过"(2016 to 2026)"等
        if not re.search(r'[A-Z][a-z]{2,}|NIH|NOT-OD|[一-鿿]{2,}', g): continue  # 必须有作者词(否则非引文)
        parts = [p.strip() for p in re.split(r'[;；]', g)]; items = []; ok = True
        for part in parts:
            ym = re.search(r'(\d{4})', part)
            if not ym: continue                 # 跳过无年份的分片(如"ALSPAC;")
            k = match_part(part, ym.group(1), by_year, nih_key)
            if not k: ok = False; break
            items.append(k)
        if ok and items:
            s, e = m.start(), m.end()
            if not any(not (e <= s2 or s >= e2) for s2, e2, _, _ in spans):
                spans.append((s, e, m.group(0), items))
    spans.sort()
    return spans


# ---------- DOI/精确文本括号引文（占位符风格正文） ----------
# 适用于正文用 "(doi:10.x; 10.y)"、"(10.x; 10.y)"、"(FDA label, …)"、"(NCT…)"、
# "(Kynexis, 2024)" 等占位代替"作者+年份"的手稿。作者-年份匹配器对这类文档一无所获。
# DOI 词元：老式 Elsevier DOI 内含年份段括号，如 10.1016/S0306-4530(98)00014-6。
# 朴素的 `10\.\d{4,9}/[^\s,;)]+` 会把它截断成 "10.1016/s0306-4530"（丢掉 (98)00014-6），
# 导致占位提取与整格匹配双双漏转。这里用"非括号段 或 (YY)x 年份段"来完整吃掉。
_DOI_TOKEN = r'10\.\d{4,9}/(?:[^\s,;；()"\']+|\([0-9]{2}[A-Za-z0-9]?\))+'
_PAREN_ANY = re.compile(r'[（(]((?:[^（）()]|\([^（）()]*\))*)[）)]')
_DOI_PART = re.compile(r'^(?:doi[:：]\s*)?(' + _DOI_TOKEN + r')[.;]?$', re.I)
_STANDALONE_DOI = re.compile(r'^\s*[（(]?\s*((?:doi[:：]\s*)?' + _DOI_TOKEN + r')\s*[）)]?\s*[.;]?\s*$', re.I)


def build_doi_matcher(item_mapping):
    """返回 doi2key: {doi(小写): ref_key}。mapping 需每项含 'doi'。"""
    return {str(v.get('doi', '')).lower(): k for k, v in item_mapping.items() if v.get('doi')}


def find_doi_citations(text, doi2key, text_map=None):
    """DOI 组括号引文 + 混合括号（作者-年份 + DOI）+ 精确文本括号引文。
    返回与 find_citations 同构的 [(start, end, display, [ref_key])]。

    - DOI 组：括号内以 ; 分隔的每个分片都是 DOI（doi: 前缀可有可无）；
      组内任一 DOI 不在 mapping -> 整组跳过（保守，避免半转换）。
    - 混合括号：括号内**非全 DOI**但包含至少一个可解析 DOI（如 "(Gao et al., 2014;
      doi:10.1016/...)"）。此时整括号转成引文域，citationItems = 括号内所有可解析
      DOI 对应条目；任一 DOI 不在 mapping -> 整组跳过。
    - text_map: {括号内文本(strip): [ref_key]}，如 {"FDA label, minocycline": ["fdamino"]}，
      用于无 DOI 的灰色文献占位（FDA 说明书 / NCT / EU CT / 公司通讯等）。"""
    out = []
    for m in _PAREN_ANY.finditer(text):
        inner = m.group(1)
        parts = [p.strip() for p in re.split(r'[;；]', inner)]
        if not parts or not all(parts):
            continue
        if all(_DOI_PART.match(p) for p in parts):
            keys = []
            for p in parts:
                d = _DOI_PART.match(p).group(1).lower()
                k = doi2key.get(d)
                if not k:
                    keys = None
                    break
                keys.append(k)
            if keys:
                out.append((m.start(), m.end(), m.group(0), keys))
            continue
        # 混合括号：非全 DOI 但含可解析 DOI（作者-年份 + DOI 同括号）
        # v3 安全约束：分片只要"看起来像 DOI"（以 10. / doi: 开头）却解析不出或
        # 查不到条目，整组跳过——否则 (10.x/known; 10.broken/z) 会被静默截断成
        # 只引 known 那条，另一条证据凭空消失（v2 实测的隐患）。
        mixed_keys = []
        for p in parts:
            mm = _DOI_PART.match(p)
            looks_doi = bool(re.match(r'^\s*(?:doi[:：]\s*)?10\.', p, re.I))
            if mm:
                k = doi2key.get(mm.group(1).lower())
                if not k:
                    mixed_keys = None
                    break
                mixed_keys.append(k)
            elif looks_doi:
                mixed_keys = None      # 形似 DOI 却无法解析 -> 整组放弃
                break
            # 非 DOI 分片（作者-年份等）忽略，但仍要求括号内至少含一个 DOI
        if mixed_keys:
            out.append((m.start(), m.end(), m.group(0), mixed_keys))
            continue
        if text_map and inner.strip() in text_map:
            out.append((m.start(), m.end(), m.group(0), list(text_map[inner.strip()])))
    return out


def match_standalone_doi(text, doi2key):
    """整段可见文本即一个 DOI（常见于表格的参考文献列）-> 返回 (ref_key, display) 或 None。
    display 保留原文（含可选的 doi: 前缀与尾部的 ; 或 .）。"""
    m = _STANDALONE_DOI.match(text or '')
    if not m:
        return None
    raw = m.group(1).lower()
    # 剥离可选的 doi: 前缀再查映射（build_doi_matcher 的 key 是纯 DOI，不带前缀）
    raw = re.sub(r'^doi[:：]\s*', '', raw, flags=re.I)
    k = doi2key.get(raw)
    if not k:
        return None
    return k, text


def _merge_spans(author_year_spans, other_spans):
    """合并两类匹配结果：other 中与 author_year 已有跨度重叠的丢弃，再按位置排序。"""
    spans = list(author_year_spans)
    for s in other_spans:
        s2_, e2 = s[0], s[1]
        if not any(not (e2 <= a or s2_ >= b) for a, b, _, _ in spans):
            spans.append(s)
    spans.sort()
    return spans


def build_combined_finder(by_year, nih_key, doi2key=None, text_map=None):
    """组合查找器：作者-年份 + DOI 组 + 精确文本括号 + 整段独立 DOI。

    insert_zotero_fields 与 verify 的零遗漏校验共用同一 finder，保证两侧口径一致
    （否则反之校验端只认作者-年份，会把 DOI 风格正文全部误报为"漏掉"）。"""
    def finder(text):
        spans = find_citations(text, by_year, nih_key)
        if doi2key:
            st = match_standalone_doi(text, doi2key)
            if st is not None:
                k, disp = st
                return _merge_spans(spans, [(0, len(text), disp, [k])])
            return _merge_spans(spans, find_doi_citations(text, doi2key, text_map))
        if text_map:
            return _merge_spans(spans, find_doi_citations(text, {}, text_map))
        return spans
    return finder


# ---------- Word 域代码 XML 构造 ----------
def _make_field_runs(instr, display, rpr):
    """返回 5 个 <w:r>：begin / instrText / separate / 显示文本 / end。"""
    runs = []
    for ftype in ('begin', 'instr', 'separate', 'text', 'end'):
        r = etree.Element(W + 'r')
        if rpr is not None: r.append(copy.deepcopy(rpr))
        if ftype in ('begin', 'separate', 'end'):
            etree.SubElement(r, W + 'fldChar').set(W + 'fldCharType', ftype)
        elif ftype == 'instr':
            it = etree.SubElement(r, W + 'instrText'); it.set(XMLSPACE, 'preserve')
            it.text = " ADDIN " + instr + " "
        else:  # text
            t = etree.SubElement(r, W + 't'); t.set(XMLSPACE, 'preserve'); t.text = display
        runs.append(r)
    return runs

def _run_text(r): return ''.join((t.text or '') for t in r.findall(W + 't'))

def _rpr_at(runs, positions, charpos):
    for (ri, s, e) in positions:
        if s <= charpos < e or charpos == e:
            return runs[ri].find(W + 'rPr')
    return None

def _slice_runs(runs, run_texts, positions, cs, ce):
    """从原 runs 切出 [cs,ce) 文字的 run(保留各 run 的 rPr)。"""
    out = []
    for (ri, s, e) in positions:
        if e <= cs or s >= ce: continue
        os_, oe = max(s, cs), min(e, ce)
        if oe <= os_: continue
        text = run_texts[ri][os_ - s:oe - s]
        if not text: continue
        r = etree.Element(W + 'r'); rpr = runs[ri].find(W + 'rPr')
        if rpr is not None: r.append(copy.deepcopy(rpr))
        t = etree.SubElement(r, W + 't'); t.set(XMLSPACE, 'preserve'); t.text = text
        out.append(r)
    return out


# ---------- 已有 Zotero 域代码检测（混合文档支持） ----------
def _detect_existing(body):
    """扫描 body，返回已存在的 Zotero 域信息。

    支持"既有域代码又有纯文字"的混合文档：已插入引文域的段落原样保留，
    bibliography/prefs 域已存在则不重复添加。
    按 begin..end 切分域并把域内所有 instrText 拼接后再判定 —— 兼容 Word/Zotero
    插件把一个域的指令拆成多个 <w:instrText> 的情况。
    返回 {"cited_paras": set(段落索引), "has_bibliography": bool, "has_preferences": bool}
    """
    paras = body.findall(W + 'p')
    has_bib = has_prefs = False
    cited = set()
    in_f = False
    cur = []
    start_para = -1
    for i, p in enumerate(paras):
        for r in p.iter(W + 'r'):
            for ch in r:
                tg = etree.QName(ch).localname
                if tg == 'fldChar':
                    ft = ch.get(W + 'fldCharType')
                    if ft == 'begin':
                        in_f = True; cur = []; start_para = i
                    elif ft == 'end' and in_f:
                        t = ''.join(cur)
                        if 'ZOTERO_ITEM CSL_CITATION' in t:
                            cited.add(start_para)
                        if 'ZOTERO_BIBL' in t:
                            has_bib = True
                        if 'ZOTERO_DOCUMENT_PREFERENCES' in t:
                            has_prefs = True
                        in_f = False; cur = []
                elif tg == 'instrText' and in_f:
                    cur.append(ch.text or '')
    return {"cited_paras": cited, "has_bibliography": has_bib, "has_preferences": has_prefs}


def detect_zotero_fields(docx_path):
    """外部入口：检测 docx 里已有的 Zotero 域。"""
    z = zipfile.ZipFile(docx_path)
    root = etree.fromstring(z.read('word/document.xml'))
    return _detect_existing(root.find(W + 'body'))


# ---------- 规范化域收集 / 域外文本 / 审计原语（v3） ----------
# 本节所有函数都以 **文档级** 顺序遍历（root.iter 顺序即文档顺序），域深度在段落之间
# 连续传递。历史上最容易出错的两个坑都源于"逐段重置深度"或"逐元素解析 JSON"：
#   1) Zotero 文献表域（BIBL）跨成百上千个段落，逐段重置深度会把文献表正文误判成
#      "域外裸 DOI"，虚增上百条假命中；域外残留扫描同样会误报。
#   2) Word/Zotero 会把一个域的指令拆成多个 <w:instrText>，必须域级拼接后再 json.loads。

def collect_fields(root, with_display=True):
    """域级收集文档中所有域。返回 [(instrText元素列表, 显示文本, 起始段落序号)]。

    - 深度跨段落连续传递：跨段域（BIBL/长 PREF）也能正确收口。
    - instrText 存 **元素**（不是字符串），便于就地重写（重编号/重链接）。
    - 显示文本 = begin..separate..end 之间的 w:t 拼接；无 separate 时为空串。
    """
    fields = []
    depth = 0; cur = None; disp = []; state = 'pre'; start_para = -1; para_i = -1
    for p in root.iter(W + 'p'):
        para_i += 1
        for r in p.iter(W + 'r'):
            for ch in r:
                tg = etree.QName(ch).localname
                if tg == 'fldChar':
                    ft = ch.get(W + 'fldCharType')
                    if ft == 'begin':
                        depth += 1
                        if depth == 1:
                            cur = []; disp = []; state = 'instr'; start_para = para_i
                    elif ft == 'separate':
                        if depth >= 1: state = 'disp'
                    elif ft == 'end':
                        if depth == 1 and cur is not None:
                            fields.append((cur, ''.join(disp) if with_display else '',
                                           start_para))
                            cur = None
                        depth = max(0, depth - 1)
                elif tg == 'instrText' and depth >= 1 and state == 'instr' and cur is not None:
                    cur.append(ch)
                elif tg == 't' and depth >= 1 and state == 'disp' and with_display:
                    disp.append(ch.text or '')
    return fields


def field_instr_text(els):
    """域级 instrText 拼接（跨 run / 跨 instrText 元素）。"""
    return ''.join(e.text or '' for e in els)


def field_payload(els):
    """解析域指令里的 JSON（CSL_CITATION / ZOTERO_BIBL / PREFERENCES 通用）。失败返回 None。"""
    t = field_instr_text(els)
    m = re.search(r'\{.*\}', t, re.S)
    if not m:
        return None
    try:
        return json.loads(m.group(0))
    except Exception:
        return None


def outside_field_paragraphs(root):
    """产出 (段落序号, 域外文本) —— 深度跨段连续传递，域内文本（引文缓存/文献表）不计。

    用途：找"真正的纯文字占位"。既有域的显示文本里常含 DOI 字样（显示文本=原文的
    转换结果），必须排除，否则会把已转换的引文误报成漏转。
    """
    depth = 0; para_i = -1
    for p in root.iter(W + 'p'):
        para_i += 1
        parts = []
        for r in p.iter(W + 'r'):
            for ch in r:
                tg = etree.QName(ch).localname
                if tg == 'fldChar':
                    ft = ch.get(W + 'fldCharType')
                    if ft == 'begin': depth += 1
                    elif ft == 'end': depth = max(0, depth - 1)
                elif tg == 't' and depth == 0:
                    parts.append(ch.text or '')
        yield para_i, ''.join(parts)


def _is_placeholder_doi(text):
    """域外文本里是否还留着 DOI 占位（doi: 前缀式或裸 DOI）。"""
    return bool(re.search(r'doi[:：]?\s*10\.\d{4,9}/', text or '', re.I)
                or re.search(r'(?<![\w/])10\.\d{4,9}/\S', re.sub(r'\([^()]*\)', '', text or '')))


def residual_placeholders(root):
    """域外残留的纯文字占位。返回 [{'para': i, 'text': str, 'kind': 'doi'|'bare'}]。

    跨段深度感知：跨段域（BIBL 文献表）内的 https://doi.org/… 不会被误报。
    """
    out = []
    for i, txt in outside_field_paragraphs(root):
        if not txt or '10.' not in txt:
            continue
        if re.search(r'doi[:：]?\s*10\.\d{4,9}/', txt, re.I):
            out.append({'para': i, 'text': txt[:120], 'kind': 'doi'})
            continue
        stripped = re.sub(r'\([^()]*\)', '', txt)
        m = re.search(r'(?<![\w/])10\.\d{4,9}/\S', stripped)
        if m:
            out.append({'para': i, 'text': stripped[max(0, m.start() - 60):m.end() + 20],
                        'kind': 'bare', 'doi': m.group(0)})
    return out


def document_item_keys(root):
    """文档现有引文域引用的全部 itemKey（用于库内存在性校验 / 死链审计）。"""
    keys = set()
    for els, _, _ in collect_fields(root):
        obj = field_payload(els)
        if not obj or 'citationItems' not in obj:
            continue
        for ci in obj.get('citationItems', []):
            for u in ci.get('uris', []):
                keys.add(str(u).rsplit('/items/', 1)[-1])
    return keys


def citation_id_stats(root):
    """citationID 统计：{'count','unique','duplicates':{id:n},'max_numeric'}。

    撞号是真实的坑：Zotero 用 citationID 索引文档内引文对象，两条不同引文共用一个
    ID 会让刷新把它们当成同一条。跨版本复制段落、手工合并文档都会造成编号世代重叠。
    """
    ids = []
    for els, _, _ in collect_fields(root):
        obj = field_payload(els)
        if not obj or 'citationItems' not in obj:
            continue
        ids.append(obj.get('citationID'))
    ids = [i for i in ids if i]
    counts = {}
    for i in ids:
        counts[i] = counts.get(i, 0) + 1
    dups = {k: v for k, v in counts.items() if v > 1}
    mx = -1
    for i in ids:
        m = re.fullmatch(r'cit(\d+)', str(i))
        if m:
            mx = max(mx, int(m.group(1)))
    return {'count': len(ids), 'unique': len(set(ids)), 'duplicates': dups,
            'max_numeric': mx, 'ids': ids}


def _citation_id_start(existing_ids, prefix='cit'):
    """新域 citationID 起始计数 = 既有同前缀数字 ID 的最大值 + 1。

    v1/v2 一律从 0 开始编号（cit0、cit1…），只要源文档已含引文域（无论来自本 skill
    的历史转换，还是用户在 Word 里刷新过、或跨版本复制段落带来的编号世代重叠），
    就必然撞号。这里让新域自动从安全起点编号。
    """
    mx = -1
    for i in existing_ids or []:
        m = re.fullmatch(re.escape(prefix) + r'(\d+)', str(i))
        if m:
            mx = max(mx, int(m.group(1)))
    return mx + 1


def _cell_has_citation(tc):
    """判断表格单元格（w:tc）是否已含引文域。同样拼接域内 instrText 以兼容 Word 拆分。"""
    in_f = False; cur = []
    for r in tc.iter(W + 'r'):
        for ch in r:
            tg = etree.QName(ch).localname
            if tg == 'fldChar':
                ft = ch.get(W + 'fldCharType')
                if ft == 'begin':
                    in_f = True; cur = []
                elif ft == 'end' and in_f:
                    if 'ZOTERO_ITEM CSL_CITATION' in ''.join(cur):
                        return True
                    in_f = False; cur = []
            elif tg == 'instrText' and in_f:
                cur.append(ch.text or '')
    return False


# ---------- 同段混合支持：段落内保留已有域，只转换纯文字引文 ----------
def _field_ranges(children):
    """返回 children 中构成 Word 域的连续子元素索引区间 [(start, end)]（含端）。

    按 fldChar begin..end 配对（兼容嵌套）。这些区间在重建段落时原样保留，
    只转换区间外的纯文字引文 —— 支持"同一段落内既有域代码又有纯文字"。
    """
    ranges = []
    i = 0
    n = len(children)
    while i < n:
        if etree.QName(children[i]).localname != 'r':
            i += 1
            continue
        if not any(etree.QName(c).localname == 'fldChar'
                   and c.get(W + 'fldCharType') == 'begin' for c in children[i]):
            i += 1
            continue
        depth = 0
        j = i
        end_at = None
        while j < n and etree.QName(children[j]).localname == 'r':
            for c in children[j]:
                if etree.QName(c).localname != 'fldChar':
                    continue
                ft = c.get(W + 'fldCharType')
                if ft == 'begin':
                    depth += 1
                elif ft == 'end':
                    depth -= 1
                    if depth == 0:
                        end_at = j
                        break
            if end_at is not None:
                break
            j += 1
        if end_at is not None:
            ranges.append((i, end_at))
            i = end_at + 1
        else:
            i += 1
    return ranges


def _convert_runs(runs, finder, uri_prefix, item_mapping, iid, csl_for, cit_n, cit_start=0):
    """把一段连续 run 中的纯文字引文转换为域。finder(text)->spans。返回 (新 children, cit_n)。"""
    run_texts = [_run_text(r) for r in runs]
    T = ''.join(run_texts)
    spans = finder(T)
    if not spans:
        return list(runs), cit_n
    positions = []
    pos = 0
    for ri, rt in enumerate(run_texts):
        positions.append((ri, pos, pos + len(rt)))
        pos += len(rt)
    new_children = []
    cursor = 0
    for (s, e, disp, items) in spans:
        if s > cursor:
            new_children += _slice_runs(runs, run_texts, positions, cursor, s)
        rpr = _rpr_at(runs, positions, s)
        cj = {"citationID": "cit" + str(cit_start + cit_n),
              "properties": {"formattedCitation": disp, "plainCitation": disp, "dontUpdate": False, "noteIndex": 0},
              "citationItems": [{"id": iid(k), "uris": [uri_prefix + item_mapping[k]['itemKey']],
                                 "itemData": csl_for(k)} for k in items],
              "schema": SCHEMA}
        new_children += _make_field_runs("ZOTERO_ITEM CSL_CITATION " + json.dumps(cj, ensure_ascii=False), disp, rpr)
        cit_n += 1
        cursor = e
    if cursor < len(T):
        new_children += _slice_runs(runs, run_texts, positions, cursor, len(T))
    return new_children, cit_n


def _process_para(p, finder, uri_prefix, item_mapping, iid, csl_for, cit_n, cit_start=0):
    """重建单个正文段落，支持同段混合：已有 Word/Zotero 域原样保留，只转换纯文字引文。

    hyperlink 与 pPr 等非 run 子元素原样保留（不转换超链接内的引文）。
    ⚠️ proofErr/lastRenderedPageBreak 是 Word 的拼写/渲染标记，会随文出现在任意两 run
    之间；若按普通元素 flush，会把一个括号引文组切成多段而整组漏转 -- 这里直接丢弃。
    书签不含文本，直通且不 flush（保留原位，避免同样切断文本块）。"""
    children = list(p)
    field_ranges = _field_ranges(children)
    new_children = []
    text_block = []

    def flush():
        nonlocal text_block, cit_n
        if not text_block:
            return
        converted, cit_n = _convert_runs(text_block, finder, uri_prefix,
                                         item_mapping, iid, csl_for, cit_n, cit_start)
        new_children.extend(converted)
        text_block = []

    fi = 0
    i = 0
    n = len(children)
    while i < n:
        if fi < len(field_ranges) and field_ranges[fi][0] == i:
            flush()
            start, end = field_ranges[fi]
            fi += 1
            new_children.extend(children[start:end + 1])
            i = end + 1
            continue
        el = children[i]
        tag = etree.QName(el).localname
        if tag == 'r':
            text_block.append(el)
        elif tag in ('proofErr', 'lastRenderedPageBreak'):
            pass  # Word 拼写/渲染标记：丢弃，不切断文本块（Word 会自行重建）
        elif tag in ('bookmarkStart', 'bookmarkEnd'):
            # 书签不含文本但需保持文档顺序：先 flush 再原位保留。
            # （若书签恰好落在一个引文组中间会切断文本块--罕见；Word 的 _GoBack
            # 类书签几乎总在段落边界，主流场景不受影响。）
            flush()
            new_children.append(el)
        elif tag == 'hyperlink':
            flush()
            new_children.append(el)
        else:  # pPr / w:ins 等
            flush()
            new_children.append(el)
        i += 1
    flush()
    for ch in list(p):
        p.remove(ch)
    for el in new_children:
        p.append(el)
    return cit_n


# ---------- 主处理（正文引文 + 参考文献表 + 首选项） ----------
def insert_zotero_fields(src_docx, out_docx, body_para_range, ref_para_indices,
                         item_mapping, uri_prefix=None, style_id="http://www.zotero.org/styles/apa",
                         style_locale="en-US", zotero_version="9.0.6", skip_existing=True,
                         citation_text_map=None, include_tables=False):
    """
    src_docx: 原始 docx。若含已有 Zotero 域代码（混合文档），skip_existing=True 时
        已插入引文域的段落原样保留，bibliography/prefs 域已存在则不重复添加。
    body_para_range: (start, end) 正文段落索引范围（end 不含），引文在这里替换
    ref_para_indices: 参考文献表段落索引列表，用参考文献表域包裹
    item_mapping: {ref_key: {itemKey, uri, csl, doi?, ...}}。含 'doi' 的条目自动启用
        DOI 组匹配（"(doi:10.x)" / "(10.x; 10.y)" 占位风格正文），与作者-年份匹配并存。
    citation_text_map: 可选 {括号内文本(strip): [ref_key]}，精确匹配无 DOI 的灰色文献
        占位（如 {"FDA label, minocycline": ["fdaminocycline"], "NCT01329991": ["nct1"]}）。
    include_tables: True 时同时处理表格内段落（含"整格即一个 DOI"的参考文献列，
        整格替换为引文域）。默认 False 保持正文段落索引语义不变。
    uri_prefix: None 时用 config.get_uri_prefix()（http://zotero.org/users/<userID>/items/）
    """
    uri_prefix = _resolve_uri_prefix(uri_prefix)
    by_year, nih = build_matcher(item_mapping)
    doi2key = build_doi_matcher(item_mapping)
    finder = build_combined_finder(by_year, nih, doi2key or None, citation_text_map)
    z = zipfile.ZipFile(src_docx); root = etree.fromstring(z.read('word/document.xml'))
    body = root.find(W + 'body'); paras = body.findall(W + 'p')
    existing = _detect_existing(body) if skip_existing else \
        {"cited_paras": set(), "has_bibliography": False, "has_preferences": False}
    item_ids = {}
    def iid(k):
        if k not in item_ids: item_ids[k] = len(item_ids) + 1
        return item_ids[k]
    def csl_for(k):
        c = copy.deepcopy(item_mapping[k]['csl']); c['id'] = iid(k); c.pop('abstract', None); return c

    def _para_text(p):
        return ''.join((t.text or '') for t in p.iter(W + 't'))

    def _standalone_cell(p):
        """整段即一个 DOI -> 整段替换为引文域。返回是否处理。"""
        if not doi2key:
            return False
        hit = match_standalone_doi(_para_text(p), doi2key)
        if not hit:
            return False
        k, disp = hit
        cj = {"citationID": "cit" + str(cit_start + cit_holder[0]),
              "properties": {"formattedCitation": disp, "plainCitation": disp, "dontUpdate": False, "noteIndex": 0},
              "citationItems": [{"id": iid(k), "uris": [uri_prefix + item_mapping[k]['itemKey']],
                                 "itemData": csl_for(k)}],
              "schema": SCHEMA}
        runs = p.findall(W + 'r')
        rpr = runs[0].find(W + 'rPr') if runs else None
        fld = _make_field_runs("ZOTERO_ITEM CSL_CITATION " + json.dumps(cj, ensure_ascii=False), disp, rpr)
        idx0 = next((j for j, ch in enumerate(p) if etree.QName(ch).localname == 'r'), len(list(p)))
        for ch in list(p):
            if etree.QName(ch).localname in ('r', 'proofErr'):
                p.remove(ch)
        for j, el in enumerate(fld):
            p.insert(idx0 + j, el)
        cit_holder[0] += 1
        return True

    # v3：新域 citationID 自动避开文档中已存在的编号（否则跨版本/复制段落必撞号）。
    # cit_start 只是编号偏移；cit_holder 仍从 0 计新增域数，函数返回值语义不变。
    cit_start = _citation_id_start(
        [obj.get('citationID') for els, _, _ in collect_fields(root) if (obj := field_payload(els))])
    cit_holder = [0]

    b0, b1 = body_para_range
    # 1) 正文引文域（支持同段混合：保留已有 Word/Zotero 域，只转换纯文字引文）
    for idx in range(b0, b1):
        cit_holder[0] = _process_para(paras[idx], finder, uri_prefix, item_mapping,
                                     iid, csl_for, cit_holder[0], cit_start)

    # 1b) 表格内段落（可选）：独立 DOI 格整格替换；其余走同一段落转换
    if include_tables:
        for tbl in body.iter(W + 'tbl'):
            for p in tbl.iter(W + 'p'):
                if _standalone_cell(p):
                    continue
                cit_holder[0] = _process_para(p, finder, uri_prefix, item_mapping,
                                             iid, csl_for, cit_holder[0], cit_start)

    cit_n = cit_holder[0]

    # 2) 参考文献表域（包裹 ref 段落；已存在则不重复添加）
    bib_n = len(ref_para_indices) if (ref_para_indices and not existing['has_bibliography']) else 0
    if ref_para_indices and not existing['has_bibliography']:
        fp = paras[ref_para_indices[0]]; lp = paras[ref_para_indices[-1]]
        bib = {"uncited": [[uri_prefix + v['itemKey']] for v in item_mapping.values() if v.get('itemKey')],
               "omitted": [], "custom": []}
        instr = "ZOTERO_BIBL " + json.dumps(bib, ensure_ascii=False) + " CSL_BIBLIOGRAPHY"
        bid = str(90000 + random.randint(0, 9999))
        bms = etree.Element(W + 'bookmarkStart'); bms.set(W + 'id', bid); bms.set(W + 'name', 'ZOTERO_BREF')
        bme = etree.Element(W + 'bookmarkEnd'); bme.set(W + 'id', bid)
        fr = fp.findall(W + 'r'); rpr = fr[0].find(W + 'rPr') if fr else None
        fld = _make_field_runs(instr, "", rpr)
        idx0 = next((i for i, ch in enumerate(fp) if etree.QName(ch).localname == 'r'), 0)
        for j, el in enumerate([bms] + fld[:3]): fp.insert(idx0 + j, el)
        re_ = etree.Element(W + 'r')
        if rpr is not None: re_.append(copy.deepcopy(rpr))
        etree.SubElement(re_, W + 'fldChar').set(W + 'fldCharType', 'end')
        lp.append(re_); lp.append(bme)

    # 3) 文档首选项域（隐藏段落，文首；已存在则不重复添加）
    if not existing['has_preferences']:
        prefs = {"style": {"styleID": style_id, "locale": style_locale, "hasBibliography": True, "bibliographyStyleHasBeenSet": False},
                 "prefs": {"fieldType": "Field", "automaticJournalAbbreviations": False},
                 "sessionID": "sess" + ''.join(random.choices(string.ascii_lowercase, k=8)),
                 "zoteroVersion": zotero_version, "dataVersion": 4}
        pp = etree.Element(W + 'p'); pPr = etree.SubElement(pp, W + 'pPr')
        sp = etree.SubElement(pPr, W + 'spacing'); sp.set(W + 'before', '0'); sp.set(W + 'after', '0'); sp.set(W + 'line', '20'); sp.set(W + 'lineRule', 'exact')
        etree.SubElement(etree.SubElement(pPr, W + 'rPr'), W + 'vanish')
        rpr_h = etree.Element(W + 'rPr'); etree.SubElement(rpr_h, W + 'vanish')
        sz = etree.SubElement(rpr_h, W + 'sz'); sz.set(W + 'val', '2')
        bid2 = str(92000 + random.randint(0, 9999))
        bms2 = etree.SubElement(pp, W + 'bookmarkStart'); bms2.set(W + 'id', bid2); bms2.set(W + 'name', 'ZOTERO_PREF')
        for r in _make_field_runs("ZOTERO_DOCUMENT_PREFERENCES " + json.dumps(prefs, ensure_ascii=False), "", rpr_h): pp.append(r)
        etree.SubElement(pp, W + 'bookmarkEnd').set(W + 'id', bid2)
        body.insert(0, pp)

    new_xml = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
    with zipfile.ZipFile(src_docx, 'r') as zin, zipfile.ZipFile(out_docx, 'w', zipfile.ZIP_DEFLATED) as zout:
        for it in zin.infolist():
            zout.writestr(it, new_xml if it.filename == 'word/document.xml' else zin.read(it.filename))
    return cit_n, bib_n


def insert_table_citations(src_docx, out_docx, item_mapping, doi_column, refnum_column,
                           uri_prefix=None, style_id="http://www.zotero.org/styles/apa",
                           style_locale="en-US", zotero_version="9.0.6", skip_existing=True):
    """
    表格类文档：把引文域放到"编号/Ref"列(refnum_column)，作者列保持纯文字。
    doi_column: DOI 所在列(用于定位 item)；refnum_column: 要写入引文域的列(如 Ref# 列)；
        refnum_column=None 时直接替换 DOI 格自身（适用于"整列就是参考文献"的表格）。
    item_mapping 需能按 doi 查: 先建 doi2key = {v['doi'].lower(): k}。
    uri_prefix: None 时用 config.get_uri_prefix()。
    skip_existing=True 时检测已有域，prefs 已存在则不重复添加。
    """
    uri_prefix = _resolve_uri_prefix(uri_prefix)
    doi2key = {v.get('doi', '').lower(): k for k, v in item_mapping.items() if v.get('doi')}
    z = zipfile.ZipFile(src_docx); root = etree.fromstring(z.read('word/document.xml'))
    body = root.find(W + 'body'); tbl = root.find('.//' + W + 'tbl')
    if tbl is None:
        return 0  # 文档无表格，直接返回
    rows = tbl.findall(W + 'tr')
    existing = _detect_existing(body) if skip_existing else \
        {"cited_paras": set(), "has_bibliography": False, "has_preferences": False}
    item_ids = {}
    def iid(k):
        if k not in item_ids: item_ids[k] = len(item_ids) + 1
        return item_ids[k]
    def csl_for(k):
        c = copy.deepcopy(item_mapping[k]['csl']); c['id'] = iid(k); c.pop('abstract', None); return c
    # v3：与 insert_zotero_fields 同样自动避开既有 tcitN 编号（cit_n 仍是新增域数）
    cit_start = _citation_id_start(
        [obj.get('citationID') for els, _, _ in collect_fields(root) if (obj := field_payload(els))],
        prefix='tcit')
    cit_n = 0
    target_col = refnum_column if refnum_column is not None else doi_column
    for ri in range(1, len(rows)):
        tcs = rows[ri].findall(W + 'tc')
        if len(tcs) <= max(doi_column, target_col): continue
        doi = ''.join((t.text or '') for t in tcs[doi_column].iter(W + 't')).strip().lower().rstrip(';.')
        key = doi2key.get(doi)
        if not key:
            for d2, k2 in doi2key.items():
                if doi and (doi in d2 or d2 in doi): key = k2; break
        if not key: continue
        ref_tc = tcs[target_col]; ref_text = ''.join((t.text or '') for t in ref_tc.iter(W + 't')).strip()
        p = ref_tc.find(W + 'p')
        if p is None: continue
        # 该 Ref 单元格已含引文域则跳过（混合文档）
        if skip_existing and _cell_has_citation(ref_tc): continue
        runs = p.findall(W + 'r'); rpr = runs[0].find(W + 'rPr') if runs else None
        cj = {"citationID": "tcit" + str(cit_start + cit_n),
              "properties": {"formattedCitation": ref_text, "plainCitation": ref_text, "dontUpdate": False, "noteIndex": 0},
              "citationItems": [{"id": iid(key), "uris": [uri_prefix + item_mapping[key]['itemKey']],
                                 "itemData": csl_for(key)}], "schema": SCHEMA}
        fld = _make_field_runs("ZOTERO_ITEM CSL_CITATION " + json.dumps(cj, ensure_ascii=False), ref_text, rpr)
        for ch in list(p):
            if etree.QName(ch).localname == 'r': p.remove(ch)
        for fr in fld: p.append(fr)
        cit_n += 1
    # 首选项域（已存在则不重复添加）
    if not existing['has_preferences']:
        prefs = {"style": {"styleID": style_id, "locale": style_locale, "hasBibliography": True, "bibliographyStyleHasBeenSet": False},
                 "prefs": {"fieldType": "Field", "automaticJournalAbbreviations": False},
                 "sessionID": "sess" + ''.join(random.choices(string.ascii_lowercase, k=8)),
                 "zoteroVersion": zotero_version, "dataVersion": 4}
        pp = etree.Element(W + 'p'); pPr = etree.SubElement(pp, W + 'pPr')
        sp = etree.SubElement(pPr, W + 'spacing'); sp.set(W + 'before', '0'); sp.set(W + 'after', '0'); sp.set(W + 'line', '20'); sp.set(W + 'lineRule', 'exact')
        etree.SubElement(etree.SubElement(pPr, W + 'rPr'), W + 'vanish')
        rpr_h = etree.Element(W + 'rPr'); etree.SubElement(rpr_h, W + 'vanish')
        etree.SubElement(rpr_h, W + 'sz').set(W + 'val', '2')
        bid = str(93000 + random.randint(0, 9999))
        bms = etree.SubElement(pp, W + 'bookmarkStart'); bms.set(W + 'id', bid); bms.set(W + 'name', 'ZOTERO_PREF')
        for r in _make_field_runs("ZOTERO_DOCUMENT_PREFERENCES " + json.dumps(prefs, ensure_ascii=False), "", rpr_h): pp.append(r)
        etree.SubElement(pp, W + 'bookmarkEnd').set(W + 'id', bid)
        body.insert(0, pp)
    new_xml = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
    with zipfile.ZipFile(src_docx, 'r') as zin, zipfile.ZipFile(out_docx, 'w', zipfile.ZIP_DEFLATED) as zout:
        for it in zin.infolist():
            zout.writestr(it, new_xml if it.filename == 'word/document.xml' else zin.read(it.filename))
    return cit_n


# ---------- 校验 ----------
def _extract_citation_fields(root):
    """提取输出文档里所有引文域的 (显示文本, formattedCitation)。"""
    fields = []
    in_f = in_disp = False
    cur = []
    disp = []
    for r in root.iter(W + 'r'):
        for ch in r:
            tg = etree.QName(ch).localname
            if tg == 'fldChar':
                ft = ch.get(W + 'fldCharType')
                if ft == 'begin':
                    in_f = True; in_disp = False; cur = []; disp = []
                elif ft == 'separate':
                    in_disp = True
                elif ft == 'end' and in_f:
                    t = ''.join(cur)
                    if 'CSL_CITATION' in t:
                        try:
                            obj = json.loads(t.split('CSL_CITATION', 1)[1].strip())
                            fmt = obj.get('properties', {}).get('formattedCitation', '')
                        except Exception:
                            fmt = ''
                        fields.append((''.join(disp), fmt))
                    in_f = in_disp = False
            elif tg == 'instrText' and in_f and not in_disp:
                cur.append(ch.text or '')
            elif tg == 't' and in_f and in_disp:
                disp.append(ch.text or '')
    return fields


def _verify_zero_miss(src_docx, out_docx, item_mapping, citation_text_map=None):
    """比对原文与输出：确保每个引文都被转换（零遗漏），且显示文本逐字一致。

    ⚠️ 必须复用 insert 用的同一 combined finder（作者-年份 + DOI 组 + 精确文本括号 +
    整段独立 DOI）：否则 DOI 风格正文会被一律误报为"漏掉"。"""
    by_year, nih = build_matcher(item_mapping)
    doi2key = build_doi_matcher(item_mapping)
    finder = build_combined_finder(by_year, nih, doi2key or None, citation_text_map)
    zsrc = zipfile.ZipFile(src_docx)
    rsrc = etree.fromstring(zsrc.read('word/document.xml'))
    expected = []
    body = rsrc.find(W + 'body')
    for p in body.findall(W + 'p'):
        text = ''.join((t.text or '') for t in p.iter(W + 't'))
        for (_s, _e, disp, _it) in finder(text):
            expected.append(disp)
    for t in body.iter(W + 'tbl'):
        for p in t.iter(W + 'p'):
            text = ''.join((tt.text or '') for tt in p.iter(W + 't'))
            for (_s, _e, disp, _it) in finder(text):
                expected.append(disp)
    zout = zipfile.ZipFile(out_docx)
    rout = etree.fromstring(zout.read('word/document.xml'))
    fields = _extract_citation_fields(rout)
    actual = []
    for (disp, fmt) in fields:
        if disp:
            actual.append(disp)
        if fmt:
            actual.append(fmt)
    # 显示文本与 formattedCitation 可能一方是字面非 ASCII、一方是 Zotero 的 RTF 转义
    # （如 en-dash: "-" 字面 vs "\uc0舑{}"）；比较前统一解码，避免误报 mismatch。
    an = set(_norm(_rtf_unescape(d)) for d in actual if d)
    # 零遗漏 + 显示文本与原文一致：每个期望引文都应出现在某域（显示文本或 formattedCitation）
    missed = [e for e in expected if _norm(e) not in an]
    # 域内一致性：显示文本应等于其 formattedCitation
    mism = [{"display": d, "formatted": f}
            for (d, f) in fields
            if f and _norm(_rtf_unescape(d)) != _norm(_rtf_unescape(f))]
    return missed, mism


# 兼容 \uc0舑{} 与裸 舑{} 两种前缀
_RTF_ESC = re.compile(r'\\uc0\\u(-?\d+)\{\}|\\u(-?\d+)\{\}')


def _rtf_unescape(s):
    """把 Zotero Word 插件写进 JSON 的 RTF Unicode 转义（\\uc0\\uN{}，N 为**十进制**码点）
    还原为字符，供显示文本/formattedCitation 逐字比较用。
    例：\\u8211{} = 十进制 8211 = 0x2013 = en-dash（不可按十六进制解，那是另一个码点）。"""
    if not s:
        return s
    def _sub(m):
        g = m.group(1) or m.group(2)
        return chr(int(g, 10) & 0xFFFF)
    return _RTF_ESC.sub(_sub, s)


def verify(out_docx, valid_item_keys, src_docx=None, item_mapping=None,
           citation_text_map=None, deep=False):
    """返回 (ok, 详情)。结构校验：fldChar 平衡、JSON 合法、URI 有效、CSL 完整。

    若同时给 src_docx 与 item_mapping，则额外做"引文零遗漏 + 显示文本逐字一致"比对：
    用原文的引文匹配器（作者-年份 + DOI 组 + 文本括号，与 insert 共用）找出所有应转换的
    引文，逐一核对是否都出现在输出域的显示文本里。citation_text_map 需与调用
    insert_zotero_fields 传的一致，否则 DOI/文本占位引用会被误报为漏掉。
    注意：一个域的指令可能被 Word 拆成多个 <w:instrText>，须把同一域内所有 instrText 拼接后再解析。

    deep=True（v3 新增）再叠加三项交付前必须过的检查：
      - citationID 唯一性（撞号会让 Zotero 刷新把两条引文当成同一条）
      - 域外残留纯文字占位（跨段深度感知，不会被 BIBL 文献表里的 DOI 误报）
      - 域计数明细（ITEM / BIBL / PREF）
    这三项进 detail，且影响 ok。
    """
    z = zipfile.ZipFile(out_docx); root = etree.fromstring(z.read('word/document.xml'))
    # 按 begin..end 切分域，拼接每个域内的所有 instrText（深度跨段连续传递）
    fields = []; depth = 0; cur = None
    for r in root.iter(W + 'r'):
        for ch in r:
            tg = etree.QName(ch).localname
            if tg == 'fldChar':
                ft = ch.get(W + 'fldCharType')
                if ft == 'begin': depth += 1; cur = [] if depth == 1 else cur
                elif ft == 'end':
                    if depth == 1 and cur is not None:
                        fields.append(''.join(cur)); cur = None
                    depth = max(0, depth - 1)
            elif tg == 'instrText' and depth >= 1 and cur is not None: cur.append(ch.text or '')
    cites = bibs = prefs = 0; bad_json = bad_uri = incomplete = with_abstract = 0
    for t in fields:
        if 'ZOTERO_ITEM CSL_CITATION' in t:
            cites += 1
            try: obj = json.loads(t.split('CSL_CITATION', 1)[1].strip())
            except: bad_json += 1; continue
            for ci in obj.get('citationItems', []):
                if not ci.get('uris') or ci['uris'][0].split('/items/')[-1] not in valid_item_keys:
                    bad_uri += 1
                elif any(k not in ci.get('itemData', {}) for k in ('type', 'title', 'author', 'issued')):
                    incomplete += 1
                elif 'abstract' in ci.get('itemData', {}):
                    # ⚠️ 真实 Zotero 插件生成的域**经常内嵌 abstract**（Word 插件刷新时带上），
                    # 不算缺陷：只计数提示，不计入失败。新写域仍应剔除（生成器已做）。
                    with_abstract += 1
        elif 'ZOTERO_BIBL' in t: bibs += 1
        elif 'ZOTERO_DOCUMENT_PREFERENCES' in t: prefs += 1
    bg = len([1 for f in root.iter(W + 'fldChar') if f.get(W + 'fldCharType') == 'begin'])
    ed = len([1 for f in root.iter(W + 'fldChar') if f.get(W + 'fldCharType') == 'end'])
    ok = bg == ed and bad_json == 0 and bad_uri == 0 and incomplete == 0
    detail = dict(cites=cites, bibs=bibs, prefs=prefs, fldChar=f"{bg}/{ed}",
                  bad_json=bad_json, bad_uri=bad_uri, incomplete=incomplete,
                  with_abstract=with_abstract)
    if src_docx and item_mapping:
        missed, mism = _verify_zero_miss(src_docx, out_docx, item_mapping, citation_text_map)
        detail["missed_citations"] = missed
        detail["display_mismatches"] = mism
        ok = ok and not missed and not mism
    if deep:
        stats = citation_id_stats(root)
        residual = residual_placeholders(root)
        detail["citation_ids"] = {k: v for k, v in stats.items() if k != 'ids'}
        detail["residual_placeholders"] = residual
        ok = ok and not stats['duplicates'] and not residual
    return ok, detail


if __name__ == '__main__':
    # 示例：见 SKILL.md。这里仅自检导入。
    print("逆向 Zotero 域代码插入器已就绪。用法见 SKILL.md。")
    print("函数: insert_zotero_fields(正文+参考文献表), insert_table_citations(表格), verify(校验)")
