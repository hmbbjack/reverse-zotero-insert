#!/usr/bin/env python3
"""
逆向 Zotero 插入 —— 前置引用摄取阶段（Front-End Reference Ingestion）
====================================================================
把 .docx 文稿里的纯文字编号参考文献（如 "[1] Smith J. Title. Journal. 2007."）
解析出元数据 → 联网（CrossRef/PubMed 搜索 API）核实 → 导入 Zotero（经 MCP）→
构建下游生成器需要的 item_mapping.json。

本模块是"纯管道"：不做任何用户交互。它只产出决策数据（resolution_summary.json、
dedup_report.json），由 ClauAu/skill 在强制检查点里呈现给用户，用户确认后才导入。

依赖: python-docx, requests, lxml  (pip install python-docx requests lxml)
与下游: 产出与 zotero_field_insert.py 逐字节兼容的 item_mapping.json；复用
zotero_mcp.call_tool（作为 zotero callable 传入，默认 zotero_mcp.call_tool）。

工作流（skill 交互态由 Claude 分步调用，并在 resolve/dedup 与 import 之间插入
强制检查点；本模块的 ingest() 是非交互兜底）:
  check_zotero_mcp -> extract_reference_paragraphs -> resolve_metadata
  -> dedup_check -> [强制检查点] -> import_to_zotero -> build_item_mapping
"""
import json
import html
import re
import time
import random
import unicodedata
from urllib.parse import quote

try:
    import config
except Exception:
    config = None

REF_KEY_MIN_CONFIDENCE = 0.6  # 低于此值标 LOW，需人工/Claude 复核

# 常见期刊/通用引用格式 -> Zotero CSL styleID（未覆盖的由 Claude 联网搜索）
_STYLE_ALIASES = {
    "apa": "http://www.zotero.org/styles/apa",
    "ieee": "http://www.zotero.org/styles/ieee",
    "nature": "http://www.zotero.org/styles/nature",
    "vancouver": "http://www.zotero.org/styles/vancouver",
    "chicago": "http://www.zotero.org/styles/chicago-author-date",
    "harvard": "http://www.zotero.org/styles/harvard-cite-them-right",
    "mla": "http://www.zotero.org/styles/modern-language-association",
    "american medical association": "http://www.zotero.org/styles/american-medical-association",
    "gb7714": "http://www.zotero.org/styles/gb7714-2015-numeric",
    "gb-t 7714": "http://www.zotero.org/styles/gb7714-2015-numeric",
    "国标": "http://www.zotero.org/styles/gb7714-2015-numeric",
    "numeric": "http://www.zotero.org/styles/ieee",
}


def resolve_style_id(name_or_journal):
    """把期刊名 / 通用格式名解析为 Zotero CSL styleID。

    命中内置别名表则直接返回；否则返回 (None, "en-US")，由 Claude 联网搜索确认
    该期刊对应的 Zotero CSL 样式后，把 style_id 传给 insert_zotero_fields。
    """
    if not name_or_journal:
        return "http://www.zotero.org/styles/apa", "en-US"
    s = _norm(name_or_journal)
    norm_aliases = {_norm(k): v for k, v in _STYLE_ALIASES.items()}
    if s in norm_aliases:
        return norm_aliases[s], "en-US"
    for k, v in norm_aliases.items():
        if k and k in s:
            return v, "en-US"
    return None, "en-US"


# ---------- 基础工具 ----------
def _norm(s):
    if not s:
        return ""
    # 保留 CJK（中文文献），NFKD 折叠拉丁重音，去标点/空白
    s = unicodedata.normalize('NFKD', s).lower()
    return re.sub(r'[^a-z0-9一-鿿]', '', s)


def _is_cjk(s):
    return any('一' <= ch <= '鿿' for ch in s)


def _title_words(s):
    """把标题拆成规范化单词集合，用于词级重叠。

    拉丁词按词；中文按整段 + 二元组（bigram），使"儿童自闭症"与"儿童自闭症干预研究"
    能在词级有重叠，让置信度对中文标题也有效。
    """
    if not s:
        return set()
    t = unicodedata.normalize('NFKD', str(s)).lower()
    words = set(re.findall(r'[a-z]+', t))
    for run in re.findall(r'[一-鿿]+', t):
        words.add(run)
        for i in range(len(run) - 1):
            words.add(run[i:i + 2])
    return words


def _first_author_lastname(csl):
    au = csl.get('author') or []
    if au:
        if au[0].get('family'):
            return au[0]['family']
        if au[0].get('literal'):
            return au[0]['literal']
    return ""


def _year(csl):
    dp = csl.get('issued', {}).get('date-parts', [[None]])
    if dp and dp[0] and dp[0][0]:
        return str(dp[0][0])
    return ""


def _write_json(path, obj):
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(obj, f, ensure_ascii=False, indent=2)


def _get_zotero(zotero):
    if zotero is not None:
        return zotero
    import zotero_mcp
    return zotero_mcp.call_tool


def _make_base_ref_key(surname, year):
    """生成 ref_key 基础段，如 'smith2007'；冲突后缀在 extract 层处理。"""
    s = _norm(surname) or 'anon'
    y = _norm(year) or 'nodate'
    return f"{s}{y}"


# ---------- 前置 MCP 配置检查 ----------
def check_zotero_mcp(zotero=None, url=None):
    """探测 Zotero MCP 是否就绪。返回 {ok, configured, error, hint}。

    未就绪时，SKILL/claude 应据此提醒用户"Zotero MCP 未配置好"并给出修复方式。
    """
    zotero = _get_zotero(zotero)
    # 用 get_collections 探测更稳妥（部分 MCP 实现无 get_libraries）；get_libraries 失败则回退
    try:
        res = zotero("get_libraries", {})
    except Exception:
        try:
            res = zotero("get_collections", {})
        except Exception as e:  # 连接失败 / 超时
            return {"ok": False, "configured": False, "error": str(e),
                    "hint": "无法连接 Zotero MCP。请确认：① Zotero 正在运行且已启用 MCP 连接器"
                            "（replication 插件/连接器监听 127.0.0.1:23120）；"
                            "② 或设置 ZOTERO_MCP_URL 环境变量指向你的 MCP 服务。"}
    if isinstance(res, dict) and res.get("error"):
        return {"ok": False, "configured": False, "error": res["error"],
                "hint": "Zotero MCP 返回错误。请确认连接器已启用、Zotero 正在运行，"
                        "或设置 ZOTERO_MCP_URL。"}
    return {"ok": True, "configured": True, "detail": res}


# ---------- docx 解析 ----------
def _detect_ref_paras(paras):
    """自动检测编号参考文献块：段首匹配 [n] / (n) / n. / n) 。

    返回连续编号段；允许编号段之间夹空段（容忍个别空行），但遇到非空、非编号段则停止。
    """
    pat = re.compile(r'^\s*(?:\[[0-9]+\]|\([0-9]+\)|[0-9]+[\.\)]\s?)')
    idx = [i for i, p in enumerate(paras) if pat.match(p.text or "")]
    if not idx:
        return []
    out = [idx[0]]
    for i in idx[1:]:
        prev = out[-1]
        # prev 到 i 之间允许全为空段
        if all(not (paras[k].text or "").strip() for k in range(prev + 1, i)):
            out.append(i)
        else:
            break
    return out


def extract_reference_paragraphs(docx_path, ref_para_indices=None, auto_detect=True):
    """返回 [(para_index, ref_key, raw_text)]。

    ref_para_indices 给定则用之；否则 auto_detect 自动检测编号块。
    ref_key 如 'smith2007'，作者+年份冲突时加 _2/_3。
    """
    from docx import Document
    doc = Document(docx_path)
    paras = doc.paragraphs
    if ref_para_indices is None:
        indices = _detect_ref_paras(paras) if auto_detect else []
    else:
        indices = list(ref_para_indices)
    used = {}
    out = []
    for i in indices:
        text = (paras[i].text or "").strip()
        if not text:
            continue
        pr = parse_numbered_reference(text)
        if not pr:
            continue
        base = pr["base_key"]
        used[base] = used.get(base, 0) + 1
        key = base if used[base] == 1 else f"{base}_{used[base]}"
        pr["ref_key"] = key
        out.append((i, key, text))
    return out


def parse_numbered_reference(text):
    """解析一条编号参考文献，返回 dict 或 None。

    支持引导符 [n]/(n)/n./n)；年份 (YYYY) 任意位置；doi:/PMID: 标记。
    字段划分是启发式（best-effort）——真正的元数据核实交给 CrossRef/PubMed。
    """
    raw = (text or "").strip()
    if not raw:
        return None
    t = re.sub(r'^[\(\[\{]?\s*\d+\s*[\)\.\}\]]?\s*', '', raw)  # 去掉前导编号
    year_m = re.search(r'\b(19|20)\d{2}\b', t)
    year = year_m.group(0) if year_m else None
    doi_m = re.search(r'(?:doi[:/]\s*)?(10\.\d{4,}/[^\s;,)]+)', t, re.I)
    doi = doi_m.group(1).rstrip('.,') if doi_m else None
    pmid_m = re.search(r'PMID[:#]?\s*(\d+)', t, re.I)
    pmid = pmid_m.group(1) if pmid_m else None

    clauses = [c for c in re.split(r'\.\s+', t) if c.strip()]

    def _pick_journal():
        """从作者/标题之后的从句里挑期刊：跳过纯 DOI/PMID 段与年份段。"""
        for cl in reversed(clauses[1:]):
            if re.match(r'^(?:doi[:/]|pmid[:#]?\s*\d)', cl, re.I):
                continue
            if re.match(r'^\s*(?:19|20)\d{2}\b', cl):
                continue
            return cl
        return None

    authors = []
    title = None
    journal = None
    if clauses:
        first = clauses[0]
        # 判定首段是否为作者部分：含逗号/分号（多作者），或含 and/&（"Smith J and Johnson K"），
        # 或短从句（≤3 词）以单/多字母首字母结尾（"Smith JD" / "Smith J"）。
        # 或为 2-4 字中文人名（"张三"）。
        # 避免把 "A study of treatment outcomes" / "Risk factors for Y" 这类标题误判为作者。
        n_sep = first.count(',') + first.count(';')
        has_and = bool(re.search(r'\s+(?:and|&)\s+', first))
        n_tok = len(first.split())
        has_initials = (n_sep > 0 or has_and
                        or (n_tok <= 3 and re.search(r'\s+[A-Z][A-Z.]{0,2}[.,]?$', first)))
        has_cjk_author = bool(re.fullmatch(r'[一-鿿]{2,4}', first))
        has_year = bool(re.search(r'\b(19|20)\d{2}\b', first))
        if (has_initials or has_cjk_author) and not has_year:
            # 首段是作者部分："Smith J, Johnson K" / "Smith JD" / "张三"
            for part in re.split(r'\s*[;,]\s*|\s+and\s+', first):
                part = part.strip()
                if not part:
                    continue
                if re.fullmatch(r'(?:et\s*al)\.?', part, re.I):
                    continue  # 跳过 "et al"，不是作者
                m = re.match(r'^(.+?)\s+([A-ZÀ-Ý][A-ZÀ-Ý.]*(?:\s+[A-ZÀ-Ý][A-ZÀ-Ý.]*)*)$', part)
                if m:
                    family = m.group(1).strip()
                    given = m.group(2).strip().rstrip('.,')
                    authors.append({"family": family, "given": given})
                elif _is_cjk(part):
                    # 中文人名整体作为姓（查询用；真正匹配由 CSL 的 family 驱动）
                    authors.append({"family": part, "given": ""})
                else:
                    authors.append({"family": part, "given": ""})
            if len(clauses) >= 2:
                title = clauses[1]
            if len(clauses) >= 3:
                journal = _pick_journal()
        else:
            title = first
            if len(clauses) >= 2:
                journal = _pick_journal()

    surname = authors[0]["family"] if authors else ((title or "").split() or ["anon"])[0]
    return {
        "base_key": _make_base_ref_key(surname, year),
        "raw_text": raw,
        "year": year,
        "title": title,
        "authors": authors,
        "journal": journal,
        "doi": doi,
        "pmid": pmid,
    }


# ---------- 元数据解析（CrossRef / PubMed 搜索） ----------
def build_query(parsed_ref):
    """构造 (provider, query_string)。query.bibliographic 为模糊查询。"""
    parts = []
    if parsed_ref.get("authors"):
        parts.append(parsed_ref["authors"][0]["family"])
    if parsed_ref.get("title"):
        parts.append(parsed_ref["title"])
    if parsed_ref.get("journal"):
        parts.append(parsed_ref["journal"])
    if parsed_ref.get("year"):
        parts.append(parsed_ref["year"])
    return "crossref", " ".join(p for p in parts if p)


def _http_get_json(url, params=None, headers=None):
    import requests
    r = requests.get(url, params=params, headers=headers or {}, timeout=30)
    r.raise_for_status()
    return r.json()


def _backoff(attempt, base=1.2):
    time.sleep(base * (2 ** attempt) + random.uniform(0, 0.5))


_last_req = {}


def _rate_limit(name, interval):
    """按 provider 保证相邻请求不低于 interval 秒，避免批量解析违反礼貌池限流。"""
    now = time.time()
    wait = interval - (now - _last_req.get(name, 0.0))
    if wait > 0:
        time.sleep(wait)
    _last_req[name] = time.time()


def search_crossref(query, rows=5):
    """CrossRef 书目搜索。礼貌池限流 ~1/1.1s，429 退避重试。"""
    _rate_limit("crossref", 1.1)
    params = {
        "query.bibliographic": query,
        "rows": rows,
        "select": "DOI,title,author,issued,container-title,type,volume,issue,page,publisher",
    }
    mailto = config.get_crossref_mailto() if config else None
    if mailto:
        params["mailto"] = mailto
    headers = {"User-Agent": f"zotero-reverse-ingest/1.0 (mailto:{mailto or 'anonymous'})"}
    for attempt in range(4):
        try:
            data = _http_get_json("https://api.crossref.org/works", params, headers)
            return data.get("message", {}).get("items", [])
        except Exception:
            _backoff(attempt)
    return []


def search_crossref_by_doi(doi):
    """按 DOI 直接从 CrossRef 取单条元数据（精确，用于参考文献里已含 DOI 的情况）。"""
    _rate_limit("crossref", 1.1)
    mailto = config.get_crossref_mailto() if config else None
    params = {"mailto": mailto} if mailto else None
    headers = {"User-Agent": f"zotero-reverse-ingest/1.0 (mailto:{mailto or 'anonymous'})"}
    try:
        data = _http_get_json(f"https://api.crossref.org/works/{quote(doi, safe='')}", params, headers)
        msg = data.get("message")
        return [msg] if msg else []
    except Exception:
        return []


def search_pubmed(query, rows=5):
    """NCBI eutils esearch + esummary。免 key ~3 req/s。"""
    _rate_limit("pubmed", 0.4)
    base = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/"
    api_key = config.get_ncbi_api_key() if config else None
    p = {"db": "pubmed", "term": query, "retmode": "json", "retmax": rows}
    if api_key:
        p["api_key"] = api_key
    try:
        ids = _http_get_json(base + "esearch.fcgi", p).get("esearchresult", {}).get("idlist", [])
        if not ids:
            return []
        p2 = {"db": "pubmed", "id": ",".join(ids), "retmode": "json"}
        if api_key:
            p2["api_key"] = api_key
        res = _http_get_json(base + "esummary.fcgi", p2).get("result", {})
        items = []
        for pid in ids:
            d = res.get(pid)
            if not d:
                continue
            authors = []
            for a in d.get("authors", []):
                nm = a.get("name", "")
                if nm:
                    parts = nm.split()
                    authors.append({"family": parts[-1], "given": " ".join(parts[:-1])})
            items.append({
                "pmid": pid,
                "title": "".join(d.get("title", "")),
                "source": d.get("source", ""),
                "pubdate": d.get("pubdate", ""),
                "authors": authors,
                "doi": next((x.get("value", "") for x in d.get("articleids", [])
                             if x.get("idtype") == "doi"), ""),
            })
        return items
    except Exception:
        return []


# ---------- CSL 构建 + 置信度 ----------
def _map_type(t):
    m = {"journal-article": "article-journal", "journal-issue": "article-journal",
         "book-chapter": "chapter", "book": "book", "proceedings-article": "paper-conference",
         "report": "report", "dissertation": "thesis", "posted-content": "article-journal",
         "webpage": "webpage", "dataset": "dataset"}
    return m.get(t, "article-journal")


def _clean_crossref_text(s):
    """CrossRef 字段常内嵌 <i>…</i> 等 HTML 标记和 HTML 实体；
    不清理会把标记原样写进 Zotero 标题（如 "Late <i>N</i>-acetylcysteine…"）。"""
    if not s:
        return s
    s = re.sub(r'<[^>]+>', '', str(s))
    s = html.unescape(s)
    return re.sub(r'\s+', ' ', s).strip()


def crossref_to_csl(item):
    """CrossRef item -> CSL-JSON（type/title/author/issued，无 abstract）。标题/期刊名已剥 HTML。"""
    csl = {"type": _map_type(item.get("type", ""))}
    title = item.get("title")
    if title:
        csl["title"] = _clean_crossref_text(title[0] if isinstance(title, list) else title)
    authors = []
    for a in item.get("author", []):
        if a.get("family"):
            authors.append({"family": a["family"], "given": a.get("given", "")})
        elif a.get("name"):
            authors.append({"literal": a["name"]})
    if authors:
        csl["author"] = authors
    issued = item.get("issued", {}).get("date-parts")
    if issued:
        csl["issued"] = {"date-parts": issued}
    ct = item.get("container-title")
    if ct:
        csl["container-title"] = _clean_crossref_text(ct[0] if isinstance(ct, list) else ct)
    for f in ("volume", "issue", "page", "publisher"):
        if item.get(f):
            csl[f] = item[f]
    if item.get("DOI"):
        csl["DOI"] = item["DOI"]
    return csl


def pubmed_to_csl(it):
    csl = {"type": "article-journal"}
    if it.get("title"):
        csl["title"] = it["title"]
    if it.get("authors"):
        csl["author"] = it["authors"]
    if it.get("source"):
        csl["container-title"] = it["source"]
    ym = re.search(r'\b(19|20)\d{2}\b', it.get("pubdate", ""))
    if ym:
        csl["issued"] = {"date-parts": [[int(ym.group(0))]]}
    if it.get("doi"):
        csl["DOI"] = it["doi"]
    return csl


def confidence_score(csl, parsed_ref):
    """0..1 启发式：首作者姓 +0.4、年份 +0.3、标题词重叠 +0.3。返回 (score, matched)。"""
    score = 0.0
    matched = []
    p_au = parsed_ref.get("authors") or []
    c_au = csl.get("author") or []
    if p_au and c_au:
        pf = _norm(p_au[0].get("family", ""))
        cf = _norm(c_au[0].get("family", "") or c_au[0].get("literal", ""))
        if pf and cf and (pf == cf or pf in cf or cf in pf):
            score += 0.4
            matched.append("author")
    p_y = parsed_ref.get("year")
    c_y = _year(csl)
    if p_y and c_y and _norm(str(p_y)) == _norm(str(c_y)):
        score += 0.3
        matched.append("year")
    p_t = parsed_ref.get("title")
    c_t = csl.get("title")
    if p_t and c_t:
        wp = _title_words(p_t)
        wc = _title_words(c_t)
        inter = wp & wc
        if wc and inter:
            score += 0.3 * (len(inter) / max(len(wc), 1))
            matched.append("title")
    return round(min(score, 1.0), 2), matched


def resolve_metadata(refs, providers=("crossref", "pubmed")):
    """批量解析。返回 {ref_key: {csl, provider, confidence, candidates}}。

    参考文献里已含 DOI 时优先按 DOI 直查 CrossRef（精确）；否则模糊搜索 CrossRef，
    置信度不足再走 PubMed 兜底。
    """
    resolved = {}
    for (_idx, key, raw) in refs:
        parsed = parse_numbered_reference(raw)
        candidates = []
        best = None
        best_score = 0.0

        def consider(cand):
            nonlocal best, best_score
            candidates.append(cand)
            if cand["confidence"] > best_score:
                best, best_score = cand, cand["confidence"]

        # 1) 若参考文献自带 DOI，先按 DOI 直查（精确匹配视为满分）
        if parsed.get("doi") and "crossref" in providers:
            for it in search_crossref_by_doi(parsed["doi"]):
                csl = crossref_to_csl(it)
                sc, matched = confidence_score(csl, parsed)
                if it.get("DOI") and _norm(parsed["doi"]) == _norm(it["DOI"]):
                    sc, matched = 1.0, list(set(matched) | {"doi"})
                consider({"csl": csl, "provider": "crossref", "doi": it.get("DOI"),
                          "confidence": sc, "matched": matched})
        # 2) 置信度不足或无 DOI 命中，转模糊搜索
        if best is None or best_score < REF_KEY_MIN_CONFIDENCE:
            _prov, q = build_query(parsed)
            if "crossref" in providers and q:
                for it in search_crossref(q):
                    csl = crossref_to_csl(it)
                    sc, matched = confidence_score(csl, parsed)
                    consider({"csl": csl, "provider": "crossref", "doi": it.get("DOI"),
                              "confidence": sc, "matched": matched})
        # 3) CrossRef 仍不足，走 PubMed 兜底
        if (best is None or best_score < REF_KEY_MIN_CONFIDENCE) and "pubmed" in providers:
            _prov, q = build_query(parsed)
            if q:
                for it in search_pubmed(q):
                    csl = pubmed_to_csl(it)
                    sc, matched = confidence_score(csl, parsed)
                    consider({"csl": csl, "provider": "pubmed", "pmid": it.get("pmid"),
                              "confidence": sc, "matched": matched})
        resolved[key] = {
            "csl": best["csl"] if best else {},
            "provider": best["provider"] if best else None,
            "confidence": best_score,
            "candidates": candidates[:5],
            "parsed": parsed,
            "resolve_error": best is None,  # 显式标记解析失败（如 DOI 死链），勿静默跳过
        }
    _write_json("resolution_summary.json", _summarize(resolved))
    return resolved


def _summarize(resolved):
    return {k: {"title": v["csl"].get("title", ""),
                "first_author": _first_author_lastname(v["csl"]),
                "year": _year(v["csl"]),
                "provider": v["provider"],
                "confidence": v["confidence"],
                "level": _confidence_level(v["confidence"]),
                "resolve_error": v.get("resolve_error", False)}
            for k, v in resolved.items()}


def _confidence_level(score):
    if score >= REF_KEY_MIN_CONFIDENCE:
        return "HIGH"
    if score > 0:
        return "LOW"
    return "UNKNOWN"


# ---------- 去重（导入前） ----------
def normalize_title(s):
    """小写/NFKD/去标点/压空白。"""
    return _norm(s)


def _results(res):
    """把 MCP 各工具五花八门的返回形态统一成条目列表。

    ⚠️ 形态不统一是实测过的坑：search_library 返回 {"results": [...]}，
    get_collection_items / get_collection_items 树 直接返回**数组**，
    另有 {"items": [...]}。旧实现只认 dict+results，遇到数组一律返回 []，
    于是"按文件夹建库索引"会静默建出空索引（v3 回归实测：285 条的文件夹返回 0 条，
    0 失败，看起来一切正常）。
    """
    if isinstance(res, list):
        return res
    if isinstance(res, dict):
        for k in ("results", "items", "data"):
            if isinstance(res.get(k), list):
                return res[k]
    return []


def _search_results(zotero, q, mode="standard", limit=5):
    res = zotero("search_library", {"q": q, "mode": mode, "limit": limit})
    return _results(res)


def _doi_search(zotero, doi):
    """按 DOI 全文搜索并用 get_item_details 逐条确认。

    ⚠️ search_library 的 minimal/standard 结果都**不含 DOI 字段**：
    旧实现拿 minimal 结果直接比 it["DOI"]，永远落空 -> 漏检库中已有条目 -> 重复导入
    （且 MCP 没有 delete_item，重复只能手动清理）。必须 fetch 详情确认。"""
    hits = []
    for it in _search_results(zotero, f'"{doi}"'):
        key = it.get("key")
        if not key:
            continue
        try:
            det = zotero("get_item_details", {"itemKey": key, "mode": "complete"})
        except Exception:
            continue
        if isinstance(det, dict) and key == det.get("key", key) and \
                normalize_doi(det.get("DOI")) == normalize_doi(doi):
            det["key"] = key
            hits.append(det)
    return hits


def normalize_doi(doi):
    """DOI 归一化：去前缀、去尾随标点、转小写。老式括号 DOI（(98)00014-6）原样保留。

    手稿里的 DOI 写法五花八门：`https://doi.org/10.x`、`doi:10.x`、`10.x.`、
    `(doi:10.x; 10.y)` 的分片、`10.1016/S0306-4530(98)00014-6`。
    直接用原串查库必然漏命中。所有比对路径都应先过这个函数。
    """
    if not doi:
        return ""
    s = str(doi).strip()
    s = re.sub(r'^[\s(（]+', '', s)                       # 前导分组括号
    s = re.sub(r'^(?:https?://(?:dx\.)?doi\.org/|doi\s*[:：]\s*)', '', s, flags=re.I)
    s = s.strip().rstrip('.,;')
    # 尾部括号属于分组而非 DOI；老式 DOI 自身以 (98) 结尾的年份段要保留
    while s and s[-1] in ')）' and not re.search(r'\(\d{2}[A-Za-z0-9]?\)$', s):
        s = s[:-1].rstrip()
    return s.lower()


def _collection_item_keys(zotero, collection_key):
    res = zotero("get_collection_items", {"collectionKey": collection_key, "limit": 500})
    return [it.get("key") for it in _results(res) if it.get("key")]


def build_library_doi_index(collection_keys=None, item_keys=None, zotero=None,
                            progress_every=0, csl_builder=None):
    """把 Zotero 库（指定文件夹 / 指定 key 集合）建成 {归一化DOI: {itemKey, uri, title, csl}}。

    **绝不静默跳过失败**：get_item_details 对已删除条目会返回 {'error': ...}，
    旧写法 `except/if error: continue` 会把这类条目悄悄丢掉，索引看着"很干净"，
    实际漏掉了死链（v2 实测：422 个 key 全部"成功"，但文档里引用的一条已删条目
    从未进入索引，直到最后单独校验 key 存在性才暴露）。
    返回 (index, failures, dups)：
      failures: [(key, error)] —— 必须报告给用户
      dups:     {doi: [key, ...]} —— 同 DOI 多条目，需用户决定合并到哪条
    """
    zotero = _get_zotero(zotero)
    if item_keys is None:
        keys = []
        for ck in (collection_keys or []):
            keys.extend(_collection_item_keys(zotero, ck))
        item_keys = sorted(set(keys))
    index, failures, dups = {}, [], {}
    for i, k in enumerate(item_keys or [], 1):
        try:
            det = zotero("get_item_details", {"itemKey": k, "mode": "complete"})
        except Exception as e:
            failures.append((k, f"{type(e).__name__}: {e}"))
            continue
        if not isinstance(det, dict) or det.get("error") or not det.get("title"):
            failures.append((k, str(det)[:120] if isinstance(det, dict) else "bad response"))
            continue
        d = normalize_doi(det.get("DOI") or "")
        if not d:
            continue
        if d in index:
            dups.setdefault(d, [index[d]["itemKey"]]).append(k)
            continue
        index[d] = {
            "itemKey": k,
            "uri": f"http://zotero.org/users/{_zotero_user_id()}/items/{k}",
            "doi": det.get("DOI") or d,
            "title": det.get("title", ""),
            "csl": csl_builder(det) if csl_builder else {},
            "itemType": det.get("itemType", ""),
        }
        if progress_every and i % progress_every == 0:
            print(f"  …{i}/{len(item_keys)}", flush=True)
    return index, failures, dups


def _zotero_user_id():
    try:
        return config.get_user_id()
    except Exception:
        return "local"


def verify_item_keys(keys, zotero=None, progress_every=0):
    """校验 itemKey 集合在库内是否真实存在（死链审计）。

    返回 (ok_keys, missing)：missing 为 [(key, error)]。转换交付前的必查项——
    既有域可能指向用户早已删除的条目，Zotero 刷新后会显示 [item removed]。
    修法：找到库内同 DOI 的规范条目后用 docx_audit.relink_item_keys() 重链接。
    """
    zotero = _get_zotero(zotero)
    ok, missing = [], []
    keys = sorted(set(keys or []))
    for i, k in enumerate(keys, 1):
        try:
            det = zotero("get_item_details", {"itemKey": k, "mode": "minimal"})
            bad = (not isinstance(det, dict) or det.get("error")
                   or not det.get("title"))
        except Exception as e:
            bad, det = True, {"error": str(e)[:80]}
        (missing.append((k, str(det.get('error'))[:100])) if bad else ok.append(k))
        if progress_every and i % progress_every == 0:
            print(f"  …{i}/{len(keys)}", flush=True)
    return ok, missing


def doc_item_dois(docx_path, zotero=None, csl_builder=None):
    """文档既有引文域引用的条目 -> {归一化DOI: {itemKey, uri, title, csl}}。

    "半刷新"手稿（部分段落已是域、部分仍是 DOI 占位）的映射必须取
    **占位 DOI ∪ 既有域引用** 的并集：只被既有域引用的文献不在占位集合里，
    少了它们，校验会把合法既有域误判为映射外。
    """
    try:
        import docx_audit
    except ImportError:
        docx_audit = None
    import zipfile as _zipfile
    from lxml import etree as _etree
    if docx_audit is not None:
        root = docx_audit.etree.fromstring(
            _zipfile.ZipFile(docx_path).read('word/document.xml'))
        keys = docx_audit.zfi.document_item_keys(root)
    else:
        from zotero_field_insert import document_item_keys
        root = _etree.fromstring(_zipfile.ZipFile(docx_path).read('word/document.xml'))
        keys = document_item_keys(root)
    idx, failures, dups = {}, [], {}
    zotero = _get_zotero(zotero)
    for k in sorted(keys):
        try:
            det = zotero("get_item_details", {"itemKey": k, "mode": "complete"})
        except Exception as e:
            failures.append((k, f"{type(e).__name__}: {e}"))
            continue
        if not isinstance(det, dict) or det.get("error") or not det.get("title"):
            failures.append((k, "not found"))
            continue
        d = normalize_doi(det.get("DOI") or "")
        if not d:
            continue
        idx[d] = {"itemKey": k,
                  "uri": f"http://zotero.org/users/{_zotero_user_id()}/items/{k}",
                  "doi": det.get("DOI") or d,
                  "title": det.get("title", ""),
                  "csl": csl_builder(det) if csl_builder else {}}
    return idx, failures, dups


def dedup_check(resolved, collection_key=None, zotero=None):
    """逐篇查重。标题路径（minimal 搜索 + 归一化比对）+ DOI 路径（全文搜索 + 详情确认）。
    标题无 exact 命中时**总是**再走 DOI 路径（标题大小写/Unicode 连字符差异会让标题
    搜索漏掉库内条目）。返回 {ref_key: {status, existing}}。不导入。"""
    zotero = _get_zotero(zotero)
    report = {}
    for key, info in resolved.items():
        csl = info.get("csl", {})
        nt = normalize_title(csl.get("title", ""))
        status = "none"
        existing = None
        if nt:
            q = csl.get("title", "") or nt  # 用原始标题查询，缓解去空格导致的相关度下降
            res = zotero("search_library", {"q": q, "mode": "minimal", "limit": 5})
            for it in _results(res):
                it_t = normalize_title(it.get("title", ""))
                hit = it_t == nt
                if it_t and not hit and (it_t in nt or nt in it_t):
                    hit = "probable"
                if hit is True:
                    status, existing = "exact", it
                    break
                elif hit == "probable" and status != "exact":
                    status, existing = "probable", it
        # DOI 路径：标题未 exact 命中就尝试（probable 也升级确认）。minimal 结果没有
        # DOI 字段，_doi_search 内部会逐条 get_item_details 确认，不会误判。
        if status != "exact" and csl.get("DOI"):
            hits = _doi_search(zotero, csl["DOI"])
            if hits:
                status, existing = "exact", hits[0]
        report[key] = {
            "status": status,
            "existing": _existing_meta(existing) if existing else None,
        }
    _write_json("dedup_report.json", report)
    return report


def audit_imported_by_doi(resolved, zotero=None):
    """导入后审计：逐 DOI 全文搜索库内命中数。

    命中数 >1 说明库内存在重复（MCP 无删除工具，只能列出让用户在 Zotero UI 里
    手动清理）。返回 {ref_key: {"hits": n, "keys": [...]}}，写 doi_audit.json。"""
    zotero = _get_zotero(zotero)
    report = {}
    for key, info in resolved.items():
        doi = info.get("csl", {}).get("DOI") or info.get("doi")
        if not doi:
            continue
        keys = [it["key"] for it in _doi_search(zotero, doi)]
        report[key] = {"hits": len(keys), "keys": keys}
    _write_json("doi_audit.json", report)
    return report


def _existing_meta(it):
    prefix = config.get_uri_prefix() if config else ""
    key = it.get("key")
    return {"key": key, "title": it.get("title", ""),
            "uri": (prefix + key) if key else ""}


# ---------- 导入（给定用户选定的 collection key） ----------
def _csl_to_field_map(csl):
    fields = {}
    itemType = _csl_to_itemtype(csl)
    if csl.get("title"):
        fields["title"] = csl["title"]
    if csl.get("container-title"):
        # 章节：书名放 bookTitle（Zotero bookSection 字段），期刊放 publicationTitle
        fields["bookTitle" if itemType == "bookSection" else "publicationTitle"] = csl["container-title"]
    if csl.get("volume"):
        fields["volume"] = str(csl["volume"])
    if csl.get("issue"):
        fields["issue"] = str(csl["issue"])
    if csl.get("page"):
        fields["pages"] = str(csl["page"])
    if csl.get("publisher"):
        fields["publisher"] = csl["publisher"]
    if csl.get("DOI"):
        fields["DOI"] = csl["DOI"]
    issued = csl.get("issued", {}).get("date-parts")
    if issued and issued[0] and issued[0][0] is not None:
        y = issued[0][0]
        m = issued[0][1] if len(issued[0]) > 1 else None
        d = issued[0][2] if len(issued[0]) > 2 else None
        date = str(y)
        if m is not None:
            try:
                date += f"-{int(m):02d}"
            except (TypeError, ValueError):
                date += f"-{m}"
        if d is not None:
            try:
                date += f"-{int(d):02d}"
            except (TypeError, ValueError):
                date += f"-{d}"
        fields["date"] = date
    return fields


def _csl_to_creators(csl):
    out = []
    for a in csl.get("author", []):
        if a.get("literal"):
            out.append({"creatorType": "author", "name": a["literal"]})
        else:
            out.append({"creatorType": "author",
                        "firstName": a.get("given", ""),
                        "lastName": a.get("family", "")})
    return out


def _csl_to_itemtype(csl):
    m = {"article-journal": "journalArticle", "chapter": "bookSection", "book": "book",
         "paper-conference": "conferencePaper", "report": "report", "thesis": "thesis",
         "webpage": "webpage", "dataset": "dataset", "posted-content": "preprint"}
    return m.get(csl.get("type"), "journalArticle")


def _extract_item_key(res):
    """从 write_item 返回里提取 Zotero item key（8 位 base62）。

    兼容常见返回形态：顶层 key/itemKey、嵌套于 item/data/result、列表、或纯文本 key。
    """
    if isinstance(res, list):
        for x in res:
            k = _extract_item_key(x)
            if k:
                return k
        return None
    if isinstance(res, str):
        s = res.strip()
        if re.fullmatch(r'[0-9A-Za-z]{8}', s):
            return s
        m = re.search(r'\b([0-9A-Za-z]{8})\b', s)
        return m.group(1) if m else None
    if isinstance(res, dict):
        if res.get("key") and len(str(res["key"])) <= 16:
            return str(res["key"])
        if res.get("itemKey"):
            return str(res["itemKey"])
        # 递归找嵌套的 key / item
        for k in ("item", "data", "result", "newItem", "created"):
            v = res.get(k)
            if isinstance(v, (dict, list)):
                kk = _extract_item_key(v)
                if kk:
                    return kk
        # 兜底：扫描所有值里像 8 位 key 的
        for k, v in res.items():
            if k.lower() in ("key", "itemkey", "id") and isinstance(v, str) and len(v) <= 16:
                return v
    return None


def _find_existing_item(csl, zotero):
    """导入前按 DOI/标题查库内是否已有该条目（幂等导入用）。
    返回 itemKey 或 None。DOI 路径走 _doi_search（详情确认），标题路径走归一化比对。"""
    doi = csl.get("DOI")
    if doi:
        hits = _doi_search(zotero, doi)
        if hits:
            return hits[0]["key"]
    nt = normalize_title(csl.get("title", ""))
    if nt:
        res = zotero("search_library", {"q": csl.get("title", ""), "mode": "minimal", "limit": 5})
        for it in _results(res):
            if normalize_title(it.get("title", "")) == nt:
                return it.get("key")
    return None


def import_to_zotero(import_plan, collection_key, zotero=None, uri_prefix=None, precheck=True):
    """按用户批准的计划导入。import_plan = {ref_key: {action, csl, zotero_key?}}。

    action: 'import' 新建 / 'reuse' 复用库中已有(zotero_key) / 'skip' 跳过。
    precheck=True 时 'import' 动作先按 DOI/标题查重，已有则自动转为复用--
    批量导入中途失败后重跑不会重复建条目（MCP 无删除工具，重复只能手动清理）。
    返回 (imported, warnings)。imported = {ref_key: {itemKey, uri, csl}}；
    warnings 列出 write_item 后未能提取到 key 的条目（避免静默丢失）。
    """
    zotero = _get_zotero(zotero)
    if uri_prefix is None:
        uri_prefix = config.get_uri_prefix() if config else ""
    imported = {}
    warnings = []
    items_to_add = []
    for key, plan in import_plan.items():
        action = plan.get("action", "import")
        csl = plan.get("csl") or {}
        if action == "skip":
            continue
        if action == "reuse":
            zk = plan.get("zotero_key")
            if zk:
                items_to_add.append(zk)
                imported[key] = {"itemKey": zk, "uri": uri_prefix + zk, "csl": csl}
            else:
                warnings.append(f"{key}: reuse 但未提供 zotero_key")
            continue
        # action == 'import'：新建
        if precheck:
            zk = _find_existing_item(csl, zotero)
            if zk:
                items_to_add.append(zk)
                imported[key] = {"itemKey": zk, "uri": uri_prefix + zk, "csl": csl,
                                 "precheck_reused": True}
                continue
        args = {
            "action": "create",
            "itemType": _csl_to_itemtype(csl),  # 必须顶层参数，不能放进 fields
            "fields": _csl_to_field_map(csl),
        }
        creators = _csl_to_creators(csl)
        if creators:
            # ⚠️ 空列表必须整个省略：Zotero 对空 creators 报
            # "Creator names cannot be empty"，会让批量导入中途崩掉（部分条目已建）
            args["creators"] = creators
        res = zotero("write_item", args)
        zk = _extract_item_key(res)
        if zk:
            items_to_add.append(zk)
            imported[key] = {"itemKey": zk, "uri": uri_prefix + zk, "csl": csl}
        else:
            warnings.append(f"{key}: write_item 未返回可用的 item key（响应: {_safe_str(res)}）")
    if items_to_add and collection_key:
        zotero("add_items_to_collection", {
            "collectionKey": collection_key, "itemKeys": items_to_add})
    return imported, warnings


def _safe_str(res):
    try:
        return json.dumps(res, ensure_ascii=False)[:200]
    except Exception:
        return str(res)[:200]


# ---------- mapping 构建（喂给下游生成器，格式逐字节兼容） ----------
def build_item_mapping(imported, uri_prefix=None):
    if uri_prefix is None:
        uri_prefix = config.get_uri_prefix() if config else ""
    mapping = {}
    for key, info in imported.items():
        csl = info.get("csl", {})
        mapping[key] = {
            "itemKey": info["itemKey"],
            "uri": info.get("uri") or (uri_prefix + info["itemKey"]),
            "csl": csl,
            "doi": csl.get("DOI", ""),
            "first_author": _first_author_lastname(csl),
            "year": _year(csl),
        }
    return mapping


# ---------- 以库为准：用 Zotero 库内元数据重建 itemData ----------
def _parse_lib_date(d):
    """库内 date 形如 "2008" / "2023-01" / "2026-05-13" -> date-parts 列表。"""
    if not d:
        return None
    m = re.match(r'^(\d{4})(?:-(\d{1,2}))?(?:-(\d{1,2}))?', str(d).strip())
    if not m:
        return None
    parts = [int(m.group(1))]
    if m.group(2):
        parts.append(int(m.group(2)))
    if m.group(3):
        parts.append(int(m.group(3)))
    return parts


def _lib_details_to_csl(det, fallback_type="article-journal"):
    """get_item_details(complete) 的返回 -> CSL-JSON（域内嵌 itemData 用）。
    条目类型沿用库内 itemType（详情缺失时退回 fallback_type）。"""
    itype = (det.get("itemType") or "").strip()
    csl = {"type": "report" if itype == "report" else fallback_type}
    if det.get("title"):
        csl["title"] = det["title"]
    authors = []
    for c in det.get("creators", []) or []:
        ln, fn = c.get("lastName", ""), c.get("firstName", "")
        if ln and fn:
            authors.append({"family": ln, "given": fn})
        elif fn and not ln:
            authors.append({"literal": fn})
        elif ln:
            # Zotero 组织作者把单位名放 lastName、firstName 为空——按 CSL 机构作者用 literal。
            # （真实 Zotero 域内嵌 itemData 对组织作者即为 {"literal": "…"}）
            authors.append({"literal": ln})
    if authors:
        csl["author"] = authors
    dp = _parse_lib_date(det.get("date"))
    csl["issued"] = {"date-parts": [dp]} if dp else {"date-parts": [[]]}
    if csl["type"] != "report":
        for src, dst in (("publicationTitle", "container-title"), ("volume", "volume"),
                         ("issue", "issue"), ("pages", "page")):
            if det.get(src):
                csl[dst] = det[src]
    if det.get("DOI"):
        csl["DOI"] = det["DOI"]
    if det.get("url"):
        csl["URL"] = det["url"]
    return csl


def apply_library_data(item_mapping, zotero=None, update_doi_field=False):
    """以 Zotero 库内元数据为准，重建 mapping 中每条的 csl（域内嵌 itemData 用）。

    复用库中已有条目时，CrossRef 数据与库内常有出入（标题大小写、Unicode 连字符、
    缺录的 DOI/卷期页等）；插入域代码前跑一遍本函数，itemData 即与库一致，
    Zotero 刷新前后表现稳定。
    返回 (mapping, doi_mismatches)；doi_mismatches = [(ref_key, mapping_doi, lib_doi)]
    （mapping 的 doi 锚点保持不变，只用于正文匹配）。"""
    zotero = _get_zotero(zotero)
    mismatches = []
    for ref_key, v in item_mapping.items():
        zk = v.get("itemKey")
        if not zk:
            continue
        try:
            det = zotero("get_item_details", {"itemKey": zk, "mode": "complete"})
        except Exception:
            continue
        if not isinstance(det, dict) or not det.get("title"):
            continue
        lib_doi = str(det.get("DOI", "")).strip()
        map_doi = str(v.get("doi", "")).strip()
        if lib_doi and map_doi and lib_doi.lower() != map_doi.lower():
            mismatches.append((ref_key, map_doi, lib_doi))
        v["csl"] = _lib_details_to_csl(det, fallback_type=v.get("csl", {}).get("type", "article-journal"))
        v["first_author"] = _first_author_lastname(v["csl"]) or v.get("first_author", "")
        v["year"] = _year(v["csl"]) or v.get("year", "")
        if update_doi_field and lib_doi:
            v["doi"] = lib_doi
    return item_mapping, mismatches


# ---------- 非交互兜底编排 ----------
def ingest(docx_path, collection_key, out="item_mapping.json", ref_para_indices=None,
           zotero=None, uri_prefix=None):
    """parse -> resolve -> dedup -> [无检查点] -> import -> mapping。

    注意：这是非交互捷径，会直接导入。交互 skill 流程应由 Claude 分步调用，
    并在 resolve/dedup 之后、import 之前插入强制检查点。
    """
    zotero = _get_zotero(zotero)
    mcp = check_zotero_mcp(zotero=zotero)
    if not mcp.get("ok"):
        return {"error": mcp.get("hint")}
    refs = extract_reference_paragraphs(docx_path, ref_para_indices)
    if not refs:
        return {"error": "未在文档中解析到编号参考文献。请确认 ref_para_indices 或编号格式（[n]/(n)/n./n)）。"}
    resolved = resolve_metadata(refs)
    if not any(v.get("csl") for v in resolved.values()):
        return {"error": "全部参考文献均未解析出元数据（联网失败或格式异常）。请检查网络/Crossref/PubMed 可达性。"}
    dup = dedup_check(resolved, collection_key=collection_key, zotero=zotero)
    import_plan = {k: {"action": "import", "csl": v["csl"]}
                   for k, v in resolved.items() if v.get("csl")}
    imported, warnings_ = import_to_zotero(import_plan, collection_key, zotero=zotero, uri_prefix=uri_prefix)
    mapping = build_item_mapping(imported, uri_prefix=uri_prefix)
    _write_json(out, mapping)
    stats = {
        "total": len(resolved),
        "resolved": sum(1 for v in resolved.values() if v.get("csl")),
        "low_confidence": sum(1 for v in resolved.values() if v["confidence"] < REF_KEY_MIN_CONFIDENCE),
        "dups": sum(1 for v in dup.values() if v["status"] != "none"),
        "imported": len(imported),
        "skipped": len(resolved) - len(imported),
        "warnings": warnings_,
    }
    return mapping, stats


if __name__ == "__main__":
    print("逆向 Zotero 插入 —— 前置引用摄取模块就绪。")
    print("函数: check_zotero_mcp / extract_reference_paragraphs / parse_numbered_reference /")
    print("      resolve_metadata / dedup_check / import_to_zotero / build_item_mapping / ingest")
    print("用途见 SKILL.md。交互流程由 Claude 分步调用，并在导入前插入强制检查点。")