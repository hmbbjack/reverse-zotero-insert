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


def crossref_to_csl(item):
    """CrossRef item -> CSL-JSON（type/title/author/issued，无 abstract）。"""
    csl = {"type": _map_type(item.get("type", ""))}
    title = item.get("title")
    if title:
        csl["title"] = title[0] if isinstance(title, list) else title
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
        csl["container-title"] = ct[0] if isinstance(ct, list) else ct
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
        }
    _write_json("resolution_summary.json", _summarize(resolved))
    return resolved


def _summarize(resolved):
    return {k: {"title": v["csl"].get("title", ""),
                "first_author": _first_author_lastname(v["csl"]),
                "year": _year(v["csl"]),
                "provider": v["provider"],
                "confidence": v["confidence"],
                "level": _confidence_level(v["confidence"])}
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


def dedup_check(resolved, collection_key=None, zotero=None):
    """逐篇 search_library 查重。返回 {ref_key: {status, existing}}。不导入。"""
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
        # 补充 DOI 匹配
        if status == "none" and csl.get("DOI"):
            doi = csl["DOI"].lower()
            res = zotero("search_library", {"q": doi, "mode": "minimal", "limit": 5})
            for it in _results(res):
                if (it.get("DOI") or "").lower() == doi:
                    status, existing = "exact", it
                    break
        report[key] = {
            "status": status,
            "existing": _existing_meta(existing) if existing else None,
        }
    _write_json("dedup_report.json", report)
    return report


def _results(res):
    if isinstance(res, dict) and isinstance(res.get("results"), list):
        return res["results"]
    return []


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


def import_to_zotero(import_plan, collection_key, zotero=None, uri_prefix=None):
    """按用户批准的计划导入。import_plan = {ref_key: {action, csl, zotero_key?}}。

    action: 'import' 新建 / 'reuse' 复用库中已有(zotero_key) / 'skip' 跳过。
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
        res = zotero("write_item", {
            "action": "create",
            "itemType": _csl_to_itemtype(csl),  # 必须顶层参数，不能放进 fields
            "fields": _csl_to_field_map(csl),
            "creators": _csl_to_creators(csl),
        })
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