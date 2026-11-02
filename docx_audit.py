#!/usr/bin/env python3
"""
docx_audit.py —— 手稿审计与修补工具箱（v3）
================================================
把"转换一份 .docx 前后必须搞清楚的事情"做成一次调用就能拿到答案的函数，
不再依赖每个会话临时手写扫描脚本（历史上同一段扫描代码被重写过多次，
每次都带进一个新 bug：逐段重置域深度、逐元素解析 JSON、字符串当元素用……）。

四类能力：
1. **扫描**  `scan_document(path)` —— 域清单/样式/占位组/表格引文列/citationID/死链，
   一条命令拿到转换前全貌（对应 SKILL.md 工作流第 0 步）。
2. **提取**  `extract_doi_groups(text)` —— 括号配平 + 老式 Elsevier DOI 完整识别。
3. **修补**  `renumber_citation_ids()` / `relink_item_keys()` —— citationID 隔离重编号、
   死链重链接（老文档的历史遗留问题）。
4. **校验**  `audit_output(src, out, item_mapping)` —— 转换后逐组对账（零遗漏 +
   显示文本逐字 + DOI 集合 + URI 指向），比"数域个数"强得多。

命令行：
    python3 docx_audit.py scan <file.docx> [--json]
    python3 docx_audit.py audit <src.docx> <out.docx> <item_mapping.json>
    python3 docx_audit.py renumber <file.docx> [--prefix sis] [--from <src.docx>]
    python3 docx_audit.py relink <file.docx> '{"OLDKEY":"NEWKEY"}'
"""

import zipfile
import json
import re
import argparse
import sys
import os

from lxml import etree

if __package__:
    from . import zotero_field_insert as zfi
    from .docx_io import load_root, save_document
    from .doi_utils import DOI_RE, extract_doi_groups, normalize_doi
    from .word_fields import parse_fields, replace_instruction
else:
    import zotero_field_insert as zfi
    from docx_io import load_root, save_document
    from doi_utils import DOI_RE, extract_doi_groups, normalize_doi
    from word_fields import parse_fields, replace_instruction

W = zfi.W
SCHEMA = zfi.SCHEMA

NUM_ONLY = re.compile(r"^\s*[\[(]?\s*\d{1,3}(\s*[-–,;]\s*\d{1,3})*\s*[\])]?\s*$")
PAREN_NUMBER = re.compile(r"[\[(]\s*\d{1,3}(\s*[,;–-]\s*\d{1,3})*\s*[\])]")


# ---------------------------------------------------------------- 提取占位
def extract_loose_dois(text, covered=()):
    """括号组之外的散落 DOI（返回 [(doi, context)]）。覆盖区间用于排除已提取的组。"""
    parts = []
    last = 0
    for s, e in covered:
        parts.append(text[last:s])
        last = e
    parts.append(text[last:])
    rest = "".join(parts)
    out = []
    for m in DOI_RE.finditer(rest):
        out.append((normalize_doi(m.group(0)), rest[max(0, m.start() - 60) : m.end() + 20]))
    return out


# ---------------------------------------------------------------- 扫描
def _cell_outside_text(tc, field_index=None):
    parsed = parse_fields(tc) if field_index is None else field_index
    return "".join(parsed.outside[p] for p in tc.iter(W + "p"))


def _tc_has_field(tc, field_index=None):
    parsed = parse_fields(tc) if field_index is None else field_index
    return any(element in parsed.protected for element in tc.iter())


def table_column_audit(root, field_index=None):
    """逐表逐列审计：区分"引文列"与"数据列"。

    表格里"纯数字格"绝大多数是数据（年龄 6–8、周期 14 天…），不是引文；
    真正要转的是：整格 DOI 占位、整格编号引文（"(1)" / "[12]"）且同行同列已是域。
    返回 [{table, rows, cols, header, columns:[{index, header, cells, with_field,
    doi_cells, number_cells, citation_cells}]}]
    """
    body = root.find(W + "body")
    parsed = parse_fields(root) if field_index is None else field_index
    report = []
    for ti, tbl in enumerate(body.findall(W + "tbl")):
        rows = tbl.findall(W + "tr")
        if not rows:
            continue
        ncols = max(len(r.findall(W + "tc")) for r in rows)
        header = [_cell_outside_text(tc, parsed).strip()[:24] for tc in rows[0].findall(W + "tc")]
        cols = []
        for ci in range(ncols):
            cells = [r.findall(W + "tc")[ci] for r in rows if ci < len(r.findall(W + "tc"))]
            with_field = doi_cells = number_cells = citation_cells = 0
            samples = {"doi": [], "number": [], "citation": []}
            for ri, tc in enumerate(cells):
                if _tc_has_field(tc, parsed):
                    with_field += 1
                    continue
                t = _cell_outside_text(tc, parsed).strip()
                if not t:
                    continue
                if DOI_RE.search(t):
                    doi_cells += 1
                    samples["doi"].append([ri, t[:60]])
                elif NUM_ONLY.match(t):
                    number_cells += 1
                    samples["number"].append([ri, t[:30]])
                elif PAREN_NUMBER.search(t) and len(t) < 60:
                    citation_cells += 1
                    samples["citation"].append([ri, t[:40]])
            cols.append(
                {
                    "index": ci,
                    "header": header[ci] if ci < len(header) else "",
                    "cells": len(cells),
                    "with_field": with_field,
                    "doi_cells": doi_cells,
                    "number_cells": number_cells,
                    "citation_cells": citation_cells,
                    "samples": samples,
                }
            )
        report.append(
            {"table": ti, "rows": len(rows), "cols": ncols, "header": header, "columns": cols}
        )
    return report


def scan_document(path, verbose=True):
    """转换前全貌扫描。返回可直接打印/落盘的 dict。"""
    root = load_root(path)
    body = root.find(W + "body")
    paras = body.findall(W + "p")
    parsed = parse_fields(root)
    fields = zfi.collect_fields(parsed)
    items = bibs = prefs = 0
    style = None
    for els, _, _ in fields:
        t = zfi.field_instr_text(els)
        if "ZOTERO_ITEM CSL_CITATION" in t:
            items += 1
        elif "ZOTERO_BIBL" in t:
            bibs += 1
        elif "ZOTERO_DOCUMENT_PREFERENCES" in t:
            prefs += 1
            obj = zfi.field_payload(els) or {}
            style = (obj.get("style") or {}).get("styleID")
    # 域外占位组（跨段深度感知，BIBL 内的文献不会混进来）
    groups = []
    for i, txt in zfi.outside_field_paragraphs(parsed):
        gs = extract_doi_groups(txt)
        for s, e, g, dois in gs:
            groups.append({"para": i, "span": [s, e], "text": g, "dois": dois})
        loose = extract_loose_dois(txt, [(s, e) for s, e, _, _ in gs])
        for d, ctx in loose:
            groups.append({"para": i, "loose": d, "context": ctx})
    mixed_paras = sorted(
        {g["para"] for g in groups if "span" in g and any(f[2] == g["para"] for f in fields)}
    )
    stats = zfi.citation_id_stats(parsed)
    keys = zfi.document_item_keys(parsed)
    from collections import Counter

    doi_count = Counter()
    for g in groups:
        if "dois" in g:
            doi_count.update(g["dois"])
        elif "loose" in g:
            doi_count.update([normalize_doi(g["loose"])])
    report = {
        "file": os.path.basename(path),
        "paragraphs": len(paras),
        "tables": len(body.findall(W + "tbl")),
        "fields": {"total": len(fields), "item": items, "bibl": bibs, "prefs": prefs},
        "structure_errors": parsed.structure_errors,
        "nested_citations": parsed.nested_count("ZOTERO_ITEM CSL_CITATION"),
        "style": style,
        "placeholder_groups": len([g for g in groups if "span" in g]),
        "loose_dois": len([g for g in groups if "loose" in g]),
        "unique_dois": len(doi_count),
        "placeholder_paras": sorted({g["para"] for g in groups}),
        "mixed_paras": mixed_paras,  # 同段既有域 + 纯文字占位
        "citation_ids": {k: v for k, v in stats.items() if k != "ids"},
        "doc_item_keys": len(keys),
        "groups": groups,
        "table_audit": table_column_audit(root, parsed),
    }
    if verbose:
        print(f"== {report['file']} ==")
        print(f"  段落 {report['paragraphs']}, 表格 {report['tables']}, 样式 {style}")
        print(f"  域: 共 {len(fields)} (引文 {items}, 文献表 {bibs}, 首选项 {prefs})")
        if report["structure_errors"] or report["nested_citations"]:
            print(
                f"  域结构异常 {report['structure_errors']}, 嵌套引文 {report['nested_citations']}"
            )
        print(
            f"  域外占位组 {report['placeholder_groups']} (散落 DOI {report['loose_dois']}), "
            f"唯一 DOI {report['unique_dois']}"
        )
        print(
            f"  占位段落 {report['placeholder_paras'][:12]}{'…' if len(report['placeholder_paras']) > 12 else ''}"
        )
        print(f"  混合段(既有域+占位): {report['mixed_paras'] or '无'}")
        print(
            f"  citationID {stats['count']} 个 / 唯一 {stats['unique']} / 重复 {len(stats['duplicates'])}"
            f" / 最大数字 {stats['max_numeric']}"
        )
        print(f"  文档引用的库条目 key: {len(keys)}")
        for t in report["table_audit"]:
            interesting = [
                c for c in t["columns"] if c["with_field"] or c["doi_cells"] or c["citation_cells"]
            ]
            if interesting:
                desc = "; ".join(
                    f'列{c["index"]}"{c["header"]}" 域{c["with_field"]}/{c["cells"]} '
                    f"DOI格{c['doi_cells']} 编号格{c['citation_cells']} 数据格{c['number_cells']}"
                    for c in interesting
                )
                print(f"  表{t['table']} ({t['rows']}×{t['cols']}): {desc}")
    return report


# ---------------------------------------------------------------- 修补
def _rewrite_field(els, obj):
    full = zfi.field_instr_text(els)
    marker = "CSL_CITATION"
    head = full[: full.index(marker) + len(marker)]
    replace_instruction(els, head + " " + json.dumps(obj, ensure_ascii=False))


def _save(root, path):
    save_document(root, path, path)


def renumber_citation_ids(
    path, new_display_texts=None, from_docx=None, prefix="sis", dry_run=False
):
    """把新插入的域的 citationID 重编到隔离命名空间（历史文档专用）。

    v3 起 insert_zotero_fields 已自动避让既有编号，正常新任务不需要本函数。
    它用于**已按旧版（从 0 编号）插入过的文档**：把新域的编号改成 prefix+序号，
    避免与旧域的 citN 世代重叠。

    识别"新域"的两种方式：
      - new_display_texts: 精确匹配显示文本集合（最可靠：从原稿重新提取的占位组文本）
      - from_docx: 给原稿路径，凡显示文本在原稿中不存在于任何既有域的，判为新域
    返回重编号数量。
    """
    root = load_root(path)
    fields = zfi.collect_fields(root)
    new_disp = None if new_display_texts is None else set(new_display_texts)
    if from_docx and new_disp is None:
        oroot = load_root(from_docx)
        old_disp = {d for _, d, _ in zfi.collect_fields(oroot) if d}
        new_disp = {d for _, d, _ in fields if d} - old_disp
    next_id = zfi._citation_id_start(zfi.citation_id_stats(root)["ids"], prefix=prefix)
    n = 0
    for els, disp, _ in fields:
        if "ZOTERO_ITEM CSL_CITATION" not in zfi.field_instr_text(els):
            continue
        if new_disp is not None and disp not in new_disp:
            continue
        obj = zfi.field_payload(els)
        if not obj:
            continue
        obj["citationID"] = f"{prefix}{next_id + n}"
        n += 1
        if not dry_run:
            _rewrite_field(els, obj)
    if n and not dry_run:
        _save(root, path)
    return n


def relink_item_keys(path, replacements, dry_run=False):
    """把域内 citationItems[].uris 的 itemKey 按 {旧: 新} 重写（修死链）。

    场景：文档里既有域指向的库条目已被删除，Zotero 刷新后会显示 [item removed]。
    用库内同 DOI 的规范条目 key 替换即可，无需重建条目。
    返回被重写的域数。
    """
    root = load_root(path)
    n = 0
    for els, _, _ in zfi.collect_fields(root):
        t = zfi.field_instr_text(els)
        if "ZOTERO_ITEM CSL_CITATION" not in t:
            continue
        obj = zfi.field_payload(els)
        if not obj:
            continue
        touched = False
        for ci in obj.get("citationItems", []):
            for k, uri in enumerate(ci.get("uris", [])):
                key = str(uri).rsplit("/items/", 1)[-1]
                if key in replacements:
                    ci["uris"][k] = str(uri).rsplit("/items/", 1)[0] + "/items/" + replacements[key]
                    touched = True
        if touched:
            n += 1
            if not dry_run:
                _rewrite_field(els, obj)
    if n and not dry_run:
        _save(root, path)
    return n


# ---------------------------------------------------------------- 转换后对账
def audit_output(src_docx, out_docx, item_mapping, verbose=True):
    """逐组对账：从原稿重新推导占位组，与输出域一一核对。

    比"数域个数"强得多：能发现整组漏转、显示文本被改写、DOI 集合错位、
    URI 指向了错误条目、组内条目数不对。
    返回 total/matched/problems 对账结果，以及 structure_errors/nested_citations 域结构计数。
    """
    key_of_doi = {
        normalize_doi(v.get("doi") or v.get("csl", {}).get("DOI")): v.get("itemKey")
        for k, v in item_mapping.items()
        if v.get("itemKey")
    }
    doi_by_key = {
        v.get("itemKey"): normalize_doi(v.get("doi") or v.get("csl", {}).get("DOI"))
        for v in item_mapping.values()
        if v.get("itemKey")
    }

    def _doi_of(ci):
        """itemData 的 DOI 优先；缺失时回退按 itemKey 从映射取。"""
        d = normalize_doi(ci.get("itemData", {}).get("DOI"))
        if not d:
            d = doi_by_key.get((ci.get("uris") or [""])[0].rsplit("/items/", 1)[-1], "")
        return d

    sroot = load_root(src_docx)
    oroot = load_root(out_docx)
    parsed = parse_fields(oroot)
    nested_citations = parsed.nested_count("ZOTERO_ITEM CSL_CITATION")
    problems = []
    total = 0
    # 原稿：按段落重建占位组
    src_groups = []
    for i, txt in zfi.outside_field_paragraphs(sroot):
        for s, e, g, dois in extract_doi_groups(txt):
            src_groups.append((i, g, dois))
    # 输出：按**显示文本多重集**匹配，不按段落号。
    # insert 会在文首插入首选项域段落，输出与原稿的段落序号整体位移，
    # 按段落号配对会全线错位（v2 实测：每组都报"未找到对应域"）。
    pool = {}
    for els, disp, para in zfi.collect_fields(parsed):
        if "ZOTERO_ITEM CSL_CITATION" not in zfi.field_instr_text(els):
            continue
        pool.setdefault(disp, []).append(els)
    for para, gtext, gdois in src_groups:
        total += 1
        cand = pool.get(gtext) or []
        if not cand:
            problems.append({"para": para, "text": gtext[:80], "why": "未找到对应域"})
            continue
        els = cand.pop(0)  # 消耗一个：同文本的重复占位按出现顺序配对
        obj = zfi.field_payload(els) or {}
        items = obj.get("citationItems", [])
        fdois = sorted({_doi_of(ci) for ci in items})
        if fdois != sorted(set(gdois)):
            problems.append(
                {
                    "para": para,
                    "text": gtext[:80],
                    "why": f"DOI 集合不符: {fdois} vs {sorted(set(gdois))}",
                }
            )
            continue
        bad_uri = [
            (ci.get("uris") or [""])[0]
            for ci in items
            if not ci.get("uris")
            or ci["uris"][0].rsplit("/items/", 1)[-1] != key_of_doi.get(_doi_of(ci))
        ]
        if bad_uri:
            problems.append({"para": para, "text": gtext[:80], "why": f"URI 指向不符: {bad_uri}"})
            continue
        if obj.get("properties", {}).get("dontUpdate") is not False:
            problems.append({"para": para, "text": gtext[:80], "why": "dontUpdate != false"})
            continue
    matched = total - len(problems)
    if verbose:
        print(f"逐组对账: {matched}/{total} 一致, 异常 {len(problems)}")
        if parsed.structure_errors or nested_citations:
            print(f"域结构异常 {parsed.structure_errors}, 嵌套引文 {nested_citations}")
        for p in problems[:10]:
            print("   ", p)
    return {
        "total": total,
        "matched": matched,
        "problems": problems,
        "structure_errors": parsed.structure_errors,
        "nested_citations": nested_citations,
    }


# ---------------------------------------------------------------- CLI
def main(argv=None):
    """Legacy argv includes the script name; console entry uses cli()."""
    parser = argparse.ArgumentParser(description="扫描、审计及修补 Zotero DOCX 域代码")
    commands = parser.add_subparsers(dest="command", required=True)
    scan = commands.add_parser("scan", help="扫描文档")
    scan.add_argument("document")
    scan.add_argument("--json", action="store_true")
    audit = commands.add_parser("audit", help="逐组对账；异常时退出码为 1")
    audit.add_argument("source")
    audit.add_argument("output")
    audit.add_argument("mapping")
    audit.add_argument("--json", action="store_true")
    renumber = commands.add_parser("renumber", help="重编 citationID")
    renumber.add_argument("document")
    renumber.add_argument("--prefix", default="sis")
    renumber.add_argument("--from", dest="source")
    renumber.add_argument("--dry-run", action="store_true")
    relink = commands.add_parser("relink", help="重链接条目 key")
    relink.add_argument("document")
    relink.add_argument("replacements", type=json.loads)
    relink.add_argument("--dry-run", action="store_true")
    args = parser.parse_args(argv[1:] if argv is not None else None)
    try:
        if args.command == "scan":
            report = scan_document(args.document, verbose=not args.json)
            if args.json:
                # Escaped Unicode keeps piped JSON valid on legacy Windows code pages.
                print(json.dumps(report, ensure_ascii=True, indent=2))
        elif args.command == "audit":
            with open(args.mapping, encoding="utf-8") as source:
                mapping = json.load(source)
            report = audit_output(args.source, args.output, mapping, verbose=not args.json)
            if args.json:
                print(json.dumps(report, ensure_ascii=True, indent=2))
            return (
                1
                if (report["problems"] or report["structure_errors"] or report["nested_citations"])
                else 0
            )
        elif args.command == "renumber":
            count = renumber_citation_ids(
                args.document, from_docx=args.source, prefix=args.prefix, dry_run=args.dry_run
            )
            print(f"{'待重编号' if args.dry_run else '重编号'} {count} 个域")
        else:
            if not isinstance(args.replacements, dict) or not all(
                isinstance(k, str) and isinstance(v, str) and k and v
                for k, v in args.replacements.items()
            ):
                raise ValueError("replacements must be a JSON object mapping nonempty keys to keys")
            count = relink_item_keys(args.document, args.replacements, dry_run=args.dry_run)
            print(f"{'待重链接' if args.dry_run else '重链接'} {count} 个域")
    except (OSError, ValueError, KeyError, zipfile.BadZipFile, etree.XMLSyntaxError) as error:
        print(f"错误: {error}", file=sys.stderr)
        return 2
    return 0


def cli():
    return main()


if __name__ == "__main__":
    sys.exit(main(sys.argv))
