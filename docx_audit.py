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
import shutil
import sys
import os

from lxml import etree

try:
    import zotero_field_insert as zfi
except ImportError:  # 直接以脚本方式运行时兜底
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    import zotero_field_insert as zfi

W = zfi.W
SCHEMA = zfi.SCHEMA

# 与生成器共用的 DOI 词元（老式 Elsevier DOI 内含 "(98)" 年份段）
DOI_TOKEN = zfi._DOI_TOKEN
# 带捕获组：fullmatch 后 group(1) 即 DOI 本体（DOI_TOKEN 内部无捕获组）
DOI_RE = re.compile('(' + DOI_TOKEN + ')', re.I)
DOI_IN_GROUP = re.compile(r'(?:doi[:：]\s*)?(' + DOI_TOKEN + r')', re.I)
NUM_ONLY = re.compile(r'^\s*[\[(]?\s*\d{1,3}(\s*[-–,;]\s*\d{1,3})*\s*[\])]?\s*$')
PAREN_NUMBER = re.compile(r'[\[(]\s*\d{1,3}(\s*[,;–-]\s*\d{1,3})*\s*[\])]')


# ---------------------------------------------------------------- 提取占位
def extract_doi_groups(text):
    """从**域外**文本里提取 `(doi:…)` 占位组（括号配平，老式 DOI 完整识别）。

    返回 [(start, end, group_text, [doi, ...])]。

    为什么不用 `\\(…\\)` 正则：老式 DOI `10.1016/S0306-4530(98)00014-6` 内含括号，
    非配平的正则会在第一个 ')' 处截断，整组漏转（实测漏 3 组）。
    为什么不用 `10\\.\\d+/[^\\s,;)]+` 提取 DOI：它同样在内层 ')' 处截断，
    得到的 `10.1016/s0306-4530` 查不到库内条目。
    """
    out = []
    for m in re.finditer(r'[(（]\s*doi[:：]\s*10\.', text, re.I):
        i = m.start()
        depth = 0
        end = None
        for j in range(i, len(text)):
            c = text[j]
            if c in '((（':
                depth += 1
            elif c in '))）':
                depth -= 1
                if depth == 0:
                    end = j + 1
                    break
        if end is None:
            continue
        group = text[i:end]
        parts = [p.strip() for p in re.split(r'[;；]', group[1:-1])]
        dois = []
        ok = bool(parts)
        for p in parts:
            mm = DOI_RE.fullmatch(p.lstrip('doi: ').strip() if p.lower().startswith('doi') else p)
            if mm is None:
                mm = DOI_IN_GROUP.fullmatch(p)
            if mm is None:
                ok = False
                break
            dois.append(mm.group(1).lower())
        if ok and dois:
            out.append((i, end, group, dois))
    return out


def extract_loose_dois(text, covered=()):
    """括号组之外的散落 DOI（返回 [(doi, context)]）。覆盖区间用于排除已提取的组。"""
    parts = []
    last = 0
    for s, e in covered:
        parts.append(text[last:s]); last = e
    parts.append(text[last:])
    rest = ''.join(parts)
    out = []
    for m in DOI_RE.finditer(rest):
        out.append((m.group(0).rstrip('.,;'), rest[max(0, m.start() - 60):m.end() + 20]))
    return out


# ---------------------------------------------------------------- 扫描
def _cell_outside_text(tc, depth0=True):
    d = 0
    parts = []
    for r in tc.iter(W + 'r'):
        for ch in r:
            ln = etree.QName(ch).localname
            if ln == 'fldChar':
                ft = ch.get(W + 'fldCharType')
                if ft == 'begin':
                    d += 1
                elif ft == 'end':
                    d = max(0, d - 1)
            elif ln == 't' and (not depth0 or d == 0):
                parts.append(ch.text or '')
    return ''.join(parts)


def _tc_has_field(tc):
    return any(c.get(W + 'fldCharType') == 'begin' for c in tc.iter(W + 'fldChar'))


def table_column_audit(root):
    """逐表逐列审计：区分"引文列"与"数据列"。

    表格里"纯数字格"绝大多数是数据（年龄 6–8、周期 14 天…），不是引文；
    真正要转的是：整格 DOI 占位、整格编号引文（"(1)" / "[12]"）且同行同列已是域。
    返回 [{table, rows, cols, header, columns:[{index, header, cells, with_field,
    doi_cells, number_cells, citation_cells}]}]
    """
    body = root.find(W + 'body')
    report = []
    for ti, tbl in enumerate(body.findall(W + 'tbl')):
        rows = tbl.findall(W + 'tr')
        if not rows:
            continue
        ncols = max(len(r.findall(W + 'tc')) for r in rows)
        header = [_cell_outside_text(tc).strip()[:24] for tc in rows[0].findall(W + 'tc')]
        cols = []
        for ci in range(ncols):
            cells = [r.findall(W + 'tc')[ci] for r in rows if ci < len(r.findall(W + 'tc'))]
            with_field = doi_cells = number_cells = citation_cells = 0
            samples = {'doi': [], 'number': [], 'citation': []}
            for ri, tc in enumerate(cells):
                if _tc_has_field(tc):
                    with_field += 1
                    continue
                t = _cell_outside_text(tc).strip()
                if not t:
                    continue
                if DOI_RE.search(t):
                    doi_cells += 1; samples['doi'].append([ri, t[:60]])
                elif NUM_ONLY.match(t):
                    number_cells += 1; samples['number'].append([ri, t[:30]])
                elif PAREN_NUMBER.search(t) and len(t) < 60:
                    citation_cells += 1; samples['citation'].append([ri, t[:40]])
            cols.append({'index': ci, 'header': header[ci] if ci < len(header) else '',
                         'cells': len(cells), 'with_field': with_field,
                         'doi_cells': doi_cells, 'number_cells': number_cells,
                         'citation_cells': citation_cells, 'samples': samples})
        report.append({'table': ti, 'rows': len(rows), 'cols': ncols,
                       'header': header, 'columns': cols})
    return report


def load_root(path):
    """读 word/document.xml 并返回 lxml root。"""
    return etree.fromstring(zipfile.ZipFile(path).read('word/document.xml'))


def scan_document(path, verbose=True):
    """转换前全貌扫描。返回可直接打印/落盘的 dict。"""
    root = load_root(path)
    body = root.find(W + 'body')
    paras = body.findall(W + 'p')
    fields = zfi.collect_fields(root)
    items = bibs = prefs = 0
    style = None
    for els, _, _ in fields:
        t = zfi.field_instr_text(els)
        if 'ZOTERO_ITEM CSL_CITATION' in t:
            items += 1
        elif 'ZOTERO_BIBL' in t:
            bibs += 1
        elif 'ZOTERO_DOCUMENT_PREFERENCES' in t:
            prefs += 1
            obj = zfi.field_payload(els) or {}
            style = (obj.get('style') or {}).get('styleID')
    # 域外占位组（跨段深度感知，BIBL 内的文献不会混进来）
    groups = []
    for i, txt in zfi.outside_field_paragraphs(root):
        gs = extract_doi_groups(txt)
        for s, e, g, dois in gs:
            groups.append({'para': i, 'span': [s, e], 'text': g, 'dois': dois})
        loose = extract_loose_dois(txt, [(s, e) for s, e, _, _ in gs])
        for d, ctx in loose:
            groups.append({'para': i, 'loose': d, 'context': ctx})
    mixed_paras = sorted({g['para'] for g in groups if 'span' in g
                          and any(f[2] == g['para'] for f in fields)})
    stats = zfi.citation_id_stats(root)
    keys = zfi.document_item_keys(root)
    from collections import Counter
    doi_count = Counter()
    for g in groups:
        if 'dois' in g:
            doi_count.update(g['dois'])
    report = {
        'file': os.path.basename(path),
        'paragraphs': len(paras),
        'tables': len(body.findall(W + 'tbl')),
        'fields': {'total': len(fields), 'item': items, 'bibl': bibs, 'prefs': prefs},
        'style': style,
        'placeholder_groups': len([g for g in groups if 'span' in g]),
        'loose_dois': len([g for g in groups if 'loose' in g]),
        'unique_dois': len(doi_count),
        'placeholder_paras': sorted({g['para'] for g in groups}),
        'mixed_paras': mixed_paras,      # 同段既有域 + 纯文字占位
        'citation_ids': {k: v for k, v in stats.items() if k != 'ids'},
        'doc_item_keys': len(keys),
        'groups': groups,
        'table_audit': table_column_audit(root),
    }
    if verbose:
        print(f"== {report['file']} ==")
        print(f"  段落 {report['paragraphs']}, 表格 {report['tables']}, 样式 {style}")
        print(f"  域: 共 {len(fields)} (引文 {items}, 文献表 {bibs}, 首选项 {prefs})")
        print(f"  域外占位组 {report['placeholder_groups']} (散落 DOI {report['loose_dois']}), "
              f"唯一 DOI {report['unique_dois']}")
        print(f"  占位段落 {report['placeholder_paras'][:12]}{'…' if len(report['placeholder_paras'])>12 else ''}")
        print(f"  混合段(既有域+占位): {report['mixed_paras'] or '无'}")
        print(f"  citationID {stats['count']} 个 / 唯一 {stats['unique']} / 重复 {len(stats['duplicates'])}"
              f" / 最大数字 {stats['max_numeric']}")
        print(f"  文档引用的库条目 key: {len(keys)}")
        for t in report['table_audit']:
            interesting = [c for c in t['columns'] if c['with_field'] or c['doi_cells'] or c['citation_cells']]
            if interesting:
                desc = '; '.join(
                    f"列{c['index']}\"{c['header']}\" 域{c['with_field']}/{c['cells']} "
                    f"DOI格{c['doi_cells']} 编号格{c['citation_cells']} 数据格{c['number_cells']}"
                    for c in interesting)
                print(f"  表{t['table']} ({t['rows']}×{t['cols']}): {desc}")
    return report


# ---------------------------------------------------------------- 修补
def _rewrite_field(els, obj):
    full = zfi.field_instr_text(els)
    marker = 'CSL_CITATION'
    head = full[:full.index(marker) + len(marker)]
    els[0].text = head + ' ' + json.dumps(obj, ensure_ascii=False)
    for e in els[1:]:
        e.text = ''


def _save(root, path):
    xml = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(path + '.tmp', 'w', zipfile.ZIP_DEFLATED) as zout:
        for it in zfinfo_iter(zin):
            zout.writestr(it, xml if it.filename == 'word/document.xml' else zin.read(it.filename))
    shutil.move(path + '.tmp', path)


def zfinfo_iter(zin):
    return zin.infolist()


def renumber_citation_ids(path, new_display_texts=None, from_docx=None, prefix='sis',
                          dry_run=False):
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
    new_disp = set(new_display_texts or [])
    if from_docx and not new_disp:
        oroot = etree.fromstring(zipfile.ZipFile(from_docx).read('word/document.xml'))
        old_disp = {d for _, d, _ in zfi.collect_fields(oroot) if d}
        new_disp = {d for _, d, _ in fields if d} - old_disp
    n = 0
    for els, disp, _ in fields:
        if 'ZOTERO_ITEM CSL_CITATION' not in zfi.field_instr_text(els):
            continue
        if new_disp and disp not in new_disp:
            continue
        obj = zfi.field_payload(els)
        if not obj:
            continue
        obj['citationID'] = f'{prefix}{n}'
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
        if 'ZOTERO_ITEM CSL_CITATION' not in t:
            continue
        obj = zfi.field_payload(els)
        if not obj:
            continue
        touched = False
        for ci in obj.get('citationItems', []):
            for k, uri in enumerate(ci.get('uris', [])):
                key = str(uri).rsplit('/items/', 1)[-1]
                if key in replacements:
                    ci['uris'][k] = str(uri).rsplit('/items/', 1)[0] + '/items/' + replacements[key]
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
    返回 {'total','matched','problems':[...]}
    """
    doi2key = zfi.build_doi_matcher(item_mapping)
    key_of_doi = {str(v.get('doi', '')).lower(): v.get('itemKey')
                  for k, v in item_mapping.items() if v.get('itemKey')}
    doi_by_key = {v.get('itemKey'): str(v.get('doi', '')).lower()
                  for v in item_mapping.values() if v.get('itemKey')}

    def _doi_of(ci):
        """itemData 的 DOI 优先；缺失时回退按 itemKey 从映射取。"""
        d = (ci.get('itemData', {}).get('DOI') or '').strip().lower()
        if not d:
            d = doi_by_key.get(ci.get('uris', [''])[0].rsplit('/items/', 1)[-1], '')
        return d
    sroot = load_root(src_docx)
    oroot = load_root(out_docx)
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
    for els, disp, para in zfi.collect_fields(oroot):
        if 'ZOTERO_ITEM CSL_CITATION' not in zfi.field_instr_text(els):
            continue
        pool.setdefault(disp, []).append(els)
    for para, gtext, gdois in src_groups:
        total += 1
        cand = pool.get(gtext) or []
        if not cand:
            problems.append({'para': para, 'text': gtext[:80], 'why': '未找到对应域'})
            continue
        els = cand.pop(0)          # 消耗一个：同文本的重复占位按出现顺序配对
        obj = zfi.field_payload(els) or {}
        items = obj.get('citationItems', [])
        fdois = sorted({_doi_of(ci) for ci in items})
        if fdois != sorted(set(gdois)):
            problems.append({'para': para, 'text': gtext[:80],
                             'why': f'DOI 集合不符: {fdois} vs {sorted(set(gdois))}'})
            continue
        bad_uri = [ci['uris'][0] for ci in items
                   if ci['uris'][0].rsplit('/items/', 1)[-1] != key_of_doi.get(_doi_of(ci))]
        if bad_uri:
            problems.append({'para': para, 'text': gtext[:80], 'why': f'URI 指向不符: {bad_uri}'})
            continue
        if obj.get('properties', {}).get('dontUpdate') is not False:
            problems.append({'para': para, 'text': gtext[:80], 'why': 'dontUpdate != false'})
            continue
    matched = total - len(problems)
    if verbose:
        print(f"逐组对账: {matched}/{total} 一致, 异常 {len(problems)}")
        for p in problems[:10]:
            print('   ', p)
    return {'total': total, 'matched': matched, 'problems': problems}


# ---------------------------------------------------------------- CLI
def main(argv):
    if len(argv) < 2:
        print(__doc__)
        return 0
    cmd = argv[1]
    if cmd == 'scan':
        path = argv[2]
        rep = scan_document(path, verbose='--json' not in argv)
        if '--json' in argv:
            rep.pop('groups', None)
            print(json.dumps(rep, ensure_ascii=False, indent=1))
    elif cmd == 'audit':
        mapping = json.load(open(argv[4])) if len(argv) > 4 else {}
        audit_output(argv[2], argv[3], mapping)
    elif cmd == 'renumber':
        prefix = 'sis'
        if '--prefix' in argv:
            prefix = argv[argv.index('--prefix') + 1]
        frm = argv[argv.index('--from') + 1] if '--from' in argv else None
        n = renumber_citation_ids(argv[2], from_docx=frm, prefix=prefix)
        print(f'重编号 {n} 个域 -> {prefix}*')
    elif cmd == 'relink':
        n = relink_item_keys(argv[2], json.loads(argv[3]))
        print(f'重链接 {n} 个域')
    else:
        print(__doc__)
    return 0


if __name__ == '__main__':
    sys.exit(main(sys.argv))
