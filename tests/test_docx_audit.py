#!/usr/bin/env python3
"""v3 专项测试：docx_audit + zotero_field_insert 的新行为。

每个用例都对应一次真实踩坑（见 SKILL.md「v3 新增」）：
  - 老式 Elsevier DOI 内含括号，整组曾整组漏转 / DOI 被截断
  - 跨段域（BIBL 文献表）被逐段重置深度误判成域外残留
  - 新域 citationID 与既有 citN 世代重叠
  - 形似 DOI 却解析不出的分片被静默截断，证据凭空消失
  - 表格"纯数字格"（数据）被误当引文
  - 死链（指向已删除条目）必须能被发现与重链接

运行: python3 -m unittest discover -s tests
"""
import json
import os
import sys
import tempfile
import unittest
import zipfile

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from docx import Document

import docx_audit as da
import zotero_field_insert as zfi
import reference_ingest as ri


# ----------------------------------------------------------------- 工具
def csl(title, year, family='Smith', container='J Test', vol='1', page='1-9', doi=None):
    out = {'type': 'article-journal', 'title': title, 'container-title': container,
           'volume': vol, 'page': page,
           'author': [{'family': family, 'given': 'A'}],
           'issued': {'date-parts': [[year]]}}
    if doi:
        out['DOI'] = doi
    return out


def mapping(*dois):
    m = {}
    for i, d in enumerate(dois, 1):
        key = f'ref_{d}'
        m[key] = {'itemKey': f'KEY{i:03d}ABC',
                  'uri': f'http://zotero.org/users/1/items/KEY{i:03d}ABC',
                  'doi': d, 'csl': csl(d, 2000 + i, doi=d)}
    return m


class TmpMixin(unittest.TestCase):
    def setUp(self):
        self._td = tempfile.TemporaryDirectory()
        self.td = self._td.name

    def tearDown(self):
        self._td.cleanup()

    def path(self, name):
        return os.path.join(self.td, name)


# ----------------------------------------------------------------- 占位提取
class TestExtractDoiGroups(unittest.TestCase):
    def test_old_style_elsevier_doi_with_parens(self):
        """老式 DOI 10.1016/S0306-4530(98)00014-6：括号配平 + DOI 完整。"""
        spans = da.extract_doi_groups('(doi:10.1016/S0306-4530(98)00014-6)')
        self.assertEqual(len(spans), 1)
        self.assertEqual(spans[0][3], ['10.1016/s0306-4530(98)00014-6'])

    def test_multi_doi_group_with_old_style_mixed(self):
        t = 'claim (doi:10.1016/S0306-4530(98)00014-6; doi:10.1002/cpz1.208) end'
        spans = da.extract_doi_groups(t)
        self.assertEqual(len(spans), 1)
        self.assertEqual(sorted(spans[0][3]),
                         ['10.1002/cpz1.208', '10.1016/s0306-4530(98)00014-6'])

    def test_trailing_paren_not_part_of_doi(self):
        """朴素的 [^\\s,;)]+ 会把分组右括号吞进 DOI。"""
        spans = da.extract_doi_groups('(doi:10.1002/cpz1.208)')
        self.assertEqual(spans[0][3], ['10.1002/cpz1.208'])

    def test_non_doi_parens_ignored(self):
        self.assertEqual(da.extract_doi_groups('a (2020) b (n = 3) c'), [])

    def test_doi_regex_never_truncates_old_style(self):
        m = da.DOI_RE.search('10.1016/S0306-4530(98)00014-6.')
        self.assertEqual(m.group(0).rstrip('.'), '10.1016/S0306-4530(98)00014-6')

    def test_matcher_agrees_with_extractor(self):
        """提取器认出的组，生成器的匹配器也必须认得出（两侧口径一致）。"""
        d2k = {'10.1016/s0306-4530(98)00014-6': 'a', '10.1002/cpz1.208': 'b'}
        for t in ['(doi:10.1016/S0306-4530(98)00014-6)',
                  '(doi:10.1002/cpz1.208; doi:10.1016/S0306-4530(98)00014-6)']:
            spans = da.extract_doi_groups(t)
            self.assertTrue(spans)
            m = zfi.find_doi_citations(t, d2k, None)
            self.assertEqual(len(m), len(spans), t)

    def test_standalone_cell_with_wrapping_parens(self):
        hit = zfi.match_standalone_doi(
            '(doi:10.1016/S0306-4530(98)00014-6)',
            {'10.1016/s0306-4530(98)00014-6': 'a'})
        self.assertIsNotNone(hit)
        self.assertEqual(hit[0], 'a')


# ----------------------------------------------------------------- 域收集
class TestCollectFields(TmpMixin):
    def test_cross_paragraph_field_is_collected_once(self):
        """BIBL 域跨多段：必须收成一个域，不能按段切碎或漏收。"""
        d = Document()
        d.add_paragraph('body text')
        for i in range(5):
            d.add_paragraph(f'[{i + 1}] Some reference with https://doi.org/10.1000/xyz{i}.')
        p = d.add_paragraph()
        src = self.path('x.docx'); d.save(src)
        # 手工把 5 段包进一个 BIBL 域（模拟 Zotero 写入的跨段文献表）
        from lxml import etree as ET
        z = zipfile.ZipFile(src); root = ET.fromstring(z.read('word/document.xml'))
        W = zfi.W
        paras = root.find(W + 'body').findall(W + 'p')[1:6]
        begin = ET.SubElement(ET.SubElement(paras[0], W + 'r'), W + 'fldChar')
        begin.set(W + 'fldCharType', 'begin')
        it = ET.SubElement(ET.SubElement(paras[0], W + 'r'), W + 'instrText')
        it.text = ' ADDIN ZOTERO_BIBL {"uncited":[],"omitted":[],"custom":[]} CSL_BIBLIOGRAPHY '
        sep = ET.SubElement(ET.SubElement(paras[0], W + 'r'), W + 'fldChar')
        sep.set(W + 'fldCharType', 'separate')
        end = ET.SubElement(ET.SubElement(paras[-1], W + 'r'), W + 'fldChar')
        end.set(W + 'fldCharType', 'end')
        xml = ET.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
        with zipfile.ZipFile(src) as zin, zipfile.ZipFile(self.path('y.docx'), 'w') as zo:
            for i in zin.infolist():
                zo.writestr(i, xml if i.filename == 'word/document.xml' else zin.read(i.filename))
        root2 = ET.fromstring(zipfile.ZipFile(self.path('y.docx')).read('word/document.xml'))
        fields = zfi.collect_fields(root2)
        self.assertEqual(len(fields), 1)
        self.assertIn('ZOTERO_BIBL', zfi.field_instr_text(fields[0][0]))

    def test_residual_scan_ignores_bibliography_dois(self):
        """文献表里的 https://doi.org/ 不得被误报成"域外残留占位"。"""
        d = Document()
        d.add_paragraph('Real placeholder (doi:10.1000/abc).')
        for i in range(3):
            d.add_paragraph(f'[{i + 1}] Ref. https://doi.org/10.2000/ref{i}.')
        src = self.path('b.docx'); d.save(src)
        from lxml import etree as ET
        z = zipfile.ZipFile(src); root = ET.fromstring(z.read('word/document.xml'))
        W = zfi.W
        paras = root.find(W + 'body').findall(W + 'p')[1:4]
        b = ET.SubElement(ET.SubElement(paras[0], W + 'r'), W + 'fldChar')
        b.set(W + 'fldCharType', 'begin')
        it = ET.SubElement(ET.SubElement(paras[0], W + 'r'), W + 'instrText')
        it.text = ' ADDIN ZOTERO_BIBL {"uncited":[]} CSL_BIBLIOGRAPHY '
        s = ET.SubElement(ET.SubElement(paras[0], W + 'r'), W + 'fldChar')
        s.set(W + 'fldCharType', 'separate')
        e = ET.SubElement(ET.SubElement(paras[2], W + 'r'), W + 'fldChar')
        e.set(W + 'fldCharType', 'end')
        xml = ET.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
        with zipfile.ZipFile(src) as zin, zipfile.ZipFile(self.path('c.docx'), 'w') as zo:
            for i in zin.infolist():
                zo.writestr(i, xml if i.filename == 'word/document.xml' else zin.read(i.filename))
        root2 = ET.fromstring(zipfile.ZipFile(self.path('c.docx')).read('word/document.xml'))
        res = zfi.residual_placeholders(root2)
        self.assertEqual([r['para'] for r in res], [0])   # 只剩正文那一处

    def test_instr_text_split_across_runs(self):
        """Word 会把一个域指令拆成多个 instrText：域级拼接后必须能解析。"""
        d = Document(); d.add_paragraph('x'); src = self.path('d.docx'); d.save(src)
        from lxml import etree as ET
        z = zipfile.ZipFile(src); root = ET.fromstring(z.read('word/document.xml'))
        W = zfi.W
        p = root.find(W + 'body').findall(W + 'p')[0]
        first = True
        for kind, txt in [('begin', None), (None, '{"citationID":"a1","citationItems":['),
                          (None, '{"id":1,"uris":["u"],"itemData":{"type":"book",'),
                          (None, '"title":"t","author":[],"issued":{"date-parts":[[2000]]}}}]}'),
                          ('end', None)]:
            r = ET.SubElement(p, W + 'r')
            if kind:
                c = ET.SubElement(r, W + 'fldChar'); c.set(W + 'fldCharType', kind)
            else:
                t = ET.SubElement(r, W + 'instrText')
                t.text = (' ADDIN ZOTERO_ITEM CSL_CITATION ' if first else '') + txt
                first = False
        xml = ET.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
        with zipfile.ZipFile(src) as zin, zipfile.ZipFile(self.path('e.docx'), 'w') as zo:
            for i in zin.infolist():
                zo.writestr(i, xml if i.filename == 'word/document.xml' else zin.read(i.filename))
        root2 = ET.fromstring(zipfile.ZipFile(self.path('e.docx')).read('word/document.xml'))
        obj = zfi.field_payload(zfi.collect_fields(root2)[0][0])
        self.assertEqual(obj['citationID'], 'a1')


# ----------------------------------------------------------------- citationID
class TestCitationIdIsolation(TmpMixin):
    def test_start_offset(self):
        self.assertEqual(zfi._citation_id_start(['cit0', 'cit7', 'cit146']), 147)
        self.assertEqual(zfi._citation_id_start([]), 0)
        self.assertEqual(zfi._citation_id_start(['IjW2SbGH', 'cit3']), 4)
        self.assertEqual(zfi._citation_id_start(['tcit0', 'tcit9'], prefix='tcit'), 10)

    def test_insert_avoids_existing_ids(self):
        """源文档已有 cit0/cit1 时，新域不得再用 0/1。"""
        d = Document()
        d.add_paragraph('Mixed para with existing field (doi:10.1000/new1).')
        src = self.path('src.docx'); d.save(src)
        # 先造一个含 2 个既有 cit 域的文档
        mp = mapping('10.1000/new1')
        mid = self.path('mid.docx')
        zfi.insert_zotero_fields(src, mid, (0, 1), [], mp)
        mp2 = mapping('10.1000/new1', '10.1000/new2')
        d2 = Document()
        d2.add_paragraph('First (doi:10.1000/new1).')
        d2.add_paragraph('Second (doi:10.1000/new2).')
        src2 = self.path('src2.docx'); d2.save(src2)
        out = self.path('out.docx')
        n = zfi.insert_zotero_fields(src2, out, (0, 2), [], mp2)[0]
        self.assertEqual(n, 2)                       # 返回值仍是"新增域数"
        root = ET.fromstring(zipfile.ZipFile(out).read('word/document.xml')) \
            if False else __import__('lxml.etree', fromlist=['etree']).fromstring(
                zipfile.ZipFile(out).read('word/document.xml'))
        st = zfi.citation_id_stats(root)
        self.assertEqual(st['duplicates'], {})
        self.assertEqual(sorted(st['ids']), ['cit0', 'cit1'])

    def test_renumber_legacy_document(self):
        d = Document()
        d.add_paragraph('Old (doi:10.1000/a). And old (doi:10.1000/b).')
        src = self.path('src3.docx'); d.save(src)
        out = self.path('out3.docx')
        zfi.insert_zotero_fields(src, out, (0, 1), [], mapping('10.1000/a', '10.1000/b'))
        n = da.renumber_citation_ids(out, new_display_texts=['(doi:10.1000/a)'], prefix='sis')
        self.assertEqual(n, 1)
        root = __import__('lxml.etree', fromlist=['etree']).fromstring(
            zipfile.ZipFile(out).read('word/document.xml'))
        ids = zfi.citation_id_stats(root)['ids']
        self.assertIn('sis0', ids)
        self.assertIn('cit1', ids)


# ----------------------------------------------------------------- 匹配安全
class TestGroupSafety(unittest.TestCase):
    def test_broken_doi_part_skips_whole_group(self):
        """形似 DOI 却查不到的���片 -> 整组跳过，不能静默截断成只剩一条。"""
        d2k = {'10.1234/ax': 'ax'}
        self.assertEqual(zfi.find_doi_citations('(10.1234/ax; 10.unknown/z)', d2k, None), [])

    def test_author_year_mixed_group_still_converts(self):
        d2k = {'10.1234/ax': 'ax'}
        spans = zfi.find_doi_citations('(Gao et al., 2014; doi:10.1234/ax)', d2k, None)
        self.assertEqual(len(spans), 1)
        self.assertEqual(spans[0][3], ['ax'])


# ----------------------------------------------------------------- DOI 归一化
class TestNormalizeDoi(unittest.TestCase):
    def test_variants(self):
        for raw in ['10.1000/abc', 'https://doi.org/10.1000/abc', 'DOI: 10.1000/abc',
                    'doi:10.1000/abc', '10.1000/abc.', '(10.1000/abc)']:
            self.assertEqual(ri.normalize_doi(raw), '10.1000/abc', raw)

    def test_old_style_keeps_inner_parens(self):
        self.assertEqual(ri.normalize_doi('(doi:10.1016/S0306-4530(98)00014-6)'),
                         '10.1016/s0306-4530(98)00014-6')


# ----------------------------------------------------------------- 表格审计
class TestTableAudit(TmpMixin):
    def test_data_numbers_not_treated_as_citations(self):
        d = Document()
        t = d.add_table(rows=3, cols=3)
        t.cell(0, 0).text = 'Strain'; t.cell(0, 1).text = 'Age (W)'; t.cell(0, 2).text = 'References'
        t.cell(1, 1).text = '8'; t.cell(2, 1).text = '12–16'
        t.cell(1, 2).text = '(doi:10.1000/aa)'; t.cell(2, 2).text = '10–14'
        src = self.path('t.docx'); d.save(src)
        from lxml import etree as ET
        root = ET.fromstring(zipfile.ZipFile(src).read('word/document.xml'))
        rep = da.table_column_audit(root)
        col1 = [c for c in rep[0]['columns'] if c['index'] == 1][0]
        col2 = [c for c in rep[0]['columns'] if c['index'] == 2][0]
        self.assertEqual(col1['number_cells'], 2)      # 8 / 12–16 是数据
        self.assertEqual(col1['doi_cells'], 0)
        self.assertEqual(col2['doi_cells'], 1)        # (doi:10.1000/aa) 要转
        self.assertEqual(col2['number_cells'], 1)     # 10–14 是数据


# ----------------------------------------------------------------- 修补
class TestRepair(TmpMixin):
    def test_relink_dead_keys(self):
        d = Document()
        d.add_paragraph('Claim (Smith, 2007) and (Jones, 2010).')
        src = self.path('s.docx'); d.save(src)
        mp = {'smith': {'itemKey': 'DEADKEY1', 'uri': 'http://zotero.org/users/1/items/DEADKEY1',
                        'csl': csl('A', 2007, 'Smith')},
              'jones': {'itemKey': 'LIVEKEY2', 'uri': 'http://zotero.org/users/1/items/LIVEKEY2',
                        'csl': csl('B', 2010, 'Jones')}}
        out = self.path('o.docx')
        zfi.insert_zotero_fields(src, out, (0, 1), [], mp)
        n = da.relink_item_keys(out, {'DEADKEY1': 'LIVEKEY2'})
        self.assertEqual(n, 1)
        from lxml import etree as ET
        root = ET.fromstring(zipfile.ZipFile(out).read('word/document.xml'))
        self.assertNotIn('DEADKEY1', zfi.document_item_keys(root))

    def test_dead_key_survives_verify_and_is_reported(self):
        """valid_item_keys 不含某 key 时 verify 必须报 bad_uri（死链可被交付前发现）。"""
        d = Document(); d.add_paragraph('Claim (Smith, 2007).')
        src = self.path('s2.docx'); d.save(src)
        mp = {'smith': {'itemKey': 'GONE', 'uri': 'http://zotero.org/users/1/items/GONE',
                        'csl': csl('A', 2007, 'Smith')}}
        out = self.path('o2.docx')
        zfi.insert_zotero_fields(src, out, (0, 1), [], mp)
        ok, detail = zfi.verify(out, {'LIVE'})
        self.assertFalse(ok)
        self.assertEqual(detail['bad_uri'], 1)


# ----------------------------------------------------------------- 逐组对账
class TestAuditOutput(TmpMixin):
    def test_group_by_group_match(self):
        dois = ('10.1016/S0306-4530(98)00014-6', '10.1002/cpz1.208', '10.1093/chemse/26.7.905')
        d = Document()
        d.add_paragraph(f'Stress (doi:{dois[0]}).')
        d.add_paragraph(f'Both (doi:{dois[1]}; doi:{dois[2]}).')
        src = self.path('s3.docx'); d.save(src)
        out = self.path('o3.docx')
        n = zfi.insert_zotero_fields(src, out, (0, 2), [], mapping(*dois))[0]
        self.assertEqual(n, 2)
        rep = da.audit_output(src, out, mapping(*dois), verbose=False)
        self.assertEqual(rep['problems'], [])
        self.assertEqual(rep['total'], 2)

    def test_detects_wrong_uri(self):
        dois = ('10.1000/ok',)
        d = Document(); d.add_paragraph(f'Claim (doi:{dois[0]}).')
        src = self.path('s4.docx'); d.save(src)
        out = self.path('o4.docx')
        zfi.insert_zotero_fields(src, out, (0, 1), [], mapping(*dois))
        bad = {'ref_10.1000/ok': {'itemKey': 'AAA', 'uri': 'http://zotero.org/users/1/items/AAA',
                                  'doi': dois[0], 'csl': csl('t', 2001)}}
        rep = da.audit_output(src, out, bad, verbose=False)
        self.assertTrue(rep['problems'])
        self.assertIn('URI', rep['problems'][0]['why'])


# ----------------------------------------------------------------- scan
class TestScanDocument(TmpMixin):
    def test_scan_reports_everything_needed_before_conversion(self):
        d = Document()
        d.add_paragraph('Body with (doi:10.1000/xyz) and (doi:10.1016/S0306-4530(98)00014-6).')
        d.add_paragraph('Second (doi:10.1000/xyz).')
        src = self.path('scan.docx'); d.save(src)
        rep = da.scan_document(src, verbose=False)
        self.assertEqual(rep['placeholder_groups'], 3)
        self.assertEqual(rep['unique_dois'], 2)
        self.assertEqual(rep['style'], None)
        self.assertEqual(rep['fields']['total'], 0)

    def test_scan_after_conversion_shows_fields_and_no_residual(self):
        dois = ('10.1000/xyz', '10.1016/S0306-4530(98)00014-6')
        d = Document()
        d.add_paragraph(f'Body (doi:{dois[0]}).')
        d.add_paragraph(f'Old style (doi:{dois[1]}).')
        src = self.path('s5.docx'); d.save(src)
        out = self.path('o5.docx')
        zfi.insert_zotero_fields(src, out, (0, 2), [], mapping(*dois),
                                 style_id='http://www.zotero.org/styles/ieee')
        rep = da.scan_document(out, verbose=False)
        self.assertEqual(rep['fields']['item'], 2)
        self.assertEqual(rep['placeholder_groups'], 0)   # 占位已全部进域
        self.assertEqual(rep['style'], 'http://www.zotero.org/styles/ieee')
        self.assertEqual(rep['citation_ids']['duplicates'], {})

    def test_deep_verify_flags_collisions(self):
        d = Document(); d.add_paragraph('Claim (Smith, 2007) and (Jones, 2010).')
        src = self.path('s6.docx'); d.save(src)
        out = self.path('o6.docx')
        mp = {'smith': {'itemKey': 'K1', 'uri': 'http://zotero.org/users/1/items/K1',
                        'csl': csl('A', 2007, 'Smith')},
              'jones': {'itemKey': 'K2', 'uri': 'http://zotero.org/users/1/items/K2',
                        'csl': csl('B', 2010, 'Jones')}}
        zfi.insert_zotero_fields(src, out, (0, 1), [], mp)
        # 人为制造撞号：两个不同引文共用同一 citationID
        from lxml import etree as ET
        z = zipfile.ZipFile(out); root = ET.fromstring(z.read('word/document.xml'))
        for els, _, _ in zfi.collect_fields(root):
            full = zfi.field_instr_text(els)
            if 'CSL_CITATION' in full:
                obj = zfi.field_payload(els)
                obj['citationID'] = 'cit0'
                da._rewrite_field(els, obj)
        xml = ET.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
        with zipfile.ZipFile(out) as zin, zipfile.ZipFile(out + '.t', 'w') as zo:
            for i in zin.infolist():
                zo.writestr(i, xml if i.filename == 'word/document.xml' else zin.read(i.filename))
        os.replace(out + '.t', out)
        ok, detail = zfi.verify(out, {'K1', 'K2'}, deep=True)
        self.assertFalse(ok)
        self.assertTrue(detail['citation_ids']['duplicates'])


if __name__ == '__main__':
    unittest.main()
