"""Word field boundaries must agree across detection, conversion, and auditing."""

import copy
import contextlib
import io
import json
from pathlib import Path
import tempfile
import unittest

from docx import Document
from lxml import etree

import docx_audit as audit
import zotero_field_insert as insert

W = insert.W


def text_run(text):
    run = etree.Element(W + "r")
    etree.SubElement(run, W + "t").text = text
    return run


def marker(kind):
    run = etree.Element(W + "r")
    etree.SubElement(run, W + "fldChar").set(W + "fldCharType", kind)
    return run


class WordFieldTests(unittest.TestCase):
    def setUp(self):
        temporary = tempfile.TemporaryDirectory()
        self.addCleanup(temporary.cleanup)
        self.directory = Path(temporary.name)
        self.source = self.directory / "source.docx"
        self.existing = self.directory / "existing.docx"
        self.output = self.directory / "output.docx"
        self.mapping = {
            "example": {
                "itemKey": "ABCD1234",
                "uri": "http://zotero.org/users/123/items/ABCD1234",
                "doi": "10.1000/example",
                "csl": {
                    "type": "article-journal",
                    "title": "Example",
                    "author": [{"family": "Smith", "given": "J"}],
                    "issued": {"date-parts": [[2007]]},
                    "DOI": "10.1000/example",
                },
            }
        }
        doc = Document()
        doc.add_paragraph("(doi:10.1000/example)")
        doc.save(self.source)
        insert.insert_zotero_fields(self.source, self.existing, (0, 1), [], self.mapping)
        self.root = audit.load_root(self.existing)
        self.body = self.root.find(W + "body")
        self.paragraph = self.body.findall(W + "p")[1]

    def save(self):
        audit.save_document(self.root, self.existing, self.existing)

    def verify(self, path=None, deep=True):
        return insert.verify(path or self.output, {"ABCD1234"}, deep=deep)

    def citation_count(self, root):
        return sum("CSL_CITATION" in (e.text or "") for e in root.iter(W + "instrText"))

    def add_markers(self, kind):
        if kind == "bookmark":
            first, last = etree.Element(W + "bookmarkStart"), etree.Element(W + "bookmarkEnd")
            first.set(W + "id", "42")
            first.set(W + "name", "citation_bookmark")
            last.set(W + "id", "42")
        elif kind == "proofErr":
            first, last = etree.Element(W + "proofErr"), etree.Element(W + "proofErr")
            first.set(W + "type", "spellStart")
            last.set(W + "type", "spellEnd")
        else:
            first = etree.Element(W + "commentRangeStart")
            last = etree.Element(W + "commentRangeEnd")
            first.set(W + "id", "42")
            last.set(W + "id", "42")
        self.paragraph.insert(len(self.paragraph) - 1, last)
        self.paragraph.insert(1, first)

    def nest_citation(self):
        inner = copy.deepcopy(list(self.paragraph))
        for run in inner:
            for instruction in run.iter(W + "instrText"):
                payload = insert.field_payload([instruction])
                payload["citationID"] = "different-id"
                audit._rewrite_field([instruction], payload)
        display = next(run for run in self.paragraph if run.find(W + "t") is not None)
        index = self.paragraph.index(display)
        self.paragraph.remove(display)
        for run in reversed(inner):
            self.paragraph.insert(index, run)
        self.save()

    def test_markers_inside_existing_citation_are_preserved(self):
        original = etree.tostring(self.paragraph)
        for kind in ("bookmark", "proofErr", "comment"):
            with self.subTest(kind=kind):
                replacement = etree.fromstring(original)
                self.body.replace(self.paragraph, replacement)
                self.paragraph = replacement
                self.add_markers(kind)
                before = etree.tostring(self.paragraph)
                self.save()
                count, _ = insert.insert_zotero_fields(
                    self.existing, self.output, (1, 2), [], self.mapping
                )
                self.assertEqual(count, 0)
                root = audit.load_root(self.output)
                self.assertEqual(self.citation_count(root), 1)
                self.assertEqual(etree.tostring(root.find(W + "body").findall(W + "p")[1]), before)
                ok, details = self.verify()
                self.assertTrue(ok, details)

    def test_marked_existing_field_does_not_block_new_citation(self):
        self.add_markers("bookmark")
        self.paragraph.append(text_run(" and (Smith, 2007)"))
        self.save()
        count, _ = insert.insert_zotero_fields(self.existing, self.output, (1, 2), [], self.mapping)
        self.assertEqual(count, 1)
        root = audit.load_root(self.output)
        self.assertEqual(self.citation_count(root), 2)
        self.assertEqual(insert.citation_id_stats(root)["ids"], ["cit0", "cit1"])
        ok, details = self.verify()
        self.assertTrue(ok, details)

    def test_existing_table_field_with_markers_is_not_converted_again(self):
        self.add_markers("proofErr")
        table = etree.SubElement(self.body, W + "tbl")
        cell = etree.SubElement(etree.SubElement(table, W + "tr"), W + "tc")
        cell.append(self.paragraph)
        self.save()
        count, _ = insert.insert_zotero_fields(
            self.existing, self.output, (0, 0), [], self.mapping, include_tables=True
        )
        self.assertEqual(count, 0)
        self.assertEqual(self.citation_count(audit.load_root(self.output)), 1)

    def test_nested_citation_fails_verification_even_with_distinct_ids(self):
        self.nest_citation()
        for deep in (False, True):
            with self.subTest(deep=deep):
                ok, details = self.verify(self.existing, deep=deep)
                self.assertFalse(ok)
                self.assertEqual(details["cites"], 2)
                self.assertEqual(details["nested_citations"], 1)

    def test_audit_cli_rejects_nested_citation(self):
        self.nest_citation()
        mapping = self.directory / "mapping.json"
        mapping.write_text(json.dumps(self.mapping), encoding="utf-8")
        output = io.StringIO()
        with contextlib.redirect_stdout(output):
            result = audit.main(
                [
                    "docx_audit.py",
                    "audit",
                    str(self.source),
                    str(self.existing),
                    str(mapping),
                    "--json",
                ]
            )
        self.assertEqual(result, 1)
        report = json.loads(output.getvalue())
        self.assertEqual(report["nested_citations"], 1)
        self.assertGreaterEqual(report["matched"], 0)

    def test_citation_inside_non_zotero_field_is_collected_and_valid(self):
        outer = insert._make_field_runs('HYPERLINK "https://example.org"', "", None)
        for node in reversed(outer[:3]):
            self.paragraph.insert(0, node)
        self.paragraph.append(outer[-1])
        self.save()
        fields = insert.collect_fields(self.root)
        self.assertEqual(len(fields), 3)  # preferences, hyperlink, citation
        ok, details = self.verify(self.existing)
        self.assertTrue(ok, details)
        self.assertEqual(details["cites"], 1)

    def test_nested_regular_field_does_not_hide_existing_citation(self):
        display = next(run for run in self.paragraph if run.find(W + "t") is not None)
        index = self.paragraph.index(display)
        self.paragraph.remove(display)
        for node in reversed(insert._make_field_runs("DATE", "2007", None)):
            self.paragraph.insert(index, node)
        detected = insert._detect_existing(self.body)
        self.assertEqual(detected["cited_paras"], {1})
        self.assertEqual(insert.collect_fields(self.root)[1][1], "2007")

    def test_cross_paragraph_field_preserves_internal_text_and_converts_tail(self):
        self.body.remove(self.paragraph)
        paragraphs = [etree.SubElement(self.body, W + "p") for _ in range(3)]
        runs = insert._make_field_runs('ZOTERO_BIBL {"uncited":[]}', "", None)
        paragraphs[0].extend(runs[:3])
        for p in paragraphs:
            p.append(text_run("(doi:10.1000/example)"))
        paragraphs[-1].append(runs[-1])
        paragraphs[-1].append(text_run(" Outside (Smith, 2007)."))
        self.save()
        count, _ = insert.insert_zotero_fields(self.existing, self.output, (1, 4), [], self.mapping)
        self.assertEqual(count, 1)
        root = audit.load_root(self.output)
        self.assertEqual(self.citation_count(root), 1)
        ok, details = self.verify()
        self.assertTrue(ok, details)
        self.assertEqual(details["bibs"], 1)

    def test_field_and_new_citation_sharing_one_run(self):
        combined = etree.Element(W + "r")
        for run in list(self.paragraph):
            combined.extend(list(run))
            self.paragraph.remove(run)
        etree.SubElement(combined, W + "t").text = " and (Smith, 2007)"
        self.paragraph.append(combined)
        self.save()
        count, _ = insert.insert_zotero_fields(self.existing, self.output, (1, 2), [], self.mapping)
        self.assertEqual(count, 1)
        ok, details = self.verify()
        self.assertTrue(ok, details)
        self.assertEqual(details["cites"], 2)

    def test_simple_field_is_preserved_and_its_instruction_can_be_repaired(self):
        instruction = next(self.paragraph.iter(W + "instrText")).text
        display = next(run for run in self.paragraph if run.find(W + "t") is not None)
        simple = etree.Element(W + "fldSimple")
        simple.set(W + "instr", instruction)
        simple.append(copy.deepcopy(display))
        self.paragraph[:] = [simple]
        self.save()
        count, _ = insert.insert_zotero_fields(self.existing, self.output, (1, 2), [], self.mapping)
        self.assertEqual(count, 0)
        self.assertEqual(audit.renumber_citation_ids(self.output, prefix="simple"), 1)
        self.assertEqual(audit.relink_item_keys(self.output, {"ABCD1234": "NEWKEY12"}), 1)
        ok, details = insert.verify(self.output, {"NEWKEY12"}, deep=True)
        self.assertTrue(ok, details)
        root = audit.load_root(self.output)
        self.assertEqual(insert.citation_id_stats(root)["ids"], ["simple0"])
        repaired = root.find(".//" + W + "fldSimple")
        self.assertIn("NEWKEY12", repaired.get(W + "instr"))
        self.assertEqual("".join(repaired.itertext()), "(doi:10.1000/example)")

    def test_text_box_fields_are_separate_and_not_counted_twice(self):
        outer = etree.Element(W + "p")
        runs = insert._make_field_runs('ZOTERO_BIBL {"uncited":[]}', "", None)
        outer.extend(runs[:3])
        outer.append(text_run("Bibliography before "))
        drawing = etree.SubElement(etree.SubElement(outer, W + "r"), W + "pict")
        textbox = etree.SubElement(drawing, W + "txbxContent")
        textbox.append(self.paragraph)
        outer.append(text_run("and after"))
        outer.append(runs[-1])
        self.body.append(outer)
        fields = insert.collect_fields(self.root)
        self.assertEqual(len(fields), 3)
        self.assertEqual(fields[1][1], "Bibliography before and after")
        self.assertEqual(fields[2][1], "(doi:10.1000/example)")
        self.assertEqual(insert.citation_id_stats(self.root)["ids"], ["cit0"])
        self.save()
        ok, details = self.verify(self.existing)
        self.assertTrue(ok, details)

    def test_nested_instruction_does_not_change_parent_instruction_or_result(self):
        outer = etree.Element(W + "p")
        prefix = insert._make_field_runs("IF ", "", None)
        outer.extend(prefix[:2])
        outer.extend(insert._make_field_runs("DOCVARIABLE example", "1", None))
        suffix = etree.Element(W + "r")
        etree.SubElement(suffix, W + "instrText").text = ' = 1 "yes" "no"'
        outer.append(suffix)
        outer.append(marker("separate"))
        outer.append(text_run("yes"))
        outer.append(marker("end"))
        fields = insert.collect_fields(outer)
        self.assertEqual(len(fields), 2)
        instruction = insert.field_instr_text(fields[0][0])
        self.assertIn('= 1 "yes" "no"', instruction)
        self.assertNotIn("DOCVARIABLE", instruction)
        self.assertEqual([display for _, display, _ in fields], ["yes", "1"])

    def test_invalid_field_markers_fail_audit_without_new_citations(self):
        for kind in ("end", "begin", "separate"):
            with self.subTest(kind=kind):
                self.paragraph.append(marker(kind))
                self.save()
                report = audit.audit_output(
                    self.existing, self.existing, self.mapping, verbose=False
                )
                self.assertEqual(report["total"], 0)
                self.assertGreater(report["structure_errors"], 0)
                ok, details = self.verify(self.existing)
                self.assertFalse(ok, details)
                self.paragraph.remove(self.paragraph[-1])


if __name__ == "__main__":
    unittest.main()
