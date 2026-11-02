"""Regression cases found during the v3.1 review."""

import contextlib
import copy
import io
import json
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch
import zipfile

from docx import Document
from lxml import etree

import docx_audit as audit
import reference_ingest as ingest
import zotero_field_insert as insert


def mapping():
    return {
        "smith": {
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


class DocumentRegressions(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.source = str(Path(self.temp.name) / "source.docx")
        self.output = str(Path(self.temp.name) / "output.docx")
        self.mapping = mapping()

    def document(self, *paragraphs):
        doc = Document()
        for text in paragraphs:
            doc.add_paragraph(text)
        doc.save(self.source)

    def convert(self, end=1, **kwargs):
        return insert.insert_zotero_fields(
            self.source,
            self.output,
            (0, end),
            [],
            self.mapping,
            uri_prefix="http://zotero.org/users/123/items/",
            **kwargs,
        )

    def test_repeated_citation_missing_one_is_rejected(self):
        self.document("First (Smith, 2007).", "Second (Smith, 2007).")
        self.convert()
        ok, details = insert.verify(self.output, {"ABCD1234"}, self.source, self.mapping, deep=True)
        self.assertFalse(ok)
        self.assertEqual(details["missed_citations"], ["(Smith, 2007)"])

    def test_repeated_citations_all_converted_pass(self):
        self.document("First (Smith, 2007).", "Second (Smith, 2007).")
        self.convert(end=2)
        ok, details = insert.verify(self.output, {"ABCD1234"}, self.source, self.mapping, deep=True)
        self.assertTrue(ok, details)

    def test_non_text_run_contents_survive_conversion(self):
        self.document("Lead\tline\nClaim (Smith, 2007).\tEnd\n")
        self.convert()
        before, after = [audit.load_root(p) for p in (self.source, self.output)]
        for tag in ("tab", "br"):
            self.assertEqual(
                len(list(before.iter(insert.W + tag))), len(list(after.iter(insert.W + tag)))
            )
        self.assertIn(
            "Lead\tline\nClaim (Smith, 2007).\tEnd\n",
            [p.text for p in Document(self.output).paragraphs],
        )

    def test_inplace_conversion_preserves_zip_members(self):
        self.document("Claim (Smith, 2007).")
        with zipfile.ZipFile(self.source) as archive:
            other_parts = {
                n: archive.read(n) for n in archive.namelist() if n != "word/document.xml"
            }
        insert.insert_zotero_fields(self.source, self.source, (0, 1), [], self.mapping)
        with zipfile.ZipFile(self.source) as archive:
            self.assertEqual(other_parts, {n: archive.read(n) for n in other_parts})
        self.assertEqual(audit.scan_document(self.source, verbose=False)["fields"]["item"], 1)

    def test_table_conversion_is_idempotent(self):
        doc = Document()
        table = doc.add_table(rows=1, cols=1)
        table.cell(0, 0).text = "10.1000/example"
        doc.save(self.source)
        self.convert(end=0, include_tables=True)
        second = str(Path(self.temp.name) / "second.docx")
        count, _ = insert.insert_zotero_fields(
            self.output, second, (0, 0), [], self.mapping, include_tables=True
        )
        self.assertEqual(count, 0)
        self.assertEqual(
            insert.citation_id_stats(audit.load_root(self.output))["ids"],
            insert.citation_id_stats(audit.load_root(second))["ids"],
        )

    def test_audit_cli_signals_failure(self):
        self.document("Unconverted (doi:10.1000/example).")
        path = Path(self.temp.name) / "mapping.json"
        path.write_text(json.dumps(self.mapping), encoding="utf-8")
        with contextlib.redirect_stdout(io.StringIO()):
            code = audit.main(["docx_audit.py", "audit", self.source, self.source, str(path)])
        self.assertEqual(code, 1)

    def test_empty_renumber_selection_changes_nothing(self):
        self.document("Claim (Smith, 2007).")
        self.convert()
        before = Path(self.output).read_bytes()
        count = audit.renumber_citation_ids(self.output, from_docx=self.output)
        self.assertEqual(count, 0)
        self.assertEqual(before, Path(self.output).read_bytes())
        self.assertEqual(audit.renumber_citation_ids(self.output, new_display_texts=[]), 0)

    def test_renumber_preserves_other_ids_without_collisions(self):
        self.document("Claim (Smith, 2007).", "Another (doi:10.1000/example).")
        self.convert(end=2)
        audit.renumber_citation_ids(self.output, new_display_texts=["(Smith, 2007)"], prefix="sis")
        audit.renumber_citation_ids(
            self.output, new_display_texts=["(doi:10.1000/example)"], prefix="sis"
        )
        self.assertEqual(insert.citation_id_stats(audit.load_root(self.output))["duplicates"], {})

    def test_mapping_uri_is_used_without_global_user_id(self):
        self.document("Claim (Smith, 2007).")
        insert.insert_zotero_fields(self.source, self.output, (0, 1), [], self.mapping)
        fields = insert.collect_fields(audit.load_root(self.output))
        payloads = [insert.field_payload(f[0]) for f in fields]
        citation = next(p for p in payloads if "citationItems" in p)
        self.assertEqual(citation["citationItems"][0]["uris"], [self.mapping["smith"]["uri"]])

    def test_reversed_field_markers_fail_even_with_equal_counts(self):
        self.document("Plain text")
        root = audit.load_root(self.source)
        paragraph = root.find(insert.W + "body").find(insert.W + "p")
        for kind in ("end", "begin"):
            marker = etree.SubElement(
                etree.SubElement(paragraph, insert.W + "r"), insert.W + "fldChar"
            )
            marker.set(insert.W + "fldCharType", kind)
        audit.save_document(root, self.source, self.output)
        ok, details = insert.verify(self.output, set())
        self.assertFalse(ok)
        self.assertGreater(details["structure_errors"], 0)

    def test_failed_atomic_write_keeps_original_document(self):
        self.document("Do not lose this document")
        before = Path(self.source).read_bytes()
        root = audit.load_root(self.source)
        module = audit.save_document.__module__
        with patch(f"{module}.os.replace", side_effect=OSError("simulated disk error")):
            with self.assertRaises(OSError):
                audit.save_document(root, self.source, self.source)
        self.assertEqual(before, Path(self.source).read_bytes())
        self.assertEqual(list(Path(self.temp.name).glob(".zotero-*")), [])

    def test_bare_doi_group_is_detected_by_scan_and_deep_verify(self):
        self.document("Claim (10.1000/example).")
        self.assertEqual(audit.scan_document(self.source, verbose=False)["placeholder_groups"], 1)
        ok, details = insert.verify(self.source, set(), deep=True)
        self.assertFalse(ok)
        self.assertTrue(details["residual_placeholders"])

    def test_image_in_citation_paragraph_survives(self):
        doc = Document()
        paragraph = doc.add_paragraph("Before ")
        drawing = etree.SubElement(paragraph.add_run()._r, insert.W + "drawing")
        etree.SubElement(
            drawing,
            "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}inline",
        )
        paragraph.add_run(" (Smith, 2007).")
        doc.save(self.source)
        self.convert()
        self.assertEqual(len(list(audit.load_root(self.output).iter(insert.W + "drawing"))), 1)


class IngestRegressions(unittest.TestCase):
    def test_old_style_doi_is_complete_at_ingestion(self):
        doi = "10.1016/S0306-4530(98)00014-6"
        parsed = ingest.parse_numbered_reference(f"[1] Smith J. Title. Journal. 1998. doi:{doi}.")
        self.assertEqual(parsed["doi"].lower(), doi.lower())

    def test_failed_dedup_does_not_write_items(self):
        calls = []

        def service(name, args):
            calls.append(name)
            return {"error": "library unavailable"}

        with self.assertRaises(RuntimeError):
            ingest.import_to_zotero(
                {"smith": {"action": "import", "csl": mapping()["smith"]["csl"]}},
                "COLL",
                zotero=service,
            )
        self.assertNotIn("write_item", calls)

    def test_invalid_import_action_is_rejected_before_writes(self):
        calls = []
        with self.assertRaises(ValueError):
            ingest.import_to_zotero(
                {"smith": {"action": "improt", "csl": mapping()["smith"]["csl"]}},
                "COLL",
                zotero=lambda *args: calls.append(args),
                precheck=False,
            )
        self.assertEqual(calls, [])

    def test_failed_library_refresh_does_not_partially_mutate_mapping(self):
        items = mapping()
        items["other"] = {**copy.deepcopy(items["smith"]), "itemKey": "OTHER123"}
        before = copy.deepcopy(items)

        def service(name, args):
            if args["itemKey"] == "OTHER123":
                raise ConnectionError("library disconnected")
            return {"title": "Updated title", "date": "2008"}

        with self.assertRaises(ConnectionError):
            ingest.apply_library_data(items, zotero=service)
        self.assertEqual(items, before)


if __name__ == "__main__":
    unittest.main()
