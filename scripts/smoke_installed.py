"""Run with the installed Python interpreter, from outside the checkout."""

import json
from pathlib import Path
import subprocess
import sys
import tempfile

from docx import Document
from reverse_zotero_insert import __version__
from reverse_zotero_insert import docx_audit, reference_ingest, zotero_field_insert, zotero_mcp

assert __version__
assert not Path(docx_audit.__file__).resolve().is_relative_to(Path(__file__).resolve().parents[1])
assert reference_ingest.normalize_doi("DOI:10.1000/Example.") == "10.1000/example"
assert callable(zotero_mcp.MCPClient().call_tool)
with tempfile.TemporaryDirectory() as directory:
    root = Path(directory)
    src, out = root / "输入.docx", root / "输出.docx"
    doc = Document()
    doc.add_paragraph("First\tline\nClaim (doi:10.1000/example).")
    doc.add_paragraph("Second (doi:10.1000/example).")
    doc.add_paragraph("[1] Smith J. Example. 2007.")
    doc.save(src)
    mapping = {
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
    scanned = subprocess.run(
        [sys.executable, "-m", "reverse_zotero_insert", "scan", str(src), "--json"],
        check=True,
        capture_output=True,
        text=True,
    )
    scan_report = json.loads(scanned.stdout)
    assert scan_report["placeholder_groups"] == 2
    assert scan_report["file"] == src.name
    assert zotero_field_insert.insert_zotero_fields(src, out, (0, 2), [2], mapping) == (2, 1)
    ok, report = zotero_field_insert.verify(out, {"ABCD1234"}, src, mapping, deep=True)
    assert ok, report
    assert "\t" in Document(out).paragraphs[1].text
    path = root / "mapping.json"
    path.write_text(json.dumps(mapping), encoding="utf-8")
    command = [
        sys.executable,
        "-m",
        "reverse_zotero_insert",
        "audit",
        str(src),
        str(out),
        str(path),
        "--json",
    ]
    good = subprocess.run(command, check=True, capture_output=True, text=True)
    assert json.loads(good.stdout)["matched"] == 2
    command[5] = str(src)
    bad = subprocess.run(command, capture_output=True, text=True)
    assert bad.returncode == 1, (bad.returncode, bad.stdout, bad.stderr)
    bad_report = json.loads(bad.stdout)
    assert bad_report["matched"] == 0
    assert len(bad_report["problems"]) == 2
    assert all(problem["why"] == "未找到对应域" for problem in bad_report["problems"])
print(
    f"Installed package {__version__}: imports, scan, conversion, deep verify, CLI exit codes PASS"
)
