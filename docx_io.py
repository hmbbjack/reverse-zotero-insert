"""Read DOCX XML and replace it atomically, preserving other archive members."""

import os
from pathlib import Path
import tempfile
import zipfile

from lxml import etree

DOCUMENT_XML = "word/document.xml"


def load_root(path):
    with zipfile.ZipFile(path) as archive:
        return etree.fromstring(archive.read(DOCUMENT_XML))


def save_document(root, source, destination):
    """Write through a sibling temporary file, including when source == destination."""
    destination = Path(destination)
    xml = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    fd, temporary = tempfile.mkstemp(prefix=".zotero-", suffix=".docx", dir=destination.parent)
    os.close(fd)
    try:
        with zipfile.ZipFile(source) as original, zipfile.ZipFile(temporary, "w") as output:
            output.comment = original.comment
            for member in original.infolist():
                output.writestr(
                    member,
                    xml if member.filename == DOCUMENT_XML else original.read(member.filename),
                )
        os.replace(temporary, destination)
    finally:
        if os.path.exists(temporary):
            os.unlink(temporary)
