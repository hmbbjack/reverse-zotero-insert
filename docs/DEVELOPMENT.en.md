# Development and publishing

[Project home](../README.en.md) · [简体中文](DEVELOPMENT.md)

## Development

Use Python 3.10+ and a virtual environment. From the repository root:

```bash
python -m pip install ".[dev]"
python -m unittest discover -s tests -v
ruff check .
ruff format --check .
```

The source modules retain their original import paths. Wheels install them under
`reverse_zotero_insert`, avoiding global names such as `config`.
Shared DOI rules live in `doi_utils.py`, atomic DOCX output in `docx_io.py`, and
MCP session handling in `zotero_mcp.MCPClient`.

## Build packages

```bash
python -m build
python scripts/build_zip.py
```

`dist/` contains a wheel, source archive, and portable ZIP. The ZIP includes project
code, tests, documentation, artwork, and the example configuration. Personal settings,
manuscripts, library data, and dependency caches must not be added to release artifacts.

Test installation outside the checkout to avoid importing source files by accident:

```bash
python -m venv /tmp/reverse-zotero-package-test
/tmp/reverse-zotero-package-test/bin/python -m pip install dist/*.whl
cd /tmp
/tmp/reverse-zotero-package-test/bin/python -m reverse_zotero_insert --help
```

Windows virtual environments use `Scripts/python.exe`.
CI is configured for Linux/Windows and Python 3.10, 3.12, and 3.13. Passing local
checks does not mean that these remote jobs have already run.

## Publish a GitHub release

1. Update the versions in `pyproject.toml`, `__init__.py`, and the MCP client, and update release notes.
2. Merge tested changes into the default branch, then create and push a matching tag such as `v3.1.0`.
3. The release workflow validates the version, runs checks, builds packages, tests the installed wheel,
   and generates SHA256 checksums.
4. It creates a **Draft Release** with attachments. Review them in GitHub Releases and select
   **Publish release** to make the version available to visitors.

The workflow uses GitHub's automatically provided token. The README's ZIP button
downloads the current `main` branch; release attachments are tied to a version tag.
PyPI publishing is not configured. Install from a GitHub URL, local source, or wheel.

## Known limits

- Offline tests and simulated MCP responses do not replace real Zotero/Word Refresh testing.
- Body conversion targets `word/document.xml`. Headers, footers, footnotes, endnotes,
  and tracked-change content need separate treatment.
- Hyperlinks, bookmarks, or embedded objects may prevent matching across structural boundaries;
  inspect verification results rather than flattening the document.
- `audit_output()` focuses on pure DOI groups. Use it alongside deep verification and library key checks.
- Library writes are not a transaction; inspect existing items before retrying a partial import.
- Collection indexing currently requests up to 500 records at once. For larger collections,
  provide a complete `item_keys` list rather than assuming a full-library scan.
