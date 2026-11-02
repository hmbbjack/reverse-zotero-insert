# User guide

[Project home](../README.en.md) · [简体中文](USAGE.md) · [Development](DEVELOPMENT.en.md)

## Configuration

Set environment variables before importing the package:

| Variable | Purpose |
| :--- | :--- |
| `ZOTERO_USER_ID` | Your actual Zotero user ID; otherwise inferred from `~/Zotero/zotero.sqlite` when available |
| `ZOTERO_URI_PREFIX` | Optional item URI prefix, including the trailing `/items/` |
| `ZOTERO_MCP_URL` | Streamable HTTP endpoint; defaults to `http://127.0.0.1:23120/mcp` |
| `CROSSREF_MAILTO` | Optional contact email for the CrossRef polite pool |
| `NCBI_API_KEY` | Optional NCBI API key |

Source users can copy `config.example.env` to `config.local.env` next to `config.py`.
Replace the example placeholders with real values. For an installed package,
environment variables are recommended. Keep personal settings out of version control.

Offline document operations do not need a running Zotero instance. Library operations
require Zotero and the [cookjohn/zotero-mcp plugin](https://github.com/cookjohn/zotero-mcp).
The MCP endpoint on port 23120 is distinct from Zotero's API on port 23119.

```python
from reverse_zotero_insert import reference_ingest as ingest

status = ingest.check_zotero_mcp()
if not status["ok"]:
    raise RuntimeError(status["hint"])
```

## Prepare an item mapping

Each mapping entry links a document reference to an existing Zotero item. It contains
an `itemKey`, a full `uri`, CSL metadata in `csl`, and, when available, a `doi`.
Use real library records; example or placeholder identities do not establish a valid link.

For numbered reference lists:

1. Extract references with `extract_reference_paragraphs()`.
2. Resolve metadata with `resolve_metadata()` and inspect confidence scores.
3. Run `dedup_check()` against the library.
4. Review the metadata, destination collection, and import/reuse/skip decisions.
5. Pass the approved plan to `import_to_zotero()`.

```python
from reverse_zotero_insert import reference_ingest as ingest

# approved_plan and collection_key come from the reviewed import decisions.
imported, warnings = ingest.import_to_zotero(approved_plan, collection_key)
if warnings:
    raise RuntimeError(warnings)
mapping = ingest.build_item_mapping(imported)
mapping, doi_mismatches = ingest.apply_library_data(mapping)
```

Inspect `doi_mismatches` before conversion. Failed MCP/deduplication requests raise
exceptions; do not interpret them as an empty library. A failed library refresh leaves
the original mapping unchanged. Remote writes cannot be rolled back as a batch, so
inspect already-created items before retrying an interrupted import.

## Convert and verify

Scan the source first and keep a backup:

```bash
reverse-zotero-insert scan manuscript.docx --json
```

The following example assumes body paragraphs 0–164 and an already prepared
`item_mapping.json`. Adjust these ranges to the document. The end of the body range
is exclusive; `ref_para_indices` identifies bibliography paragraphs to wrap in a field.

```python
import json
from reverse_zotero_insert import zotero_field_insert as fields
from reverse_zotero_insert import docx_audit as audit

with open("item_mapping.json", encoding="utf-8") as source:
    mapping = json.load(source)

fields.insert_zotero_fields(
    "manuscript.docx",
    "manuscript_zotero.docx",
    body_para_range=(0, 165),
    ref_para_indices=[],
    item_mapping=mapping,
    style_id="http://www.zotero.org/styles/ieee",
)

ok, details = fields.verify(
    "manuscript_zotero.docx",
    {item["itemKey"] for item in mapping.values()},
    "manuscript.docx",
    mapping,
    deep=True,
)
if not ok:
    raise RuntimeError(details)
report = audit.audit_output("manuscript.docx", "manuscript_zotero.docx", mapping)
if report["problems"] or report["structure_errors"] or report["nested_citations"]:
    raise RuntimeError(report)
```

Prefer a new output file. Atomic writing also protects against truncation when the
same source/output path is accidentally supplied.

For tables, `include_tables=True` handles table paragraphs; use a body range of `(0, 0)`
to process only tables. `insert_table_citations()` targets the first table, using a
DOI column to locate items and a selected reference column for the citation fields.
Set `refnum_column=None` to replace the DOI cell itself.

## Understand the checks

- `verify()` checks field structure, citation JSON, item keys, and required CSL fields.
  Nested Zotero citations fail verification; valid nesting of ordinary Word fields is preserved.
- Providing the source and mapping enables occurrence-by-occurrence omission checks.
- `deep=True` adds duplicate citation IDs and residual DOI placeholders.
- `audit_output()` reconciles pure DOI parenthetical groups by display text, DOI, and item key.
- `verify_item_keys()` separately checks existence in the actual library. Local verification
  against supplied keys does not prove that those items still exist remotely.

```bash
reverse-zotero-insert audit manuscript.docx manuscript_zotero.docx item_mapping.json --json
reverse-zotero-insert renumber manuscript_zotero.docx --from manuscript.docx --dry-run
```

Audit success requires an empty `problems` list, `structure_errors == 0`, and `nested_citations == 0`.
Shared parsing protects existing fields across paragraphs and around bookmarks and proofing markers,
so repeated conversion leaves their cached citation text intact.

Audit exit codes: `0` no discrepancies, `1` reconciliation or field structure errors, `2` argument/file errors.
After XML checks pass, open the result in the target Word/Zotero environment and verify refresh behavior.
