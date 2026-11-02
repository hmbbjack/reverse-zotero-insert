"""DOI parsing shared by ingestion, conversion, and auditing."""

import re

DOI_TOKEN = r"10\.\d{4,9}/(?:[^\s,;；()（）\"']+|\([^\s()]*\))+"
DOI_RE = re.compile("(" + DOI_TOKEN + ")", re.IGNORECASE)
DOI_PREFIX = r"(?:https?://(?:dx\.)?doi\.org/|doi\s*[:：]\s*)?"
DOI_PART = re.compile(DOI_PREFIX + "(" + DOI_TOKEN + ")", re.IGNORECASE)
PAREN_GROUP = re.compile(r"[（(]((?:[^（）()]|\([^（）()]*\))*)[）)]")


def normalize_doi(value):
    """Remove labels and surrounding punctuation, preserving balanced DOI parentheses."""
    if not value:
        return ""
    text = str(value).strip().lstrip("(（").strip()
    text = re.sub("^" + DOI_PREFIX, "", text, flags=re.IGNORECASE)
    text = text.rstrip(".,;；。 ")
    for opening, closing in (("(", ")"), ("（", "）")):
        while text.endswith(closing) and text.count(closing) > text.count(opening):
            text = text[:-1].rstrip(".,;；。 ")
    return text.lower()


def extract_doi_groups(text):
    """Return (start, end, display, dois) for complete, pure DOI groups."""
    groups = []
    for match in PAREN_GROUP.finditer(text):
        parts = [part.strip() for part in re.split(r"[;；]", match.group(1))]
        if parts and all(DOI_PART.fullmatch(part) for part in parts):
            groups.append(
                (match.start(), match.end(), match.group(), [normalize_doi(part) for part in parts])
            )
    return groups
