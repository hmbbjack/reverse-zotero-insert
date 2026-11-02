"""Build a portable source ZIP from an explicit list; never include local settings."""

import argparse
import ast
from pathlib import Path
import zipfile

ROOT = Path(__file__).resolve().parents[1]
RELEASE_FILES = (
    "__init__.py",
    "__main__.py",
    "config.py",
    "docx_audit.py",
    "docx_io.py",
    "doi_utils.py",
    "reference_ingest.py",
    "zotero_field_insert.py",
    "zotero_mcp.py",
    "README.md",
    "README.en.md",
    "LICENSE",
    "SKILL.md",
    "config.example.env",
    "requirements.txt",
    "pyproject.toml",
    "MANIFEST.in",
    ".gitignore",
)


def version():
    for node in ast.parse((ROOT / "__init__.py").read_text(encoding="utf-8")).body:
        if isinstance(node, ast.Assign) and any(
            isinstance(target, ast.Name) and target.id == "__version__" for target in node.targets
        ):
            return ast.literal_eval(node.value)
    raise ValueError("Missing package version")


def build_zip(output):
    output.mkdir(parents=True, exist_ok=True)
    folder = f"reverse-zotero-insert-{version()}"
    destination = output / f"{folder}.zip"
    files = [ROOT / name for name in RELEASE_FILES]
    for name, pattern in (
        ("tests", "*.py"),
        ("scripts", "*.py"),
        ("docs", "*.md"),
        ("assets", "*.png"),
    ):
        files.extend(sorted((ROOT / name).glob(pattern)))
    with zipfile.ZipFile(destination, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for path in files:
            archive.write(path, f"{folder}/{path.relative_to(ROOT).as_posix()}")
    return destination


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", type=Path, default=ROOT / "dist")
    print(build_zip(parser.parse_args().output))
