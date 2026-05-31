"""Shared test fixtures for mathtypejx."""

import os
import struct
import tempfile
import zipfile
from pathlib import Path

import pytest

# Path to test fixture directory
FIXTURES_DIR = Path(__file__).resolve().parent / "fixtures"


def _load_ole_binary(index: int = 1) -> bytes:
    """Load a real OLE binary from the test fixtures."""
    path = FIXTURES_DIR / f"oleObject{index}.bin"
    if path.exists():
        return path.read_bytes()
    raise FileNotFoundError(f"Fixture not found: {path}")


# ── Fixture: temp working directory ──────────────────────────────
@pytest.fixture
def temp_dir():
    with tempfile.TemporaryDirectory(prefix="mathtypejx_test_") as d:
        yield Path(d)


# ── Fixture: real OLE binary from MathType formula ───────────────
@pytest.fixture
def mock_ole_binary():
    """Return a real MathType OLE binary from test fixtures."""
    return _load_ole_binary(1)


# ── Fixture: real OLE binary #2 (different formula) ─────────────
@pytest.fixture
def mock_ole_binary_2():
    return _load_ole_binary(2)


# ── Fixture: real MathType docx ──────────────────────────────────
@pytest.fixture
def sample_docx_with_mathtype():
    """Real docx with MathType formulas from test fixtures."""
    path = FIXTURES_DIR / "sample_mathtype.docx"
    if not path.exists():
        pytest.skip("Real MathType docx fixture not available")
    return path


# ── Fixture: docx with multiple MathType formulas ─────────────────
@pytest.fixture
def sample_docx_multi_formula():
    """Alias for the real docx (which has 20 formulas)."""
    return sample_docx_with_mathtype()


# ── Fixture: docx with NO MathType formulas ──────────────────────
@pytest.fixture
def sample_docx_no_formulas(temp_dir):
    """Create a .docx without any formulas."""
    docx_path = temp_dir / "no_formula.docx"

    content_types = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml"
   ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"""

    root_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

    document_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body><w:p><w:r><w:t>Plain text, no formulas.</w:t></w:r></w:p></w:body>
</w:document>"""

    with zipfile.ZipFile(str(docx_path), "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("_rels/.rels", root_rels)
        z.writestr("word/document.xml", document_xml)

    return docx_path
