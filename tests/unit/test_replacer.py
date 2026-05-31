"""Unit tests for replacer module."""

import zipfile
from pathlib import Path

import pytest
from lxml import etree

from mathtypejx.replacer import (
    replace_formulas,
    _repack_docx,
    _update_content_types,
    _replace_in_part,
)
from mathtypejx.models import FormulaInfo, FormulaStatus, RiskLevel

NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS_M = "http://schemas.openxmlformats.org/officeDocument/2006/math"


def _make_omml_formula(fid="F0001", rid="rId6", omml=None):
    """Helper to create a converted FormulaInfo."""
    if omml is None:
        omml = f'<m:oMath xmlns:m="{NS_M}"><m:r><m:t>x</m:t></m:r></m:oMath>'
    return FormulaInfo(
        formula_id=fid,
        ole_name="oleObject1.bin",
        part_name="word/document.xml",
        rels_path="word/_rels/document.xml.rels",
        relationship_id=rid,
        prog_id="Equation.DSMT4",
        status=FormulaStatus.CONVERTED,
        risk_level=RiskLevel.AUTO_REPLACE,
        omml=omml,
    )


class TestReplacer:
    """Tests for OLE→OMML replacement in docx XML."""

    def test_replace_single_formula(self, temp_dir):
        """Replace one OLE formula with OMML in document.xml."""
        # Setup: create a workdir with document.xml containing a MathType OLE
        workdir = temp_dir / "work"
        workdir.mkdir()
        (workdir / "word").mkdir()
        (workdir / "word" / "embeddings").mkdir(parents=True)

        doc_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="{NS_W}"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            xmlns:o="urn:schemas-microsoft-com:office:office"
            xmlns:v="urn:schemas-microsoft-com:vml">
  <w:body>
    <w:p>
      <w:r><w:t>Text </w:t></w:r>
      <w:r>
        <w:object>
          <v:shape o:ole=""/>
          <o:OLEObject r:id="rId6" ProgID="Equation.DSMT4"/>
        </w:object>
      </w:r>
      <w:r><w:t> more.</w:t></w:r>
    </w:p>
  </w:body>
</w:document>"""

        (workdir / "word" / "document.xml").write_text(doc_xml, encoding="utf-8")

        # Create Content_Types.xml
        ct = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml"
   ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"""
        (workdir / "[Content_Types].xml").write_text(ct)

        # Create _rels
        (workdir / "_rels").mkdir()
        (workdir / "_rels" / ".rels").write_text(
            """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""
        )

        formula = _make_omml_formula()
        output = temp_dir / "output.docx"

        replace_formulas(workdir, [formula], str(output))

        # Verify output
        assert output.exists()
        assert formula.status == FormulaStatus.REPLACED

        # Check the output docx contains OMML
        with zipfile.ZipFile(str(output), "r") as z:
            doc_content = z.read("word/document.xml").decode("utf-8")
            assert "oMath" in doc_content
            # Original OLE should be gone
            assert "OLEObject" not in doc_content
            # Text around formula should be preserved
            assert "Text" in doc_content
            assert "more" in doc_content

    def test_failed_formula_not_replaced(self, temp_dir):
        """Failed formulas should not be replaced in the document."""
        workdir = temp_dir / "work"
        workdir.mkdir()
        (workdir / "word").mkdir()

        doc_xml = f"""<?xml version="1.0"?>
<w:document xmlns:w="{NS_W}"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            xmlns:o="urn:schemas-microsoft-com:office:office"
            xmlns:v="urn:schemas-microsoft-com:vml">
  <w:body><w:p><w:r><w:object>
    <o:OLEObject r:id="rId6" ProgID="Equation.DSMT4"/>
  </w:object></w:r></w:p></w:body>
</w:document>"""
        (workdir / "word" / "document.xml").write_text(doc_xml, encoding="utf-8")

        ct = """<?xml version="1.0"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
</Types>"""
        (workdir / "[Content_Types].xml").write_text(ct)
        (workdir / "_rels").mkdir()
        (workdir / "_rels" / ".rels").write_text(
            """<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""
        )

        # Formula with FAILED status
        formula = FormulaInfo(
            formula_id="F0001",
            ole_name="oleObject1.bin",
            part_name="word/document.xml",
            rels_path="word/_rels/document.xml.rels",
            relationship_id="rId6",
            prog_id="Equation.DSMT4",
            status=FormulaStatus.FAILED,
            error_message="Conversion failed",
        )

        output = temp_dir / "output_failed.docx"
        replace_formulas(workdir, [formula], str(output))

        # OLE should still be present
        with zipfile.ZipFile(str(output), "r") as z:
            doc_content = z.read("word/document.xml").decode("utf-8")
            assert "OLEObject" in doc_content

    def test_replace_preserves_text(self, temp_dir):
        """Text content around formulas should survive replacement."""
        workdir = temp_dir / "work"
        workdir.mkdir()
        (workdir / "word").mkdir()
        (workdir / "word" / "embeddings").mkdir(parents=True)

        doc_xml = f"""<?xml version="1.0"?>
<w:document xmlns:w="{NS_W}"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            xmlns:o="urn:schemas-microsoft-com:office:office"
            xmlns:v="urn:schemas-microsoft-com:vml">
  <w:body>
    <w:p>
      <w:r><w:t>Before</w:t></w:r>
      <w:r><w:object><o:OLEObject r:id="rId6" ProgID="Equation.DSMT4"/></w:object></w:r>
      <w:r><w:t>After</w:t></w:r>
    </w:p>
  </w:body>
</w:document>"""
        (workdir / "word" / "document.xml").write_text(doc_xml, encoding="utf-8")

        ct = """<?xml version="1.0"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
</Types>"""
        (workdir / "[Content_Types].xml").write_text(ct)
        (workdir / "_rels").mkdir()
        (workdir / "_rels" / ".rels").write_text(
            """<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""
        )

        formula = _make_omml_formula()
        output = temp_dir / "output_text.docx"
        replace_formulas(workdir, [formula], str(output))

        with zipfile.ZipFile(str(output), "r") as z:
            doc_content = z.read("word/document.xml").decode("utf-8")
            assert "Before" in doc_content
            assert "After" in doc_content
