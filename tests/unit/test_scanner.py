"""Unit tests for scanner module."""

from mathtypejx.scanner import scan_docx, read_ole_binary
from mathtypejx.models import FormulaStatus, FormulaType


class TestScanner:
    """Tests for the docx formula scanner."""

    def test_find_mathtype_formulas(self, sample_docx_with_mathtype):
        """Real docx has 20 MathType OLE formulas."""
        formulas, workdir = scan_docx(str(sample_docx_with_mathtype))
        mathtype = [f for f in formulas if f.formula_type == FormulaType.MATHTYPE_OLE]
        assert len(mathtype) == 20
        assert all(f.prog_id == "Equation.DSMT4" for f in mathtype)

    def test_no_formulas(self, sample_docx_no_formulas):
        """Scan a docx with no formulas."""
        formulas, workdir = scan_docx(str(sample_docx_no_formulas))
        mathtype = [f for f in formulas if f.formula_type == FormulaType.MATHTYPE_OLE]
        assert len(mathtype) == 0

    def test_read_ole_binary(self, sample_docx_with_mathtype):
        """Verify OLE binary is readable from word/embeddings/."""
        formulas, workdir = scan_docx(str(sample_docx_with_mathtype))
        f = formulas[0]
        ole = read_ole_binary(workdir, f.ole_name)
        assert ole is not None
        assert len(ole) == 3584


class TestConsistency:
    """Consistency validation tests."""

    def test_formula_count_after_scan(self, sample_docx_with_mathtype):
        """Scanned count matches expected (20 for the real fixture)."""
        formulas, _ = scan_docx(str(sample_docx_with_mathtype))
        mathtype = [f for f in formulas if f.formula_type == FormulaType.MATHTYPE_OLE]
        assert len(mathtype) == 20

    def test_unique_formula_ids(self, sample_docx_with_mathtype):
        """Each formula gets a unique ID."""
        formulas, _ = scan_docx(str(sample_docx_with_mathtype))
        ids = [f.formula_id for f in formulas]
        assert len(ids) == len(set(ids))

    def test_part_name_consistency(self, sample_docx_with_mathtype):
        """All formulas from document.xml have consistent part_name."""
        formulas, _ = scan_docx(str(sample_docx_with_mathtype))
        doc_formulas = [f for f in formulas if f.part_name == "word/document.xml"]
        assert len(doc_formulas) == len(formulas)

    def test_ole_name_resolved_from_rels(self, sample_docx_with_mathtype):
        """OLE names are correctly resolved via .rels mapping."""
        formulas, _ = scan_docx(str(sample_docx_with_mathtype))
        for f in formulas:
            assert f.ole_name.startswith("oleObject")
            assert f.ole_name.endswith(".bin")
