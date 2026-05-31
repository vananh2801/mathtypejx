"""Unit tests for converter module (MathML→OMML path)."""

import pytest
from lxml import etree
from mathtypejx.converter import (
    MATHML_NS,
    _normalize_mathml,
    _mathml_to_omml,
    _assess_risk,
    convert_formula,
)
from mathtypejx.models import FormulaInfo, FormulaStatus, RiskLevel
from mathtypejx.validator import validate_conversion_quality


# Sample MathML strings for testing
SIMPLE_MATHML = (
    '<math xmlns="http://www.w3.org/1998/Math/MathML">'
    "<mfrac><mi>a</mi><mi>b</mi></mfrac>"
    "</math>"
)

COMPLEX_MATHML = (
    '<math xmlns="http://www.w3.org/1998/Math/MathML">'
    "<mrow>"
    "<mi>x</mi><mo>=</mo>"
    "<mfrac>"
    "<mrow><mo>-</mo><mi>b</mi><mo>±</mo><msqrt><msup><mi>b</mi><mn>2</mn></msup>"
    "<mo>-</mo><mn>4</mn><mi>a</mi><mi>c</mi></msqrt></mrow>"
    "<mrow><mn>2</mn><mi>a</mi></mrow>"
    "</mfrac>"
    "</mrow>"
    "</math>"
)


class TestMathMLNormalize:
    """MathML normalization tests."""

    def test_preserves_valid_mathml(self):
        result = _normalize_mathml(SIMPLE_MATHML)
        assert "mfrac" in result
        assert "http://www.w3.org/1998/Math/MathML" in result

    def test_adds_missing_namespace(self):
        mathml_no_ns = "<math><mfrac><mi>a</mi><mi>b</mi></mfrac></math>"
        result = _normalize_mathml(mathml_no_ns)
        root = etree.fromstring(result.encode("utf-8"))
        assert etree.QName(root).namespace == MATHML_NS
        assert etree.QName(root[0]).namespace == MATHML_NS

    def test_handles_xml_declaration(self):
        mathml_with_decl = (
            '<?xml version="1.0" encoding="UTF-8"?>\n'
            '<math xmlns="http://www.w3.org/1998/Math/MathML">'
            "<mi>x</mi></math>"
        )
        result = _normalize_mathml(mathml_with_decl)
        assert "<mi>x</mi>" in result

    def test_handles_empty_string(self):
        result = _normalize_mathml("")
        assert result == ""

    def test_wraps_bare_subscript_text(self):
        mathml = "<math><msub><mrow>A</mrow><mrow>6</mrow></msub></math>"
        result = _normalize_mathml(mathml)
        root = etree.fromstring(result.encode("utf-8"))
        ns = {"m": MATHML_NS}
        assert root.xpath("string(.//m:msub/*[1]/m:mi)", namespaces=ns) == "A"
        assert root.xpath("string(.//m:msub/*[2]/m:mn)", namespaces=ns) == "6"

    def test_tokenizes_bare_mrow_text(self):
        result = _normalize_mathml("<math><mrow>A+6</mrow></math>")
        root = etree.fromstring(result.encode("utf-8"))
        names = [etree.QName(e).localname for e in root.xpath(".//*[local-name() != 'math' and local-name() != 'mrow']")]
        texts = [e.text for e in root.xpath(".//*[local-name() != 'math' and local-name() != 'mrow']")]
        assert names == ["mi", "mo", "mn"]
        assert texts == ["A", "+", "6"]

    def test_tokenizes_child_tail_text(self):
        result = _normalize_mathml("<math><mrow><mi>A</mi>+6x</mrow></math>")
        root = etree.fromstring(result.encode("utf-8"))
        tokens = [
            (etree.QName(e).localname, e.text)
            for e in root.xpath(".//*[local-name()='mi' or local-name()='mn' or local-name()='mo']")
        ]
        assert tokens == [("mi", "A"), ("mo", "+"), ("mn", "6"), ("mi", "x")]

    def test_namespaces_multiscript_tags(self):
        mathml = (
            "<math><mmultiscripts><mi>A</mi><mi>i</mi><none/>"
            "<mprescripts/><mi>j</mi><none/></mmultiscripts></math>"
        )
        root = etree.fromstring(_normalize_mathml(mathml).encode("utf-8"))
        for tag in ["mmultiscripts", "mprescripts", "none"]:
            elems = root.xpath(f".//*[local-name()='{tag}']")
            assert elems
            assert all(etree.QName(elem).namespace == MATHML_NS for elem in elems)

    def test_namespaces_spacing_and_alignment_tags(self):
        mathml = "<math><mrow><maligngroup/><mi>x</mi><malignmark/><mspace/><mpadded><mi>y</mi></mpadded></mrow></math>"
        root = etree.fromstring(_normalize_mathml(mathml).encode("utf-8"))
        for tag in ["maligngroup", "malignmark", "mspace", "mpadded"]:
            elems = root.xpath(f".//*[local-name()='{tag}']")
            assert elems
            assert all(etree.QName(elem).namespace == MATHML_NS for elem in elems)


class TestMathMLToOMML:
    """MathML to OMML conversion tests."""

    def test_simple_conversion(self):
        omml = _mathml_to_omml(SIMPLE_MATHML)
        if omml is None:
            pytest.skip("MML2OMML.XSL not available")
        assert "oMath" in omml

    def test_complex_conversion(self):
        omml = _mathml_to_omml(COMPLEX_MATHML)
        if omml is None:
            pytest.skip("MML2OMML.XSL not available")
        assert "oMath" in omml

    @pytest.mark.parametrize("mathml", [
        '<math xmlns="http://www.w3.org/1998/Math/MathML"><mroot><mi>x</mi><mn>3</mn></mroot></math>',
        '<math xmlns="http://www.w3.org/1998/Math/MathML"><mtable><mtr><mtd><mi>a</mi></mtd><mtd><mi>b</mi></mtd></mtr><mtr><mtd><mi>c</mi></mtd><mtd><mi>d</mi></mtd></mtr></mtable></math>',
        '<math xmlns="http://www.w3.org/1998/Math/MathML"><mfenced><mi>x</mi></mfenced></math>',
    ])
    def test_style_smoke_conversion_quality(self, mathml):
        omml = _mathml_to_omml(mathml)
        if omml is None:
            pytest.skip("MML2OMML.XSL not available")
        assert "oMath" in omml
        quality = validate_conversion_quality(mathml, omml)
        assert quality["valid"], quality["errors"]

    @pytest.mark.parametrize("mathml", [
        '<math xmlns="http://www.w3.org/1998/Math/MathML"><mover><mi>x</mi><mo>¯</mo></mover></math>',
        '<math xmlns="http://www.w3.org/1998/Math/MathML"><munderover><mo>∑</mo><mi>i</mi><mi>n</mi></munderover><mi>x</mi></math>',
    ])
    def test_style_smoke_conversion_reports_quality(self, mathml):
        omml = _mathml_to_omml(mathml)
        if omml is None:
            pytest.skip("MML2OMML.XSL not available")
        assert "oMath" in omml
        quality = validate_conversion_quality(mathml, omml)
        assert "valid" in quality
        assert "errors" in quality
        assert "warnings" in quality


class TestRiskAssessment:
    """Risk level classification tests."""

    def test_simple_formula_auto_replace(self):
        risk = _assess_risk('<math xmlns="http://www.w3.org/1998/Math/MathML"><mi>x</mi></math>', "")
        assert risk == RiskLevel.AUTO_REPLACE

    def test_simple_fraction_spot_check(self):
        risk = _assess_risk(SIMPLE_MATHML, "")
        assert risk == RiskLevel.SPOT_CHECK

    def test_complex_formula_higher_risk(self):
        risk = _assess_risk(COMPLEX_MATHML, "")
        assert risk in (RiskLevel.SPOT_CHECK, RiskLevel.MANUAL_REVIEW)

    def test_empty_mathml_blocked(self):
        risk = _assess_risk("", "")
        assert risk == RiskLevel.BLOCKED

    def test_invalid_mathml_blocked(self):
        risk = _assess_risk("invalid", "")
        assert risk == RiskLevel.BLOCKED

    def test_mspace_manual_review(self):
        risk = _assess_risk('<math xmlns="http://www.w3.org/1998/Math/MathML"><mspace/></math>', "")
        assert risk == RiskLevel.MANUAL_REVIEW

    def test_matrix_manual_review(self):
        mathml = '<math xmlns="http://www.w3.org/1998/Math/MathML"><mtable><mtr><mtd><mi>a</mi></mtd></mtr></mtable></math>'
        risk = _assess_risk(mathml, "")
        assert risk == RiskLevel.MANUAL_REVIEW

    def test_mroot_manual_review(self):
        mathml = '<math xmlns="http://www.w3.org/1998/Math/MathML"><mroot><mi>x</mi><mn>3</mn></mroot></math>'
        risk = _assess_risk(mathml, "")
        assert risk == RiskLevel.MANUAL_REVIEW

    def test_mfenced_spot_check(self):
        mathml = '<math xmlns="http://www.w3.org/1998/Math/MathML"><mfenced><mi>x</mi></mfenced></math>'
        risk = _assess_risk(mathml, "")
        assert risk in (RiskLevel.SPOT_CHECK, RiskLevel.MANUAL_REVIEW)


class TestConverterEdgeCases:
    """Edge case tests for the converter."""

    def test_convert_no_ole_data(self):
        """Converting a formula with no OLE data should fail."""
        formula = FormulaInfo(
            formula_id="F0100",
            ole_name="empty.bin",
            part_name="word/document.xml",
            rels_path="r",
            relationship_id="r1",
            prog_id="Equation.DSMT4",
            ole_data=None,
        )
        result = convert_formula(formula)
        assert not result
        assert formula.status == FormulaStatus.FAILED

    def test_normalize_non_xml(self):
        """Normalize handles non-XML gracefully."""
        result = _normalize_mathml("just plain text")
        assert isinstance(result, str)

    def test_bad_omml_quality_blocks_conversion(self, monkeypatch):
        """Converted-but-lossy OMML should not be marked CONVERTED."""
        formula = FormulaInfo(
            formula_id="F0101",
            ole_name="bad.bin",
            part_name="word/document.xml",
            rels_path="r",
            relationship_id="r1",
            prog_id="Equation.DSMT4",
            ole_data=b"ole",
        )
        mathml = (
            '<math xmlns="http://www.w3.org/1998/Math/MathML">'
            "<msub><mi>A</mi><mn>6</mn></msub></math>"
        )
        omml = (
            '<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">'
            "<m:sSub><m:e><m:r><m:t>A</m:t></m:r></m:e><m:sub/></m:sSub>"
            "</m:oMath>"
        )
        monkeypatch.setattr("mathtypejx.converter._ole_to_mathml", lambda *args, **kwargs: mathml)
        monkeypatch.setattr("mathtypejx.converter._mathml_to_omml", lambda *args, **kwargs: omml)

        result = convert_formula(formula)

        assert not result
        assert formula.status == FormulaStatus.FAILED
        assert formula.risk_level == RiskLevel.BLOCKED
        assert "quality validation" in formula.error_message
        assert formula.quality_errors

    def test_good_omml_quality_allows_conversion(self, monkeypatch):
        """Quality-valid OMML should be marked CONVERTED."""
        formula = FormulaInfo(
            formula_id="F0102",
            ole_name="good.bin",
            part_name="word/document.xml",
            rels_path="r",
            relationship_id="r1",
            prog_id="Equation.DSMT4",
            ole_data=b"ole",
        )
        mathml = (
            '<math xmlns="http://www.w3.org/1998/Math/MathML">'
            "<msub><mi>A</mi><mn>6</mn></msub></math>"
        )
        omml = (
            '<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">'
            "<m:sSub>"
            "<m:e><m:r><m:t>A</m:t></m:r></m:e>"
            "<m:sub><m:r><m:t>6</m:t></m:r></m:sub>"
            "</m:sSub>"
            "</m:oMath>"
        )
        monkeypatch.setattr("mathtypejx.converter._ole_to_mathml", lambda *args, **kwargs: mathml)
        monkeypatch.setattr("mathtypejx.converter._mathml_to_omml", lambda *args, **kwargs: omml)

        result = convert_formula(formula)

        assert result
        assert formula.status == FormulaStatus.CONVERTED
        assert formula.quality_errors == []
