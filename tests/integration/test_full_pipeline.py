"""Integration tests for the full conversion pipeline."""

import json

from mathtypejx.scanner import scan_docx, read_ole_binary
from mathtypejx.extractor import extract_mtef, detect_mtef_version
from mathtypejx.converter import convert_formula
from mathtypejx.replacer import replace_formulas
from mathtypejx.validator import (
    make_report,
    validate_formula_count,
    validate_output_docx,
    validate_conversion_quality,
)
from mathtypejx.models import FormulaStatus, FormulaType


class TestFullPipeline:
    """End-to-end pipeline tests."""

    def test_pipeline_no_formulas(self, sample_docx_no_formulas, temp_dir):
        """Full pipeline on a docx without formulas."""
        output = temp_dir / "output.docx"
        formulas, workdir = scan_docx(str(sample_docx_no_formulas))
        mathtype = [f for f in formulas if f.formula_type == FormulaType.MATHTYPE_OLE]
        assert len(mathtype) == 0

        replace_formulas(workdir, formulas, str(output))
        val = validate_output_docx(str(output))
        assert val["valid"], f"Validation errors: {val['errors']}"

    def test_pipeline_scan_extract_flow(self, sample_docx_with_mathtype):
        """Scan → Extract flow works on real MathType docx."""
        formulas, workdir = scan_docx(str(sample_docx_with_mathtype))
        mathtype = [f for f in formulas if f.formula_type == FormulaType.MATHTYPE_OLE]
        assert len(mathtype) == 20

        # Extract first 3
        for f in mathtype[:3]:
            ole = read_ole_binary(workdir, f.ole_name)
            f.ole_data = ole
            assert ole is not None
            ok = extract_mtef(f)
            assert ok, f"Extraction failed for {f.ole_name}: {f.error_message}"
            version = detect_mtef_version(f.mtef_bytes)
            assert version == 5
            assert f.status == FormulaStatus.EXTRACTED

    def test_formula_count_conservation(self, sample_docx_with_mathtype, temp_dir):
        """Formula count conservation across full pipeline."""
        formulas, workdir = scan_docx(str(sample_docx_with_mathtype))
        output = temp_dir / "output_multi.docx"

        for f in formulas:
            ole = read_ole_binary(workdir, f.ole_name)
            f.ole_data = ole
            if ole:
                extract_mtef(f)
                if f.status == FormulaStatus.EXTRACTED:
                    convert_formula(f)

        replace_formulas(workdir, formulas, str(output))

        report = make_report(str(sample_docx_with_mathtype), str(output), formulas)
        result = validate_formula_count(report)
        assert result["conserved"], (
            f"Formula count not conserved: detected={result['total_detected']}, "
            f"succeeded={result['succeeded']}, failed={result['failed']}"
        )

    def test_report_json_serializable(self, sample_docx_with_mathtype, temp_dir):
        """Conversion report can be serialized to JSON."""
        formulas, workdir = scan_docx(str(sample_docx_with_mathtype))
        output = temp_dir / "output.docx"

        for f in formulas[:3]:  # Test with first 3 only
            ole = read_ole_binary(workdir, f.ole_name)
            f.ole_data = ole
            if ole:
                extract_mtef(f)
                if f.status == FormulaStatus.EXTRACTED:
                    convert_formula(f)

        replace_formulas(workdir, formulas, str(output))
        report = make_report(str(sample_docx_with_mathtype), str(output), formulas)
        data = report.to_dict()
        json_str = json.dumps(data, indent=2)
        assert len(json_str) > 0
        assert "formulas" in data
        assert "total_ole_objects" in data

    def test_output_docx_valid(self, sample_docx_with_mathtype, temp_dir):
        """Output docx passes structural validation."""
        formulas, workdir = scan_docx(str(sample_docx_with_mathtype))
        output = temp_dir / "output_valid.docx"

        for f in formulas[:3]:
            ole = read_ole_binary(workdir, f.ole_name)
            f.ole_data = ole
            if ole:
                extract_mtef(f)
                if f.status == FormulaStatus.EXTRACTED:
                    convert_formula(f)

        replace_formulas(workdir, formulas, str(output))
        val = validate_output_docx(str(output))
        assert val["valid"], f"Validation failed: {val['errors']}"


class TestConsistencyChecks:
    """Consistency validation across pipeline phases."""

    def test_no_formula_loss_in_scan(self, sample_docx_with_mathtype):
        """Scan finds all 20 MathType formulas."""
        formulas, _ = scan_docx(str(sample_docx_with_mathtype))
        mathtype = [f for f in formulas if f.formula_type == FormulaType.MATHTYPE_OLE]
        assert len(mathtype) == 20

    def test_status_progression(self, sample_docx_with_mathtype):
        """Status progresses: DETECTED → EXTRACTED → CONVERTED."""
        formulas, workdir = scan_docx(str(sample_docx_with_mathtype))
        f = formulas[0]
        assert f.status == FormulaStatus.DETECTED

        ole = read_ole_binary(workdir, f.ole_name)
        f.ole_data = ole
        extract_mtef(f)
        assert f.status == FormulaStatus.EXTRACTED

        if f.status == FormulaStatus.EXTRACTED:
            convert_formula(f)
            assert f.status in (FormulaStatus.CONVERTED, FormulaStatus.FAILED)
            if f.status == FormulaStatus.CONVERTED:
                quality = validate_conversion_quality(f.mathml, f.omml)
                assert quality["valid"], quality["errors"]
