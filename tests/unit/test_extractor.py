"""Unit tests for extractor module."""

import struct

from mathtypejx.extractor import extract_mtef, detect_mtef_version, inspect_ole
from mathtypejx.models import FormulaInfo, FormulaStatus, FormulaType


class TestExtractor:
    """Tests for the MTEF extractor."""

    def test_extract_from_valid_ole(self, mock_ole_binary):
        """Extract MTEF from a valid OLE binary."""
        formula = FormulaInfo(
            formula_id="F0001",
            ole_name="test.bin",
            part_name="word/document.xml",
            rels_path="word/_rels/document.xml.rels",
            relationship_id="rId1",
            prog_id="Equation.DSMT4",
            ole_data=mock_ole_binary,
        )

        ok = extract_mtef(formula)
        assert ok
        assert formula.status == FormulaStatus.EXTRACTED
        assert formula.mtef_bytes is not None
        assert len(formula.mtef_bytes) > 0

    def test_extract_empty_ole(self):
        """Handle empty OLE data gracefully."""
        formula = FormulaInfo(
            formula_id="F0002",
            ole_name="empty.bin",
            part_name="word/document.xml",
            rels_path="word/_rels/document.xml.rels",
            relationship_id="rId2",
            prog_id="Equation.DSMT4",
            ole_data=b"",
        )

        ok = extract_mtef(formula)
        assert not ok
        assert formula.status == FormulaStatus.FAILED
        assert formula.error_message is not None

    def test_extract_none_ole(self):
        """Handle None OLE data."""
        formula = FormulaInfo(
            formula_id="F0003",
            ole_name="none.bin",
            part_name="word/document.xml",
            rels_path="word/_rels/document.xml.rels",
            relationship_id="rId3",
            prog_id="Equation.DSMT4",
            ole_data=None,
        )

        ok = extract_mtef(formula)
        assert not ok
        assert formula.status == FormulaStatus.FAILED

    def test_detect_mtef_version(self, mock_ole_binary):
        """MTEF version detection works."""
        formula = FormulaInfo(
            formula_id="F0004",
            ole_name="test.bin",
            part_name="word/document.xml",
            rels_path="word/_rels/document.xml.rels",
            relationship_id="rId1",
            prog_id="Equation.DSMT4",
            ole_data=mock_ole_binary,
        )

        extract_mtef(formula)
        version = detect_mtef_version(formula.mtef_bytes)
        assert version is not None
        assert version == 5  # MTEF v5 start byte

    def test_detect_version_none(self):
        """detect_mtef_version returns None for empty data."""
        assert detect_mtef_version(b"") is None
        assert detect_mtef_version(None) is None

    def test_inspect_ole_info(self, mock_ole_binary):
        """Inspect OLE returns useful info."""
        info = inspect_ole(mock_ole_binary)
        assert "streams" in info
        assert "header" in info
        assert info["header"] is not None
        assert info["mtef_version"] == 5

    def test_inspect_bad_data(self):
        """Inspect handles bad data gracefully."""
        info = inspect_ole(b"not an ole file")
        assert "error" in info


class TestConsistency:
    """Consistency checks for extraction."""

    def test_extraction_idempotent(self, mock_ole_binary):
        """Extracting the same data twice gives same result."""
        f1 = FormulaInfo(
            formula_id="F0010", ole_name="a.bin",
            part_name="word/document.xml", rels_path="r", relationship_id="r1",
            prog_id="Equation.DSMT4", ole_data=mock_ole_binary,
        )
        f2 = FormulaInfo(
            formula_id="F0011", ole_name="a.bin",
            part_name="word/document.xml", rels_path="r", relationship_id="r1",
            prog_id="Equation.DSMT4", ole_data=mock_ole_binary,
        )

        extract_mtef(f1)
        extract_mtef(f2)
        assert f1.mtef_bytes == f2.mtef_bytes

    def test_mtef_size_positive(self, mock_ole_binary):
        """Extracted MTEF has positive size."""
        formula = FormulaInfo(
            formula_id="F0012", ole_name="a.bin",
            part_name="word/document.xml", rels_path="r", relationship_id="r1",
            prog_id="Equation.DSMT4", ole_data=mock_ole_binary,
        )
        extract_mtef(formula)
        assert formula.mtef_bytes is not None
        assert len(formula.mtef_bytes) > 0
