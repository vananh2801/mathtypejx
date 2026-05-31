"""Data models for the MathType conversion pipeline."""

from dataclasses import dataclass, field
from enum import Enum
from pathlib import Path
from typing import Optional


class FormulaStatus(Enum):
    """Status of a formula in the conversion pipeline."""

    DETECTED = "detected"            # Found in docx, not yet processed
    EXTRACTED = "extracted"          # MTEF bytes extracted from OLE
    CONVERTED = "converted"          # MathML/OMML generated
    REPLACED = "replaced"            # Successfully replaced in docx
    FAILED = "failed"                # Conversion failed
    SKIPPED = "skipped"              # Intentionally skipped (not MathType)


class RiskLevel(Enum):
    """Risk classification for quality gating."""

    AUTO_REPLACE = "auto_replace"        # Simple formula, safe to auto-replace
    SPOT_CHECK = "spot_check"             # Medium complexity, sample check
    MANUAL_REVIEW = "manual_review"      # High complexity, human review needed
    BLOCKED = "blocked"                   # Conversion failed, keep original


class FormulaType(Enum):
    """Type of formula detected in the document."""

    MATHTYPE_OLE = "mathtype-ole"           # MathType OLE object (ProgID: Equation.DSMT4)
    EQUATION_EDITOR_3 = "equation-editor-3" # Equation Editor 3.x OLE
    OMML_NATIVE = "omml-native"             # Already OMML
    UNKNOWN = "unknown"


@dataclass
class FormulaInfo:
    """Represents a single formula found in a docx document."""

    # Identity
    formula_id: str                          # Unique ID within the conversion run
    ole_name: str                            # e.g. "oleObject1.bin"

    # Location in docx
    part_name: str                           # e.g. "word/document.xml"
    rels_path: str                           # e.g. "word/_rels/document.xml.rels"
    relationship_id: str                     # e.g. "rId8"
    prog_id: str                             # e.g. "Equation.DSMT4"
    para_index: int = -1                     # Index within the part XML
    run_index: int = -1                      # Index within the paragraph

    # Conversion state
    status: FormulaStatus = FormulaStatus.DETECTED
    risk_level: RiskLevel = RiskLevel.AUTO_REPLACE
    formula_type: FormulaType = FormulaType.MATHTYPE_OLE
    error_message: Optional[str] = None

    # Data payloads (populated through pipeline)
    ole_data: Optional[bytes] = None         # Raw OLE binary (oleObjectN.bin)
    mtef_bytes: Optional[bytes] = None       # Extracted MTEF bytes (after 28-byte header)
    mtef_version: Optional[int] = None       # MTEF version (3, 4, 5)
    mathml: Optional[str] = None             # MathML XML string
    omml: Optional[str] = None               # OMML XML string (m:oMath)

    # Conversion quality diagnostics
    quality_errors: list[str] = field(default_factory=list)
    quality_warnings: list[str] = field(default_factory=list)

    # Metadata
    is_inline: bool = True                   # True = inline formula, False = display


@dataclass
class ConversionReport:
    """Summary report for a conversion run."""

    input_docx: str = ""
    output_docx: str = ""
    total_ole_objects: int = 0
    formulas: list[FormulaInfo] = field(default_factory=list)

    @property
    def succeeded(self) -> int:
        return sum(1 for f in self.formulas if f.status == FormulaStatus.REPLACED)

    @property
    def failed(self) -> int:
        return sum(1 for f in self.formulas if f.status == FormulaStatus.FAILED)

    @property
    def skipped(self) -> int:
        return sum(1 for f in self.formulas if f.status == FormulaStatus.SKIPPED)

    @property
    def by_risk(self) -> dict:
        counts = {r: 0 for r in RiskLevel}
        for f in self.formulas:
            counts[f.risk_level] += 1
        return counts

    def to_dict(self) -> dict:
        return {
            "input_docx": self.input_docx,
            "output_docx": self.output_docx,
            "total_ole_objects": self.total_ole_objects,
            "succeeded": self.succeeded,
            "failed": self.failed,
            "skipped": self.skipped,
            "by_risk": {k.value: v for k, v in self.by_risk.items()},
            "formulas": [
                {
                    "formula_id": f.formula_id,
                    "ole_name": f.ole_name,
                    "part_name": f.part_name,
                    "prog_id": f.prog_id,
                    "status": f.status.value,
                    "risk_level": f.risk_level.value,
                    "formula_type": f.formula_type.value,
                    "error_message": f.error_message,
                    "quality_errors": f.quality_errors,
                    "quality_warnings": f.quality_warnings,
                    "mtef_version": f.mtef_version,
                }
                for f in self.formulas
            ],
        }
