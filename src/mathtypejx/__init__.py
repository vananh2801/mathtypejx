"""MathType OLE formula to OMML pipeline converter for Word docx.

Usage:
    from mathtypejx import convert_mathtype_to_omml

    report = convert_mathtype_to_omml("input.docx", "output.docx")
    print(f"Converted {report.succeeded}/{report.total_ole_objects} formulas")
"""

__version__ = "0.1.0"

from .service import (
    convert_mathtype_to_omml,
    convert_batch,
    health_check,
)
from .models import (
    ConversionReport,
    FormulaInfo,
    FormulaStatus,
    RiskLevel,
)
