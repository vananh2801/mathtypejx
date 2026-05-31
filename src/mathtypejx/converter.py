"""Converter: MTEF → MathML → OMML conversion chain."""

import re
from pathlib import Path
from typing import Optional

from lxml import etree

from .models import FormulaInfo, FormulaStatus, RiskLevel

# MTEF parsing is handled by the native Python parser in mathtypejx.mtef.

# Common XSL paths
_DEFAULT_XSL_PATHS = [
    Path(r"C:\Program Files (x86)\Microsoft Office\root\Office16\MML2OMML.XSL"),
    Path(r"C:\Program Files\Microsoft Office\root\Office16\MML2OMML.XSL"),
]

# MathML namespace
MATHML_NS = "http://www.w3.org/1998/Math/MathML"

_MATHML_TAGS = {
    "math", "semantics", "annotation", "annotation-xml",
    "mrow", "mi", "mn", "mo", "mtext", "ms", "mglyph",
    "mfrac", "msqrt", "mroot", "msub", "msup", "msubsup",
    "munder", "mover", "munderover", "mfenced", "menclose",
    "mstyle", "mpadded", "mphantom", "mtable", "mtr", "mtd",
    "mmultiscripts", "mprescripts", "none", "mspace",
    "maligngroup", "malignmark", "mlabeledtr", "maction",
}
_TOKEN_TAGS = {"mi", "mn", "mo", "mtext", "ms", "mglyph"}
_OPERATOR_CHARS = set("+-−=*/×÷±∓<>≤≥()[]{}|,.;:⋅·^_∈∉≈≠≡∝∞∑∏∫∮√")
_TOKEN_RE = re.compile(r"\d+(?:\.\d+)?|[A-Za-zΑ-Ωα-ω]+|[^\s]")
_RISK_ORDER = {
    RiskLevel.AUTO_REPLACE: 0,
    RiskLevel.SPOT_CHECK: 1,
    RiskLevel.MANUAL_REVIEW: 2,
    RiskLevel.BLOCKED: 3,
}
_MANUAL_REVIEW_TAGS = {
    "mmultiscripts", "mprescripts", "maligngroup", "malignmark",
    "mphantom", "mpadded", "mspace", "mtable", "mtr", "mtd",
    "munderover", "mroot", "menclose",
}
_SPOT_CHECK_TAGS = {
    "mfrac", "msqrt", "msub", "msup", "msubsup",
    "munder", "mover", "mfenced",
}
_NARY_OPERATORS = {"∑", "∏", "∫", "∮", "⋂", "⋃"}
_FUNCTION_NAMES = {"sin", "cos", "tan", "log", "ln", "lim", "max", "min"}


def convert_formula(
    formula: FormulaInfo,
    xsl_path: Optional[str] = None,
) -> bool:
    """Run the full conversion chain for a single formula.

    Args:
        formula: FormulaInfo with ole_data populated.
        xsl_path: Path to MML2OMML.XSL. Auto-detected if None.
    Returns:
        True if conversion succeeded.
    """
    if not formula.ole_data:
        formula.status = FormulaStatus.FAILED
        formula.error_message = "No OLE data to convert"
        return False

    # Step 1: MTEF → MathML (pure Python)
    mathml = _ole_to_mathml(formula.ole_data, formula.ole_name)
    if mathml is None:
        formula.status = FormulaStatus.FAILED
        formula.error_message = "MTEF to MathML conversion failed"
        return False

    formula.mathml = mathml

    # Step 2: Normalize MathML
    mathml = _normalize_mathml(mathml)
    formula.mathml = mathml

    # Step 3: MathML → OMML (via XSLT)
    omml = _mathml_to_omml(mathml, xsl_path=xsl_path)
    if omml is None:
        formula.status = FormulaStatus.FAILED
        formula.error_message = "MathML to OMML conversion failed"
        return False

    formula.omml = omml

    # Step 3.5: Clean OMML of unsupported character noise
    omml = _clean_omml_unsupported(omml)
    formula.omml = omml

    # Step 4: Validate generated OMML before allowing replacement
    from .validator import validate_conversion_quality
    quality = validate_conversion_quality(mathml, omml)
    formula.quality_errors = quality["errors"]
    formula.quality_warnings = quality["warnings"]
    if not quality["valid"]:
        formula.status = FormulaStatus.FAILED
        formula.risk_level = RiskLevel.BLOCKED
        formula.error_message = (
            "OMML quality validation failed: "
            + "; ".join(quality["errors"][:3])
        )
        return False

    # Step 5: Assess risk level
    formula.risk_level = _assess_risk(mathml, omml)

    formula.status = FormulaStatus.CONVERTED
    return True


def _ole_to_mathml(
    ole_data: Optional[bytes],
    ole_name: str,
) -> Optional[str]:
    """Convert OLE binary to MathML using pure Python MTEF parser."""
    if not ole_data:
        return None

    from .mtef import mtef_to_mathml
    return mtef_to_mathml(ole_data)


def _normalize_mathml(mathml: str) -> str:
    """Normalize MathML before XSLT conversion.

    Fixes:
    - XML declaration parsing quirks
    - Missing real MathML namespaces on elements
    - Bare text in container elements that Office XSLT otherwise drops
    - Empty root degree → msqrt conversion
    """
    try:
        root = etree.fromstring(mathml.encode("utf-8"))
    except etree.XMLSyntaxError:
        mathml = mathml.strip()
        if mathml.startswith("<?xml"):
            idx = mathml.find("?>")
            if idx > 0:
                mathml = mathml[idx + 2:].strip()
        try:
            root = etree.fromstring(mathml.encode("utf-8"))
        except etree.XMLSyntaxError:
            return mathml

    # Always ensure MathML namespace (XSLT outputs elements without ns)
    _ensure_mathml_namespace(root)
    _normalize_bare_text_nodes(root)
    _normalize_root_degree(root)
    # Replace <mtext> → <mi mathvariant="normal"> so MML2OMML.XSL
    # produces <m:sty m:val="p"/> instead of <m:nor/>.  The latter
    # switches to the document body font whose baseline may differ
    # from the math font, making units look like subscripts.
    _mtext_to_normal_mi(root)

    return etree.tostring(root, encoding="unicode", pretty_print=False)


def _local_name(elem) -> str:
    """Return an element's local tag name."""
    try:
        return etree.QName(elem).localname
    except ValueError:
        return str(elem.tag)


def _mtext_to_normal_mi(root) -> None:
    """Replace <mtext> with <mi mathvariant="normal"> in-place.

    MML2OMML.XSL maps ``<mtext>`` → ``<m:rPr><m:nor/></m:rPr>`` which
    switches the run to the document body font.  When the body font
    (e.g. 宋体 / SimSun) has different baseline metrics than the math
    font (Cambria Math), the text appears shifted downward — easily
    mistaken for a subscript.

    ``<mi mathvariant="normal">`` maps to ``<m:rPr><m:sty m:val="p"/></m:rPr>``
    instead.  That keeps the math font and merely disables math-italic,
    so the baseline stays aligned with the rest of the formula.
    """
    mtext_tag = f"{{{MATHML_NS}}}mtext"
    mi_tag = f"{{{MATHML_NS}}}mi"
    for elem in list(root.iter()):
        if elem.tag == mtext_tag:
            elem.tag = mi_tag
            if "mathvariant" not in elem.attrib:
                elem.set("mathvariant", "normal")


def _ensure_mathml_namespace(root) -> None:
    """Recursively put known MathML elements in the MathML namespace."""
    for elem in root.iter():
        if not isinstance(elem.tag, str):
            continue
        qname = etree.QName(elem)
        local = qname.localname
        if local in _MATHML_TAGS and qname.namespace != MATHML_NS:
            elem.tag = f"{{{MATHML_NS}}}{local}"


def _normalize_bare_text_nodes(root) -> None:
    """Wrap bare text in MathML containers as mi/mn/mo/mtext tokens."""
    for elem in list(root.iter()):
        if not isinstance(elem.tag, str):
            continue
        local = _local_name(elem)
        if local in _TOKEN_TAGS:
            continue

        if elem.text and elem.text.strip():
            tokens = _text_to_mathml_tokens(elem.text)
            elem.text = None
            for index, token in enumerate(tokens):
                elem.insert(index, token)

        child_index = 0
        while child_index < len(elem):
            child = elem[child_index]
            if child.tail and child.tail.strip():
                tokens = _text_to_mathml_tokens(child.tail)
                child.tail = None
                insert_at = child_index + 1
                for offset, token in enumerate(tokens):
                    elem.insert(insert_at + offset, token)
                child_index += len(tokens)
            child_index += 1


def _text_to_mathml_tokens(text: str) -> list[etree._Element]:
    """Convert bare MathML text into token elements."""
    tokens = []
    for match in _TOKEN_RE.finditer(text):
        value = match.group(0)
        if not value.strip():
            continue
        tag = _classify_mathml_token(value)
        elem = etree.Element(f"{{{MATHML_NS}}}{tag}")
        elem.text = value
        tokens.append(elem)
    return tokens


def _classify_mathml_token(token: str) -> str:
    """Classify a text token as a MathML token element name."""
    if re.fullmatch(r"\d+(?:\.\d+)?", token):
        return "mn"
    if all(ch in _OPERATOR_CHARS for ch in token):
        return "mo"
    if re.fullmatch(r"[A-Za-zΑ-Ωα-ω]+", token):
        return "mi"
    if len(token) == 1 and token.isalpha():
        return "mi"
    return "mtext"


def _normalize_root_degree(root) -> None:
    """Convert <mroot> with empty first child to <msqrt>.

    MML2OMML.XSL creates an empty <m:deg> for <mroot> when the degree is
    empty, which fails OMML validation. Converting to <msqrt> avoids this.
    """
    MATHML_NS = "http://www.w3.org/1998/Math/MathML"
    for mroot in list(root.iter()):
        if not isinstance(mroot.tag, str):
            continue
        if etree.QName(mroot).localname != "mroot":
            continue
        children = [c for c in mroot if isinstance(c.tag, str)]
        if len(children) >= 1 and (
            len(children[0]) == 0
            and not (children[0].text or "").strip()
            and not children[0].tail
        ):
            # Empty degree — convert mroot to msqrt
            msqrt = etree.Element(f"{{{MATHML_NS}}}msqrt")
            for child in children[1:]:
                mroot.remove(child)
                msqrt.append(child)
            # Copy tail text
            msqrt.tail = mroot.tail
            parent = mroot.getparent()
            if parent is not None:
                idx = list(parent).index(mroot)
                parent.remove(mroot)
                parent.insert(idx, msqrt)


def _mathml_to_omml(mathml: str, xsl_path: Optional[str] = None) -> Optional[str]:
    """Convert MathML to OMML using MML2OMML.XSL."""
    xsl_file = _find_xsl(xsl_path)
    if xsl_file is None:
        return None

    try:
        mathml_doc = etree.fromstring(mathml.encode("utf-8"))
        xsl_doc = etree.parse(str(xsl_file))
        transform = etree.XSLT(xsl_doc)
        result = transform(mathml_doc)
        omml = etree.tostring(result, encoding="unicode", pretty_print=False)
        return omml
    except Exception:
        return None


def _find_xsl(xsl_path: Optional[str] = None) -> Optional[Path]:
    """Find the MML2OMML.XSL file."""
    if xsl_path:
        p = Path(xsl_path)
        if p.exists():
            return p

    for p in _DEFAULT_XSL_PATHS:
        if p.exists():
            return p

    return None


def _assess_risk(mathml: str, omml: str) -> RiskLevel:
    """Assess replacement risk based on complexity and style-sensitive constructs."""
    try:
        root = etree.fromstring(mathml.encode("utf-8"))
        node_count = sum(1 for _ in root.iter())

        if node_count < 15:
            risk = RiskLevel.AUTO_REPLACE
        elif node_count < 40:
            risk = RiskLevel.SPOT_CHECK
        else:
            risk = RiskLevel.MANUAL_REVIEW

        if _mathml_contains_any(root, _MANUAL_REVIEW_TAGS):
            risk = _max_risk(risk, RiskLevel.MANUAL_REVIEW)
        if (
            _mathml_contains_any(root, _SPOT_CHECK_TAGS)
            or _mathml_has_known_nary_operator(root)
            or _mathml_has_known_function(root)
        ):
            risk = _max_risk(risk, RiskLevel.SPOT_CHECK)
        return risk
    except Exception:
        return RiskLevel.BLOCKED


def _max_risk(left: RiskLevel, right: RiskLevel) -> RiskLevel:
    return left if _RISK_ORDER[left] >= _RISK_ORDER[right] else right


def _mathml_contains_any(root, tag_names: set[str]) -> bool:
    for elem in root.iter():
        if isinstance(elem.tag, str) and _local_name(elem) in tag_names:
            return True
    return False


def _mathml_has_known_nary_operator(root) -> bool:
    return any(token in _NARY_OPERATORS for token in _mathml_token_texts(root))


def _mathml_has_known_function(root) -> bool:
    return any(token in _FUNCTION_NAMES for token in _mathml_token_texts(root))


def _mathml_token_texts(root) -> list[str]:
    tokens = []
    for elem in root.iter():
        if isinstance(elem.tag, str) and _local_name(elem) in {"mi", "mn", "mo", "mtext", "ms"}:
            text = "".join(elem.itertext()).strip()
            if text:
                tokens.append(text)
    return tokens


def _clean_omml_unsupported(omml: str) -> str:
    """Remove 'Unsupported character' noise text from OMML output.

    The MML2OMML.XSL inserts 'Unsupported character' text for code points
    it can't map. This strips those text nodes from the OMML so they don't
    appear as noise in the output document.
    """
    if "Unsupported" not in omml:
        return omml

    try:
        root = etree.fromstring(omml.encode("utf-8"))
    except etree.XMLSyntaxError:
        return omml

    OMML_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
    unsupported_re = re.compile(r"Unsupported\s*characters?\s*", re.IGNORECASE)
    changed = False
    for mt in list(root.iter(f"{{{OMML_NS}}}t")):
        if mt.text and "Unsupported" in mt.text:
            mt.text = unsupported_re.sub("", mt.text).strip()
            if not mt.text:
                # Remove the entire <m:r> that contains this empty <m:t>
                parent = mt.getparent()
                if parent is not None and etree.QName(parent).localname == "r":
                    grandparent = parent.getparent()
                    if grandparent is not None:
                        grandparent.remove(parent)
                        changed = True

    if changed:
        return etree.tostring(root, encoding="unicode", pretty_print=False)
    return omml


def convert_mathml_file_to_omml(
    mathml_path: str,
    xsl_path: Optional[str] = None,
) -> Optional[str]:
    """Convert a MathML file to OMML (utility for testing)."""
    mathml = Path(mathml_path).read_text(encoding="utf-8")
    return _mathml_to_omml(mathml, xsl_path=xsl_path)
