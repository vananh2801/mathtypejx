"""Validator: quality gating for the MathType conversion pipeline."""

import re
import zipfile
from collections import Counter
from pathlib import Path
from typing import Optional

from lxml import etree

from .models import ConversionReport, FormulaInfo, FormulaStatus, RiskLevel


# Minimum expected structure inside a valid docx
REQUIRED_DOCX_ENTRIES = [
    "[Content_Types].xml",
    "word/document.xml",
]

MATHML_NS = "http://www.w3.org/1998/Math/MathML"
OMML_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"

_MATHML_TOKEN_TAGS = {"mi", "mn", "mo", "mtext", "ms"}
_EMPTY_SLOT_TAGS = {"e", "sub", "sup", "num", "den", "deg", "lim", "fName", "chr"}

# Empty expression slots can be valid in structures such as matrix cells.
# Empty sub/sup slots silently drop script content and must remain blocking.
_ACCEPTABLE_EMPTY_SLOTS = {"e", "lim"}

# Empty sub/sup are valid placeholders inside sPre (prescripts) and
# sSubSup when only one of the pair is present (e.g. isotope notation
# with only a presuperscript and no presubscript).
_SLOTS_ACCEPTABLE_EMPTY_IN = {"sPre", "sSubSup"}
_STRUCTURE_RULES = {
    "mfrac": ("f",),
    "msub": ("sSub", "sSubSup", "nary"),
    "msup": ("sSup", "sSubSup", "nary"),
    "msubsup": ("sSubSup", "nary"),
    "msqrt": ("rad",),
    "mroot": ("rad",),
    # "mtable" and "mtr" omitted — MML2OMML.XSL matrix conversion is
    # structurally complex; OMML matrix shape is validated separately
    # in _validate_matrix_shape which does detailed row/cell counting,
    "mfenced": ("d",),
    "menclose": ("borderBox", "box"),
    "munder": ("limLow", "groupChr"),
    "mover": ("limUpp", "acc", "bar", "groupChr"),
    "munderover": ("limLow", "limUpp", "nary"),
    "mmultiscripts": ("sPre", "sSubSup", "sSub", "sSup"),
}
_NARY_OPERATORS = {"∑", "∏", "∫", "∮", "⋂", "⋃"}
_ACCENT_MARKS = {"¯", "‾", "^", "~", "˙", "¨", "→", "←", "↔"}

# Characters that MML2OMML.XSL represents structurally rather than as text
_STRUCTURAL_CHARS = _ACCENT_MARKS | {
    "\u200b",  # ZERO WIDTH SPACE (MathType spacing)
    "\u2009",  # THIN SPACE
    "\u200a",  # HAIR SPACE
    "\u205f",  # MEDIUM MATHEMATICAL SPACE
    "\u2003",  # EM SPACE
    "/",        # slash (can be converted to bevelled fraction)
    "\ufe39",  # Presentation form
    "\u2322",  # FROWN / over-arc (converted to m:acc in OMML)
    "\u2323",  # SMILE / under-arc (converted to m:acc in OMML)
    "\u2211",  # N-ARY SUMMATION (∑ — converted to m:nary in OMML)
    "\u222b",  # INTEGRAL (∫ — converted to m:nary in OMML)
    "\u220f",  # N-ARY PRODUCT (∏ — converted to m:nary in OMML)
    "\u2210",  # N-ARY COPRODUCT (∐ — converted to m:nary in OMML)
    "\u22c3",  # N-ARY UNION (⋃ — converted to m:nary in OMML)
    "\u22c2",  # N-ARY INTERSECTION (⋂ — converted to m:nary in OMML)
}
_FUNCTION_NAMES = {"sin", "cos", "tan", "log", "ln", "lim", "max", "min"}
_EQUIVALENT_CHARS = str.maketrans({
    "−": "-",
    "–": "-",
    "—": "-",
    "，": ",",
    "（": "(",
    "）": ")",
})


def validate_conversion_quality(mathml: str, omml: str) -> dict:
    """Validate semantic quality of a MathML → OMML conversion."""
    result = {"valid": True, "errors": [], "warnings": [], "metrics": {}}

    try:
        mathml_root = etree.fromstring(mathml.encode("utf-8"))
    except Exception as e:
        result["valid"] = False
        result["errors"].append(f"MathML parse error: {e}")
        return result

    try:
        omml_root = etree.fromstring(omml.encode("utf-8"))
    except Exception as e:
        result["valid"] = False
        result["errors"].append(f"OMML parse error: {e}")
        return result

    omml_root_name = _local_name(omml_root)
    omml_root_ns = etree.QName(omml_root).namespace
    if omml_root_ns != OMML_NS or omml_root_name not in {"oMath", "oMathPara"}:
        result["errors"].append(f"Unexpected OMML root: {omml_root.tag}")

    empty_slots = _find_empty_omml_slots(omml_root)
    for slot in empty_slots:
        result["errors"].append(f"Empty OMML critical slot: m:{slot}")

    mathml_tokens = _extract_mathml_tokens(mathml_root)
    omml_tokens = _extract_omml_tokens(omml_root)
    omml_text = _extract_omml_text(omml_root)
    missing_tokens = _missing_tokens(mathml_tokens, omml_tokens, omml_text)
    for token, expected, found in missing_tokens:
        result["errors"].append(
            f"OMML token loss: expected token {token!r} count {expected}, found {found}"
        )

    missing_structures = _missing_structures(mathml_root, omml_root)
    for mathml_tag, expected, found in missing_structures:
        result["errors"].append(
            f"OMML structure loss: MathML {mathml_tag} count {expected}, OMML count {found}"
        )

    style_errors, style_warnings = _validate_style_semantics(mathml_root, omml_root)
    result["errors"].extend(style_errors)
    result["warnings"].extend(style_warnings)

    result["metrics"] = {
        "mathml_tokens": mathml_tokens,
        "omml_tokens": omml_tokens,
        "omml_text": omml_text,
        "empty_slots": empty_slots,
        "missing_tokens": missing_tokens,
        "missing_structures": missing_structures,
        "style_errors": style_errors,
        "style_warnings": style_warnings,
    }
    result["valid"] = not result["errors"]
    return result


def _local_name(elem) -> str:
    return etree.QName(elem).localname


def _find_empty_omml_slots(root) -> list[str]:
    empty = []
    for elem in root.iter():
        if not isinstance(elem.tag, str):
            continue
        qname = etree.QName(elem)
        if qname.namespace != OMML_NS or qname.localname not in _EMPTY_SLOT_TAGS:
            continue
        # Allow empty m:deg when m:degHide is "on" (square root notation)
        if qname.localname == "deg" and _deg_is_hidden(elem):
            continue
        # Allow empty m:e in contexts such as matrix cells.
        if qname.localname in _ACCEPTABLE_EMPTY_SLOTS:
            continue
        # Allow empty m:sub / m:sup inside m:sPre and m:sSubSup
        # (placeholders for absent script positions).
        if qname.localname in {"sub", "sup"}:
            parent = elem.getparent()
            if parent is not None:
                parent_qn = etree.QName(parent)
                if parent_qn.localname in _SLOTS_ACCEPTABLE_EMPTY_IN:
                    continue
        if not _visible_text(elem):
            empty.append(qname.localname)
    return empty


def _deg_is_hidden(elem) -> bool:
    """Check if an m:deg element has m:degHide set to 'on'."""
    parent = elem.getparent()
    if parent is None:
        return False
    for child in parent:
        if isinstance(child.tag, str) and etree.QName(child).localname == "radPr":
            for prop in child:
                if isinstance(prop.tag, str) and etree.QName(prop).localname == "degHide":
                    # OMML on/off values: on, 1, true (absent val means on)
                    val = prop.get(f"{{{OMML_NS}}}val")
                    if val is None or val.lower() in ("on", "1", "true"):
                        return True
    return False


def _visible_text(elem) -> str:
    text = "".join(text for text in elem.itertext() if text).strip()
    if text:
        return text
    for attr_name, attr_value in elem.attrib.items():
        if etree.QName(attr_name).localname == "val" and attr_value:
            return attr_value
    return ""


def _extract_mathml_tokens(root) -> list[str]:
    tokens = []
    for elem in root.iter():
        if not isinstance(elem.tag, str):
            continue
        qname = etree.QName(elem)
        if qname.localname in _MATHML_TOKEN_TAGS:
            token = _normalize_token("".join(elem.itertext()))
            # Strip structural characters that OMML represents differently
            token = _strip_structural_chars(token)
            if token:
                tokens.append(token)
    return tokens


def _extract_omml_tokens(root) -> list[str]:
    tokens = []
    for elem in root.iter():
        if not isinstance(elem.tag, str):
            continue
        qname = etree.QName(elem)
        if qname.namespace == OMML_NS and qname.localname == "t":
            for token in _split_omml_text(elem.text or ""):
                normalized = _normalize_token(token)
                if normalized:
                    tokens.append(normalized)
        elif qname.namespace == OMML_NS and qname.localname == "d":
            # Delimiters live in m:dPr/m:begChr, m:endChr and m:sepChr
            # (defaults: "(" and ")"), not in m:t text runs.
            tokens.extend(_delimiter_tokens(elem))
    return tokens


def _delimiter_tokens(d_elem) -> list[str]:
    props = {}
    for child in d_elem:
        if isinstance(child.tag, str) and etree.QName(child).localname == "dPr":
            for prop in child:
                if isinstance(prop.tag, str):
                    props[etree.QName(prop).localname] = prop.get(f"{{{OMML_NS}}}val")
    n_items = sum(
        1 for c in d_elem
        if isinstance(c.tag, str) and etree.QName(c).localname == "e"
    )
    out = []
    beg = props.get("begChr", "(")
    end = props.get("endChr", ")")
    sep = props.get("sepChr", "|")
    for ch in (beg, end):
        ch = _normalize_token(ch or "")
        if ch:
            out.append(ch)
    sep = _normalize_token(sep or "")
    if sep:
        out.extend([sep] * max(n_items - 1, 0))
    return out


def _extract_omml_text(root) -> str:
    parts = []
    for elem in root.iter():
        if not isinstance(elem.tag, str):
            continue
        qname = etree.QName(elem)
        if qname.namespace == OMML_NS and qname.localname == "t" and elem.text:
            parts.append(elem.text)
    text = _normalize_token("".join(parts))
    return _strip_structural_chars(text)


def _split_omml_text(text: str) -> list[str]:
    return re.findall(r"\d+(?:\.\d+)?|[A-Za-zΑ-Ωα-ω]+|[^\s]", text)


def _normalize_token(token: str) -> str:
    token = token.translate(_EQUIVALENT_CHARS)
    return re.sub(r"\s+", "", token)


def _strip_structural_chars(token: str) -> str:
    """Remove characters that MML2OMML.XSL represents structurally.

    Accent marks (¯, ^, ~, etc.), zero-width spaces, and certain
    spacing characters are converted to OMML structure elements
    (m:acc, m:bar, m:chr) rather than text runs, so they won't
    appear in m:t elements during token comparison.
    """
    result = token
    for ch in sorted(_STRUCTURAL_CHARS, key=len, reverse=True):
        result = result.replace(ch, "")
    return result


def _missing_tokens(
    mathml_tokens: list[str],
    omml_tokens: list[str],
    omml_text: str,
) -> list[tuple[str, int, int]]:
    expected = Counter(mathml_tokens)
    missing = []
    for token, count in expected.items():
        found = _count_token_occurrences(token, omml_tokens, omml_text)
        if found < count:
            missing.append((token, count, found))
    return missing


def _count_token_occurrences(token: str, omml_tokens: list[str], omml_text: str) -> int:
    exact_count = Counter(omml_tokens).get(token, 0)
    if not token:
        return 0
    text_count = omml_text.count(token)
    return max(exact_count, text_count)


def _missing_structures(mathml_root, omml_root) -> list[tuple[str, int, int]]:
    missing = []
    mathml_counts = Counter(
        etree.QName(elem).localname
        for elem in mathml_root.iter()
        if isinstance(elem.tag, str)
    )
    omml_counts = Counter(
        etree.QName(elem).localname
        for elem in omml_root.iter()
        if isinstance(elem.tag, str) and etree.QName(elem).namespace == OMML_NS
    )
    for mathml_tag, omml_tags in _STRUCTURE_RULES.items():
        expected = mathml_counts.get(mathml_tag, 0)
        if expected == 0:
            continue
        found = sum(omml_counts.get(tag, 0) for tag in omml_tags)
        if found < expected:
            missing.append((mathml_tag, expected, found))
    return missing


def _validate_style_semantics(mathml_root, omml_root) -> tuple[list[str], list[str]]:
    errors = []
    warnings = []
    errors.extend(_validate_mroot_degree(mathml_root, omml_root))
    errors.extend(_validate_matrix_shape(mathml_root, omml_root))
    errors.extend(_validate_delimiters(mathml_root, omml_root))
    errors.extend(_validate_accents_and_bars(mathml_root, omml_root))
    nary_errors, nary_warnings = _validate_nary_and_limits(mathml_root, omml_root)
    errors.extend(nary_errors)
    warnings.extend(nary_warnings)
    warnings.extend(_validate_functions(mathml_root, omml_root))
    warnings.extend(_validate_style_sensitive_constructs(mathml_root, omml_root))
    errors.extend(_validate_multiscripts(mathml_root, omml_root))
    return errors, warnings


def _count_mathml_tag(root, tag_name: str) -> int:
    return sum(
        1 for elem in root.iter()
        if isinstance(elem.tag, str) and _local_name(elem) == tag_name
    )


def _count_omml_tag(root, tag_name: str) -> int:
    return sum(
        1 for elem in root.iter()
        if isinstance(elem.tag, str)
        and etree.QName(elem).namespace == OMML_NS
        and etree.QName(elem).localname == tag_name
    )


def _count_non_empty_omml(root, tag_name: str) -> int:
    return sum(
        1 for elem in root.iter()
        if isinstance(elem.tag, str)
        and etree.QName(elem).namespace == OMML_NS
        and etree.QName(elem).localname == tag_name
        and _visible_text(elem)
    )


def _has_omml_any(root, tag_names: set[str]) -> bool:
    for elem in root.iter():
        if not isinstance(elem.tag, str):
            continue
        qname = etree.QName(elem)
        if qname.namespace == OMML_NS and qname.localname in tag_names:
            return True
    return False


def _mathml_text_for(elem) -> str:
    return _normalize_token("".join(elem.itertext()))


def _find_mathml_elements(root, tag_name: str) -> list:
    return [
        elem for elem in root.iter()
        if isinstance(elem.tag, str) and _local_name(elem) == tag_name
    ]


def _validate_mroot_degree(mathml_root, omml_root) -> list[str]:
    mroot_count = _count_mathml_tag(mathml_root, "mroot")
    if not mroot_count:
        return []
    deg_count = _count_non_empty_omml(omml_root, "deg")
    if deg_count < mroot_count:
        return [
            f"OMML mroot degree loss: MathML mroot count {mroot_count}, "
            f"non-empty OMML m:deg count {deg_count}"
        ]
    return []


def _validate_matrix_shape(mathml_root, omml_root) -> list[str]:
    tables = _find_mathml_elements(mathml_root, "mtable")
    if not tables:
        return []
    omml_matrix_count = _count_omml_tag(omml_root, "m")
    omml_delim_count = _count_omml_tag(omml_root, "d")
    omml_eqarr_count = _count_omml_tag(omml_root, "eqArr")
    # MML2OMML.XSL may represent matrices as:
    #   m:m (explicit matrix), m:d (bracketed/braced), or
    #   m:eqArr (equation array — piecewise functions with braces)
    total_matrix_structures = omml_matrix_count + omml_delim_count + omml_eqarr_count
    if total_matrix_structures < len(tables):
        return [f"OMML matrix loss: MathML mtable count {len(tables)}, "
                f"OMML m:m={omml_matrix_count} m:d={omml_delim_count}"]

    mathml_rows = 0
    mathml_cells = 0
    max_mathml_cells = 0
    for table in tables:
        rows = [child for child in table if isinstance(child.tag, str) and _local_name(child) in {"mtr", "mlabeledtr"}]
        mathml_rows += len(rows)
        for row in rows:
            cells = [child for child in row if isinstance(child.tag, str) and _local_name(child) == "mtd"]
            mathml_cells += len(cells)
            max_mathml_cells = max(max_mathml_cells, len(cells))

    omml_rows = [
        elem for elem in omml_root.iter()
        if isinstance(elem.tag, str)
        and etree.QName(elem).namespace == OMML_NS
        and etree.QName(elem).localname == "mr"
    ]
    # eqArr uses <m:e> directly as row equivalents, not <m:mr>
    omml_eqarr_elements = [
        elem for elem in omml_root.iter()
        if isinstance(elem.tag, str)
        and etree.QName(elem).namespace == OMML_NS
        and etree.QName(elem).localname == "e"
        and etree.QName(elem.getparent()).localname == "eqArr"
    ] if omml_eqarr_count > 0 else []

    omml_cell_count = 0
    max_omml_cells = 0
    for row in omml_rows:
        cells = [
            child for child in row
            if isinstance(child.tag, str)
            and etree.QName(child).namespace == OMML_NS
            and etree.QName(child).localname == "e"
        ]
        omml_cell_count += len(cells)
        max_omml_cells = max(max_omml_cells, len(cells))

    # eqArr rows: each <m:e> is one row
    omml_row_count = len(omml_rows) + len(omml_eqarr_elements)

    errors = []
    if omml_row_count < mathml_rows:
        errors.append(f"OMML matrix row loss: MathML rows {mathml_rows}, OMML rows {omml_row_count}")
    if omml_cell_count < mathml_cells and not omml_eqarr_elements:
        # Only check cell count for non-eqArr matrices (eqArr cell structure is different)
        errors.append(f"OMML matrix cell loss: MathML cells {mathml_cells}, OMML cells {omml_cell_count}")
    return errors


def _validate_delimiters(mathml_root, omml_root) -> list[str]:
    fenced_count = _count_mathml_tag(mathml_root, "mfenced")
    if fenced_count and _count_omml_tag(omml_root, "d") < fenced_count:
        return [f"OMML delimiter loss: MathML mfenced count {fenced_count}, OMML m:d count {_count_omml_tag(omml_root, 'd')}"]
    return []


def _validate_accents_and_bars(mathml_root, omml_root) -> list[str]:
    errors = []
    for mover in _find_mathml_elements(mathml_root, "mover"):
        children = [child for child in mover if isinstance(child.tag, str)]
        if len(children) >= 2 and _mathml_text_for(children[1]) in _ACCENT_MARKS:
            if not _has_omml_any(omml_root, {"acc", "bar", "groupChr"}):
                errors.append("OMML accent/bar loss: MathML mover accent present, OMML lacks m:acc/m:bar/m:groupChr")
                break
    return errors


def _validate_nary_and_limits(mathml_root, omml_root) -> tuple[list[str], list[str]]:
    errors = []
    warnings = []
    has_nary_token = any(token in _NARY_OPERATORS for token in _extract_mathml_tokens(mathml_root))
    has_nary_limits = False
    for elem in list(_find_mathml_elements(mathml_root, "munder")) + list(_find_mathml_elements(mathml_root, "mover")) + list(_find_mathml_elements(mathml_root, "munderover")):
        children = [child for child in elem if isinstance(child.tag, str)]
        if children and _mathml_text_for(children[0]) in _NARY_OPERATORS:
            has_nary_limits = True
            break
    has_omml_nary = _has_omml_any(omml_root, {"nary", "limLow", "limUpp"})
    if has_nary_limits and not has_omml_nary:
        errors.append("OMML n-ary limit loss: MathML n-ary limits present, OMML lacks m:nary/m:limLow/m:limUpp")
    elif has_nary_token and not has_omml_nary:
        warnings.append("OMML n-ary style warning: n-ary operator preserved as text without m:nary structure")
    return errors, warnings


def _validate_functions(mathml_root, omml_root) -> list[str]:
    function_tokens = [token for token in _extract_mathml_tokens(mathml_root) if token in _FUNCTION_NAMES]
    if function_tokens and not _has_omml_any(omml_root, {"func", "fName"}):
        return ["OMML function style warning: MathML function names present without m:func/m:fName structure"]
    return []


def _validate_style_sensitive_constructs(mathml_root, omml_root) -> list[str]:
    warnings = []
    if _count_mathml_tag(mathml_root, "mphantom"):
        warnings.append("OMML style warning: MathML mphantom present; verify phantom spacing")
    if _count_mathml_tag(mathml_root, "mspace") or _count_mathml_tag(mathml_root, "mpadded"):
        warnings.append("OMML style warning: MathML spacing/padding present; verify visual spacing")
    if _count_mathml_tag(mathml_root, "maligngroup") or _count_mathml_tag(mathml_root, "malignmark"):
        warnings.append("OMML style warning: MathML alignment markers present; verify equation alignment")
    return warnings


def _validate_multiscripts(mathml_root, omml_root) -> list[str]:
    errors = []
    if _count_mathml_tag(mathml_root, "mmultiscripts") and not _has_omml_any(omml_root, {"sPre", "sSubSup", "sSub", "sSup"}):
        errors.append("OMML multiscript loss: MathML mmultiscripts present, OMML lacks script structure")
    if _count_mathml_tag(mathml_root, "mprescripts") and not _has_omml_any(omml_root, {"sPre"}):
        errors.append("OMML prescript loss: MathML mprescripts present, OMML lacks m:sPre")
    return errors


def validate_output_docx(docx_path: str) -> dict:
    """Validate the output docx for structural integrity.

    Returns a dict with keys: valid, errors, warnings.
    """
    result = {"valid": True, "errors": [], "warnings": []}

    path = Path(docx_path)
    if not path.exists():
        result["valid"] = False
        result["errors"].append(f"Output file not found: {docx_path}")
        return result

    # Check it's a valid ZIP
    try:
        with zipfile.ZipFile(docx_path, "r") as z:
            namelist = z.namelist()
    except zipfile.BadZipFile:
        result["valid"] = False
        result["errors"].append("Output file is not a valid ZIP")
        return result

    # Check required entries
    for entry in REQUIRED_DOCX_ENTRIES:
        if entry not in namelist:
            result["valid"] = False
            result["errors"].append(f"Missing required entry: {entry}")

    # Validate XML well-formedness of key parts
    xml_parts = [
        "word/document.xml",
        "[Content_Types].xml",
    ]
    with zipfile.ZipFile(docx_path, "r") as z:
        for part in xml_parts:
            if part not in namelist:
                continue
            try:
                etree.fromstring(z.read(part))
            except etree.XMLSyntaxError as e:
                result["valid"] = False
                result["errors"].append(f"XML parse error in {part}: {e}")

    return result


def validate_formula_count(report: ConversionReport) -> dict:
    """Check formula count conservation: input = output + failed + skipped."""
    total = report.total_ole_objects
    accounted = report.succeeded + report.failed + report.skipped

    return {
        "conserved": total == accounted,
        "total_detected": total,
        "succeeded": report.succeeded,
        "failed": report.failed,
        "skipped": report.skipped,
        "accounted": accounted,
        "loss": total - accounted,
    }


def make_report(
    input_docx: str,
    output_docx: str,
    formulas: list[FormulaInfo],
) -> ConversionReport:
    """Create a ConversionReport from a completed pipeline run."""
    report = ConversionReport(
        input_docx=input_docx,
        output_docx=output_docx,
        total_ole_objects=len(formulas),
        formulas=formulas,
    )
    return report


def try_word_open(docx_path: str) -> dict:
    """Attempt to open the docx with Word COM to verify it's not corrupt.

    Returns dict with 'opened' (bool) and 'error' (optional str).
    Requires Word COM to be available on Windows.
    """
    result = {"opened": False, "error": None}
    try:
        import win32com.client
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        try:
            doc = word.Documents.Open(str(Path(docx_path).resolve()), ReadOnly=True)
            doc.Close(False)
            result["opened"] = True
        finally:
            word.Quit()
    except ImportError:
        result["error"] = "pywin32 not installed; Word COM validation skipped"
    except Exception as e:
        result["error"] = str(e)
    return result


def health_check() -> dict:
    """Check that all external dependencies for the pipeline are available.

    Returns a dict with status of each dependency.
    """
    checks = {}

    # Python packages
    for pkg in ["lxml", "olefile"]:
        try:
            __import__(pkg)
            checks[f"python-{pkg}"] = "ok"
        except ImportError:
            checks[f"python-{pkg}"] = "missing"

    return checks
