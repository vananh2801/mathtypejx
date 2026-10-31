"""Converter: MTEF → MathML → OMML conversion chain."""

import re
from pathlib import Path
from typing import Optional

from lxml import etree

from .models import FormulaInfo, FormulaStatus, RiskLevel

# MTEF parsing is handled by the native Python parser in mathtypejx.mtef.

# Common XSL paths
_DEFAULT_XSL_PATHS = [
    Path(__file__).resolve().parent.parent.parent.parent / "MML2OMML.XSL",
    Path("./MML2OMML.XSL"),
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
    math_font: Optional[str] = None,
) -> bool:
    """Run the full conversion chain for a single formula.

    Args:
        formula: FormulaInfo with ole_data populated.
        xsl_path: Path to MML2OMML.XSL. Auto-detected if None.
    Returns:
        True if conversion succeeded.
    """
    if not formula.ole_data and not (
        formula.source == "picture" and formula.mtef_bytes
    ):
        formula.status = FormulaStatus.FAILED
        formula.error_message = "No OLE data to convert"
        return False

    # Step 1: MTEF → MathML (pure Python)
    if formula.source == "picture":
        from .mtef import mtef_data_to_mathml
        mathml = mtef_data_to_mathml(formula.mtef_bytes)
    else:
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

    if math_font:
        formula.omml = apply_math_font(formula.omml, math_font)

    formula.status = FormulaStatus.CONVERTED
    return True


def apply_math_font(omml: str, font: str) -> str:
    """Set the font (w:rFonts) of every math run, e.g. "Cambria Math"."""
    m_ns = "http://schemas.openxmlformats.org/officeDocument/2006/math"
    w_ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    try:
        root = etree.fromstring(omml.encode("utf-8"))
    except etree.XMLSyntaxError:
        return omml
    for run in root.iter(f"{{{m_ns}}}r"):
        w_rpr = run.find(f"{{{w_ns}}}rPr")
        if w_rpr is None:
            w_rpr = etree.Element(f"{{{w_ns}}}rPr")
            m_rpr = run.find(f"{{{m_ns}}}rPr")
            run.insert(0 if m_rpr is None else run.index(m_rpr) + 1, w_rpr)
        fonts = w_rpr.find(f"{{{w_ns}}}rFonts")
        if fonts is None:
            fonts = etree.Element(f"{{{w_ns}}}rFonts")
            w_rpr.insert(0, fonts)
        for attr in ("ascii", "hAnsi", "cs"):
            fonts.set(f"{{{w_ns}}}{attr}", font)
    return etree.tostring(root, encoding="unicode")


def _ole_to_mathml(
    ole_data: Optional[bytes],
    ole_name: str,
) -> Optional[str]:
    """Convert OLE binary to MathML using pure Python MTEF parser."""
    if not ole_data:
        return None

    from .mtef import mtef_to_mathml
    return mtef_to_mathml(ole_data)

# Danh sách các hàm toán học phổ biến (sắp xếp từ dài đến ngắn để ưu tiên khớp trước)
_MATH_FUNCTIONS = sorted([
    "arcsin", "arccos", "arctan", "arccot",
    "sinh", "cosh", "tanh", "coth",
    "sin", "cos", "tan", "cot", "sec", "csc",
    "log", "ln", "lg", "exp",
    "lim", "max", "min", "det", "dim", "deg", "gcd",
    "tg", "cotg",
], key=len, reverse=True)

def _normalize_functions_in_parent(parent: etree._Element) -> None:
    """Gộp các thẻ <mi>s</mi><mi>i</mi><mi>n</mi> đứng liền nhau thành <mi mathvariant="normal">sin</mi>
    và tự động chèn khoảng cách toán học mỏng (\\u2009) nếu phía sau là số/biến (vd: sin 3x).
    """
    i = 0
    while i < len(parent):
        matched_fn = None
        matched_count = 0

        # Thử khớp các hàm toán học bắt đầu từ vị trí con thứ i
        for fn in _MATH_FUNCTIONS:
            fn_len = len(fn)
            current_text = ""
            count = 0
            while i + count < len(parent) and len(current_text) < fn_len:
                ch = parent[i + count]
                if not isinstance(ch.tag, str):
                    break
                loc = _local_name(ch)
                if loc not in ("mi", "mtext"):
                    break
                t = ch.text or ""
                if not t:
                    break
                current_text += t
                count += 1

            if current_text.lower() == fn:
                matched_fn = fn
                matched_count = count
                break

        if matched_fn and matched_count > 0:
            # Kiểm tra xem có cần chèn khoảng cách mỏng (\u2009) phía sau tên hàm không
            add_space = False
            next_idx = i + matched_count
            if next_idx < len(parent):
                next_elem = parent[next_idx]
                if isinstance(next_elem.tag, str):
                    next_loc = _local_name(next_elem)
                    next_text = (next_elem.text or "").strip()
                    # Không chèn khoảng cách nếu ngay sau là dấu ngoặc mở hoặc đã có dấu cách
                    if next_loc == "mo" and next_text in ("(", "[", "{", "|", "（", "【"):
                        add_space = False
                    elif next_text.startswith(("\u00a0", "\u2009", " ")):
                        add_space = False
                    elif next_loc in ("mn", "mi", "mfrac", "msqrt", "mroot"):
                        # Ví dụ: sin đi liền với 3x, 2x, x -> cần cách ra: sin 3x
                        add_space = True
                    elif next_loc == "mo" and next_text and next_text not in (",", ";", ".", ")", "]", "}", "）", "】"):
                        add_space = True

            # Tạo thẻ <mi> chuẩn toán học: thẳng đứng (mathvariant="normal")
            fn_elem = etree.Element(f"{{{MATHML_NS}}}mi")
            fn_elem.set("mathvariant", "normal")
            fn_elem.text = matched_fn + ("\u2009" if add_space else "")

            # Giữ lại tail text nếu có
            fn_elem.tail = parent[i + matched_count - 1].tail

            # Xóa các thẻ cũ bị tách lẻ và chèn thẻ hàm mới vào
            for _ in range(matched_count):
                parent.remove(parent[i])
            parent.insert(i, fn_elem)
            i += 1
        else:
            i += 1


def _normalize_functions(elem: etree._Element) -> None:
    """Duyệt đệ quy toàn bộ cây MathML để chuẩn hóa tất cả các hàm toán học."""
    for child in list(elem):
        _normalize_functions(child)
    if len(elem) > 0:
        _normalize_functions_in_parent(elem)

def _normalize_function_scripts_in_parent(parent: etree._Element) -> None:
    """Gộp các tiền tố bị bỏ rơi vào cơ số của số mũ/chỉ số (vd: s + i + n^6 -> (sin)^6)."""
    i = 0
    while i < len(parent):
        child = parent[i]
        if not isinstance(child.tag, str):
            i += 1
            continue

        loc = _local_name(child)
        if loc in ("msup", "msub", "msubsup"):
            if len(child) > 0:
                base_elem = child[0]
                base_text = "".join(base_elem.itertext()).strip()

                # Nếu bản thân base_elem đã là hàm hoàn chỉnh (vd: <msup><mi>sin</mi><mn>2</mn></msup>)
                if base_text.lower() in _MATH_FUNCTIONS:
                    if _local_name(base_elem) == "mi":
                        base_elem.set("mathvariant", "normal")
                    _add_script_trailing_space(parent, i, child)
                    i += 1
                    continue

                # Tìm các ký tự đứng liền trước child (vd: 's', 'i' đứng trước 'n^6')
                prefix_items = []
                j = i - 1
                while j >= 0:
                    prev = parent[j]
                    if not isinstance(prev.tag, str):
                        break
                    prev_loc = _local_name(prev)
                    if prev_loc not in ("mi", "mtext"):
                        break
                    txt = prev.text or ""
                    if not txt or not txt.isalpha():
                        break
                    prefix_items.insert(0, (j, txt))
                    if sum(len(t) for _, t in prefix_items) >= 6:
                        break
                    j -= 1

                # Thử ghép các tiền tố với ký tự cơ số xem có tạo thành tên hàm không
                matched_fn = None
                matched_prefix_count = 0
                for start_k in range(len(prefix_items)):
                    candidate = "".join(t for _, t in prefix_items[start_k:]) + base_text
                    if candidate.lower() in _MATH_FUNCTIONS:
                        matched_fn = candidate.lower()
                        matched_prefix_count = len(prefix_items) - start_k
                        break

                if matched_fn:
                    # Xóa các ký tự tiền tố thừa đứng trước (xóa 's', 'i')
                    del_start = i - matched_prefix_count
                    for _ in range(matched_prefix_count):
                        parent.remove(parent[del_start])

                    child_idx = del_start

                    # Thay thế cơ số cũ ('n') bằng cả hàm chuẩn: <mi mathvariant="normal">sin</mi>
                    new_base = etree.Element(f"{{{MATHML_NS}}}mi")
                    new_base.set("mathvariant", "normal")
                    new_base.text = matched_fn
                    new_base.tail = base_elem.tail
                    child.replace(base_elem, new_base)

                    # Tự động chèn khoảng cách mỏng sau số mũ nếu phía sau là biến (vd: sin^6 x)
                    _add_script_trailing_space(parent, child_idx, child)

                    i = child_idx + 1
                    continue

        i += 1


def _add_script_trailing_space(parent: etree._Element, idx: int, elem: etree._Element) -> None:
    """Tự động thêm khoảng cách mỏng (\\u2009) sau số mũ nếu liền sau là biến/số (vd: sin^6 x)."""
    next_idx = idx + 1
    if next_idx < len(parent):
        next_elem = parent[next_idx]
        if isinstance(next_elem.tag, str):
            next_loc = _local_name(next_elem)
            next_txt = (next_elem.text or "").strip()
            # Không cách nếu phía sau là dấu ngoặc đóng hoặc dấu chấm/phẩy
            if next_loc == "mo" and next_txt in (")", "]", "}", ",", ";", ".", "）", "】"):
                return
            if next_txt.startswith(("\u00a0", "\u2009", " ")):
                return
            if next_loc in ("mn", "mi", "mfrac", "msqrt", "mroot") or (
                next_loc == "mo" and next_txt not in ("(", "[", "{", "|")
            ):
                if not (elem.tail or "").startswith(("\u00a0", "\u2009", " ")):
                    elem.tail = "\u2009" + (elem.tail or "")


def _normalize_function_scripts(elem: etree._Element) -> None:
    """Duyệt đệ quy toàn bộ cây để sửa tất cả các hàm có số mũ / chỉ số."""
    for child in list(elem):
        _normalize_function_scripts(child)
    if len(elem) > 0:
        _normalize_function_scripts_in_parent(elem)

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
    _normalize_empty_scripts(root)
    # Replace <mtext> → <mi mathvariant="normal"> so MML2OMML.XSL
    # produces <m:sty m:val="p"/> instead of <m:nor/>.  The latter
    # switches to the document body font whose baseline may differ
    # from the math font, making units look like subscripts.
    _mtext_to_normal_mi(root)

    _normalize_function_scripts(root)
    _normalize_functions(root)

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


def _normalize_empty_scripts(root) -> None:
    """Replace msub/msup/msubsup whose scripts are all empty by base + space.

    MathType encodes the spaces inside text-like runs (e.g. "khi n le") as
    sub/sup slots with no content; emitting them as real scripts produces
    empty m:sSubSup nodes and glues the words together.
    """
    # <mmultiscripts> with an empty base is a MathType artefact for a script
    # nested inside another script (e.g. x^{ {}^{2} }).  Flatten it to the
    # single real script it carries; drop it when nothing is inside.
    ms_tag = f"{{{MATHML_NS}}}mmultiscripts"
    none_tag = f"{{{MATHML_NS}}}none"
    pre_tag = f"{{{MATHML_NS}}}mprescripts"
    for ms in list(root.iter(ms_tag)):
        kids = [c for c in ms if isinstance(c.tag, str)]
        if not kids or "".join(kids[0].itertext()).strip():
            continue  # real base: leave as is
        scripts = [
            c for c in kids[1:]
            if c.tag not in (none_tag, pre_tag) and "".join(c.itertext()).strip()
        ]
        if len(scripts) > 1:
            continue
        if scripts:
            replacement = scripts[0]
            ms.remove(replacement)
        else:
            replacement = etree.Element(f"{{{MATHML_NS}}}mrow")
        replacement.tail = ms.tail
        parent = ms.getparent()
        if parent is not None:
            parent.replace(ms, replacement)

    script_tags = {f"{{{MATHML_NS}}}{t}" for t in ("msub", "msup", "msubsup")}
    for elem in list(root.iter()):
        if elem.tag not in script_tags:
            continue
        kids = [c for c in elem if isinstance(c.tag, str)]
        if len(kids) < 2:
            continue
        if any("".join(c.itertext()).strip() for c in kids[1:]):
            continue
        parent = elem.getparent()
        if parent is None:
            continue
        idx = parent.index(elem)
        base = kids[0]
        space = etree.Element(f"{{{MATHML_NS}}}mtext")
        space.text = "\u00a0"
        space.tail = elem.tail
        parent.remove(elem)
        parent.insert(idx, base)
        parent.insert(idx + 1, space)


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
        # Standard MathML order is (base, index): an empty index means a
        # plain square root.
        if len(children) == 2 and (
            len(children[1]) == 0
            and not (children[1].text or "").strip()
        ):
            msqrt = etree.Element(f"{{{MATHML_NS}}}msqrt")
            mroot.remove(children[0])
            msqrt.append(children[0])
            msqrt.tail = mroot.tail
            parent = mroot.getparent()
            if parent is not None:
                idx = list(parent).index(mroot)
                parent.remove(mroot)
                parent.insert(idx, msqrt)
            continue
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

# Biến cache toàn cục lưu transformer XSLT
_CACHED_XSLT_TRANSFORM = None
_CACHED_XSLT_PATH = None

def _mathml_to_omml(mathml: str, xsl_path: Optional[str] = None) -> Optional[str]:
    global _CACHED_XSLT_TRANSFORM, _CACHED_XSLT_PATH
    """Convert MathML to OMML.

    Uses Microsoft's MML2OMML.XSL when it is available (best fidelity).
    Otherwise falls back to the pure-Python ``mathml2omml`` package, which
    makes the pipeline usable on Linux/macOS without Microsoft Office.
    """

    xsl_file = _find_xsl(xsl_path)
    if xsl_file is None:
        return _mathml_to_omml_python(mathml)

    try:
        # Tái sử dụng đối tượng XSLT đã biên dịch để không phải parse lại file đĩa
        if _CACHED_XSLT_TRANSFORM is None or _CACHED_XSLT_PATH != str(xsl_file):
            xsl_doc = etree.parse(str(xsl_file))
            _CACHED_XSLT_TRANSFORM = etree.XSLT(xsl_doc)
            _CACHED_XSLT_PATH = str(xsl_file)

        mathml_doc = etree.fromstring(mathml.encode("utf-8"))
        result = _CACHED_XSLT_TRANSFORM(mathml_doc)
        omml = etree.tostring(result, encoding="unicode", pretty_print=False)
        return omml
    except Exception:
        # Nếu XSLT gặp lỗi bất thường thì tự động fallback về mathml2omml
        return _mathml_to_omml_python(mathml)


_BAR_CHARS = {"¯", "‾", "―", "_", "−", "-", "\u0305", "\u203e"}
_MARK_PREFIX = "ACCMARK"
_OMML_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"


def _extract_accents(root) -> dict:
    """Swap accent movers/unders for markers (mathml2omml mishandles them).

    ``<mover accent><base/><mo>^</mo></mover>`` becomes
    ``<mrow><mi>ACCMARKn</mi><mrow>base</mrow></mrow>``; _apply_accents()
    later turns the marker back into a real <m:acc> or <m:bar>.
    """
    marks: dict[str, tuple[str, str, str]] = {}
    for elem in list(root.iter("mover", "munder")):
        kids = [c for c in elem if isinstance(c.tag, str)]
        if len(kids) != 2 or kids[1].tag != "mo":
            continue
        attr = elem.get("accent") if elem.tag == "mover" else elem.get("accentunder")
        char = "".join(kids[1].itertext()).strip()
        if attr != "true" or len(char) != 1:
            continue
        pos = "top" if elem.tag == "mover" else "bot"
        kind = "bar" if char in _BAR_CHARS else "acc"
        if char == "^":
            char = "\u0302"  # combining circumflex (Word's default accent)
        key = f"{_MARK_PREFIX}{len(marks)}X"
        marks[key] = (kind, char, pos)

        marker = etree.Element("mi")
        marker.text = key
        base = etree.Element("mrow")
        base.append(kids[0])
        wrapper = etree.Element("mrow")
        wrapper.append(marker)
        wrapper.append(base)
        wrapper.tail = elem.tail
        parent = elem.getparent()
        if parent is None:
            continue
        parent.replace(elem, wrapper)
    return marks


def _apply_accents(omml: str, marks: dict):
    """Turn ACCMARK runs produced above into <m:acc>/<m:bar> structures."""
    ns = _OMML_NS
    q = lambda t: f"{{{ns}}}{t}"
    try:
        root = etree.fromstring(f'<root xmlns:m="{ns}">{omml}</root>'.encode("utf-8"))
    except etree.XMLSyntaxError:
        return None

    for _ in range(len(marks) + 1):
        found = None
        for t in root.iter(q("t")):
            if (t.text or "").startswith(_MARK_PREFIX) and t.text in marks:
                found = t
                break
        if found is None:
            break
        run = found.getparent()
        holder = run.getparent()          # <m:e> of the wrapping box
        kind, char, pos = marks[found.text]
        base_nodes = [c for c in holder if c is not run]
        box = holder.getparent() if holder is not None else None
        if kind == "bar":
            node = etree.Element(q("bar"))
            pr = etree.SubElement(node, q("barPr"))
            etree.SubElement(pr, q("pos")).set(q("val"), pos)
        else:
            node = etree.Element(q("acc"))
            pr = etree.SubElement(node, q("accPr"))
            etree.SubElement(pr, q("chr")).set(q("val"), char)
        e = etree.SubElement(node, q("e"))
        for c in base_nodes:
            e.append(c)
        if box is not None and box.tag == q("box") and box.getparent() is not None:
            box.getparent().replace(box, node)
        else:
            holder.remove(run)
            holder.append(node)
    return "".join(etree.tostring(c, encoding="unicode") for c in root)


_OPENERS = {"(": ")", "[": "]", "{": "}", "|": "|", "⟨": "⟩", "⌊": "⌋", "⌈": "⌉"}
_CLOSERS = {")", "]", "}", "⟩", "⌋", "⌉", "|"}


def _wrap_trailing_closers(root) -> None:
    """Keep a lone closing bracket that follows a table/expression.

    MathType "system" brackets such as ``a∥c / b∥c / a≠b }`` give
    ``<mrow><mtable/><mo>}</mo></mrow>``; mathml2omml silently drops the
    unmatched ``}``.  Wrap the preceding siblings in an <mfenced> with only
    a closing character so the bracket survives as an OMML delimiter.
    """
    for mrow in list(root.iter("mrow")):
        kids = [c for c in mrow if isinstance(c.tag, str)]
        if len(kids) < 2:
            continue
        last = kids[-1]
        text = (last.text or "").strip()
        if last.tag != "mo" or text not in _CLOSERS:
            continue
        # Skip when an opener is present among the siblings (matched pair).
        if any(k.tag == "mo" and (k.text or "").strip() in _OPENERS
               and _OPENERS[(k.text or "").strip()] == text for k in kids[:-1]):
            continue
        fenced = etree.Element("mfenced")
        fenced.set("open", "")
        fenced.set("close", text)
        fenced.set("separators", "")
        for k in kids[:-1]:
            mrow.remove(k)
            fenced.append(k)
        mrow.remove(last)
        mrow.insert(0, fenced)


def _unwrap_plain_boxes(omml: str) -> str:
    """Remove the redundant <m:box> wrappers mathml2omml emits for <mrow>.

    A box without properties is a pure grouping node; it only clutters the
    equation tree and makes editing harder in Word/WPS.
    """
    ns = "http://schemas.openxmlformats.org/officeDocument/2006/math"
    wrapped = f'<root xmlns:m="{ns}">{omml}</root>'
    try:
        root = etree.fromstring(wrapped.encode("utf-8"))
    except etree.XMLSyntaxError:
        return omml
    box = f"{{{ns}}}box"
    e_tag = f"{{{ns}}}e"
    box_pr = f"{{{ns}}}boxPr"
    changed = True
    while changed:
        changed = False
        for b in list(root.iter(box)):
            kids = [c for c in b if isinstance(c.tag, str)]
            if any(c.tag == box_pr for c in kids):
                continue
            inner = [c for c in kids if c.tag == e_tag]
            if len(inner) != 1:
                continue
            parent = b.getparent()
            if parent is None:
                continue
            idx = parent.index(b)
            for offset, child in enumerate(list(inner[0])):
                parent.insert(idx + offset, child)
            parent.remove(b)
            changed = True
            break
    # <m:rad> must contain <m:deg>; for plain square roots hide the degree.
    rad = f"{{{ns}}}rad"
    rad_pr = f"{{{ns}}}radPr"
    deg_tag = f"{{{ns}}}deg"
    deg_hide = f"{{{ns}}}degHide"
    val_attr = f"{{{ns}}}val"
    for r in root.iter(rad):
        if r.find(deg_tag) is not None:
            continue
        pr = r.find(rad_pr)
        if pr is None:
            pr = etree.Element(rad_pr)
            r.insert(0, pr)
        if pr.find(deg_hide) is None:
            hide = etree.SubElement(pr, deg_hide)
            hide.set(val_attr, "1")
        pos = list(r).index(pr) + 1
        r.insert(pos, etree.Element(deg_tag))
    out = etree.tostring(root[0], encoding="unicode") if len(root) else omml
    return out


def _mathml_to_omml_python(mathml: str) -> Optional[str]:
    """Pure-Python MathML -> OMML fallback (no Microsoft Office needed)."""
    try:
        import mathml2omml
    except ImportError:
        return None

    try:
        root = etree.fromstring(mathml.encode("utf-8"))
        # mathml2omml only understands un-prefixed element names.
        for elem in root.iter():
            if isinstance(elem.tag, str):
                elem.tag = etree.QName(elem).localname
        _wrap_trailing_closers(root)
        marks = _extract_accents(root)
        etree.cleanup_namespaces(root)
        clean = etree.tostring(root, encoding="unicode")
        omml = mathml2omml.convert(clean)
        if not omml:
            return None
        if marks:
            omml = _apply_accents(omml, marks)
            if omml is None:
                return None
        omml = _unwrap_plain_boxes(omml)
        # Declare the OMML namespace so the fragment parses standalone.
        omml = omml.strip()
        if "xmlns:m=" not in omml.split(">", 1)[0]:
            omml = omml.replace(
                "<m:oMath",
                '<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"',
                1,
            )
        return omml
    except Exception as e:
        # Thêm dòng này để in ra lỗi và chuỗi MathML gây lỗi:
        print(f"[DEBUG mathml2omml error]: {e}\nMathML: {mathml}")
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
