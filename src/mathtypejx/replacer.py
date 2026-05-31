"""Replacer: replace MathType OLE objects in docx XML with OMML math elements."""

import shutil
import zipfile
from pathlib import Path
from typing import Optional

from lxml import etree

from .models import FormulaInfo, FormulaStatus, RiskLevel

# Namespaces used in replacement
NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS_O = "urn:schemas-microsoft-com:office:office"
NS_V = "urn:schemas-microsoft-com:vml"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_M = "http://schemas.openxmlformats.org/officeDocument/2006/math"

NSMAP = {
    "w": NS_W,
    "o": NS_O,
    "v": NS_V,
    "r": NS_R,
    "m": NS_M,
}

# Content type to add for OMML
MATH_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.mathml.equation+xml"

# Override types to add to [Content_Types].xml
CT_TYPES_ADDITIONS = {
    "/word/document.xml": [
        MATH_CONTENT_TYPE,
    ],
}


def replace_formulas(
    workdir: Path,
    formulas: list[FormulaInfo],
    output_docx: str,
) -> list[FormulaInfo]:
    """Replace OLE formulas with OMML in all XML parts and repack the docx.

    Only formulas with status=CONVERTED are replaced.
    Failed formulas keep their original OLE objects.

    Args:
        workdir: Path to the unpacked docx working directory.
        formulas: List of FormulaInfo objects from the pipeline.
        output_docx: Path for the output .docx file.

    Returns:
        The updated list of formulas (with status updated).
    """
    # Group formulas by part_name
    by_part: dict[str, list[FormulaInfo]] = {}
    for f in formulas:
        if f.status == FormulaStatus.CONVERTED:
            by_part.setdefault(f.part_name, []).append(f)

    # Process each part
    for part_name, part_formulas in by_part.items():
        part_path = workdir / part_name
        if not part_path.exists():
            for f in part_formulas:
                f.status = FormulaStatus.FAILED
                f.error_message = f"Part file not found: {part_name}"
            continue

        _replace_in_part(part_path, part_formulas)

    # Update [Content_Types].xml
    _update_content_types(workdir)

    # Repack the docx
    _repack_docx(workdir, Path(output_docx))

    return formulas


def _replace_in_part(part_path: Path, formulas: list[FormulaInfo]) -> None:
    """Replace OLE objects in a single XML part."""
    parser = etree.XMLParser(remove_blank_text=False)
    tree = etree.parse(str(part_path), parser)
    root = tree.getroot()

    # Build rid → FormulaInfo map (only for CONVERTED formulas)
    rid_to_formula: dict[str, FormulaInfo] = {}
    for f in formulas:
        if f.status == FormulaStatus.CONVERTED:
            rid_to_formula[f.relationship_id] = f

    if not rid_to_formula:
        return

    # Group OLEObjects by their parent w:r.
    # Multiple formulas can share the same <w:r>; replacing one run
    # breaks the parent chain for all other OLEObjects in that run.
    # We collect all OLEs per run and replace the run once.
    run_oles: dict[int, list[tuple[object, FormulaInfo]]] = {}
    run_map: dict[int, object] = {}  # id(run) → run element
    for ole_elem in root.iter():
        if ole_elem.tag != f"{{{NS_O}}}OLEObject":
            continue
        rid = ole_elem.get(f"{{{NS_R}}}id", "")
        formula = rid_to_formula.get(rid)
        if formula is None:
            continue

        # Walk up to find w:object or w:pict
        obj_elem = ole_elem
        while obj_elem is not None and (
            obj_elem.tag != f"{{{NS_W}}}object"
            and obj_elem.tag != f"{{{NS_W}}}pict"
        ):
            obj_elem = obj_elem.getparent()

        if obj_elem is None:
            continue

        # Find the w:r
        run_elem = obj_elem
        while run_elem is not None and run_elem.tag != f"{{{NS_W}}}r":
            run_elem = run_elem.getparent()

        if run_elem is None:
            continue

        run_id = id(run_elem)
        if run_id not in run_oles:
            run_oles[run_id] = []
            run_map[run_id] = run_elem
        run_oles[run_id].append((ole_elem, formula))

    changed = False
    for run_id, entries in run_oles.items():
        run_elem = run_map[run_id]
        # Collect all OMML nodes from all formulas in this run
        omml_fragments = []
        for ole_elem, formula in entries:
            try:
                omml_xml = formula.omml.strip()
                if omml_xml.startswith("<?xml"):
                    idx = omml_xml.find("?>")
                    if idx >= 0:
                        omml_xml = omml_xml[idx + 2:].strip()
                omml_node = etree.fromstring(omml_xml.encode("utf-8"))
                omml_fragments.append(omml_node)
                formula.status = FormulaStatus.REPLACED
            except Exception as e:
                formula.status = FormulaStatus.FAILED
                formula.error_message = f"OMML parse error: {e}"

        if not omml_fragments:
            continue

        # Replace the entire w:r with all OMML fragments
        parent = run_elem.getparent()
        if parent is None:
            for _, formula in entries:
                if formula.status != FormulaStatus.REPLACED:
                    formula.status = FormulaStatus.FAILED
                    formula.error_message = "Parent run has no parent"
            continue

        idx = parent.index(run_elem)
        parent.remove(run_elem)

        for i, omml_node in enumerate(omml_fragments):
            parent.insert(idx + i, omml_node)
        changed = True

    if changed:
        tree.write(
            str(part_path),
            encoding="utf-8",
            xml_declaration=False if root.tag.startswith("{") else True,
            standalone=None,
        )


def _do_replace(ole_elem, formula: FormulaInfo) -> bool:
    """Replace a single OLE object run with an OMML element.

    Walks up from o:OLEObject to find w:object or w:pict → w:r,
    then replaces the entire w:r with the OMML content.

    Some documents embed MathType using VML (w:pict) instead of the
    standard OLE container (w:object). Both paths are supported.
    """
    # Walk up to find the <w:object> or <w:pict> parent
    obj_elem = ole_elem
    while obj_elem is not None and (
        obj_elem.tag != f"{{{NS_W}}}object"
        and obj_elem.tag != f"{{{NS_W}}}pict"
    ):
        obj_elem = obj_elem.getparent()

    if obj_elem is None:
        return False

    # Walk up to find the <w:r> (run) that contains this object
    run_elem = obj_elem
    while run_elem is not None and run_elem.tag != f"{{{NS_W}}}r":
        run_elem = run_elem.getparent()

    if run_elem is None:
        return False

    # Parse the OMML string
    try:
        # OMML from MML2OMML.XSL may be a full document or fragment
        omml_xml = formula.omml.strip()
        # Remove XML declaration if present
        if omml_xml.startswith("<?xml"):
            idx = omml_xml.find("?>")
            if idx >= 0:
                omml_xml = omml_xml[idx + 2:].strip()

        omml_nodes = etree.fromstring(omml_xml.encode("utf-8"))

        # If the root is <m:oMathPara>, insert it directly
        # If it's <m:oMath>, wrap in <w:r> for inline, or
        # insert directly for display
        if omml_nodes.tag == f"{{{NS_M}}}oMathPara":
            # Display formula — insert at paragraph level
            para_elem = run_elem
            while para_elem is not None and para_elem.tag != f"{{{NS_W}}}p":
                para_elem = para_elem.getparent()

            if para_elem is not None:
                parent = para_elem.getparent()
                if parent is not None:
                    idx = parent.index(para_elem)
                    parent.remove(para_elem)
                    parent.insert(idx, omml_nodes)
                    return True
            # Fallback: replace the run
            parent = run_elem.getparent()
            idx = parent.index(run_elem)
            parent.remove(run_elem)
            parent.insert(idx, omml_nodes)
            return True

        elif omml_nodes.tag == f"{{{NS_M}}}oMath":
            # Inline formula — wrap in a <w:r> to maintain paragraph structure
            parent = run_elem.getparent()
            if parent is not None:
                idx = parent.index(run_elem)
                parent.remove(run_elem)
                parent.insert(idx, omml_nodes)
                return True
            # parent is None — tree modification already removed this run

        else:
            # Unknown OMML root element, insert as-is
            parent = run_elem.getparent()
            if parent is not None:
                idx = parent.index(run_elem)
                parent.remove(run_elem)
                parent.insert(idx, omml_nodes)
                return True

    except etree.XMLSyntaxError as e:
        formula.error_message = f"Invalid OMML XML: {e}"
        return False

    return False


def _update_content_types(workdir: Path) -> None:
    """Ensure [Content_Types].xml includes OMML math content types."""
    ct_path = workdir / "[Content_Types].xml"
    if not ct_path.exists():
        return

    tree = etree.parse(str(ct_path))
    root = tree.getroot()
    ns = "http://schemas.openxmlformats.org/package/2006/content-types"

    changed = False
    existing = {
        ov.get("PartName")
        for ov in root
        if ov.tag == f"{{{ns}}}Override"
    }

    # We need a Default for math content type if not present
    # OMML uses the mathml equation content type
    default_exists = any(
        d.get("Extension") == "xml"
        for d in root
        if d.tag == f"{{{ns}}}Default"
    )

    # Add Override for document.xml with math content type
    for part_name in ["/word/document.xml"]:
        if part_name not in existing:
            ov = etree.SubElement(root, f"{{{ns}}}Override")
            ov.set("PartName", part_name)
            ov.set("ContentType", MATH_CONTENT_TYPE)
            changed = True

    if changed:
        ct_path.write_bytes(
            etree.tostring(
                root,
                encoding="utf-8",
                xml_declaration=True,
                standalone=True,
                pretty_print=True,
            )
        )


def _repack_docx(workdir: Path, output_path: Path) -> None:
    """Repack the working directory into a new .docx (ZIP) file."""
    output_path.parent.mkdir(parents=True, exist_ok=True)

    with zipfile.ZipFile(str(output_path), "w", zipfile.ZIP_DEFLATED) as z:
        for fpath in sorted(workdir.rglob("*")):
            if fpath.is_file():
                arcname = str(fpath.relative_to(workdir)).replace("\\", "/")
                z.write(str(fpath), arcname)
