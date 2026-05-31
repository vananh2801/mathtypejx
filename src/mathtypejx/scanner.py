"""Scanner: scan a docx file for MathType OLE formula objects."""

import shutil
import tempfile
import zipfile
from pathlib import Path
from typing import Optional

from lxml import etree

from .models import FormulaInfo, FormulaType

# XML namespaces used in OOXML
NS = {
    "w":   "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r":   "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "o":   "urn:schemas-microsoft-com:office:office",
    "v":   "urn:schemas-microsoft-com:vml",
    "mc":  "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "wps": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
    "wp":  "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "a":   "http://schemas.openxmlformats.org/drawingml/2006/main",
    "m":   "http://schemas.openxmlformats.org/officeDocument/2006/math",
}

# ProgIDs that indicate MathType or Equation Editor formulas
MATHTYPE_PROG_IDS = {
    "Equation.DSMT4",   # MathType 4+
    "Equation.3",       # Equation Editor 3.x
    "Equation.2",       # Equation Editor 2.x
}

# Parts in a docx that can contain formulas
XML_PARTS = [
    "word/document.xml",
    "word/header1.xml",
    "word/header2.xml",
    "word/header3.xml",
    "word/footer1.xml",
    "word/footer2.xml",
    "word/footer3.xml",
    "word/footnotes.xml",
    "word/endnotes.xml",
]


def _unzip_docx(docx_path: Path, workdir: Path) -> None:
    """Extract docx (ZIP) contents to workdir."""
    with zipfile.ZipFile(docx_path, "r") as z:
        z.extractall(workdir)


def _rels_path_for_part(part_name: str) -> str:
    """Given a part like 'word/document.xml', return its .rels path."""
    name = Path(part_name).name
    parent = Path(part_name).parent
    return str(parent / "_rels" / f"{name}.rels").replace("\\", "/")


def _resolve_rid_to_target(workdir: Path, part_name: str, rid: str) -> Optional[str]:
    """Resolve an r:id to the target path using the part's .rels file."""
    rels_path = workdir / _rels_path_for_part(part_name)
    if not rels_path.exists():
        return None

    try:
        tree = etree.parse(str(rels_path))
    except etree.XMLSyntaxError:
        return None

    for rel in tree.getroot():
        rel_id = rel.get("Id")
        if rel_id == rid:
            target = rel.get("Target", "")
            # Target is relative to the part's directory; make it relative to docx root
            part_dir = Path(part_name).parent
            resolved = str((part_dir / target).resolve().relative_to(Path.cwd()))
            return resolved.replace("\\", "/")
    return None


def scan_docx(docx_path: str, workdir: Optional[str] = None) -> tuple[list[FormulaInfo], Path]:
    """Scan a docx file and return all detected MathType OLE formulas.

    Args:
        docx_path: Path to the input .docx file.
        workdir: Optional working directory path. If None, a temp dir is used.

    Returns:
        Tuple of (list of FormulaInfo, path to the unpacked working directory).
    """
    docx_path = Path(docx_path).resolve()
    if workdir:
        wd = Path(workdir)
        wd.mkdir(parents=True, exist_ok=True)
    else:
        wd = Path(tempfile.mkdtemp(prefix="mathtypejx_"))

    _unzip_docx(docx_path, wd)

    formulas: list[FormulaInfo] = []
    formula_counter = 0

    for part_name in XML_PARTS:
        part_path = wd / part_name
        if not part_path.exists():
            continue

        try:
            parser = etree.XMLParser(remove_blank_text=False)
            tree = etree.parse(str(part_path), parser)
        except etree.XMLSyntaxError:
            continue

        # Build RID → ole_name map from .rels file
        rid_to_ole: dict[str, str] = {}
        rels_path = wd / _rels_path_for_part(part_name)
        if rels_path.exists():
            try:
                rels_tree = etree.parse(str(rels_path))
                for rel in rels_tree.getroot():
                    rid = rel.get("Id", "")
                    target = rel.get("Target", "")
                    if "oleObject" in target:
                        ole_name = Path(target).name
                        rid_to_ole[rid] = ole_name
            except etree.XMLSyntaxError:
                pass

        # Find OLEObject elements
        for ole_elem in part_path_ole_objects(tree, part_name):
            rid = ole_elem.get(
                "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id",
                ""
            )
            prog_id = ole_elem.get("ProgID", "")

            ole_name = rid_to_ole.get(rid, f"oleObject_unknown_{formula_counter}.bin")

            formula_counter += 1
            formula_id = f"F{formula_counter:04d}"

            formula_type = FormulaType.MATHTYPE_OLE
            if "Equation.DSMT4" in prog_id:
                formula_type = FormulaType.MATHTYPE_OLE
            elif "Equation.3" in prog_id or "Equation.2" in prog_id:
                formula_type = FormulaType.EQUATION_EDITOR_3
            else:
                formula_type = FormulaType.UNKNOWN

            # Determine para/run index by walking up the tree
            para_idx, run_idx = _find_location(ole_elem)

            formulas.append(FormulaInfo(
                formula_id=formula_id,
                ole_name=ole_name,
                part_name=part_name,
                rels_path=_rels_path_for_part(part_name),
                relationship_id=rid,
                prog_id=prog_id,
                para_index=para_idx,
                run_index=run_idx,
                formula_type=formula_type,
            ))

    return formulas, wd


def part_path_ole_objects(tree, part_name: str):
    """Yield all o:OLEObject elements in a parsed XML tree."""
    for ole in tree.iter():
        if ole.tag == f"{{{NS['o']}}}OLEObject":
            prog_id = ole.get("ProgID", "")
            if prog_id in MATHTYPE_PROG_IDS:
                yield ole


def _find_location(elem) -> tuple[int, int]:
    """Walk up from elem to find paragraph index and run index."""
    para_idx = -1
    run_idx = -1

    # Walk up to find <w:r> and <w:p>
    current = elem
    while current is not None:
        tag = current.tag
        if tag == f"{{{NS['w']}}}r" and run_idx == -1:
            parent = current.getparent()
            if parent is not None:
                for i, child in enumerate(parent):
                    if child is current:
                        run_idx = i
                        break
        elif tag == f"{{{NS['w']}}}p":
            parent = current.getparent()
            if parent is not None:
                for i, child in enumerate(parent):
                    if child is current:
                        para_idx = i
                        break
            break
        current = current.getparent()

    return para_idx, run_idx


def read_ole_binary(workdir: Path, ole_name: str) -> Optional[bytes]:
    """Read the raw OLE binary data from word/embeddings/ in the unpacked docx."""
    ole_path = workdir / "word" / "embeddings" / ole_name
    if ole_path.exists():
        return ole_path.read_bytes()
    return None
