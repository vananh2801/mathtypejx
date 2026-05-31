"""Service API: single-call MathType OLE → OMML conversion for Word docx."""

import re
import shutil
import tempfile
import zipfile
from pathlib import Path
from typing import Optional, Callable
from dataclasses import dataclass, field
from concurrent.futures import ThreadPoolExecutor, as_completed

from lxml import etree

from .models import FormulaInfo, FormulaStatus, RiskLevel, ConversionReport
from .scanner import scan_docx, read_ole_binary
from .extractor import extract_mtef, detect_mtef_version
from .converter import convert_formula
from .replacer import replace_formulas
from .validator import validate_output_docx, make_report, validate_formula_count


# ── Public API ─────────────────────────────────────────────────

def convert_mathtype_to_omml(
    docx_path: str,
    output_path: Optional[str] = None,
    *,
    remove_edit_info: bool = True,
    parallel: bool = True,
    max_workers: int = 8,
    progress_callback: Optional[Callable[[str, int, int], None]] = None,
) -> ConversionReport:
    """Convert all MathType OLE formulas in a Word docx to OMML.

    This is the single entry point. It handles everything:
      1. Scan the docx for MathType OLE objects
      2. Extract MTEF binary data from each OLE
      3. Convert MTEF → MathML (pure Python) → OMML (via XSLT)
      4. Replace OLE objects with OMML in the document XML
      5. Remove 排版/校对 edit info from the document body
      6. Validate the output document integrity

    Args:
        docx_path: Path to input .docx file.
        output_path: Path for output .docx. If None, creates a copy
            with '_omml' suffix alongside the input.
        remove_edit_info: Strip 排版/校对/校正 lines from document.
        parallel: Use threaded parallel conversion for formulas.
        max_workers: Max threads for parallel conversion (default 8).
        progress_callback: Optional callback(phase, current, total)
            called during conversion phases.

    Returns:
        ConversionReport with per-formula status, risk levels, and stats.

    Raises:
        FileNotFoundError: If docx_path does not exist.
        RuntimeError: If the docx cannot be read as a valid ZIP.
    """
    input_path = Path(docx_path).resolve()
    if not input_path.exists():
        raise FileNotFoundError(f"Input file not found: {input_path}")

    if output_path is None:
        output_path = str(input_path.parent / f"{input_path.stem}_omml.docx")
    output_path = Path(output_path).resolve()

    _notify(progress_callback, "scan", 0, 1)

    # ── Phase 1: Scan ─────────────────────────────────────
    formulas, workdir = scan_docx(str(input_path))
    mathtype_formulas = [
        f for f in formulas
        if f.prog_id in ("Equation.DSMT4", "Equation.3", "Equation.2")
    ]

    if not mathtype_formulas:
        shutil.copy2(str(input_path), str(output_path))
        report = make_report(str(input_path), str(output_path), [])
        _notify(progress_callback, "done", 0, 0)
        return report

    _notify(progress_callback, "extract", 0, len(mathtype_formulas))

    # ── Phase 2: Extract ──────────────────────────────────
    extracted = []
    for i, formula in enumerate(mathtype_formulas):
        ole_data = read_ole_binary(workdir, formula.ole_name)
        formula.ole_data = ole_data
        if ole_data is None:
            formula.status = FormulaStatus.FAILED
            formula.error_message = "OLE binary not found"
            continue
        if extract_mtef(formula):
            formula.mtef_version = detect_mtef_version(formula.mtef_bytes)
            extracted.append(formula)
        _notify(progress_callback, "extract", i + 1, len(mathtype_formulas))

    _notify(progress_callback, "convert", 0, len(extracted))

    # ── Phase 3: Convert (parallel or serial) ─────────────
    if parallel and len(extracted) > 1:
        _convert_parallel(extracted, max_workers, progress_callback)
    else:
        for i, formula in enumerate(extracted):
            try:
                convert_formula(formula)
            except Exception as e:
                formula.status = FormulaStatus.FAILED
                formula.error_message = str(e)
            _notify(progress_callback, "convert", i + 1, len(extracted))

    _notify(progress_callback, "replace", 0, 1)

    # ── Phase 4: Replace ─────────────────────────────────
    replace_formulas(workdir, formulas, str(output_path))

    # ── Phase 5: Remove edit info ─────────────────────────
    if remove_edit_info:
        _strip_edit_info(str(output_path))

    # ── Phase 6: Validate ─────────────────────────────────
    validate_output_docx(str(output_path))

    # Clean up
    try:
        shutil.rmtree(workdir, ignore_errors=True)
    except Exception:
        pass

    report = make_report(str(input_path), str(output_path), formulas)
    _notify(progress_callback, "done", report.succeeded, report.total_ole_objects)
    return report


def health_check() -> dict:
    """Check if all external dependencies are available.

    Returns dict mapping dependency name to status string.
    """
    from .validator import health_check as _hc
    return _hc()


# ── Bulk API ───────────────────────────────────────────────────

def convert_batch(
    docx_paths: list[str],
    output_dir: str,
    *,
    remove_edit_info: bool = True,
    max_workers: int = 4,
    progress_callback: Optional[Callable[[str, int, int], None]] = None,
) -> list[ConversionReport]:
    """Convert multiple docx files. Each file is processed independently.

    For large batches, processing is sequential per-file but formulas
    within each file can be converted in parallel.
    """
    out_dir = Path(output_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    reports = []

    for i, path in enumerate(docx_paths):
        name = Path(path).stem
        out_path = out_dir / f"{name}_omml.docx"
        _notify(progress_callback, "file", i + 1, len(docx_paths))
        report = convert_mathtype_to_omml(
            path, str(out_path),
            remove_edit_info=remove_edit_info,
            parallel=True, max_workers=max_workers,
        )
        reports.append(report)

    _notify(progress_callback, "done", len(reports), len(docx_paths))
    return reports


# ── Internals ──────────────────────────────────────────────────

# Edit info patterns (排版/校对/校正/一校/二校/三校/编辑/制图)
_EDIT_PATTERNS = [
    re.compile(r'排版[：:].*'),
    re.compile(r'校对[：:].*'),
    re.compile(r'校正[：:].*'),
    re.compile(r'一校[：:].*'),
    re.compile(r'二校[：:].*'),
    re.compile(r'三校[：:].*'),
    re.compile(r'编辑[：:].*'),
    re.compile(r'制图[：:].*'),
]


def _strip_edit_info(docx_path: str) -> list[str]:
    """Remove 排版/校对/校正 info paragraphs from a docx in-place."""
    tmp_path = docx_path + ".tmp"
    removed = []

    with zipfile.ZipFile(docx_path, "r") as zin:
        with zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == "word/document.xml":
                    root = etree.fromstring(data)
                    NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                    for p in list(root.iter(f"{{{NS_W}}}p")):
                        texts = "".join(
                            t.text or "" for t in p.iter(f"{{{NS_W}}}t")
                        )
                        for pat in _EDIT_PATTERNS:
                            if pat.match(texts.strip()):
                                removed.append(texts.strip()[:80])
                                parent = p.getparent()
                                if parent is not None:
                                    parent.remove(p)
                                break
                    data = etree.tostring(
                        root, encoding="utf-8",
                        xml_declaration=data.strip().startswith(b"<?xml"),
                    )
                zout.writestr(item, data)

    Path(docx_path).unlink()
    Path(tmp_path).rename(docx_path)
    return removed


def _convert_parallel(
    formulas: list[FormulaInfo],
    max_workers: int,
    progress_callback=None,
) -> None:
    """Convert formulas in parallel using ThreadPoolExecutor.

    Conversion is independent per formula, so threads provide useful
    parallelism across OLE objects and XSLT work.
    """
    from threading import Lock

    total = len(formulas)
    done = 0
    lock = Lock()

    def convert_one(f):
        try:
            convert_formula(f)
        except Exception as e:
            f.status = FormulaStatus.FAILED
            f.error_message = str(e)
        nonlocal done
        with lock:
            nonlocal done
            done += 1
            _notify(progress_callback, "convert", done, total)
        return f

    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures = [executor.submit(convert_one, f) for f in formulas]
        for _ in as_completed(futures):
            pass


def _notify(callback, phase: str, current: int, total: int) -> None:
    """Call progress callback if provided."""
    if callback:
        try:
            callback(phase, current, total)
        except Exception:
            pass
