"""CLI entry point for the MathType to OMML pipeline."""

import argparse
import json
import sys
from pathlib import Path
from typing import Optional

from .models import FormulaStatus, RiskLevel
from .scanner import scan_docx, read_ole_binary
from .extractor import extract_mtef, detect_mtef_version
from .converter import convert_formula
from .replacer import replace_formulas
from .validator import (
    make_report,
    validate_formula_count,
    validate_output_docx,
    health_check,
    try_word_open,
)


def main(argv: Optional[list[str]] = None) -> int:
    parser = argparse.ArgumentParser(
        prog="mathtypejx",
        description="Convert MathType OLE formulas to OMML in Word docx files.",
    )
    sub = parser.add_subparsers(dest="command")

    # convert command
    conv = sub.add_parser("convert", help="Convert a docx file")
    conv.add_argument("input", help="Input .docx file path")
    conv.add_argument("-o", "--output", required=True, help="Output .docx file path")
    conv.add_argument("--workdir", help="Working directory (temp if not specified)")
    conv.add_argument("--xsl", help="Path to MML2OMML.XSL")
    conv.add_argument("--keep-workdir", action="store_true",
                       help="Do not clean up working directory after conversion")
    conv.add_argument("--word-validate", action="store_true",
                       help="Try to open output with Word COM after conversion")
    conv.add_argument("--report", help="Write conversion report JSON to this path")
    conv.add_argument("--quiet", action="store_true", help="Suppress progress output")

    # health command
    sub.add_parser("health", help="Check dependency availability")

    args = parser.parse_args(argv)

    if args.command == "health":
        return _cmd_health()
    elif args.command == "convert":
        return _cmd_convert(args)
    else:
        parser.print_help()
        return 1


def _cmd_convert(args) -> int:
    """Run the full conversion pipeline."""
    input_docx = Path(args.input).resolve()
    if not input_docx.exists():
        print(f"ERROR: Input file not found: {input_docx}", file=sys.stderr)
        return 1

    output_docx = Path(args.output).resolve()

    log = lambda msg: None if args.quiet else print(f"  {msg}")

    # Phase 1: Scan
    log("Scanning for MathType OLE formulas...")
    workdir = None
    try:
        formulas, workdir = scan_docx(str(input_docx), workdir=args.workdir)
    except Exception as e:
        print(f"ERROR: Failed to scan docx: {e}", file=sys.stderr)
        return 1

    mathtype_formulas = [
        f for f in formulas
        if f.prog_id in ("Equation.DSMT4", "Equation.3", "Equation.2")
    ]
    log(f"Found {len(mathtype_formulas)} MathType/Equation Editor formulas "
        f"({len(formulas)} total OLE objects)")

    if not mathtype_formulas:
        log("No MathType formulas to convert. Creating copy of input.")
        import shutil
        shutil.copy2(str(input_docx), str(output_docx))
        return 0

    # Phase 2: Extract MTEF from each OLE
    for formula in mathtype_formulas:
        log(f"Extracting: {formula.ole_name} (r:id={formula.relationship_id})")
        ole_data = read_ole_binary(workdir, formula.ole_name)
        formula.ole_data = ole_data

        if ole_data is None:
            formula.status = FormulaStatus.FAILED
            formula.error_message = f"OLE binary not found: {formula.ole_name}"
            log(f"  FAILED: {formula.error_message}")
            continue

        ok = extract_mtef(formula)
        if ok:
            formula.mtef_version = detect_mtef_version(formula.mtef_bytes)
            log(f"  Extracted MTEF v{formula.mtef_version} ({len(formula.mtef_bytes)} bytes)")
        else:
            log(f"  FAILED: {formula.error_message}")

    # Phase 3: Convert MTEF → MathML → OMML
    for formula in mathtype_formulas:
        if formula.status != FormulaStatus.EXTRACTED:
            continue
        log(f"Converting: {formula.ole_name}")
        try:
            ok = convert_formula(
                formula,
                xsl_path=args.xsl,
            )
            if ok:
                log(f"  Converted → OMML (risk: {formula.risk_level.value})")
            else:
                log(f"  FAILED: {formula.error_message}")
        except Exception as e:
            formula.status = FormulaStatus.FAILED
            formula.error_message = str(e)
            log(f"  FAILED: {e}")

    # Phase 4: Replace in docx
    log("Replacing OLE objects with OMML in document XML...")
    try:
        replace_formulas(workdir, formulas, str(output_docx))
    except Exception as e:
        print(f"ERROR: Replacement failed: {e}", file=sys.stderr)
        return 1

    # Phase 5: Validate
    log("Validating output docx...")
    val_result = validate_output_docx(str(output_docx))
    if not val_result["valid"]:
        for err in val_result["errors"]:
            log(f"  VALIDATION ERROR: {err}")

    # Word COM validation
    if args.word_validate:
        log("Checking with Word COM...")
        word_result = try_word_open(str(output_docx))
        if word_result["opened"]:
            log("  Word opened successfully")
        else:
            log(f"  Word check: {word_result['error']}")

    # Report
    report = make_report(str(input_docx), str(output_docx), formulas)
    count_result = validate_formula_count(report)
    log(f"Formula count conserved: {count_result['conserved']}")

    if args.report:
        report_path = Path(args.report)
        report_path.parent.mkdir(parents=True, exist_ok=True)
        report_path.write_text(
            json.dumps(report.to_dict(), indent=2, ensure_ascii=False),
            encoding="utf-8",
        )
        log(f"Report written to: {args.report}")

    # Summary
    succeeded = report.succeeded
    failed = report.failed
    total = len(mathtype_formulas)
    print(f"\nConversion complete: {succeeded}/{total} succeeded, {failed} failed")

    if not args.quiet:
        for f in formulas:
            marker = "OK" if f.status == FormulaStatus.REPLACED else f.status.value.upper()
            risk_tag = f.risk_level.value if f.risk_level != RiskLevel.AUTO_REPLACE else ""
            extra = f" [{risk_tag}]" if risk_tag else ""
            err = f" — {f.error_message}" if f.error_message else ""
            print(f"  [{marker}] {f.formula_id} {f.ole_name}{extra}{err}")

    # Cleanup
    if not args.keep_workdir and workdir is not None:
        import shutil
        try:
            shutil.rmtree(workdir, ignore_errors=True)
        except Exception:
            pass

    return 0 if failed == 0 else 2


def _cmd_health() -> int:
    """Check dependency availability."""
    print("MathTypeJX Health Check\n")
    results = health_check()
    all_ok = True
    for name, status in sorted(results.items()):
        ok = status.startswith("ok") or status == "ok"
        if not ok:
            all_ok = False
        marker = "OK" if ok else "MISSING"
        detail = status if not ok else (status[3:] if status.startswith("ok:") else "")
        print(f"  [{marker}] {name} {detail}")
    print(f"\nOverall: {'READY' if all_ok else 'NOT READY — install missing dependencies'}")
    return 0 if all_ok else 1


if __name__ == "__main__":
    sys.exit(main())
