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
from .service import convert_mathtype_to_omml


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
    conv.add_argument("-w", "--workers", type=int, default=8,
                      help="Number of cpu threads (default: 8)")
    conv.add_argument("--font", default="Cambria Math",
                      help="Math font (default: Cambria Math)")

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

    tasks: list[tuple[Path, Path]] = []

    # Phase 1: Scan
    if input_docx.is_file():
        # Single file
        if output_docx.is_dir() or str(args.output).endswith(("/", "\\")):
            output_docx.mkdir(parents=True, exist_ok=True)
            dst_file = output_docx / input_docx.name
        else:
            dst_file = output_docx
        tasks.append((input_docx, dst_file))

    elif input_docx.is_dir():
        # Whole folder
        output_docx.mkdir(parents=True, exist_ok=True)
        docx_files = [
            f for f in sorted(input_docx.rglob("*.docx"))
            if not f.name.startswith(("~$", ".")) and output_docx not in f.resolve().parents
        ]
        for src_file in docx_files:
            rel = src_file.relative_to(input_docx)
            dst_file = output_docx / rel
            tasks.append((src_file, dst_file))

    if not tasks:
        print("ERROR: Failed to scan docx files.", file=sys.stderr)
        return 1

    # Phase 2: Call service from service.py to process files
    total_files = len(tasks)
    total_formulas = 0
    total_succeeded = 0
    total_failed = 0
    all_reports = []

    for i, (src, dst) in enumerate(tasks, 1):
        if total_files > 1:
            print(f"[{i}/{total_files}] Processing: {src.name}")
        dst.parent.mkdir(parents=True, exist_ok=True)

        def _progress(phase, cur, total):
            if not args.quiet:
                print(f"  [{phase.upper()}] {cur}/{total}...", end="\r", flush=True)

        try:
            report = convert_mathtype_to_omml(
                str(src),
                str(dst),
                parallel=True,
                max_workers=args.workers,
                xsl_path=args.xsl,
                math_font=args.font,
                progress_callback=_progress,
            )
            total_formulas += report.total_ole_objects
            total_succeeded += report.succeeded
            total_failed += report.failed
            all_reports.append(report.to_dict())

            succeeded = report.succeeded
            failed = report.failed
            total = report.total_ole_objects
            print(f"Conversion complete: {succeeded}/{total} succeeded, {failed} failed")

            if not args.quiet:
                failed_formulas = [f for f in report.formulas if f.status != FormulaStatus.REPLACED]
                if failed_formulas:
                    print(f"  -> Detail: {len(failed_formulas)} failed (keep OLE):")
                    for f in failed_formulas:
                        risk_tag = f.risk_level.value if f.risk_level != RiskLevel.AUTO_REPLACE else ""
                        extra = f" [{risk_tag}]" if risk_tag else ""
                        err = f" — {f.error_message}" if f.error_message else ""
                        print(f"    ! [FAILED] {f.formula_id} ({f.ole_name}){extra}{err}")

        except Exception as e:
            print(f"  -> Failed: {e}", file=sys.stderr)
            total_failed += 1

    # Phase 3: export report to file
    if args.report:
        report_path = Path(args.report)
        report_path.parent.mkdir(parents=True, exist_ok=True)
        report_path.write_text(
            json.dumps(reports, indent=2, ensure_ascii=False),
            encoding="utf-8",
        )
        log(f"Report written to: {args.report}")

    return 0 if total_failed == 0 else 2


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
