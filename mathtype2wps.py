#!/usr/bin/env python3
"""Chuyển công thức MathType (OLE) trong .docx thành công thức Word gốc (OMML)
để WPS Office / Word / LibreOffice chỉnh sửa được.

Dùng được trên Linux, KHÔNG cần Microsoft Office.

Ví dụ:
    python3 mathtype2wps.py bai1.docx
    python3 mathtype2wps.py bai1.docx bai2.docx -o ketqua/
    python3 mathtype2wps.py thu_muc_tai_lieu/ -o ketqua/ --report
"""
from __future__ import annotations

import argparse
import json
import shutil
import sys
import tempfile
from pathlib import Path
import os
from concurrent.futures import ThreadPoolExecutor

# Cho phép chạy trực tiếp từ thư mục repo mà không cần pip install
_SRC = Path(__file__).resolve().parent / "src"
if _SRC.is_dir():
    sys.path.insert(0, str(_SRC))

try:
    from mathtypejx.converter import convert_formula
    from mathtypejx.extractor import detect_mtef_version, extract_mtef
    from mathtypejx.models import FormulaStatus
    from mathtypejx.replacer import replace_formulas
    from mathtypejx.scanner import read_ole_binary, scan_docx
    from mathtypejx.validator import make_report, validate_output_docx
    from mathtypejx.wmf import scan_picture_formulas
except ImportError as exc:  # pragma: no cover
    sys.exit(
        f"Thiếu thư viện: {exc}\n"
        "Cài bằng: pip install lxml olefile python-docx mathml2omml"
    )

MATHTYPE_IDS = ("Equation.DSMT4", "Equation.3", "Equation.2")


def convert_one(src: Path, dst: Path, xsl: str | None, quiet: bool,
                font: str | None = "Cambria Math", workers: int = 8) -> dict:
    """Chuyển 1 file. Trả về dict thống kê."""
    workdir = Path(tempfile.mkdtemp(prefix="mt2wps_"))
    try:
        formulas, workdir = scan_docx(str(src), workdir=str(workdir))
        mt = [f for f in formulas if f.prog_id in MATHTYPE_IDS]

        # Công thức dạng ảnh WMF còn chứa dữ liệu MathType (không có OLE)
        pics = scan_picture_formulas(workdir, start_counter=len(formulas))
        n_ole = len(mt)
        mt = mt + pics

        if not mt:
            shutil.copy2(src, dst)
            return {"file": src.name, "total": 0, "ok": 0, "failed": 0,
                    "note": "không có công thức MathType (đã sao chép nguyên bản)"}

        for f in mt:
            if f.source == "picture":
                continue  # MTEF đã được trích từ ảnh WMF
            f.ole_data = read_ole_binary(workdir, f.ole_name)
            if f.ole_data is None:
                f.status = FormulaStatus.FAILED
                f.error_message = f"Không thấy {f.ole_name}"
                continue
            if extract_mtef(f):
                f.mtef_version = detect_mtef_version(f.mtef_bytes)

        # -------------------------------------------------------------
        # THAY THẾ ĐOẠN CONVERT ĐƠN NHÂN BẰNG MULTI-THREADS TẠI ĐÂY
        # -------------------------------------------------------------
        to_convert = [f for f in mt if f.status == FormulaStatus.EXTRACTED]
        if not quiet and to_convert:
            print(f"    -> Đang xử lý song song {len(to_convert)} công thức trên {workers} luồng CPU...", flush=True)

        def _worker_task(formula):
            try:
                convert_formula(formula, xsl_path=xsl, math_font=font)
            except Exception as exc:  # noqa: BLE001
                formula.status = FormulaStatus.FAILED
                formula.error_message = str(exc)

        if workers > 1 and len(to_convert) > 1:
            with ThreadPoolExecutor(max_workers=workers) as executor:
                list(executor.map(_worker_task, to_convert))
        else:
            for f in to_convert:
                _worker_task(f)
        # -------------------------------------------------------------

        replace_formulas(workdir, mt, str(dst))
        check = validate_output_docx(str(dst))
        report = make_report(str(src), str(dst), mt)

        failed = [f for f in mt if f.status == FormulaStatus.FAILED]
        if not quiet:
            for f in failed:
                print(f"    ! {f.formula_id} ({f.ole_name}): {f.error_message}")
            if not check["valid"]:
                for err in check["errors"]:
                    print(f"    ! kiểm tra file đầu ra: {err}")
        return {
            "file": src.name,
            "total": len(mt),
            "ok": report.succeeded,
            "failed": report.failed,
            "from_pictures": len(pics),
            "ole": n_ole,
            "output_valid": check["valid"],
            "failures": [
                {"id": f.formula_id, "ole": f.ole_name, "error": f.error_message}
                for f in failed
            ],
            "report": report.to_dict(),
        }
    finally:
        shutil.rmtree(workdir, ignore_errors=True)


def collect_inputs(paths: list[str], skip_dir: Path | None = None):
    """Trả về (danh sách (file, đường dẫn tương đối), danh sách file khác).

    Với thư mục: đường dẫn tương đối tính từ thư mục đó, nên cấu trúc thư mục
    con được giữ nguyên. Nếu truyền nhiều đầu vào, thêm tên thư mục gốc vào
    đường dẫn tương đối để tránh trùng tên.
    """
    docs: list[tuple[Path, Path]] = []
    others: list[tuple[Path, Path]] = []
    multi = len(paths) > 1
    skip = skip_dir.resolve() if skip_dir else None

    for p in map(Path, paths):
        if p.is_dir():
            base = p.name if multi else ""
            for x in sorted(p.rglob("*")):
                if not x.is_file() or x.name.startswith(("~$", ".")):
                    continue
                if skip and skip in x.resolve().parents:
                    continue  # bỏ qua thư mục kết quả nếu nằm trong thư mục nguồn
                rel = Path(base) / x.relative_to(p)
                (docs if x.suffix.lower() == ".docx" else others).append((x, rel))
        elif p.suffix.lower() == ".docx" and p.exists():
            docs.append((p, Path(p.name)))
        else:
            print(f"Bỏ qua (không phải .docx / không tồn tại): {p}", file=sys.stderr)
    return docs, others


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("inputs", nargs="+", help="File .docx hoặc thư mục chứa .docx")
    ap.add_argument("-o", "--outdir",
                    help="Thư mục xuất. Cấu trúc thư mục con được giữ nguyên, "
                         "tên file giữ nguyên (không thêm hậu tố)")
    ap.add_argument("-s", "--suffix", default=None,
                    help="Hậu tố tên file đầu ra. Mặc định: không có khi dùng -o, "
                         "'_wps' khi ghi cạnh file gốc")
    ap.add_argument("--copy-others", action="store_true",
                    help="Sao chép cả các file không phải .docx (pdf, ảnh, .doc...) "
                         "để thư mục kết quả đầy đủ như thư mục gốc")
    ap.add_argument("--xsl", help="Đường dẫn MML2OMML.XSL nếu có (không bắt buộc)")
    ap.add_argument("--font", default="Cambria Math",
                    help='Font cho công thức (mặc định: "Cambria Math"; '
                         'truyền "" để giữ font mặc định)')
    ap.add_argument("--report", action="store_true",
                    help="Ghi báo cáo JSON chi tiết cạnh mỗi file đầu ra")
    ap.add_argument("-q", "--quiet", action="store_true")
    ap.add_argument("-w", "--workers", type=int, default=os.cpu_count() or 8,
                help="Số luồng xử lý song song (mặc định: toàn bộ số luồng CPU, vd: 16)")
    args = ap.parse_args()

    outdir = Path(args.outdir) if args.outdir else None
    suffix = args.suffix if args.suffix is not None else ("" if outdir else "_wps")

    docs, others = collect_inputs(args.inputs, skip_dir=outdir)
    if not docs:
        print("Không có file .docx nào để xử lý.", file=sys.stderr)
        return 1
    if outdir:
        outdir.mkdir(parents=True, exist_ok=True)

    total_ok = total_fail = 0
    failed_files: list[str] = []
    summary: list[dict] = []
    for i, (src, rel) in enumerate(docs, 1):
        dst_dir = (outdir / rel.parent) if outdir else src.parent
        dst_dir.mkdir(parents=True, exist_ok=True)
        dst = dst_dir / f"{rel.stem}{suffix}.docx"
        if dst.resolve() == src.resolve():
            print(f"Bỏ qua {src}: trùng đường dẫn file đầu ra "
                  "(hãy dùng -o hoặc --suffix)", file=sys.stderr)
            continue
        print(f"[{i}/{len(docs)}] {rel}")
        try:
            res = convert_one(src, dst, args.xsl, args.quiet, args.font or None, workers=args.workers)
        except Exception as exc:  # noqa: BLE001
            print(f"    LỖI: {exc}", file=sys.stderr)
            total_fail += 1
            failed_files.append(str(rel))
            continue
        total_ok += res["ok"]
        total_fail += res["failed"]
        if res["failed"]:
            failed_files.append(f"{rel} ({res['failed']} công thức)")
        note = res.get("note") or f"{res['ok']}/{res['total']} công thức"
        if res.get("from_pictures"):
            note += f" (trong đó {res['from_pictures']} công thức dạng ảnh WMF)"
        print(f"    {note}")
        summary.append({k: res[k] for k in ("file", "total", "ok", "failed")}
                       | {"path": str(rel)})
        if args.report:
            dst.with_name(dst.stem + ".report.json").write_text(
                json.dumps(res, indent=2, ensure_ascii=False), encoding="utf-8")

    if outdir and args.copy_others:
        for src, rel in others:
            target = outdir / rel
            target.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(src, target)
        print(f"\nĐã sao chép {len(others)} file khác (không phải .docx).")

    print(f"\nTổng: {len(docs)} file .docx, {total_ok} công thức chuyển được, "
          f"{total_fail} lỗi (công thức lỗi vẫn giữ nguyên dạng ảnh/OLE cũ).")
    if failed_files:
        print("File còn lỗi cần xem lại:")
        for f in failed_files:
            print(f"  - {f}")
    return 0 if total_fail == 0 else 2


if __name__ == "__main__":
    sys.exit(main())
