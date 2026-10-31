"""Formulas stored as WMF pictures that carry their MathType (MTEF) data.

MathType writes the equation source into META_ESCAPE/MFCOMMENT records of
the metafile ("AppsMFCC" + "Design Science, Inc." + MTEF bytes).  Documents
where the OLE object was stripped (or "converted to picture") keep that
data, so such pictures can still be turned back into editable equations.
"""

from __future__ import annotations

import struct
from pathlib import Path
from typing import Optional

from lxml import etree

from .models import FormulaInfo, FormulaStatus
from .scanner import XML_PARTS, _rels_path_for_part

NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS_V = "urn:schemas-microsoft-com:vml"
NS_O = "urn:schemas-microsoft-com:office:office"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"

_SIGNATURE = b"Design Science, Inc.\x00"
_PLACEABLE = bytes.fromhex("d7cdc69a")


def extract_mtef_from_wmf(data: bytes) -> Optional[bytes]:
    """Return the MTEF bytes embedded in a MathType WMF, or None."""
    off = 22 if data[:4] == _PLACEABLE else 0
    pos = off + 18  # skip standard METAHEADER
    total = None
    payload = b""
    while pos + 6 <= len(data):
        size, func = struct.unpack_from("<IH", data, pos)
        if size < 3:
            break
        if func == 0x0626 and pos + 10 <= len(data):  # META_ESCAPE
            esc, count = struct.unpack_from("<HH", data, pos + 6)
            chunk = data[pos + 10: pos + 10 + count]
            if esc == 15 and chunk[:8] == b"AppsMFCC" and len(chunk) >= 18:
                size_total, size_chunk = struct.unpack_from("<II", chunk, 10)
                body = chunk[18:]
                if body.startswith(_SIGNATURE):
                    body = body[len(_SIGNATURE):]
                if total is None:
                    total = size_total
                payload += body[:size_chunk] if size_chunk else body
        pos += size * 2
    if not payload:
        return None
    if total:
        payload = payload[:total]
    # MTEF streams start with version 3 or 5
    return payload if payload[:1] in (b"\x03", b"\x05") else None


def _resolve(workdir: Path, part_name: str, target: str) -> Path:
    if target.startswith("/"):
        return workdir / target.lstrip("/")
    return (workdir / part_name).parent / target


def _enclosing(elem, tags):
    while elem is not None and elem.tag not in tags:
        elem = elem.getparent()
    return elem


def scan_picture_formulas(
    workdir: Path,
    start_counter: int = 0,
) -> list[FormulaInfo]:
    """Find WMF pictures that contain MTEF and are not OLE previews."""
    formulas: list[FormulaInfo] = []
    counter = start_counter
    container_tags = {f"{{{NS_W}}}object", f"{{{NS_W}}}pict", f"{{{NS_W}}}drawing"}

    for part_name in XML_PARTS:
        part_path = workdir / part_name
        rels_path = workdir / _rels_path_for_part(part_name)
        if not part_path.exists() or not rels_path.exists():
            continue
        try:
            tree = etree.parse(str(part_path))
            rels = {
                r.get("Id"): r.get("Target", "")
                for r in etree.parse(str(rels_path)).getroot()
                if r.get("TargetMode") != "External"
            }
        except etree.XMLSyntaxError:
            continue

        seen: dict[str, Optional[bytes]] = {}
        done_rids: set[str] = set()
        for elem in tree.getroot().iter():
            if elem.tag == f"{{{NS_V}}}imagedata":
                rid = elem.get(f"{{{NS_R}}}id", "")
            elif elem.tag == f"{{{NS_A}}}blip":
                rid = elem.get(f"{{{NS_R}}}embed", "")
            else:
                continue
            if not rid or rid in done_rids or rid not in rels:
                continue
            target = rels[rid]
            if not target.lower().endswith(".wmf"):
                continue
            holder = _enclosing(elem, container_tags)
            if holder is None:
                continue
            if holder.find(f".//{{{NS_O}}}OLEObject") is not None:
                continue  # preview of an OLE object: handled elsewhere

            if target not in seen:
                img = _resolve(workdir, part_name, target)
                try:
                    seen[target] = extract_mtef_from_wmf(img.read_bytes())
                except OSError:
                    seen[target] = None
            mtef = seen[target]
            if not mtef:
                continue

            counter += 1
            done_rids.add(rid)
            formulas.append(FormulaInfo(
                formula_id=f"F{counter:04d}",
                ole_name=Path(target).name,
                part_name=part_name,
                rels_path=_rels_path_for_part(part_name),
                relationship_id=rid,
                prog_id="MathType.WMF",
                status=FormulaStatus.EXTRACTED,
                mtef_bytes=mtef,
                mtef_version=mtef[0],
                source="picture",
            ))
    return formulas
