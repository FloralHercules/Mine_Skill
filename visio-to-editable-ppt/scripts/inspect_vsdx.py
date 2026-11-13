"""Inspect one Visio page without Office or modifying the source drawing.

Run against export_visio.ps1's flattened.vsdx before building native PPTX.
Original drawings may also be inspected; unresolved groups are explicitly
reported so their children cannot be silently discarded or rasterized.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import math
import posixpath
import re
import sys
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

V = "http://schemas.microsoft.com/office/visio/2012/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS = {"v": V, "r": R}


def resolve_part(base_part: str, target: str) -> str:
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join(posixpath.dirname(base_part), target))


def rels_part(part: str) -> str:
    return posixpath.join(posixpath.dirname(part), "_rels", posixpath.basename(part) + ".rels")


def read_rels(archive, part):
    try:
        root = ET.fromstring(archive.read(rels_part(part)))
    except KeyError:
        return {}
    return {rel.get("Id"): dict(rel.attrib) for rel in root}


def cells(element):
    return {cell.get("N"): cell.get("V", "") for cell in element.findall("v:Cell", NS)}


def number(values, name, default=0.0):
    try:
        value = float(values.get(name, default))
        return value if math.isfinite(value) else default
    except (TypeError, ValueError):
        return default


def detect_mathtype(binary):
    marker = "Equation.DSMT4"
    if marker.encode("ascii") in binary or marker.encode("utf-16le") in binary:
        match = re.search(rb"MathType [0-9.]+ Equation", binary)
        return marker, match.group().decode("ascii") if match else "MathType Equation"
    return None, None


def inspect(source: Path, page_name: str | None, slide_width_in: float | None):
    warnings = []
    with zipfile.ZipFile(source) as archive:
        pages_part = "visio/pages/pages.xml"
        pages = ET.fromstring(archive.read(pages_part))
        page_list = [{"page_id": int(page.get("ID")), "name": page.get("Name"),
                      "name_u": page.get("NameU")} for page in pages]
        if page_name is None:
            if len(page_list) != 1:
                raise ValueError("Choose --page-name from: " + ", ".join(p["name"] or p["name_u"] for p in page_list))
            page_name = page_list[0]["name"] or page_list[0]["name_u"]
        matches = [p for p in pages if page_name in (p.get("Name"), p.get("NameU"))]
        if len(matches) != 1:
            raise ValueError(f"Expected one page named {page_name!r}, found {len(matches)}")
        page = matches[0]
        page_rel = page.find("v:Rel", NS)
        if page_rel is None:
            raise ValueError("Requested page has no package relationship")
        page_rel_id = page_rel.get(f"{{{R}}}id")
        page_link = read_rels(archive, pages_part)[page_rel_id]
        if page_link.get("TargetMode") == "External":
            raise ValueError("An external page relationship cannot be converted locally")
        page_part = resolve_part(pages_part, page_link["Target"])
        page_xml = ET.fromstring(archive.read(page_part))
        page_rels = read_rels(archive, page_part)
        page_sheet = page.find("v:PageSheet", NS)
        if page_sheet is None:
            raise ValueError("Requested page has no PageSheet dimensions")
        page_cells = cells(page_sheet)
        width = number(page_cells, "PageWidth")
        height = number(page_cells, "PageHeight")
        if width <= 0 or height <= 0:
            raise ValueError("PageWidth and PageHeight must be positive saved dimensions")
        output_width = slide_width_in if slide_width_in is not None else width
        if output_width <= 0:
            raise ValueError("--slide-width-in must be positive")
        output_height = output_width * height / width
        objects = []
        shapes = page_xml.find("v:Shapes", NS)
        for order, shape in enumerate(list(shapes) if shapes is not None else []):
            sid = int(shape.get("ID"))
            values = cells(shape)
            foreign = shape.find("v:ForeignData", NS)
            foreign_type = foreign.get("ForeignType") if foreign is not None else None
            source_part = None
            ole_binary = None
            prog_id = None
            ole_class_name = None
            binary_sha256 = None
            foreign_attrs = dict(foreign.attrib) if foreign is not None else {}
            if foreign is not None:
                rel = foreign.find("v:Rel", NS)
                if rel is None:
                    warnings.append({"source_id": sid, "code": "foreign-without-payload",
                                     "message": "Foreign object has no embedded-data relationship."})
                else:
                    rid = rel.get(f"{{{R}}}id")
                    relationship = page_rels.get(rid)
                    if relationship is None or relationship.get("TargetMode") == "External":
                        warnings.append({"source_id": sid, "code": "unresolved-foreign-payload",
                                         "message": f"Foreign relationship {rid!r} is missing or external."})
                    else:
                        source_part = resolve_part(page_part, relationship["Target"])
                        if foreign_type == "Object":
                            binary = archive.read(source_part)
                            ole_binary = posixpath.basename(source_part)
                            binary_sha256 = hashlib.sha256(binary).hexdigest()
                            prog_id, ole_class_name = detect_mathtype(binary)
                            if prog_id is None:
                                warnings.append({"source_id": sid, "code": "unknown-ole-class",
                                                 "message": "Unknown embedded OLE class; preserve its binary but do not invent a ProgID."})
            text = "\n".join("".join(t.itertext()).strip() for t in shape.findall("v:Text", NS)).strip()
            descendants = list(shape.findall(".//v:Shape", NS))
            is_group = shape.get("Type") == "Group" or bool(descendants)
            if is_group:
                warnings.append({"source_id": sid, "code": "unresolved-group",
                                 "message": "Run the private-copy flattening exporter; this group must not be replaced with a picture."})
            obj = {
                "source_id": sid, "source_parent_id": None, "z_order": order,
                "name": shape.get("Name", ""), "name_u": shape.get("NameU", ""),
                "page_name": page.get("Name") or page.get("NameU"), "page_part": page_part,
                "shape_type": shape.get("Type"), "is_group": is_group,
                "descendant_ids": [int(s.get("ID")) for s in descendants],
                "raster": foreign is not None, "foreign_type": foreign_type,
                "foreign_attributes": foreign_attrs, "source_part": source_part,
                "ole_binary": ole_binary, "prog_id": prog_id, "ole_class_name": ole_class_name,
                "binary_sha256": binary_sha256, "text": text,
                "pin_x": number(values, "PinX"), "pin_y": number(values, "PinY"),
                "width": number(values, "Width"), "height": number(values, "Height"),
                "loc_pin_x": number(values, "LocPinX"), "loc_pin_y": number(values, "LocPinY"),
                "angle": number(values, "Angle"), "flip_x": bool(number(values, "FlipX")),
                "flip_y": bool(number(values, "FlipY")), "source_cells": values,
                "source_cell_formulas": {c.get("N"): c.get("F") for c in shape.findall("v:Cell", NS) if c.get("F")},
                "asset_name": f"shape-{sid}.png" if foreign is not None else None,
            }
            objects.append(obj)
        all_shape_count = len(page_xml.findall(".//v:Shape", NS))
    return {
        "schema_version": "visio-editable-manifest.v1", "source": str(source.resolve()),
        "source_page": page.get("Name") or page.get("NameU"), "source_page_id": int(page.get("ID")),
        "source_page_part": page_part, "pages": page_list,
        "page_width_in": width, "page_height_in": height,
        "slide_width_in": output_width, "slide_height_in": output_height,
        "slide_width_pt": output_width * 72, "slide_height_pt": output_height * 72,
        "slide_width_emu": round(output_width * 914400), "slide_height_emu": round(output_height * 914400),
        "source_page_cells": page_cells, "objects": objects, "warnings": warnings,
        "requires_review": bool(warnings), "counts": {
            "top_level_shapes": len(objects), "all_shapes": all_shape_count,
            "text_shapes": sum(bool(o["text"]) for o in objects),
            "foreign_shapes": sum(o["raster"] for o in objects),
            "ole_shapes": sum(o["foreign_type"] == "Object" for o in objects),
            "known_mathtype_objects": sum(o["prog_id"] == "Equation.DSMT4" for o in objects),
            "unresolved_groups": sum(o["is_group"] for o in objects),
        },
    }


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--source", type=Path, required=True)
    parser.add_argument("--page-name")
    parser.add_argument("--output", type=Path)
    parser.add_argument("--slide-width-in", type=float)
    parser.add_argument("--list-pages", action="store_true")
    args = parser.parse_args()
    if args.list_pages:
        with zipfile.ZipFile(args.source) as archive:
            pages = ET.fromstring(archive.read("visio/pages/pages.xml"))
        print(json.dumps([{"page_id": int(p.get("ID")), "name": p.get("Name"),
                           "name_u": p.get("NameU")} for p in pages], ensure_ascii=True, indent=2))
        return
    result = inspect(args.source, args.page_name, args.slide_width_in)
    if args.output:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
        print(json.dumps({"manifest": str(args.output.resolve()), "counts": result["counts"],
                          "requires_review": result["requires_review"]}, ensure_ascii=True))
    else:
        print(json.dumps(result, ensure_ascii=True, indent=2))


if __name__ == "__main__":
    main()
