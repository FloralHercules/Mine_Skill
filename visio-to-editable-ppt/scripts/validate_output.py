"""Validate the PowerPoint-normalized editable result and its source evidence.

Usage: python validate_output.py --work-dir WORK
Reports structural/content evidence in WORK/validation.json. This does not
replace the required visual comparison of the original and rendered previews.
"""
from __future__ import annotations

import argparse
from collections import Counter
import copy
import hashlib
import json
import posixpath
import re
import sys
from pathlib import Path
from urllib.parse import unquote, urlsplit
import zipfile

from lxml import etree as ET

P = "http://schemas.openxmlformats.org/presentationml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
NS = {"p": P, "a": A, "r": R, "mc": MC}
SHAPE_TAGS = {f"{{{P}}}{name}" for name in ("sp", "cxnSp", "grpSp", "graphicFrame", "pic", "contentPart")}
NONVISUAL = {
    "sp": "p:nvSpPr/p:cNvPr", "cxnSp": "p:nvCxnSpPr/p:cNvPr",
    "grpSp": "p:nvGrpSpPr/p:cNvPr", "graphicFrame": "p:nvGraphicFramePr/p:cNvPr",
    "pic": "p:nvPicPr/p:cNvPr",
}


def sha256_file(path):
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def read_json(path):
    return json.loads(path.read_text(encoding="utf-8-sig"))


def normalize_text(value):
    return re.sub(r"\s+", "", value)


def relationship_owner(part):
    if part == "_rels/.rels":
        return ""
    folder, filename = posixpath.split(part)
    if posixpath.basename(folder) != "_rels" or not filename.endswith(".rels"):
        raise ValueError(f"Invalid relationship part path: {part}")
    return posixpath.join(posixpath.dirname(folder), filename[:-5])


def resolve_target(owner, target):
    parsed = urlsplit(target)
    if parsed.scheme or parsed.netloc:
        raise ValueError("Internal relationship has an external URI")
    path = unquote(parsed.path)
    if not path:
        resolved = owner
    elif path.startswith("/"):
        resolved = posixpath.normpath(path.lstrip("/"))
    else:
        resolved = posixpath.normpath(posixpath.join(posixpath.dirname(owner), path))
    if resolved == ".." or resolved.startswith("../"):
        raise ValueError("Relationship escapes the package root")
    return resolved


def effective_xml(original):
    """Use one primary branch of each MC wrapper, never Choice plus Fallback.

    Input is saved by the same desktop PowerPoint that produced the check
    report. Its first Choice is the primary representation; Fallback is used
    only when no Choice exists. Report the number of wrappers handled.
    """
    root = copy.deepcopy(original)
    count = 0
    while True:
        wrappers = root.xpath(".//mc:AlternateContent", namespaces=NS)
        if not wrappers:
            break
        wrapper = wrappers[0]
        branches = wrapper.findall(f"{{{MC}}}Choice")
        branch = branches[0] if branches else wrapper.find(f"{{{MC}}}Fallback")
        if branch is None:
            raise ValueError("MC AlternateContent has no Choice or Fallback")
        parent = wrapper.getparent()
        index = parent.index(wrapper)
        children = list(branch)
        parent.remove(wrapper)
        for offset, child in enumerate(children):
            parent.insert(index + offset, child)
        count += 1
    return root, count


def shape_name(shape):
    local = ET.QName(shape).localname
    selector = NONVISUAL.get(local)
    entry = shape.find(selector, NS) if selector else None
    return entry.get("name", "") if entry is not None else ""


def validate(work):
    report = {
        "schema_version": "visio-editable-validation.v1", "work_dir": str(work),
        "verified_pptx": str(work / "verified.pptx"), "status": "fail", "passed": False,
        "checks": [], "metrics": {},
        "claim_boundary": "Package, source integrity, editable text, OLE, and desktop-object checks only; visual review is still required.",
        "visual_review_required": True,
    }

    def check(name, passed, **details):
        report["checks"].append({"name": name, "passed": bool(passed), **details})
        return bool(passed)

    required = ("verified.pptx", "manifest.json", "assembly-report.json", "powerpoint-check.json",
                "conversion-context.json", "export-report.json")
    missing = [name for name in required if not (work / name).is_file()]
    if not check("required_evidence_files", not missing, missing=missing):
        return report
    try:
        manifest = read_json(work / "manifest.json")
        assembly = read_json(work / "assembly-report.json")
        powerpoint = read_json(work / "powerpoint-check.json")
        context = read_json(work / "conversion-context.json")
        exported = read_json(work / "export-report.json")
        source_path = Path(context["source"])
        expected_source_hash = context["source_sha256"].lower()
        actual_source_hash = sha256_file(source_path)
        check("original_source_unchanged", actual_source_hash == expected_source_hash,
              source=str(source_path), expected_sha256=expected_source_hash, actual_sha256=actual_source_hash)
        check("source_evidence_consistent",
              exported.get("source_sha256", "").lower() == expected_source_hash
              and Path(exported["source"]).resolve() == source_path.resolve()
              and exported.get("source_unchanged") is True,
              export_sha256=exported.get("source_sha256"), context_sha256=expected_source_hash)
        report["verified_sha256"] = sha256_file(work / "verified.pptx")
        report["verified_byte_count"] = (work / "verified.pptx").stat().st_size

        with zipfile.ZipFile(work / "verified.pptx") as archive:
            names = archive.namelist()
            duplicates = [name for name, count in Counter(names).items() if count != 1]
            crc_bad = archive.testzip()
            check("zip_integrity", crc_bad is None and not duplicates,
                  crc_failure=crc_bad, duplicate_parts=duplicates)
            names_set = set(names)
            xml_parts = {}
            xml_errors = []
            for name in names:
                if name.endswith((".xml", ".rels")):
                    try:
                        xml_parts[name] = ET.fromstring(archive.read(name), ET.XMLParser(resolve_entities=False))
                    except ET.XMLSyntaxError as error:
                        xml_errors.append({"part": name, "error": str(error)})
            if not check("xml_well_formed", not xml_errors, errors=xml_errors):
                return report
            relationships = {}
            rel_errors = []
            external_count = 0
            internal_count = 0
            for name, root in xml_parts.items():
                if not name.endswith(".rels"):
                    continue
                owner = relationship_owner(name)
                mapping = {}
                if owner and owner not in names_set:
                    rel_errors.append({"part": name, "error": "Relationship owner is absent", "owner": owner})
                for rel in root:
                    rid = rel.get("Id")
                    if not rid or rid in mapping:
                        rel_errors.append({"part": name, "id": rid, "error": "Missing or duplicate relationship ID"})
                    item = dict(rel.attrib)
                    mapping[rid] = item
                    if rel.get("TargetMode") == "External":
                        external_count += 1
                        continue
                    internal_count += 1
                    try:
                        resolved = resolve_target(owner, rel.get("Target", ""))
                        item["resolved"] = resolved
                        if resolved not in names_set:
                            rel_errors.append({"part": name, "id": rid, "target": rel.get("Target"),
                                               "resolved": resolved, "error": "Internal target is absent"})
                    except ValueError as error:
                        rel_errors.append({"part": name, "id": rid, "error": str(error)})
                relationships[owner] = mapping
            check("internal_relationships", not rel_errors, internal_count=internal_count,
                  external_count=external_count, errors=rel_errors)

            presentation_part = "ppt/presentation.xml"
            presentation = xml_parts[presentation_part]
            slides = presentation.findall("p:sldIdLst/p:sldId", NS)
            slide_parts = []
            for slide in slides:
                relationship = relationships[presentation_part].get(slide.get(f"{{{R}}}id"), {})
                if not relationship.get("Type", "").endswith("/slide") or not relationship.get("resolved"):
                    raise ValueError("Presentation references a missing or invalid slide relationship")
                slide_parts.append(relationship["resolved"])
            if not check("single_slide", len(slide_parts) == 1, slide_count=len(slide_parts), parts=slide_parts):
                return report
            slide_part = slide_parts[0]
            slide, mc_count = effective_xml(xml_parts[slide_part])
            tree = slide.find("p:cSld/p:spTree", NS)
            if tree is None:
                raise ValueError("Slide has no shape tree")
            top_shapes = [el for el in tree if el.tag in SHAPE_TAGS]
            named = {}
            for shape in top_shapes:
                named.setdefault(shape_name(shape), []).append(shape)
            objects = manifest["objects"]
            names_expected = [f"Visio_{obj['source_id']}" for obj in objects]
            object_issues = []
            for name in names_expected:
                if len(named.get(name, [])) != 1:
                    object_issues.append({"name": name, "actual_count": len(named.get(name, []))})
            unexpected = [name for name in named if name not in set(names_expected)]
            check("source_object_coverage", not object_issues and not unexpected,
                  expected=len(objects), actual=len(top_shapes), missing_or_duplicate=object_issues,
                  unexpected=unexpected)

            label_checks = []
            for obj in objects:
                if not obj.get("text"):
                    continue
                name = f"Visio_{obj['source_id']}"
                found = named.get(name, [])
                actual = "".join(found[0].xpath(".//a:t/text()", namespaces=NS)) if len(found) == 1 else None
                matched = actual is not None and normalize_text(actual) == normalize_text(obj["text"])
                label_checks.append({"source_id": obj["source_id"], "matched": matched,
                                     **({} if matched else {"expected": obj["text"], "actual": actual})})
            label_failures = [result for result in label_checks if not result["matched"]]
            check("editable_text_exact", not label_failures and assembly.get("labels") == len(label_checks),
                  expected_labels=len(label_checks), matched_labels=len(label_checks) - len(label_failures),
                  assembly_labels=assembly.get("labels"), failures=label_failures,
                  normalization="Whitespace only; wording and punctuation are preserved exactly.")

            expected_hashes = Counter(item["sha256"].lower() for item in assembly["ole_binaries"])
            manifest_hashes = Counter(obj["binary_sha256"].lower() for obj in objects if obj.get("ole_binary") and obj.get("binary_sha256"))
            actual_hashes = Counter()
            ole_errors = []
            ole_objects = slide.findall(".//p:oleObj", NS)
            actual_ole_parts = []
            for ole in ole_objects:
                rid = ole.get(f"{{{R}}}id")
                relationship = relationships.get(slide_part, {}).get(rid, {})
                part = relationship.get("resolved")
                if not relationship.get("Type", "").endswith("/oleObject") or not part:
                    ole_errors.append({"relationship": rid, "error": "Embedded OLE relationship is missing or invalid"})
                    continue
                if ole.find("p:embed", NS) is None:
                    ole_errors.append({"relationship": rid, "error": "OLE object is linked rather than embedded"})
                    continue
                binary_hash = hashlib.sha256(archive.read(part)).hexdigest()
                actual_hashes[binary_hash] += 1
                actual_ole_parts.append({"part": part, "sha256": binary_hash, "prog_id": ole.get("progId")})
            expected_ole = sum(bool(obj.get("ole_binary")) for obj in objects)
            check("original_ole_binaries_preserved", not ole_errors and actual_hashes == expected_hashes == manifest_hashes
                  and len(ole_objects) == expected_ole == assembly.get("ole_objects"),
                  expected_objects=expected_ole, actual_objects=len(ole_objects),
                  expected_sha256_multiset=dict(expected_hashes), actual_sha256_multiset=dict(actual_hashes),
                  source_sha256_multiset=dict(manifest_hashes),
                  failures=ole_errors, objects=actual_ole_parts,
                  alternate_content_wrappers_handled=mc_count)

            check("powerpoint_object_counts", powerpoint.get("slides") == 1
                  and powerpoint.get("top_level_objects") == len(top_shapes) == len(objects)
                  and powerpoint.get("ole_count") == len(ole_objects) == expected_ole,
                  powerpoint=powerpoint, package_top_level_objects=len(top_shapes),
                  package_ole_objects=len(ole_objects), manifest_objects=len(objects))
            native_shapes = slide.findall(".//p:sp", NS)
            text_shapes = [shape for shape in native_shapes if shape.xpath(".//a:t/text()", namespaces=NS)]
            report["metrics"] = {
                "native_shapes": len(native_shapes), "native_text_shapes": len(text_shapes),
                "groups": len(slide.findall(".//p:grpSp", NS)),
                "native_connectors": len(slide.findall(".//p:cxnSp", NS)),
                "top_level_objects": len(top_shapes), "ole_objects": len(ole_objects),
                "path_commands": sum(len(slide.findall(".//a:" + op, NS)) for op in
                                     ("moveTo", "lnTo", "cubicBezTo", "quadBezTo", "arcTo")),
                "slide_xml_bytes": len(archive.read(slide_part)),
                "geometry_complexity_policy": "Reported only; no arbitrary complexity threshold.",
                "alternate_content_primary_branches": mc_count,
            }
    except Exception as error:
        check("validation_execution", False, exception_type=type(error).__name__, error=str(error))
    report["passed"] = bool(report["checks"]) and all(entry["passed"] for entry in report["checks"])
    report["status"] = "pass" if report["passed"] else "fail"
    return report


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--work-dir", type=Path, required=True)
    args = parser.parse_args()
    work = args.work_dir.resolve()
    report = validate(work)
    work.mkdir(parents=True, exist_ok=True)
    destination = work / "validation.json"
    destination.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({"status": report["status"], "report": str(destination),
                      "failed_checks": [c["name"] for c in report["checks"] if not c["passed"]],
                      "metrics": report["metrics"]}, ensure_ascii=True))
    if not report["passed"]:
        raise SystemExit(1)


if __name__ == "__main__":
    main()
