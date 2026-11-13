"""Office-free regression tests for editable-PPTX validation.

All drawings, binaries, and Office-check reports here are synthetic fixtures
in a temporary directory. Passing these tests is not a native Office render
claim. Run: python test_validate_output.py
"""
from __future__ import annotations

import copy
import hashlib
import json
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest
import zipfile

from lxml import etree as ET

import validate_output as validator

P = validator.P
A = validator.A
R = validator.R
MC = validator.MC
REL = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
NS = validator.NS


def element(parent, namespace, tag_name, **attrs):
    return ET.SubElement(parent, f"{{{namespace}}}{tag_name}", **{key: str(value) for key, value in attrs.items()})


def xml(root):
    return ET.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def write_json(path, data):
    path.write_text(json.dumps(data, indent=2), encoding="utf-8")


def hash_bytes(data):
    return hashlib.sha256(data).hexdigest()


def write_zip(path, parts):
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, data in parts.items():
            archive.writestr(name, data)


def read_zip(path):
    with zipfile.ZipFile(path) as archive:
        return {name: archive.read(name) for name in archive.namelist()}


def relationship_xml(entries):
    root = ET.Element(f"{{{REL}}}Relationships", nsmap={None: REL})
    for rid, relation_type, target in entries:
        element(root, REL, "Relationship", Id=rid, Type=R + "/" + relation_type, Target=target)
    return xml(root)


def add_nonvisual(parent, kind, sid):
    mapping = {"text": ("nvSpPr", "cNvSpPr"), "ole": ("nvGraphicFramePr", "cNvGraphicFramePr")}
    outer, inner = mapping[kind]
    properties = element(parent, P, outer)
    element(properties, P, "cNvPr", id=sid, name=f"Visio_{sid}")
    element(properties, P, inner)
    element(properties, P, "nvPr")


def ole_frame(sid, rid):
    frame = ET.Element(f"{{{P}}}graphicFrame")
    add_nonvisual(frame, "ole", sid)
    transform = element(frame, P, "xfrm")
    element(transform, A, "off", x=0, y=sid * 10000)
    element(transform, A, "ext", cx=1000000, cy=200000)
    data = element(element(frame, A, "graphic"), A, "graphicData",
                   uri="http://schemas.openxmlformats.org/presentationml/2006/ole")
    ole = element(data, P, "oleObj", showAsIcon=0, progId="Equation.DSMT4")
    ole.set(f"{{{R}}}id", rid)
    element(ole, P, "embed")
    return frame


def make_fixture(work):
    # Two distinct OLE instances intentionally contain identical bytes. This
    # detects validators that compare hash sets instead of hash multisets.
    binary = b"synthetic-equation-bytes:Equation.DSMT4:fixture-only"
    binary_hash = hash_bytes(binary)
    original = work / "original.vsdx"
    write_zip(original, {
        "visio/embeddings/oleObject1.bin": binary,
        "visio/embeddings/oleObject2.bin": binary,
        "fixture.txt": b"Synthetic source package; not intended to open in Visio.",
    })
    original_hash = validator.sha256_file(original)
    write_json(work / "conversion-context.json", {"source": str(original), "source_sha256": original_hash})
    write_json(work / "export-report.json", {
        "source": str(original), "source_sha256": original_hash, "source_unchanged": True,
    })
    objects = [{"source_id": 10, "text": "Local geometry", "raster": False}]
    binaries = []
    for sid, number in ((20, 1), (30, 2)):
        objects.append({"source_id": sid, "text": "", "raster": True,
                        "ole_binary": f"oleObject{number}.bin", "binary_sha256": binary_hash,
                        "prog_id": "Equation.DSMT4"})
        binaries.append({"source_id": sid, "sha256": binary_hash, "part": f"ppt/embeddings/ole{number}.bin"})
    write_json(work / "manifest.json", {"objects": objects})
    write_json(work / "assembly-report.json", {"labels": 1, "ole_objects": 2, "ole_binaries": binaries})
    write_json(work / "powerpoint-check.json", {
        "slides": 1, "ole_count": 2, "top_level_objects": 3,
        "fixture_only": True, "open_seconds": 0, "render_seconds": 0,
    })

    presentation = ET.Element(f"{{{P}}}presentation", nsmap={"p": P, "r": R})
    slide_ids = element(presentation, P, "sldIdLst")
    slide_id = element(slide_ids, P, "sldId", id=256)
    slide_id.set(f"{{{R}}}id", "rId1")
    element(presentation, P, "sldSz", cx=9144000, cy=6858000)
    slide = ET.Element(f"{{{P}}}sld", nsmap={"p": P, "a": A, "r": R, "mc": MC})
    tree = element(element(slide, P, "cSld"), P, "spTree")
    properties = element(tree, P, "nvGrpSpPr")
    element(properties, P, "cNvPr", id=1, name="")
    element(properties, P, "cNvGrpSpPr")
    element(properties, P, "nvPr")
    element(tree, P, "grpSpPr")
    shape = element(tree, P, "sp")
    add_nonvisual(shape, "text", 10)
    element(shape, P, "spPr")
    body = element(shape, P, "txBody")
    element(body, A, "bodyPr")
    element(body, A, "lstStyle")
    paragraph = element(body, A, "p")
    # Deliberate extra whitespace must be accepted, without relaxing wording.
    element(element(paragraph, A, "r"), A, "t").text = "Local \n geometry"
    wrapper = element(tree, MC, "AlternateContent")
    choice = element(wrapper, MC, "Choice", Requires="p")
    primary = ole_frame(20, "rId1")
    choice.append(primary)
    fallback = element(wrapper, MC, "Fallback")
    # Duplicate logical content across compatibility branches must count once.
    fallback.append(copy.deepcopy(primary))
    tree.append(ole_frame(30, "rId2"))

    types = ET.Element(f"{{{CT}}}Types", nsmap={None: CT})
    element(types, CT, "Default", Extension="rels", ContentType="application/vnd.openxmlformats-package.relationships+xml")
    element(types, CT, "Default", Extension="xml", ContentType="application/xml")
    element(types, CT, "Default", Extension="bin", ContentType="application/vnd.openxmlformats-officedocument.oleObject")
    element(types, CT, "Override", PartName="/ppt/presentation.xml",
            ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml")
    element(types, CT, "Override", PartName="/ppt/slides/slide1.xml",
            ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml")
    write_zip(work / "verified.pptx", {
        "[Content_Types].xml": xml(types),
        "_rels/.rels": relationship_xml([("rId1", "officeDocument", "ppt/presentation.xml")]),
        "ppt/presentation.xml": xml(presentation),
        "ppt/_rels/presentation.xml.rels": relationship_xml([("rId1", "slide", "slides/slide1.xml")]),
        "ppt/slides/slide1.xml": xml(slide),
        "ppt/slides/_rels/slide1.xml.rels": relationship_xml([
            ("rId1", "oleObject", "../embeddings/ole1.bin"),
            ("rId2", "oleObject", "../embeddings/ole2.bin"),
        ]),
        "ppt/embeddings/ole1.bin": binary, "ppt/embeddings/ole2.bin": binary,
    })


class EditableValidationTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(prefix="visio-editable-validation-")
        self.addCleanup(self.temp.cleanup)
        self.work = Path(self.temp.name)
        make_fixture(self.work)

    def checks(self, report):
        return {check["name"]: check for check in report["checks"]}

    def assert_check_failed(self, name):
        report = validator.validate(self.work)
        self.assertFalse(report["passed"], report)
        self.assertEqual(report["status"], "fail")
        self.assertIn(name, self.checks(report), report)
        self.assertFalse(self.checks(report)[name]["passed"], report)
        return report

    def change_text(self):
        parts = read_zip(self.work / "verified.pptx")
        slide = ET.fromstring(parts["ppt/slides/slide1.xml"])
        slide.find(".//a:t", NS).text = "Wrong label"
        parts["ppt/slides/slide1.xml"] = xml(slide)
        write_zip(self.work / "verified.pptx", parts)

    def test_complete_fixture_passes_and_compatibility_branch_counts_once(self):
        report = validator.validate(self.work)
        self.assertTrue(report["passed"], report)
        self.assertEqual(report["status"], "pass")
        self.assertEqual(report["metrics"]["top_level_objects"], 3)
        self.assertEqual(report["metrics"]["ole_objects"], 2)
        self.assertEqual(report["metrics"]["alternate_content_primary_branches"], 1)
        hashes = self.checks(report)["original_ole_binaries_preserved"]
        self.assertEqual(list(hashes["actual_sha256_multiset"].values()), [2])
        self.assertEqual(hashes["source_sha256_multiset"], hashes["actual_sha256_multiset"])

    def test_changed_editable_text_fails(self):
        self.change_text()
        self.assert_check_failed("editable_text_exact")

    def test_changed_ole_bytes_fail_even_with_valid_zip_crc(self):
        parts = read_zip(self.work / "verified.pptx")
        parts["ppt/embeddings/ole1.bin"] += b"changed"
        write_zip(self.work / "verified.pptx", parts)
        report = self.assert_check_failed("original_ole_binaries_preserved")
        self.assertTrue(self.checks(report)["zip_integrity"]["passed"])

    def test_source_manifest_hash_still_detects_rewritten_assembly_evidence(self):
        parts = read_zip(self.work / "verified.pptx")
        parts["ppt/embeddings/ole1.bin"] += b"changed"
        write_zip(self.work / "verified.pptx", parts)
        assembly_path = self.work / "assembly-report.json"
        assembly = validator.read_json(assembly_path)
        assembly["ole_binaries"][0]["sha256"] = hash_bytes(parts["ppt/embeddings/ole1.bin"])
        write_json(assembly_path, assembly)
        report = self.assert_check_failed("original_ole_binaries_preserved")
        check = self.checks(report)["original_ole_binaries_preserved"]
        self.assertEqual(check["expected_sha256_multiset"], check["actual_sha256_multiset"])
        self.assertNotEqual(check["source_sha256_multiset"], check["actual_sha256_multiset"])

    def test_changed_original_source_fails(self):
        with (self.work / "original.vsdx").open("ab") as stream:
            stream.write(b"source changed")
        self.assert_check_failed("original_source_unchanged")

    def test_missing_internal_target_fails(self):
        parts = read_zip(self.work / "verified.pptx")
        del parts["ppt/embeddings/ole1.bin"]
        write_zip(self.work / "verified.pptx", parts)
        self.assert_check_failed("internal_relationships")

    def test_cli_writes_success_then_failure_report_with_exit_status(self):
        command = [sys.executable, str(Path(validator.__file__).resolve()), "--work-dir", str(self.work)]
        passed = subprocess.run(command, capture_output=True, text=True)
        self.assertEqual(passed.returncode, 0, passed.stdout + passed.stderr)
        self.assertEqual(validator.read_json(self.work / "validation.json")["status"], "pass")
        self.change_text()
        failed = subprocess.run(command, capture_output=True, text=True)
        self.assertEqual(failed.returncode, 1, failed.stdout + failed.stderr)
        self.assertEqual(validator.read_json(self.work / "validation.json")["status"], "fail")


if __name__ == "__main__":
    unittest.main(verbosity=2)
