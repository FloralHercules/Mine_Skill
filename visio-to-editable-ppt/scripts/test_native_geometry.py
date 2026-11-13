"""Synthetic fixtures only: no user diagram, coordinates, or extracted media."""
import json
from pathlib import Path
import tempfile
import unittest

from lxml import etree as ET
from native_geometry import Converter, NS, SVG, convert_work_dir, main


FIXTURE = f'''<svg xmlns="{SVG}" width="4in" height="2in" viewBox="0 0 400 200">
 <style>.outline{{fill:none;stroke:#94a3b8;stroke-width:2;stroke-dasharray:8,4}}</style>
 <g id="shape11-1"><rect x="10" y="10" width="120" height="70" rx="8" class="outline"/></g>
 <g id="shape12-2"><line x1="150" y1="30" x2="220" y2="60" stroke="#123456" stroke-width="2"/></g>
 <g id="shape13-3" transform="translate(15,10)"><ellipse cx="270" cy="70" rx="35" ry="20" fill="#00aaff" stroke="#000000"/></g>
 <g id="shape14-4"><path d="M 300 120 L 330 140 L 300 160 L 270 140 Z" fill="#ffeeaa" stroke="#555555"/></g>
</svg>'''


def manifest():
    return {"slide_width_pt": 288, "slide_height_pt": 144, "page_width_in": 4,
            "objects": [{"source_id": sid, "raster": False, "source_cells": {}}
                        for sid in (11, 12, 13, 14)]}


class NativeGeometryTests(unittest.TestCase):
    def convert_fixture(self, data=None, overrides=None, svg=FIXTURE):
        return Converter(ET.fromstring(svg.encode()), data or manifest(), overrides).convert()

    def test_simple_shapes_remain_native_and_dash_does_not_explode(self):
        root, report = self.convert_fixture()
        self.assertEqual(report["unresolved"], [])
        self.assertEqual(report["native_geometry_shapes"], 4)
        self.assertEqual(report["custom_paths"], 0)
        self.assertEqual(report["geometry_commands"], 0)
        self.assertEqual([n.get("prst") for n in root.findall(".//a:prstGeom", NS)],
                         ["roundRect", "line", "ellipse", "diamond"])
        dash = root.find("object/p:sp/p:spPr/a:ln/a:custDash/a:ds", NS)
        self.assertEqual(dash.attrib, {"d": "400000", "sp": "200000"})
        ellipse = root.find("object[@source_id='13']/p:sp/p:spPr/a:xfrm/a:off", NS)
        self.assertEqual(int(ellipse.get("x")), round(250*288*12700/400))

    def test_glow_reads_generic_source_cells(self):
        data = manifest()
        data["objects"][0]["source_cells"] = {"GlowSize": "0.1", "GlowColor": "#ffcc00", "GlowColorTrans": "0.4"}
        root, report = self.convert_fixture(data)
        self.assertFalse(report["unresolved"])
        glow = root.find("object/p:sp/p:spPr/a:effectLst/a:glow", NS)
        self.assertEqual(int(glow.get("rad")), 91440)
        self.assertEqual(glow.find("a:srgbClr/a:alpha", NS).get("val"), "60000")

    def test_3d_rotation_fails_until_explicit_projection_supplied(self):
        data = manifest()
        data["objects"][0]["source_cells"] = {"RotationXAngle": "0.3", "RotationType": "1"}
        root, report = self.convert_fixture(data)
        self.assertEqual(root.get("status"), "unresolved")
        self.assertIn("geometry-overrides.json", report["unresolved"][0]["action"])
        override = {"11": {"points": [[10, 15], [100, 20], [90, 60], [5, 55]],
                           "closed": True, "style": {"fill": "#808080", "fill-opacity": "0.2"}}}
        fixed, report = self.convert_fixture(data, override)
        self.assertEqual(report["unresolved"], [])
        self.assertEqual(len(fixed.find("object/p:sp/p:spPr/a:custGeom/a:pathLst/a:path", NS)), 5)

    def test_unknown_visible_svg_is_reported_not_dropped(self):
        svg = FIXTURE.replace('<rect x="10"', '<use href="#missing"/><rect x="10"')
        _, report = self.convert_fixture(svg=svg)
        self.assertEqual(report["unresolved"][0]["source_id"], "11")
        self.assertIn("Unsupported visible SVG <use>", report["unresolved"][0]["action"])

    def test_flattened_nested_manifest_does_not_duplicate_child(self):
        svg = f'<svg xmlns="{SVG}" width="4in" viewBox="0 0 400 200"><g id="group20-1" transform="translate(10,10)"><g id="shape21-2"><rect x="5" y="5" width="30" height="20" fill="#00ff00"/></g></g></svg>'
        data = {"slide_width_pt": 288, "slide_height_pt": 144,
                "objects": [{"source_id": 20}, {"source_id": 21}]}
        _, report = self.convert_fixture(data, svg=svg)
        self.assertEqual(report["by_source"], {"20": 0, "21": 1})

    def test_cli_outputs_and_nonzero_failure_status(self):
        with tempfile.TemporaryDirectory() as temporary:
            work = Path(temporary)
            (work/"source.svg").write_text(FIXTURE, encoding="utf8")
            (work/"manifest.json").write_text(json.dumps(manifest()), encoding="utf8")
            report = convert_work_dir(work)
            self.assertEqual(report["status"], "ready")
            self.assertTrue((work/"geometry-map.xml").exists())
            data = manifest()
            data["objects"][1]["source_cells"] = {"RotationYAngle": 1}
            (work/"manifest.json").write_text(json.dumps(data), encoding="utf8")
            self.assertEqual(main(["--work-dir", str(work)]), 2)


if __name__ == "__main__":
    unittest.main()
