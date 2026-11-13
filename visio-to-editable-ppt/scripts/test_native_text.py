"""Focused synthetic tests: original text, styling and native coordinates."""
import json
import tempfile
import unittest
from pathlib import Path

from lxml import etree as ET

from native_text import NativeTextBuilder, UnsupportedTextError, NS


SVG = '''<svg xmlns="http://www.w3.org/2000/svg"
 xmlns:v="http://schemas.microsoft.com/visio/2003/SVGExtensions/"
 viewBox="0 0 200 100" class="page">
<style>.page {font-size:12px} .cn {font-family:SimSun;font-size:1em;fill:#123456}
.en {font-family:Times New Roman;font-weight:bold;font-size:1em}
.multi {font-family:SimSun;font-size:10px;fill:#000000}</style>
<g transform="translate(10 20)"><g id="shape7-1" transform="translate(5,3)">
<v:textRect cx="30" cy="10" width="60" height="20"/>
<text x="0" y="14" class="cn"><v:paragraph v:horizAlign="1"/>局部 <tspan class="en">VGGT</tspan> 窗口</text>
</g></g>
<g transform="translate(100,50) rotate(90)"><g id="custom-label" transform="translate(5 0)">
<v:textRect cx="10" cy="10" width="20" height="20"/>
<text x="0" y="7" class="multi"><v:paragraph v:horizAlign="1"/>第一行<v:newlineChar/><tspan x="0" dy="1.2em">第二行</tspan></text>
</g></g></svg>'''


class NativeTextTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.work = Path(self.temp.name)
        (self.work / "source.svg").write_text(SVG, encoding="utf8")
        self.manifest = {"slide_width_pt": 400, "slide_height_pt": 200, "objects": [
            {"source_id": 7, "text": "局部 VGGT 窗口"},
            {"source_id": "second", "svg_id": "custom-label", "text": "第一行\n第二行"}]}
        self.write_manifest()

    def tearDown(self):
        self.temp.cleanup()

    def write_manifest(self):
        (self.work / "manifest.json").write_text(json.dumps(self.manifest, ensure_ascii=False), encoding="utf8")

    def test_two_labels_preserve_native_text_styles_and_coordinates(self):
        builder = NativeTextBuilder(self.work)
        report = builder.write()
        self.assertEqual(report["native_textboxes"], 2)
        root = ET.parse(str(self.work / "text-map.xml"))
        self.assertEqual(len(root.findall(".//p:sp", NS)), 2)
        one = root.find("object[@source_id='7']/p:sp", NS)
        self.assertEqual("".join(one.xpath(".//a:t/text()", namespaces=NS)), "局部 VGGT 窗口")
        runs = one.findall("p:txBody/a:p/a:r", NS)
        self.assertEqual([run.find("a:t", NS).text for run in runs], ["局部 ", "VGGT", " 窗口"])
        self.assertEqual(runs[1].find("a:rPr/a:latin", NS).get("typeface"), "Times New Roman")
        self.assertEqual(runs[1].find("a:rPr", NS).get("b"), "1")
        self.assertEqual(runs[0].find("a:rPr", NS).get("sz"), "2400")
        self.assertEqual(runs[0].find("a:rPr/a:solidFill/a:srgbClr", NS).get("val"), "123456")
        # Nested translations total (15,23), canvas scale2. Box local(0,0,60,20).
        xf = one.find("p:spPr/a:xfrm", NS)
        self.assertEqual(dict(xf.find("a:off", NS).attrib), {"x": str(30*12700), "y": str(46*12700)})
        self.assertEqual(dict(xf.find("a:ext", NS).attrib), {"cx": str(120*12700), "cy": str(40*12700)})
        two = root.find("object[@source_id='second']/p:sp", NS)
        paras = two.findall("p:txBody/a:p", NS)
        self.assertEqual(["".join(p.xpath("a:r/a:t/text()", namespaces=NS)) for p in paras], ["第一行", "第二行"])
        self.assertEqual(paras[0].find("a:pPr/a:lnSpc/a:spcPts", NS).get("val"), "2400")
        # Local center(10,10), inner translate5, outer rotation90 + (100,50), scale2 => (180,130).
        xf = two.find("p:spPr/a:xfrm", NS)
        self.assertEqual(xf.get("rot"), str(90*60000))
        self.assertEqual(dict(xf.find("a:off", NS).attrib), {"x": str(160*12700), "y": str(110*12700)})
        self.assertEqual(dict(xf.find("a:ext", NS).attrib), {"cx": str(40*12700), "cy": str(40*12700)})

    def test_skew_stops_instead_of_misplacing_text(self):
        (self.work / "source.svg").write_text(SVG.replace('translate(5,3)', 'skewX(20)'), encoding="utf8")
        with self.assertRaisesRegex(UnsupportedTextError, "Unsupported text transform"):
            NativeTextBuilder(self.work).write()
        self.assertFalse((self.work / "text-map.xml").exists())

    def test_nonuniform_scaling_stops_instead_of_stretching_glyphs(self):
        (self.work / "source.svg").write_text(SVG.replace('translate(5,3)', 'scale(2,1)'), encoding="utf8")
        with self.assertRaisesRegex(UnsupportedTextError, "nonuniform text scaling"):
            NativeTextBuilder(self.work).write()

    def test_missing_internal_space_is_a_content_error(self):
        self.manifest["objects"][0]["text"] = "局部VGGT 窗口"
        self.write_manifest()
        with self.assertRaisesRegex(ValueError, "Text mismatch"):
            NativeTextBuilder(self.work).write()


if __name__ == "__main__":
    unittest.main()
