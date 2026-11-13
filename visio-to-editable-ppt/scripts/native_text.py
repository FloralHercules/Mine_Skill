"""Convert Visio SVG text into one editable PowerPoint text box per label.

Inputs in --work-dir: source.svg and manifest.json with slide_width_pt,
slide_height_pt, and objects [{source_id, text, svg_id?}].  SVG groups normally
use Visio's shape<ID>-<export sequence> IDs. An explicit svg_id also works.

This module emits Open XML fragments, not a standalone PPTX. The assembler must
allocate unique final shape IDs. Unsupported text geometry raises an actionable
error; it never rasterizes or silently approximates the label.
"""
from __future__ import annotations

import argparse
import json
import math
import re
import unicodedata
from pathlib import Path

from lxml import etree as ET

S = "http://www.w3.org/2000/svg"
V = "http://schemas.microsoft.com/visio/2003/SVGExtensions/"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS = {"s": S, "v": V, "a": A, "p": P}
EMU_PER_PT = 12700
IDENTITY = (1.0, 0.0, 0.0, 1.0, 0.0, 0.0)
NUMBER = r"[-+]?(?:\d*\.\d+|\d+\.?\d*)(?:[eE][-+]?\d+)?"


class UnsupportedTextError(ValueError):
    """Source text needs a deliberate conversion strategy before proceeding."""


def multiply(left, right):
    """Composition of SVG affine matrices: left(right(point))."""
    a, b, c, d, e, f = left
    g, h, i, j, k, ell = right
    return (a*g+c*h, b*g+d*h, a*i+c*j, b*i+d*j, a*k+c*ell+e, b*k+d*ell+f)


def transform_point(matrix, x, y):
    a, b, c, d, e, f = matrix
    return a*x+c*y+e, b*x+d*y+f


def parse_transform(value):
    matrix = IDENTITY
    cursor = 0
    for match in re.finditer(r"([A-Za-z]+)\s*\(([^)]*)\)", value or ""):
        if (value or "")[cursor:match.start()].strip(" ,\t\r\n"):
            raise UnsupportedTextError(f"Invalid transform syntax: {value!r}")
        cursor = match.end()
        name, args = match.groups()
        tokens = re.findall(NUMBER, args)
        if re.sub(NUMBER, "", args).strip(" ,\t\r\n"):
            raise UnsupportedTextError(f"Invalid transform arguments: {value!r}")
        nums = list(map(float, tokens))
        if name == "matrix" and len(nums) == 6:
            operation = tuple(nums)
        elif name == "translate" and len(nums) in (1, 2):
            operation = (1, 0, 0, 1, nums[0], nums[1] if len(nums) == 2 else 0)
        elif name == "scale" and len(nums) in (1, 2):
            operation = (nums[0], 0, 0, nums[-1], 0, 0)
        elif name == "rotate" and len(nums) in (1, 3):
            radians = math.radians(nums[0])
            co, si = math.cos(radians), math.sin(radians)
            operation = (co, si, -si, co, 0, 0)
            if len(nums) == 3:
                cx, cy = nums[1:]
                operation = multiply((1, 0, 0, 1, cx, cy),
                                     multiply(operation, (1, 0, 0, 1, -cx, -cy)))
        else:
            raise UnsupportedTextError(f"Unsupported text transform {name!r}: {value!r}")
        matrix = multiply(matrix, operation)
    if (value or "")[cursor:].strip(" ,\t\r\n"):
        raise UnsupportedTextError(f"Invalid transform syntax: {value!r}")
    if not all(math.isfinite(x) for x in matrix):
        raise UnsupportedTextError("Text transform contains a nonfinite number")
    return matrix


def full_transform(node):
    matrix = IDENTITY
    for parent in list(node.iterancestors())[::-1] + [node]:
        matrix = multiply(matrix, parse_transform(parent.get("transform", "")))
    return matrix


def normalize_text(value):
    """Ignore layout line breaks only; preserve meaningful intra-line spaces."""
    return unicodedata.normalize("NFC", value).replace("\r", "").replace("\n", "")


def el(parent, prefix, local_name, **attrs):
    namespace = {"a": A, "p": P}[prefix]
    return ET.SubElement(parent, f"{{{namespace}}}{local_name}",
                         {key: str(value) for key, value in attrs.items()})


def numeric(value, description):
    if not re.fullmatch(NUMBER, value or ""):
        raise UnsupportedTextError(f"Expected a scalar {description}, found {value!r}")
    result = float(value)
    if not math.isfinite(result):
        raise UnsupportedTextError(f"Nonfinite {description}")
    return result


def color_hex(value):
    named = {"black": "000000", "white": "FFFFFF", "red": "FF0000", "blue": "0000FF"}
    value = value.strip()
    if value.lower() in named:
        return named[value.lower()]
    if re.fullmatch(r"#[0-9a-fA-F]{3}", value):
        return "".join(x*2 for x in value[1:]).upper()
    if re.fullmatch(r"#[0-9a-fA-F]{6}", value):
        return value[1:].upper()
    match = re.fullmatch(r"rgb\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)", value)
    if match and all(int(c) <= 255 for c in match.groups()):
        return "".join(f"{int(c):02X}" for c in match.groups())
    raise UnsupportedTextError(f"Unsupported text color {value!r}")


class NativeTextBuilder:
    def __init__(self, work_dir, *, manifest_name="manifest.json", svg_name="source.svg"):
        self.work_dir = Path(work_dir)
        self.manifest = json.loads((self.work_dir / manifest_name).read_text(encoding="utf-8-sig"))
        self.root = ET.parse(str(self.work_dir / svg_name),
                             ET.XMLParser(huge_tree=True, resolve_entities=False, no_network=True)).getroot()
        self.objects = {}
        for obj in self.manifest["objects"]:
            if obj.get("text"):
                key = str(obj["source_id"])
                if key in self.objects:
                    raise ValueError(f"Duplicate text source_id {key}")
                self.objects[key] = obj
        width = float(self.manifest["slide_width_pt"])
        height = float(self.manifest["slide_height_pt"])
        values = re.split(r"[\s,]+", self.root.get("viewBox", "").strip())
        if len(values) != 4:
            raise UnsupportedTextError("Visio SVG requires an explicit four-number viewBox")
        min_x, min_y, vb_width, vb_height = map(float, values)
        if min(width, height, vb_width, vb_height) <= 0:
            raise ValueError("Slide and viewBox dimensions must be positive")
        self.canvas = (width/vb_width, 0, 0, height/vb_height,
                       -min_x*width/vb_width, -min_y*height/vb_height)
        self.groups = {}
        self.by_id = {node.get("id"): node for node in self.root.iter() if node.get("id")}
        for node in self.root.iter(f"{{{S}}}g"):
            match = re.fullmatch(r"shape(\d+)-\d+", node.get("id", ""))
            if match:
                self.groups[match.group(1)] = node
        self.css = {}
        for style in self.root.findall(".//s:style", NS):
            for selector, declarations in re.findall(r"([^{}]+)\{([^}]+)\}", style.text or ""):
                # Visio exports resolved class styles. Other selectors may not
                # target text predictably, so require the caller to resolve them.
                for simple in selector.strip().split(","):
                    if re.fullmatch(r"\.[\w-]+", simple.strip()):
                        self.css[simple.strip()[1:]] = self.declarations(declarations)
                    else:
                        raise UnsupportedTextError(f"Unsupported SVG CSS selector {simple.strip()!r}")

    @staticmethod
    def declarations(value):
        return {k.strip(): v.strip() for k, v in
                (item.split(":", 1) for item in value.split(";") if ":" in item)}

    def style(self, node, inherited=None):
        result = dict(inherited or {"font-size-value": 16.0, "font-family": "Times New Roman",
                                   "fill": "#000000", "font-weight": "normal", "opacity": "1"})
        parent_size = result["font-size-value"]
        properties = {}
        for name in node.get("class", "").split():
            if name not in self.css:
                raise UnsupportedTextError(f"Unresolved SVG style class {name!r}")
            properties.update(self.css[name])
        properties.update({k: v for k, v in node.attrib.items() if k in
                           ("font-family", "font-size", "font-weight", "font-style", "fill",
                            "fill-opacity", "opacity", "text-decoration", "letter-spacing",
                            "writing-mode", "baseline-shift", "text-anchor")})
        properties.update(self.declarations(node.get("style", "")))
        for key, value in properties.items():
            if key == "font-size":
                match = re.fullmatch(f"({NUMBER})(em|px|pt|%)?", value)
                if not match:
                    raise UnsupportedTextError(f"Unsupported font size {value!r}")
                number = float(match.group(1))
                unit = match.group(2)
                result["font-size-value"] = number * (parent_size if unit == "em" else
                                                       parent_size/100 if unit == "%" else 1)
            elif key == "opacity":
                result[key] = str(float(inherited.get("opacity", 1) if inherited else 1) * float(value))
            else:
                result[key] = value
        for name in ("letter-spacing", "baseline-shift"):
            if result.get(name, "normal") not in ("normal", "0", "0px", "baseline"):
                raise UnsupportedTextError(f"Unsupported text {name}={result[name]!r}")
        if result.get("writing-mode", "horizontal-tb") not in ("horizontal-tb", "lr", "lr-tb"):
            raise UnsupportedTextError("Vertical text requires explicit conversion")
        if result.get("text-decoration", "none") not in ("none", "underline", "line-through"):
            raise UnsupportedTextError("Unsupported text decoration")
        return result

    def inherited_style(self, node):
        result = None
        for ancestor in list(node.iterancestors())[::-1]:
            result = self.style(ancestor, result)
        return result

    def parse_lines(self, text):
        inherited = self.inherited_style(text)
        base = self.style(text, inherited)
        lines = [[]]
        steps = []

        def append(value, style):
            if value:
                lines[-1].append((value, dict(style)))

        def newline(step=None):
            if lines[-1]:
                lines.append([])
                steps.append(step)
            elif step is not None and steps:
                steps[-1] = step

        def visit(node, parent_style):
            current = self.style(node, parent_style)
            if node is not text:
                if node.get("transform") or node.get("rotate") or node.get("textLength"):
                    raise UnsupportedTextError("Individually transformed text spans are unsupported")
                if node.get("dx") not in (None, "0", "0px"):
                    raise UnsupportedTextError("Explicit inline glyph offsets need manual reconstruction")
                if node.get("x") is not None:
                    numeric(node.get("x"), "tspan x")
                    dy = node.get("dy")
                    if dy:
                        match = re.fullmatch(f"({NUMBER})(em|px|pt)?", dy)
                        if not match:
                            raise UnsupportedTextError(f"Unsupported line spacing {dy!r}")
                        step = float(match.group(1)) * (parent_style["font-size-value"]
                                                       if match.group(2) == "em" else 1)
                        newline(step)
                    elif node.get("y"):
                        raise UnsupportedTextError("Absolute tspan baselines require explicit line-spacing conversion")
                    elif lines[-1]:
                        raise UnsupportedTextError("Inline absolute tspan x cannot be represented as ordinary rich text")
                elif node.get("dy") not in (None, "0", "0px"):
                    raise UnsupportedTextError("Inline baseline shifts are unsupported")
            append(node.text, current)
            for child in node:
                if child.tag == f"{{{S}}}tspan":
                    visit(child, current)
                elif child.tag == f"{{{V}}}newlineChar":
                    newline()
                elif child.tag not in (f"{{{V}}}paragraph", f"{{{V}}}tabList"):
                    raise UnsupportedTextError(f"Unsupported SVG text child {child.tag!r}")
                append(child.tail, current)

        if text.get("rotate") or text.get("textLength"):
            raise UnsupportedTextError("Per-glyph text rotation or textLength stretching is unsupported")
        for coordinate in ("x", "y"):
            if text.get(coordinate) is not None:
                numeric(text.get(coordinate), f"text {coordinate}")
        visit(text, inherited)
        return [line for line in lines if line], steps, base

    def build(self, source_id):
        key = str(source_id)
        source = self.objects[key]
        group = self.by_id.get(source.get("svg_id")) if source.get("svg_id") else self.groups.get(key)
        if group is None:
            raise UnsupportedTextError(f"Missing SVG group for source_id={key}")
        texts = group.findall("s:text", NS)
        rect = group.find("v:textRect", NS)
        if len(texts) != 1 or rect is None:
            raise UnsupportedTextError(f"Source {key} needs exactly one SVG text and Visio textRect")
        text = texts[0]
        matrix = multiply(self.canvas, full_transform(text))
        a, b, c, d, _, _ = matrix
        sx, sy = math.hypot(a, b), math.hypot(c, d)
        if min(sx, sy) < 1e-12 or a*d-b*c <= 0:
            raise UnsupportedTextError(f"Source {key} has singular or reflected text transform")
        if abs(a*c+b*d) > sx*sy*1e-7:
            raise UnsupportedTextError(f"Source {key} has skewed text geometry")
        # Exporters round viewBox dimensions. Tolerate rounding, not visible
        # nonuniform glyph stretching, which PowerPoint cannot reproduce here.
        if abs(sx-sy) > max(sx, sy)*1e-4:
            raise UnsupportedTextError(f"Source {key} has nonuniform text scaling ({sx:g}, {sy:g})")
        cx, cy = transform_point(matrix, numeric(rect.get("cx"), "textRect cx"),
                                 numeric(rect.get("cy"), "textRect cy"))
        width = numeric(rect.get("width"), "textRect width")*sx
        height = numeric(rect.get("height"), "textRect height")*sy
        if min(width, height) <= 0:
            raise UnsupportedTextError(f"Source {key} has empty text bounds")
        lines, steps, base = self.parse_lines(text)
        actual = "".join(value for line in lines for value, _ in line)
        if normalize_text(actual) != normalize_text(source["text"]):
            raise ValueError(f"Text mismatch for source {key}: {source['text']!r} versus {actual!r}")
        shape = ET.Element(f"{{{P}}}sp", nsmap={"p": P, "a": A})
        nv = el(shape, "p", "nvSpPr")
        el(nv, "p", "cNvPr", id=1, name=f"Visio_{key}_Text", descr=source["text"])
        el(nv, "p", "cNvSpPr", txBox=1)
        el(nv, "p", "nvPr")
        props = el(shape, "p", "spPr")
        angle = math.degrees(math.atan2(b, a)) % 360
        xfrm = el(props, "a", "xfrm", **({"rot": round(angle*60000)} if abs(angle) > 1e-7 else {}))
        el(xfrm, "a", "off", x=round((cx-width/2)*EMU_PER_PT), y=round((cy-height/2)*EMU_PER_PT))
        el(xfrm, "a", "ext", cx=round(width*EMU_PER_PT), cy=round(height*EMU_PER_PT))
        el(el(props, "a", "prstGeom", prst="rect"), "a", "avLst")
        el(props, "a", "noFill")
        el(el(props, "a", "ln"), "a", "noFill")
        body = el(shape, "p", "txBody")
        anchor = {"top": "t", "center": "ctr", "bottom": "b"}.get(source.get("vertical_align", "center"))
        if anchor is None:
            raise UnsupportedTextError(f"Unsupported vertical_align for source {key}")
        body_pr = el(body, "a", "bodyPr", wrap="none", lIns=0, tIns=0, rIns=0, bIns=0,
                     anchor=anchor, anchorCtr=0, vert="horz")
        el(body_pr, "a", "noAutofit")
        el(body, "a", "lstStyle")
        paragraph = text.find("v:paragraph", NS)
        align_value = paragraph.get(f"{{{V}}}horizAlign", "1") if paragraph is not None else "1"
        align = {"0": "l", "1": "ctr", "2": "r"}.get(align_value)
        if align is None:
            raise UnsupportedTextError(f"Unsupported paragraph alignment {align_value!r}")
        for index, line in enumerate(lines):
            para = el(body, "a", "p")
            para_pr = el(para, "a", "pPr", algn=align, marL=0, marR=0, indent=0, fontAlgn="base")
            if len(lines) > 1:
                step = (steps[index] if index < len(steps) and steps[index] is not None else
                        next((step for step in reversed(steps) if step), base["font-size-value"]*1.2))
                if step <= 0:
                    raise UnsupportedTextError("Non-positive text line spacing")
                el(el(para_pr, "a", "lnSpc"), "a", "spcPts", val=round(step*sy*100))
            el(el(para_pr, "a", "spcBef"), "a", "spcPts", val=0)
            el(el(para_pr, "a", "spcAft"), "a", "spcPts", val=0)
            el(para_pr, "a", "buNone")
            for value, style in line:
                run = el(para, "a", "r")
                run_pr = el(run, "a", "rPr", lang=source.get("lang", "zh-CN"), altLang="en-US",
                            sz=round(style["font-size-value"]*sy*100),
                            b=int(style.get("font-weight") in ("bold", "700", "800", "900")),
                            i=int(style.get("font-style") == "italic"))
                if style.get("text-decoration") == "underline":
                    run_pr.set("u", "sng")
                elif style.get("text-decoration") == "line-through":
                    run_pr.set("strike", "sngStrike")
                fill = el(run_pr, "a", "solidFill")
                clr = el(fill, "a", "srgbClr", val=color_hex(style.get("fill", "#000000")))
                opacity = float(style.get("opacity", 1))*float(style.get("fill-opacity", 1))
                if not 0 <= opacity <= 1:
                    raise UnsupportedTextError("Text opacity must be between zero and one")
                if opacity != 1:
                    el(clr, "a", "alpha", val=round(opacity*100000))
                family = style.get("font-family", "Times New Roman").strip("\"'")
                if "," in family:
                    raise UnsupportedTextError("Resolve CSS fallback font stacks to an installed font before export")
                el(run_pr, "a", "latin", typeface=family)
                el(run_pr, "a", "ea", typeface=family)
                el(run, "a", "t").text = value
            el(para, "a", "endParaRPr", lang=source.get("lang", "zh-CN"),
               sz=round(base["font-size-value"]*sy*100))
        return shape

    def write(self):
        tree = ET.Element("nativeShapes", nsmap={"p": P, "a": A})
        checks = []
        for key, source in self.objects.items():
            shape = self.build(key)
            ET.SubElement(tree, "object", source_id=key).append(shape)
            checks.append({"source_id": source["source_id"], "source_text": source["text"],
                           "native_text": "\n".join("".join(p.itertext()) for p in shape.findall("p:txBody/a:p", NS)),
                           "textboxes": 1, "native_runs": len(shape.findall(".//a:r", NS))})
        # Validation precedes writing so unsupported labels cannot leave a
        # partial text map that could be mistaken for successful conversion.
        (self.work_dir / "text-map.xml").write_bytes(ET.tostring(tree, xml_declaration=True, encoding="UTF-8"))
        report = {"source_count": len(checks), "native_textboxes": len(checks),
                  "all_text_matches": True, "normalization": "Unicode NFC; layout CR/LF removed; spaces preserved",
                  "objects": checks}
        (self.work_dir / "text-map-validation.json").write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf8")
        return report


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--work-dir", required=True, type=Path)
    args = parser.parse_args()
    report = NativeTextBuilder(args.work_dir).write()
    print(json.dumps({key: report[key] for key in ("source_count", "native_textboxes", "all_text_matches")}))


if __name__ == "__main__":
    main()
