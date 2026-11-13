#!/usr/bin/env python3
"""Convert evaluated Visio SVG geometry into simple, editable DrawingML.

Usage: python native_geometry.py --work-dir WORK

Inputs in WORK:
  source.svg: Visio SVG export. The original is read only.
  manifest.json: slide_width_pt, slide_height_pt, optional page_width_in, and
    objects [{source_id, raster, source_cells, optional svg_id}]. Source-cell
    values are evaluated Visio values, with length cells in inches.
  geometry-overrides.json (optional): {"objects": {"SOURCE_ID": {
    "points": [[x,y], ...], "closed": true, "style": {"fill": "#808080"}
  }}}. Points use absolute source-SVG viewBox coordinates, not slide units.

Outputs: geometry-map.xml and geometry-report.json. Native geometry replaces
metafile fragments; text, images, and OLE are assembled separately. Source SVG
3D projections are NOT reliable: objects with 3D rotation require an explicit
projected-polygon override. Unsupported visible geometry also fails closed.
The report always lists actionable unresolved items; exit status 2 means the
XML must not be delivered. There are no document-specific IDs or coordinates.

Requires lxml. This module does not use python-pptx or run Office applications.
"""
from __future__ import annotations

import argparse
import json
import math
import re
import sys
from pathlib import Path

from lxml import etree as ET

SVG = "http://www.w3.org/2000/svg"
DML = "http://schemas.openxmlformats.org/drawingml/2006/main"
PML = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS = {"s": SVG, "a": DML, "p": PML}
IDENTITY = (1.0, 0.0, 0.0, 1.0, 0.0, 0.0)
NUMBER = r"[-+]?(?:\d*\.\d+|\d+\.?\d*)(?:[eE][-+]?\d+)?"
PATH_TOKEN = r"[MmLlHhVvCcSsQqTtAaZz]|" + NUMBER
DRAWABLE = {"path", "rect", "ellipse", "circle", "line", "polyline", "polygon"}
SKIP = {"title", "desc", "metadata", "defs", "style", "text", "tspan"}


class GeometryError(ValueError):
    """An unsupported construct that must be resolved before delivery."""


def matrix_product(m, n):
    a, b, c, d, e, f = m
    g, h, i, j, k, l = n
    return (a*g+c*h, b*g+d*h, a*i+c*j, b*i+d*j, a*k+c*l+e, b*k+d*l+f)


def apply_matrix(matrix, xy):
    a, b, c, d, e, f = matrix
    x, y = xy
    return a*x+c*y+e, b*x+d*y+f


def parse_transform(value):
    result = IDENTITY
    for name, arguments in re.findall(r"(\w+)\s*\(([^)]+)\)", value or ""):
        v = [float(x) for x in re.findall(NUMBER, arguments)]
        if not all(math.isfinite(x) for x in v):
            raise GeometryError("SVG transform has nonfinite values")
        if name == "translate" and len(v) in (1, 2):
            m = (1, 0, 0, 1, v[0], v[1] if len(v) == 2 else 0)
        elif name == "scale" and len(v) in (1, 2):
            m = (v[0], 0, 0, v[-1], 0, 0)
        elif name == "rotate" and len(v) in (1, 3):
            angle = math.radians(v[0])
            c, s = math.cos(angle), math.sin(angle)
            m = (c, s, -s, c, 0, 0)
            if len(v) == 3:
                m = matrix_product(matrix_product((1, 0, 0, 1, v[1], v[2]), m),
                                   (1, 0, 0, 1, -v[1], -v[2]))
        elif name == "matrix" and len(v) == 6:
            m = tuple(v)
        elif name == "skewX" and len(v) == 1:
            m = (1, 0, math.tan(math.radians(v[0])), 1, 0, 0)
        elif name == "skewY" and len(v) == 1:
            m = (1, math.tan(math.radians(v[0])), 0, 1, 0, 0)
        else:
            raise GeometryError(f"Unsupported SVG transform: {name}({arguments})")
        result = matrix_product(result, m)
    remainder = re.sub(r"\w+\s*\([^)]+\)", "", value or "").strip(" ,\t\r\n")
    if remainder:
        raise GeometryError(f"Unparsed SVG transform: {remainder}")
    return result


def arc_to_cubics(start, rx, ry, phi, large, sweep, end):
    """SVG endpoint-parameterized arc, converted to <=90-degree cubic pieces."""
    rx, ry = abs(rx), abs(ry)
    if not rx or not ry:
        return [("L", end)]
    if start == end:
        return []
    if large not in (0, 1) or sweep not in (0, 1):
        raise GeometryError("SVG arc flags must be 0 or 1")
    angle = math.radians(phi)
    co, si = math.cos(angle), math.sin(angle)
    dx, dy = (start[0]-end[0])/2, (start[1]-end[1])/2
    xp, yp = co*dx+si*dy, -si*dx+co*dy
    ratio = xp*xp/(rx*rx)+yp*yp/(ry*ry)
    if ratio > 1:
        rx *= math.sqrt(ratio)
        ry *= math.sqrt(ratio)
    denominator = rx*rx*yp*yp+ry*ry*xp*xp
    factor = math.sqrt(max(0, (rx*rx*ry*ry-denominator)/denominator))
    if bool(large) == bool(sweep):
        factor = -factor
    cxp, cyp = factor*rx*yp/ry, -factor*ry*xp/rx
    cx = co*cxp-si*cyp+(start[0]+end[0])/2
    cy = si*cxp+co*cyp+(start[1]+end[1])/2
    u = ((xp-cxp)/rx, (yp-cyp)/ry)
    v = ((-xp-cxp)/rx, (-yp-cyp)/ry)
    theta = math.atan2(u[1], u[0])
    delta = math.atan2(u[0]*v[1]-u[1]*v[0], u[0]*v[0]+u[1]*v[1])
    if not sweep and delta > 0:
        delta -= 2*math.pi
    if sweep and delta < 0:
        delta += 2*math.pi
    count = max(1, math.ceil(abs(delta)/(math.pi/2)))
    step = delta/count

    def xy(t):
        return (cx+rx*co*math.cos(t)-ry*si*math.sin(t),
                cy+rx*si*math.cos(t)+ry*co*math.sin(t))

    def derivative(t):
        return (-rx*co*math.sin(t)-ry*si*math.cos(t),
                -rx*si*math.sin(t)+ry*co*math.cos(t))

    result = []
    for k in range(count):
        t0, t1 = theta+k*step, theta+(k+1)*step
        alpha = 4/3*math.tan(step/4)
        p0, p1, d0, d1 = xy(t0), xy(t1), derivative(t0), derivative(t1)
        result.append(("C", (p0[0]+alpha*d0[0], p0[1]+alpha*d0[1]),
                       (p1[0]-alpha*d1[0], p1[1]-alpha*d1[1]), p1))
    result[-1] = (*result[-1][:-1], end)
    return result


def parse_path(value):
    if re.sub(PATH_TOKEN+r"|[\s,]+", "", value):
        raise GeometryError("Unsupported token in SVG path")
    tokens = re.findall(PATH_TOKEN, value)
    index, command, pos, first, previous, control = 0, None, (0., 0.), (0., 0.), None, None
    arity = {"M": 2, "L": 2, "H": 1, "V": 1, "C": 6, "S": 4,
             "Q": 4, "T": 2, "A": 7, "Z": 0}
    output, raw = [], []
    while index < len(tokens):
        if tokens[index].isalpha():
            command = tokens[index]
            index += 1
        if command is None:
            raise GeometryError("SVG path data has no command")
        upper, relative = command.upper(), command.islower()
        count = arity[upper]
        try:
            values = [float(v) for v in tokens[index:index+count]]
        except ValueError as exc:
            raise GeometryError("SVG path has an incomplete segment") from exc
        if len(values) != count or not all(math.isfinite(v) for v in values):
            raise GeometryError("SVG path has incomplete or nonfinite coordinates")
        index += count
        raw.append((upper, values))

        def pair(j):
            return (values[j]+(pos[0] if relative else 0),
                    values[j+1]+(pos[1] if relative else 0))

        if upper == "Z":
            output.append(("Z",))
            pos, command = first, None
        elif upper in ("M", "L"):
            pos = pair(0)
            output.append((upper, pos))
            if upper == "M":
                first, command = pos, "l" if relative else "L"
        elif upper == "H":
            pos = values[0]+(pos[0] if relative else 0), pos[1]
            output.append(("L", pos))
        elif upper == "V":
            pos = pos[0], values[0]+(pos[1] if relative else 0)
            output.append(("L", pos))
        elif upper == "C":
            p1, p2, end = pair(0), pair(2), pair(4)
            output.append(("C", p1, p2, end))
            control, pos = p2, end
        elif upper == "S":
            p1 = ((2*pos[0]-control[0], 2*pos[1]-control[1])
                  if previous in ("C", "S") else pos)
            p2, end = pair(0), pair(2)
            output.append(("C", p1, p2, end))
            control, pos = p2, end
        elif upper == "Q":
            p1, end = pair(0), pair(2)
            output.append(("Q", p1, end))
            control, pos = p1, end
        elif upper == "T":
            p1 = ((2*pos[0]-control[0], 2*pos[1]-control[1])
                  if previous in ("Q", "T") else pos)
            end = pair(0)
            output.append(("Q", p1, end))
            control, pos = p1, end
        elif upper == "A":
            end = pair(5)
            output.extend(arc_to_cubics(pos, *values[:5], end))
            pos = end
        previous = upper
    return output, raw


def rectangle_path(x, y, width, height, radius=0):
    radius = min(radius, width/2, height/2)
    if not radius:
        return [("M", (x, y)), ("L", (x+width, y)),
                ("L", (x+width, y+height)), ("L", (x, y+height)), ("Z",)]
    r = radius
    d = (f"M{x+r} {y} H{x+width-r} A{r} {r} 0 0 1 {x+width} {y+r} "
         f"V{y+height-r} A{r} {r} 0 0 1 {x+width-r} {y+height} "
         f"H{x+r} A{r} {r} 0 0 1 {x} {y+height-r} "
         f"V{y+r} A{r} {r} 0 0 1 {x+r} {y} Z")
    return parse_path(d)[0]


def bounds(commands):
    points = [p for command in commands for p in command[1:]]
    if not points:
        raise GeometryError("Empty geometry")
    xs, ys = zip(*points)
    return min(xs), min(ys), max(xs)-min(xs), max(ys)-min(ys)


def diamond_bounds(commands):
    """Recognize four axis-symmetric vertices, independently of source IDs."""
    if not commands or commands[-1][0] != "Z" or any(c[0] not in ("M", "L", "Z") for c in commands):
        return None
    points = [c[1] for c in commands if len(c) > 1]
    if len(points) > 1 and math.dist(points[0], points[-1]) < 1e-7:
        points.pop()
    if len(points) != 4:
        return None
    x, y, width, height = bounds(commands)
    if not width or not height:
        return None
    expected = [(x+width/2, y), (x+width, y+height/2),
                (x+width/2, y+height), (x, y+height/2)]
    tolerance = max(width, height)*1e-4
    if all(any(math.dist(p, q) <= tolerance for q in expected) for p in points):
        return x, y, width, height
    return None


def color(parent, value, opacity=1):
    value = (value or "#000000").strip().lower()
    if value == "none":
        ET.SubElement(parent, f"{{{DML}}}noFill")
        return
    aliases = {"black": "000000", "white": "ffffff", "red": "ff0000",
               "green": "008000", "blue": "0000ff", "yellow": "ffff00",
               "gray": "808080", "grey": "808080", "transparent": "000000"}
    if value == "transparent":
        opacity = 0
    if value.startswith("#"):
        value = value[1:]
        if len(value) == 3:
            value = "".join(c*2 for c in value)
    elif value.startswith("rgb("):
        parts = [p.strip() for p in value[4:-1].split(",")]
        if len(parts) != 3:
            raise GeometryError(f"Unsupported color: {value}")
        values = [round(float(p[:-1])*2.55) if p.endswith("%") else round(float(p)) for p in parts]
        value = "".join(f"{max(0,min(255,v)):02x}" for v in values)
    else:
        value = aliases.get(value, value)
    if not re.fullmatch(r"[0-9a-f]{6}", value):
        raise GeometryError(f"Unsupported fill/stroke color {value!r}; use an explicit solid color or projected override")
    fill = ET.SubElement(parent, f"{{{DML}}}solidFill")
    channel = ET.SubElement(fill, f"{{{DML}}}srgbClr", val=value.upper())
    if opacity < 0.99999:
        ET.SubElement(channel, f"{{{DML}}}alpha", val=str(round(max(0, min(1, opacity))*100000)))


class Converter:
    def __init__(self, svg_root, manifest, overrides=None):
        self.root, self.manifest, self.overrides = svg_root, manifest, overrides or {}
        self.css = {}
        for sheet in svg_root.findall("s:style", NS):
            for name, body in re.findall(r"\.([\w-]+)\s*\{([^}]+)\}", sheet.text or ""):
                self.css[name] = dict(re.findall(r"([\w-]+)\s*:\s*([^;]+)", body))
        self.lookup = {el.get("id"): el for el in svg_root.iter()
                       if isinstance(el.tag, str) and el.get("id")}
        self.shape_lookup = {re.match(r"(?:shape|group)(\d+)-", key).group(1): el
                             for key, el in self.lookup.items()
                             if re.match(r"(?:shape|group)\d+-", key)}
        view = [float(v) for v in re.findall(NUMBER, svg_root.get("viewBox", ""))]
        if len(view) != 4 or view[2] <= 0 or view[3] <= 0:
            raise GeometryError("source.svg must have a positive four-value viewBox")
        sx = float(manifest["slide_width_pt"])*12700/view[2]
        sy = float(manifest["slide_height_pt"])*12700/view[3]
        if not math.isfinite(sx+sy) or min(sx, sy) <= 0:
            raise GeometryError("Manifest slide dimensions must be finite and positive")
        if abs(sx/sy-1) > 0.005:
            raise GeometryError("Slide and source SVG aspect ratios differ; set matching slide dimensions before converting")
        self.to_slide = sx, 0, 0, sy, -view[0]*sx, -view[1]*sy
        source_width = manifest.get("page_width_in")
        if not source_width:
            size = re.fullmatch(r"\s*("+NUMBER+r")\s*(in|pt|mm|cm)\s*", svg_root.get("width", ""))
            if size:
                source_width = float(size[1])*{"in": 1, "pt": 1/72, "mm": 1/25.4, "cm": 1/2.54}[size[2]]
        self.effect_scale = (float(manifest["slide_width_pt"])*12700/float(source_width)
                             if source_width else None)
        self.warnings = []
        self.flat_ids = {str(o["source_id"]) for o in manifest["objects"]}

    def style_for(self, element):
        style = {"fill": "none", "stroke": "none"}
        for el in list(element.iterancestors())[::-1]+[element]:
            for name in el.get("class", "").split():
                style.update(self.css.get(name, {}))
            style.update(dict(re.findall(r"([\w-]+)\s*:\s*([^;]+)", el.get("style", ""))))
            for name in ("fill", "stroke", "fill-opacity", "stroke-opacity", "stroke-width",
                         "stroke-dasharray", "stroke-linecap", "stroke-linejoin", "opacity",
                         "filter", "visibility", "display", "marker-end", "marker-start",
                         "clip-path", "mask"):
                if el.get(name) is not None:
                    style[name] = el.get(name)
        return style

    def full_transform(self, element):
        matrix = self.to_slide
        for el in list(element.iterancestors())[::-1]+[element]:
            matrix = matrix_product(matrix, parse_transform(el.get("transform")))
        return matrix

    def geometry_elements(self, group, source_id):
        for el in group:
            if not isinstance(el.tag, str) or not el.tag.startswith("{"+SVG+"}"):
                continue
            tag = ET.QName(el).localname
            if tag in SKIP:
                continue
            style = self.style_for(el)
            if style.get("display") == "none" or style.get("visibility") == "hidden":
                continue
            nested = re.match(r"(?:shape|group)(\d+)-", el.get("id", ""))
            if nested and nested[1] != source_id and nested[1] in self.flat_ids:
                continue  # This child has its own flattened manifest entry.
            if style.get("filter", "none") != "none":
                # Visio adds separate shadow-only geometry. These are not source
                # objects and must not be converted into extra raster fragments.
                if el.get("id", "").startswith("shadow") or any(a.get("id", "").startswith("shadow") for a in el.iterancestors()):
                    self.warnings.append({"source_id": source_id, "reason": "Decorative SVG shadow skipped; add a native PPT shadow if required"})
                    continue
                raise GeometryError("Visible SVG filter needs explicit handling or a geometry override")
            if style.get("clip-path", "none") != "none" or style.get("mask", "none") != "none":
                raise GeometryError("Clipped/masked native geometry needs an explicit projected polygon override")
            if tag == "g":
                yield from self.geometry_elements(el, source_id)
            elif tag in DRAWABLE:
                yield el
            else:
                raise GeometryError(f"Unsupported visible SVG <{tag}>; flatten manifest children, mark true imagery as raster, or provide an override")

    def make_shape(self, name, commands, style, matrix, preset=None, local_box=None, adjustment=None):
        world = [(op, *[apply_matrix(matrix, p) for p in pts]) for op, *pts in commands]
        stroke_width = float(style.get("stroke-width", 1))
        if stroke_width < 0:
            raise GeometryError("Negative stroke width")
        matrix_scale = math.hypot(matrix[0], matrix[1])
        for key, index, neighbor_index in (("marker-start", 0, 1), ("marker-end", -1, -2)):
            if style.get(key, "none") == "none" or len(world) < 2:
                continue
            marker = self.lookup.get(style[key].split("#")[-1].rstrip(")"))
            if marker is None:
                raise GeometryError(f"Missing SVG marker: {style[key]}")
            if not world[index][1:]:
                raise GeometryError("Marker on closed path needs explicit reconstruction")
            distance = abs(float(marker.get("refX", "0")))*stroke_width*matrix_scale
            target, neighbor = world[index][-1], world[neighbor_index][-1]
            if index == -1 and world[-1][0] in ("C", "Q"):
                neighbor = world[-1][-2]
            elif index == 0 and world[1][0] in ("C", "Q"):
                neighbor = world[1][1]
            dx, dy = target[0]-neighbor[0], target[1]-neighbor[1]
            length = math.hypot(dx, dy)
            if length:
                world[index] = (*world[index][:-1], (target[0]+distance*dx/length, target[1]+distance*dy/length))
        x0, y0, width, height = bounds(world)
        rotation = 0
        orthogonal = abs(matrix[0]*matrix[2]+matrix[1]*matrix[3]) <= max(1, matrix_scale*math.hypot(matrix[2], matrix[3]))*1e-8
        if preset in ("ellipse", "rect", "roundRect", "diamond") and not orthogonal:
            preset = None  # Keep true sheared geometry in one native custom path.
        if preset in ("ellipse", "rect", "roundRect", "diamond") and local_box:
            lx, ly, lw, lh = local_box
            center = apply_matrix(matrix, (lx+lw/2, ly+lh/2))
            width, height = lw*matrix_scale, lh*math.hypot(matrix[2], matrix[3])
            x0, y0 = center[0]-width/2, center[1]-height/2
            rotation = math.degrees(math.atan2(matrix[1], matrix[0]))
        x, y, w, h = round(x0), round(y0), max(1, round(width)), max(1, round(height))
        shape = ET.Element(f"{{{PML}}}sp")
        nv = ET.SubElement(shape, f"{{{PML}}}nvSpPr")
        ET.SubElement(nv, f"{{{PML}}}cNvPr", id="1", name=name)
        ET.SubElement(nv, f"{{{PML}}}cNvSpPr")
        ET.SubElement(nv, f"{{{PML}}}nvPr")
        properties = ET.SubElement(shape, f"{{{PML}}}spPr")
        xf = ET.SubElement(properties, f"{{{DML}}}xfrm")
        if rotation:
            xf.set("rot", str(round((rotation % 360)*60000)))
        if preset == "line":
            if world[-1][-1][0] < world[0][-1][0]:
                xf.set("flipH", "1")
            if world[-1][-1][1] < world[0][-1][1]:
                xf.set("flipV", "1")
        ET.SubElement(xf, f"{{{DML}}}off", x=str(x), y=str(y))
        ET.SubElement(xf, f"{{{DML}}}ext", cx=str(w), cy=str(h))
        if preset:
            geom = ET.SubElement(properties, f"{{{DML}}}prstGeom", prst=preset)
            av = ET.SubElement(geom, f"{{{DML}}}avLst")
            if adjustment is not None:
                ET.SubElement(av, f"{{{DML}}}gd", name="adj", fmla="val "+str(adjustment))
        else:
            geom = ET.SubElement(properties, f"{{{DML}}}custGeom")
            for name in ("avLst", "gdLst", "ahLst", "cxnLst"):
                ET.SubElement(geom, f"{{{DML}}}{name}")
            ET.SubElement(geom, f"{{{DML}}}rect", l="0", t="0", r="r", b="b")
            paths = ET.SubElement(geom, f"{{{DML}}}pathLst")
            path = ET.SubElement(paths, f"{{{DML}}}path", w=str(w), h=str(h))
            tags = {"M": "moveTo", "L": "lnTo", "C": "cubicBezTo", "Q": "quadBezTo", "Z": "close"}
            for command in world:
                op = ET.SubElement(path, f"{{{DML}}}{tags[command[0]]}")
                for px, py in command[1:]:
                    ET.SubElement(op, f"{{{DML}}}pt", x=str(round(px-x0)), y=str(round(py-y0)))
        opacity = float(style.get("opacity", 1))
        color(properties, style.get("fill", "none"), float(style.get("fill-opacity", 1))*opacity)
        line = ET.SubElement(properties, f"{{{DML}}}ln", w=str(max(1, round(stroke_width*matrix_scale))),
                             cap={"round": "rnd", "square": "sq"}.get(style.get("stroke-linecap"), "flat"))
        color(line, style.get("stroke", "none"), float(style.get("stroke-opacity", 1))*opacity)
        dash = [float(n) for n in re.findall(NUMBER, style.get("stroke-dasharray", ""))]
        if dash:
            if not stroke_width or any(v < 0 for v in dash):
                raise GeometryError("Dash lengths and stroke width must be valid positive values")
            if len(dash) % 2:
                dash *= 2
            custom = ET.SubElement(line, f"{{{DML}}}custDash")
            for d, gap in zip(dash[::2], dash[1::2]):
                ET.SubElement(custom, f"{{{DML}}}ds", d=str(round(max(d/stroke_width, .01)*100000)), sp=str(round(gap/stroke_width*100000)))
        else:
            ET.SubElement(line, f"{{{DML}}}prstDash", val="solid")
        ET.SubElement(line, f"{{{DML}}}round")
        for attribute, tag in (("marker-start", "headEnd"), ("marker-end", "tailEnd")):
            if style.get(attribute, "none") != "none":
                ET.SubElement(line, f"{{{DML}}}{tag}", type="triangle", w="med", len="med")
        ET.SubElement(properties, f"{{{DML}}}effectLst")
        return shape

    def convert_element(self, element, source_id, index):
        style = self.style_for(element)
        if style.get("fill") == "none" and style.get("stroke") == "none":
            return None
        tag, matrix = ET.QName(element).localname, self.full_transform(element)
        preset = box = adjustment = None
        if tag == "path":
            commands, raw = parse_path(element.get("d", ""))
            code = "".join(op for op, _ in raw)
            if len(commands) == 2 and commands[0][0] == "M" and commands[1][0] == "L":
                preset = "line"
            if code == "MAAZ":
                arcs = [v for op, v in raw if op == "A"]
                start, middle, end = raw[0][1], arcs[0][-2:], arcs[1][-2:]
                if (all(abs(a-b) < 1e-6 for a, b in zip(arcs[0][:3], arcs[1][:3]))
                        and abs(arcs[0][2]) < 1e-6 and math.dist(start, end) < 1e-4):
                    rx, ry = arcs[0][:2]
                    cx, cy = (start[0]+middle[0])/2, (start[1]+middle[1])/2
                    if min(rx, ry) > 0 and abs(math.dist(start, middle)/(2*rx)-1) < 1e-3 and abs(start[1]-middle[1]) < 1e-3:
                        box, preset = (cx-rx, cy-ry, 2*rx, 2*ry), "ellipse"
            if code == "MLALALALAZ":
                box = bounds(commands)
                radius = next(v[0] for op, v in raw if op == "A")
                preset, adjustment = "roundRect", min(50000, round(radius/min(box[2:])*100000))
            diamond = diamond_bounds(commands)
            if diamond:
                box, preset = diamond, "diamond"
        elif tag == "rect":
            x, y, w, h = [float(element.get(k, 0)) for k in ("x", "y", "width", "height")]
            if w <= 0 or h <= 0:
                raise GeometryError("Rectangle dimensions must be positive")
            radius = min(float(element.get("rx", element.get("ry", 0))), w/2, h/2)
            commands, box = rectangle_path(x, y, w, h, radius), (x, y, w, h)
            preset = "roundRect" if radius else "rect"
            adjustment = round(radius/min(w, h)*100000) if radius else None
        elif tag in ("ellipse", "circle"):
            cx, cy = float(element.get("cx", 0)), float(element.get("cy", 0))
            rx, ry = float(element.get("rx", element.get("r", 0))), float(element.get("ry", element.get("r", 0)))
            if min(rx, ry) <= 0:
                raise GeometryError("Ellipse radii must be positive")
            commands = parse_path(f"M{cx-rx} {cy} A{rx} {ry} 0 0 1 {cx+rx} {cy} A{rx} {ry} 0 0 1 {cx-rx} {cy} Z")[0]
            box, preset = (cx-rx, cy-ry, 2*rx, 2*ry), "ellipse"
        elif tag == "line":
            commands = [("M", (float(element.get("x1", 0)), float(element.get("y1", 0)))),
                        ("L", (float(element.get("x2", 0)), float(element.get("y2", 0))))]
            preset = "line"
        else:
            values = [float(v) for v in re.findall(NUMBER, element.get("points", ""))]
            if len(values) < 4 or len(values) % 2:
                raise GeometryError("Invalid polygon/polyline coordinates")
            commands = [("M" if i == 0 else "L", (values[i], values[i+1])) for i in range(0, len(values), 2)]
            if tag == "polygon":
                commands.append(("Z",))
                diamond = diamond_bounds(commands)
                if diamond:
                    box, preset = diamond, "diamond"
        return self.make_shape(f"Visio_{source_id}_geometry_{index}", commands, style, matrix, preset, box, adjustment)

    def apply_glow(self, shape, cells):
        size = float(cells.get("GlowSize", 0))
        if size <= 0:
            return
        if not self.effect_scale:
            raise GeometryError("GlowSize is in inches; supply manifest page_width_in or physical SVG width to scale it")
        value = cells.get("GlowColor")
        if not value:
            raise GeometryError("GlowSize is present but GlowColor is missing; resolve inherited source cells")
        opacity = 1-float(cells.get("GlowColorTrans", 0))
        effects = shape.find("p:spPr/a:effectLst", NS)
        glow = ET.SubElement(effects, f"{{{DML}}}glow", rad=str(round(size*self.effect_scale)))
        temp = ET.Element("temporary")
        color(temp, value, opacity)
        color_node = temp.find("a:solidFill/a:srgbClr", NS)
        if color_node is None:
            raise GeometryError("GlowColor must resolve to a solid color")
        glow.append(color_node)

    def convert_override(self, source_id, group, override):
        points = override.get("points")
        if not isinstance(points, list) or len(points) < 2:
            raise GeometryError("geometry-overrides entry needs at least two projected points in source SVG viewBox coordinates")
        try:
            if any(not isinstance(p, (list, tuple)) or len(p) != 2 for p in points):
                raise GeometryError("Each override point must contain exactly two coordinates")
            points = [(float(p[0]), float(p[1])) for p in points]
        except (TypeError, ValueError) as exc:
            raise GeometryError("Invalid override points") from exc
        if len(points) < 2 or not all(math.isfinite(v) for p in points for v in p):
            raise GeometryError("Invalid or nonfinite override points")
        # The original style may still be useful even where SVG projected geometry
        # is wrong. Explicit style values always take precedence.
        original = next((e for e in group.iter() if isinstance(e.tag, str) and ET.QName(e).namespace == SVG and ET.QName(e).localname in DRAWABLE), group)
        style = self.style_for(original)
        style.update(override.get("style", {}))
        style.pop("filter", None)
        commands = [("M" if i == 0 else "L", p) for i, p in enumerate(points)]
        if override.get("closed", True):
            if len(points) < 3:
                raise GeometryError("Closed override requires at least three projected points")
            commands.append(("Z",))
        return self.make_shape(f"Visio_{source_id}_projected", commands, style, self.to_slide)

    def convert(self):
        output = ET.Element("nativeShapes", nsmap={"p": PML, "a": DML})
        unresolved, counts = [], {}
        for item in self.manifest["objects"]:
            if item.get("raster"):
                continue
            sid = str(item["source_id"])
            entry = ET.SubElement(output, "object", source_id=sid)
            try:
                group = self.lookup.get(item.get("svg_id")) if item.get("svg_id") else self.shape_lookup.get(sid)
                if group is None:
                    raise GeometryError("Source SVG group is missing; supply svg_id or check exported page")
                cells = item.get("source_cells", {})
                override = self.overrides.get(sid)
                rotations = {k: float(cells.get(k, 0)) for k in ("RotationXAngle", "RotationYAngle", "RotationZAngle")}
                if any(abs(v) > 1e-9 for v in rotations.values()) and not override:
                    raise GeometryError("Visio 3D rotation is not preserved reliably in SVG. Render the source object, derive its projected corners, and add geometry-overrides.json for this ID")
                if override:
                    shape = self.convert_override(sid, group, override)
                    self.apply_glow(shape, cells)
                    entry.append(shape)
                else:
                    for index, element in enumerate(self.geometry_elements(group, sid)):
                        shape = self.convert_element(element, sid, index)
                        if shape is not None:
                            self.apply_glow(shape, cells)
                            entry.append(shape)
                counts[sid] = len(entry)
            except (GeometryError, ValueError, TypeError, IndexError) as exc:
                for child in list(entry):
                    entry.remove(child)
                entry.set("unresolved", "true")
                unresolved.append({"source_id": sid, "action": str(exc)})
        output.set("status", "unresolved" if unresolved else "ready")
        report = {"status": output.get("status"), "source_objects": len(counts),
                  "native_geometry_shapes": sum(counts.values()),
                  "custom_paths": len(output.findall(".//a:path", NS)),
                  "preset_shapes": len(output.findall(".//a:prstGeom", NS)),
                  "geometry_commands": sum(len(p) for p in output.findall(".//a:path", NS)),
                  "by_source": counts, "warnings": self.warnings, "unresolved": unresolved}
        return output, report


def convert_work_dir(work_dir):
    work_dir = Path(work_dir).resolve()
    manifest = json.loads((work_dir/"manifest.json").read_text(encoding="utf-8-sig"))
    parser = ET.XMLParser(huge_tree=True, resolve_entities=False, no_network=True)
    root = ET.parse(str(work_dir/"source.svg"), parser).getroot()
    overrides_path = work_dir/"geometry-overrides.json"
    overrides = json.loads(overrides_path.read_text(encoding="utf-8-sig")) if overrides_path.exists() else {}
    converter = Converter(root, manifest, overrides.get("objects", {}))
    output, report = converter.convert()
    (work_dir/"geometry-map.xml").write_bytes(ET.tostring(output, encoding="UTF-8", xml_declaration=True))
    (work_dir/"geometry-report.json").write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    return report


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--work-dir", type=Path, required=True)
    args = parser.parse_args(argv)
    try:
        report = convert_work_dir(args.work_dir)
    except (GeometryError, KeyError, ValueError, OSError, ET.XMLSyntaxError) as exc:
        report = {"status": "unresolved", "unresolved": [{"action": str(exc)}]}
        if args.work_dir.is_dir():
            (args.work_dir/"geometry-report.json").write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
            # Never leave a previously successful map next to a failed report.
            failed = ET.Element("nativeShapes", status="unresolved", nsmap={"p": PML, "a": DML})
            (args.work_dir/"geometry-map.xml").write_bytes(ET.tostring(failed, encoding="UTF-8", xml_declaration=True))
    print(json.dumps(report, ensure_ascii=True))
    return 2 if report.get("unresolved") else 0


if __name__ == "__main__":
    sys.exit(main())
