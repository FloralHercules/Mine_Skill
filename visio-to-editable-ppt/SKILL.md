---
name: visio-to-editable-ppt
description: Convert a page of a local Visio VSDX into an editable PowerPoint slide, retaining native shapes, editable text, separate images, and embedded MathType formulas. Use for Visio转PPT、VSDX转可编辑PPT、框架图转原生形状 or fixing a converted diagram that lags after metafile ungrouping. Not for screenshot-to-diagram reconstruction or ordinary slide design.
---

# Visio to Editable PowerPoint

Rebuild the selected Visio page as a single slide that remains practical to edit. Preserve the source layout, stacking order, colors, text, and formulas. Rectangles, rounded boxes, ellipses, arrows, and connectors should be native PowerPoint objects; photographs remain separate images. Keep MathType formulas as their original embedded OLE objects with rendered previews.

## Conversion

1. Resolve the source VSDX, requested page, and destination. List pages with `scripts/inspect_vsdx.py --source INPUT.vsdx --list-pages` if needed. Use the page name rather than assuming a page number. Work in a fresh private folder; the source drawing must remain unchanged.
2. This implementation needs Windows desktop Visio and PowerPoint, PowerShell, and Python 3.10+ with `lxml` and `Pillow`. For the bundled Python runtime, discover `load_workspace_dependencies` when available; do not hardcode a runtime or plugin version. If Office is unavailable, explain the missing dependency and preserve the requested editable output rather than silently substituting a flat image.
3. Run the bundled pipeline using the discovered Python executable:

   ```text
   python scripts/convert.py --source INPUT.vsdx --page "页面名称" --work-dir PRIVATE_WORK --output OUTPUT.pptx
   ```

   Keep the original aspect ratio. Use `--slide-width-in 13.333` when a smaller slide is wanted or the original exceeds PowerPoint's size limit. Use `--overwrite-output` only for the exact output file the task authorizes replacing.
4. Inspect `validation.json`, `geometry-report.json`, `text-map-validation.json`, and the exported images. Compare `reference.png` with `flattened-reference.png` first, then compare `preview.png` with the original. Check text wrapping, formulas, arrowheads, transparency, projected planes, and layer order. A successful script exit is not a visual approval.
5. Deliver the verified PPTX with a clickable local file link and a short description of what is editable. Report any material visual or editing limitation. Do not claim MathType activation was tested merely because the embedded bytes were preserved; opening formulas requires a compatible local editor.

## Fidelity and responsiveness

- Do not use one full-slide image, a whole-slide SVG, or an embedded Visio object as a substitute for editable slide elements.
- Do not use EMF/WMF ungrouping as the primary conversion route. It can explode dashed lines, text, and fills into many paths. Preserve dashes and arrowheads as native line properties; use small custom paths only where needed.
- Do not turn photographs into thousands of vector fragments. Keep them as independent raster objects with their original crop and placement.
- The pipeline ungroups only a private Visio copy so nested formulas can be extracted. If that changes appearance, resolve the source group behavior before delivery.
- SVG exports may lose Visio 3D projection or effects. The geometry converter stops for unsupported projection. Read [conversion-notes.md](references/conversion-notes.md) for evidence-backed native geometry overrides and recovery. Never guess missing corners or conceal a failed conversion with a bitmap.
- Source fonts and labels take precedence. Preserve rich text and mixed Chinese/English text; do not impose a new font scheme unless requested.
- Performance is checked by counting objects/path commands and opening/rendering in PowerPoint. Measure opening and rendering separately; reduced file size alone does not establish responsiveness.

## Repairs and testing

Use `--resume` to rebuild from exports for the same source/page/settings after a documented fix; it verifies the original source hash. It does not restart an incomplete Office export. For that failure, use a new work folder after checking the Office state. Never terminate all Visio or PowerPoint processes to clear one failed conversion.

Read [conversion-notes.md](references/conversion-notes.md) when extending the supported geometry, resolving an OLE class, or debugging export placement. After modifying helpers, run `python -m unittest discover -s scripts -p "test_*.py"` and a small end-to-end Office conversion. Keep test drawings and generated output outside the skill package.
