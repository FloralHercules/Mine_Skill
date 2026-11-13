# Conversion details and recovery

## Data flow

`export_visio.ps1` creates a private drawing and exports the original reference, flattens groups there, then exports a second reference, evaluated SVG, and per-object raster previews. `inspect_vsdx.py` resolves the chosen page and embedded-data relationships from that private VSDX. `native_geometry.py` and `native_text.py` write separate native DrawingML maps. `assemble_pptx.py` combines them, keeping z-order and reusing source OLE binaries. `powerpoint.ps1` opens, renders, and saves the result; `validate_output.py` checks the saved artifact and original source.

Source shape IDs are only local join keys for a particular run. Page IDs, shape IDs, canvas dimensions, object counts, and source paths must never be baked into the skill.

## Raster placement

Individual `Shape.Export` can discard transformed bounds. The exporter instead selects the object with a thin visible magenta full-page frame, exports on white with white transparency, then removes that frame and alpha-crops the result. Crop coordinates are measured against the full canvas and mapped to slide points. The frame is never saved to the source or final drawing. Inspect fine white detail: color-key transparency can remove intentional white content; when observed, use a source-native raster extraction or a verified alpha mask instead of accepting missing detail.

Nested foreign objects must be exposed in the private copy before selection export. Blank or missing PNGs are failures. Avoid generating hundreds of one-object VSDX packages as an extraction strategy; it is costly and can destabilize Visio.

## Native geometry overrides

Visio's SVG can omit 3D projection while retaining only the original flat rectangle. The exporter writes isolated `assets/diagnose-ID.png` references when source 3D cells are nonzero. The converter reports unsupported objects and exits rather than flattening their geometry to images.

For a planar quadrilateral, determine projected corners from the original reference or an accurate isolated rendering, then supply a native polygon override. Coordinates are in the SVG viewBox coordinate system, with top-left origin, not PNG pixels or PowerPoint points. Record the source and measurement method in a `note` beside the override.

```json
{
  "objects": {
    "SOURCE_SHAPE_ID": {
      "points": [[12, 20], [140, 35], [125, 85], [5, 70]],
      "closed": true,
      "style": {"fill": "#808080", "fill-opacity": "0.3"},
      "note": "Example coordinates only; replace with measured source projection."
    }
  }
}
```

Use `--geometry-overrides FILE.json --resume` after the exports completed. The polygon remains a native editable custom shape. Do not reuse example or previous-project coordinates. Nonplanar extrusions need an appropriate small native construction and a visual comparison, not an arbitrary quadrilateral.

`geometry-report.json` also identifies unsupported visible SVG. Add support only after checking the source semantics. Decorative shadow warnings need visual assessment; do not silently drop a substantive diagram element. The geometry helper supports native glow from saved Visio cells when SVG omits it.

## Text and formulas

The text helper uses evaluated SVG text positioning and Visio text-box metadata. It preserves editable rich-text runs, colors, font sizes, rotations, and line breaks. It rejects transformations that ordinary text boxes cannot reproduce faithfully. Compare exact label text after assembly and after PowerPoint saves the deck.

MathType objects use their source binary, confirmed class `Equation.DSMT4`, and a separate rendered preview. Unknown OLE classes must be identified rather than mislabeled as MathType. Keep the original binary intact, and compare hashes after the PowerPoint round trip. Office may rename embedded files; compare hashes as a multiset rather than relying on filenames. Compatibility fallback XML may duplicate an object's representation; count actual PowerPoint OLE objects rather than summing every fallback branch.

## Operational details

- Large SVGs with embedded imagery need `lxml.etree.XMLParser(huge_tree=True)`. Treat these local document inputs as untrusted; do not enable external entity or network resolution.
- Windows PowerShell 5 can misread UTF-8 scripts without a BOM under `-File`. The Python runner explicitly reads scripts as UTF-8, retaining Chinese paths and page names.
- Each Office phase has a bounded wait. The scripts close documents and application instances they created, not arbitrary user sessions. A timeout may leave a helper instance; inspect ownership before cleaning up.
- A native PPTX may render faster but take a different amount of time to open. Report measured behavior accurately. No universal shape-count threshold guarantees responsiveness.
- Keep the installed skill free of user drawings, embedded images, private paths, work manifests, previews, and generated decks.
