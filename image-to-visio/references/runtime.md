# Runtime and recovery

Requirements: desktop Microsoft Visio, desktop MathType registered as `Equation.DSMT4`, Microsoft Word with working `MTCommand_TeXToggle`, Windows Python, `pywin32` and `Pillow`. Native diagrams without formulas need no Word/MathType.

Python loads normal packages or the skill's `.runtime` folder. `VISIO_SKILL_PYTHONPATH` can point to another compatible dependency directory. Compiled packages must match Python version/architecture. On another computer:

```text
python -m pip install --target <skill>/.runtime pywin32 Pillow
```

Do not bundle proprietary fonts, MathType or Office. Do not automatically change macro security, licensing or startup preferences. Report missing prerequisites accurately.

## Observed interoperability details

- Use `Visio.InvisibleApp` and close only owned applications/documents. Never attach to and quit an arbitrary active Office instance.
- Set `Char.Font` to Times New Roman, `Char.AsianFont` to SimSun and explicit Chinese runs. Character indexes are UTF-16 code units.
- Visio `PasteSpecial(65536)` inserts OLE; value 7 is OEM text.
- Keep Visio events enabled for OLE insertion. Native rendering may defer recalculation.
- This machine stalled after about ten newly generated formulas in one chain. This is an observation, not a universal MathType limit. Workers use at most six formulas per fresh Visio AND Word session, saving each object and resuming missing IDs.
- Timeouts preserve checkpoints and terminate only the owned Python worker, not Office. Inspect remaining owned application state before retrying. Never kill Office by executable name or assume an old PID still belongs to the task.
- The master is never the partially rendered page. A formula failure leaves a working checkpoint and does not publish an incomplete master.

## Project semantics

`visio-project.json` stores the project-relative master path and imported figure/page history. Global skills must not hard-code project paths.

Appending uses matching `Copy(1)` and `Paste(1)` to preserve coordinates and native/OLE objects. Compare existing page shape counts, text content and per-page MathType counts after round-trip. Stage, version-check, back up and atomically commit. Photos and MathType data remain embedded.

If the master is open and locked, retain the validated candidate, explain that file locking is the blocker, and resume after the user closes it. Do not force-close Office or silently choose another final file.

## Provenance

Developed from the local Remote-Things workflow, informed by:

- https://github.com/deermiya/visio-skill — native Visio COM reconstruction.
- https://github.com/felimet/mathtype-for-word — TeX Toggle / MathType OLE.
- https://learn.microsoft.com/en-us/office/vba/api/visio.page.paste — coordinate-preserving paste.
- https://learn.microsoft.com/en-us/office/vba/api/visio.vispastespecialcodes — OLE paste format.

This implementation does not require those repositories to be installed.
