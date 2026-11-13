# Scene schema

One JSON scene represents one page. Coordinates use an upper-left origin, x right, y down. Default `units_per_inch`: 100. Text `size` and line `width` use scene units; `line_weight` uses points. Large diagrams may use a larger custom canvas.

```json
{
  "schema_version": 1,
  "title": "方法框架",
  "width": 900,
  "height": 500,
  "source_image": "source.png",
  "elements": [
    {"kind":"box","id":"panel","x":20,"y":20,"w":860,"h":460,"fill":"#EEF7FF","line":"#80B9E0","round":12},
    {"kind":"text","id":"title","x":50,"y":40,"w":800,"h":40,"text":"方法框架 Method Framework","size":26,"bold":true,"align":1},
    {"kind":"math","id":"loss","x":340,"y":300,"w":220,"h":55,"tex":"L=\\sum_{i=1}^{N}x_i^2"}
  ]
}
```

IDs are unique ASCII names beginning with a letter, followed by letters, digits, `_` or `-`. Keep IDs stable when correcting figures; avoid deriving them from list positions if edits could reorder elements.

Common rectangle-like fields: `x,y,w,h`; `w,h` positive.

| Kind | Additional fields |
|---|---|
| `box` | `fill`, `line`, `round`, optional `line_weight` |
| `text` | `text`, `size`, `bold`, `align` (0 left, 1 center, 2 right), `valign` (0 top, 1 middle, 2 bottom), `color` |
| `oval` | `fill`, `line`, optional `transparent` |
| `poly` | `points: [[x,y],...]`, `fill`, `line`; repeat first point to close |
| `line` | `x1,y1,x2,y2`, `color`, `width`, `arrow` |
| `photo` | `file` relative to scene, OR `crop: [left,top,right,bottom]` in scene coordinates from `source_image` |
| `math` | `tex`, placement box `x,y,w,h`; preserves aspect ratio |

Elements are painted in order. Put backgrounds first. Formulas are inserted last, so reserve their boxes. Photos are independent images. Native line endpoints are independently editable; extend the renderer if a task specifically needs glued dynamic connectors, and do not claim these lines are glued.

Split long bilingual captions into Chinese and English boxes when useful. Do not turn normal prose into formulas. Use MathType-supported TeX and check subscripts, superscripts, signs and delimiters visually.
