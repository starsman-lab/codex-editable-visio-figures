# Spec Schema

`visio_apply_spec.ps1` accepts a UTF-8 JSON file with this general shape:

```json
{
  "page": {
    "name": "Figure1",
    "width": 10.0,
    "height": 6.0
  },
  "shapes": [
    {
      "id": "panel-a",
      "type": "rect",
      "x": 0.5,
      "y": 0.5,
      "width": 4.0,
      "height": 2.2,
      "text": "Encoder",
      "fill": "#E8F1FF",
      "line": "#3B82F6",
      "lineWeight": 0.018,
      "fontSize": 10,
      "rounding": 0.08
    },
    {
      "id": "module-b",
      "type": "ellipse",
      "x": 5.2,
      "y": 0.8,
      "width": 1.4,
      "height": 1.0,
      "text": "Fusion",
      "fill": "#EEFCEB",
      "line": "#2F855A",
      "fontSize": 9
    },
    {
      "id": "arrow-1",
      "type": "connector",
      "from": [4.5, 1.6],
      "to": [5.2, 1.3],
      "line": "#334155",
      "lineWeight": 0.014,
      "endArrow": 13
    },
    {
      "id": "caption-a",
      "type": "text",
      "x": 0.5,
      "y": 3.0,
      "width": 3.5,
      "height": 0.4,
      "text": "Panel A",
      "fontSize": 12,
      "textColor": "#111827"
    }
  ]
}
```

## Coordinate system

- `x` and `y` use a top-left origin in inches.
- `width` and `height` use inches.
- The drawing script converts these values to the Visio page coordinate system internally.

## Page block

- `page.name`: target page name
- `page.width`: page width in inches
- `page.height`: page height in inches

## Supported shape types

### rect

Fields:
- `x`, `y`, `width`, `height`
- optional `text`, `fill`, `line`, `lineWeight`, `fontSize`, `textColor`, `rounding`

### ellipse

Fields:
- `x`, `y`, `width`, `height`
- optional `text`, `fill`, `line`, `lineWeight`, `fontSize`, `textColor`

### text

Fields:
- `x`, `y`, `width`, `height`, `text`
- optional `fontSize`, `textColor`, `line`, `fill`

Text blocks are drawn as invisible or lightly styled rectangles containing editable Visio text.

### line

Fields:
- `from`: `[x, y]`
- `to`: `[x, y]`
- optional `line`, `lineWeight`, `endArrow`

### connector

Same as `line`. Use `connector` when the semantic role is an arrow between modules.

## Color format

Use hex colors such as `#3B82F6`.

## Notes

- Keep the first spec simple. Add fine-grained shapes only after the page-level layout is stable.
- If exact font family control becomes important, use the scripted draft first and then refine in the live Visio UI.

