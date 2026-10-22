# StatusDot

A 24px square that shows the result for one of the seven checklist items in a course file.

| State | Look | Glyph |
|---|---|---|
| ok | solid `ok`, `surface` glyph | ✓ |
| part | `part-soft` fill, 2px inset `part` ring | ! |
| miss | `miss-soft` fill, 1.5px inset `miss` ring | ✕ |
| none | `none-soft` fill, `none` glyph | · |

Always pair the hue with its glyph. Put the item's finding in the `title` so it shows on hover. When a course file has nothing uploaded, replace all seven dots with one dashed "Nothing uploaded yet" bar instead of seven grey dots.
