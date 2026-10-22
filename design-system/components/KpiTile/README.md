# KpiTile

A number tile in the Department overview, shaped like a file folder. Pressing it filters the grid to the course files it counts, and pressing it again clears the filter.

- `raised` fill, 1px `line` border, `radius-lg` on every corner except the top left, where the folder tab sits. The number uses the `kpi-number` style and the label uses `eyebrow`-sized text in `muted`.
- The folder tab on top shows the tile's meaning: `ok-soft` with an `ok` edge for Complete, `part-soft` with a `gold` edge for In progress, `accent-soft` for Ready to review, `miss-soft` with a `miss` edge for Overdue, solid `ok` for Approved and solid `gold` for Open comments.
- When the count is 0, the number turns `muted` and the tab goes plain, so the tiles with a non-zero count stand out.
- When pressed: tile and tab turn `fill` with `fill-ink` text, with `aria-pressed="true"`.
- Lay out eight tiles in a 4 × 2 grid. On phones use 2 columns.
