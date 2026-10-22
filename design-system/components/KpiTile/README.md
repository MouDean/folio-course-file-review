# KpiTile

A number tile in the summary band. Pressing it filters the grid to the course files it counts, and pressing it again clears the filter.

- `raised` fill, 1px `line` border, `radius-lg`. The number uses the `kpi-number` style and the label uses `eyebrow`-sized text in `muted`.
- A 3px bar at the bottom edge, in the tile's signal colour, shows its meaning: `ok` for Complete, `part` for In progress, `accent` for Ready to review, `miss` for Overdue.
- When the count is 0, the number turns `muted` and the bar is hidden, so the tiles with a non-zero count stand out.
- When pressed: `accent-soft` fill and an `accent` ring, with `aria-pressed="true"`.
- Lay out eight tiles in a 4 × 2 grid. On phones use 2 columns.
