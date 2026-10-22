# StatusPill

A rounded label for a course file's checker status or the reviewer's decision.

- Checker status: **Complete** (`ok-soft`/`ok`), **In progress** (`part-soft`/`part`), **Not due yet** (`none-soft`/`muted`), **Overdue** (solid `miss`, `surface` text).
- Reviewer decision: **Approved** (solid `ok`), **Ready to review** (`accent-soft`/`accent` with an `accent-line` ring), **Needs revision** (`miss-soft`/`miss`), **Not reviewed** (plain `muted` text, no fill, so it fades into the background).

Use 12px, weight 600, `radius-pill`. Use solid fills only for final states (Approved, Overdue).
