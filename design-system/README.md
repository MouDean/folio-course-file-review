# Folio design system

This folder is the design system for Folio: the colours, type, spacing and status components that the live page uses. The same content is published as a Claude Design System artifact (private to Mousa): https://claude.ai/artifact/Y6GsVuX6yNBXRQCZX8tvkw

## Files

```
tokens.json          the source of truth: colours (light + dark), type styles, spacing, radii, shadows
tokens.css           generated from tokens.json (CSS custom properties + type classes); do not edit by hand
build_tokens.py      writes tokens.css and checks that src/top.html uses the same colours
components/          one folder per component: README.md (guidelines) + preview.html (rendered in the artifact)
  StatusDot/         the 24px square for one checklist item
  StatusPill/        checker status and reviewer decision labels
  KpiTile/           the number tiles in the summary band
  Cover/             the cover image shown at the top of the artifact
```

## Changing a colour or style

1. Edit `tokens.json`.
2. Make the same change in the `:root` and dark-theme blocks at the top of `src/top.html` (the page keeps its own copy so it stays a single file).
3. Run `python3 design-system/build_tokens.py`. It rewrites `tokens.css` and reports any colour in `src/top.html` that no longer matches `tokens.json`.
4. Rebuild the page with `python3 build.py`.
5. Ask Claude to update the Design System artifact so it matches.

Every text colour pair must keep at least 4.5:1 contrast in both themes (3:1 for decorative glyphs).

The PDF and Excel exports (`src/export.js`, palette `XC`) still use the earlier v1 colours. Update them separately if the reports should match.

---

## Brand book

Folio is the course file review tool of the Department of Education (College of Arts, Education and Social Sciences, Abu Dhabi University). This system describes how Folio looks and speaks, so new screens, exports and emails stay consistent with the live page.

### Content fundamentals

- **Plain, exact English.** Short sentences, no jargon. Say what is missing and where: "Previous offering syllabus missing (add it, or a note if this is the first delivery)."
- **Name things the way the department does.** *Course file*, never "folder" or "record". The seven items keep their numbers and short names: 1 Checklist, 2 Syllabus, 3 Materials, 4 Assessment, 5 Attendance, 6 CRC, 7 SET.
- **Status words are fixed.** Complete · In progress · Not due yet · Overdue for the checker; Not reviewed · Ready to review · Approved · Needs revision for the reviewer. Don't invent synonyms.
- **Calm about empty folders.** An empty course file before its due date is "Nothing uploaded yet", not an error. Red is only for overdue or missing.
- **Student work is counted, never listed.** No student names on screen, in exports or in emails.
- **Emails** open with "Dear Dr. <first name>," and have no sign-off, because Outlook adds the signature.

### Visual foundations

**Colour.** A quiet green-grey neutral ground (`bg`, `surface`, `raised`, `sunk`) carries everything. The one brand colour is ADU maroon (`accent`). It marks the primary action (Scan SharePoint), the semester chip, focus rings, sort arrows and the selected filter, and nothing else. Three signal hues carry meaning and are never decoration:

| Meaning | Token | Use |
|---|---|---|
| Meets the model | `ok` / `ok-soft` | filled dot ✓, Approved, complete bars |
| Partly there | `part` / `part-soft` | outlined dot !, In progress, due within 7 days |
| Missing / overdue | `miss` / `miss-soft` | outlined dot ✕, Overdue, Needs revision |
| Not uploaded yet | `none` / `none-soft` | dot ·, dashed "Nothing uploaded yet" bar |

Status is never shown by hue alone. Each dot also has a glyph (✓ ! ✕ ·) and a fill style (solid, outlined, soft), so it still reads in greyscale and for colour-blind reviewers. Every text pair meets 4.5:1 in both themes.

**Type.** *Archivo* (display) for the product name, headings, course codes in the sheet and KPI numbers. *Public Sans* (body) for everything else. *JetBrains Mono* only for copyable emails and key hints. Numbers use tabular figures. Eyebrows and table headers are 11px uppercase with 0.06em tracking, in `muted`.

**Space and shape.** Four spacing steps from `space-1` 4px to `space-5` 20px. Radii grow with the size of the container: `radius-sm` for dots, `radius-md` for controls, `radius-lg` for tiles and cards, `radius-xl` for panels, and `radius-pill` for status pills. Cards are separated with hairline borders (`line`) and a faint `shadow-sm`. Only the side sheet and the toast use `shadow-lg`.

**Layout.** One working screen, from top to bottom:

1. Header: brand, semester and scan controls.
2. Action toolbar.
3. Due-date cards.
4. Summary band (KPI tiles beside progress by program).
5. Filter bar.
6. Grid of course files × the seven items, grouped by program. Program groups fold, and the columns sort.

Detail opens in a right-hand side sheet with previous/next buttons. On phones (≤760px) the grid rows become cards.

### Components

- **StatusDot**: the 24px square for one checklist item (ok, part, miss, none).
- **StatusPill**: a rounded state label for the checker status or the reviewer decision.
- **KpiTile**: a number tile that filters the grid when pressed, with a 3px signal bar at the bottom edge.

### Iconography

Folio uses simple 1.8px-stroke line icons (24px grid, round caps) next to button labels: add, mail, download, document, settings, refresh, chevrons. The glyphs ✓ ! ✕ · are the only symbols inside status dots. Folio has no illustrations or emoji.

### Logo

The Folio mark is two stacked pages behind a folder. It is drawn in `accent` with a `surface` interior and an `ok` check, as an inline SVG in the page header. It isn't uploaded here yet.
