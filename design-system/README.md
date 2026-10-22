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

### Visual foundations — "The Archive" (v3, October 2026)

Folio looks like a well-kept paper archive: a dark-teal cover over warm paper, with folder tabs for tiles and program groups. The v2 look is kept in `backup_v05_design2/` and the v4 "Bento workspace" in `backup_v08_bento/`.

**Layout.** One page, top to bottom:

- **Cover:** a dark-teal header (`cover`) holding the Folio mark and serif wordmark, the department line, the semester chip and the last-scan time. The ADU logo sits on the right above the program menu and the saffron Scan SharePoint button. A saffron scalloped edge runs along the bottom.
- **Toolbar:** Add course file, Email reminders, Download report, Export PDF / Excel, then the Auto / Light / Dark switch, the Sound switch and Settings.
- **Deadline strip:** one card per term due date with a calendar tile; the next deadline gets a saffron edge and badge. Winter and Summer show a single "Short semester" card.
- **Department overview:** eight folder-tab KPI tiles. **Progress by program:** a stacked bar per program.
- **Filters and table:** the course files grouped by program under dark-teal folder tabs that fold open and closed, then the legend and "How the checker works".

Clicking a deadline, a tile or a program bar filters the table. On phones the cover stacks and the table becomes cards.

**Colour.** Warm paper (`bg`) with cream panels (`surface`). Deep teal (`accent`, `fill`, `cover`) is the main colour, and saffron (`gold`) marks the scalloped edge, the semester chip, the Scan button and the next deadline. The signal colours are green `ok`, amber `part`, red `miss` and grey `none`. Dark mode swaps to charcoal-teal with mint and saffron accents. Every text pair meets 4.5:1 in both themes.

**Type.** Plus Jakarta Sans for everything, including the wordmark, panel titles and numbers (weights 700–800), chosen because it is plain and easy to read. IBM Plex Mono for course codes and dates.

**Sound** (made in the page, no audio files, muted with the Sound switch in the toolbar):
- a two-note chime when a scan completes, and a low tone if it fails;
- a soft rising pair when a course file is approved or a comment is resolved;
- a click for smaller actions (settings saved, marked as sent or emailed);
- a short arpeggio the first time a program reaches 100% complete, or every course file in the semester is approved (once per program per semester, per browser).

### Components

- **StatusDot**: the 24px square for one checklist item (ok, part, miss, none).
- **StatusPill**: a rounded state label for the checker status or the reviewer decision.
- **KpiTile**: a folder-shaped number tile that filters the grid when pressed; its tab shows the signal colour.

### Iconography

Folio uses simple 1.8px-stroke line icons (24px grid, round caps) next to button labels: add, mail, download, document, settings, refresh, chevrons. The glyphs ✓ ! ✕ · are the only symbols inside status dots. Folio has no illustrations or emoji.

### Logo

Two marks sit on the cover. On the left is the Folio mark: two stacked pages behind a folder, with saffron outlines on teal and a sand-coloured check, drawn as inline SVG. On the right is the white Abu Dhabi University logo (`assets/adu-logo-white.png`, taken from adu.ac.ae), shown 42px tall above the scan controls. A large, faint crop of the ADU woven mark (`assets/adu-mark-white.png`, 9% opacity) replaces the earlier star lattice as the cover's background. Only use the ADU logo in its original white or colour versions, and never recolour or distort it. A brand board for this identity is in Canva ("Folio brand identity board").

### Type scale (October 2026)

The root size is 14.5px and body text is 13px. The wordmark is about 33px, KPI numbers about 24px, panel titles about 15px, and table text 12–13px with codes at 11.6px.
