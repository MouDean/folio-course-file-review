# Folio — Course File Review

Folio is a single-page tool for reviewing university course files on SharePoint against a model course file. It was built for the Department of Education (College of Arts, Education and Social Sciences, Abu Dhabi University) and runs as a Claude artifact with a shared database and a Microsoft 365 connector.

## What it does

- **Scans SharePoint** course-file folders (by program, one course file, or everything) and checks each against the seven sections of the model course file: 1 Checklist, 2 Syllabus, 3 Materials, 4 Assessment, 5 Attendance, 6 CRC, 7 SET.
- **Applies department rules**: graded samples per assessment and level (High / Medium / Low) — 1 for classes under 30, 2 for 30–60, 3 for more than 60 — and the syllabus must state the section and campus.
- **Skips unchanged and approved course files** to keep scans fast; a "Full" option re-reads everything.
- **Reviewer workflow**: approve / needs revision, accept an item as done, comments with open / fixed / resolved status.
- **Semesters**: Fall, Winter, Spring and Summer are kept separately, with a switcher at the top.
- **Schedule import** (.xlsx) to update class sizes, terms and campuses.
- **Emails**: per-course-file review emails and batch reminders, opened in Outlook for the reviewer to send.
- **Reports**: quick Excel report, plus formal PDF or styled Excel exports for a course file, an instructor, a program or the whole department, with charts and a signed sign-off block.
- Light / dark / auto theme.

Student work is only counted, never listed — no student names appear anywhere.

## Repository layout

```
src/top.html      page markup and styles
src/classify.js   course-file classification against the model course file
src/app.js        app logic (scanning, review, emails, semesters, settings)
src/export.js     PDF / Excel export (inserted into app.js at /*@@EXPORT@@*/)
build.py          builds dist/folio.html from src/
dist/folio.html   the built page that is published as the artifact
```

## Build

```
python3 build.py
```

This writes `dist/folio.html`. Publish that file as the artifact (same URL) to update the live tool.

## Runtime requirements

The page uses these artifact capabilities: `db` (shared database: courses, config, people, comments), `user` (profile), `mcp` (Microsoft 365 `read_resource` for SharePoint reads) and `downloads` (saving reports).
External libraries load on demand from cdnjs: SheetJS 0.18.5 (quick report, schedule import) and pdf-lib 1.17.1 (PDF export). The styled Excel export needs no library.

## Notes

- SharePoint site and drive identifiers are set at the top of `src/app.js`.
- Settings (due dates, reviewer name, email, signature) live in the shared database, not in the code.
