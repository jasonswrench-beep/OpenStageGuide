# Stage 4: accessibility and print

Written 2026-10-31. The plan's test for this stage (the rebuild plan of 2026-10-30 (kept outside the repository)): "A review list separates errors, warnings and human checks. Screen-reader pass (only when you are away, as agreed). Grayscale and print check. A reviewed PDF export route." The author asked for building to continue without sign-off documents, so this is a plan, not a request for approval; the choices it makes are small and easy to change.

## What already exists

- **Review guide** groups what it finds as *To fix* (errors), *To look at* (warnings) and *To check yourself* (human checks), from `guideIssues` and the structure rules in `src/guide-check.mjs` (title, language, headings, table headers, pictures, typed lists, empty paragraphs, contrast of styles). It is explicitly not Word's own checker and not a screen-reader test.

## What this stage adds

1. **The workbook is checked as it is made** (`src/workbook-checks.mjs`). The same structure rules run on the finished workbook's parts. The result is shown under the Workbook screen's message as *to fix*, *to look at* and *to check yourself*. It does not block the workbook (an accessibility finding is something to fix in the copy, which the author can edit), unlike the leak check, which does.
2. **A print check.** For the chosen look, in colour or in black and white: body text and the accent as greys against the page (4.5 to 1), the body size, the margins, and a plain note that the icons are coloured pictures and carry no meaning on their own. Nothing in the workbook is told by colour alone: that is a human check.
3. **A reviewed PDF route.** On the Review screen, *Save as PDF* runs the review first. With anything to fix, no PDF is made and the screen says so. With nothing to fix, it asks Word for its own PDF (`getFileAsync` with the PDF type), checks the file for the markers of a tagged PDF (a structure tree, "marked", a language, a title) and saves it as a download. If Word did not tag it, the screen says how to make Word do so (File > Export > Create PDF/XPS > Options > Document structure tags for accessibility). The tag check reads the file's plain text, so a PDF that stores its catalogue compressed is reported as "could not tell", not as untagged.

## What stays with a person (all in `TESTING-TODO.md`)

- Word's own Check Accessibility on the guide, the workbook and the assessment sections.
- A screen-reader read-through (NVDA). It logs everything spoken on the whole computer, so it is done only when the author is away, after asking, and the logs are deleted.
- A printed page or print to PDF in colour and in black and white.
- That the PDF Word makes is really tagged and reads well.

## Not in this stage

Repairing an untagged PDF (that is the author's separate PDF remediation project), a logo, and changing Word's own export options from the add-in (Office.js cannot).
