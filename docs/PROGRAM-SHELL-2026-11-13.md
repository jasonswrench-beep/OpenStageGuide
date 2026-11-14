# The program's shell: what is built

2026-11-13, branch `app/spike`. This is step 4 of `STANDALONE-APP-PLAN-2026-11-13.md` in its first form: the suite's frame around the pane of the add-in, running in a browser with a small local server. It is not yet the program the author will install (that is a Tauri window with a signed setup, step 6), and it is not yet an editor (step 2).

## Run it

```
node scripts/project-server.mjs --folder "C:\path\to\guides" --port 3850
```

Open `http://127.0.0.1:3850/`. With no `--folder` the guides go in `OpenStage Guides` inside the person's Documents folder as Windows has it (read from the registry; the folder in OneDrive when OneDrive holds Documents), which is the folder the window uses. (Until 2026-11-14 this was the Documents folder inside the home folder, which is a different folder on a computer where OneDrive holds Documents.) The activities the person saved in the activity library, the workbook's look and the agenda's start time are kept in `OpenStage Guide settings.json` in that folder, by the server (`src/program-settings.mjs`); they used to be kept in the page's local storage, which the window's page does not keep from one start to the next, because its address has a new port every time. The server listens on this computer only, answers only to its own address, and never takes a path from the page: a guide is named by its file name in the folder. `--file` still opens the bare pane on one project, as the spike did (`/taskpane.html?project=1`).

## What it has

- **The frame**, after the suite's other programs (the copies are credited in `NOTICE`): a navy bar with the product's name, the open guide's name and whether it is saved (in words), Search, and the three things done all the time (Review guide, Update guide, Make workbook, the last one filled); a row of fourteen menus, **File, Edit, Insert, Format, View, Start, Sections, Activities, Slides, Assessment, eLearning, Check & Plan, Make, Help** (the tool menus were regrouped on 2026-11-14: the Word add-in's ribbon needed groups of three, the program does not), built from the add-in's 21 tools and four that only the program has, Add Lesson (the lesson half of Add module), Learning Outcomes (the block form on Objective), E-Learning Cues (the block form on the e-learning kinds) and Accessibility (the document checks, which Review guide no longer lists) (`program/ui/commands.js` reads `PROGRAM_MENUS` in `src/guide-tools.mjs`, so a command cannot exist without its tool); **Ctrl+K** finds any command by its name, **Ctrl+/** lists the keys, **F10** goes to the menus.
- **The Guide's own colour pair** (accepted by the author on 2026-11-13; adding it to the suite's design guide in the Author repository still needs his separate yes): a blue, `#0E7490` for anything you press or read, `#155E75` under the pointer, `#E3F4F7` for the chosen item, `#67E8F9` for the name on the navy bar. The contrast of each pairing is measured in `tests/program-files.test.mjs`. There is a dark scheme of its own, like OpenAssess's.
- **A welcome page** (New guide..., Open guide..., the guides there are, as buttons), **New guide** and **Open guide** dialogs in the suite's dialog look (a Close button that does what Escape does, focus on the first thing to fill in, focus back where it was).
- **An open guide**: the guide drawn as a page (a read-only view made from the project's paragraphs: headings, blocks with their labels, notes, forms with their lines, fields, tables and check boxes, cards, the print line), and the add-in's tools beside it in a frame. The guide's document, and its saving, live in the program's page; the page draws the guide again as soon as anything in it changes.
- **The tools' page** is the pane's own page with only Office.js left out (it sends a framed page away), served at `/program/tools`; the pane puts on the program's look (`program/pane.css`). It has no document of its own: it uses the program's by message (`src/guide-remote.mjs`), is told how saving stands, and passes the suite's keys (Ctrl+K, Ctrl+/, F10) on to the program's page. Files it makes (the workbook, a backup, an agenda) are real downloads.

## Safety it keeps

The page writes text and attributes only (never markup); it reaches nothing but its own server; its pages carry a content security policy with no inline script or style; changing anything (making, opening or closing a guide, saving) needs this page on this address in the form the program sends; a guide's name never reaches outside the guides folder; a guide that could not be read is never written over; each guide's first save after it is opened keeps the old file as a dated copy (`.<date>-<nn>.bak`, the newest three); every read and save carries the token the program gave the page when the guide was opened or made, so a page left over from another tab is refused instead of writing over the guide that is open (`stale-guide`, 409), and it stops saving; saves run one at a time through a temporary name nobody can guess; every HTML page carries a policy and `frame-ancestors 'self'`. Tests: `tests/program-commands.test.mjs`, `program-preview.test.mjs`, `program-server.test.mjs`, `program-files.test.mjs`, and the real-browser check `scripts/check-program-browser.mjs` (needs Playwright; 40 checks, among them that only the program's page stores the guide and that a workbook downloads), and `tests/guide-remote.test.mjs` for the message link between the page and the tools.

## What it does not have yet

- **An editing surface (E1 and E2)**: the words of a paragraph can be changed on the page, paragraphs can be chosen together, moved, deleted, added, cut, copied, pasted and given another kind, from the keyboard or the **Edit** menu (see `EDITING-SURFACE-DESIGN-2026-11-13.md`). Not yet: tables, pictures, the time line's minutes, Bold and Italic (E3 and E4); until then they are done with the tools beside the guide.
- **The tools beside it keep the add-in's look inside** (green boxes, small type). Restyling them in the suite's look, and moving each tool's buttons into the menus, comes with the editing surface and the Forms tool.
- **No Forms tool** (F3): adding it to the add-in's tool list would change the add-in's ribbon, so the program needs a tool list of its own.
- **No export menu**, no welcome-page sample guide, no Save as, no recent-list controls, no About page beyond a short dialog.
- **The program's name is OpenStage Guide** (answered by the author, 2026-11-13). **The file ending** (`.guide-project.json`) is still a decision for the author.
- **Nothing has been tried in Word, with a screen reader, or on a second computer.**

## Zoom and the page's height (2026-11-14)

- **Zoom is the document's.** The View menu (Zoom in, Zoom out, Actual size; Ctrl+Alt+=, Ctrl+Alt+-, Ctrl+Alt+0; the mouse wheel with Ctrl held over the guide) draws the
  guide's page between 50% and 300%, as a word processor does. It is CSS `zoom` on the guide's page, set by the variable `--doc-zoom` on the program's page (`program/style.css`,
  `program/ui/app.js`, `src/zoom.mjs`); the menus, the tools and the dialogs keep their size, nothing in the guide or in a Word file changes, and the level is kept in the settings
  file (`osg.zoom.v1`). Ctrl+= and Ctrl+Shift+= stay subscript and superscript in the guide's words, which is why zoom has Ctrl+Alt. The window's own (web view) zoom is left off,
  and the window has no part in this.
- **The whole page never scrolls.** A long guide once made the program's page scroll as well as the guide's pane (and left empty space below the guide and the tools): the hidden
  hint for screen readers inside the guide is positioned absolutely, and the guide's pane was not its containing block, so the pane did not clip it. `.work-guide` is now
  `position: relative`, and `check:program` fails without it.
