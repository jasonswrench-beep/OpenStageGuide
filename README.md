# OpenStage Guide

A tool for helping training and development professionals create facilitator guides and participant workbooks.

You write the facilitator guide once, in modules, lessons and blocks (what to say, ask, show and do, with timing and notes). The program checks it, keeps its tables and cross-references current, and makes the participant workbook from the same guide, with nothing meant for the facilitator in it. Handouts, forms and worksheets are written into the guide and used by the facilitator's script and by the workbook. The workbook is made as a Word file, and so can the guide itself be (File, Save as Word file), with its styles; the guide's styles can also be saved as a Word template, and the styles of a Word file taken in. Four families of free, easy-to-read fonts (Atkinson Hyperlegible, Lexend, Source Sans 3, Source Serif 4) come with the program and are carried inside a Word file that uses them. OpenDyslexic is offered too, by name only: it is not included, so it has to be installed on the computer, and a Word file only names it.

**Status: alpha 0.1.** The program has its own window and a signed Windows setup, on the [releases page](https://github.com/jasonswrench-beep/OpenStageGuide/releases), and it also runs from source in a browser on your own computer (below). Its editing surface is not finished (words and whole paragraphs can be written, given a look (bold, italic, color, size, font, a character style), aligned, put in styles of your own, chosen, moved, deleted, copied and pasted on the page; tables are filled in on the page, and new blocks are still added and changed with the tools beside it), and nothing has been tried yet with a screen reader. The Word files the program makes pass the Word schema validator and open in a real Word; that is one machine's Word, read-only. [docs/PROGRAM-SHELL-2026-11-13.md](docs/PROGRAM-SHELL-2026-11-13.md) says exactly what exists and what does not.

A guide can also be started from a **session template** (16 ready-made sessions, with 46 timed activities in the activity library), and the run sheet, the preparation checklist, five layouts of the workbook, card decks and a presentation file are made from the same activities; what was made says when the guide has changed since ([docs/SESSION-MATERIALS-2026-11-23.md](docs/SESSION-MATERIALS-2026-11-23.md)).

## Run it

Node 20 or later. No other software.

```
npm ci --ignore-scripts
node scripts/project-server.mjs --folder "C:\path\to\guides" --port 3850
```

On Windows, double-clicking `Start OpenStage Guide.cmd` does both steps (it installs what is needed the first time, starts the server in a minimized window, and opens the program in your browser once it answers; if it does not start, what it said is shown). Otherwise open <http://127.0.0.1:3850/>. With no `--folder`, guides go in `OpenStage Guides` inside your Documents folder, wherever Windows has put it (in OneDrive, when OneDrive holds Documents): the same folder the program's window uses, so a guide made in a browser is there in the window too. Beside the guides the program keeps one small file, `OpenStage Guide settings.json`, with the activities you saved in the activity library, the look you chose for the workbook and the agenda's start time, so they are there the next time. A guide is one file ending in `.guide-project.json` (plain JSON, so it can be opened, kept and backed up like any file). The server answers only on this computer, never takes a path from the page, checks that a save comes from the page that has the guide open (a page left over from another tab is refused, so it can never write its guide over another), saves one at a time, and keeps a dated copy of the guide (`.<date>.bak`, the newest three) the first time it saves after the guide is opened. **View > Zoom in, Zoom out and Actual size** (Ctrl+Alt+=, Ctrl+Alt+-, Ctrl+Alt+0, or the mouse wheel with Ctrl held over the guide) change how big the guide's page is drawn on the screen, from 50% to 300%; the menus and the tools keep their size, nothing in the guide changes, and the level is kept for next time. **Help > Check for updates...** asks GitHub whether a newer release is published, only when you choose it, and installs nothing ([docs/UPDATE-CHECK.md](docs/UPDATE-CHECK.md)).

## Tests

```
npm test
```

Six checks drive a real browser and need [Playwright](https://playwright.dev/) installed somewhere (this repository does not depend on it). Set `PLAYWRIGHT_DIR` to the `playwright` package folder, then:

```
npm run check:program
npm run check:project
npm run check:settings
npm run check:closing
npm run check:design
npm run check:sessions
```

After changing the list of icons in `program/icons.json`, rebuild `program/ui/ui-icons.js` with `npm run icons -- --lucide <path to OpenStage Author's editor/web/icons/lucide.json>`.

One test is skipped unless `OPENSTAGE_AUTHOR_VALIDATOR` names the `validate.py` of OpenStage Author, or Author is in its usual place on the author's PC (`C:\Dev\openstage-author`): the project file Guide makes for e-learning (schema 0.49.0, Author's current format) is checked against that program's own validator.

## What is where

| Folder | What it holds |
| --- | --- |
| `program/` | The program's page: the suite's navy frame, menus, command search, welcome page, dialogs, and the drawn guide. |
| `src/` | The guide's core: the model, the Word-document XML it is made of, forms, the participant workbook and its leak check, review, e-learning screens, assessments, and the tools' pane. |
| `scripts/` | The local server (`project-server.mjs`), the build snapshot it serves from, the icon builder, and the browser checks. |
| `public/` | The tools' page (`taskpane.html`), its style, the training icons and the product icon. |
| `shell/` | The program's window (Tauri): starts the server and shows it, and makes the Windows setup. Not needed to use the program in a browser. See `shell/README.md`. |
| `tests/` | The automatic tests. |
| `docs/` | The plan, the forms and handouts specification, the review of example guides, and the specifications the first screens were built from (`docs/background`, `docs/features`). |

## How it relates to the Word add-in

OpenStage Guide began as an add-in for Microsoft Word, in the repository [OpenStage-Guide](https://github.com/jasonswrench-beep/OpenStage-Guide). The two share one core: the same guide model, the same tools and the same checks. This repository holds the program only. The add-in's manifest, installer, certificate and publishing scripts stay in that repository. The code that drives Word (`src/word.mjs`, and the Word half of `src/guide-word.mjs`) is still here because the tools' pane reads and writes its document through one small adapter; the program uses the adapter that keeps the guide in the project file and never loads Office.js.

## Licence

Apache License 2.0 for the code and the product icon (see `LICENSE` and `NOTICE`). The original fourteen training icons are by Jason S. Wrench under CC BY-NC-SA 4.0; the terms are in `public/assets/ICON-LICENSE.txt`. Third-party components are listed in `NOTICE`.
