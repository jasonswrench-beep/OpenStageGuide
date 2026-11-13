# OpenStage Guide

A tool for helping training and development professionals create facilitator guides and participant workbooks.

You write the facilitator guide once, in modules, lessons and blocks (what to say, ask, show and do, with timing and notes). The program checks it, keeps its tables and cross-references current, and makes the participant workbook from the same guide, with nothing meant for the facilitator in it. Handouts, forms and worksheets are written into the guide and used by the facilitator's script and by the workbook. The workbook is made as a Word file; taking the guide itself out as a Word file, when the author chooses to, is planned and is not built yet.

**Status: prototype 0.1.** This is the first form of the program: a local web server and a page that run in a browser on your own computer. It is not yet a program you install, its editing surface is not finished (words and whole paragraphs can be written, chosen, moved, deleted, copied and pasted on the page; tables, pictures and new blocks are still added and changed with the tools beside it), and nothing has been tried yet with a screen reader or in Word. [docs/PROGRAM-SHELL-2026-11-13.md](docs/PROGRAM-SHELL-2026-11-13.md) says exactly what exists and what does not.

## Run it

Node 20 or later. No other software.

```
npm ci --ignore-scripts
node scripts/project-server.mjs --folder "C:\path\to\guides" --port 3850
```

Open <http://127.0.0.1:3850/>. With no `--folder`, guides go in `Documents\OpenStage Guides`. A guide is one file ending in `.guide-project.json` (plain JSON, so it can be opened, kept and backed up like any file). The server answers only on this computer, never takes a path from the page, and keeps the previous copy of a guide as `.bak` the first time it saves after the guide is opened.

## Tests

```
npm test
```

Two checks drive a real browser and need [Playwright](https://playwright.dev/) installed somewhere (this repository does not depend on it). Set `PLAYWRIGHT_DIR` to the `playwright` package folder, then:

```
npm run check:program
npm run check:project
```

After changing the list of icons in `program/icons.json`, rebuild `program/ui/ui-icons.js` with `npm run icons -- --lucide <path to OpenStage Author's editor/web/icons/lucide.json>`.

One test is skipped unless `OPENSTAGE_AUTHOR_VALIDATOR` names the `validate.py` of OpenStage Author (the project file it makes for e-learning is checked against that program's own validator).

## What is where

| Folder | What it holds |
| --- | --- |
| `program/` | The program's page: the suite's navy frame, menus, command search, welcome page, dialogs, and the drawn guide. |
| `src/` | The guide's core: the model, the Word-document XML it is made of, forms, the participant workbook and its leak check, review, e-learning screens, assessments, and the tools' pane. |
| `scripts/` | The local server (`project-server.mjs`), the build snapshot it serves from, the icon builder, and the browser checks. |
| `public/` | The tools' page (`taskpane.html`), its style, the training icons and the product icon. |
| `tests/` | The automatic tests. |
| `docs/` | The plan, the forms and handouts specification, the review of example guides, and the specifications the first screens were built from (`docs/background`, `docs/features`). |

## How it relates to the Word add-in

OpenStage Guide began as an add-in for Microsoft Word, in the repository [OpenStage-Guide](https://github.com/jasonswrench-beep/OpenStage-Guide). The two share one core: the same guide model, the same tools and the same checks. This repository holds the program only. The add-in's manifest, installer, certificate and publishing scripts stay in that repository. The code that drives Word (`src/word.mjs`, and the Word half of `src/guide-word.mjs`) is still here because the tools' pane reads and writes its document through one small adapter; the program uses the adapter that keeps the guide in the project file and never loads Office.js.

## Licence

Apache License 2.0 for the code and the product icon (see `LICENSE` and `NOTICE`). The original fourteen training icons are by Jason S. Wrench under CC BY-NC-SA 4.0; the terms are in `public/assets/ICON-LICENSE.txt`. Third-party components are listed in `NOTICE`.
