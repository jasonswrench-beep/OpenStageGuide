# The program's shell: what is built

2026-11-13, branch `app/spike`. This is step 4 of `STANDALONE-APP-PLAN-2026-11-13.md` in its first form: the suite's frame around the pane of the add-in, running in a browser with a small local server. It is not yet the program the author will install (that is a Tauri window with a signed setup, step 6), and it is not yet an editor (step 2).

## Run it

```
node scripts/project-server.mjs --folder "C:\path\to\guides" --port 3850
```

Open `http://127.0.0.1:3850/`. With no `--folder` the guides go in `Documents\OpenStage Guides`. The server listens on this computer only, answers only to its own address, and never takes a path from the page: a guide is named by its file name in the folder. `--file` still opens the bare pane on one project, as the spike did (`/taskpane.html?project=1`).

## What it has

- **The frame**, after the suite's other programs (the copies are credited in `NOTICE`): a navy bar with the product's name, the open guide's name and whether it is saved (in words), Search, and the three things done all the time (Review guide, Update guide, Make workbook, the last one filled); a row of eight menus, **File, Guide, Write, Bring in, E-learning, Check, Make, Help**, built from the same 21 tools the add-in's ribbon is (`program/ui/commands.js` reads `src/guide-tools.mjs`, so a command cannot exist without its tool); **Ctrl+K** finds any command by its name, **Ctrl+/** lists the keys, **F10** goes to the menus.
- **The Guide's own colour pair** (accepted by the author on 2026-11-13; adding it to the suite's design guide in the Author repository still needs his separate yes): a blue, `#0E7490` for anything you press or read, `#155E75` under the pointer, `#E3F4F7` for the chosen item, `#67E8F9` for the name on the navy bar. The contrast of each pairing is measured in `tests/program-files.test.mjs`. There is a dark scheme of its own, like OpenAssess's.
- **A welcome page** (New guide..., Open guide..., the guides there are, as buttons), **New guide** and **Open guide** dialogs in the suite's dialog look (a Close button that does what Escape does, focus on the first thing to fill in, focus back where it was).
- **An open guide**: the guide drawn as a page (a read-only view made from the project's paragraphs: headings, blocks with their labels, notes, forms with their lines, fields, tables and check boxes, cards, the print line), and the add-in's tools beside it in a frame. The guide's document, and its saving, live in the program's page; the page draws the guide again as soon as anything in it changes.
- **The tools' page** is the pane's own page with only Office.js left out (it sends a framed page away), served at `/program/tools`; the pane puts on the program's look (`program/pane.css`). It has no document of its own: it uses the program's by message (`src/guide-remote.mjs`), is told how saving stands, and passes the suite's keys (Ctrl+K, Ctrl+/, F10) on to the program's page. Files it makes (the workbook, a backup, an agenda) are real downloads.

## Safety it keeps

The page writes text and attributes only (never markup); it reaches nothing but its own server; its pages carry a content security policy with no inline script or style; changing anything (making, opening or closing a guide, saving) needs this page on this address in the form the program sends; a guide's name never reaches outside the guides folder; a guide that could not be read is never written over; each guide's first save after it is opened keeps the old file as `.bak`. Tests: `tests/program-commands.test.mjs`, `program-preview.test.mjs`, `program-server.test.mjs`, `program-files.test.mjs`, and the real-browser check `scripts/check-program-browser.mjs` (needs Playwright; 40 checks, among them that only the program's page stores the guide and that a workbook downloads), and `tests/guide-remote.test.mjs` for the message link between the page and the tools.

## What it does not have yet

- **No editing surface.** The guide is read-only in the program; it is still written through the tools' forms.
- **The tools beside it keep the add-in's look inside** (green boxes, small type). Restyling them in the suite's look, and moving each tool's buttons into the menus, comes with the editing surface and the Forms tool.
- **No Forms tool** (F3): adding it to the add-in's tool list would change the add-in's ribbon, so the program needs a tool list of its own.
- **No export menu**, no welcome-page sample guide, no Save as, no recent-list controls, no About page beyond a short dialog.
- **The file ending** (`.guide-project.json`) and the program's name are decisions for the author.
- **Nothing has been tried in Word, with a screen reader, or on a second computer.**
