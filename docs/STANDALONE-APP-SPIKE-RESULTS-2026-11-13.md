# OpenStage Guide as a program of its own: spike results

2026-11-13, branch `app/spike` (from `rebuild/core` at e06e80f). Local only. The plan is `STANDALONE-APP-PLAN-2026-11-13.md`.

## The question

Can the pane's screens run on a guide kept in a project file, with no Word anywhere, without rewriting them?

## The answer so far: yes

The pane already talks to its document through one interface, and the in-memory document the browser preview uses implements all of it. The spike adds a way to take that document's state out, put it back, and keep it in a file. The pane, the guide model, the cue and review rules and the Word-format builders are unchanged.

## What was built

| File | What it is |
| --- | --- |
| `src/guide-project.mjs` | The project file format, version 1. Pure: `cleanGuide`, `serializeProject`, `parseProject`, `ProjectError`, limits. Refuses a wrong kind of file, a newer version, a wrong shape, anything oversize; drops unknown fields and prototype-reaching keys. |
| `src/guide-word.mjs` | `exportState()` and `importState()` on the in-memory document. `importState` checks that every paragraph's XML is a Word paragraph or table before it takes anything. |
| `src/guide-project-adapter.mjs` | Wraps the document in the same interface and saves a moment after any change. Refuses changes before the project has opened or after it failed to open. Never writes over a file it could not read. A failed save keeps the guide marked "not saved" and tries again with the next change. |
| `src/guide-project-web.mjs` | The browser's read and store calls (`GET` and `PUT /project`). |
| `scripts/project-server.mjs` | A small server for one project file. Loopback only; refuses a wrong Host, a foreign Origin or cross-site fetch, a wrong content type, a project that does not validate, anything oversize. Writes through a temporary file and a rename, and keeps the file as found at start once as `.bak`. |
| `src/app.mjs` | `?project=1`: the preview, with the guide read from and saved to the project file. The header says whether it is saved. |

Run it: `node scripts/project-server.mjs --file "C:\path\My guide.project.json" --port 3850`, then open `http://127.0.0.1:3850/taskpane.html?project=1`.

## Checks

- 26 new tests (`tests/guide-project.test.mjs` 8, `tests/guide-project-adapter.test.mjs` 12, `tests/project-server.test.mjs` 6). The main one: a guide written in the pane is saved, read by a new session, and **the Word-format package built from it is byte-identical** to the original's, and a pane mounted on it shows the same guide. Others cover debouncing, a failing disk, an unreadable or newer file, a damaged paragraph, and changes before opening. The server tests cover the guards (another host name, another web site, a cross-site fetch, a form post, a wrong method), a project that does not validate, the one-time backup and the size limit.
- A real-browser check (headless Chromium, 29 checks): write a guide, see the file; reload and see the guide; keep working; a broken file is explained and left byte for byte as it was; a repaired file opens; seven refusals by the server and none of them changed the file. The script is `scripts/check-project-browser.mjs` (it needs Playwright: set `PLAYWRIGHT_DIR`).
- The full suite (`node --test tests/*.test.mjs`): everything passes except **one test that already fails on `rebuild/core` without any of this work**: `tests/guide-author.test.mjs` ("the project passes OpenStage Author's own validator"). The narration export to Author writes project format 0.48.0, and Author's current validator wants 0.49.0 (Author moved on 2026-11-10). Author's editor offers to update such a project, so this is a test and export version to bring up to date, not a broken file. Not touched here.
- `tests/static-export.test.mjs` (the privacy guard) names the one new request, the project file read and store on the program's own server. **`PRIVACY.md` needs a section for the program before it is released**; it is not edited here because the statement is the add-in's.

## What this does not prove

- **No editing surface yet.** The guide is still built by the pane's forms. Typing in the guide itself is the next big piece.
- **The look is the pane's old look.** The suite shell (navy bar, menus, welcome page) is not started. The screenshot shows the Word pane's header and its "Loaded build identity" strip, which will not exist in the program.
- **Pictures** are referenced by the paragraphs' XML (icons are loaded from the site). The project file does not yet carry pictures of its own.
- **The file is JSON.** A real project needs to hold pictures, so it will become a zip; the format check and the adapter would stay.
- The browser's own XML parser flags unbalanced tags; the test's xmldom does not. The check therefore also requires every paragraph's root to be a Word paragraph or table.
- Nothing was tried in Word, and nothing here touches the add-in's behaviour: Word mode is as before.

## Next

1. The suite shell around the pane (the OpenAssess pattern), with the colour decision.
2. The editing surface (styled paragraphs, typing, Undo), with a screen reader pass.
3. The project as a zip with its pictures, and a file extension.
4. Outputs: web page, PDF (is a tagged PDF possible?), Word as the optional one.
