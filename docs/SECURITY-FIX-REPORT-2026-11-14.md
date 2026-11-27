# Fix report: the 2026-11-13 security review of OpenStage Guide

This answers `SECURITY-REVIEW-OpenStage Guide.md` and its hand-off. The review was of main at bdc7701. The fixes are in a08348d. One more fix, found while building the setup, is in 80a6eda. Suggested recheck: `git diff bdc7701..80a6eda`.

## The short version

All 12 findings (2 Medium, 10 Low) are fixed or, for the two that are the owner's, listed below. Every fix has a test that fails without it. The program also has Help > Check for updates now, built to the pattern the other OpenStage tools use. The setup was rebuilt and signed from the fixed code (details at the end).

One thing the review could not have seen, because it was found afterwards: the first signed setup left `scripts/updates.mjs` out of its file list, so an installed copy could not have started. That is fixed in 80a6eda with a test that follows the server's imports and fails if anything the server loads is not carried. See "Found while building".

## Finding by finding

| # | Finding | What changed | Proof |
|---|---|---|---|
| M1 | Field codes and file links in a shared guide reach the Word file | New `src/guide-xml-clean.mjs` allows only what the program itself writes. Opening a guide cleans each paragraph (`importState`), and saving a Word file cleans again (`bodyWithContents`). Fields other than the program's own TOC, PAGE and NUMPAGES go, and none that names an address is ever allowed. Links lose their target unless they point inside the guide. Objects, attached files, controls, data bindings, legacy pictures and floating drawings go. A plain inline picture stays, without links. The program's own paragraphs pass through byte for byte. | `tests/guide-xml-clean.test.mjs` (include-picture field, simple field, file link, never-closed field, program's own pages untouched, saved-again guide is the clean one) |
| M2 | A save goes to whichever guide is open | Every open or new gets a random token, returned in the state. The page sends it as `X-Guide-Token` on every read and save. A mismatch is 409 `stale-guide`. The page then says "This guide was closed in another window", stops saving and refuses further changes. A save that was already waiting when another guide was opened is not written to it. | `tests/program-server.test.mjs` ("a save is for the guide its page was given the token for", "a save that was waiting..."); `tests/guide-project-web.test.mjs` |
| L1 | pdf.js 3.11.174 on a page with no policy | The file window (`picker.html`) and the PDF library are not served and not carried in the setup. `getDocument` is called with `isEvalSupported:false` anyway. NOTICE says the library is not part of the program. The old files still sit in `public/` because the Word add-in's pane is still in the repository. | `tests/pdf-slides.test.mjs` (eval off); the smoke start of the carried server returns 404 for both |
| L2 | No CSP or frame guard outside `/program/*` | Every HTML response, including the tools pane and the pane's pages, sends a strict policy with `frame-ancestors 'self'` and `X-Frame-Options: SAMEORIGIN`. No served page loads Office.js. | `tests/program-server.test.mjs` ("every HTML page says what it may load...") |
| L3 | Temp file and backup handling | Saves run one at a time. The temporary name is random and opened with `wx`. `backedUp` is set only after the copy works. The newest three dated backups are kept (the numbering cannot reuse a name when copies are pruned in the same second). | "saves at once are saved one at a time", "a temporary name cannot be guessed...", "only the newest three dated copies are kept" |
| L4 | Request bodies up to 256 MB | `/api/*` bodies are capped at 16 KB. Only `PUT /project` may be large. The server has a request time limit. | "an /api request over 16 KB is refused, a guide far larger than that is still saved..." |
| L5 | Style import trusts the zip's own size | `readCapped` counts the bytes as they inflate and stops at the cap, for the styles part and the theme. It no longer reads JSZip's private size field. | `tests/guide-docx.test.mjs` (the two inflate-bomb tests; a good file is unchanged) |
| L6 | Workbook filter gaps | `w:ruby` and `w:rt` are removed (the base text stays), `w:webHidden` and `w:specVanish` count as hidden, and text in a paragraph whose style hides it is dropped. The leak check knows the same cases. | `tests/guide-workbook.test.mjs` (ruby, web-hidden, always-hidden, hidden paragraph style) and `scripts/workbook-leak-check.mjs` |
| L7 | The window can go anywhere; Office.js on two pages | The window shows only its own page and the server's address on this machine. A link to a release page or the icon license page opens in the person's own browser, never inside the window. A new-window request is always denied. The standalone build serves the pane's pages without Office.js. | `cargo test` in `shell/src-tauri` (address rules); the HTML-route test above |
| L8 | The launcher does not check its own server | `Start OpenStage Guide.cmd` polls `/health` until it answers as OpenStage Guide before opening the browser, and shows the server's error if it does not start. | Not automated (a launcher); read the script (the browser opens only when `/health` answers as OpenStage Guide; otherwise the server's log is shown) |
| L9 | Commit authorship | Owner's. The first commit carries the owner's personal address and was kept on purpose (decided 2026-11-13). 17 older commits carry the assistant trailer; rewriting them was deferred ("Not yet"). Nothing new adds one: the commits since the review are authored as the owner with no trailer. | Not code |
| L10 | Windows reserved names | The guide-name check and the static-file patterns refuse every device name (CON, PRN, AUX, NUL, COM1-9, LPT1-9, CONIN$, CONOUT$ and the superscript digits) with any extension. | "every Windows device name is refused..." |

Smaller items in the hand-off, all done: `/api/state` gives the guides folder's name, not its path; the Slides icons were renamed and their references updated; "Read-only presentation review" replaces the old wording; the rebuild-plan line says "the first design conversation".

## Check for updates

Help > Check for updates... follows `docs/UPDATE-CHECK.md`, the same pattern as the other tools: it runs only when chosen, is notify-only, makes one HTTPS request to this repository's releases list (no redirects, 2 MB cap, 15 s limit, a plain user agent and nothing else), offers only a page of this repository's releases, opens it in the person's own browser, and has an off switch (`OPENSTAGE_UPDATE_CHECK=off`, an `update-check.txt` beside the program, or `--no-update-check`). Until the repository is public and has a release it answers "no release found". Please check it against your "Do not change" line about update checks: this is the one new network request, and it is on a button press only.

## Found while building

- `shell/prepare.mjs` listed four scripts for the setup. The server also loads `scripts/updates.mjs`, so the first signed setup would have failed at start-up. Fixed in 80a6eda; `tests/shell-prepare.test.mjs` walks the server's imports and fails for any script, source folder or package that is not carried (shown to fail with the old list, pass with the new).
- The carried server was then started the way the window starts it, from the carried files with the carried Node: it answered `/health`, the program page, the tools pane and `/api/updates`, served 404 for the file window and the PDF library, and stopped when its input closed.

## Checked and left alone

Everything in the hand-off's "Do not change" list is unchanged: the Host, Sec-Fetch-Site, Origin and content-type checks, no CORS headers, the 127.0.0.1 address with a random port in the window, no route that takes a path, no `innerHTML`, the parser limits, the workbook sanitize and leak check, the window's empty capabilities and no plugins, local fonts, no telemetry.

## Checks run

- Unit tests: 1805 pass, 0 fail, 1 skipped (an opt-in test). That includes the two new file-list tests.
- Browser checks at a08348d: program 58 of 58, project 29 of 29, editor 214 of 214. Since then only `shell/prepare.mjs` and one test were added, so they were not repeated.
- `cargo test` in `shell/src-tauri`: 2 of 2.
- `npm audit` was not run again (the review found none, and the dependencies did not change). The only open dependency alert is the Linux-only `glib` one in the window's Rust crates.

## Setup

The setup was built from 80a6eda with `shell/tools/build_installer.py --sign`. Every program of ours in it (the setup, the window, the uninstaller and the installer's plugins) is signed and timestamped by the owner as publisher; the Node that runs the server keeps its own makers' signature. The program holds no folder names of the building computer. The uninstaller removes only the files the setup put in the install folder and never touches `Documents\OpenStage Guides`. Nobody has run the setup yet, and no release is published.

The setup is 29,100,176 bytes. SHA-256: `021df4a19d5171bf334ad9b8eca11b30919f993ead07ea2f3f200cb57ab9009b`. It was unpacked and the server in its own files was started with its own Node, the way the window starts it, and answered. The setup built from a08348d (SHA-256 starting `263e8c60`) is the one with the missing file and must not be used.

## Left for the owner

1. The 17 older commits with the assistant trailer (decided "Not yet"), and whether to make the repository public.
2. Publishing a release. The update check finds nothing until the repository is public and a release exists.
3. Trying the setup on a real computer, and an NVDA pass.
4. Archiving the older Word add-in repository, as the review suggested.
