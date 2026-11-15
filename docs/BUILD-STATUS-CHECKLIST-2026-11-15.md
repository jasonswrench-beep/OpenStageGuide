# OpenStage Guide: what is finished and what is not (checklist)

Brought up to date 2026-11-15 after the whole kit was built (the author's decision: no build, signing or push until the kit is done). Made 2026-11-15, at the author's request ("do we have a formal checklist for what is finished in this build and what still needs to be finished?"). There was none: the facts were spread over the plan, the slice records in the specs, the release notes and the author's notes. This file collects them in one place. Keep it up to date at every release (section 9 says how).

**Reading it.** `[x]` is finished and has a test or a check that fails if it breaks. `[~]` is partly done, and the line says which part is missing. `[ ]` is not done. *(per docs)* means it comes from a record written when the piece was built and was not looked at again for this list. The newest published setup is alpha 8 (a plain release; alphas 1 to 7 stay marked "pre-release" on GitHub). **Alpha 9 is built and signed (2026-11-15) and is not pushed or published.** It holds what was built after alpha 8: **Save as PDF**, **Bullets**, **moving a whole module or lesson**, **pictures inside table cells**, **copying a picture between guides**, and the Layout tool's **icon column** and **slide picture beside the script**.

## 1. The program and its installer

- [x] A window of its own (Tauri), with the program's own local server inside, started and stopped with the window.
- [x] A signed setup (Jason Wrench, timestamped) built by `python -I tools/build_installer.py --sign`; installs for one user with no administrator password; updates in place and keeps guides.
- [x] Published on GitHub as releases (alphas 1 to 7 went out marked "pre-release"; from alpha 8 a release is a plain release, not marked so); the setup has a checksum file; downloads without signing in.
- [x] **Help > Check for updates...** (asks GitHub only when chosen, installs nothing); can be turned off by `OPENSTAGE_UPDATE_CHECK=off` or an `update-check.txt`.
- [x] **Clear previous guides** (home screen): leaves the guides off the list, deletes nothing; **Show the N guides left off** brings them back; Open guide lists every guide. Built after alpha 7, in alpha 8 (built and signed, setup SHA-256 `189f62ac...c71f`, build `5105a1e7`, in `installers-signed-2026-11-15-alpha8`; pushed as tag b89c28f and published as a plain release).
- [x] **Alpha 9 built and signed, not published** (2026-11-15, release label on 7b08e59, built from bf7ad88; setup SHA-256 `7ce745cd...4199`, 29,190,216 bytes, build `df04cfed`, in `installers-signed-2026-11-15-alpha9`): checked as in section 9 steps 1 to 6 (source and unpacked). Pushing and publishing wait for the author's yes.
- [x] Closing the window saves the open guide first.
- [x] What the author sets (activity library, workbook look, agenda start, zoom) is kept in `OpenStage Guide settings.json` in the guides folder.
- [x] One guides folder for the window and the browser version (`Documents\OpenStage Guides`).
- [ ] The setup has not been tried on a second computer, or by anyone but its author.
- [ ] A virtual-machine check of the installer and the uninstaller (clean Windows) has not been done.
- [ ] Alpha 6 is not published (built and signed; the setup is in `installers-signed-2026-11-14-alpha6`). **It will not be.** Alpha 7 holds alpha 6's fixes and the whole kit; it was built and signed on 2026-11-15 on the author's go (setup SHA-256 `3fc537c5...a8a0`, build `0bd2e09a`, in `installers-signed-2026-11-15-alpha7`) and checked as in section 9 steps 2 to 6; it was pushed and published as v0.1.0-alpha.7 (marked "pre-release" then) on the author's yes, downloaded again without signing in (same hash, valid signature), and an alpha 5 copy's Check for updates was told it is out. Nothing of the kit is committed to the public repository yet: the commits are local only.

## 2. The suite look and the shell

- [x] Navy bar (name, open guide, saved or not, Search, three buttons), a row of 14 menus, welcome page with recent guides, Ctrl+K command search, Ctrl+/ key list, F10 to the menus, dialogs.
- [x] The Guide's own blue pair, colors checked for contrast *(per docs)*.
- [x] Zoom for the guide's page (View menu, formatting bar, Ctrl+Alt+= and friends, Ctrl and the wheel), 50% to 300%, kept for next time.
- [x] Menu lists stay inside the window (alpha 6) and every tool has one name in its menu, in More tools and as its heading (alpha 6).
- [~] The menu row fits on one line from about 1366 pixels of window width; narrower, the last menus move to a second line (they still work).
- [~] Windows high contrast mode: looked at here with the browser engine's forced-colors emulation (the program, the Layout tool, the web page): borders, black on white, every control outlined. Also looked at in the dark setting, which found and fixed headings that were navy on navy. *Not tried in Windows' own high contrast mode on a real PC.* The tools beside the guide stay a light panel in the dark setting (readable, but not matched to the dark page).
- [ ] **A screen-reader pass (NVDA) of the whole program has not been done.** It waits for the author's go (it records everything spoken on the computer).

## 3. Writing in the program (the editing surface)

- [x] Typing, Undo and Redo, styled paragraphs, whole-block commands (delete, move, copy, cut, paste), paste of several lines.
- [x] Tables you fill in (cells, check boxes, rating grids), the Insert menu (lines, box, fields, check list, table, rating grid, cards, picture, source line), pictures with a description and a caption.
- [x] Looks for a paragraph (alignment) and for words inside a sentence (the formatting bar's bold, italic, underline, strike, sub, super, color and clear), the Styles panel, styles taken from a Word file, a Word template made from the guide's styles.
- [x] Four families of free fonts shipped and embedded in Word files; OpenDyslexic is offered by name only.
- [x] The Review checks how words look (color contrast, tiny text, bold lines that look like headings), in **Accessibility**.
- [x] **Edit > Add table column right, Delete table column, Join cell with the next, Split joined cell, Table title and description...**: columns keep the table as wide as it was, the header row is not joined, cells merged downwards are left alone with the reason; checked in a real browser.
- [~] **Edit > Picture details...** changes a picture's description, whether it is only decoration, its width (one of three) and its caption after it is in. Not done: a free size.
- [x] **Edit > Picture in this cell...** puts a picture inside a table cell (kept no wider than the cell's room), changes it and takes it out; **Picture details...** works on it too. Unit tests and a browser check; *the Word file with a picture in a cell was not opened in a real Word*.
- [x] **Copying a picture between guides:** the copy buffer belongs to the page and carries the pictures, so copy in one guide and paste in another brings the picture along. Unit tests and a browser check.
- [x] **Format > Bullets** (Ctrl+Shift+L, and a button on the formatting bar) makes the selected lines a bulleted list and back; all bulleted lines share one list; the Word file, the web page and the guide's own view draw it. *Not opened in a real Word.*
- [x] **Moving a whole module or lesson:** drag its heading (a line shows the drop place), or Alt+Shift+Up / Down, or Edit > Move module or lesson up / down; everything under it goes along, the numbers are made again, and it is one step of Undo. *The drag was driven in Chromium, not in the real window.*
- [ ] The editing surface is not finished: new blocks are added and changed with the tools beside the page.

## 4. The tools (29: 21 from the add-in, 8 only in the program)

- [x] All 29 open from their menus, show the piece of the Guide screen they name, and have a test and a browser check. Add lesson, Learning outcomes, E-learning cues, Accessibility, Forms, Lesson details, Reference pages and Layout exist only in the program.
- [x] Review guide (what the guide says) and Accessibility (the document's checks) are two screens.
- [x] Timing and agenda, Learning alignment, Delivery versions, Appearance, Generate materials (Training at a Glance, Graphic Cues page, module overview), Participant workbook with its leak check, Assessments, E-learning screens, Slides import and update, Activity library.
- [x] **File > Save a backup...** gives a copy of the guide's own project file (words, pictures, styles, settings) named for the guide and the day; **Restore a backup...** reads one and opens it as a new guide of its own, so nothing is written over. Both are the program's own commands, and the tools of the same names ask the program to do them.
- [x] **A PDF made by the program: File > Save as PDF...** (see section 7). The Review screen's button is called "How to make a PDF": it reviews the guide and then says how to make a tagged PDF, by the program's own command or through a Word file.
- [x] The Help tool lists the program's own keys (the add-in's keys and its link to instructions are gone).
- [x] Ctrl+Z and Ctrl+Y pressed in the tools (not in a box being typed in) are the guide's Undo and Redo.
- [~] The Word add-in's words are gone from the pane's messages ("choose Edit, then Undo", "on the page", "Make the agenda as a Word file"). Left: the typed-list message of the Review ("use Word's list tools instead") and the older panels of the add-in that the program never shows.

## 5. Forms and handouts (the spec's slices)

- [x] **F1** model and numbering (by unit, letter, sequence, none), bookmarks and citations, Update guide renumbers.
- [x] **F2** elements (writing lines and box, field rows, check list, table, rating grid, source line) and the workbook keeps or leaves out a form by its audience.
- [x] **F4** self-scored instrument (statements, a grid to add up by scale, what each total means).
- [x] **F5** card sets (2, 4, 6 or 8 to a page).
- [x] **F6** print line ("Print: 1 per person, Letter, portrait.") and the copies-to-make arithmetic.
- [x] **F7** the Review checks for forms (missing, duplicate, unnamed, empty, no header row, audience, print line).
- [x] **F3 the Forms tool** (in the working tree, not in a setup): the list of forms with number, audience and print line; New form; the numbering choice; how many people, teams and tables, with the copies to make; 15 templates (reading handout, worksheet, note-taking grid, checklist, rating form, self-scored instrument, observer form, evaluation form, action plan, commitment card, role cards, match-up cards, two-by-two, force-field, planning checklist); and the **"Which form" choice on the Handout cue**, so the script cites a form by choosing it (Update guide keeps the name right; editing such a block keeps the citation).
- [x] **F1b** a "Forms used" column on Training at a Glance (only when blocks cite forms), and "Participant handouts: ..." and "Facilitator forms: ..." lines in a module's overview.
- [x] **F8** a landscape form is a section of its own in the guide's Word file, the participant workbook (with its own footer) and the handout pack; the **handout pack** makes one Word file of the forms for everyone from a lesson, a module or the whole guide. *Not checked in a real Word or with the Word schema validator (its Python module is not installed here); checked by tests and as well-formed XML.*
- [x] Review guide's warning about print lines says where to give the number of people, teams or tables (the Forms tool), and clears when it is given.

## 6. The rest of the kit the plan calls for (plan section 3)

Built 2026-11-15; what each slice does is in `docs/KIT-SPEC-2026-11-15.md` ("As built"). Each has unit tests; the ones with a screen have browser checks.

- [x] **K1** Lesson details: purpose, materials, method, range of minutes, interaction, group size, new day, optional and "choose n of m", facilitator only, read by the Review.
- [x] **K2** A wider time model: a stated length, days with a clock that starts again, ranges, set-up and clean-up; the Review compares typed times with the stated length.
- [x] **K3** Script mechanics: expected answer, optional, "slide or flip chart" as one instruction, five more blocks and five more techniques (their icons are stand-ins drawn here, the author's to replace).
- [x] **K4** Reference pages: glossary and references kept in order, figure captions with source numbered by Update guide, pull quotes, the licence and permissions page.
- [x] **K5** The planning kit: timeline, technology and materials, room set-up, roles grid (facilitator only), pre-work, follow-up; the built Materials list page.
- [x] **K6** Layouts for the Word file: each module on a new page, a running header (title or module), mirrored footers with the author's words, a one-page summary, mixed orientation (from F8), **each block's icon in a column at the left, and a slide's picture beside its script** (a borderless table of one row for each such block; the guide on the page does not change). *Checked as zip and markup and through the program; never opened in a real Word.*
- [x] **K7** Import: **File > Import a Word file...** (a file this program or the add-in made) and **File > Paste text...**, each as a new guide beside the others.

## 7. Outputs

- [x] The participant workbook and the facilitator's copy, made from the guide; the handout pack (section 5).
- [x] A guide and a workbook saved as Word files (File > Save as Word file), passing the Word schema validator and opening in a real Word (one machine's Word, read-only) *(per docs)*.
- [x] **File > Save as web page...**: one file with its style sheet and pictures inside, no address of anywhere else, a language, a skip link, one main region, headings with addresses and a contents that links to them, table headers, a dark setting, forced colors and a print style; printed to a tagged PDF it keeps its structure and language *(measured here)*.
- [~] **File > Print...** (Ctrl+P) opens the window's print dialog with a print style (the guide alone, on white, over as many sheets as it needs). *Measured here:* the PDF Chromium makes by itself is **not tagged**; asked to tag, it is. *Not checked:* what the window's own print dialog saves. The route through Word stays too: save the Word file, then make the PDF in Word.
- [~] **File > Save as PDF...** makes a PDF itself: the web page's text goes to Microsoft Edge in its headless mode (Windows 11 has Edge), and the file that comes back is saved. *Measured here:* the PDF Edge makes this way is **tagged** (headings, paragraphs, language, title); the server reads it back for a structure tree and the program says so if there is none. The page is refused if it holds a script; one PDF at a time; a plain message when Edge is not there. *Not checked:* a Windows 11 with no Edge, and what a screen reader makes of the result.
- [~] A tagged-PDF check: done for the engine's PDF call, the web page (above) and the PDF the program makes, with pypdf; not done for a PDF made from the program's Word files, and PDFRemediator remains the fallback.

## 8. Words, testing and release process

- [x] 1,966 automatic tests (one opt-in test skipped) and 5 window tests; browser checks of the program (184, with the Forms tool, the handout pack, backup and restore, undo from the tools, Word import, pasted text, the web page, Save as PDF, print, the kit's tools and the icon column), of saving a guide (29), of the editor (274, with table columns, joined cells, picture details, pictures in cells, copying between guides, bullets and moving a module), of settings (7) and of closing (6). All pass on the working tree of 2026-11-15 after alpha 8.
- [x] The Word files with the icon column and the slide picture beside the script (icons only, slides only, both) pass the Word schema validator (the docx skill's `validate.py`); the icon-column file made through the program passes it too. *Not opened in a real Word.*
- [x] Every build is run through the steps in section 9 before it is published.
- [ ] Flaky checks seen while building and made steadier: the editor check's "Update guide numbers it" (the button is pressed again), the program check's wait after a restore (it waits for the restored guide). A real race was found and fixed: a restore or an import was put into the new guide before it had finished opening (the program now waits for it, and the save, web page, print and backup commands wait for the guide to open).
- [ ] The program has been used by nobody but its author (apart from the manual's author, who tried every tool in a scratch copy and found the problems listed in the release notes).
- [ ] The program manual (written by the manual session) is on alpha 3; it will be refreshed for alpha 5 and 6 when the author says.
- [ ] Questions waiting for the author: the file name ending, whether the add-in stays maintained next to the program, and where a public download page would live.

## 9. The steps every build goes through (the release checklist)

0. **The kit is finished and the author has said go** (this build: alpha 7, which includes alpha 6's fixes).
1. All automatic tests and the five browser checks pass on the source.
2. The release label is bumped and committed; the signed setup is built from the committed code (`shell/tools/build_installer.py --sign`).
3. The setup's signature, and the window's and the uninstaller's, are valid and timestamped.
4. The setup is unpacked: every program file is byte for byte the file of the commit, and `release.txt` carries the label.
5. The five browser checks pass again on the unpacked files with the setup's own Node.
6. The signed window is started from the unpacked setup and looked at (menus, About build number).
7. Only on the author's yes: push, publish the release (a plain release, not marked pre-release) with the setup and its checksum, download it again without signing in (hash and signature), and ask for updates from a copy of the previous release.
8. The record is added to `shell/README.md` and this list is updated; the manual session is told what changed.
