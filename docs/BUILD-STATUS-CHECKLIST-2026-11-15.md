# OpenStage Guide: what is finished and what is not (checklist)

Made 2026-11-15, at the author's request ("do we have a formal checklist for what is finished in this build and what still needs to be finished?"). There was none: the facts were spread over the plan, the slice records in the specs, the release notes and the author's notes. This file collects them in one place. Keep it up to date at every release (section 9 says how).

**Reading it.** `[x]` is finished and has a test or a check that fails if it breaks. `[~]` is partly done, and the line says which part is missing. `[ ]` is not done. *(per docs)* means it comes from a record written when the piece was built and was not looked at again for this list. The newest published setup is alpha 5; alpha 6 is built and signed but not published; the Forms tool below is in the working tree and in no setup yet.

## 1. The program and its installer

- [x] A window of its own (Tauri), with the program's own local server inside, started and stopped with the window.
- [x] A signed setup (Jason Wrench, timestamped) built by `python -I tools/build_installer.py --sign`; installs for one user with no administrator password; updates in place and keeps guides.
- [x] Published as GitHub pre-releases (alpha 1 to 5); the setup has a checksum file; downloads without signing in.
- [x] **Help > Check for updates...** (asks GitHub only when chosen, installs nothing); can be turned off by `OPENSTAGE_UPDATE_CHECK=off` or an `update-check.txt`.
- [x] Closing the window saves the open guide first.
- [x] What the author sets (activity library, workbook look, agenda start, zoom) is kept in `OpenStage Guide settings.json` in the guides folder.
- [x] One guides folder for the window and the browser version (`Documents\OpenStage Guides`).
- [ ] The setup has not been tried on a second computer, or by anyone but its author.
- [ ] A virtual-machine check of the installer and the uninstaller (clean Windows) has not been done.
- [ ] Alpha 6 is not published (built and signed; the setup is in `installers-signed-2026-11-14-alpha6`).

## 2. The suite look and the shell

- [x] Navy bar (name, open guide, saved or not, Search, three buttons), a row of 14 menus, welcome page with recent guides, Ctrl+K command search, Ctrl+/ key list, F10 to the menus, dialogs.
- [x] The Guide's own blue pair, colors checked for contrast *(per docs)*.
- [x] Zoom for the guide's page (View menu, formatting bar, Ctrl+Alt+= and friends, Ctrl and the wheel), 50% to 300%, kept for next time.
- [x] Menu lists stay inside the window (alpha 6) and every tool has one name in its menu, in More tools and as its heading (alpha 6).
- [~] The menu row fits on one line from about 1366 pixels of window width; narrower, the last menus move to a second line (they still work).
- [ ] Windows high contrast mode: the plan requires it to keep working; no check of it in the program is recorded.
- [ ] **A screen-reader pass (NVDA) of the whole program has not been done.** It waits for the author's go (it records everything spoken on the computer).

## 3. Writing in the program (the editing surface)

- [x] Typing, Undo and Redo, styled paragraphs, whole-block commands (delete, move, copy, cut, paste), paste of several lines.
- [x] Tables you fill in (cells, check boxes, rating grids), the Insert menu (lines, box, fields, check list, table, rating grid, cards, picture, source line), pictures with a description and a caption.
- [x] Looks for a paragraph (alignment) and for words inside a sentence (the formatting bar's bold, italic, underline, strike, sub, super, color and clear), the Styles panel, styles taken from a Word file, a Word template made from the guide's styles.
- [x] Four families of free fonts shipped and embedded in Word files; OpenDyslexic is offered by name only.
- [x] The Review checks how words look (color contrast, tiny text, bold lines that look like headings), in **Accessibility**.
- [ ] Adding and deleting a table's columns, merged cells, editing a table's caption and description after it is in.
- [ ] Changing a picture's description, size or caption after it is in (delete and insert again); pictures inside table cells; copying a picture between guides.
- [ ] A bullet toggle; dragging a whole module or lesson (the Course outline tool moves them).
- [ ] The editing surface is not finished: new blocks are added and changed with the tools beside the page.

## 4. The tools (26: 21 from the add-in, 5 only in the program)

- [x] All 26 open from their menus, show the piece of the Guide screen they name, and have a test and a browser check. Add lesson, Learning outcomes, E-learning cues, Accessibility and Forms exist only in the program.
- [x] Review guide (what the guide says) and Accessibility (the document's checks) are two screens.
- [x] Timing and agenda, Learning alignment, Delivery versions, Appearance, Generate materials (Training at a Glance, Graphic Cues page, module overview), Participant workbook with its leak check, Assessments, E-learning screens, Slides import and update, Activity library.
- [ ] **Save a backup... and Restore a backup...** do not work as their names say (copy the guide's own file instead).
- [ ] **Save as PDF** in Review guide does not make a PDF (save as a Word file and make the PDF in Word).
- [ ] The Help tool's key list and its link to instructions do not work in the program.
- [ ] Ctrl+Z pressed while the keyboard is in the tools does not undo the guide (use Edit, Undo).
- [ ] Several tools still use words from the Word add-in ("click in your document and press Ctrl+Z", "use Word's list tools instead", "Make the agenda in Word").

## 5. Forms and handouts (the spec's slices)

- [x] **F1** model and numbering (by unit, letter, sequence, none), bookmarks and citations, Update guide renumbers.
- [x] **F2** elements (writing lines and box, field rows, check list, table, rating grid, source line) and the workbook keeps or leaves out a form by its audience.
- [x] **F4** self-scored instrument (statements, a grid to add up by scale, what each total means).
- [x] **F5** card sets (2, 4, 6 or 8 to a page).
- [x] **F6** print line ("Print: 1 per person, Letter, portrait.") and the copies-to-make arithmetic.
- [x] **F7** the Review checks for forms (missing, duplicate, unnamed, empty, no header row, audience, print line).
- [x] **F3 the Forms tool** (in the working tree, not in a setup): the list of forms with number, audience and print line; New form; the numbering choice; how many people, teams and tables, with the copies to make; 15 templates (reading handout, worksheet, note-taking grid, checklist, rating form, self-scored instrument, observer form, evaluation form, action plan, commitment card, role cards, match-up cards, two-by-two, force-field, planning checklist); and the **"Which form" choice on the Handout cue**, so the script cites a form by choosing it (Update guide keeps the name right; editing such a block keeps the citation).
- [ ] **F1b** a "Forms used" column on Training at a Glance, and "Participant handouts: 5.2, 5.3" on a module or lesson header.
- [ ] **F8** a form of another orientation in its own section, and a handout pack (all the forms of a lesson, a unit or the programme in one file).
- [ ] Review guide warns about print lines until the number of people, teams or tables is given in Forms (the tool now has the place for it; the warning says where).

## 6. The rest of the kit the plan calls for (plan section 3)

- [ ] Segment, module and unit headers (purpose, materials, method, minute range, optional flag, "choose n of m").
- [ ] A wider time model (ranges, clock start plus duration, set-up before and clean-up after, a start time for each day, group-size rules, interaction mode, a stated programme total, Review of typed times against it).
- [ ] Script mechanics (an expected-answer line after Ask and Debrief, an optional flag, "slide or flip chart" as one instruction, a library of named techniques).
- [ ] Reference pages (glossary, references, figure captions with source credits, pull quote, licence or permission page).
- [ ] The planning kit (timeline and checklists, a materials list by person and team, room set-up, roles grid).
- [ ] Layouts for the outputs (slide thumbnail beside the notes, dividers, running headers, margin icons, mixed orientation, mirrored footers, one-page summaries).
- [ ] Import of a guide made with the add-in; pasted text.

## 7. Outputs

- [x] The participant workbook and the facilitator's copy, made from the guide.
- [x] A guide and a workbook saved as Word files (File > Save as Word file), passing the Word schema validator and opening in a real Word (one machine's Word, read-only) *(per docs)*.
- [ ] A web page made from the guide.
- [ ] **A PDF made by the program.** Whether the window's print-to-PDF gives a tagged (accessible) PDF has not been checked; PDFRemediator is the fallback.
- [ ] A tagged-PDF check of anything the program produces.

## 8. Words, testing and release process

- [x] 1,847 automatic tests (one opt-in test skipped) and 5 window tests; browser checks of the program (126, with the Forms tool), of saving a guide (29), of the editor (217), of settings (7) and of closing (6). All pass on the working tree of 2026-11-15.
- [x] Every build is run through the steps in section 9 before it is published.
- [ ] The program has been used by nobody but its author (apart from the manual's author, who tried every tool in a scratch copy and found the problems listed in the release notes).
- [ ] The program manual (written by the manual session) is on alpha 3; it will be refreshed for alpha 5 and 6 when the author says.
- [ ] Questions waiting for the author: the file name ending, whether the add-in stays maintained next to the program, and where a public download page would live.

## 9. The steps every build goes through (the release checklist)

1. All automatic tests and the five browser checks pass on the source.
2. The release label is bumped and committed; the signed setup is built from the committed code (`shell/tools/build_installer.py --sign`).
3. The setup's signature, and the window's and the uninstaller's, are valid and timestamped.
4. The setup is unpacked: every program file is byte for byte the file of the commit, and `release.txt` carries the label.
5. The five browser checks pass again on the unpacked files with the setup's own Node.
6. The signed window is started from the unpacked setup and looked at (menus, About build number).
7. Only on the author's yes: push, publish the pre-release with the setup and its checksum, download it again without signing in (hash and signature), and ask for updates from a copy of the previous release.
8. The record is added to `shell/README.md` and this list is updated; the manual session is told what changed.
