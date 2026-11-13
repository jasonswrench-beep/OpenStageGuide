# The editing surface: design for the author's review

2026-11-13. This is step 2 of `STANDALONE-APP-PLAN-2026-11-13.md`. It is a design, not code: nothing in it is built. It is written to be read once and answered in a few lines (section 8).

## 1. What it must do

The author writes and changes the whole guide on the page: types the words, changes a paragraph's kind, adds, deletes and moves paragraphs and blocks, fills in tables, forms and fill-in elements, adds pictures with their descriptions, and takes a step back with Undo. Every tool and check that works today keeps working on what was typed. It has to be usable by keyboard and with a screen reader from the first version, not added afterwards.

## 2. What the code already does (found by reading it and measuring two guides)

- **The guide is a list of styled paragraphs.** Each one is Word-format XML plus its words and its style's name. The style carries the meaning (module, lesson, cue, notes, form title). The core reads the guide only through one document interface of about thirty methods; the program's document is an in-memory one that the project file saves and restores (`createDemoAdapter`, `createProjectAdapter`).
- **That interface changes a guide only by whole operations:** insert blocks, delete or move a range (each checked against the words the author last saw), and a few edits to the words of a heading, time line, form title, citation or question. There is no operation that changes the words of one paragraph. The nearest thing, `typeOver`, exists for tests and rewrites only the first piece of text in a paragraph.
- **A table is one block, not many paragraphs.** The document reports a table's cells as paragraphs, but the whole table's XML sits on the first cell and the others carry none. Editing a cell means editing inside that one table.
- **Two JavaScript pages hold the guide today.** The tools run in a frame inside the program's page. The frame owns the document and saves it; the program's page only draws what the frame saved. An editor in the program's page would be a second writer to the same file, and the frame would overwrite its changes with its own stale copy at the next save. This has to be settled before any typing is built (section 3).
- **What paragraphs are made of** (two guides built with the tools: the 20-minute sample and a guide with blocks and every kind of form; 63 ordinary paragraphs, table cells not counted). About 7 in 10 are plain words, some with a bold label run in front ("Facilitator notes:"). About 1 in 6 are cue lines: icon, label, a tab, then words. About 1 in 9 hold something that is not words: a picture, a bookmark, a citation inside a sentence, a tab, a line break, or the contents field. Real guides that were pasted in will differ; this is a floor to plan from, not a forecast.
- **Headings and times are partly generated.** "Module 2:", "Lesson 2.1:", "Handout 1.1:", the time sentence with its running total, and the contents table are written by Update guide. An author who could type over them would break them.

## 3. The first thing to settle: one page, one document

**Recommendation: put the tools into the program's page and remove the frame** (call it E0). The editor, the tools and the save then share one document in one place. This is also where the plan was heading: the tools' buttons are to move into the menus and dialogs, which needs them in the same page. The cost is real: the tools' markup (about 800 lines with `guide-` ids) has to be hosted in a container, their styles scoped so they do not fight the program's, and the server's special page for the frame retired. The 35-step browser check of the program is the test: it must pass with the frame gone.

**Fallback:** keep the frame and build the editor inside it, as part of the tools' page, styled in the suite's look. Nothing about the document changes, but the editor then lives in the page that is meant to disappear, and the work is thrown away when the tools move.

I would do E0 first as a short spike and report before building on it.

**What E0 became (as built, same day).** The spike found a third way that gets the same result with far less risk, so that is what was built. The tools' page and stylesheet are written to own their whole page: they read their elements by id from the document, add page-level key handlers, and their stylesheet styles `body`, `header` and plain element names. Hosting them inside the program's page would have meant scoping all of that. So the frame stays, and the guide's document moved out of it:

- The document, with its saving, now lives in the program's page (`createDemoAdapter` inside `createProjectAdapter`, started when a guide is opened). The page draws the guide from it at once after any change, not after a save and a second read of the file.
- The tools in the frame use it through `src/guide-remote.mjs`: the same document interface, each call carried by `postMessage` and answered. A tool's block is built in the frame and only the finished Word-format text travels. Only the interface's own methods can be called; importing and exporting the whole state, moving the cursor and undoing cannot be reached from the frame. Both ends check who sent a message (the other page, same address).
- One writer: the browser check now records every request that stores the project and fails if any came from the tools' frame.
- The frame goes away tool by tool as the tools become menus and dialogs in the program's page; nothing about the document changes when that happens. The editor (E1) is built in the program's page against the same document.

Two things the move turned up, both fixed:

- **"Make workbook" gave the author nothing.** The in-memory document's `openDocument` and `offerDownload` only recorded what they were given (they exist for the tests), so in the program the workbook, the backup, the agenda and the narration script were "opened" nowhere and never delivered. Now the tools find there is no Word window to open a file in, fall back to their download path, and a real download happens. The browser check downloads a workbook and looks at the file. The tools' wording still says "Word could not open it for you, so it was downloaded"; the words for the program are a later pass.
- **The suite's keys did not work while the author was in the tools.** Ctrl+K, Ctrl+/ and F10 were heard by the program's page only, and a frame is its own document. They are now passed on from the tools to the page.

## 4. How it works for the author

**Rows, and one paragraph in edit at a time.** The guide stays drawn as it is now: headings are headings, forms are forms. Each paragraph is a row. Arrow keys move between rows. Enter, F2 or a click starts editing the row: its words become a text box (a real `textarea`, which is the control screen readers and spell-check handle best) that grows with its words. Escape ends editing and leaves the row selected. This is how outline editors work, and OpenAssess already edits with auto-growing text boxes.

**Typing flows.** While editing: Enter at the end of the words starts the next paragraph and edits it at once; Up on the first line and Down on the last line move into the neighbouring row, still editing; Backspace at the very start joins the paragraph to the one before. Shift+Enter is a line break inside the paragraph. The author can type a whole script without leaving the keyboard.

**What a row holds depends on its style.** The part the author types is the words; the part the program writes is shown as a fixed lead or as a field next to it.

| Style | Author types | Shown, not typed |
| --- | --- | --- |
| Guide title | the title (also the document's title) | |
| Module, Lesson | the title | "Module 2:", "Lesson 2.1:" |
| Time | the minutes, in a number box | "Time: 10 minutes. Cumulative time: ..." |
| Cue and Facilitator cue | the words, with "(5 min)" left in them as now | the icon and the cue's name; a menu on it changes the cue |
| Cue body, notes, normal text, question, option, overview, form text | the words | a label such as "Facilitator notes:" in front |
| Form title | the title, and who it is for | "Handout 1.1:", the bookmark |
| Contents entries | nothing | generated by Update guide |
| A table cell | the cell's words | |

**Changing a paragraph's kind.** A "Paragraph" menu (and a key) lists the styles that make sense and changes the style; the row redraws with the right lead. A bullet toggle makes a paragraph a list item.

**Blocks.** Shift+Up and Shift+Down select whole rows. Delete, cut, copy, paste, Alt+Up and Alt+Down (move) act on them. Pasted text becomes paragraphs, one per line, in the style the author is in; text pasted from Word comes in as plain words (no stray fonts or colours). Mouse drag to move is added later and never replaces the keys.

**Tables and fill-in elements.** A table is one row that opens into a grid: arrow keys move between cells, Enter edits a cell, and add/remove row and column are menu commands. A header row stays a header row. Check boxes in cells stay check boxes. Lines, boxes, field rows and the other form elements are inserted from an Insert menu, made by the builders that already exist, and edited in the same grid or text box.

**Pictures.** Insert picture asks for the file and the description in one step and refuses to finish without a description or a "this picture is only decoration" tick, as the tools do now. A picture row shows the picture, its caption as editable words, and Describe / Size / Remove commands.

**Undo and Redo.** One history, shared with the tools, so Undo reverses whatever was done last, typed or tool-made. Typing in one paragraph is one step until the author leaves the paragraph or pauses. Redo is new.

## 5. What it changes in the code

- **The document interface gets a few new operations, in the program's document only** (the Word adapter is left alone): change one paragraph's words, style or list level; split and join paragraphs; insert paragraphs; edit a table cell; add or remove table rows and columns; redo. Each is checked against the words the editor last saw, like the existing ones, so a stale row is refused and not applied to the wrong paragraph, and each takes one history step.
- **Words are changed in the XML properly, not by pattern.** A paragraph is read into pieces: a leading label, the editable text, the things that are not text (icon, tab, picture, check box, citation, bookmark, field). Changing the words changes the text pieces and leaves the others untouched. This is the one place the editor can corrupt a guide, so it gets the most tests: every paragraph kind the builders make is read into pieces and written back byte-for-byte when nothing changed.
- **Drawing is incremental.** The page now redraws the whole guide after each save. The editor redraws only the rows that changed, so focus and the caret are never torn down while the author types. Rows need stable identities for this; they are made in the view and are not written to the file.
- **The cursor is the focused row.** The core already has a cursor and a selection event that the tools use ("you are in this block"); the editor sets and fires them as focus moves, so the tools keep following the author.
- **New files:** `program/ui/editor.js` (the page), a pure `program/ui/editor-model.js` or `src/` module (paragraph into pieces and back, the style table above, the commands), and tests beside them.

## 6. Order of work

Each step ends with automatic tests, a real-browser check, screenshots for the author, and no push without a yes.

- **E0. One page, one document.** *Done 2026-11-13, in the form described in section 3 ("What E0 became"): the document lives in the program's page and the tools use it by message.* Program check 40 of 40, project check 29 of 29, 1582 tests passing.
- **E1. Words.** Rows, keyboard model, editing the words of plain, heading, time and cue paragraphs, cursor and selection sync, autosave, undo and redo. The paragraph-into-pieces code and its round-trip tests.
- **E2. Structure.** Split and join, insert, delete, move, row selection, copy, cut, paste, the Paragraph menu, lists.
- **E3. Tables and fill-in elements.** The grid, rows and columns, form elements inserted and edited. (This is also the Forms tool, F3.)
- **E4. Pictures with descriptions, and inline Bold and Italic** (if the author wants them, see question 1).
- **E5. A screen reader pass** with NVDA (only when you are away, and I ask first), keyboard-only pass, a big guide (hundreds of paragraphs) for speed and the history's memory, and the polish the passes turn up.

The Word export, the web page and PDF outputs, and the Tauri window and installer come after, as in the plan.

## 7. Risks

- **A text box cannot hold a citation or a check box inside a sentence.** A sentence with a citation in the middle is shown as piece, chip, piece, each piece its own small text box. That is about 1 in 9 paragraphs in the guides measured, and the ones that matter (citations) are rare. If it feels wrong in use, the alternative is a small rich-text box per paragraph with the pieces as chips, which would also give Bold and Italic; it is more work and less certain with screen readers.
- **Screen readers and a page full of text boxes** have to be tried, not assumed. E5 does that. Until then, no claim.
- **Table XML.** Edits inside a table go through the first cell's XML; add or remove row and column are where this will go wrong first. Round-trip tests come before the grid.
- **History memory.** The document keeps a whole copy of the paragraph list for each step. That is fine for a normal guide and wasteful for a very long one; steps are grouped and capped, and E5 measures it.
- **Big change to the page.** E0 touches the part the 35-step check covers. It is a spike for that reason.
- **Edits made outside the program** (in Word after an export) do not come back; this is stated in the plan and stays true.

## 8. For the author to decide

1. **Bold and Italic inside a sentence: day one or later?** My suggestion is later (E4). Today's guides use no direct formatting; the emphasis they have lives in styles (labels, notes). If you want Bold and Italic from the start, say so: they would be character styles, so they survive a restyle, and the editor would need the rich-text box described under the first risk.
2. **Is "one paragraph in edit at a time, with rows and arrow keys" the feel you want,** rather than a single flowing page where the caret runs through everything as in Word? The first is far safer for the file and for screen readers; the second is what Word users expect. You can try the first as soon as E1 is built and say if it is wrong.
3. **E0 as the first step** (tools into the program's page), or the fallback of building the editor inside the frame.
4. **NVDA in E5:** same rule as always: only when you are away, after I ask.

## 9. E1 as built (2026-11-13)

**What it does.** The guide is drawn as rows, one for each paragraph (a table is one row). Arrow keys, Home and End move between rows; Enter, F2 or a click changes the words; Escape stops. While the words are being changed: Enter starts the next paragraph with the words after the caret (the first press in a heading goes on to the next row instead); Backspace at the start joins the paragraph to the one before, or takes an empty paragraph away; Up or Left at the start and Down or Right at the end go to the paragraph before or after; Ctrl+Z, Ctrl+Y and Ctrl+Shift+Z are Undo and Redo for everything, the tools' changes too. Each change is on the document and is saved a moment later; typing in one paragraph is one step of Undo. The tools beside the guide follow the editor's cursor ("Your cursor is in the Ask block").

**What can be changed.** The words of any paragraph made of a lead and plain words: Normal text, headings (the "Module 1:" or "Lesson 1.1:" number is shown beside the box and is not typed), cue lines (the cue's picture and name stay; "(5 min)" is part of the words), notes (the "Facilitator notes:" label stays), body text, list items (Enter keeps the list). The time line, the contents lines, form lines and boxes and fields, tables, and any paragraph that holds a picture, a check box, a citation, a bookmark, a field or a line break cannot be changed by the editor. They are drawn, can be moved to with the keys, say why when Enter is pressed, and are never rewritten: the browser check shows a time line is on disk byte for byte as it was after a run of edits around it.

**Where it is.** `src/guide-edit.mjs` (reading a paragraph into a lead and words, and writing new words, splitting and joining, all pure), `replaceParagraphs`, `redo` and grouped Undo steps in the in-memory document (`src/guide-word.mjs`), `program/ui/editor.js` (the page) and `program/ui/editor-model.js` (which rows are drawn again, what each key does). Tests: `guide-edit`, `guide-edit-adapter`, `program-editor-model` (and `program-preview` for the rows). Real-browser check: `scripts/check-editor-browser.mjs`, 40 checks.

**Things the build found that are worth knowing.**

- Paragraphs with nothing in them were not drawn at all, so a new line made by Enter would have had no row. The editor's view draws them as rows in the look of their style.
- A row whose paragraph did not change is kept as it is when the others are redrawn, so a row that had been a text box had to be redrawn by hand when Enter or Backspace left it. A screenshot showed it; the 40 checks of the first run had not. There is now a check that only one text box exists at a time.
- A forwarded key (F10 pressed while the keyboard is in the tools) reaches the menus by message, so a check that read the focus at once was wrong about half the time. It now waits.

**Not done in E1 (so, E2 onward).** Whole-paragraph and block commands (delete, move, insert with a style, Paragraph menu, bullet toggle, row selection, copy, cut and paste), paste of several lines as several paragraphs (a pasted line break becomes a space for now), the time line's minutes box, tables and the fill-in elements, pictures, Bold and Italic, and the screen reader pass (not run: it needs the author away and his go). No claim is made about screen readers yet: the row and text-box structure follows the pattern in section 4, and a real pass comes in E5.

## 10. E2 as built (2026-11-13)

**What it adds.** Whole paragraphs. The **Edit** menu (Undo, Redo, Cut, Copy, Paste, Add paragraph below, Kind of paragraph..., Move up, Move down, Delete paragraph) and the keys for each: Shift+Down and Shift+Up choose rows together (the choice is shown and announced: "3 paragraphs selected."), Ctrl+A chooses all, Alt+Up and Alt+Down move the chosen rows (or the one the author is on), Delete deletes them, Ctrl+Enter adds a paragraph below and goes into it, Ctrl+C, Ctrl+X and Ctrl+V copy, cut and paste. Everything is one step of the one Undo history. **Kind of paragraph** is a dialog of five kinds that are styles the guide already has (Normal text, Block text, Facilitator note, Heading 1, Heading 2); the title, modules, lessons, form titles, cue lines and labelled notes keep their kind and say why. A paragraph pasted from inside the guide comes back exactly as it was (its kind included); text from anywhere else becomes a paragraph per line, in the kind of the paragraph it is pasted after. Several lines pasted into a paragraph's words become a paragraph each (blank lines dropped; in a heading they join with spaces). A table is one row that covers all its cells: it can be chosen, moved and deleted, and says so when it cannot be copied (not in this version).

**Found by the checks, fixed.** Three real defects, each caught by the browser check and not by the unit tests: a row that got the keyboard by Tab or by the mouse (not by the editor's own arrow keys) was not the row commands acted on, so Delete and Cut acted on the wrong paragraph; after Undo the chosen rows stayed "chosen" though they were now different paragraphs, so the next Cut took the wrong one; and a refused write to the system clipboard became an unhandled error. The checks also had to learn to wait out the 0.4 second autosave before reading the file.

**Where.** `KINDS`, `kindChange` and `withStyle` in `src/guide-edit.mjs`; the commands in `program/ui/editor.js`; the keys in `program/ui/editor-model.js`; the Edit menu in `program/ui/commands.js`; `chooseOne` in `program/ui/dialogs.js`; a table's `span` in `program/ui/preview.js`. Real-browser check `scripts/check-editor-browser.mjs` is now 62 checks. 1620 of 1621 tests pass (one opt-in skip).

**Not in E2.** A bullet toggle (a bulleted list needs a list definition in the document; the activity tool makes them), moving a whole module or lesson with its content as one unit (the tools' outline does that), drag and drop, copying tables, and the time line's minutes. The next step is E3: tables and the fill-in elements, which is also the Forms tool (F3).

## 11. E3, first part: the Insert menu and forms from templates (2026-11-13)

**What it adds.** An **Insert** menu: **Form...** (a handout, worksheet, checklist or other form from one of eight templates: Reading handout, Blank worksheet, Note-taking grid, Checklist, Rating form with evidence, Observer form, Action plan, Commitment card), then the elements a form is made of, each with a dialog for what it needs: Lines to write on, Writing box, Row of fields, Check list, Table to fill in, Rating grid, Cards to cut apart, Source line. A form gets its title, its print line ("Print: 1 per person, portrait.") and its elements; its number is written by Update guide (the browser check presses Update guide and sees "Checklist A: Preparation list"). Whatever is inserted goes after the row the author is on, is chosen afterwards (so it can be moved or deleted), and is one step of Undo. A problem with what was asked for (12 lines at most, 8 columns at most, a title is needed) is explained in the dialog, which stays open. This is the first version of the Forms tool (slice F3 of the forms specification); what is left of F3 is editing a form as a form (its settings, its numbering scheme, the copies list) and more templates.

**Where.** `src/guide-forms-templates.mjs` (templates and `buildForm`, pure), `program/ui/forms.js` (the dialogs, which build with the same builders as everything else), `askFields` and `chooseOne` in `program/ui/dialogs.js`, `insertBlock` in the editor, the Insert menu in `program/ui/commands.js`. Tests: `guide-forms-templates` (eight templates read back as forms with their own ids, no Review issue that stops a guide, Update guide numbers them), `program-commands` (the menu). The editor browser check is 74 checks.

**Next in E3.** Editing the cells of a table in the guide (a grid inside the row, arrow keys between cells, Enter to change a cell's words, add and delete a row or a column). Tables made by Insert, and the checklist and rating grid above, are drawn today as read-only tables.

## 12. E3, second part: the cells of a table (2026-11-13)

**What it adds.** A table in the guide is now written in. Enter on a table row goes into its cells; the arrow keys, Home, End, Ctrl+Home and Ctrl+End move between them; Enter or F2 (or a click) changes a cell's words, in a text box in the cell. While changing a cell's words: Enter goes down a row, Tab and Shift+Tab to the next and last cell that can be written in, and Tab from the last cell adds a row; Escape stops. Ctrl+Enter adds a row below the one the author is in, and Ctrl+Delete deletes it (or Edit, Add table row below and Delete table row); the header row stays, and says why. A cell that holds a check box, a picture or several paragraphs is reached but not opened: it says it is changed with the tools. Typing in one cell is one step of Undo; a row added or deleted is one step. Header cells (column names) are changed the same way and stay header cells.

**What is not touched.** `src/guide-table.mjs` changes the words of a cell through the same code that changes the words of a paragraph, so a cell is held to the same rules, and nothing else in the table is rewritten: widths, the header-row mark, the caption and description, shading, other cells and check boxes are as they were (tests read every table the builders make, change every editable cell and change it back, and compare). A row added is made from a row that is there (check boxes unchecked, each with an id of its own); under a header it is a plain body row.

**Where.** `src/guide-table.mjs` (`readTable`, `setCellWords`, `addRow`, `deleteRow`, pure), `replaceTable` in the in-memory document (replaces a table and its cells, guarded by the first cell's words, grouped Undo for typing), the grid code in `program/ui/editor.js`, `cellKey` and `cellWordsKey` in `program/ui/editor-model.js`, the row and column on each drawn cell in `program/ui/preview.js`. Tests: `guide-table` (9), `program-editor-model`, `program-commands`. The editor browser check is 93 checks; 1637 of 1638 tests pass (one opt-in skip).

**Not yet.** Adding and deleting columns (the column widths of the grid and of every cell must be re-divided), merged cells, a table's caption and description (they are read by screen readers; the Insert dialog sets the caption), copying a table, and the Forms tool's own editing of a form's settings.

## 13. E4: pictures, and two defects found underneath them (2026-11-13)

**What it adds.** **Insert, Picture...**: a file chooser (PNG, JPEG or GIF), a description in a few words (required, or a tick that says the picture is only decoration, as the tools do), a width, and an optional caption that becomes a paragraph under the picture. The picture is drawn in the guide from the guide's own copy (a data address; nothing is fetched), with its description as alt text; it is saved in the project file, survives closing and opening, and moves, copies (as a paragraph) and deletes with the other rows. Pictures added through the tools ("Add a picture to this block") now keep their picture too.

**The two defects.** (1) The in-memory document, which is what the program uses, ignored the pictures a block was made with, so a picture added through the tools was lost as soon as the block was inserted (the paragraph named a picture that did not exist). (2) The package a Word document is made from gave every picture, the cue icons included, the same one-pixel stand-in, so the participant workbook made in the program had blank icons. Now the document keeps the pictures under ids of their own (a block's ids are renamed so two blocks do not share "their first picture"), the project file carries them (`media`: PNG, JPEG or GIF only, as base64 of a bounded size; SVG is refused because it can carry script), what is saved leaves out pictures no paragraph names any more (Undo brings a deleted picture's paragraph and its picture back), and the package holds the real picture of each paragraph and the real cue icons, loaded from the program's own server. The browser check downloads the workbook and finds real pictures in it.

**Where.** `media` in `src/guide-project.mjs` (validated whole, or the project is not opened), `keepPictures`, `mediaOf`, `exportState` and `sourcePackage` in `src/guide-word.mjs`, the pictures sent with an insert in `src/guide-remote.mjs`, `picture` drawing in `program/ui/preview.js`, the dialog in `program/ui/forms.js`. Tests: `guide-pictures-program` (7), `guide-project`, `program-preview`. The editor browser check is 101 checks, the program check 41; 1646 of 1647 tests pass.

**Also in this round.** `Start OpenStage Guide.cmd`: double-click to start the server and open the program in the browser (it asks for Node.js if that is missing and installs the two packages the first time).

**Not yet.** Changing a picture's description, size or caption after it is in (delete and insert again for now), pictures inside table cells, copying a picture between guides, and the Word export of the guide itself.

## 14. Formatting, step 1: how a paragraph sits (2026-11-13)

The author asked for a mini word processor: formatting by hand, and styles that carry over to Word. It is built in five steps; this is the first. Order and reasoning: alignment and spacing need no rich-text box, so they come first; the Styles panel and the Word template (steps 2 and 3) are definitions, not text, so they come next; formatting inside a sentence (step 4) is the one that needs the rich-text box, so it comes late and is tested hardest; the accessibility checks (step 5) need the formatting to exist.

**Step 1 as built.** A **Format** menu (eleven menus now): Align left, Center, Align right and Justify with Word's own keys (Ctrl+L, Ctrl+E, Ctrl+R, Ctrl+J, on a row and while changing the words, where the caret stays put), **Paragraph...** (alignment, space before and after, line spacing, indent from the left and the first line, in points and multiples of a line), and **Clear paragraph formatting** (the style decides again). It acts on the paragraph the author is on or the ones chosen together, as one step of Undo. The paragraph's own settings are drawn over its style in the guide. The dialog says "Justified (the gaps between words vary, which is harder to read)" at the choice.

**What it protects.** `src/guide-format.mjs` writes the paragraph's properties in the order Word's schema gives (a file with them in another order can be refused by Word); clearing every setting gives back exactly the XML there was; a paragraph whose spacing the program sets (a line to write on, a writing box of a set height, a row of fields, the time line, the contents) or a table is refused with the reason, and a paragraph with a line height of a set size keeps it. Values are bounded (space 0 to 144 points, line spacing 0.8 to 3, indents 0 to 216 points).

**Found on the way.** A check that pressed Home in a text box expected the start of the words; in a paragraph that wraps, Home goes to the start of the line. Tests and checks: `guide-format` (8), preview, commands, keys; editor check 112, program check 41, suite 1658 of 1659.

## 15. Formatting, step 2: the Styles panel (2026-11-13)

**What it adds.** **Format, Styles...** opens a panel with every style of the guide and what it comes to (its font, size, color, bold, italic, underline, strikethrough, alignment, spacing, indents, keep with next, shading), a sample that redraws as the settings change, and a contrast reading for the text against its shading. **Apply to this style** changes the style, so every paragraph in it follows, in the guide and in the Word file made from it. Only what was changed is stored, so the rest goes on following the style the style is based on. **Put back as the program makes it** takes the changes away. **New style...** makes a style of your own, starting as a copy of the one chosen and following it. **Delete style** removes one of your own, and refuses while a paragraph is in it (the paragraph would be left with a style that is not there). **Edit, Kind of paragraph** now offers your styles after the five kinds. **Format, New style from this paragraph...** keeps the paragraph's alignment, spacing and indents as a style and puts the paragraph in it; one Undo takes back both. An error in the panel (a size under 6 points, a color that is not six digits) is explained in the panel and nothing is changed.

**How it is kept.** Not in the styles themselves: as edits laid over the program's own definitions each time the styles part is made (`src/guide-styles-edit.mjs`: `overrides` by style id, `custom` styles with their own id, name and base), kept in the project file (`styles`, validated whole, with ids and names checked once), in Undo's history with the paragraphs, and applied to the styles part that the Word file and the Review are given. Taking an edit away gives back the original definitions byte for byte. Everything is written as real Word style properties in the order Word's schema gives, with custom styles carrying the name, the base, "next is itself" and the gallery flag, so they are the styles a Word user sees. A look applied through the Appearance tool now becomes the same edits (before, the program recorded it and did nothing with it), including its borders, as one step of Undo.

**What it does not do.** A style of your own is not carried into the participant workbook (the workbook has its own look and its own list of styles); its paragraphs are left out of it like other text the author typed. Character styles (for words inside a sentence) come with step 4. Borders can be set by a look but not yet in the panel.

**Found on the way.** The program's own OSG styles are Word "custom styles" too, so "custom" could not mean "the author's own"; the author's styles are told apart by being in the edits. A style whose only change was a border was never written. A spec's indents needed wider limits for styles (the program's own styles hang by 120 points). Tests and checks: `guide-styles-edit` (16), `guide-styles-document` (7), project, commands; editor check 133, program 41, project 29; suite 1682 of 1683.

## 16. Formatting, step 3: Word files (2026-11-13)

**What it adds.** Three commands. **File, Save as Word file...** makes a Word document from the guide (its paragraphs, tables, pictures and cue icons, its styles as the author changed them, a real contents field, the guide's title as the file's title) and gives it to the author as a download. **File, Save styles as a Word template...** makes a Word template (.dotx) that holds the guide's styles and one empty paragraph, so a document started from it looks like the guide; the styles are real Word styles and appear in Word's Styles gallery. **Format, Styles from a Word file...** reads the styles of a Word document or template and shows each paragraph style with a sample, a line saying how it looks in words, and where it goes: a style of the file that has the name of one of the guide's styles is chosen to change that style; any other can be added as a style of the author's own (not chosen until the author chooses it). **Bring in** does it as one step of Undo and says what it did.

**Reading a Word file.** The file is only read: nothing in it is opened in Word, run or changed. The reader takes `word/styles.xml` and the two fonts of the theme (so a style that asks for "the body font" gets Calibri or whatever the theme names), and works out what each style comes to once the styles it is based on and the document's defaults (the space after and the line spacing Word gives every paragraph) are counted. What the guide cannot hold is brought to the nearest value it can (a 200-point size becomes 96, 9000 points of space becomes 144) or left out (a font name with odd characters), and the dialog says so. A style that says nothing about bold takes bold away from the guide's style it replaces, as the panel does. Limits: 25 MB for the file, 6 MB for the styles part, 3000 parts, 150 styles offered.

**What the Word file is.** A zip with the content types, the package relationships, the document, its relationships, the styles, a small settings part, the lists and pictures the guide uses, and the title in `docProps/core.xml`. No entries for folders (some services that open Word files refuse a file that has them). The contents lines the program shows (plain lines in styles TOC 1 and TOC 2) are written as one real contents field with those lines as its cached result; Word fills in the page numbers when the author presses Ctrl+A, F9 (the dialog says so). The two contents styles are added to the styles part when the body uses them and the part does not have them.

**Checked against Word itself.** The document and the template pass the Word schema validator, and a real Word opened both (22 paragraphs and one contents field; the contents updated to page numbers; the title property read back; the author's `Callout` style present with its font and shading in both). That is one machine's Word, opened read-only from a script; it is not a substitute for the author opening the files.

**Where.** `src/guide-docx.mjs` (zip out, styles in), `src/guide-styles-import.mjs` (matching, fitting, the plan), `program/ui/word-files.js` (the download and the two dialogs), the commands in `program/ui/commands.js`, and `public/vendor/jszip.min.js`, loaded by the program page (the add-in's own copy of the zip library). Tests: `guide-docx` (13), `guide-styles-import` (12), commands; editor check 153, program 41, project 29.

**Found on the way.** The buttons at the foot of every dialog were cut off by 18 pixels (a negative margin the dialog's own scrolling box clipped); they now sit flush with the dialog's edge. A style's name in Word is lower case for the built-in ones ("heading 1"); the dialogs now say "Heading 1".

**Not yet, and not tried.** The browser's own download is how the file reaches the author; it works in the browser check, and has not been tried in the program's own window. Character styles in a file (they come with step 4), table styles, and the styles' numbering are not read. A style of a file is never brought in by id, only by name or as a new style.

## 17. Formatting, step 4: looks for words inside a sentence (2026-11-13)

**What it adds.** The words of a paragraph can now have a look of their own: bold, italic, underline, strikethrough, raised (superscript) or lowered (subscript), a color, a font and a size, and a character style. **Format** has Bold (Ctrl+B), Italic (Ctrl+I), Underline (Ctrl+U), Strikethrough, Superscript (Ctrl+Shift+=), Subscript (Ctrl+=), **Text...** (font, size, color and the looks in one dialog), **Character style...** and **Clear text formatting** (Ctrl+Space); a formatting bar above the guide has the same looks and the alignments as buttons, with the looks the chosen words have shown as pressed. While the words are being changed, a look goes on the words chosen (with none chosen, the next words typed, as in a word processor); from a menu, on the words chosen a moment ago, otherwise on every word of the paragraph or the paragraphs chosen. A toggle is decided once, from the first paragraph, and done to all. Each is one step of Undo, apart from the typing before and after it. The Styles panel can make **character styles** (a look with a name, kept in the project and in the Word file as a real Word character style); changing one changes every word that has it, and one in use cannot be deleted.

**How it is kept.** The words of a paragraph are held as spans: stretches of words, each with the Word run properties it had (`src/guide-runs.mjs`). The properties are kept as Word wrote them, so what the program does not model (the language, a theme font, a ligature setting) goes through every edit untouched; only the settings changed are written, in the order Word's schema gives, and a look taken off is removed so the style decides again (a look switched off against a style that has it on is written as an explicit off). Editing the words, Enter in the middle of words and Backspace at the start now keep each stretch's look (before, editing a paragraph with a bold word made the whole paragraph one plain run: formatting was lost the first time someone typed). Typing after a bold word is bold; typing at the start takes the first look; a new paragraph made by Enter starts with the look of the words it was cut from, except an empty one made at the end of the words, which starts plain.

**The box the words are typed in** (`program/ui/richbox.js`) replaces the plain text box for a paragraph. It shows each stretch with its own look as it is typed. The browser types, deletes and moves the caret in it; after each change the list of spans is read back from what the browser left, and the elements are drawn again only when the browser left them in a shape the box would not have made. It refuses everything that could bring markup in from outside: the browser's own bold and italic keys, paste (only the plain words come in; several lines are made paragraphs by the editor, as before), dropping text, the browser's own undo, a new line. It answers the same way the text box did (value, selectionStart, selectionEnd, setSelectionRange, setRangeText), so the editor's keys and saving did not change. Table cells still use the plain text box: a cell's words have no looks of their own yet.

**What the Review says.** See section 18.

**Found on the way.** (1) The tools' **Review did not work inside the program at all**: the documents it reads cannot be sent between the page and the tools' frame (the browser refused with "XMLDocument object could not be cloned"). They now cross as text and are read again; a test now sends every message through a real structured clone, which is what two pages do. (2) A command from a menu took the words that were chosen a moment ago for the row the keyboard was on, and the keyboard could be anywhere once the menu closed, so a look could land on the first paragraph of the guide; the words chosen are now kept for their own row, and forgotten as soon as the author moves on. (3) The Styles panel gave the keyboard back to a row even when the author had already gone on into the guide.

**Checked.** `guide-runs` (16), `guide-check-looks` (8), preview, commands, editor model, remote; editor check 201 (four runs in a row), program 41, project 29; 1737 of 1738 tests (one opt-in skip). A sample Word file with every look passes the Word schema validator and a real Word read each look (bold, italic, the red Georgia 14, strikethrough, underline, and the character style with its own bold and color).

**Not yet, and not tried.** The browser's own spelling and grammar underlines are the browser's, not Word's. A look on words inside a table cell, in a heading, or in a paragraph that also holds a picture or a check box (such a paragraph is still changed with the tools). Highlight color. Typing a look at the end of the words and pressing Enter to go on with it. The box has been tried with the keyboard and the mouse in a browser, not with a screen reader or an input method (the NVDA pass waits for the author), and not in the program's own window.

## 18. Formatting, step 5: what the Review says about how words look (2026-11-13)

**What it adds.** Six checks in Review guide (the same code the Word add-in's pane and the command-line check use), each saying what is wrong and what to do:

| Check | What it finds | Level |
|---|---|---|
| CONTRAST | Words the author colored that are hard to read: measured against the paragraph's shading (or the page), 4.5 to 1, or 3 to 1 for words of 18 points or more (14 and bold) | To fix |
| TEXT-SIZE | Words under 10 points, whether by their own size or their style's | To look at |
| JUSTIFIED | A paragraph justified by hand or by its style: the gaps between words vary, which is hard to read, especially with dyslexia | To look at |
| FAUX-HEADING | A short paragraph, all bold or all large, with no full stop, in an ordinary style (Normal, Block text, a note, or a style of the author's) and no heading style: screen reader users cannot find it by heading | To look at |
| COLOR-ONLY | Color is the only thing that sets some words apart (a color that is not a grey, with no bold, italic, underline, strikethrough, size or character style beside it, in a paragraph that also has other words) | To look at |
| UNDERLINE | Underlined words that are not a link: on a screen they look like links | To look at |

TEXT-SIZE and JUSTIFIED say so when they pass; CONTRAST says how many colored stretches meet the contrast it needs. A guide the program made itself raises none of them (a test builds a standard guide and checks that nothing is flagged). Each finding has a "Go there" link to the first place it was found.

**Not claimed.** As before, this is not Word's Accessibility Checker and not a screen-reader test; the Review's own line says so. The checks look at how words are marked up, not at whether the words are plain. Colors in pictures and in tables' own styles are not examined beyond what was there.

## 19. The font pack (2026-11-13)

**What it adds.** Four families of free fonts made to be easy to read come with the program (`program/fonts`, listed in `src/guide-fonts.mjs`): **Atkinson Hyperlegible** (the Braille Institute's, regular, bold, italic, bold italic), **Lexend** (made to make reading easier, regular and bold), **Source Sans 3** and **Source Serif 4** (Adobe's, four styles each). About 2.8 MB. The font boxes (Format, Text... and the Styles panel) offer them first, each with a line about what it is for and "comes with the program", then the fonts Word documents usually have; any name can still be typed. The page draws words in them whatever is installed on the computer (the stylesheet's `@font-face` rules, served by the program's own server).

**In a Word file.** A Word file only names a font, so on a computer without it Word substitutes another and the layout changes. When the guide (its styles or its words) uses one of the program's fonts, **Save as Word file** and **Save styles as a Word template** now carry the font inside the file, in Word's own way: scrambled font files in `word/fonts`, a font table that names them, the relationships, the content type for them, and the "embed fonts" setting. Only the families the file uses are carried, all their styles (so bold and italic are right), and the status line says which. A file that uses none of them is exactly as before. Other fonts are still only named. A file with all four families is about 1.3 MB larger.

**Where it came from, and what is unchanged.** Each file is the makers' own, downloaded from the fonts' official repositories (Google's Atkinson Hyperlegible and Lexend, Adobe's Source Sans and Source Serif), unmodified. Each is a static TrueType font whose own settings allow embedding (checked by a test). Each is under the SIL Open Font License 1.1, which allows use, shipping and embedding in documents; the license text of each is in `program/fonts` beside it, and NOTICE credits each.

**Checked against Word itself.** None of these fonts is installed on this computer, so a real Word could only draw them from inside the file. Word laid out the same words in a file with and without the fonts: with them, to within 0.3 points of the widths the font files themselves give (Source Serif 322.5 against 322.74, Atkinson Hyperlegible 176.25 against 176.12, Lexend 224.25 against 224.53 once its 120-point indent is taken off); without them, 429, 171.75 and 197.25, which is Word's substitute. That shows the scrambling is right and Word loads the fonts. Source Sans 3 could not be told apart this way, because Word substitutes the older Source Sans Pro, which this computer has and which has the same widths. The files also pass the Word schema validator. Word's PDF export hangs in a hidden Word on this computer even for a file with no embedded fonts, so no PDF was made.

**Left out, on purpose.** OpenDyslexic: its only published files are PostScript-outlined OpenType, which Word cannot carry inside a file, so it could be drawn in the program but never travel; the evidence that it helps is also thin. It can be added as a font that is only named if the author wants it. Fonts in the participant workbook: the workbook is made by the tools and keeps its own look, so it does not carry the pack. Fonts in a heading or a table cell are the looks-inside-a-sentence work that is still to do.

**Checked.** `guide-fonts` (13: the files, the licenses, NOTICE, the stylesheet, the scrambling with a known key, what a file carries and what it does not, the size limit, the server's rules), program-files (the one fetch allowed); editor check 209 (twice), program 41, project 29; 1750 of 1751 tests (one opt-in skip).
