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
