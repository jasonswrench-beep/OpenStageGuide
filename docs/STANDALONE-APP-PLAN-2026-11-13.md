# OpenStage Guide as a program of its own: plan

Draft, 2026-11-13, updated the same day with the spike result (`STANDALONE-APP-SPIKE-RESULTS-2026-11-13.md`) and the review of 16 real training documents (`EXAMPLES-REVIEW-2026-11-13.md`). Only the spike is built: the pane runs on a project file, with no Word.

**Where it lives (decided 2026-11-13, after this plan was written):** the program has a repository of its own, `OpenStageGuide`, made by the author. The plan below was written in the Word add-in's repository, `OpenStage-Guide`, where "this repository" and the branch `app/spike` mean that one; the add-in stays there as it is, and this repository started from the program's files only. That answers the first half of question 4 in section 7 (where the code lives); where a public download would live is still open.

## 1. Why, and what the author has decided

The Word add-in cannot be given to just anyone. Microsoft's free ways to install one are a per-PC testing route, or an administrator inside one organization. The public route (the Marketplace) is closed to the author today. Every user also depends on the author's web server for the add-in's files.

The author wants it workable and deployable by anyone, as a complete facilitator guide and workbook production tool, like the other programs of the suite. Decisions given on 2026-11-13:

1. **Authors write the guide in the program itself.** Word is no longer the editor.
2. **Export to Word is an option,** left to each person. Nothing needs Word, and nothing is Word-only.
3. **It must look like the rest of the suite,** following the most recent versions of those programs.
4. **The scope is the whole kit** (author, 2026-11-13): the script, **and** the forms and handouts, the planning kit, the reference pages and the layouts around it, as the examples show. The first piece is specified in `FORMS-AND-HANDOUTS-SPEC-2026-11-13.md`.

The add-in stays as it is (published, maintained for faults) while the program is built. Whether it is ever retired is a later decision.

The author also asked for a range of real facilitator guides and workbooks to be looked at before going further. They show that the words a facilitator says are mostly covered already; what is missing is the **pages around them**: forms and handouts, planning kits, reference pages, segment headers, time shown in more ways, and page layouts. That widens the scope of the program, and it is the reason section 3 now has more rows than it had.

## 2. What carries over (measured 2026-11-13 by a direct search; a closer look comes first)

- 96 source files in `src/`, about 14,100 lines. **Two of them touch Word**: `guide-word.mjs` (754 lines) and `word.mjs` (133 lines); `app.mjs` and `picker.mjs` mention it a couple of times each. The rest is plain JavaScript: the pane's screens (`guide-pane.mjs`, 2,020 lines), the guide model, the cue and review rules, the workbook, assessment, slide and delivery builders, the writers of Word-format files.
- 126 test files.
- **The pane already talks to its document through one interface**: about 30 methods (`paragraphs`, `cursorIndex`, `insert`, `go`, `deleteRange`, `moveRange`, `applyEdits`, `reviewParts`, `sourcePackage`, `exportPdf`, the saved settings such as own cues, look, slide memory and version memory, and so on). There are two implementations: the Word one, and an **in-memory demo document** (about 190 lines) used by the browser preview (`node scripts/server.mjs --preview`, port 3844) and by tests.
- The document the pane understands is **a list of styled paragraphs** (Module, Lesson, Time, Cue, Notes, Assessment, Question, Option, tables and so on). The guide is read from the styles and the text, with no hidden record. That is why a program of our own can hold the same thing.

So the program's document is the demo document made real: the same paragraphs, kept in a project file, with a real editing surface over them.

## 3. What is new

| Piece | What it is | Size |
| --- | --- | --- |
| **The editing surface** | Typing and formatting inside the program: styled paragraphs, bold and italic, lists, **tables with fill-in elements** (check boxes, blank lines, signature and date lines), pictures with captions, footnotes, Undo and Redo, cut, copy and paste. Not a full word processor. Keyboard and screen reader use from the start. | The largest piece |
| **The project file** | One file per guide (the paragraphs, the saved settings, the pictures). Autosave, and a backup copy. The file name extension must not clash with Capture Express's `.osguide.zip`. | Small |
| **Outputs** | PDF and a web page made from the program's own view of the guide; Word as the optional third. The workbook and facilitator copies already have builders. | Medium; **tagged (accessible) PDF is unproven, see risks** |
| **Forms and handouts** | A form is a thing of its own: numbered (by unit, 5.4, or by letter), with fill-in elements, scoring grids, evaluation forms and cut-apart cards, a print specification (copies per person or per team, paper size, orientation), cited from the outline (a "Forms used" column) and from the script with automatic cross-references. The participant workbook holds the same forms. 12 of the 16 examples have them. | Medium to large; the biggest gap the examples show |
| **Segment, module and unit headers** | Purpose, materials, method, minute range, optional flag, "choose n of m"; a module fact page; the unit's lists of handouts and slides built from the guide, not typed. | Small to medium |
| **A wider time model** | Ranges, a clock start plus a duration, set-up before the first segment and clean-up after the last, a start time for each day, group-size rules, an interaction-mode tag (whole group, small group, solo), a stated programme total, breaks. The Review checks typed times against the stated total. The examples show six different ways of giving time. | Medium; extends Timing and Agenda and Delivery Versions |
| **Script mechanics** | An expected-answer line after Ask, Debrief and Option cues, an optional flag, "slide or flip chart" as one instruction, a library of named techniques in the activity library. | Small |
| **Reference pages** | Glossary, references, figure captions with source credits, pull quote with attribution, licence or permission page. | Medium |
| **The planning kit** | Planning timeline and checklists (owner, lead time, notes), a materials list grouped for facilitators, each person and each team with quantities, a room set-up picture and list, a roles grid, pre-work and follow-up as items. | Medium |
| **Layouts for the outputs** | Properties of an output look, not of the content: slide thumbnail beside or above the notes, section dividers, running headers, a margin column of icons, mixed orientation, mirrored footers, one-page summary sheets. | Medium |
| **Import** | A guide made with the add-in (its structure is in the styles, so it can probably be read), pasted text, and the slide import that exists. | Small to medium |
| **The suite shell** | Navy bar, menus, welcome page, search, key list, dialogs: section 4. | Medium |
| **The program and its installer** | Same pattern as OpenAssess: a plain-JavaScript interface, a local server inside the program, a Tauri window, a signed NSIS setup. | Medium |

## 4. The look: the suite's, as of the latest versions

Source of truth: `design/DESIGN-GUIDE.md` (version 1) and `design/openstage-ui.css` in the Author repository. Every product keeps a copy of the style file. The closest model is **OpenAssess** (plain JavaScript, no build step, local server, Tauri window): its ports of the menu bar, pop-up menus, command search, key list and frame, and its icon generator, are the ones to copy.

- **Navy bar:** the product's logo and name, the name of the open guide, saved or not, search, and at most five buttons (proposed: Save, Undo, Redo, Review guide, Make workbook).
- **Menu row**, from the 21 tools in 7 groups the add-in already has: **File** (New, Open, Close, Save, Save a backup, Restore a backup, Export: PDF, web page, Word), **Write** (Add module, Add activity, Instructional cues, Activity library), **Bring in** (Import slides, Update from slides), **E-learning** (Screens, Assessments, Learning alignment), **Check** (Review guide, Timing and agenda, Delivery versions), **Make** (Appearance, Generate materials, Participant workbook), **Help**. Everything is in a menu, with its icon and key. (Regrouped on 2026-11-14 into fourteen menus: see PROGRAM-SHELL-2026-11-13.md.)
- **Work area:** the guide itself is the biggest thing on screen, with the outline on one side and properties in tabs and sections on the other, as Author does.
- **Welcome page:** New guide and Open guide as large buttons, recent guides (with a way to clear the list), and the sample guide as a card that opens a copy.
- Ctrl+K command search, Ctrl+/ key list, F10 for the menus, dialogs with a Close x and a sticky foot, 24 by 24 pixel minimum buttons, Lucide icons, Segoe UI. Windows high contrast must keep working.
- **Its own colour.** The guide says every product has a pair. For the Guide, **accepted by the author on 2026-11-13**: a teal-blue, action `#0E7490` (white words on it 5.4:1), hover `#155E75` (7.3:1), tint `#E3F4F7` (action words on it 4.7:1), name on the navy bar `#67E8F9` (11.9:1). Violet, rose, gold, indigo and orange are taken; green is the success colour.
- **Needs the author's yes before anything is changed outside this repository:** the design guide currently says "OpenStage Guide ... stay as they are", and adding a Guide column to it, to `icons.json` and to the revamp checklist means editing the Author repository's `design/` folder, which other products copy from.

## 5. The order of work

Each step ends with checks run by script, screenshots for the author, and nothing pushed without a yes.

1. **Spike: the core without Word.** A project adapter (the demo document plus a file on disk). Run the existing pane screens in the browser against it. Test: write a guide, save, close, reopen, and every adapter test passes against it as it does against the demo document. This answers how reusable the code really is.
   **Done 2026-11-13** as `d718730` on the branch `app/spike`; results in `STANDALONE-APP-SPIKE-RESULTS-2026-11-13.md`.
2. **The editing surface**, first in a plain form (styled paragraphs, typing, Undo) **with tables and fill-in elements from the start**, then widened. A screen reader pass on it before building further on it.
3. **Model additions from the examples**, in this order, each a small piece of the shared core with its tests and one tool in the program: forms and handouts; segment headers and the wider time model; script cues; reference pages; the planning kit. Each also gets Review checks the examples suggest: agenda names that differ from activity names, typed times that do not add up to the stated day, a form that is cited but missing, a wrong language flag.
4. **The suite shell** in the OpenAssess pattern: bar, menus, welcome, search, key list, dialogs, the colour pair. It can be built alongside steps 2 and 3.
5. **Outputs:** web page and PDF with the layouts above, then the optional Word export and import of add-in guides.
6. **The program and installer:** Tauri window, signed setup, smoke test of the installed program by script.
7. **The manual:** the Explainer session rewrites the Quick Start and User Guide for the program; the pictures are retaken.

Where it lives (recommended): a new `app/` folder in this repository, on its own branch, importing the same `src/` core, so the add-in and the program cannot drift. It can move to a repository of its own if the author prefers; OpenAssess is its own.

## 6. Risks and what is not known

- **The editor is the big one.** It has to be good enough that nobody misses Word for writing a guide, and accessible. Tables and pictures are where this gets hard, and the examples make them day-one, with fill-in elements. The demo document already flattens a table into its cells' paragraphs.
- **The scope grew.** Forms, the planning kit and layouts are a product around the script. Build them in the order above and release when the script, the forms and the outputs work, not when every example can be reproduced.
- **The examples are other people's work.** They were used to see what exists, nothing was copied, and none of them ships as a template. Which of the five older "program" layouts (overview, objectives, outline table, directions, reproducible forms) to offer as a starting template is the author's choice.
- **Tagged PDF.** The add-in gets an accessible (tagged) PDF from Word. A program of our own has to produce one from its own view. Whether the Windows web view's print-to-PDF tags the file has not been checked. PDFRemediator is the fallback.
- **Edits made outside the program** (in Word after an export) do not come back. The program is the source of truth, and the Word export says so.
- **Two products to keep in step** while the add-in lives. The shared core keeps that small.
- **Not covered by the 2026-11-13 measurement:** how the 2,000-line pane screens reach Word through the interface in detail; whether the guide's saved settings all fit in a project file; how much of the 126 test files assume Word-shaped documents.
- **Signing and releases.** The author has Azure Artifact Signing working for other programs. A public pre-release needs a public place: this repository is private.

## 7. Questions only the author can answer

1. ~~The program's name~~ **Answered 2026-11-13: it stays OpenStage Guide.** The file extension is still open (not `.osguide`, which Capture Express uses).
2. ~~The Guide's colour pair (section 4 proposes teal-blue).~~ **Answered 2026-11-13: the blue is accepted** (`#0E7490`, `#155E75`, `#E3F4F7`, `#67E8F9`).
3. Whether the add-in is kept going next to the program, or frozen once the program can do everything it does.
4. Where the code lives (this repository, recommended, or a new one) and where a public download would live.
5. What the editor must handle on day one. The examples say: **tables and fill-in elements, pictures with captions**; footnotes soon; comments and tracked changes appear nowhere. Confirm.
6. Whether the suite's design files may get a Guide column.
7. ~~Is the program's scope the script plus the forms and planning kit around it, or the script alone?~~ **Answered 2026-11-13: the whole kit.**
8. Should the participant workbook hold the same forms the facilitator pages cite, so a number and a page reference stay true?
9. Should a handout carry a print specification (copies, paper, orientation)?
10. Are slide-shaped and self-study outputs wanted, or only guides and workbooks?
11. Should any of the five older "program" layouts be offered as a template to start from?
