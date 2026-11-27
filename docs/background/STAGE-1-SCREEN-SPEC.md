# Stage 1 screen spec: Start a guide, Add, Review

Draft 2026-10-30, for the author's sign-off before any add-in code is written. It follows the rebuild plan of 2026-10-30 (kept outside the repository). Layout details assume the cue-label layout wins the Stage 0 comparison; if it does not, only the inserted paragraphs change, not the screens.

## The rules every screen follows

- One clear next action per screen, and its button is visible without scrolling.
- Everything the pane does is real content inserted into the Word document. Word's own Undo (Ctrl+Z) takes it back. There is no hidden record to fall out of step, so no "operations paused" state exists.
- The pane finds the guide by reading the document: its styles and headings. If the author edits by hand, the pane reads the new state.
- Messages say what happened and what to do next, in plain words. No jargon, no codes.
- Every control works from the keyboard, in a sensible Tab order, with visible focus and a text label.
- The old pane stays available behind a "Legacy tools" link at the bottom. Stage 1 does not touch it.

## Screen 1: Start a guide

Shown when the document has no guide yet.

```
OpenStage Guide

Start a guide
Guide title   [______________________]
[x] Add the standard pages
    Cover, contents, Getting Started, Training at a Glance

[ Start guide ]

Try the sample guide
Legacy tools
```

- **Start guide** inserts the title page and standard pages at the very top of the document. Existing text is left exactly as it was, below the new pages, and the pane says so.
- **Try the sample guide** inserts a short finished example so a first-time author can see what the result looks like. It is ordinary content and can be deleted or undone.
- No form beyond the title. Look and brand choices arrive in Stage 3, with sensible defaults until then.

## Screen 2: Build

Shown once a guide exists. One screen, top to bottom.

```
How to Reframe Conflict            Total: 20 min

Contents (click to go there)
  Module 1: Recognize the conflict
    Lesson 1.1: Reframe the conversation   20 min

[ Add module ]  [ Add lesson ]

Add a block
Find a cue  [____________]
  Objective   Say   Ask   Activity   Key point
  Time check  Transition  ...        (list, arrow keys, Enter)

Text   [__________________________]
Notes  [__________________________]  (facilitator only)
Minutes [__]   Shown to: (o) Everyone ( ) Facilitator only

[ Insert block ]
```

- **Contents** is built by reading the document's headings. Clicking a line moves the cursor there.
- **Add module** and **Add lesson** insert a heading (module and lesson styles, both real headings so Word's Navigation pane shows them), a lesson time block with cumulative time, and place the cursor ready for the first block. A new module ends with a Transition block.
- **Add a block** inserts at the cursor. If the cursor is inside a paragraph it inserts after that paragraph. A cue is chosen from a list by mouse or arrow keys; typing in the box narrows the list. Enter in the list moves to Text; Ctrl+Enter anywhere in the form inserts.
- **Notes** become a shaded facilitator-notes paragraph that begins with the words "Facilitator notes:". **Shown to** decides whether a block is kept when the learner workbook is made in Stage 3. It is recorded as a paragraph style, so it survives copy and paste.
- **Minutes** feed the lesson time block. **Update times** (a small link beside Total) rewrites the cumulative times in each lesson block; it can be undone.
- Cues come from a catalog (name, icon, default audience). Adding a cue later means adding one catalog row and one icon, not changing code. The first release ships the author's 14 icons, with text cues such as Say and Ask mapped to them. The sample guide's longer list, including the virtual-class cues, joins as icons are made.

## Screen 3: Review

Opened from a "Review guide" link on the Build screen.

```
Review this guide
2 to fix   3 to look at   4 to check yourself

To fix
  Lesson 1.2 has no time block.              [Go there]
  A picture has no alt text.                 [Go there]
To look at
  Heading levels skip from 1 to 3.           [Go there]
  ...
To check yourself
  Does each cue read sensibly out loud?
  Run Word's own check: Review > Check Accessibility.
```

- The first two groups come from the same rules as `scripts/docx-structure-check.mjs`: title, language, heading levels, table header rows, alt text or decorative marks, colour contrast, plus guide rules such as a lesson without a time block or a time that does not add up.
- The pane says plainly that this is not Word's checker and not a screen-reader test, and points to Word's checker.

## What Stage 1 does not do

Assessment items, the learner workbook, looks and brand colours, slide import, delivery versions and backup. They stay in the legacy pane or come in later stages.

## As built (2026-10-30), where it differs from this draft

- **Build screen order:** add a module or lesson, add a block, tidy (Update guide, Review guide, Refresh), then Contents. Contents moved to the bottom so the most used controls come first.
- **Transition:** Add module inserts the heading and a Transition block. A Transition counts as the module's block, not any one lesson's, so a lesson with only the Transition is still reported as having no blocks. Add lesson puts new lessons before it as long as the cursor is on the module heading or inside the module.
- **Cumulative time** is the minutes taught before a lesson starts, not the running total at its end. The sample lessons from Stage 0 showed the other reading; tell me if you prefer it.
- **Minutes:** a lesson's time is the sum of its block minutes when any block has minutes, otherwise the number in its time block. Update guide writes that back. It also renumbers modules and lessons that are written as "Module N: Title" and "Lesson M.N: Title".
- **Start guide** also sets the document's title property, so Word's check does not report a missing title.
- **Built since:** keyboard shortcuts and the pane opening with its document (2026-11-01: see "Keyboard shortcuts and opening with the document" below), the Graphic Cues page and the generated Training at a Glance table (2026-10-31: see "Getting Started pages" below).
- **Ribbon:** the Open Designer, Add Module, Add Activity, Course Outline and Review Guide buttons all open the Guide pane (Review Guide opens its Review screen). The other ribbon buttons still open the old screens, which sit behind "Legacy tools".

- **After the first native test (2026-10-31)** the Build screen gained "In your document" (what the cursor is in, with Edit this block, lesson or module), an "Edit a block" mode on the block form, placement by block, lesson and module instead of "after the cursor's paragraph", and a pane that follows the cursor and Undo. See `STAGE-1-NATIVE-RESULTS-2026-10-31.md`. A module's Transition is still inserted with the module, and always stays at the end of it: lessons and blocks are placed before it.

## Getting Started pages: Training at a Glance and Graphic Cues (built 2026-10-31, the author said "yes, as buttons")

Two buttons under "Getting Started pages" on the Build screen. Each builds one page from the guide as it is now. Nothing rebuilds a page by itself, not even Update guide, so the pane only says whether each page "has not been built yet", "matches your guide" or "is out of date" (read from the document, no hidden record).

- **Training at a Glance** (`src/guide-pages.mjs`): an intro line, a table, and a closing line "Total teaching time: N minutes." The table has three columns, "Module or lesson", "Minutes" and "Starts at minute", a real header row, a caption and a description, no merged or blank cells. A module row (lightly shaded, its name in the cue-label style) gives its lessons' minutes added up and where it starts; a lesson row gives its minutes and its start. A lesson's minutes follow the rule Update guide uses (its blocks' minutes if any, else its time block), and "starts at" is the cumulative time before the lesson. It needs at least one lesson.
- **Graphic Cues**: an intro line and a two-column table, "Cue" and "What it means". Each cue row has the picture (decorative) and the cue's name in one cell, and the cue's help text in the other, so no cell is empty and a screen reader hears the name. Only the cues the guide really uses are listed, in the catalogue's order, with Time first when any lesson has a time block. It needs at least one block.
- **Where it goes:** under its own Heading 2. Pressing the button again replaces the intro, table and closing line it wrote earlier and nothing else; anything the author typed on the page, above or below, stays. On a Glance page still holding the untouched placeholder text, the placeholder is replaced. If the heading is missing, the heading is added: Training at a Glance before the first module, Graphic Cues before Program Preparation (so after About This Guide), else before Training at a Glance, else before the first module. The cursor lands at the start of the new page.
- **Checked by the tests, not in Word:** both pages pass the review and the structure check with header rows; the learner workbook leaves both out and still passes its leak check; Undo takes a page out in one step.

## Module overview (built 2026-10-31, the author said "yes, as a button")

A third button, "Build module overview", in the same section (now headed "Pages built from your guide"). It builds the page that opens a module, for the module the cursor is in. The note under the buttons says whether that module's overview is built, current or out of date, and how many modules have one. Like the other two pages, nothing rebuilds it by itself.

- **Under the module heading, before its first lesson:** the author's lines *Goal*, *Overview*, *Materials needed* and *References and resources*, each a bold label and one sentence to type over; and between them the part the guide can write: "Time: N minutes in M lessons." and a table with columns Lesson, Minutes and Objective (a header row, a caption, no merged or blank cells; a lesson with no Objective block reads "No objective yet."). Minutes follow the rule Update guide uses. It needs at least one lesson.
- **Rebuilding rewrites only the time line and the table,** found by their shape anywhere between the module heading and its first lesson. The four lines are written once and never touched again; if only they are left, the button adds just the time line and table after them.
- **A new style, OSG Overview,** marks all of it (a left bar like the facilitator-only blocks, and the label says what each line is). A new first lesson in a module with an overview goes after it, not above.
- **All of it is for the facilitator.** The learner workbook leaves it out and says "1 module overview" among what is left out; the author's own text is still the only thing listed with a way to it. The leak check searches for its words as it does for notes and keys.
- **Checked by the tests, not in Word:** the page's exact paragraphs, a rebuild that keeps the author's words, the note, the review, the structure check, Undo in one step, and the workbook with its leak check.

## Activities as labelled parts (built 2026-10-31, the author said "yes, build them")

On the Add a block form, when the Activity cue is chosen a box appears: "Write it as labelled parts". Ticked, the Text box becomes "Title of the activity" and five fields appear: Objective, Materials, Preparation, Instructions (one step to a line, required) and Case study. Another cue puts the form back to plain text. It is still the one Activity block, so it is placed, edited, undone and counted like every other block.

- **What is written:** the usual cue line (picture, "Activity", the title and "(N min)"), then each part in its own labelled paragraph in the order Objective, Materials, Preparation, Instructions, Case study. A part's first line starts with its label in bold; further lines are further paragraphs of the same style. Empty parts are left out. The block's own Facilitator notes still come last, with their "Facilitator notes:" label.
- **Who sees what:** Materials and Preparation are facilitator-only, so they are written in the shaded facilitator-notes style and never reach the learner workbook. The title, Objective, Instructions and Case study are for whoever the block is shown to (Everyone by default; Facilitator only drops all of it). The minutes stay on the title line, where the guide already reads them, so the lesson's time adds up as before.
- **Reading it back:** Edit this block opens a block written this way with its parts filled in and the box ticked. A block is read as parts only when every paragraph after its first line is a labelled part (in the right style), a further line of one, or a facilitator note, and it has instructions; anything else opens as plain text, so no words are lost.
- **Checked by the tests, not in Word:** the exact paragraphs and styles, the form, the errors, an edit that replaces the whole activity once, minutes feeding the lesson, the review, and the workbook (objective, instructions and case study kept; materials, preparation and notes left out; facilitator-only left out whole).

## Keyboard shortcuts and opening with the document (built 2026-11-01, the author allowed the manifest change)

The manifest (now version 1.1.0.1, so Word re-reads it) differs from the previously accepted one in two ways only: the version, and the shortcuts file below. The previous manifest is in git (`git show f93b8a1:manifest.xml`).

- **Reopening the pane with the document is NOT available (changed 2026-11-01).** Version 1.1.0.0 made the Open button a ShowTaskpane action marked `Office.AutoShowTaskpaneWithDocument`, which is what Word needs before a document may ask for its pane to reopen with it. In the author's desktop Word, clicking that button did nothing at all; the same manifest with the button running the `openDesigner` function again opened the pane (recorded on video, from the public site). So 1.1.0.1 has the button run `openDesigner` as before (its shortcut works the same way), no button has a ShowTaskpane action, and `AUTO_OPEN_SUPPORTED` in `src/guide-word.mjs` is false: the pane does not ask to reopen, and takes the request out of a document that carries it. What is not known: whether the fault is the special id or any ShowTaskpane action with this shared runtime. A later experiment could try (a) a ShowTaskpane action with an ordinary id, and (b) the special id on a second, separate control; each needs a Word restart to try.
- **`<ExtendedOverrides Url=".../shortcuts.json">`** sits directly below `VersionOverrides` and names `public/shortcuts.json`, now served by the local server and carried in the static export at the root. The shared runtime that shortcuts need was already configured (SharedRuntime 1.1, a long-lived runtime on the pane page).
- **Eight shortcuts**, all in the `Ctrl+Shift+Alt+letter` pattern Microsoft recommends, from the document with the pane open or not: **G** open the Guide pane, **R** review the guide, **M** next module, **N** previous module, **L** next lesson, **K** previous lesson, **C** go to the Contents page, **T** go to Training at a Glance. Next and previous go to the nearest module or lesson heading from the cursor; past the end, or with no Contents or Glance heading, the cursor stays and the pane's status line says why. Word needs Microsoft 365 version 2408 or later on Windows. The first time a key is used that Word already uses, Word asks which to keep; Office remembers the answer per user.
- **Checked by the tests, not in Word:** the manifest's shape (ExtendedOverrides' place and address, the Open button running `openDesigner` and no ShowTaskpane action, the shared runtime, every other ribbon function still there), the shortcuts file against Microsoft's rules (unique ids and names, a registered action for each, one letter plus modifiers, no repeats), each navigation target, each action run against a document, and the public manifest carrying both with the host's address. The manifest also passed Microsoft's own check (`office-addin-manifest validate`). Whether Word accepts it and honours the keys is for the run-through.

## Sign-off (2026-10-30)

The author accepted the suggested answer to each of the five points below. Point 2 stays provisional until Stage 1 proves it in real Word.

1. **Existing text.** Is "standard pages go at the very top, existing text stays below" right for a document you have already begun writing?
2. **Heading styles.** Modules and lessons use styles based on Heading 1 and 2 so they can be told apart from other headings. I must prove in Stage 1 that Word's Navigation pane and checker treat them as headings. If they do not, modules and lessons use Heading 1 and 2 directly.
3. **First-release cues.** Your 14 icons plus the text cues Say and Ask, or a longer list now?
4. **"Shown to" choices.** Everyone and Facilitator only for Stage 1, with Learner only added when the workbook arrives in Stage 3?
5. **Time wording.** The lesson time block reads "Time: N minutes. Cumulative time: M minutes." Is that the wording you want?

## Done when (checked in your Word)

Starting works on a blank document and on one with text, and nothing existing is lost. Contents shows the heading structure. Add module, lesson or block takes three clicks or fewer, and each has a keyboard route. Ctrl+Z undoes each insert. Save, close and reopen, and the pane picks the guide up. Word's Check Accessibility shows no errors. You finish a first guide without reading instructions.
