# Session templates, activities, materials and decks (alpha 11), as built 2026-11-23

Built from the owner's product specification of 2026-11-22 (session templates, activities, materials and decks) and its template and activity catalog (16 templates, 46 activities, 102 agenda blocks).
All four milestones of the specification are in this release. This page says what was built, the decisions that shaped it, what is not done, and what was checked.

## What it adds

| Milestone | What the program now does | Where |
|---|---|---|
| 1. Authoring | **Start > Session templates**: browse 16 ready-made sessions (10 launch, 6 expansion pack) by words, purpose, length and set; preview one without changing anything (outcome, prerequisites, why the activities come in that order, materials, ways to adapt, the agenda with its clock times); put a copy in the guide as a lesson or as a module; save a lesson as one of your own templates, choosing whether its case studies and notes are kept. **Activities > Activity library** lists the 46 catalog activities with the 13 starters, filtered by phase and by how people work, and says where an activity will go. The agenda shows each activity's phase. The Review reports timed steps that do not add up to an activity's minutes, a step with no minutes, a need that names an activity that is not there, and an activity made from a replaced catalog version. | `src/session-catalog.mjs`, `session-text.mjs`, `session-templates.mjs`, `guide-pane-templates.mjs` |
| 2. Materials | **Make > Session materials**: a run sheet (clock times, set-up, supplies, timed steps with the clock time of each, debrief prompts, what goes wrong, other ways to take part; private notes only when asked) and a preparation checklist (each supply once, with the activities that use it, print / bring / open on a screen, for each person, pair, group or room, how many from the people and teams the guide gives, paper, status). The participant workbook has five **layouts** (Facilitated Workshop, Skills Practice, Case and Discussion, Reflection and Transfer, Assessment and Evidence). Everything made is listed with its status: Ready, Review needed (and what changed, named), Excluded; one output can be made again without touching the others. | `session-materials.mjs`, `workbook-layouts.mjs`, `session-record.mjs`, `guide-pane-materials.mjs` |
| 3. Decks | **Make > Decks and slides**: card decks (prompt or scenario cards, a participant toolkit, facilitator pocket cards) edited with a preview of both sides, an overflow check, a print sheet on Letter or A4 (1 to 8 to a page, fronts and backs mirrored for the flip, crop marks, page order for a printer that prints both sides or for two passes), a text view, and a line naming the deck in its activity's materials. The session as an editable 16 by 9 presentation file with native text and speaker notes, one slide for the opening, each activity, each break and the close. | `session-decks.mjs`, `card-print.mjs`, `pptx-export.mjs`, `guide-pane-decks.mjs` |
| 4. Reuse | The six expansion templates are the additional pack (a filter on the template list). **Bring changes back from a presentation**: slides are matched to activities by name, only differences are listed, the author ticks a subset, a change to something the guide also changed is marked and not ticked, an unmatched slide is never applied. Optional AI help: the part that does not need a model (see below). | `session-slide-import.mjs`, `session-assist.mjs` |

## Decisions

- **No hidden record in the guide.** The guide stays ordinary paragraphs. An activity put in from a template says where it came from in a visible, facilitator-only `Source:` part ("Scenario rehearsal (A13, version 1)"), and a lesson in a `Template:` fact. Bookmarks were not used: the editor copies them as they are, so a pasted block would have two owners of one name.
- **One new project member, `sessions`** (like `design` and `forms`), holds only what cannot be read from the paragraphs: which outputs were made, a fingerprint of each activity's words that matter to that output, and the card decks being edited. An output is "Review needed" when those words change; nothing is rewritten on its own, and an earlier copy that was edited by hand is never overwritten.
- **A session is a lesson.** A template is applied as a lesson (in the module the cursor is in) or as a module of its own; each row of the template is an Activity block written as labelled parts. Eight new parts exist (expected output, grouping, phase, needs, timed steps, debrief prompts, likely challenges, access options, source); the participant sees only the ones written for them, so the workbook's leak check holds.
- **The catalog was completed, not copied.** The spreadsheet's steps are written for the facilitator, so each activity also has a participant task, what it needs from earlier work, debrief prompts, likely challenges, materials apart from access, and a mode; each template has a family, prerequisites and why its blocks come in that order. All of it is written for this tool.
- **The reference tabs about another product's catalog are not in the repository,** and no competitor product name is in the data. Method names are only names; the catalog's "origin" credits were not carried over because they could not be checked (see TESTING-TODO.md).
- **The program is the target.** Nothing here touches the Word add-in's repository; the core pieces are in `src/` and the add-in picks them up when it takes this core.
- **No AI model is built in and nothing is sent anywhere.** `src/session-assist.mjs` holds what does not depend on one: what a helper may be given (only what the author selected, never the source line or private notes), a check of what it proposes, proposals shown as the same list of changes to accept, edit or discard, and a missing, failing or badly answering helper changing nothing. The page says that no helper is set up.
- **The spec's colour tokens were not adopted.** The program keeps the suite's look; phases are shown in words, never by colour alone.

## Not done, and said so

- Importing and exporting a template pack as a file (the expansion pack is part of the program; your own templates are kept in the program's settings file).
- A delivery-mode filter on the templates: no template in the catalog is for one mode only, so the filter would have one value.
- The AI helper itself (a provider, its key, and a decision about what may leave the computer).
- Template lengths of 58, 57, 77 and 73 minutes keep their honest totals (no buffer rows added).
- A shared "Preparation checklist inserted into the guide" (it is a Word file); the card decks are kept in the guide's file and made as a Word file.
- The Word add-in's task pane widths (270, 320, 400 px) and host matrix: they belong to that repository.

## Checks

- 2,226 automatic tests (the catalog's totals and steps, the template tool, the run sheet and checklist, the workbook layouts through the leak check and structure check, the record and its status, the decks and slides, the slide import, the AI seam).
- `npm run check:sessions`: 30 checks in a real browser (the menus, the panels in the tools frame, what is written into the guide and its file, the Word and presentation files made, a change flagging what was made). The other five browser checks pass as before (192, 29, 7, 6, 42) and the editor's 289.
- The run sheet, the checklist, the five workbook layouts, the card print sheets (Letter and A4) and the slides pass the Word and presentation schema validators; the presentation file was also opened in the presentation program installed on this computer without a repair prompt. None of these was opened in a real Word.
- Found on the way and fixed: the workbook's settings part listed its elements in an order the Word schema does not allow.
