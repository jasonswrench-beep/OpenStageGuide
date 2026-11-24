# The design desk (development list items 7 to 11), as built 2026-11-16

Learning design lives in Guide, in front of its learning alignment, agenda and timing, not in a tool of its own (the settled decision of the list). It is a **Design** menu with five screens that open one
**design desk**, which takes the place of the guide and its tools while it is open (**Back to the guide** gives them back as they were). AI is not used anywhere in it: every part works with no model.
The author's words and choices are what is kept; the program advises and never refuses.

## Where the notes are kept

In the guide's own project file, as one more member of `guide`: `design`, the text of the notes (`src/design-model.mjs`), like the form settings. `readDesign` and `cleanDesign` check them when the desk is
opened (unknown fields dropped, every id a list points at must exist, every list has a limit, a message fit to show when something is beyond them). The notes are saved 0.6 seconds after the last change
and when the desk is left, and travel with a backup, a restore and **Make a new guide from the storyboard**. They do not take part in Undo: Ctrl+Z is still the guide's.

## The five screens

| Screen | What it does | Code |
|---|---|---|
| **Needs and objectives** (item 7) | A needs analysis (the problem, the learners, what they do and should do, what is in the way, limits, how success is known) with a plain verdict on whether training can help. Objectives written freely or built from the verbs of Bloom's six levels (twelve a level); each is read for its verb: the level it gives, a weak verb ("understand", "know"...) with verbs that could stand in, two verbs in one objective, a verb with nothing after it, too many words. All advice. | `src/design-bloom.mjs`, `program/ui/design-needs.js` |
| **Models and activities** (item 9) | Nine named models (Nine events of instruction, First principles of instruction, ARCS, Backward design, Experiential learning cycle, 5E, 4C/ID, Action mapping, SAM) and ADDIE, each with its events and ideas for activities that fit each; one click adds an activity to the plan, to be linked to the objectives it serves. Each model names the work it comes from, cited in APA 7 by the program; with **Prism Reference Manager** connected it is cited by PRM (its local API) and the work is kept in the person's library, and either can be put in the guide's list of references. | `src/design-models.mjs`, `src/prm-client.mjs`, `scripts/prm-link.mjs`, `program/ui/design-activities.js` |
| **Course map** (item 8) | Guide's learning alignment, extended into a map from objectives to activities to assessments. The assessments are planned here (kind, the level it checks at, the objectives it checks, item slots for each). Two maps: **the plan's** (from the notes) and **the guide's** (from what the guide holds, by the same likeness rule as Learning alignment), each with a row for each objective, its level, lesson, activities and assessments, and the gaps in words: no activity, no assessment, an assessment checking below the objective's level, a weak verb, no lesson. Each is a table for a syllabus or an accreditation report: a CSV (UTF-8 with a mark), or a landscape form put in the guide. | `src/design-map.mjs`, `program/ui/design-mapview.js` |
| **Storyboard** (item 10) | Modules and lessons with minutes, the objectives each teaches, its activities and its assessments, moved and removed with buttons. From it: **a new guide** (the guide's own course plan: times, an Objective block for each objective, an Activity block for each activity with its model in the facilitator's notes, an Assessment block, a Transition at the end of each module; it carries the notes); **a skeleton for OpenStage Author** (a scene for each lesson, a slide for each lesson, activity and assessment, the levels in the hand-off notes); **a skeleton for OpenStage Scroll** (an `.openscroll`: a section for each module, a lesson for each lesson with its objectives and activities as lists); **item slots for OpenAssess** (a question table, an essay row for each slot with a placeholder stem, tagged with the objective, its level, the module as chapter, the lesson as section, and "slot"). | `src/design-storyboard.mjs`, `program/ui/design-board.js` |
| **Evaluation plan** (item 11) | Measures by Kirkpatrick's four levels (reaction, learning, behavior, results): what is measured (statements on a scale, a score before and after, one score, a work measure before and after), when, of whom, for which objectives. What is missing is said in words. It goes to PrismStatistics as a `.pstat` (a variable for each measure with its level in the label, empty rows, the plan in the notes), and to the guide as a table. | `src/design-evaluation.mjs`, `program/ui/design-eval.js` |

## Checks

- Unit tests for each module (`tests/design-*.test.mjs`, `tests/prm-client.test.mjs`): the Bloom verbs and checks, the notes, the models and their citations, the maps, the storyboard's four outputs, the evaluation file, the client of Prism Reference Manager against a stand-in.
- **The files are held to the other tools' own checks**: the Author skeleton passes Author's `tools/validate.py` with 0 errors (runs when `OPENSTAGE_AUTHOR_VALIDATOR` names its `validate.py`); the Scroll course passes OpenScroll's published `course-v3.schema.json` (opt-in: `OPENSCROLL_SCHEMA` and `OPENSCROLL_AJV_DIR`; it passed); both `.pstat` files pass `checkPstat` of `@openstage/pstat` 1.2.0.
- `npm run check:design` is a browser check of the desk: the menu, the screens, what is kept in the project file, back to the guide, the files it makes, a new guide from the storyboard.

## What the core does not offer yet

In `CORE-REQUESTS.md`: a profile for an instrument and its scoring guide and for an evaluation plan (G1), the accessibility rules in a form a JavaScript tool can call (G2), the bank reader (G3), a slot format for
OpenAssess (G4), a place for objectives in Author's and Scroll's formats (G5), a shared client of Prism Reference Manager's API (G6), and a way to open a file in another tool (G7).

## Not done, on purpose

- No AI drafting of needs, objectives or storyboards (item text: "AI stays optional"; it is absent, not optional here). A model could later propose; a person would still accept.
- The map's levels for an assessment item come from what the author says; an item in the guide has no level of its own.
- Reverse-scored items: Guide's self-scored instrument has none, so the `.pstat` notes cannot say so.
- The real Prism Reference Manager, PrismStatistics, OpenAssess, Author and Scroll have not been run on these files by this session: see `TESTING-TODO.md`.
