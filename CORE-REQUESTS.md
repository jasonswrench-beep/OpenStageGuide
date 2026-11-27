# Requests to the shared core

Things OpenStage Guide needs from `openstage-core` that the core does not offer yet. Guide does not build its own version of a shared piece when the core could offer it; where it
had to write something to finish its list, the entry says what and says what replaces it once the core has the piece. Newest last.

Core as read, 2026-11-27: `CORE-NOTES.md` on its GitHub `main`, and its tags `pstat-v1.2.0`, `pstat-v1.3.0`, `item-analysis-v1.1.0`, `item-analysis-v2.0.0`, `prismqda-stats-v1.0.0`, `capture-v0.1.0`,
`a11y-v0.1.0` and `bank-v0.1.0`. The accessibility rules (C2) and the bank reader (C4) are published; Guide uses neither yet (see G2 and G3). Nothing unpublished was read or used.

**What Guide builds on from the core, as published.** `@openstage/pstat` 1.2.0 (from its release, no sign-in) is a development dependency: Guide's tests check every `.pstat` file it
writes with the package's `checkPstat`, and the browser checks do the same on the files the program saves. Guide writes the files itself, as the core's notes say a tool with no runtime
dependency may (the program's page makes the zip with JSZip), and writes no profile, because the core has none for what Guide sends (G1). Nothing of the core's code is copied.

## G1. A `.pstat` profile for an instrument and its scoring guide, and one for an evaluation plan

**For:** development list item 6 (rating and self-scored instruments to PrismStatistics with their scoring guide) and item 11 (an evaluation plan by Kirkpatrick level that PrismStatistics can open).

Guide sends two kinds of file today, both plain format 1 with no `profile`: one variable for each statement (its words as the label, the points of its scale as value labels, ordinal), empty
rows for the answers, and a `respondent` number. The scoring guide cannot ride as data, so it is written as plain sentences in `source.report.notes`, which PrismStatistics shows when it
opens the file: how each scale is added up ("Scale "Confidence": add q01, q03, q05, then divide by 3"), what each range of totals means, and for an evaluation plan each level's measures,
when, of whom, and for which objectives. A person can read it; PrismStatistics cannot use it.

Asked: a profile for each, with a part, written the way section 7.2 and 7.3 are, so PrismStatistics can compute the scale scores and show the plan without anyone retyping it:

- `OpenStage Guide instrument 1`, part `instrument.json`: the scales (name, the variables by `id`, how they are combined: sum or mean, which items are reverse-scored and on what scale), the
  ranges of totals and what each means, and the scale's points.
- `OpenStage Guide evaluation plan 1`, part `evaluation-plan.json`: for each measure its Kirkpatrick level (1 to 4), kind, when, of whom, and the objectives it is for (their words and level); the
  variables it gave rise to, by `id`.

Guide has no notion of a reverse-scored item yet; if the profile has one, Guide will add it to the self-scored instrument. A test set of two or three small instruments with the scale scores worked
by R would hold the writer and PrismStatistics's reader to the same figures. Until then Guide keeps the notes.

## G2. The accessibility rules (C2) as a package a JavaScript tool can run, for the single-page web export and the PDF

**For:** development list item 4 (its own tagged PDF export, with C2 on the single-page web export).

Guide's File > Save as web page and File > Save as PDF make one HTML page of the guide (`src/guide-html.mjs`), which Edge prints to a tagged PDF with bookmarks. Guide checks the guide's document
(color contrast, tiny text, headings, picture descriptions, table headers, forms) in its Review, and checks the PDF it gets back for its structure tree, language, title and outline. It has no check
of the page's own markup against the WCAG 2.1 AA success criteria, and the core's list says C2 is the one rule set for that.

Asked: when C2 is published, a form a JavaScript program can call without a Rust build: a function that takes a page's HTML (and says it is a single page, a course package or a help site) and gives
back findings, each with a rule id, the success criterion, a status in the core's words (pass, fail, needs review, not tested) and where in the page; with the "warn, never block" rule and the
opt-in blockers as a setting the caller passes. Guide will run it on the web export in the Review (Accessibility), list the findings, and never stop a save or a PDF because of one. It writes no rule
of its own meanwhile.

**Partly answered (2026-11-27).** The core published A11Y-1 revision 1.0 (tag `a11y-v0.1.0`): the engine, in Rust with a command-line program, checks web pages, Word and PowerPoint among others (PDFs not
yet), and `@openstage/a11y` 0.1.0 gives the catalogue, the status words, contrast and the settings in JavaScript, but does not check a page. A build of the engine for the browser is on the core's roadmap.
Guide uses neither yet; the request for a call that takes a page's HTML stands.

## G3. The bank reader (C4)

**For:** the design desk (items 7 to 11) and the existing bank import.

Guide reads OpenAssess's `.openassess` banks in `src/guide-bank.mjs` (it reads format 1 and takes the questions with their objectives, levels and chapters). The core lists C4 as not started and
Guide's list names it as a user. Asked: when C4 is made, start from what that file does, and say when to move Guide onto it; it holds no code the core would need to keep.

**Answered (2026-11-27).** `@openstage/bank` 0.1.0 (tag `bank-v0.1.0`, BANK-1 revision 1.0) started from `src/guide-bank.mjs`. Guide has not moved onto it yet. The core's notes for the move: keep
`questionToItem` (what a guide item can hold) and read `q.key` and `q.options` in place of the raw body; the reader hands over retired questions (`metadata.status` "retired") and leaves the choice to
Guide; a picture is in `q.pictures`, an equation in `q.equations`, and their placeholders are in `q.stem.html`.

## G4. How a design hands item slots to OpenAssess

**For:** development list item 10 (empty item slots tagged with objectives to OpenAssess).

OpenAssess has no import of slots: its question table (.xlsx or .csv) refuses a row with no question, and no status marks a question as a placeholder. Guide therefore writes a question table
(`src/design-storyboard.mjs`, `itemSlots`): an essay row for each slot with a placeholder stem ("[Slot] Write a question that checks: ..."), `Objective`, `Bloom`, `Chapter` (the module),
`Section` (the lesson) and `Tags` ("slot; quiz; Understand"), `ID` ("guide-o1-s2-1"). A person saves it into a bank with OpenAssess's Import.

Asked (to the core and OpenAssess's session): a small spec of what a design sends to an assessment tool: an objective (its words, its level, an id that stays), the module and lesson, the number
of items wanted, and the kind of check; and in OpenAssess, an importer that accepts a row with an objective and a level and no stem, and a "slot" status that keeps it out of tests and exports until it
is written. Guide will write that instead of the placeholder stems.

## G5. A place for objectives in the course formats

**For:** development list item 10 (a skeleton into Author or Scroll).

OpenStage Author's project (schema 0.52) and OpenScroll's course (schema 3) have no field for a lesson's objectives or their Bloom level; both refuse unknown keys. Guide writes them as words: a slide
and a note in Author's hand-off notes file, a statement and a list on the lesson's page in Scroll. Asked: an optional list of objectives (their words, level and id) on a scene or lesson in both formats, so a
skeleton keeps what a design said and a bank (C4) can be linked to it.

## G6. A client of Prism Reference Manager's local API, and two things from it

**For:** development list item 9 (each design choice cited through Prism Reference Manager's local API).

Prism Reference Manager's list (its item 3) says other tools cite through its API. Guide wrote the client itself because the core has none: `src/prm-client.mjs` (pairing, `capture.check`,
`refs.add`, `cite.render`) and `scripts/prm-link.mjs`, made on the program's server because a page cannot call the API and the token must not reach it. It is checked against a stand-in that answers as
PRM's `docs/local-api.md` and code (api.rs, dispatch.rs) say, not against the real program. Asked: a shared client in the core when more than one tool needs it (PrismQDA and PrismStatistics reports
will), which Guide will move onto and delete its own. For Prism Reference Manager's own session: a call that gives its version and whether a library is open (today only `library.info` after pairing
says anything, and there is no way to find the port), and a documented request and answer for each call (the documentation lists the calls and not their shapes; Guide's came from the code).

## G7. Opening a file in another tool

**For:** development list item 10.

Guide's hand-offs (the Author skeleton, the Scroll file, the item slots) are files a person carries: none of the three tools can be asked to open one from outside. That is how the suite's design plan says
tools link, and Guide builds it that way. If C1's "send to" path is made general (a way to ask an installed tool to open a file), Guide would use it for these three. Recorded, not urgent.
