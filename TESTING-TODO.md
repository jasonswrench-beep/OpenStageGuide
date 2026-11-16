# What only you can check

Written 2026-11-16, in the round that did the development list of 2026-11-15 (items 1 to 11). Each line says what to do and what it should do. Everything the program could check by itself
(the automatic tests, the browser checks on the program, the editor, the design desk and the files it makes, Author's validator, Scroll's published schema, the core's `checkPstat`) was run and is
in the commit notes; what is below needs a person, a screen reader, another program or another computer. Tick a line when you have done it, and write on it what you saw.

## Item 5 of the list: your checks

- [ ] **A screen reader pass (NVDA) of the whole program.** It records everything spoken on the computer, so it waits for your go. Include what is new: **Insert > Block...** (the chooser, then typing in the block
  that appears), **File > Save for PrismStatistics...**, the **Design menu** and the five screens of the **design desk** (the tabs: arrow keys move between them; each list of ticks is a group with a name; a
  warning about an objective is read after the box as you type). Say what is read wrongly or not at all.
- [ ] **The Word files in a real word processor.** Open in Word (and, if you have it, LibreOffice) a guide, a workbook, a handout pack, a guide with **each block's icon in a column** and **a slide's picture
  beside its script**, a **bulleted list**, a **picture in a table cell**, and the **course map** and **evaluation plan** tables put in a guide from the design desk (landscape form). They pass the Word schema
  validator and were never opened in a real Word. Look for: icons beside their words, pictures where they were put, bullets, the landscape pages.
- [ ] **A second computer.** Install the setup (see below for the one to use), start it, make a guide, save it as Word, web page and PDF, and uninstall. Needs Edge for the PDF (Windows 11 has it). Does
  Check for updates say what it should?

## What is waiting for your yes

- [ ] **Publish.** Nothing from this round is pushed or published. The commits are on `main` here, and alpha 9 (built and signed before this round) is not published either; its fixes are being made in another
  working folder (`C:\Dev\OpenStageGuide-a9`, branch `alpha9-fixes`) by another session, and will need merging with this branch. Tell me when, and in what order: push, then a build of the next alpha with
  everything, then the release as in section 9 of `docs/BUILD-STATUS-CHECKLIST-2026-11-15.md`.
- [ ] **The line "Co-Authored-By: Claude" at the end of my commits.** Your older commits have none. If you want none, say so and I will rewrite the messages of this round's local commits before anything is pushed.

## Item 2: alpha 8 and the Forms

- [x] Alpha 8 is published as a plain release (checked on GitHub 2026-11-16: not marked pre-release, setup and checksum attached). The **Forms tool is in alpha 7 and alpha 8** (checked in each tag's source).
- [ ] Download alpha 8 once more, install it and open **Activities > Forms**. (The release was downloaded again and its signature checked when it was published; this is only whether you see what you expect.)

## Item 6: PrismStatistics opens the file

- [ ] Make a guide with **Insert > Form > Self-scored instrument** and a **Rating grid**, then **File > Save for PrismStatistics...** and open the `.pstat` file in **PrismStatistics**. It should open with a variable for each
  statement, the statement as its label and the points as value labels, empty rows, and the scoring guide in the notes of the import. The file passes the core's checker and what the program reads, but it has
  not been opened in PrismStatistics's own window here. Say if the notes are shown, and whether you would rather the scale totals were made for you (that needs the profile asked for in `CORE-REQUESTS.md` G1).

## Items 7 to 11: the design desk

- [ ] **Use it on a real course.** Design menu > Needs and objectives: write the needs analysis and a few objectives. Does the advice about verbs help or nag? Are the verb lists the ones you would use?
- [ ] **The models.** Are the nine models the ones you want, and are the activity ideas under each fit to your teaching? The ideas are mine, not the models'. Add or take away what you like (it is the list in
  `src/design-models.mjs`).
- [ ] **The citations.** The program writes each in APA 7 itself. Seven have a DOI that was confirmed with Scite (Merrill 2002, Keller 1987, Krathwohl 2002, Sweller 1988, Branch 2009, Mayer 2009, van Merriënboer and
  Kirschner); the rest are books with no DOI: **check each edition and year against your copy** (Gagné 1985, 4th ed.; Wiggins and McTighe 2005; Kolb 1984; Bybee et al. 2006; Moore 2017; Allen and Sites 2012;
  Anderson and Krathwohl 2001; Kirkpatrick and Kirkpatrick 2006, 3rd ed.). The 3rd edition of *Ten steps to complex learning* is dated 2018 here and 2017 in Scite's record of the online edition.
- [ ] **Prism Reference Manager.** Start it, turn on **Allow local connections** in its Settings (port 23119 unless you changed it), then in the desk (Models and activities) choose **Connect** and **Allow** in
  its window, then **Cite this with Prism Reference Manager**. The program's client was checked against a stand-in that answers as PRM's documentation and code say; **it was never run against the real
  program.** If it fails, the message says what PRM said. Also check: the work shows under Machine-created in PRM's library, the citation is APA 7, and **Put this reference in the guide** adds it to the
  guide's list of references.
- [ ] **The course map.** Plan a few assessments, read the map, **Put it in the guide**, save the CSV and open it in a spreadsheet (accents right?). Would your syllabus or an accreditation report take this table as it is? What
  column is missing?
- [ ] **A new guide from the storyboard.** Does the guide that comes out have the shape you would start from? (Each lesson: its time, an Objective block, its activities, its assessments; a Transition at the end of
  each module.)
- [ ] **The Author skeleton.** Unzip the file, open the `.osproj` folder in **OpenStage Author**. It passes Author's own validator. Does it open without asking to be updated, and are the scenes and slides as you expect?
- [ ] **The Scroll skeleton.** Import the `.openscroll` file in **OpenScroll**. It passes Scroll's published JSON schema; it was not imported into the real Scroll.
- [ ] **The item slots in OpenAssess.** In OpenAssess choose Import, pick the CSV, review, save it into a bank. Each slot should arrive as an essay question with its objective, Bloom level, chapter and section, and the tag
  "slot". OpenAssess may warn about an essay with no answers; say what it says.
- [ ] **The evaluation plan in PrismStatistics.** Save it for PrismStatistics and open it: one variable for each measure (the level in its label), empty rows, the plan in the notes.

## Things I noticed and did not fix

- **OpenStage Author is moving ahead of the format its list named.** The list said Author's validator expects schema 0.49; that is what Author has published (its `main` and the branch it is on, checked
  in its committed `tools/validate.py`), and Guide writes 0.49.0, which that validator passes with 0 errors. Author's working copy, being worked on by its own session, is now at 0.50 (it adds the
  optional `accessibilityBlockers`, the opt-in publish blockers) and then 0.51, unpublished, and its validator says a 0.49 project "needs updating" (its editor offers to). So the two tests that run
  Author's validator **skip, with that reason printed, while Author's working copy is ahead**, and fail for any other finding. When Author publishes its new version, Guide should write it: it is one
  constant (`AUTHOR_SCHEMA_VERSION` in `src/guide-author.mjs`) and a look at what the new version asks.
- Running the whole automatic suite at once, one test of `guide-lesson-details` failed once while another session was changing the same working folder, and passes by itself (17 of 17). I think it was that.
- Rows added to the guide with Insert > Block come with no words. The Review warns of a block with no words (BLOCK-EMPTY); it does not stop anything.
- The Reference Manager port has to be told (23119 unless changed): it has no way to be found. The token is kept in `OpenStage Guide Reference Manager.json` beside your guides, and is not shown to the page.
