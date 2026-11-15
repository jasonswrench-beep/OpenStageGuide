# Forms and handouts spec

Draft 2026-11-13, for the author's sign-off before any code is written. It is step 3 of `STANDALONE-APP-PLAN-2026-11-13.md` and answers the first gap in `EXAMPLES-REVIEW-2026-11-13.md`. It reuses the machinery that already exists: audience is a paragraph style, the guide is read from its paragraphs with no hidden record, "Update guide" regenerates numbers, and a second check looks at the finished file.

## What was asked

"Yes, the whole kit" (2026-11-13): the program holds the script **and** the forms, planning kit, reference pages and layouts around it. Forms come first because 12 of the 16 example documents have them, and the Guide today can say only "give out a handout".

## What a form is

A **form** is a numbered, titled piece of one or more pages that a person fills in or reads: a learner (worksheet, self-assessment, action plan), a facilitator (planning checklist, answer key), an observer (feedback form), an organiser (materials and room checklists). The examples show these kinds:

| Kind | What it is | Seen as |
| --- | --- | --- |
| Reading handout | Text to read, with a title and a source line | Most guides |
| Worksheet | Labelled boxes or lines to write in | Most guides |
| Note-taking grid | Rows and columns, some cells blank, some filled | Standards guide, planning guides |
| Checklist | Tick boxes with who, when and notes columns | Planning timelines, equipment lists |
| Rating form | Statements against a scale, plus an evidence or comment column | Self-assessments |
| Self-scored instrument | Answer, add the columns, fill in a grid or graph, read the meaning of the score | Team and conflict surveys |
| Observer form | Criteria with Met, Not met and Comments, plus name, date, time and signature | Practice and assessment sessions |
| Evaluation form | Header fields, agree scales, quality scales, comment lines | Session feedback |
| Contract or commitment card | Fields and signature lines, sometimes two to a page or in duplicate | Closing activities |
| Card set | Role cards, match-up cards, question cards, cut apart | Role plays and games |
| Template page | A fixed shape such as a two-by-two or a force-field | Thinking tools |

## The parts a form is made of (the elements)

First release, in this order of need:

1. **Title** and **number** (generated, see below).
2. **Instruction or reading text:** ordinary paragraphs in the form's own styles.
3. **Writing area:** ruled lines (1 to 12) or an empty box of a set height.
4. **Field row:** a label with a line after it (Name, Date, Time, Signature, Organisation), short, medium or full width. Several fields can share a row.
5. **Check list:** real check boxes, each with its text, and optional Who, When and Notes columns.
6. **Table:** a header row of named columns; rows that are filled in, blank, or hold check boxes, number boxes or writing space; an optional **sample row** (shown as an example); an optional **totals row**; a header row that repeats on each page.
7. **Rating grid:** statements against a scale. This is the existing Likert item (rating statements) plus an optional **Evidence** column.
8. **Score grid:** each statement maps to a scale; the form shows the totals per scale and, per scale, the meaning of the score by range (the self-scored instrument).
9. **Card set:** a grid of cards (2, 4, 6 or 8 to a page) with cut lines, each with its text; used with the copies rule.
10. **Source line:** a short credit or permission line under the form.

Later: picture with a caption (the program's editor will have it), two-column layout (a table without borders is allowed here only if its header cells name the columns), fillable fields on screen.

## Where forms live, and how they are cited

- **A form is written once**, in a **Forms** place: either after each unit's script, or collected in the Appendix. The author chooses per guide.
- **The script cites a form by choosing it**, not by typing its number. The Handout cue gets a "Which form" choice. It then shows "Handout 5.4: Standards in depth" and, in the workbook, the page. Move, rename or renumber a form and **Update guide** corrects every citation, as it already corrects heading numbers and time lines.
- **Numbering scheme**, chosen once per guide: by unit (module number, then a sequence: 5.4), by letter (A, B, C), by sequence (1, 2, 3), or none. Suggested default: by unit when the guide has modules, otherwise letters.
- **The outline and the unit header list the forms.** The Training at a Glance page gets a **Forms used** column. A module or lesson header lists "Participant handouts: 5.2, 5.3", built from the guide, not typed.
- **How it is stored** (the one technical decision): a form's first paragraph carries a Word bookmark named for the form's id, and a citation is a content control tagged with that id. Both are ordinary Word-format constructs, so the project file holds them as paragraphs like everything else, Word export keeps them as real bookmarks and fields, and nothing is hidden. Resolving a citation reads the bookmark; a citation whose bookmark is gone is a Review error.

## Print specification

Each form can say **how it is printed**: copies per person, per team of a size, per table, per room, or a fixed number; paper size; orientation; one-sided or two-sided; and "editable download" or not. The facilitator copy shows it as a line under the title ("Print: 1 per person, Letter, landscape"). If the guide records the expected group (number of people, number of teams or tables), the guide also builds a **copies to make** list for each lesson and for the whole programme, which feeds the materials list in the planning kit.

## The workbook and handout packs

- A form for **everyone** goes into the participant workbook automatically, in order, with its fill-in elements and its own orientation (a landscape form inside a portrait workbook is a section of its own). A form for the **facilitator or an observer** never does. The leak check already looks at the finished file; it learns the form styles and refuses a facilitator-only form that reached the workbook, and checks that a score grid shows no answer key.
- "Turn to page 31" becomes a field that the output fills with its own page number. The guide cites the form; each output supplies the page.
- A **handout pack** collects the forms of one lesson, one unit or the whole programme into one file for copying. It builds on the delivery pack that exists.

## Accessibility

Header rows are marked as header rows. Check boxes are real controls with names that say which statement they belong to, as the Likert items already are. A table is used for data only: a form never uses a table to place things on the page. Blank lines are paragraph borders, not typed underscores. Nothing depends on colour alone. A form with no title, or a table with no named columns, is a Review error.

## Review checks this adds

- A form is cited but does not exist; a form exists and is never cited (a note).
- A form cited as "give out" has no print specification (a note), or says "per team" when no teams are recorded.
- A score grid has a statement that maps to no scale, or to two.
- A table has no header row; a field row has no label.
- A facilitator-only form is cited from a script block that is for everyone.
- Typed numbers that disagree with the generated ones (an imported guide).

## How it looks in the program

A **Forms** tool in the Sections menu (called Write until 2026-11-14), like Add module and Add activity: a list of the guide's forms with numbers, a **New form** choice that offers the templates below, and an editor for a form that uses the editing surface's table support. Citation is the Handout cue's new "Which form" choice. In the suite look: one main button and a "..." menu, sections that open and close, each element a card.

First templates: Reading handout; Blank worksheet; Note-taking grid; Checklist (who, when); Rating form with evidence; Self-scored instrument; Observer form (met, not met, comments, signature, date); Evaluation form; Action plan table (the existing action-plan item, reused); Commitment card; Role cards; Match-up cards; Two-by-two; Force-field; Planning checklist.

## Slices

Each slice ends with checks run by script, a screenshot, and nothing pushed without a yes.

- **F1 Model and numbering.** The form record, bookmarks, citations, the numbering schemes, Update guide, the Forms used column. Tests on the model; no screens yet.
- **F2 Elements.** Title, text, lines and boxes, field rows, check list, table (with header row, sample row, totals), rating grid with evidence; Word-format builders and tests; the workbook includes them; the leak check learns them.
- **F3 The Forms tool** in the program, with the first templates.
- **F4 Score grid and self-scored instruments.**
- **F5 Card sets.**
- **F6 Print specification and the copies-to-make list.**
- **F7 Review checks.**
- **F8 Orientation sections and the handout pack.**

## Not in this stage

Fillable fields on screen or in a PDF (paper first; the model keeps room for it); branching or logic; collecting responses; the add-in (it is maintained, not extended; the builders live in the shared core, so nothing prevents it later); picture elements until the editor has them.

## Decisions for the author

1. **Numbering default.** By unit when the guide has modules, else letters. Suggested: yes.
2. **Paper first.** Forms are for pens on paper; fillable on screen later. Suggested: yes.
3. **Expected group in the guide** (people, teams, tables) so copy totals can be built. Suggested: yes, optional.
4. **Where forms sit.** After each unit, or in the Appendix, chosen per guide. Suggested: yes, default after each unit when numbering is by unit.
5. **Landscape forms inside a portrait guide and workbook.** Suggested: yes.
6. **The first template list** above: add or remove any.
7. **Citation by choosing the form** instead of typing its number. Suggested: yes.

## Sign-off (2026-11-13)

The author approved the suggested answers to all seven decisions: numbering by unit when the guide has modules, otherwise letters; paper first; an optional expected group in the guide; forms after each unit by default when numbering is by unit; landscape forms inside a portrait guide and workbook; the first template list as written; citation by choosing the form. Build starts with slice F1.

## As built, slice F1 (2026-11-13)

Built and tested; the full suite passes except the one test that already failed before this work (Author's validator wants project format 0.49.0).

- `src/guide-forms.mjs` (pure): the form title ("Handout 5.4: Standards in depth", read as label, number, title), the numbering schemes (auto, unit, letter, sequence, none; a form outside a module takes a letter, and so does the Appendix), reading forms and citations from the paragraphs, the plan Update guide acts on, and the Review issues (FORM-MISSING as a failure; FORM-DUPLICATE, FORM-NO-ID, FORM-NAME, FORM-UNCITED as warnings; FORM-UNCITED only appears once the guide cites forms at all). A new id is never one a citation still uses.
- `src/guide-forms-ooxml.mjs`: the form title paragraph with its bookmark (`osg_form_<id>`) and the citation content control (tag `osg-ref:<id>`).
- A new style, OSG Form Title (outline level 3, not in the contents), and a `form-title` kind.
- Paragraphs may carry `marks` ({bookmarks, refs}), read from their XML by the in-memory document and rebuilt when a project opens; a Word document through the add-in reports none, so it simply has no forms. Citations and bookmarks are read from ordinary paragraphs only, not from table cells.
- `updatePlan(paras,{forms})` adds `form` and `citation` edits; the document applies them; the pane reads the guide's scheme from the document (`formSettings`, kept in the project file as `forms`).
- 18 tests, including a pane test where Update guide renumbers forms and fixes citations by the scheme the guide chose.

**Deferred to F1b:** the Forms used column on Training at a Glance. It changes a generated page that many tests pin, so it is done after the elements (F2), when the column has more to show. It will appear only for guides that have forms.

## As built, slice F2 (2026-11-13)

Built and tested. The full suite passes except the one test that already failed before this work.

- **Elements** (`src/guide-forms-ooxml.mjs`): form text, ruled writing lines (1 to 12), a writing box of a set height, field rows (one to four labelled fields, each with a line made by a tab leader), a source line, a checklist (real check boxes, up to three more columns), a table to fill in (named columns of text, blank, number or check kinds; an example row; empty rows; a totals row; a repeating header row), and a rating grid (statements against 2 to 7 points, with an optional Evidence column). Every table has a header row Word knows about, a caption and a description, and every check box has a name. Limits are written down in `FORM_LIMITS`.
- **Styles:** OSG Form Title Facilitator (a form for the facilitator only, shaded), OSG Form Text, OSG Form Line, OSG Form Field, OSG Form Box, OSG Form Source, each with a model kind.
- **Workbook:** a form's audience comes from its title style. A form for everyone is kept whole with its tables and check boxes; a form for the facilitator is left out whole and its words are added to what the leak check watches for; the facilitator title style is a facilitator style in the leak check. A citation inside a kept block line keeps its words and loses its tag; the bookmark is dropped. A table or check box inside a form follows the form, and a table headed "Lesson" inside a form is not mistaken for a module overview. The pane's summary counts forms for everyone and forms left for the facilitator.
- **Found while testing:** the leak check treats short generic words such as Done and Item as facilitator text when a facilitator form's table has them, unless `allowedTexts` is passed. The pane passes it; tests of the leak check must too.
- 22 tests (elements 12, workbook 8, and the earlier ones kept passing).

**Not yet:** the Forms tool and its templates (F3), score grids (F4), card sets (F5), print specification and copies (F6), the remaining Review checks (F7), orientation sections and handout packs (F8), the Forms used column (F1b).

## As built, slices F6 and F7 (2026-11-13)

Built and tested. The full suite passes except the one test that already failed before this work.

**F6, the print line and the copies to make**
- A form can have a **print line** directly under its title, in a style of its own (OSG Form Print): "Print: 1 per person, Letter, landscape." It reads copies (a number per person, team, table or room, "of 4" for a team size, or a fixed "30 copies"), paper (Letter or A4), orientation, sides, "editable download", and keeps any other words as notes. It is for the facilitator: the workbook leaves it out, the leak check watches for its words, and the workbook does not define its style.
- The guide's form settings may record **how many people, teams and tables** there will be (`group`). `copiesToMake` multiplies each print line by them (a team of four with only the number of people is worked out) and lists the forms that cannot be added up, with the reason. The Forms tool and the planning kit's materials list will show it.

**F7, Review checks**
- FORM-EMPTY-TITLE (a failure), FORM-EMPTY (nothing in the form), FORM-FIELD-LABEL (a field with no label), FORM-TABLE-HEADER (a failure: a table in a form with no header row; each table is judged on its own), FORM-AUDIENCE (a form for the facilitator cited from a block for everyone), FORM-PRINT-NAME, FORM-PRINT-GROUP (a print line that needs a group size the guide does not give) and FORM-PRINT-NONE (a form that is cited has no print line). Together with F1's FORM-MISSING, FORM-DUPLICATE, FORM-NO-ID, FORM-NAME and FORM-UNCITED.
- 19 tests.

**Still to build:** the Forms tool (F3, waits for the program's shell), score grids (F4), card sets (F5), orientation sections and handout packs (F8), the Forms used column (F1b).

## As built, slices F4 and F5 (2026-11-13)

Built and tested. The full suite passes except the one test that already failed before this work.

- **Card sets (F5):** `cardSet` writes each card as one bordered paragraph in a new style (OSG Form Card) with a dashed line between cards, the line to cut along; `perPage` (2, 4, 6 or 8) starts a new page after that many cards; the label sits above. A card set in a form for the facilitator stays out of the workbook.
- **Self-scored instrument (F4):** `scoreInstrument` writes three tables, each with a header row, a caption and a description: the statements to rate (a number box for each score), a grid to add the scores up by scale (the statements to add, a total box, and optionally "divide by" and an average box), and what each range of totals means (for one scale, or for every scale). It will not build an instrument in which a statement is in no scale or in two, a scale names a statement that is not there, two scales share a name, a range runs backwards or has no meaning, or a range names a scale that does not exist. Scores are written, not ticked.
- Both read as parts of their forms, so the Review finds nothing wrong with them, and both go into the workbook whole for learners.
- 9 tests.

**Still to build:** the Forms tool and its templates (F3, waits for the program's shell: adding a tool to the add-in's tool list would change its ribbon, so the program needs a tool list of its own), orientation sections and handout packs (F8), the Forms used column (F1b).

## As built, slice F3, the Forms tool (2026-11-15)

Built and tested in the working tree (not yet in a published setup). The full suite passes (1,842 tests; the program's browser check is 124, with the Forms tool opened in a real browser).

- **The tool** is in the Activities menu, after Activity library (the spec said the Write menu, which became Sections and Activities on 2026-11-14), `forms` in `PROGRAM_TOOLS`, section `guide-sec-forms` of the Guide screen. It is a program-only tool: the add-in's ribbon is unchanged.
- **The list**: every form in the guide, as written on the page, with who it is for and its print line, and a Go there link. A guide with none says so and points to New form.
- **New form...** asks the program's page for its own New form dialog (the page's `osg-run` message; the same dialog as Insert > Form...), which puts the form where the cursor is.
- **Numbering**: the scheme (by unit when the guide has modules otherwise letters, by unit, letters, numbers, none) is chosen in the tool, kept in the guide's form settings, and Update guide uses it.
- **Copies to make**: how many people, teams and tables (each optional, 1 to 9999) are kept in the same settings; the print lines are added up in a table, and a form that cannot be added up is listed with the reason. The Review's FORM-PRINT-GROUP warnings now say where to give the number, and clear when it is given.
- **Fifteen templates** (the spec's list): the eight there were, and the self-scored instrument, evaluation form, role cards, match-up cards, two-by-two, force-field and planning checklist.
- **Found while building this**: every form made from a template carries a print line "1 per person", so the Review's "how many people there will be" warning appeared for every new form with no way to answer it. That was the manual author's second finding on alpha 5.

- **Which form** (the Handout cue): choosing the Handout cue shows a "Which form" list of the guide's forms (a form for the facilitator is marked as such). A block cites the form chosen: the words typed come first, then the form's name as a citation (a content control tagged with the form's id), then the minutes. A form alone is enough for a block. Update guide keeps the name and number in the block right when the form is renumbered, moved or renamed (tested through a change of scheme). **Edit this block** on such a block shows the form and the words before it, and Update block keeps the citation, changes it to another form, or takes it away with No form. A block that cites a form anywhere else (in the middle of its words, in its notes, in a block that is not a Handout) is not rewritten by the pane, which says so instead, so the link is never lost quietly.

## As built, slices F1b and F8 (2026-11-15)

Built and tested in the working tree (not in a published setup).

- **F1b, Forms used.** A guide whose blocks cite forms gets a fourth column, **Forms used**, on Training at a Glance: each lesson lists the forms its blocks cite, and a module row lists those of its lessons (a form is named by its label and number, "Handout 1.1", or by its title while it has no number). A guide that cites no form keeps the three columns it always had, and the page says it is out of date when a block starts to cite a form. A module's overview names the forms its blocks cite in lines under its lessons table: "Participant handouts: ..." for forms for everyone and "Facilitator forms: ..." for the others. Both are found again by their shape, so building the page again replaces them. `src/guide-pages.mjs` (`formsCited`, `lessonForms`).
- **F8, a landscape form is a section of its own.** A form whose print line says landscape is put in its own Word section in the guide's Word file (the portrait section ends before it and a portrait one follows), in the participant workbook (with a footer of its own whose page number sits at the wider page's margin; the cover keeps its first page without a footer), and in a handout pack. Two landscape forms in a row share one section; a guide that ends in one ends in a landscape section. `landscapeFlags` and `formSpans` in `src/guide-forms.mjs`; `bodyWithContents` in `src/guide-docx.mjs`; `buildWorkbook` in `src/guide-workbook.mjs` (the leak check expects the second footer part). Not checked in a real Word, nor with the Word schema validator (its Python module was not installed on the computer used); the files were checked by the tests and as well-formed XML.
- **F8, the handout pack.** The Forms tool's **Handout pack** makes one Word file of the forms for everyone from the lesson the cursor is in, the module it is in, or the whole guide: the forms inside that part and those cited from it, each starting on a page of its own (a change of orientation starts one already), without the print lines and without a form for the facilitator only. It is made from the guide's own styles and fonts like File > Save as Word file. `src/guide-handouts.mjs`; the program's page makes and gives the file (`saveHandoutPack` in `program/ui/word-files.js`).

**Left in the spec:** fillable fields on screen or in a PDF, and a pack in an order other than the guide's.

