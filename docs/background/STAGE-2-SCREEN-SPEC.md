# Stage 2 screen spec: the assessment gallery

Draft 2026-10-31, for the author's sign-off before any code is written. It follows the rebuild plan of 2026-10-30 (kept outside the repository) (Stage 2) and the Stage 1 rules (`STAGE-1-SCREEN-SPEC.md`). Stage 1 is proven in the author's Word; this stage reuses its machinery: styles for audience, block-aware placement, Edit, Update guide, Undo, and a real-Word check after each step.

## What was asked

"An assessment section where I click the assessment item type and quickly insert it." Learner version and facilitator key kept separate. The result passes Word's accessibility checker and survives save and reopen.

## What this builds on

- **The author's own template** (`Facilitator's Guide Template - 2023.docx`): *Level One Assessment* (a short survey: five-point agree/disagree statements grouped under headings such as Trainer/Facilitator, Logistics and Program, then "Program Improvement Suggestions"), *Level Two Assessment* (true/false, fill in the blank, multiple choice with a to d, matching in two columns), *Level Three Assessment* (a follow-up survey: did the training succeed on each learning objective, five-point scale). Each starts with a "Dear Participant" note and program information lines.
- **OpenAssess** (the author's separate question-bank project) already defines the question types and their answer rules: multiple choice, multiple response, true/false, fill in the blank, matching, ordering, numeric, short answer, essay (hotspot and calculated come later there). Stage 2 uses the same names and the same answer-key rules, so a bank question and a guide question mean the same thing. It does not read a bank file in this stage.
- **The older assessment code in this repo** (scorer, respondent files, exporter, 43 KB panel) stays in Legacy tools, untouched. It is a different thing: it scores responses and exports files. The gallery only inserts items into the guide.

## The item types

First release (nine). Each inserts in one click plus a short form.

| Type | For | The author fills in | What the learner version holds | What the key holds |
| --- | --- | --- | --- | --- |
| Multiple choice | Level Two | Question, 2 to 6 options, which is correct | Question, lettered options, a line for the answer | Correct letter, why, points |
| Multiple response | Level Two | Question, options, which are correct | Question, lettered options, "choose all that apply", answer line | Correct letters, why, points |
| True or false | Level Two | Statement, true or false | Statement, "True or false:" answer line | Correct answer, why, points |
| Fill in the blank | Level Two | Sentence with the blank marked, accepted words | Sentence with "[blank]" shown, answer line | Accepted answers, points |
| Short answer | Level Two | Question, expected points | Question, 2 to 5 blank lines | Expected points, points |
| Matching | Level Two | Items and their matches, extra choices | A table with a header row (Item, Your match) and a lettered list of choices | Item-to-letter pairs, points |
| Reflection or essay | Levels Two and Three | Prompt, optional rubric rows | Prompt and writing lines | Model answer or rubric, points |
| Likert scale (was Rating statements) | Levels One and Three | A heading, several statements, the scale (five-point agree, five-point succeeded, or your own labels) | A table with a header row: statement, then one column per rating | Nothing (a survey has no key); optional objective per statement |
| Open comment | Levels One and Three | Prompt, number of lines | Prompt and writing lines | Nothing |

Later, not in this stage: ordering, numeric, semantic differential, performance checklist or rubric, action plan, pasting many questions in the plain-text standard format at once, and linking to an OpenAssess bank.

## What an item looks like in the guide

Everything is real Word content in styles, so Undo, the Navigation pane and the accessibility checker all work. Illustration (multiple choice):

```
Question 1. Which step comes first when reframing a disagreement?          style: OSG Question
a. Name the feeling                                                         style: OSG Option, a real lettered list
b. Describe the facts
c. Ask for a change
Your answer:                                                                style: OSG Answer Line (a bordered line, not typed underscores)
| Answer key: b. Describing the facts keeps the first words neutral.        style: OSG Answer Key, shaded, facilitator only
| Points: 1. Objective: Reframe a disagreement.
```

- The question, options and answer line are for **everyone**. The answer key is **facilitator only**, in the same way as facilitator notes: a shaded paragraph whose style marks it, starting with the words "Answer key:". It is never hidden text, never white text, never a hidden row. The Stage 3 learner workbook leaves out every facilitator-only paragraph and is checked for leaks.
- **Question numbers** ("Question 1.") are kept in step by Update guide, the same way module and lesson numbers are, restarting in each assessment section. Option letters are a real Word list, so they are never typed by hand.
- **Tables** are used only where the content really is a table (rating statements, the matching items): a header row, no merged cells, no nested tables. A multiple-choice item is never a table.
- **Check boxes and answer spaces**: see decision 4.
- **Points** and **objective** are optional. They appear in the key only.
- **Each assessment is a section** with a heading (so it is in the contents), an editable "Dear Participant" note and the program information lines from your template. "Add a Level One, Two or Three assessment" inserts that starter; items then go inside it.

## The screen

A second screen beside Build, reached by a two-button switch at the top of the pane (Guide | Assessments). One focused screen, the main button always visible.

```
Assessments                                     [ Guide | Assessments ]

Your cursor is in: Level Two Assessment, Question 2     (same "In your document" idea)
[ Edit this question ]

Start a section
[ Level One survey ] [ Level Two knowledge check ] [ Level Three follow-up ]

Add an item
Check knowledge   [Multiple choice] [Multiple response] [True or false] [Fill in the blank] [Short answer] [Matching] [Reflection]
Surveys           [Likert scale] [Open comment]

(choosing one shows its short form)
Question  [__________]
Option a  [__________]   [ ] correct
Option b  [__________]   [ ] correct
[ add an option ]
Why it is right (optional) [__________]
Points [__]   Objective [__________]
[ Insert multiple choice ]
```

- An item goes **after the item the cursor is in**, never into the middle of one (the Stage 1 placement rule, extended to assessment items).
- **Edit this question** reads an item back into its form and replaces it in place, as Edit this block does.
- The form for each type has only the fields in the table above. Ctrl+Enter inserts. Every control has a text label and works from the keyboard.
- Update guide also renumbers the questions and refreshes the contents.

## What I cannot prove without Word (each is on the native checklist)

- That a lettered list inserted with an item keeps its letters and restarts for each item, and that the numbering survives Undo, save and reopen.
- That check boxes (if chosen) arrive as real, clickable check boxes and that Word's checker has no complaint about them.
- That a table with a header row arrives with its header row marked as a header, which the checker reads.
- That the answer-key styles arrive on every paragraph (the Stage 1 last-paragraph problem is handled the same way: a spare final paragraph and a read-back).
- That Word's accessibility checker is happy with a whole assessment section.

## Not in this stage

The learner workbook and its leak check (Stage 3). Looks and brand colours (Stage 3). Scoring, response collection and the old scorer (stay in Legacy tools). PDF and print (Stage 4). Reading an OpenAssess bank (later).

## Sign-off (2026-10-31)

The author accepted the suggested answer to each of the five decisions below.

## As built, slice A (2026-10-31)

Built in two slices so a problem in real Word on one cannot block the other.

- **Slice A (built):** the Assessments screen and its Guide | Assessments switch; the three section starters (Level One, Two, Three); seven item types, which are the ones that need only paragraphs and a lettered list: multiple choice, multiple response, true or false, fill in the blank, short answer, reflection or essay, open comment. Items insert after the item the cursor is in, are edited in place with "Edit this item", are renumbered by Update guide, and keep the Stage 1 safeguards (spare final paragraph, read-back, Undo).
- **Slice B (built after the first run in Word, at the author's request, see below):** matching and rating statements, the two types that need a table with a header row and, for rating statements, clickable check boxes. They were held back so a problem with the table and check-box risks could not hide a problem with the list and key-style ones; the author asked for them before the next Word test.
- **Where items go:** a section goes after the last module or section, or after the section the cursor is in; items need the cursor inside a section.
- **Key format:** each key is a run of facilitator-only paragraphs: "Answer key (multiple choice): b.", then optional "Why:", "Points:" and "Objective:" paragraphs. The first one names the type so the item can be read back and edited.
- **Writing lines:** a bordered paragraph style, not typed underscores. Review does not count them as empty paragraphs.
- **Open comment** has no number and no key. Question numbers restart in each section.
- **Not done yet:** Review checks that the key letter matches an option.

## Changes after the first run in real Word (2026-10-31, from the author's Round 4 recording)

What the recording showed in real Word: Level One, Two and Three sections, the heading in the Navigation Pane, "Question 1." with a real a. to d. list, the "Your answer:" line and the shaded key all came out right. What went wrong was the pane. The author filled in a whole item with the cursor outside a section, pressed Insert item three times, and saw a refusal at the top of the pane, out of sight from the button at the bottom, with no way forward there.

- **Items go in:** a line under "Add an item" says where an item would go before anything is typed ("Items go in: Level Two Assessment, after Question 1 (true or false)."), or that there is no section here yet. It follows the cursor.
- **The refusal sits above Insert item** (not at the top of the pane), says nothing was lost, and has a "Start a Level Two section here" button. Starting the section that way keeps what was typed and puts the focus on Insert item. Other problems with the item form (missing question, a bad points value) appear in the same place. A time-out from Word still shows at the top.
- **Section buttons say what they are:** "Level One: reaction survey", "Level Two: knowledge check (quiz)", "Level Three: follow-up survey", each with a one-line description under it.
- **Smaller header:** the tagline is gone and the icon and build box are smaller, about 20% of the pane's height instead of about 32%. The build box stays in the header, visible, because an earlier test requires the loaded build to be visible there (it is how a stale pane is caught); it was not moved to the footer or into a collapsed note.
- **Switching between Guide and Assessments clears the last screen's message.**

## As built, slice B (2026-10-31): matching and rating statements

- **Matching** (Level Two). The author types one pair to a line, the item then its match, separated by | or a tab (so a line pasted from a spreadsheet works), and optionally some extra choices that match nothing. The item is: the question with "Write the letter of the best match for each item." after it; a table with a header row (No., Item, Your match) and a "Letter:" in each match cell; "Choices:" and a real lettered list of the choices, in alphabetical order so the order of the items does not give the answers away; then the key, "Answer key (matching): 1 = b; 2 = d; 3 = a.", with the optional Why, Points and Objective. Up to ten pairs and five extra choices.
- **Likert scale** (called Rating statements until 2026-11-07; Levels One and Three). The author types a heading or instruction and one statement to a line, and picks a scale: five-point agreement, five-point success, or their own two to seven labels. The item is the heading (a prompt, with no number) and a table with a header row: Statement, then one column per rating, and a row for each statement with a real Word check box content control under each rating. Each check box is named for a screen reader ("The trainer was prepared.: Agree"). There is no key and no line to write on. Up to twenty statements.
- **Tables:** a header row marked as a header, no merged cells, no nested table, no blank cell, a caption and a description, and the columns fill the text width. Cells are in new styles (OSG Table Head, OSG Table Cell, OSG Rating Box, OSG Item Note), so the pane reads an item back from its styles like every other.
- **Editing:** Edit this item reads either type back into its form (pairs and extras recovered from the table and the key, the scale from the header row) and Update item replaces the whole item. The adapter deletes a replaced item's table as a table; deleting only its paragraphs would leave an emptied table behind.
- **Where an item goes:** after the whole item the cursor is in, from any cell. From inside a table the adapter works out where the table ends before inserting.
- **Not proven in real Word yet** (each is on the author's checklist): the table arriving with its header row marked; the check boxes arriving as clickable content controls and the checker's verdict on them; a table at the very end of an insert followed by the spare paragraph, and whether an empty paragraph is left after it; deleting a table when an item is edited; matching's "Letter:" cells and the checker's verdict on them; the checker's verdict on the whole section; save, close and reopen.

## As built, slice C (2026-10-31): four more item types

Added from the plan's own list (the rebuild plan of 2026-10-30 (kept outside the repository), "Proposed assessment item types"), from the same patterns, with no new kind of risk in Word beyond the tables and check boxes of slice B.

- **Ordering (Level Two).** The author types the items in the right order, one to a line (3 to 8, each different). In the document they are a lettered list in a fixed mixed-up order (alphabetical, or backwards if that is already the right order), with a line "Your order, by letter:". The key gives the right order by letter: "Answer key (ordering): c, b, a."
- **Performance task (Level Two).** A task, with "Your response or notes:" and lines to write on (6 by default). The checklist goes in the key, one paragraph for each criterion with optional points ("Criterion 1: Listens without interrupting (2 points)"), so the learner workbook leaves it out like any key. The author types a criterion and, after a | , its points.
- **Semantic differential (Levels One and Three).** A table with a header row: "Left word", the numbers 1 to 5 (or 7), "Right word"; each row is a pair of opposite words with a check box for each point between them, named for a screen reader ("Boring to Interesting: 2 of 5"). One to twelve pairs. No key.
- **Action plan (Levels One and Three).** For each action (1 to 5), a heading "Action 1" and four lines to write on labelled "Opportunity:", "What could get in the way:", "What I will do:" and "By when:". No table, so no blank cells; no key.
- **Reading back and editing** work for all four; each is told apart from the types that look like it (a rating table by its first column's heading, an action plan by its headings).
- **Nothing left unbuilt in this stage's list.** (The numeric type, the rubric, pasting many questions and importing from an OpenAssess bank were built on 2026-11-01: see "Numeric", "Performance task with a rubric", "Paste many questions" and "Import from an OpenAssess bank" below.)

## Import from an OpenAssess bank (built 2026-11-01)

The last item the plan listed ("linking to an OpenAssess bank"), built as an import, as the author chose. At the bottom of the Assessments screen: choose a bank file (.openassess), see what it holds, choose a folder and where to start, and add questions as items in the section the cursor is in, up to 50 at a time (the start moves on by itself, so the button is pressed again for the next 50). Each question is added once and is then an ordinary item, with no link back to the bank.

- **The file:** a zip with `manifest.json` (format `openassess-bank`, format version 1), one `questions/q_*.json` per question and `folders.json` for the folder tree, as OpenAssess's `packages/assess/src/container/layout.ts` and `model/question.ts` describe. A file that is not a bank, is damaged or is a newer format version is refused with the reason. Retired questions are left out. Questions come in the folder order the author arranged, then the rest. The guide reads the bank; it never writes to it.
- **Types:** multiple choice, multiple response, true or false, fill in the blank (one blank), short answer, matching (up to 10 pairs, up to 5 extra choices), ordering (3 to 8 items), numeric (an exact value, a tolerance, or a range, which is written as its middle plus or minus half its width, and the message says so), and essay (as a reflection: its model answer, and its rubric rows as lines in the key; a placeholder, reported, when it has no model answer). Points, learning objective and the "correct" feedback (as why it is right) come across.
- **Skipped, never guessed, and named with the reason:** a question with a picture or an equation, hotspot and calculated questions, a fill in the blank with more than one blank, a true or false with no answer given, no or too many correct options, more than 6 options, too few or too many pairs or items, and anything else the builder would refuse. They are listed above the button, for they cannot be fixed here.
- **Checked by the tests, not in Word:** every type, every refusal, the file checks, the folder order, a real bank file written by OpenAssess's own code (kept as `tests/fixtures/openassess-written-by-openassess.openassess`), the pane (the summary, a folder, the start and window, 60 questions in two goes, one-step Undo, outside a section) and the workbook with its leak check.

## Performance task with a rubric (built 2026-11-01)

A fifteenth item type beside the performance task with a checklist, which is unchanged. The form takes the levels (2 to 5, one to a line, best first, with optional points after a bar; a new rubric starts with Exceeds | 3, Meets | 2, Approaching | 1) and the criteria (one to a line: the criterion, then what each level looks like, separated by bars, in the levels' order). The learner gets the task and lines to write on. The key holds the rubric as a table: a header row (Criterion, then each level with its points) and a row for each criterion, up to 8, with a caption and description and no merged or blank cells.

- **Facilitator only, or shown to learners (the author chose "both options", 2026-11-01):** by default the table is written in two key styles (OSG Key Table Head and OSG Key Table Cell), shaded like the other keys, so the learner workbook leaves all of it out with the rest of the key. A box, "Show the rubric to learners", writes it instead in the ordinary table styles under the task (before the lines to write on), so the workbook includes it. Edit this item opens it with the box as it was and can move the table either way.
- **Checked when typed:** the number of levels, different level names, numeric points, at least one criterion, and for each line the right number of descriptions with none empty, each refusal naming the line and what to write.
- **Reads back:** Edit this item opens the levels and the criteria lines; Update item replaces the whole table once.
- **Checked by the tests, not in Word:** the exact paragraphs and styles, the table's header row and cells, every refusal, the read-back, the form (its fields, starting levels, insert, edit, replace once, an error that keeps what was typed), the workbook with its leak check, the review and the structure check.

## Numeric (built 2026-11-01)

A fourteenth item type, in the quiz group with the other types that have a key. The form shows the correct number (labelled "Correct number"), how close an answer can be and still count (optional), and the units (optional). The learner gets the question and one line to write the number on; if the key has units, the line says so ("Your answer in km per hour:"). The key reads "Answer key (numeric): 60.5 (plus or minus 0.5) km per hour", with no closing full stop so the number is never taken for one with a decimal point, followed by the usual Why, Points and Objective lines.

- **Checked when typed:** the number must be a plain number (a sign and a decimal point are fine; commas and spaces are refused with an example), the tolerance a number that is not negative, the units up to 30 characters. Each refusal says what to type.
- **Reads back:** Edit this item opens the number, tolerance and units in their own boxes; a key the author typed by hand that does not fit the shape is kept as it is.
- **Not in the paste layout:** the test-bank format has no numeric type, so a `Type: N` line is named as not read and the question is added with Add an item.
- **Checked by the tests, not in Word:** the paragraphs, the key round trip, every refusal, the form (its fields, insert, edit and replace once), the workbook (the line to write on is there, the key is not) and the review.

## Paste many questions (built 2026-11-01)

At the bottom of the Assessments screen, a box and the button "Add the pasted questions". It reads the plain-text layout test banks use (what the author's OpenAssess project calls the Standard Format; the guide reads the common part of it, checked against OpenAssess's own example files) and puts every question into the section the cursor is in, after the item the cursor is in, numbered on from the items already there (the pasted numbers are only used to match an Answers: list). One Ctrl+Z takes all of them out. At most 50 questions at a time.

- **Read:** a numbered question (`1)` or `1.`) whose text may run over several lines; lettered options (`a.` or `a)`) with a `*` before the correct one, or an `Answer:` / `ANS:` line after the options, or an `Answers:` list (to its first blank line, so questions can follow it). A `Type:` line before the number picks the type: none or MC (multiple choice), MA or MR (multiple response), TF (true or false; also found from True and False options), E (essay, kept as reflection), S (short answer), F (fill in the blank; the options and the Answers: list are the accepted answers) or MT (matching, each option written `a. Item = Match`). A `Points:` line before the number gives its points. Feedback lines (`~`, `@`, `#`), `Title:` and publisher notes (DIF, REF, OBJ, TOP, BLOOM, MSC, NAT, LOC, PTS) are skipped and counted. A `Type:`, `Points:` or `Title:` line in the middle of a question's own text, with no blank line before it and no options yet, is part of the text.
- **Refused, with the question's number and line:** no correct answer marked, two marked without Type: MA, an answer that is not one of the options, option letters that skip or do not start at a, fewer than two or more than six options, matching options with no `=`, a question with no options and no Type: line (it cannot be told from a heading), a Type: this does not read (ORD, FMB and the like: add those with Add an item), and a question number used twice when an Answers: list is keyed by it (chapters that restart at 1). The builder's own refusals come back the same way.
- **All or nothing:** if any question cannot be read, nothing is added, what was pasted stays in the box, and each problem is listed.
- **Placeholders:** an essay or short-answer question with no model answer gets "Add a model answer, or what to look for." (or "Add what a good answer includes.") in its key, and the message says so.
- **Checked by the tests, not in Word:** the layout cases above (about 24 tests), the pane (numbering on from existing items, one-step Undo, nothing added on a problem, outside a section, 50 at once), reading the pasted items back for editing, and the workbook with its leak check. The reader was also run over every example file in OpenAssess's `fixtures/standard-format` (a one-off check, not part of the suite); every question that file's own notes call readable is read, and the ones they call broken are named.
- **Not proven in Word:** these four have only run against stand-ins, like the rest of Stage 2.

## Decisions for the author

1. **The first nine item types** above, and the later ones deferred. Add or drop any?
2. **Answer keys** as a shaded "Answer key:" paragraph under each item, facilitator only, the same as facilitator notes?
3. **Question numbers** kept by Update guide ("Question 1.", restarting in each assessment section), with option letters as a real list?
4. **How learners answer.** On paper only, in Word only, or both? Suggested: both. Multiple choice and similar items get a bordered answer line to write on; rating statements get check boxes that can be clicked in Word and print as empty boxes. If it is paper only, plain boxes and lines do the job.
5. **Level sections** inserted as starters from your template (Level One survey, Level Two knowledge check, Level Three follow-up) with the "Dear Participant" note and program information lines, all editable?
