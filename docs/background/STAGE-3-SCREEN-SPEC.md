# Stage 3 screen spec: the learner workbook and its looks

Draft 2026-10-31, for the author's sign-off before any code is written. It follows the rebuild plan of 2026-10-30 (kept outside the repository) (Stage 3) and reuses the Stage 1 and 2 machinery: audience is a style, the document is the source of truth, and each step is checked in real Word.

## What was asked

"Makes the facilitator's guide and the learner workbooks from the same course." The plan's test for this stage: the workbook opens without the add-in; the saved file holds no facilitator text or keys in the body, comments, hidden content or custom XML; the look and brand apply.

## What it is

A **separate Word file** made from the guide: the learner-visible part, in a look the author chooses. It is a one-way copy. The guide stays the master. Change the guide, make the workbook again; edits made in a workbook do not flow back. Nothing in the guide is changed by making one.

## How the guide marks who sees what (already built)

- Every block has an audience, chosen when it is added (Everyone or Facilitator only). The facilitator-only cues are **Say**, **Trainer notes** and **Transition** by default; the other thirteen are for everyone. The audience is a paragraph style (OSG Cue Facilitator or OSG Cue), not hidden text.
- **Facilitator notes** under any block are one style (OSG Facilitator Notes).
- **Answer keys** are one style (OSG Answer Key).
- A cue's extra lines are in a style (OSG Cue Body) that does not say who they are for, so the workbook decides it from the cue above them. A Say block's second line goes with the Say block.

## How the workbook is made: kept, not hidden

The workbook is **built from the paragraphs that are kept**, never from a copy with things deleted or hidden afterwards. Anything not on the list below is not in the file at all.

| Kept | Left out |
| --- | --- |
| The guide title, the module, lesson and assessment-section headings | Facilitator-only blocks and their extra lines (Say, Trainer notes, Transition, and any block set to "Facilitator only") |
| Blocks for everyone, with their icons and extra lines | Facilitator notes |
| Everything inside an assessment section except keys: the note, the information lines, questions, options, answer lines, matching and rating tables with their check boxes | Answer keys |
| | Time blocks (the minutes and cumulative time) |
| | The Getting Started pages and the facilitator's contents table |
| | Your own typed paragraphs outside an assessment section (listed for you, see decision 1) |

The workbook then **adds**: a title page ("Participant workbook" in place of "Facilitator guide", and a Name line), a fresh contents table, and a Learning objectives page collected from the Objective blocks (see decision 2).

## The leak check

Run on the finished file before it is offered, and again by me on the file you save:

1. No paragraph in a facilitator style, and none of the style definitions for them.
2. Every facilitator paragraph's words, searched for in all of the file's text (body, header, footer, notes), are absent.
3. No comments, tracked changes, hidden text, footnotes or endnotes, custom XML parts, or content controls other than the rating check boxes.
4. The document properties (title, author, subject, keywords, comments) hold no guide text beyond the workbook's title.
5. Only the expected parts are in the file. Anything else makes it fail.

If a check fails, the workbook is **not offered** and the pane says which check failed. It is a checker I also run on its own, against the file you save, so the result does not rest on the code that made it.

## The screen

A third screen beside Guide and Assessments, reached from the same switch.

```
Workbook                                         [ Guide | Assessments | Workbook ]

From your guide
2 modules, 3 lessons, 11 blocks for learners, 6 assessment items.
Left out for learners: 5 facilitator-only blocks, 4 sets of facilitator notes,
6 answer keys, 3 time blocks, 2 paragraphs of your own text.      [ Show what was left out ]

Look
( ) Clearline            spacious single column
( ) Editorial Study      reference style, serif headings
( ) Practice Lab         numbered responses
( ) Field Guide          step-by-step
( ) Reflection Journal   prompt and response

Your name or organization   [__________]      Accent colour [#164B6B]
( ) In colour   ( ) Black and white

Writing lines after each question or activity   [ 3 ]   (0 for none)

[ Make the workbook ]
```

- Pressing **Make the workbook** opens it in a **new Word window**. Save it where you want; the pane suggests the name "<guide title> - Participant workbook".
- A line under the button says what was checked: "Leak check passed: 5 facilitator-only blocks, 4 sets of notes and 6 answer keys are not in the file."
- **Show what was left out** lists your own text that was left out, each with a Go there link, so nothing leaves the workbook without you being able to see it.
- If Word cannot open it for you, the same file is offered as a download.

## The looks

The five looks already in the product are kept as the choices: Clearline, Editorial Study, Practice Lab, Field Guide and Reflection Journal. Each sets the fonts, the accent colour, the sizes, line spacing and page margins. They already refuse a colour pair with text contrast below 4.5 to 1, an unapproved font, a body size below 11 points, and sizes or margins outside a supported range, and that check is reused.

- A look is applied by **writing the guide's own styles again** in the workbook (headings, blocks, questions, tables), so every part follows it, and so a person who edits the workbook in Word keeps working with real styles.
- **Brand:** organization name, accent colour, colour or black-and-white. In black and white, the accent becomes dark grey and nothing is explained by colour alone. A footer carries the organization name, the title and "Page X of Y".
- Page size Letter or A4, from the same settings. A logo is **later** (the old code can check one; it is not in this stage).

## What I cannot prove without Word (each is on the checklist)

- That Word opens the file in a new window with no "repair" message, and that the workbook's styles, icons, numbered lettered lists and check boxes arrive intact.
- That the contents table shows its entries (it may need F9 in the new window).
- That each look and the black-and-white mode look right, print as intended and keep contrast.
- That Word's Check Accessibility has no errors on the workbook.
- That the saved file passes the leak check, run by me on your Save As copy.

## Slices

- **3a: the workbook.** Kept and left out, the added title page and contents, the leak check, the Workbook screen with the summary, opening in a new window. One look (Clearline), no brand fields.
- **3b: the looks.** The other four looks, brand fields, the footer, black and white, writing lines.

Run 3a in Word first, so a problem with opening the file or the leak check cannot hide a problem with the looks.

## Not in this stage

A logo. Cover pages beyond the title page. Changing the facilitator guide's own look (it keeps its navy and its icons). PDF and print export (Stage 4). Taking edits back from a workbook into the guide. The old "family workbook" and "delivery pack" generators, which stay in Legacy tools, untouched.

## Sign-off (2026-10-31)

The author accepted the suggested answer to each of the five decisions below.

## As built, slice 3a (2026-10-31)

- **What is kept** is decided by one function (`classify` in `src/guide-workbook.mjs`), once for the pane's summary and again for the file, so the two cannot disagree. Blocks for everyone and their extra lines are kept; a facilitator-only block takes its extra lines with it; notes, keys, time blocks, the front pages and the author's own typed paragraphs are left out. An assessment section is kept except its keys, and a heading ends a section.
- **The file** has a title page ("Participant workbook" and a Name line), a contents table whose entries are already written (it shows them without F9, and F9 adds page numbers), and a Learning objectives page from the Objective blocks. Each module and each assessment starts a new page. The file holds only the styles it uses, its own numbering for the lettered lists and the bullets, and the icons it needs.
- **Also taken out of the kept paragraphs:** a block's minutes ("(5 min)", however the runs are split), tracked changes (insertions and moves are accepted, deletions dropped) and formatting revisions, comments, hidden text (including text hidden by a character style), links, notes, fields (flattened to their results), text boxes and floating pictures, and the author's own content controls. The rating check boxes and the icons stay. All of this is done to a copy before anything is read from it, so hidden words cannot reach the title page, the contents or the learning objectives.
- **A tightening of decision 1, from an independent review (2026-10-31).** An assessment section's note and information lines are kept (plain paragraphs straight after its heading, up to eight), and so are its items. Text typed among the items or after the last one is now left out and listed, like the rest of your own text; it was kept at first, and the review showed a facilitator's remark typed there would have gone into the workbook with the check passing. A paragraph in a style the pane does not use (a pasted key whose style Word renamed, say) is left out wherever it is.
- **A table is kept whole or not at all.** Headings alone are not enough to make a workbook; it needs a block for everyone or an assessment.
- **The leak check** (`src/workbook-leak.mjs`) checks the finished file's parts, relationships, styles, hidden content (revisions, comments, notes, text boxes, alternate content, links, fields, content controls), properties and words. It searches for each facilitator paragraph, and for the author's own left-out text, three ways: whole, without a block's minutes, and without its label ("Facilitator notes:", "Answer key (...):", a cue's name; this only when at least twelve letters remain), in the text and in the words that describe pictures and tables. A facilitator paragraph's words count as a leak only when the file holds them more often than the guide's own kept content does, so a learner block that happens to repeat a facilitator line is not refused, but a facilitator paragraph that has been added is. Words under eight letters are matched as a whole paragraph only. The same check runs on a saved file with `node scripts/workbook-leak-check.mjs "workbook.docx" --source "guide.docx"`, which is how I check the copy you save; for a copy Word has saved again it accepts the parts and properties Word adds (its app properties, theme, font table, the creator, and the contents entries' links and page-number fields) and still searches them.
- **An independent review** (a second agent, given the code and asked to break it) found nine ways in or around the check, and each is now a test: unrecognised text kept in a section, hidden words reaching the generated pages, text boxes and revisions copied through, a minutes marker surviving a split run, an Objective field copied from an Objective block being refused, one shared word being refused, fields and moved text mishandled, and a saved copy that could never pass.
- **The screen:** a third button, Workbook. It shows what is for learners and what is left out, lists your own text that was left out with a Go there link for each, and has one button. The message about the result sits above the button.
- **Look:** Clearline only, applied by writing the styles again from the look's values (fonts, sizes, accent). No brand fields, footer or writing lines yet: those are slice 3b.
- **Not proven in real Word:** `createDocument` opening the file in a new window (a download is the fallback); the file opening with no repair message; the styles, icons, lettered lists, bullets and check boxes arriving intact; Aptos being installed (Word substitutes another font if it is not); the contents entries and F9; Word's Check Accessibility; and my leak check on your saved copy.

## As built, slice 3b (2026-10-31)

- **Looks:** the five existing looks (Clearline, Editorial Study, Practice Lab, Field Guide, Reflection Journal) are offered as a list on the Workbook screen, Clearline first. Each writes the workbook's styles with its own heading and body fonts, sizes, line spacing and margins, and its accent.
- **Brand:** an organization name (in the footer), an accent colour (typed as #RRGGBB, with or without the #), In colour or Black and white, and Letter or A4. The existing contrast rules still apply, and the accent also has to read as heading text on the page (at least 4.5 to 1), or the workbook is refused with the reason.
- **Black and white:** the accent becomes dark grey, and every colour in the kept content (the tables' header shading and borders, a line's border) becomes the grey of the same lightness. A typed-in accent is not used. The icons are still coloured pictures.
- **Footer:** the organization and the title on the left, "Page X of Y" on the right as real page-number fields, on every page but the first. The leak check reads the footer's words too, and allows exactly one footer part.
- **Lines to write on:** after an Ask, Discussion, Activity or Workbook block (and its extra lines), "Your notes:" and as many bordered lines as chosen in all, 0 to 6, 3 by default. They use the same style as the other lines to write on, so they are not counted as empty paragraphs. How far apart they are comes from the look's response-line setting.
- **Remembered:** the look and the brand choices are kept in the browser's local storage, a convenience only. If it is missing, refused or damaged, the pane starts on the first look.
- **Not proven in real Word:** each look on screen and in print; Aptos being installed; the footer's page numbers updating; black and white in print preview; whether the line spacing is comfortable to write on.

## Decisions for the author

1. **What counts as for learners.** The allow-list above: blocks for everyone, all of an assessment section except its keys. Your own typed paragraphs outside an assessment section are left out and listed, because the pane cannot tell whether you meant them for learners. Suggested: yes, with the list. (The alternative is to keep them, which risks a facilitator remark reaching learners.)
2. **What the workbook adds.** A title page with a Name line, a new contents table, and a Learning objectives page collected from the Objective blocks. Time blocks and the Getting Started pages are left out. Suggested: yes to all.
3. **Writing space.** After Ask, Discussion, Activity and Workbook blocks, a labelled "Your notes:" with bordered lines (not typed underscores), 3 by default, 0 to 6 in the pane. Suggested: yes.
4. **The looks and brand.** The five existing looks, with organization name, accent colour, colour or black and white, a footer with page numbers, Letter or A4. Logo later. Suggested: yes.
5. **How it is delivered.** It opens in a new Word window, and you save it; it is a one-way copy you make again whenever the guide changes. A download if Word cannot open it. Suggested: yes.
