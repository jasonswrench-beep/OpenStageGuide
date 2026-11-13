# Rebuild plan: an accessible LeaderGuide Pro for Word

Written 2026-10-30. This is a plan only: nothing in it is built and no code changed. Decisions that are the author's to make are in section 8, each with a recommendation.

**How to read the evidence.** *Confirmed* = seen in a recording, a saved file, the code or a test run. *Judgment* = my reading, not tested. *Native* = ran in the author's real Word. *Automated* = ran against made-up test documents.

## 1. The goal (the author's words, 2026-10-30)

- Something like LeaderGuide Pro: guides built from modules, lessons and content blocks, with icons and timing.
- Functional and accessible, and it looks nice and professional.
- Makes the facilitator's guide and the learner workbooks from the same course.
- Standardized like the guide used in class, with more it can do.
- An assessment section where you click the type of item you want and insert it quickly.
- The workbook families are deliberate look options.

LeaderGuide Pro itself, from the public Great Circle Learning walkthrough page and the supplied teardown: a cover and acknowledgements, a table of contents, modules of lessons built from content blocks, an optional appendix, and "Collections" that set fonts, icons and colors. The participant guide is extracted from the facilitator guide. Neither source describes accessibility, which is where this product differs.

### Standing constraints (from the original request that opened this work)

- Preserve source754's features.
- No broad cache clearing, speculative registry edits, certificate or trust changes, or installations without asking first.
- Marketplace acceptance is never a prerequisite for a private test.
- Separate confirmed findings from hypotheses, and native Word results from automated tests.

### Original design intent (the first Claude conversation and spec, read 2026-10-24)

The first spec (working title GuideSmith, researched from Great Circle Learning's public pages) asked for:

- An Office web add-in rather than a Windows-only plug-in, producing ordinary `.docx` with real Word styles so the file opens correctly without the add-in.
- LeaderGuide Pro's structure: Module, Lesson, Content Block. Each block is a table row with an icon column, a main column and a notes column. Facilitator default is three columns and participant default is two, changeable per block.
- The standard page inventory: cover, acknowledgements, contents, a Getting Started set (About This Guide, Graphic Cues, The Program in Perspective, Program Preparation, Training at a Glance), module overview pages, lesson tables that open with a time block showing cumulative time, a Transition block ending each module, and an appendix. The participant guide is leaner and adds Purpose and Learning Objectives.
- Collections: the same style and icon names with different looks, so one click restyles a guide. This is the author's "workbook families".
- Structure protected by locked content controls and a "Show structure" toggle, instead of the vendor's red warning text.
- A first-run guide with a sample project.

Revision 2 (2026-10-24) later made a single-column narrative the default layout and added delivery versions, slide import with reviewed updates, and recovery. This plan keeps the original structure (module, lesson, block, icons, timing) and keeps those additions parked. On 2026-10-30 the author decided that neither guide should use the three-column layout table, so that both stay accessible (see D1).

## 2. Where things stand

Confirmed:
- The accepted checkpoint (source754) passes 754 automated tests. Its README says native authoring load "remains unreliable".
- First real course creation (screen recording, 2026-10-30): the Wrench starter refused a brand-new blank document with a "requires a new blank DOCX" message. The document stayed untouched, so nothing was lost.
- A saved copy of that document passes the add-in's own blank check. The failure happens in Word's live capture and its cause is still unknown. A diagnostic build (branch `diagnose-blank-check`) was made to name the failing step.
- Earlier the same day a "managed field was deleted" error paused every operation in another document.
- The pane is one 149 KB file (`src/app.mjs`) driving every panel. The Course tab alone holds the starter choices, five detail fields and the create button, and the button sits below the fold.

Judgment:
- The "managed field was deleted" pause is the safety machinery working as designed. The blank-document refusal is that same machinery rejecting a document that really was blank. Either way the design gives the author errors instead of a guide: every field is a locked content control checked against a saved record, with journals, receipts and a pause-on-mismatch rule. The author does not need it to meet the goals above.
- The feature list is mostly intended. The drift is in the engine and the screens, not the scope.

## 3. Principles

1. **Insert, don't manage.** Each action puts real, visible content into the Word document. The document is the source of truth, not a hidden record that Word text must match.
2. **Word's own Undo works.** No journals, receipts or pause state. Structure may keep Word's "cannot be deleted by accident" protection, as the original spec proposed, but the add-in never stops working because of a mismatch: it reads what is in the document.
3. **Audience is a style.** Facilitator-only and learner content are marked with real Word paragraph styles (visible, survive copy and paste). Content controls are optional wrappers, never locks. *(Decision D2.)*
4. **Accessible by construction**, following the spec's ACCESS01 to ACCESS03: native heading hierarchy, real lists, labeled prompts, no tables used for layout (real data tables only, each with a header row), document title and language, icons beside text labels so the same instruction is not announced twice.
5. **Never lose or leak content.** Participant output excludes answer keys by construction, not by hiding them (spec LEARN02, A03).
6. **Each stage is proven in the author's real Word before the next starts.** Automated tests stay, but a native result is never claimed from an automated one.
7. **Change control** (from the spec): a new feature enters a stage only if it is needed to pass that stage's stated test.
8. **Preserve source754.** Nothing is deleted. `main` and `shared-folder-install` keep the accepted build, and new work happens on `rebuild/core`. A source754 feature is retired from the screen only when its replacement works or the author signs it off explicitly.

## 4. What stays, what is replaced, what is parked

| Layer | Files | Decision |
| --- | --- | --- |
| Pane screens | `app.mjs`, `*-panel.mjs`, `taskpane.html`, `style.css` | Replace with a short, task-led pane (Build, Assessments, Look, Output). |
| State and integrity | `document-store`, `checkpoint`, `cue-ownership`, `frontmatter-package`, `icon-controller`, `model`, `word` | Replace with the insert-based model. Retire journal, receipt and pause-on-mismatch for new guides; keep deletion protection on structure. |
| Output generators | `facilitator-output`, `delivery-pack`, `materials-output`, `workbook-family-*`, `workbook-presentation`, `workbook-logo`, brand presets | Keep. Adapt to read the new guide. |
| Assessment | `assessment-output`, `-scoring`, `-file`, `-starters` | Keep the Likert and semantic-differential logic as two item types. Replace the 43 KB panel with a click-to-insert gallery. |
| Icons, licences | `public/assets`, `LICENSE`, `NOTICE` | Keep. |
| Parked (code stays, hidden) | delivery plans, objective coverage, PowerPoint import, backup and restore, research section, measurement fields | Stay reachable through the Legacy pane and are never deleted. They return to the new screens by change control, or are retired only with the author's explicit sign-off, feature by feature. |

## 5. The small content model

Guide, then Module, then Lesson, then Block. A block has a **type** (the fourteen icons: Objective, Key point, Lecturette, Slide, Media, Training instrument, Learning activity, Participant workbook, Handout, Discussion, Time on task, Trainer notes, Transition, Assessment), an **audience** (facilitator, participant or both) and optional **minutes**. Lesson times roll up to module and guide totals. Cue labels (Ask, Explain, Discuss, Debrief, Play audio, Capture, Game instructions, Highlight, Key points) are text labels mapped to the existing icons, as in the spec's ICON01. The supplied sample guide (see Reference) uses about 47 cues, so the cue set is a catalog that grows, not a fixed fourteen. A lesson opens with a time block showing cumulative time, and each module ends with a Transition block. "Start a guide" creates the standard pages listed in section 1; the participant guide is generated from the facilitator guide's participant content. Each block is laid out as ordinary paragraphs, not as a table: a cue label with its icon, then the content, then any facilitator notes as a shaded paragraph style. Tables appear only where the content is really tabular (Training at a Glance, matching items, rubrics), each with a header row. A learning-activity block is a small labelled set (objective, materials, time, preparation, instructions), as in the author's own classroom guide, and each module opens with a "content to be covered" list with minutes, which is a real data table with a header row.

## 6. Stages and the real-Word test for each

| Stage | After it you can | Done when (checked in real Word) |
| --- | --- | --- |
| 0. Steady ground | Trust the install again. | Server back on the accepted build; new branch `rebuild/core` from `shared-folder-install`; fixtures and a native test checklist written; which Word API versions each insert needs confirmed on your Word; one sample lesson built three ways from the same content: cue-label paragraphs with shaded notes (the proposed default), plain narrative, and, for comparison only, a three-column table. I check each one's structure; you run Word's checker on each; a screen-reader read-through happens when you are away. |
| 1. Start a guide | Start a guide in a blank or an existing document with the standard pages, add modules, lessons and blocks with icons and time. | Works on a blank and on an already-written document. Navigation pane shows the heading hierarchy. Add module, lesson or block in three clicks or fewer, with a keyboard route. Ctrl+Z undoes each insert. Save, close, reopen and the pane picks the guide up. Word's accessibility checker shows no errors. You finish the first run without reading instructions. |
| 2. Assessment gallery | Click an item type and insert it. | Each type inserts in one click plus a short form. Learner version and facilitator key are separate. Result passes the checker. Survives save and reopen. |
| 3. Learner workbook and looks | Generate the participant workbook in a chosen look. | The workbook opens without the add-in. I inspect the saved file and find no facilitator text or keys in the body, comments, hidden content or custom XML. The look and brand apply. |
| 4. Accessibility and print | Review and export with confidence. | A review list separates errors, warnings and human checks. Screen-reader pass (only when you are away, as agreed). Grayscale and print check. A reviewed PDF export route. |
| 5. Distribution | Give it to others. | Public hosting on GitHub, a second manifest, an install route and a privacy statement, per `DEPLOYMENT-OPTIONS-2026-10-30.md`. |

### Proposed assessment item types (you edit this list)

- **Check knowledge (Level Two):** multiple choice (one answer), multiple response, true/false, matching, ordering, fill in the blank, short answer, essay or reflection prompt, performance task with checklist or rubric.
- **Surveys and follow-up (Levels One and Three):** rating scale (Likert), semantic differential, open comment, action plan with opportunity and barriers.

Each item inserts a learner version with writing space and a separate facilitator key, rationale and scoring note, plus an optional link to an objective. The spec forbids answer-key leakage by hiding text, so keys live only in facilitator-styled content.

## 7. First-run requirements (the "intuitive" problem)

Seen in the recording and treated as requirements: one clear next action per screen; the main button visible without scrolling; plain-language messages that say what to do next; no dead-end state that only a new document can fix; no browser autofill popups on text fields; a built-in sample guide; the icon legend generated from the icons actually used.

## 8. Decisions for the author

| ID | Decision | Recommendation |
| --- | --- | --- |
| D1 | Layout default: LeaderGuide-style three zones (icon, content, notes; the original spec) or a single-column narrative with icon and label (revision 2, DOC01). | **Decided by the author, 2026-10-30: no three-column layout table in either guide, so both stay accessible.** (The author first chose three zones, then set that aside after the accessibility discussion.) Default: each block is a cue label with its icon, then the content, then facilitator notes as a shaded paragraph style. It keeps the LeaderGuide feel and reads in a straight line. The plain narrative stays as an alternative. Real data tables (Training at a Glance, matching, rubrics) are still allowed, with header rows. The icon is marked decorative and the text label always carries the meaning. Stage 0 compares the options before the default is fixed in code. |
| D2 | Mark audience with styles or with locked content controls. | Styles. |
| D3 | The assessment item types in section 6. | Confirm or edit the list. Check first whether the separate OpenStage-Assess project already defines item types to reuse. |
| D4 | Keep the old pane as "Legacy". The original request was to preserve source754's features. | Keep it available. Retire a feature only when its replacement works or you sign it off explicitly. |
| D5 | Where the work lives. | New branch `rebuild/core`. `source754` stays the accepted checkpoint. The repo goes public only after a scan for machine names and local paths. |
| D6 | First platform. | Windows desktop Word. Mac and web later, per the spec. |
| D7 | How native tests run. | You run a short checklist and send a screen recording or a Save As copy. I read the files. |

## 9. Risks

- Office.js may not support every insert I want on every Word build. Stage 0 checks this on your Word before design is fixed.
- Real Word documents differ from test documents. The blank-check failure is an example. Stage gates exist for this reason.
- I cannot see your Word directly. All native evidence comes from your recordings and saved files.
- The cue-label layout leans on hanging indents, shading and a decorative icon. Stage 0 checks how Word's checker and a screen reader treat it, against plain narrative and a three-column table, before it is fixed as the default.
- Form-style learner prompts may read differently in different screen readers. Stage 4 tests this rather than assuming.

## 10. Next steps after approval

1. Switch the server back to the accepted build and record where each branch stands.
2. A one-page screen specification for Stage 1 (three mock screens: Start a guide, Add, Review) for your sign-off before any code.
3. Build Stage 1 on `rebuild/core` and test it in your Word.

## Reference: what the LeaderGuide Pro recording shows (supplied 2026-10-30)

The recording is about 2.7 minutes of the product's Word ribbon with the product deactivated, so it shows menus and tooltips only, not finished pages. I have not heard its audio. What the menus show:

- **Module/Lesson Control:** Module and Lesson insert buttons; a menu to start a guidebook from a design template, insert archived modules or lessons, delete, reorder, archive, add an appendix section, and "About this Guidebook".
- **Go To:** jump to the table of contents or Training at a Glance, next module, next lesson, and content-block menus (instructor-led, virtual, special tables, custom, numbered step blocks).
- **Content blocks:** an instructor-led list and a virtual instructor-led list (cues such as Breakout, Chat, Do This, Microphone, Polling, Presenter 1 and 2, Producer, Raise Hand, Upload Files, Vote, Webcam, Website, plus three custom slots), and "Content Block Bundles".
- **Special:** PowerPoint slide images, picture blocks, show or hide icons, "Sync Guidebook Blocks (Participant and Facilitator)", keyword tags, and "Remove Knowledge Check Answers". The last two confirm the same concerns this plan has: linking the two guides, and keeping answers out of learner copies.
- **Text and table formatting:** a style gallery, paste as plain text, an answer entry field, numbered lists, a title style, and a generated Training at a Glance table.
- **Insert:** capture lines for notes and a task pane for graphic cues that includes the time marker for content blocks.

The supplied sample guide is for a virtual class and uses these virtual cues, so the catalog needs them. Open question for the author: which cues ship in the first release. The fourteen icons do not include the virtual ones, and the author has said more icons can be made.

## Reference: the supplied sample facilitator guide (Great Circle Learning, 2026)

A 48-page guide built with LeaderGuide Pro and exported to PDF, so it shows the finished product, not the Word source. I read its text and its tags locally with a script; nothing was sent anywhere. My tools could not draw its pages, so layout details come from text positions and tags, not from seeing them.

- **Parts, in order:** acknowledgements; table of contents (introductory materials, then modules; roman numerals for front matter); Getting Started (About This Guide, Graphic Cues, The Program in Perspective, Program Preparation, Training at a Glance); then the modules. A one-page summary opens each module: goal, time, number of lessons, overview, materials needed, references and resources. Each lesson has its title, cumulative time and lesson time, then its blocks. A running header and footer carry the guide title, copyright, revision date and page number.
- **Block layout:** the guide describes a block as a margin icon, a title line and the text. In the PDF each lesson page is one tagged table: icon or slide picture on the left, facilitator text in the middle, notes on the right, with tagged header cells.
- **Cue catalog (about 47):** module blocks are Goal, Time, Overview, Materials Needed, References and Resources, Break. Lesson blocks are Activity Instructions, Audio, Breakout, Capture, Case Study Instructions, Chat, Check Computer, Data Tables, Do This, Evaluation Instructions, Files, Flipchart, Game Instructions, Handout, Highlight, Important, Key Points, Lab, Microphone, Objective, Picture, Polling, PPT Notes, PPT Slide, Presenter 1 and 2, Process, Producer, QA, Raise Hand, Role Play, Say This, Steps, Time, Tools, Transition, Video, Vote, Webcam, Website, Workbook.
- **Text patterns:** "Say Something Like", "Ask" with "Listen For", "Note", "Time Check" (what to shorten when running long), "Screen Share", a poll with its launch instruction, "Answer in Chat", "Transition To", best-practice and warning boxes, slide thumbnails with animation click counts, and references to participant-guide pages.
- **Accessibility of that export (this PDF only):** it is tagged, with a language and a title, and its tables have header cells. Rows with one cell beside rows with three suggest merged cells. 131 of 219 tagged pictures carry alt text. Only four headings are tagged, so the Navigation pane would have little to show. These are facts about one PDF export, not about the product's Word files.

## Reference: the author's own classroom guide (supplied 2026-10-30)

Appendix B, "Conflict Management Skills Training: Facilitator Guide", from the author's textbook. It is a picture of what the author's guides look like today, not a specification: the author wants the product to do more than these documents do. The *Facilitator's Guide Template - 2023* supplied the same day is the same file as the copy already read.

- **Shape:** each module opens with its purpose, learning objectives and a "content to be covered" list with minutes (a table with an icon column, a content column and a minutes column), then key points, discussion questions and an activity list that names the materials each activity needs. Each activity is a set of labelled parts: objective, materials, time, preparation, instructions, with any case study after.
- **What my structure check finds in it (not confirmed in Word):** no real headings, because its H1 and H2 styles may lack outline levels; no header row on its table; no alt text on 10 of 19 pictures; 78 paragraphs that are typed bullets or numbers.

## Sources

Supplied documents: *Revised Product Spec for Training Guide Design* (revision 2, 2026-10-24), *Product Spec: Word Add-in for Learner and Leader Guide Design*, *Facilitator's Guide Template - 2023*, the Classroom Test participant workbook. Public page: the Great Circle Learning LeaderGuide Pro participant guide walkthrough. A screen recording of the LeaderGuide Pro ribbon and a 48-page sample facilitator guide built with it (both supplied 2026-10-30). This repository's code and tests. The 2026-10-30 screen recording and saved test document. The original Claude conversation that produced the first spec, pasted by the author on 2026-10-30. A shared ChatGPT link the author sent the same day could not be read (HTTP 403), so nothing from it is used.
