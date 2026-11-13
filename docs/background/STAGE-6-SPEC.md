# Stage 6: more cues, pictures and slides, looks, and E-Learning

Written 2026-11-06 for the author's sign-off. No code has been written for this stage. Each part says what the author
sees, what is built on what already exists, what is unproven in Word, and the decision that is the author's to make.
Each recommended answer is the one I would build if you say "accept all".

## Where this starts

- The guide has 16 cues (Objective, Say, Ask, Discussion, Activity, Key point, Lecture, Slide, Media, Training instrument,
  Workbook, Handout, Assessment, Time on task, Trainer notes, Transition). Adding a cue is one row in `src/guide-cues.mjs`
  and one picture file; the label always carries the meaning and the picture is decorative.
- Pictures: the guide only inserts its own cue pictures (decorative). Nothing inserts the author's pictures or slides.
- Looks: the facilitator guide has one built-in look. The workbook already has five (Clearline, Editorial Study, Practice
  Lab, Field Guide, Reflection Journal) with brand colour, footer and black and white, and the style code already takes a look
  as a handful of values (`styleDefinitions(theme)`).
- E-Learning: nothing. The first spec listed it as a non-goal; this stage changes that at the author's request.

## Part 1. The larger cue catalogue (about 48 cues)

**What the author sees.** The cue list in Add a block is grouped, so 48 choices stay easy to scan:

| Group | Cues |
| --- | --- |
| Teaching | Objective, Say, Ask, Discussion, Lecture, Key point, Highlight, Important, Process, Steps, Tools |
| Activities and practice | Activity, Do this, Lab, Role play, Game, Case study, Time on task, Break |
| Slides and media | Slide, Slide notes, Picture, Video, Audio, Media, Flipchart, Data tables |
| Learner materials | Workbook, Handout, Files, Training instrument, Assessment, Evaluation, Capture |
| Virtual class | Chat, Breakout, Polling, Vote, Raise hand, Microphone, Webcam, Website, Check computer, Presenter 1, Presenter 2, Producer, Q and A |
| For you | Trainer notes, Transition, and three custom cues the author names once (for example "Debrief") |

That is the 16 now plus about 32 new ones, and it covers the reference guide's block list. Module-level items (Goal,
Overview, Materials needed, References and resources) are already on the module overview page, so they are not repeated here.

**Shown to** defaults come with each cue (Everyone or Facilitator only), and the author can change them per block as now.
Virtual-class cues and Presenter / Producer cues default to Facilitator only, so they never reach the workbook. Handout,
Workbook, Files, Case study, Activity, Do this and Evaluation default to Everyone.

**Pictures for the new cues.** The author's 14 icons stay the only artwork. A new cue borrows the nearest of the 14 (as Ask
and Discussion already share one) until the author makes its own: replacing it later is one file named after the cue plus
one catalogue row. I will not draw icons in the author's style. The Graphic Cues page lists whichever cues the guide uses.

**Unproven in Word:** nothing new in kind (same blocks as now). Needs a check that no two cue labels read back as one
another, and that the pane stays usable with 48 choices (grouped list, keyboard and screen reader).

**Decision 1.** Add all of these now, borrowing icons? Recommended: yes. Alternative: only the virtual-class group now.

## Part 2. Inserting pictures and slides

**Picture block.** Add a block > Picture. The author chooses a PNG, JPG or GIF from the PC. The pane makes the author write
alt text (or tick "decorative"), offers a caption and a width (default fits the page), shrinks very large files so the
document stays small, and inserts the picture in its own paragraph under the cue line. Review's existing alt-text check
already finds any picture without alt text.

**Slide block from a presentation.** A reader of a `.pptx` file cannot draw its slides without PowerPoint, and automatic
slide drawing was deferred in the first spec. So the author does one step in PowerPoint (File > Export > PNG, all slides),
then in the pane chooses the `.pptx` and the exported pictures. The pane matches them in order and, for each slide, writes:

- a Slide cue line with the slide number and title;
- the slide's picture (alt text = the slide's title and text);
- the slide's presenter notes as a facilitator-notes paragraph (the existing slide reader already gets title, text, notes);
- "Click 1, click 2 ..." when the slide has animation steps (new, small: counts the click effects).

The author can insert all slides into a lesson at once, or pick which ones. Slide pictures default to Facilitator only,
with a "Show in the workbook" choice for handout slides. If the author has no export, the pane still inserts the slide cue
lines with notes and no picture, and says so.

**Unproven in Word:** a picture package inserted next to the spare-paragraph trick (icons already use it, so likely fine);
very large inserts; the matching of pictures to slides when file names are not in order (I would sort by the number in the
name and ask the author to confirm when counts differ).

**Decision 2.** Is "export the slides as pictures first, then pick them" acceptable? Recommended: yes. Alternative: slide cue
lines with text and notes only, and the author pastes pictures by hand.

## Part 3. Facilitator-guide looks

**What the author sees.** A Look screen with the same five names as the workbook, so a guide and its workbook can match, plus:
accent colour, heading font, body font and size, line spacing, page size, and a header and footer (guide title, copyright
line, revision date, page number: the finished reference guide has these and ours has none yet). A preview sentence shows each
look's headings and a shaded notes paragraph before it is applied.

**Two moments to choose one.**
1. On Start a guide: the look is written with the first insert (low risk: the styles do not exist yet).
2. Later, on a guide that already has content: the pane changes the OSG styles in place, so every block restyles at once.
   This uses Word's style settings (colour, font, size, spacing, shading). Not every setting is available to a pane, so this
   is the part to prove first; if a setting is refused the pane applies the rest and lists what it could not change.

The check that colours are readable (contrast 4.5 to 1) already exists for the workbook and is reused. Looks never change
the words or the structure, only styles.

**Unproven in Word:** changing existing styles from the pane; whether it is one undo step; header and footer from the pane.

**Decision 3.** Same five looks and brand settings for both guides, chosen on Start a guide and changeable later? Recommended:
yes. Icon sets per look (a look with its own artwork) wait until the author has more than one set.

## Part 4. Turning a guide into E-Learning

The first spec made E-Learning a non-goal, so this is new. The author's own tool for building E-Learning is OpenStage Author,
so the useful shape is: **design the course in the guide, then hand it to Author to build.** In four pieces, smallest first.

1. **E-Learning cues.** A group of cues for E-Learning designers: Screen title, On-screen text, Narration (the voice-over
   script), Visual direction, Interaction (what the learner clicks or drags), Knowledge check (an assessment item that
   follows the screen), Branch (if the learner chooses X go to Y), Feedback (correct or not), Animation note, Media asset,
   Accessibility note, Developer note. Narration, Visual direction, Animation note and Developer note default to Facilitator
   only so they stay out of a learner copy.
2. **Storyboard look.** A look for E-Learning: each lesson becomes a run of screens, each screen written as labelled parts
   (Screen title, Visual, On-screen text, Narration, Interaction, Notes), the way the Activity cue already writes labelled
   parts. No layout tables, so it stays accessible.
3. **Checks and exports for E-Learning professionals.**
   - Review gains an E-Learning check: every screen has narration and on-screen text, every picture has alt text, every
     interaction says what the learner does and what feedback follows, every Knowledge check points at a real item.
   - An estimated running time per screen and per lesson from the narration (about 150 words a minute, adjustable), shown
     in Review and written into the contents note.
   - A narration script export (a plain Word file, screen by screen) for the voice talent or text-to-speech, and a text
     transcript.
4. **Send to OpenStage Author.** The pane writes a file that Author opens: one lesson becomes one lesson in Author, each
   screen a page with its on-screen text, narration script, picture and alt text, and the Knowledge checks become Author quiz
   items. Author then does the building, the interactions and the SCORM or xAPI packaging it already does.
   **This needs Author's file format.** I would read Author's documented format to write the file. Anything Author itself must
   learn to open (an import) is a change in Author's own project and its own session, so I would ask you before touching it.

**Unproven:** all of it in real Word and in Author. Parts 1 to 3 of this list are pure writing and reading, as the rest of the
guide is.

**Decision 4.** Build in this order: cues and storyboard look first, then the checks and the narration script, then the hand-off
to Author? Recommended: yes. Alternative: build the hand-off to Author first (more useful sooner, but the Author format and the
Author-side import need to be settled first).

## Order and size

1. Part 1 (cues): small, mostly data. 2. Part 2 (pictures, then slides). 3. Part 3 (looks; start-time look, then change-later).
4. Part 4 (cues, storyboard look, checks and script, then the Author hand-off). Each is built, tested with my own tests, listed
in TESTING-TODO.md for what only Word can show, and then you try it. Nothing is published or pushed without your yes.

## Questions I could not settle alone

- Which of the reference list's cues you do **not** want (I would rather add one too many than lose one).
- The names of your three custom cues.
- Whether E-Learning designers who are not instructors should get a start option, "Start an E-Learning storyboard", with the
  E-Learning cues and storyboard look already chosen (recommended: yes, it is a small addition to Start a guide).

## As built (2026-11-06)

All four parts are built in the source and tested with Node (1279 tests); none has been in Word, and the hand-off has not been opened
in OpenStage Author (its own validator accepts what it writes). Where the build differs from the spec above, or the spec left a choice:

- **Part 1.** 61 cues in seven groups (the E-Learning group of part 4 is the seventh). The author's three own cues are kept in the
  document as a custom property (`osg.customCues`). Cues without a picture of their own borrow one of the fourteen.
- **Part 2.** Any block can carry a picture; a block that holds one cannot be rewritten from the pane (the pane says so and what to do
  instead). Slide pictures are matched to slides by the number in their file names. A slide's animation clicks are counted from the
  presentation file. Slides are facilitator only unless the author ticks the workbook box.
- **Part 3.** Looks are the workbook's five families plus a sixth, Storyboard. They change the guide's styles through Word's style
  settings (never tried in Word), and styles not made yet are made from the look by the next insert. **The Start screen offers
  "Standard styles" first** until the style settings are proven; then the default can become Clearline. Page size and margins are not part of a look.
- **Part 4.** Branch is facilitator only by default (the spec left it open). Screens are read from the blocks, not stored. The narration
  is timed at 150 words a minute (a constant, not yet a setting). Send to OpenStage Author writes one scene per lesson, one slide per
  screen with a heading, the on-screen text and the narration as the voice-over script; everything else the storyboard says goes to a
  notes file. **Not built:** turning pictures, interactions, branches and Knowledge checks into Author objects (a Knowledge check is only
  named in the notes), and any change inside OpenStage Author itself.

### Added after the first native test (2026-11-06)

Word's task pane never opens a file chooser (`<input type="file">` does nothing there), so choosing a picture, or a presentation and its
slide pictures, happens in a small window of its own (Office's dialog API, the page `picker.html`, `src/file-pick.mjs`): it reads and
shrinks the files and sends them to the pane as text messages. The pane's picture and slide buttons open that window. This has not been
tried in Word yet. The same problem affects the older question-bank file import. A list of the new icons to create is in
`docs/ICONS-TO-CREATE.md`.
