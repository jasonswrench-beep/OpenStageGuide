// The tools of OpenStage Guide, in one place: what the ribbon shows, which function each button runs, and what the side panel
// holds when that button is pressed. The manifest (scripts/build-manifest.mjs), the ribbon commands (ribbon.mjs) and the pane
// (guide-pane.mjs, app.mjs) all read this list, so a button cannot exist without its panel or the other way round.
//
// Planned with the author on 2026-11-06 and 2026-11-07 (docs/RIBBON-PLAN.md). Seven groups of three, twenty-one buttons (Word draws a
// group of more than three with mixed icon sizes; regrouped at the author's yes on 2026-11-07). A button opens
// ONE tool in the pane, not the whole guide page. "More tools" inside a panel lists only the few tools that belong with it
// (the `related` list), never all twenty-one.
//
// Where a tool's panel comes from (`kind`):
//   guide   a part of the Guide screen: `part` is the tab panel it lives in, `sections` the pieces of it that show
//   screen  a whole screen of the new pane: assess, workbook, review or help
//   legacy  a panel from the older pane that is not redesigned yet; the button opens it as it is (`view` names it).
//           None is left: Save Backup and Restore Backup, the last two, became guide tools on 2026-11-08.

export const GROUPS=Object.freeze([
  {id:'start',label:'Start',tools:['openDesigner','courseSettings','courseOutline']},
  {id:'write',label:'Write',tools:['addModule','addActivity','instructionalCues']},
  {id:'bring',label:'Bring In',tools:['activityLibrary','importSlides','updateFromSlides']},
  {id:'elearning',label:'E-Learning & Assessment',tools:['eLearningScreens','assessments','learningAlignment']},
  {id:'check',label:'Check & Plan',tools:['reviewGuide','timingAndAgenda','deliveryVersions']},
  {id:'make',label:'Make',tools:['appearance','generateMaterials','participantWorkbook']},
  {id:'protect',label:'Protect',tools:['saveBackup','restoreBackup','help']}
]);

// The menus of the standalone program (program/ui/commands.js). The ribbon above had to be seven groups of three because of Word, and the program is not
// bound by that: its menus are grouped by the work they are for, from the author's list of 2026-11-14, and a tool may be in a menu of its own
// (Slides, Assessment). Four of the tools below exist only in the program (PROGRAM_TOOLS): they open an existing panel set up for one job, or the Add module
// form's lesson half, or half of Review guide. Save and Restore a backup are in File and Help is in Help, as before.
export const PROGRAM_MENUS=Object.freeze([
  {id:'start',label:'Start',tools:['openDesigner','courseSettings','courseOutline']},
  {id:'sections',label:'Sections',tools:['addModule','learningOutcomes','addLesson','lessonDetails']},
  {id:'activities',label:'Activities',tools:['addActivity','instructionalCues','activityLibrary','forms']},
  {id:'slides',label:'Slides',tools:['importSlides','updateFromSlides']},
  {id:'assessment',label:'Assessment',tools:['assessments','learningAlignment']},
  {id:'elearning',label:'eLearning',tools:['eLearningScreens','eLearningCues']},
  {id:'check',label:'Check & Plan',tools:['reviewGuide','timingAndAgenda','accessibility']},
  {id:'make',label:'Make',tools:['deliveryVersions','appearance','generateMaterials','participantWorkbook']}
]);

// The pieces of the Guide screen that a tool can show on its own. Anything not listed for a tool is hidden while it is open.
export const GUIDE_SECTIONS=Object.freeze([
  'guide-sec-module','guide-sec-lesson','guide-sec-block','guide-sec-owncues',
  'guide-sec-import','guide-sec-update',
  'guide-sec-look','guide-sec-standard','guide-sec-built',
  'guide-sec-move','guide-sec-keeptidy','guide-sec-contents','guide-sec-library','guide-sec-agenda','guide-sec-forms','guide-sec-details','guide-sec-align','guide-sec-versions',
  'guide-sec-backup','guide-sec-restore'
]);

const tool=(id,t)=>[id,Object.freeze({id,command:id,...t,related:Object.freeze(t.related||[])})];

export const TOOLS=Object.freeze(Object.fromEntries([
  tool('openDesigner',{label:'Open Designer',icon:'OpenDesigner',group:'start',kind:'guide',part:'organize',sections:['guide-sec-keeptidy'],titlebar:true,selected:true,
    title:'Your guide',note:'The guide at a glance, and where your cursor is.',
    tip:'Open your course workspace and outline.',related:['courseSettings','courseOutline','reviewGuide','participantWorkbook','help']}),
  tool('courseSettings',{label:'Course Settings',icon:'CourseSettings',group:'start',kind:'guide',part:'look',sections:['guide-sec-standard'],
    title:'Course settings',note:'The standard pages, and where a guide starts.',
    tip:'Add the standard pages a guide opens with, and the Appendix.',related:['openDesigner','appearance','generateMaterials']}),
  tool('addModule',{label:'Add Module',icon:'AddModule',group:'write',kind:'guide',part:'add',sections:['guide-sec-module'],selected:true,
    title:'Add a module',note:'A module goes after the module your cursor is in.',
    tip:'Add a module with its title; the program numbers it and gives it a Transition block.',related:['addLesson','learningOutcomes','courseOutline']}),
  tool('addActivity',{label:'Add Activity',icon:'AddActivity',group:'write',kind:'guide',part:'add',sections:['guide-sec-block'],selected:true,preset:{cue:'activity',parts:true},
    title:'Add an activity',note:'An activity written as labelled parts.',
    tip:'Add a teaching activity with a Wrench instructional icon.',related:['addModule','instructionalCues','activityLibrary']}),
  tool('instructionalCues',{label:'Instructional Cues',icon:'InstructionalCues',group:'write',kind:'guide',part:'add',sections:['guide-sec-block','guide-sec-owncues'],selected:true,preset:{parts:false},
    title:'Instructional cues',note:'Words to say, or the points to make.',
    tip:'Insert a teaching icon and its words into the guide.',related:['addActivity','addModule','activityLibrary','importSlides']}),
  tool('activityLibrary',{label:'Activity Library',icon:'ActivityLibrary',group:'bring',kind:'guide',part:'add',sections:['guide-sec-library'],selected:true,
    title:'Activity library',note:'Ready-made activities to put into your guide, and the ones you save.',
    tip:'Put a ready-made activity into your guide, or save your own to use again.',related:['addActivity','instructionalCues','importSlides','forms']}),
  tool('courseOutline',{label:'Course Outline',icon:'CourseOutline',group:'start',kind:'guide',part:'organize',sections:['guide-sec-move','guide-sec-contents'],
    title:'Course outline',note:'Choose a line to go there. Move or delete the lesson or module your cursor is in.',
    tip:'See modules, activities and timing; jump to their text in Word.',related:['addModule','openDesigner']}),
  tool('importSlides',{label:'Import Slides',icon:'ImportSlides',group:'bring',kind:'guide',part:'slides',sections:['guide-sec-import'],
    title:'Import slides',note:'Choose the presentation. For pictures, also choose a PDF of it.',
    tip:'Put a presentation\'s slides into a lesson, each with its title, notes and picture.',related:['updateFromSlides','eLearningScreens']}),
  tool('updateFromSlides',{label:'Update from Slides',icon:'UpdateFromSlides',group:'bring',kind:'guide',part:'slides',sections:['guide-sec-update'],
    title:'Update from slides',note:'Choose the changed presentation (and a new PDF). Only differences are listed.',
    tip:'Bring changes from a presentation into slides already in the guide.',related:['importSlides','reviewGuide']}),
  tool('eLearningScreens',{label:'E-Learning Screens',icon:'ELearningScreens',group:'elearning',kind:'guide',part:'elearning',sections:[],
    title:'E-Learning screens',note:'Add a screen with its usual blocks, then make the narration script.',
    tip:'Design E-Learning screen by screen, with narration counted and timed.',related:['eLearningCues','importSlides','assessments']}),
  tool('assessments',{label:'Assessments',icon:'Assessments',group:'elearning',kind:'screen',screen:'assess',ownHeading:true,
    title:'Assessments',note:'Start a section, then add items to it.',
    tip:'Write survey and quiz items, paste questions, or take them from a question bank.',related:['learningAlignment','reviewGuide']}),
  tool('learningAlignment',{label:'Learning Alignment',icon:'LearningAlignment',group:'elearning',kind:'guide',part:'organize',sections:['guide-sec-align'],
    title:'Learning alignment',note:'Which objectives have an activity and an assessment item.',
    tip:'See which objectives have an activity and an assessment.',related:['assessments','addActivity']}),
  tool('reviewGuide',{label:'Review Guide',icon:'ReviewGuide',group:'check',kind:'screen',screen:'review',ownHeading:true,
    title:'Review guide',note:'Checks for missing text, unfinished prompts and timing.',
    tip:'Find missing text, objectives, unfinished prompts and timing issues.',related:['accessibility','learningAlignment','timingAndAgenda','openDesigner']}),
  tool('timingAndAgenda',{label:'Timing & Agenda',icon:'TimingAndAgenda',group:'check',kind:'guide',part:'organize',sections:['guide-sec-agenda'],
    title:'Timing and agenda',note:'Times for each lesson, from a start time you choose.',
    tip:'Calculate activity start and end times for your workshop.',related:['deliveryVersions','reviewGuide']}),
  tool('deliveryVersions',{label:'Delivery Versions',icon:'DeliveryVersions',group:'check',kind:'guide',part:'organize',sections:['guide-sec-versions'],
    title:'Delivery versions',note:'One-, three- or six-hour plans from the same guide.',
    tip:'Make a delivery pack from a saved plan.',related:['timingAndAgenda','participantWorkbook']}),
  tool('appearance',{label:'Appearance',icon:'Appearance',group:'make',kind:'guide',part:'look',sections:['guide-sec-look'],
    title:'Appearance',note:'The guide\'s look: fonts, colors, header and footer.',
    tip:'Choose the look of the guide: fonts, colors, header and footer.',related:['generateMaterials','courseSettings','participantWorkbook']}),
  tool('generateMaterials',{label:'Generate Materials',icon:'GenerateMaterials',group:'make',kind:'guide',part:'look',sections:['guide-sec-built'],
    title:'Generate materials',note:'Pages built from your guide when you ask.',
    tip:'Build Training at a Glance, the Graphic Cues page and a module overview.',related:['appearance','participantWorkbook']}),
  tool('participantWorkbook',{label:'Participant Workbook',icon:'ParticipantWorkbook',group:'make',kind:'screen',screen:'workbook',ownHeading:true,
    title:'Participant workbook',note:'Blocks for learners. Facilitator-only blocks are left out.',
    tip:'Preview and download learner instructions and writing space.',related:['appearance','generateMaterials','reviewGuide']}),
  tool('saveBackup',{label:'Save Backup',icon:'SaveBackup',group:'protect',kind:'guide',part:'organize',sections:['guide-sec-backup'],
    title:'Save a backup',note:'A complete copy of the guide, saved as a Word file.',
    tip:'Prepare and verify a full Word master backup with teaching assets.',related:['restoreBackup','help']}),
  tool('restoreBackup',{label:'Restore Backup',icon:'RestoreBackup',group:'protect',kind:'guide',part:'organize',sections:['guide-sec-restore'],
    title:'Restore a backup',note:'Opens a saved backup as a separate document.',
    tip:'Verify a saved backup and prepare its preserved Word master as a separate copy.',related:['saveBackup','help']}),
  tool('help',{label:'Help',icon:'Help',group:'protect',kind:'screen',screen:'help',ownHeading:true,
    title:'Help',note:'How to use the tools, and a finished example.',
    tip:'Try the sample guide, see the keyboard shortcuts, and read the instructions.',related:['openDesigner','saveBackup']})
]));

// The tools only the program has. They are not in the ribbon's list (TOOLS, TOOL_IDS and GROUPS stay the add-in's twenty-one in seven groups of three).
export const PROGRAM_TOOLS=Object.freeze(Object.fromEntries([
  tool('addLesson',{label:'Add Lesson',icon:'AddLesson',group:'sections',kind:'guide',part:'add',sections:['guide-sec-lesson'],selected:true,
    title:'Add a lesson',note:'A lesson goes after the lesson your cursor is in.',
    tip:'Add a lesson with its title and minutes; the program numbers it and gives it a time line.',related:['addModule','learningOutcomes','instructionalCues','lessonDetails']}),
  tool('learningOutcomes',{label:'Learning Outcomes',icon:'LearningOutcomes',group:'sections',kind:'guide',part:'add',sections:['guide-sec-block'],selected:true,preset:{cue:'objective',parts:false},
    title:'Learning outcomes',note:'What learners will be able to do, written as an Objective block.',
    tip:'Write what learners will be able to do after a lesson, as an Objective block.',related:['learningAlignment','addLesson','instructionalCues']}),
  tool('eLearningCues',{label:'E-Learning Cues',icon:'ELearningCues',group:'elearning',kind:'guide',part:'add',sections:['guide-sec-block','guide-sec-owncues'],selected:true,preset:{cue:'screen-title',parts:false},
    title:'E-learning cues',note:'The kinds of block for designing a course screen by screen.',
    tip:'Insert a screen title, visual direction, on-screen text, narration or another e-learning block.',related:['eLearningScreens','instructionalCues','assessments']}),
  tool('accessibility',{label:'Accessibility',icon:'Accessibility',group:'check',kind:'screen',screen:'review',ownHeading:true,
    title:'Accessibility',note:'The document checks: title, language, headings, pictures, tables, color and size of text.',
    tip:'Check the guide for what a screen reader and a reader with low vision need: heading levels, picture descriptions, table headers, color contrast and text size.',related:['reviewGuide','appearance','timingAndAgenda']}),
  tool('forms',{label:'Forms',icon:'Forms',group:'activities',kind:'guide',part:'organize',sections:['guide-sec-forms'],
    title:'Forms',note:'The handouts, worksheets and forms in your guide, how they are numbered and how many copies to make.',
    tip:'See every form in the guide with its number and print line, start a new one, choose how forms are numbered, and say how many people, teams and tables there will be to work out the copies to make.',related:['addActivity','instructionalCues','reviewGuide']}),
  tool('lessonDetails',{label:'Lesson Details',icon:'LessonDetails',group:'sections',kind:'guide',part:'organize',sections:['guide-sec-details'],selected:true,
    title:'Lesson details',note:'Purpose, materials, method, range of minutes, how people work together and whether the lesson may be left out.',
    tip:'Write down a lesson\'s purpose, materials and method, the minutes the facilitator may take, how people work together, the size of the group, and whether it is optional or one to choose from.',related:['addLesson','timingAndAgenda','reviewGuide']})
]));

export const TOOL_IDS=Object.freeze(Object.keys(TOOLS));
export const toolFor=id=>TOOLS[id]||PROGRAM_TOOLS[id]||null;

// What each tool is called in the program, in one list: in its menu, in "More tools" and as the heading of its panel. (The `label` of a tool above is the
// Word ribbon's name for it, "Open Designer" and "Add Lesson", and stays the add-in's.) The menus, the strip and the headings used to say three
// different things; the manual's author found the lists disagreeing on 2026-11-14.
export const PROGRAM_NAMES=Object.freeze({
  openDesigner:'Guide overview',
  courseSettings:'Course settings',
  courseOutline:'Course outline',
  addModule:'Add module',
  addLesson:'Add lesson',
  learningOutcomes:'Learning outcomes',
  addActivity:'Add activity',
  instructionalCues:'Instructional cues',
  activityLibrary:'Activity library',
  forms:'Forms',
  lessonDetails:'Lesson details',
  importSlides:'Import slides',
  updateFromSlides:'Update from slides',
  eLearningScreens:'E-learning screens',
  eLearningCues:'E-learning cues',
  accessibility:'Accessibility',
  assessments:'Assessments',
  learningAlignment:'Learning alignment',
  reviewGuide:'Review guide',
  timingAndAgenda:'Timing and agenda',
  deliveryVersions:'Delivery versions',
  appearance:'Appearance',
  generateMaterials:'Generate materials',
  participantWorkbook:'Participant workbook',
  saveBackup:'Save a backup',
  restoreBackup:'Restore a backup',
  help:'Help'
});
export const programName=id=>PROGRAM_NAMES[id]||toolFor(id)?.label||id;
export const isTool=id=>Object.hasOwn(TOOLS,id)||Object.hasOwn(PROGRAM_TOOLS,id);
export const groupOf=id=>GROUPS.find(g=>g.tools.includes(id))||null;

// The tools listed under "More tools" for one tool: the few that belong with it, in the order they are listed here.
export const relatedTools=id=>(toolFor(id)?.related||[]).map(r=>toolFor(r)).filter(Boolean);
