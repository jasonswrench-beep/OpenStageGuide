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

// The pieces of the Guide screen that a tool can show on its own. Anything not listed for a tool is hidden while it is open.
export const GUIDE_SECTIONS=Object.freeze([
  'guide-sec-module','guide-sec-block','guide-sec-owncues',
  'guide-sec-import','guide-sec-update',
  'guide-sec-look','guide-sec-standard','guide-sec-built',
  'guide-sec-move','guide-sec-keeptidy','guide-sec-contents','guide-sec-library','guide-sec-agenda','guide-sec-align','guide-sec-versions',
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
    title:'Add a module or lesson',note:'A module goes after the module your cursor is in; a lesson after the lesson.',
    tip:'Create a module with a learning objective.',related:['addActivity','instructionalCues','courseOutline']}),
  tool('addActivity',{label:'Add Activity',icon:'AddActivity',group:'write',kind:'guide',part:'add',sections:['guide-sec-block'],selected:true,preset:{cue:'activity',parts:true},
    title:'Add an activity',note:'An activity written as labelled parts.',
    tip:'Add a teaching activity with a Wrench instructional icon.',related:['addModule','instructionalCues','activityLibrary']}),
  tool('instructionalCues',{label:'Instructional Cues',icon:'InstructionalCues',group:'write',kind:'guide',part:'add',sections:['guide-sec-block','guide-sec-owncues'],selected:true,preset:{parts:false},
    title:'Instructional cues',note:'Words to say, or the points to make.',
    tip:'Insert a teaching icon and its words into the guide.',related:['addActivity','addModule','activityLibrary','importSlides']}),
  tool('activityLibrary',{label:'Activity Library',icon:'ActivityLibrary',group:'bring',kind:'guide',part:'add',sections:['guide-sec-library'],selected:true,
    title:'Activity library',note:'Ready-made activities to put into your guide, and the ones you save.',
    tip:'Put a ready-made activity into your guide, or save your own to use again.',related:['addActivity','instructionalCues','importSlides']}),
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
    tip:'Design E-Learning screen by screen, with narration counted and timed.',related:['importSlides','assessments']}),
  tool('assessments',{label:'Assessments',icon:'Assessments',group:'elearning',kind:'screen',screen:'assess',ownHeading:true,
    title:'Assessments',note:'Start a section, then add items to it.',
    tip:'Write survey and quiz items, paste questions, or take them from a question bank.',related:['learningAlignment','reviewGuide']}),
  tool('learningAlignment',{label:'Learning Alignment',icon:'LearningAlignment',group:'elearning',kind:'guide',part:'organize',sections:['guide-sec-align'],
    title:'Learning alignment',note:'Which objectives have an activity and an assessment item.',
    tip:'See which objectives have an activity and an assessment.',related:['assessments','addActivity']}),
  tool('reviewGuide',{label:'Review Guide',icon:'ReviewGuide',group:'check',kind:'screen',screen:'review',ownHeading:true,
    title:'Review guide',note:'Checks for missing text, unfinished prompts and timing.',
    tip:'Find missing text, objectives, unfinished prompts and timing issues.',related:['learningAlignment','timingAndAgenda','openDesigner']}),
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

export const TOOL_IDS=Object.freeze(Object.keys(TOOLS));
export const toolFor=id=>TOOLS[id]||null;
export const isTool=id=>Object.hasOwn(TOOLS,id);
export const groupOf=id=>GROUPS.find(g=>g.tools.includes(id))||null;

// The tools listed under "More tools" for one tool: the few that belong with it, in the order they are listed here.
export const relatedTools=id=>(TOOLS[id]?.related||[]).map(r=>TOOLS[r]).filter(Boolean);
