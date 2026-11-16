// The learning design models the activity picker draws on, each with what it is, its events or steps, some activities that fit each, and the work it comes from.
// Every model names its source, so that each design choice can be tied to the scholarship behind it: the citation is formatted here in APA 7 as a fallback,
// and by Prism Reference Manager when it is connected (src/design-cite.mjs). The events are the models' own steps in short words, and the ideas are this program's
// suggestions, not the models' text.
//
// `doi` is set only where the DOI was confirmed (Scite, 2026-11-16). Books have no DOI here: their editions and years are to be checked against the copy in hand
// (TESTING-TODO.md), and the program says "book" and nothing more about how far they were checked.

const A=(family,given)=>({family,given});

export const SOURCES=Object.freeze({
  gagne1985:{type:'book',authors:[A('Gagné','R. M.')],year:1985,title:'The conditions of learning and theory of instruction',edition:'4th ed.',publisher:'Holt, Rinehart and Winston'},
  merrill2002:{type:'article-journal',authors:[A('Merrill','M. D.')],year:2002,title:'First principles of instruction',container:'Educational Technology Research and Development',volume:'50',issue:'3',pages:'43–59',doi:'10.1007/BF02505024'},
  keller1987:{type:'article-journal',authors:[A('Keller','J. M.')],year:1987,title:'Development and use of the ARCS model of instructional design',container:'Journal of Instructional Development',volume:'10',issue:'3',pages:'2–10',doi:'10.1007/BF02905780'},
  wiggins2005:{type:'book',authors:[A('Wiggins','G.'),A('McTighe','J.')],year:2005,title:'Understanding by design',edition:'Expanded 2nd ed.',publisher:'Association for Supervision and Curriculum Development'},
  kolb1984:{type:'book',authors:[A('Kolb','D. A.')],year:1984,title:'Experiential learning: Experience as the source of learning and development',publisher:'Prentice-Hall'},
  bybee2006:{type:'report',authors:[A('Bybee','R. W.'),A('Taylor','J. A.'),A('Gardner','A.'),A('Van Scotter','P.'),A('Powell','J. C.'),A('Westbrook','A.'),A('Landes','N.')],year:2006,title:'The BSCS 5E instructional model: Origins and effectiveness',publisher:'BSCS'},
  vanmerrienboer2018:{type:'book',authors:[A('van Merriënboer','J. J. G.'),A('Kirschner','P. A.')],year:2018,title:'Ten steps to complex learning: A systematic approach to four-component instructional design',edition:'3rd ed.',publisher:'Routledge',doi:'10.4324/9781315113210'},
  moore2017:{type:'book',authors:[A('Moore','C.')],year:2017,title:'Map it: The hands-on guide to strategic training design',publisher:'Montesino Press'},
  allen2012:{type:'book',authors:[A('Allen','M. W.'),A('Sites','R.')],year:2012,title:'Leaving ADDIE for SAM: An agile model for developing the best learning experiences',publisher:'ASTD Press'},
  branch2009:{type:'book',authors:[A('Branch','R. M.')],year:2009,title:'Instructional design: The ADDIE approach',publisher:'Springer',doi:'10.1007/978-0-387-09506-6'},
  anderson2001:{type:'book',authors:[A('Anderson','L. W.'),A('Krathwohl','D. R.')],editors:true,year:2001,title:'A taxonomy for learning, teaching, and assessing: A revision of Bloom\'s taxonomy of educational objectives',publisher:'Longman'},
  krathwohl2002:{type:'article-journal',authors:[A('Krathwohl','D. R.')],year:2002,title:'A revision of Bloom\'s taxonomy: An overview',container:'Theory Into Practice',volume:'41',issue:'4',pages:'212–218',doi:'10.1207/s15430421tip4104_2'},
  kirkpatrick2006:{type:'book',authors:[A('Kirkpatrick','D. L.'),A('Kirkpatrick','J. D.')],year:2006,title:'Evaluating training programs: The four levels',edition:'3rd ed.',publisher:'Berrett-Koehler'},
  mayer2009:{type:'book',authors:[A('Mayer','R. E.')],year:2009,title:'Multimedia learning',edition:'2nd ed.',publisher:'Cambridge University Press',doi:'10.1017/CBO9780511811678'},
  sweller1988:{type:'article-journal',authors:[A('Sweller','J.')],year:1988,title:'Cognitive load during problem solving: Effects on learning',container:'Cognitive Science',volume:'12',issue:'2',pages:'257–285',doi:'10.1207/s15516709cog1202_4'},
});

export const MODELS=Object.freeze([
  {id:'gagne',name:'Nine events of instruction',by:'Gagné',source:'gagne1985',kind:'A lesson from start to finish',
    summary:'Nine things that happen, in order, in a lesson that works: from getting attention to helping people keep and use what they learned.',
    useWhen:'You are designing one lesson or one live session and want nothing left out.',
    events:[
      {name:'Gain attention',ideas:['A real problem or a short story','A surprising fact or question','A poll or a show of hands']},
      {name:'Tell learners the objective',ideas:['Say what they will be able to do, and why it matters','Show the finished product']},
      {name:'Recall prior learning',ideas:['Ask what they already do','A quick pair share on past experience']},
      {name:'Present the content',ideas:['A short talk with an example','A demonstration','A worked example']},
      {name:'Guide the learning',ideas:['A job aid or checklist','A case to read and discuss','Hints that fade']},
      {name:'Elicit performance',ideas:['Practice in pairs','A role play','A hands-on task']},
      {name:'Give feedback',ideas:['Peer feedback with a checklist','Facilitator coaching during practice']},
      {name:'Assess performance',ideas:['A skills check','A short quiz','A work sample']},
      {name:'Enhance retention and transfer',ideas:['An action plan for the job','A commitment card','A follow-up message in a week']},
    ]},
  {id:'merrill',name:'First principles of instruction',by:'Merrill',source:'merrill2002',kind:'A course built on real tasks',
    summary:'Learning is promoted when it centers on real problems, activates what people know, shows them, lets them try, and brings it into their work.',
    useWhen:'The goal is to do a task well, and the course can be built around real tasks.',
    events:[
      {name:'Problem: start from a real task',ideas:['A case drawn from the learners\' own work','The whole task shown first']},
      {name:'Activation: build on what they know',ideas:['Ask for their experience','A short diagnostic']},
      {name:'Demonstration: show it',ideas:['A model performance with the steps named','A before and after example']},
      {name:'Application: let them do it',ideas:['Practice with coaching that fades','A scenario with choices']},
      {name:'Integration: bring it into their work',ideas:['Plan where they will use it','Teach it to a colleague','A reflection on what changes']},
    ]},
  {id:'arcs',name:'ARCS: attention, relevance, confidence, satisfaction',by:'Keller',source:'keller1987',kind:'Keeping people motivated',
    summary:'Four things to plan for so people want to learn: get attention, show it matters to them, build their confidence, and let them feel the result.',
    useWhen:'Motivation is the worry: a required course, a hard topic, or a tired group.',
    events:[
      {name:'Attention',ideas:['A question or a puzzle','A change of pace or of activity','A real example with a surprise in it']},
      {name:'Relevance',ideas:['Link each topic to their job','Let them choose an example','Tell who has used it and how']},
      {name:'Confidence',ideas:['Clear objectives and criteria','Small steps that build','Practice with feedback before any assessment']},
      {name:'Satisfaction',ideas:['Use it on something real, early','Recognition of what they did','Feedback that says what worked']},
    ]},
  {id:'backward',name:'Backward design (Understanding by design)',by:'Wiggins and McTighe',source:'wiggins2005',kind:'A whole course, planned from the end',
    summary:'Plan from the result: first what people should be able to do, then the evidence that they can, and only then the activities.',
    useWhen:'You are planning a course or a program and want every activity to earn its place.',
    events:[
      {name:'Stage 1: Desired results',ideas:['Write the objectives','Name the big ideas that matter most']},
      {name:'Stage 2: Evidence',ideas:['Decide the performance that shows it','Plan the assessment before the lessons']},
      {name:'Stage 3: Learning plan',ideas:['Choose activities that lead to the evidence','Order them from simple to whole task']},
    ]},
  {id:'kolb',name:'Experiential learning cycle',by:'Kolb',source:'kolb1984',kind:'Learning from doing',
    summary:'A cycle of doing, looking back, drawing a lesson from it, and trying it again.',
    useWhen:'The skill is best learned by doing and then thinking about what happened: leadership, coaching, customer work.',
    events:[
      {name:'Concrete experience',ideas:['A simulation or role play','A real task with a debrief to follow']},
      {name:'Reflective observation',ideas:['Debrief questions: what happened, what did you notice','Journal for five minutes']},
      {name:'Abstract conceptualization',ideas:['Name the principle the group found','Compare with a model or a rule of thumb']},
      {name:'Active experimentation',ideas:['Try it again, changed','Plan what to test back at work']},
    ]},
  {id:'5e',name:'The 5E instructional model',by:'Bybee and others',source:'bybee2006',kind:'Learning by inquiry',
    summary:'Five phases that move from curiosity to explanation to use to checking.',
    useWhen:'The topic is a concept or a process that people can explore before they are told about it.',
    events:[
      {name:'Engage',ideas:['A puzzling case','A prediction to make and check']},
      {name:'Explore',ideas:['A hands-on task before any explanation','A data set to look at in groups']},
      {name:'Explain',ideas:['Learners explain first, then the facilitator names the concept','A short input to tie it together']},
      {name:'Elaborate',ideas:['Apply it to a new case','Extend with a harder problem']},
      {name:'Evaluate',ideas:['A performance task','Self-assessment against the objectives']},
    ]},
  {id:'4cid',name:'Four-component instructional design (4C/ID)',by:'van Merriënboer and Kirschner',source:'vanmerrienboer2018',kind:'A complex skill, taught whole',
    summary:'A complex skill is taught through whole tasks, with the information they need, the steps to follow and practice on the parts that need it.',
    useWhen:'The job is complex (a clinician, a technician, a manager) and no single lesson will teach it.',
    events:[
      {name:'Learning tasks',ideas:['Whole-task cases from easy to hard','Support that fades as they improve']},
      {name:'Supportive information',ideas:['How to reason about this kind of task','Examples of expert approaches']},
      {name:'Procedural information',ideas:['Step-by-step job aids shown just when they are needed']},
      {name:'Part-task practice',ideas:['Drills on a step that must become automatic']},
    ]},
  {id:'action',name:'Action mapping',by:'Moore',source:'moore2017',kind:'Training that changes what people do',
    summary:'Start from the business goal, list what people must do, find out why they do not, and give them practice in the situations where it matters.',
    useWhen:'There is a performance problem and you need to decide what training, if any, should do about it.',
    events:[
      {name:'Business goal',ideas:['Write one measurable goal for the work']},
      {name:'What people must do',ideas:['List the actions on the job that would reach the goal']},
      {name:'Why they do not do it now',ideas:['Knowledge, skill, motivation, environment: only some need training']},
      {name:'Practice activities',ideas:['A realistic scenario with a choice and a consequence for each action']},
      {name:'Only the information needed',ideas:['A job aid or reference in place of a lecture']},
    ]},
  {id:'sam',name:'Successive approximation model (SAM)',by:'Allen and Sites',source:'allen2012',kind:'A project plan with quick rounds',
    summary:'Plan the project as short rounds of design, build and review with the people who will use the course, so problems show early.',
    useWhen:'The course is a project with a team, a deadline and people who can review early versions.',
    events:[
      {name:'Preparation',ideas:['Gather what is known and what is needed','Agree the aim and the limits']},
      {name:'Iterative design',ideas:['Quick sketches or a storyboard to react to','A prototype of one lesson']},
      {name:'Iterative development',ideas:['Build in rounds','Review each round with learners and fix what shows up']},
    ]},
  {id:'addie',name:'ADDIE: analyze, design, develop, implement, evaluate',by:'Branch',source:'branch2009',kind:'The phases of a design project',
    summary:'The five phases that most design projects pass through, each with what goes in and what comes out.',
    useWhen:'You want a plain frame to plan the phases of a project, or to explain them to others.',
    events:[
      {name:'Analyze',ideas:['A needs analysis','A description of the learners and the setting']},
      {name:'Design',ideas:['Objectives, assessments and a plan for the activities']},
      {name:'Develop',ideas:['Write and build the materials','Review them with a sample of learners']},
      {name:'Implement',ideas:['Run the course','Prepare the facilitators']},
      {name:'Evaluate',ideas:['The evaluation plan, by level','Fix what the evidence shows']},
    ]},
]);

export const modelById=id=>MODELS.find(m=>m.id===id)||null;
export const sourceFor=id=>{const m=modelById(id);return m?SOURCES[m.source]:SOURCES[id]||null;};

/** What the program cites on its other screens besides the models: the taxonomy of objectives, the four levels of evaluation, and two principles of design. */
export const REFERENCE_SOURCES=Object.freeze([
  {id:'bloom',use:'The levels and verbs of the objectives',sources:['anderson2001','krathwohl2002']},
  {id:'kirkpatrick',use:'The four levels of the evaluation plan',sources:['kirkpatrick2006']},
  {id:'mayer',use:'Words and pictures on screens',sources:['mayer2009']},
  {id:'load',use:'Cognitive load',sources:['sweller1988']},
]);

// ---- the citation, formatted in APA 7 (the program's own, for when no Reference Manager is connected)

const names=list=>{
  const one=a=>`${a.family}, ${a.given}`;
  if(list.length===1)return one(list[0]);
  if(list.length===2)return `${one(list[0])}, & ${one(list[1])}`;
  if(list.length<=20)return `${list.slice(0,-1).map(one).join(', ')}, & ${one(list.at(-1))}`;
  return `${list.slice(0,19).map(one).join(', ')}, . . . ${one(list.at(-1))}`;
};
/** The reference-list entry, as plain text: "Merrill, M. D. (2002). First principles of instruction. Educational Technology Research and Development, 50(3), 43–59. https://doi.org/10.1007/BF02505024" */
export function referenceText(s){
  const who=s.editors?`${names(s.authors)} (Eds.).`:names(s.authors);
  const stop=t=>/[.?!]$/.test(t)?t:`${t}.`;
  const doi=s.doi?` https://doi.org/${s.doi}`:'';
  const lead=`${who.endsWith('.')?who:`${who}`} (${s.year}).`;
  if(s.type==='article-journal')return `${lead} ${stop(s.title)} ${s.container}, ${s.volume}${s.issue?`(${s.issue})`:''}, ${s.pages}.${doi}`;
  return `${lead} ${stop(`${s.title}${s.edition?` (${s.edition})`:''}`)} ${s.publisher}.${doi}`;
}
/** The citation in the words of a sentence: "(Merrill, 2002)", "(Wiggins & McTighe, 2005)", "(Bybee et al., 2006)". */
export function inTextCitation(s){
  const f=s.authors.map(a=>a.family);
  const who=f.length===1?f[0]:f.length===2?`${f[0]} & ${f[1]}`:`${f[0]} et al.`;
  return `(${who}, ${s.year})`;
}
