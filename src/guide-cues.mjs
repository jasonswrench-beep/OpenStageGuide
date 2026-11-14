// The cue catalog. Adding a cue means adding one row here (and, when it has its own picture, one icon file in
// public/assets/icons-doc). A cue is a label the author sees and a picture that sits beside it. The text label always
// carries the meaning; the picture is marked decorative in the document.
//
// The rows are in group order, which is also the order of the Graphic Cues page. Every cue has a picture of its own: the author's
// fourteen, and 48 more drawn on 2026-11-06 (see public/assets/icons-doc/README.txt). The icon is the picture's file name there.
export const CUE_GROUPS=Object.freeze([
  {id:'teaching',label:'Teaching'},
  {id:'activities',label:'Activities and practice'},
  {id:'slides-media',label:'Slides and media'},
  {id:'materials',label:'Learner materials'},
  {id:'virtual',label:'Virtual class'},
  {id:'elearning',label:'E-Learning'},
  {id:'you',label:'For you'}
]);
export const CUES=Object.freeze([
  // Teaching
  {id:'objective',group:'teaching',label:'Objective',icon:'objective',audience:'everyone',help:'What learners will be able to do.'},
  {id:'say',group:'teaching',label:'Say',icon:'say',audience:'facilitator',help:'Words to say, or the points to make.'},
  {id:'ask',group:'teaching',label:'Ask',icon:'questions',audience:'everyone',help:'A question for the group.'},
  {id:'discussion',group:'teaching',label:'Discussion',icon:'discussion',audience:'everyone',help:'A topic for the group to talk through.'},
  {id:'lecture',group:'teaching',label:'Lecture',icon:'lecture',audience:'everyone',help:'A short talk.'},
  {id:'key-point',group:'teaching',label:'Key point',icon:'key-point',audience:'everyone',help:'A point learners should keep.'},
  {id:'highlight',group:'teaching',label:'Highlight',icon:'highlight',audience:'everyone',help:'Something to call attention to.'},
  {id:'important',group:'teaching',label:'Important',icon:'important',audience:'everyone',help:'A warning or a point that must not be missed.'},
  {id:'process',group:'teaching',label:'Process',icon:'process',audience:'everyone',help:'How something works, from start to finish.'},
  {id:'steps',group:'teaching',label:'Steps',icon:'steps',audience:'everyone',help:'A numbered set of steps to follow.'},
  {id:'tools',group:'teaching',label:'Tools',icon:'tools',audience:'everyone',help:'Tools, templates or job aids to use.'},
  // Activities and practice
  {id:'activity',group:'activities',label:'Activity',icon:'interaction',audience:'everyone',help:'Something learners do.'},
  {id:'do-this',group:'activities',label:'Do this',icon:'do-this',audience:'everyone',help:'A short task for learners to do right now.'},
  {id:'lab',group:'activities',label:'Lab',icon:'lab',audience:'everyone',help:'Hands-on practice with a tool or system.'},
  {id:'role-play',group:'activities',label:'Role play',icon:'role-play',audience:'everyone',help:'Learners act out a situation. Say who plays whom.'},
  {id:'game',group:'activities',label:'Game',icon:'game',audience:'everyone',help:'A game or contest, with its rules.'},
  {id:'case-study',group:'activities',label:'Case study',icon:'case-study',audience:'everyone',help:'A case for learners to read and work through.'},
  {id:'time-on-task',group:'activities',label:'Time on task',icon:'time-on-task',audience:'everyone',help:'Time learners spend working.'},
  {id:'break',group:'activities',label:'Break',icon:'break',audience:'everyone',help:'A break. Give its length in minutes.'},
  // Slides and media
  {id:'slide',group:'slides-media',label:'Slide',icon:'slides',audience:'everyone',help:'Which slide to show and what to do with it.'},
  {id:'slide-notes',group:'slides-media',label:'Slide notes',icon:'slide-notes',audience:'facilitator',help:'What to say or do with the slide. Facilitator only.'},
  {id:'picture',group:'slides-media',label:'Picture',icon:'picture',audience:'everyone',help:'A picture to show or hand out.'},
  {id:'video',group:'slides-media',label:'Video',icon:'video',audience:'everyone',help:'A video to play, and where to start and stop.'},
  {id:'audio',group:'slides-media',label:'Audio',icon:'audio',audience:'everyone',help:'An audio clip to play.'},
  {id:'media',group:'slides-media',label:'Media',icon:'media',audience:'everyone',help:'A video, audio clip or picture to use.'},
  {id:'flipchart',group:'slides-media',label:'Flipchart',icon:'flipchart',audience:'facilitator',help:'What to write or draw on the flipchart.'},
  {id:'data-tables',group:'slides-media',label:'Data tables',icon:'data-tables',audience:'everyone',help:'A table of figures to read or discuss.'},
  // Learner materials
  {id:'workbook',group:'materials',label:'Workbook',icon:'workbook',audience:'everyone',help:'Where learners write in the workbook.'},
  {id:'handout',group:'materials',label:'Handout',icon:'hand-out',audience:'everyone',help:'A handout to give out.'},
  {id:'files',group:'materials',label:'Files',icon:'files',audience:'everyone',help:'Files learners need, and where to get them.'},
  {id:'training-instrument',group:'materials',label:'Training instrument',icon:'training-instrument',audience:'everyone',help:'A tool, survey or instrument learners complete.'},
  {id:'assessment',group:'materials',label:'Assessment',icon:'assessment',audience:'everyone',help:'A check of what learners know or can do.'},
  {id:'evaluation',group:'materials',label:'Evaluation',icon:'evaluation',audience:'everyone',help:'How learners rate the session.'},
  {id:'capture',group:'materials',label:'Capture',icon:'capture',audience:'everyone',help:'A place for learners to capture notes or ideas.'},
  // Virtual class
  {id:'chat',group:'virtual',label:'Chat',icon:'chat',audience:'facilitator',help:'What learners type in the chat, and when.'},
  {id:'breakout',group:'virtual',label:'Breakout',icon:'breakout',audience:'facilitator',help:'Send learners to breakout rooms: task, groups and time.'},
  {id:'polling',group:'virtual',label:'Polling',icon:'polling',audience:'facilitator',help:'A poll to launch, with its question and choices.'},
  {id:'vote',group:'virtual',label:'Vote',icon:'vote',audience:'facilitator',help:'A quick vote, by tool or by hand.'},
  {id:'raise-hand',group:'virtual',label:'Raise hand',icon:'raise-hand',audience:'facilitator',help:'Ask learners to raise a hand, real or virtual.'},
  {id:'microphone',group:'virtual',label:'Microphone',icon:'microphone',audience:'facilitator',help:'When learners should unmute and speak.'},
  {id:'webcam',group:'virtual',label:'Webcam',icon:'webcam',audience:'facilitator',help:'When cameras should be on or off.'},
  {id:'website',group:'virtual',label:'Website',icon:'website',audience:'facilitator',help:'A web page to open or share.'},
  {id:'check-computer',group:'virtual',label:'Check computer',icon:'check-computer',audience:'facilitator',help:'Check that learners can see and hear before you start.'},
  {id:'presenter-1',group:'virtual',label:'Presenter 1',icon:'presenter-1',audience:'facilitator',help:'What the first presenter says or does.'},
  {id:'presenter-2',group:'virtual',label:'Presenter 2',icon:'presenter-2',audience:'facilitator',help:'What the second presenter says or does.'},
  {id:'producer',group:'virtual',label:'Producer',icon:'producer',audience:'facilitator',help:'What the producer does behind the scenes.'},
  {id:'qa',group:'virtual',label:'Q and A',icon:'qa',audience:'everyone',help:'Time for learners’ questions.'},
  // E-Learning: for designing a course screen by screen. What the voice says, what the learner reads and does, and what the
  // developer needs. Narration, Visual direction, Branch, Animation note and Developer note are for the people who build the
  // course, so they are facilitator only by default and never reach a learner copy.
  {id:'screen-title',group:'elearning',label:'Screen title',icon:'screen-title',audience:'everyone',help:'The name of a screen. Each screen starts with one.'},
  {id:'visual-direction',group:'elearning',label:'Visual direction',icon:'visual-direction',audience:'facilitator',help:'The picture, animation or layout for the screen. Facilitator only.'},
  {id:'on-screen-text',group:'elearning',label:'On-screen text',icon:'on-screen-text',audience:'everyone',help:'The words learners read on the screen.'},
  {id:'narration',group:'elearning',label:'Narration',icon:'narration',audience:'facilitator',help:'What the voice says, as it will be recorded. Facilitator only.'},
  {id:'interaction',group:'elearning',label:'Interaction',icon:'click-interaction',audience:'everyone',help:'What the learner does (click, drag, choose) and what happens.'},
  {id:'knowledge-check',group:'elearning',label:'Knowledge check',icon:'knowledge-check',audience:'everyone',help:'A question that follows the screen. Name the question it uses, for example Question 3.'},
  {id:'feedback',group:'elearning',label:'Feedback',icon:'feedback',audience:'everyone',help:'What the learner is told after a choice: right, wrong, and why.'},
  {id:'branch',group:'elearning',label:'Branch',icon:'branch',audience:'facilitator',help:'If the learner chooses this, go to that screen. Facilitator only.'},
  {id:'animation-note',group:'elearning',label:'Animation note',icon:'animation-note',audience:'facilitator',help:'What appears or moves, and when. Facilitator only.'},
  {id:'media-asset',group:'elearning',label:'Media asset',icon:'media-asset',audience:'everyone',help:'A picture, video or sound file the screen uses, and where it is kept.'},
  {id:'accessibility-note',group:'elearning',label:'Accessibility note',icon:'accessibility-note',audience:'everyone',help:'Captions, alt text, keyboard use or reading order for the screen.'},
  {id:'developer-note',group:'elearning',label:'Developer note',icon:'developer-note',audience:'facilitator',help:'A note for whoever builds the screen. Facilitator only.'},
  // For you
  {id:'trainer-notes',group:'you',label:'Trainer notes',icon:'trainer-notes',audience:'facilitator',help:'Advice for the facilitator.'},
  {id:'transition',group:'you',label:'Transition',icon:'transitions',audience:'facilitator',help:'How to move to the next topic.'}
]);
export const AUDIENCES=Object.freeze([
  {id:'everyone',label:'Everyone'},
  {id:'facilitator',label:'Facilitator only'}
]);

// Up to three cues of the author's own, named once for the whole guide. The names live in the document (see the Word adapter),
// and the pane hands them in when it loads a guide, so every module that looks cues up sees them.
export const CUSTOM_CUE_SLOTS=3;
export const CUSTOM_CUE_MAX=24;
export const CUSTOM_ICON='custom-cue';
let customCues=Object.freeze([]);
export const allCues=()=>[...CUES,...customCues];
export const cueById=id=>allCues().find(c=>c.id===id)||null;
export const cuesInGroup=group=>allCues().filter(c=>c.group===group);

// Names that cannot be told apart from an existing label, or from each other. A name that begins with another label and a space
// ("Say more" next to "Say") is refused too: a cue line is recognised by the label it starts with.
function related(a,b){const x=a.toLowerCase(),y=b.toLowerCase();return x===y||x.startsWith(`${y} `)||y.startsWith(`${x} `);}
export function checkCustomCueNames(names){
  const raw=Array.from({length:CUSTOM_CUE_SLOTS},(_,i)=>String(names?.[i]??'').replace(/\s+/g,' ').trim());
  const problems=[],taken=CUES.map(c=>c.label);
  raw.forEach((name,i)=>{
    if(!name)return;
    const slot=`Custom cue ${i+1}`;
    if(name.length>CUSTOM_CUE_MAX)problems.push(`${slot} is too long: use ${CUSTOM_CUE_MAX} letters or fewer.`);
    else if(!/^[\p{L}\p{N}]/u.test(name))problems.push(`${slot} must start with a letter or a number.`);
    else if(/[\p{Cc}:()]/u.test(name))problems.push(`${slot} cannot contain a colon or round brackets ( ).`);
    else{
      const clash=[...taken,...raw.slice(0,i).filter(Boolean)].find(t=>related(name,t));
      if(clash)problems.push(`${slot} ("${name}") is too like "${clash}". Choose a name that does not begin with, or repeat, another cue's name.`);
    }
  });
  return {ok:!problems.length,problems,names:raw};
}
// Replaces the author's own cues. Names that are not acceptable are left out rather than half-used.
export function setCustomCues(names=[]){
  const {names:clean,problems}=checkCustomCueNames(names),bad=new Set();
  if(problems.length)clean.forEach((n,i)=>{if(n&&problems.some(p=>p.startsWith(`Custom cue ${i+1} `)))bad.add(i);});
  customCues=Object.freeze(clean.flatMap((label,i)=>label&&!bad.has(i)?[Object.freeze({id:`custom-${i+1}`,group:'you',label,icon:CUSTOM_ICON,audience:'everyone',help:`Your own cue: ${label}.`,custom:true})]:[]));
  return customCues;
}
export const customCueNames=()=>Array.from({length:CUSTOM_CUE_SLOTS},(_,i)=>customCues.find(c=>c.id===`custom-${i+1}`)?.label||'');

// An Activity can be written as labelled parts, in this order. Each part is a labelled paragraph; the ones for the facilitator
// are written in the facilitator-notes style, so they are shaded and never reach the learner workbook. The activity's minutes
// are the "(N min)" on its first line, as for every block, and its title is that line's text.
export const ACTIVITY_PARTS=Object.freeze([
  {id:'objective',label:'Objective',for:'everyone',help:'What learners will be able to do after the activity.'},
  {id:'materials',label:'Materials',for:'facilitator',help:'What you need to have ready. Facilitator only.'},
  {id:'preparation',label:'Preparation',for:'facilitator',help:'What to do before the activity starts. Facilitator only.'},
  {id:'instructions',label:'Instructions',for:'everyone',help:'What learners do, one step to a line.'},
  {id:'caseStudy',label:'Case study',for:'everyone',help:'A case for learners to read, if the activity uses one.'}
]);
export const TIME_ICON='time-on-task';
export const ICON_FILES=Object.freeze([...new Set([TIME_ICON,CUSTOM_ICON,...CUES.map(c=>c.icon)])]);
export const iconUrl=icon=>`/assets/icons-doc/${icon}.png`;
