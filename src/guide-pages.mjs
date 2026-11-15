// The two Getting Started pages the pane builds from the guide itself: Training at a Glance (a table of the modules and
// lessons with their times) and Graphic Cues (the cues the guide uses, with their pictures). Both are ordinary Word content
// in the guide's own styles. They are built when the author chooses the button and are never changed by anything else, so
// the pane only tells the author when one has fallen out of date. Pure functions: the guide in, OOXML and positions out.
import {para,run,plural,iconRun,heading,GLANCE_PROMPT} from './guide-ooxml.mjs';
import {cell,row,table,head,PAGE} from './guide-assess-ooxml.mjs';
import {STYLE_IDS} from './guide-styles.mjs';
import {allCues,TIME_ICON} from './guide-cues.mjs';
import {readGuide,lessonMinutes,kindOf,cueFromText,readBlock} from './guide-model.mjs';
import {readForms} from './guide-forms.mjs';
import {readFacts} from './guide-facts.mjs';

const alnum=s=>String(s??'').replace(/[^\p{L}\p{N}]+/gu,'').toLowerCase();
const isHeading2=p=>/^heading 2$/i.test(String(p.style||''))||/^Heading2$/.test(String(p.builtIn||''));
const headingIndex=(paras,name)=>paras.findIndex(p=>isHeading2(p)&&alnum(p.text)===alnum(name));
const CENTRED='<w:jc w:val="center"/>';
const MODULE_FILL='F1F3F6';

// Where the page under the heading at `gi` ends: at the next heading, module, lesson, assessment or title.
function sectionEnd(paras,gi){
  const stop=i=>{
    const p=paras[i],k=kindOf(p);
    return k==='module'||k==='lesson'||k==='assessment'||k==='guide-title'||k==='title'||/^heading\b/i.test(String(p.style||''))||/^Heading/.test(String(p.builtIn||''));
  };
  let i=gi+1;
  while(i<paras.length&&!stop(i))i++;
  return i;
}
// The paragraphs this page's own button wrote, found anywhere under its heading so that words the author typed above or
// below them change nothing: an intro line, a table whose first header cell reads `firstHead`, and (for the Glance) a closing
// line. Cells are paragraphs, as Word counts them. null when they are not there.
function findBuilt(paras,gi,{intro,firstHead,closing=null}){
  const end=sectionEnd(paras,gi);
  for(let from=gi+1;from<end;from++){
    if(alnum(paras[from].text)!==alnum(intro)||kindOf(paras[from])!=='other')continue;
    let i=from+1;
    if(kindOf(paras[i]||{})!=='table-head'||alnum(paras[i].text)!==alnum(firstHead))continue;
    while(['table-head','table-cell'].includes(kindOf(paras[i]||{})))i++;
    if(closing&&paras[i]&&kindOf(paras[i])==='other'&&alnum(paras[i].text).startsWith(alnum(closing)))i++;
    return {start:from,end:i-1};
  }
  return null;
}
// Where a page with no heading of its own goes: before the first of the named Heading 2 pages that exists, else before the
// first module, lesson or assessment. `after` is the paragraph to insert after, -1 meaning the very top.
function before(paras,names){
  for(const name of names){const i=headingIndex(paras,name);if(i>=0)return i-1;}
  const first=paras.findIndex(p=>['module','lesson','assessment'].includes(kindOf(p)));
  return (first<0?paras.length:first)-1;
}
const sameWords=(paras,span,want)=>{
  const have=paras.slice(span.start,span.end+1).map(p=>alnum(p.text)),expect=want.map(alnum);
  return have.length===expect.length&&have.every((t,i)=>t===expect[i]);
};

// ---- the forms a part of the guide cites
// A form's short name for a list: "Handout 1.1" (its label and number), or its whole title when it has no number yet.
const shortName=f=>f.numbered&&f.number?`${f.label} ${f.number}`:f.text;
// The forms cited from the paragraphs `start` to `end`, each once, in the order they are cited: all of them, those for everyone (the participants' handouts)
// and those for the facilitator only. A citation of a form no longer in the guide is not listed (the Review says so).
export function formsCited(paras,start,end,read=readForms(paras)){
  const all=[],everyone=[],facilitator=[];
  for(const c of read.cites){
    if(c.index<start||c.index>end)continue;
    const f=read.byId.get(c.id);
    if(!f)continue;
    const name=shortName(f);
    if(all.includes(name))continue;
    all.push(name);
    (f.audience==='facilitator'?facilitator:everyone).push(name);
  }
  return {all,everyone,facilitator};
}
// For each lesson (by the index of its heading), the forms its blocks cite. Empty when the guide cites no form at all, which is how the pages know to leave
// out their forms columns and lines.
export function lessonForms(paras,guide){
  const read=readForms(paras),byLesson=new Map();
  if(!read.cites.length)return byLesson;
  for(const l of guide.modules.flatMap(m=>m.lessons)){
    const names=formsCited(paras,l.index,l.end,read).all;
    if(names.length)byLesson.set(l.index,names);
  }
  return byLesson;
}

// For each lesson (by the index of its heading) that says what method it uses, the method (kit slice K1). Empty when no lesson says, which is how the pages
// know to leave out their method columns.
export function lessonMethods(paras,guide){
  const byLesson=new Map();
  for(const l of guide.modules.flatMap(m=>m.lessons)){
    const {facts}=readFacts(paras,l);
    if(facts.method)byLesson.set(l.index,facts.method);
  }
  return byLesson;
}

// ---- Training at a Glance
export const GLANCE_HEADING='Training at a Glance';
export const GLANCE_INTRO='The minutes each module and lesson takes, and the minute of the session at which each one starts.';
export const GLANCE_HEAD=['Module or lesson','Minutes','Starts at minute'];
const GLANCE_WIDTHS=[5560,1900,1900];
// Two columns are added only when there is something to put in them: Method (a lesson says what method it uses) and Forms used (blocks cite forms). A guide
// with neither keeps the three it always had.
const GLANCE_METHOD_HEAD='Method',GLANCE_FORMS_HEAD='Forms used';
const GLANCE_WIDTHS_BY={method:[4300,1200,1300,2560],forms:[4000,1200,1300,2860],both:[3300,1100,1200,1700,2060]};
const glanceColumns=(used,methods)=>{
  const method=methods.size>0,forms=used.size>0;
  return {method,forms,heads:[...GLANCE_HEAD,...(method?[GLANCE_METHOD_HEAD]:[]),...(forms?[GLANCE_FORMS_HEAD]:[])],widths:method&&forms?GLANCE_WIDTHS_BY.both:method?GLANCE_WIDTHS_BY.method:forms?GLANCE_WIDTHS_BY.forms:GLANCE_WIDTHS};
};
const GLANCE_CLOSING='Total teaching time';
const glanceTotal=total=>`${GLANCE_CLOSING}: ${plural(total,'minute')}.`;
const GLANCE_SHAPE={intro:GLANCE_INTRO,firstHead:GLANCE_HEAD[0],closing:GLANCE_CLOSING};

// One row for each module (its lessons' minutes added up) and each lesson. A lesson starts at the minutes taught before it,
// the same cumulative time the lesson's own time block shows.
export function glanceRows(guide,used=new Map(),methods=new Map()){
  const rows=[];let cumulative=0;
  for(const mod of guide.modules){
    const start=cumulative,inModule=[];
    const lessons=mod.lessons.map(l=>{
      const minutes=lessonMinutes(l),names=used.get(l.index)||[],r={kind:'lesson',name:l.text||'Untitled lesson',minutes,starts:cumulative,method:methods.get(l.index)||'',forms:names.join(', ')};
      for(const n of names)if(!inModule.includes(n))inModule.push(n);
      cumulative+=minutes;return r;
    });
    if(mod.index>=0)rows.push({kind:'module',name:mod.text||'Untitled module',minutes:cumulative-start,starts:start,method:'',forms:inModule.join(', ')});
    rows.push(...lessons);
  }
  return {rows,total:cumulative};
}
const glanceWords=(guide,used=new Map(),methods=new Map())=>{
  const {rows,total}=glanceRows(guide,used,methods),c=glanceColumns(used,methods);
  return [GLANCE_INTRO,...c.heads,...rows.flatMap(r=>[r.name,String(r.minutes),String(r.starts),...(c.method?[r.method]:[]),...(c.forms?[r.forms]:[])]),glanceTotal(total)];
};

// `used` is lessonForms' map and `methods` lessonMethods': what each holds decides which of the two extra columns the table has (a module row lists the forms of
// its lessons, and no method).
export function glanceBody(guide,{withHeading=false,used=new Map(),methods=new Map()}={}){
  const {rows,total}=glanceRows(guide,used,methods),c=glanceColumns(used,methods);
  if(!rows.some(r=>r.kind==='lesson'))throw new Error('Add a module and a lesson first. Training at a Glance lists them with their times.');
  const widths=c.widths;
  const trs=[
    row(c.heads.map((t,i)=>head(t,widths[i],i>0&&i<3)),{header:true}),
    ...rows.map(r=>{
      const fill=r.kind==='module'?MODULE_FILL:undefined;
      let at=3;
      return row([
        cell(widths[0],para(STYLE_IDS.tableCell,run(r.name,r.kind==='module'?STYLE_IDS.cueLabel:undefined)),{fill}),
        cell(widths[1],para(STYLE_IDS.tableCell,run(String(r.minutes)),CENTRED),{fill}),
        cell(widths[2],para(STYLE_IDS.tableCell,run(String(r.starts)),CENTRED),{fill}),
        ...(c.method?[cell(widths[at++],para(STYLE_IDS.tableCell,r.method?run(r.method):''),{fill})]:[]),
        ...(c.forms?[cell(widths[at++],para(STYLE_IDS.tableCell,r.forms?run(r.forms):''),{fill})]:[])
      ]);
    })
  ];
  const extras=[c.method?'the method of each lesson':'',c.forms?'the forms its blocks cite':''].filter(Boolean).join(' and ');
  return (withHeading?heading(2,GLANCE_HEADING):'')+para(null,run(GLANCE_INTRO))
    +table(widths,trs,{caption:'Training at a Glance',description:`Each module and lesson of the guide, with the minutes it takes and the minute at which it starts${extras?`, and ${extras}`:''}. Module rows give the total for their lessons.`})
    +para(null,run(glanceTotal(total)));
}

// Whether the page is in the document and whether it still says what the guide says now.
export function glanceState(paras){
  const gi=headingIndex(paras,GLANCE_HEADING),built=gi>=0?findBuilt(paras,gi,GLANCE_SHAPE):null;
  const guide=readGuide(paras);
  return {built:!!built,stale:built?!sameWords(paras,built,glanceWords(guide,lessonForms(paras,guide),lessonMethods(paras,guide))):false};
}
// What building the page does to the document: what it replaces (a table built before, or the untouched placeholder and
// nothing else), where it goes, and whether it must bring its own heading.
export function glancePlacement(paras){
  const gi=headingIndex(paras,GLANCE_HEADING);
  if(gi<0)return {withHeading:true,after:before(paras,[]),remove:null,replaces:null};
  const built=findBuilt(paras,gi,GLANCE_SHAPE);
  if(built)return {withHeading:false,after:built.end,remove:{start:built.start,end:built.end,guard:paras[built.start].text},replaces:'page'};
  const end=sectionEnd(paras,gi);
  for(let i=gi+1;i<end;i++)
    if(kindOf(paras[i])==='other'&&alnum(paras[i].text)===alnum(GLANCE_PROMPT))return {withHeading:false,after:i,remove:{start:i,end:i,guard:paras[i].text},replaces:'placeholder'};
  return {withHeading:false,after:gi,remove:null,replaces:null};
}

// ---- Graphic Cues
export const CUES_HEADING='Graphic Cues';
export const CUES_INTRO='These pictures mark the kind of block you are reading. The word beside each picture says the same thing, so the picture is a help and never the only clue.';
export const CUES_HEAD=['Cue','What it means'];
const CUES_WIDTHS=[3000,PAGE-3000];
const CUES_SHAPE={intro:CUES_INTRO,firstHead:CUES_HEAD[0]};
const TIME_ENTRY={id:'time',label:'Time',icon:TIME_ICON,help:'How long a lesson takes, and how many minutes have gone by when it starts.'};

// The cues the guide really uses, in the catalog's order (Time first, since each lesson opens with one).
export function cuesUsed(paras){
  const used=new Set();let time=false;
  for(const p of paras){
    const k=kindOf(p);
    if(k==='time')time=true;
    else if(k==='cue'){const c=cueFromText(p.text);if(c)used.add(c.id);}
  }
  return [...(time?[TIME_ENTRY]:[]),...allCues().filter(c=>used.has(c.id))];
}
const cuesWords=entries=>[CUES_INTRO,...CUES_HEAD,...entries.flatMap(e=>[e.label,e.help])];

export function cuesBody(entries,ctx,{withHeading=false}={}){
  if(!entries.length)throw new Error('Add some blocks first. The Graphic Cues page shows the cues your guide uses.');
  const [w0,w1]=CUES_WIDTHS;
  const trs=[
    row(CUES_HEAD.map((t,i)=>head(t,CUES_WIDTHS[i])),{header:true}),
    // The picture is decorative and sits beside the cue's name, so every cell has words and a screen reader hears the name.
    ...entries.map(e=>row([
      cell(w0,para(STYLE_IDS.tableCell,iconRun(e.icon,ctx)+run(` ${e.label}`,STYLE_IDS.cueLabel))),
      cell(w1,para(STYLE_IDS.tableCell,run(e.help)))
    ]))
  ];
  return (withHeading?heading(2,CUES_HEADING):'')+para(null,run(CUES_INTRO))
    +table(CUES_WIDTHS,trs,{caption:'Graphic Cues',description:'Each cue picture this guide uses, with its name and what it means.'});
}

export function cuesState(paras){
  const gi=headingIndex(paras,CUES_HEADING),built=gi>=0?findBuilt(paras,gi,CUES_SHAPE):null;
  return {built:!!built,stale:built?!sameWords(paras,built,cuesWords(cuesUsed(paras))):false};
}
export function cuesPlacement(paras){
  const gi=headingIndex(paras,CUES_HEADING);
  // Graphic Cues sit between About This Guide and The Program in Perspective, as in the original specification.
  if(gi<0)return {withHeading:true,after:before(paras,['The Program in Perspective','Program Preparation',GLANCE_HEADING]),remove:null,replaces:null};
  const built=findBuilt(paras,gi,CUES_SHAPE);
  if(built)return {withHeading:false,after:built.end,remove:{start:built.start,end:built.end,guard:paras[built.start].text},replaces:'page'};
  return {withHeading:false,after:gi,remove:null,replaces:null};
}

// ---- Module overview
// The page under a module's heading, before its first lesson. Some of it the guide can write (the module's time, and a table
// of its lessons with their minutes and objectives) and some only the author can (goal, overview, materials, references). The
// first build writes both, with a sentence to type over for each of the author's lines. After that the button rewrites only the
// part it can write, found by its shape, so the author's lines are never touched. All of it is for the facilitator.
export const OVERVIEW_HEAD=['Lesson','Minutes','Objective'];
const OVERVIEW_WIDTHS=[3200,1100,PAGE-4300];
const OVERVIEW_TOP=[
  ['Goal:','Say what learners will be able to do by the end of this module.'],
  ['Overview:','Describe what happens in this module and how it connects to the rest of the program.'],
  ['Key message:','State the one idea learners should keep from this module.']
];
const OVERVIEW_BOTTOM=[
  ['Materials needed:','List the materials, equipment and room set-up this module needs.'],
  ['Teaching aids:','List the slides, handouts and equipment used to teach this module.'],
  ['References and resources:','List any readings, links or other resources.']
];
const NO_OBJECTIVE='No objective yet.';
const overviewLine=([label,text])=>para(STYLE_IDS.overview,run(label,STYLE_IDS.cueLabel)+run(` ${text}`));
const timeLine=facts=>`Time: ${plural(facts.total,'minute')} in ${plural(facts.lessons.length,'lesson')}.`;

// The module the cursor is in (the last module heading at or before it), as readGuide describes it, or null.
export function moduleAt(paras,at){
  return readGuide(paras).modules.filter(m=>m.index>=0&&m.index<=at).at(-1)||null;
}
// What the overview says about a module: each lesson with its minutes (the same rule Update guide uses) and its objectives.
export function overviewFacts(paras,mod){
  const lessons=mod.lessons.map(l=>{
    const objectives=[];
    for(let i=l.index+1;i<=l.end;i++)
      if(kindOf(paras[i])==='cue'&&cueFromText(paras[i].text)?.id==='objective'){
        const b=readBlock(paras,i);
        if(b?.text.trim())objectives.push(b.text.replace(/\s*\n\s*/g,' ').trim());
      }
    return {name:l.text||'Untitled lesson',minutes:lessonMinutes(l),objective:objectives.join('; ')||NO_OBJECTIVE,method:readFacts(paras,l).facts.method};
  });
  // The forms the module's blocks cite, for the participants and for the facilitator only (empty unless the guide cites forms).
  const cited=formsCited(paras,mod.index,mod.end);
  // The slides of the module's lessons, counted from the Slide blocks (kit slice K1).
  let slides=0;
  for(let i=mod.index+1;i<=mod.end;i++)if(kindOf(paras[i])==='cue'&&cueFromText(paras[i].text)?.id==='slide')slides++;
  return {name:mod.text||'This module',lessons,total:lessons.reduce((n,l)=>n+l.minutes,0),handouts:cited.everyone,facilitatorForms:cited.facilitator,slides};
}
// The lines under the lessons table that name the forms: "Participant handouts: Handout 1.1, Checklist 1.2." and, when there are any, "Facilitator forms: ...".
const formLines=facts=>[...(facts.handouts?.length?[['Participant handouts:',` ${facts.handouts.join(', ')}.`]]:[]),...(facts.facilitatorForms?.length?[['Facilitator forms:',` ${facts.facilitatorForms.join(', ')}.`]]:[]),...(facts.slides>0?[['Slides:',` ${plural(facts.slides,'slide')}.`]]:[])];
const FORM_LINE=/^[^A-Za-z]*(Participant handouts|Facilitator forms|Slides):/;
// With any lesson that says its method the lessons table has a Method column; with none it keeps its three.
const overviewColumns=facts=>{const method=facts.lessons.some(l=>l.method);return {method,heads:method?['Lesson','Minutes','Method','Objective']:OVERVIEW_HEAD,widths:method?[2700,1000,1700,PAGE-5400]:OVERVIEW_WIDTHS};};
const overviewWords=facts=>{
  const c=overviewColumns(facts);
  return [timeLine(facts),...c.heads,...facts.lessons.flatMap(l=>[l.name,String(l.minutes),...(c.method?[l.method]:[]),l.objective]),...formLines(facts).map(([a,b])=>`${a}${b}`)];
};

// withAuthored is true for a first build, which also writes the author's own lines; otherwise only the time line and the lessons table.
export function overviewBody(facts,{withAuthored=false}={}){
  if(!facts.lessons.length)throw new Error('Add a lesson to this module first. The overview lists its lessons with their minutes.');
  const c=overviewColumns(facts),widths=c.widths;
  const trs=[
    row(c.heads.map((t,i)=>head(t,widths[i],i===1)),{header:true}),
    ...facts.lessons.map(l=>row([
      cell(widths[0],para(STYLE_IDS.tableCell,run(l.name))),
      cell(widths[1],para(STYLE_IDS.tableCell,run(String(l.minutes)),CENTRED)),
      ...(c.method?[cell(widths[2],para(STYLE_IDS.tableCell,l.method?run(l.method):''))]:[]),
      cell(widths[c.method?3:2],para(STYLE_IDS.tableCell,run(l.objective)))
    ]))
  ];
  const generated=para(STYLE_IDS.overview,run('Time:',STYLE_IDS.cueLabel)+run(timeLine(facts).slice('Time:'.length)))
    +table(widths,trs,{caption:`Lessons in ${facts.name}`,description:`Each lesson in ${facts.name}, with the minutes it takes${c.method?', its method':''} and its objective.`})
    +formLines(facts).map(([a,b])=>para(STYLE_IDS.overview,run(a,STYLE_IDS.cueLabel)+run(b))).join('');
  return withAuthored?OVERVIEW_TOP.map(overviewLine).join('')+generated+OVERVIEW_BOTTOM.map(overviewLine).join(''):generated;
}

// What lies between a module's heading and its first lesson, block or other heading.
function overviewRegion(paras,mod){
  const stop=i=>{
    const p=paras[i],k=kindOf(p);
    return k!=='overview'&&k!=='table-head'&&k!=='table-cell'&&k!=='other'||/^heading\b/i.test(String(p.style||''))||/^Heading/.test(String(p.builtIn||''));
  };
  let end=mod.index+1;
  while(end<paras.length&&!stop(end))end++;
  return {start:mod.index+1,end};
}
// The time line and lessons table this button wrote, found by their shape anywhere in the region, and the author's own lines.
function findOverview(paras,mod){
  const {start,end}=overviewRegion(paras,mod);
  let built=null;
  for(let from=start;from<end&&!built;from++){
    if(kindOf(paras[from])!=='overview'||!/^[^A-Za-z]*Time:/.test(paras[from].text))continue;
    let i=from+1;
    if(kindOf(paras[i]||{})!=='table-head'||alnum(paras[i].text)!==alnum(OVERVIEW_HEAD[0]))continue;
    while(['table-head','table-cell'].includes(kindOf(paras[i]||{})))i++;
    // The lines that name the forms come straight after the table, and are part of what the button wrote.
    while(kindOf(paras[i]||{})==='overview'&&FORM_LINE.test(paras[i].text))i++;
    built={start:from,end:i-1};
  }
  let lastOwn=-1;
  for(let i=start;i<end;i++)if(kindOf(paras[i])==='overview'&&!(built&&i>=built.start&&i<=built.end))lastOwn=i;
  return {built,lastOwn};
}
export function overviewState(paras,mod){
  const {built}=findOverview(paras,mod);
  return {built:!!built,stale:built?!sameWords(paras,built,overviewWords(overviewFacts(paras,mod))):false};
}
export function overviewPlacement(paras,mod){
  const {built,lastOwn}=findOverview(paras,mod);
  if(built)return {withAuthored:false,after:built.end,remove:{start:built.start,end:built.end,guard:paras[built.start].text},replaces:'generated'};
  // The author's own lines are already there (the time line and table were deleted): add only what the guide can write.
  if(lastOwn>=0)return {withAuthored:false,after:lastOwn,remove:null,replaces:null};
  return {withAuthored:true,after:mod.index,remove:null,replaces:null};
}
// How many modules have an overview and how many of those are out of date, for the note under the button.
export function overviewsSummary(paras){
  const mods=readGuide(paras).modules.filter(m=>m.index>=0),states=mods.map(m=>overviewState(paras,m));
  return {total:mods.length,built:states.filter(s=>s.built).length,stale:states.filter(s=>s.stale).length};
}
