// The words of the session catalog as the guide writes them: the labelled parts of an Activity block made from a catalog activity (with its source line and its
// timed steps), the row a template lays down for it, and the entry the activity library lists for it. Pure functions over src/session-catalog.mjs and strings; the Review
// (guide-model.mjs) and the templates and library tools (session-templates.mjs) share them, so a step written here is read back the same way.
import {ACTIVITIES} from './session-catalog.mjs';

/** The program's settings file keeps the author's own session templates under this name (src/program-settings.mjs). */
export const TEMPLATES_KEY='osg.sessionTemplates.v1';
const BY_ID=new Map(ACTIVITIES.map(a=>[a.id,a]));
export const activityById=id=>BY_ID.get(String(id))||null;

// ---- the source line: "Scenario rehearsal (A13, version 1)". A copy of an activity says where it came from, in words the facilitator can read and copy with the block.
export const sourceText=a=>`${a.title} (${a.id}, version ${a.version})`;
const SOURCE=/\((A\d\d), version (\d+)\)\s*$/;
export function readSource(text){
  const m=SOURCE.exec(String(text??'').trim());
  return m?{id:m[1],version:Number(m[2])}:null;
}

// ---- timed steps: "3 min: Assign performer, counterpart and observer". The minutes come first; a dash, colon or full stop after them is allowed.
const STEP=/^\s*(\d{1,3})\s*min(?:ute)?s?\b\s*[:.\-–—]?\s*(.*)$/i;
export const stepLine=s=>`${s.min} min: ${s.text}`;
/** The lines of a Timed steps part: the steps that give minutes, the total, and the lines that give none (a step with no minutes cannot be added up). */
export function readSteps(text){
  const steps=[],without=[];
  for(const line of String(text??'').split(/\r?\n/).map(l=>l.trim()).filter(Boolean)){
    const m=STEP.exec(line);
    if(m&&Number(m[1])>0)steps.push({min:Number(m[1]),text:m[2].trim()});
    else without.push(line);
  }
  return {steps,without,minutes:steps.reduce((n,s)=>n+s.min,0)};
}

// ---- needs from earlier work: a line may name an earlier activity and its output: Output of "Priority vote": Ranked shortlist and voting rule
const NEED_FROM=/^Output of "([^"]{1,200})"/;
export const needFrom=title=>`Output of "${String(title).replace(/"/g,"'")}"`;
export const readNeedFrom=line=>{const m=NEED_FROM.exec(String(line??'').trim());return m?m[1]:null;};

// ---- what a catalog activity is, as the labelled parts of an Activity block (every part is a string; an empty one is left out when it is written)
export function catalogParts(a){
  return {
    output:a.output,grouping:a.grouping,phase:a.phase,needs:a.needs||'',
    materials:a.materials.join('\n'),
    steps:a.steps.map(stepLine).join('\n'),
    instructions:a.task.join('\n'),
    debrief:a.debrief.join('\n'),challenges:a.challenges.join('\n'),access:a.access,
    source:sourceText(a)
  };
}
/** The row a template lays down for an activity: a title, its minutes and its parts, which is also what an author's own template keeps. */
export const rowEntry=a=>({title:a.title,minutes:a.minutes,parts:catalogParts(a),notes:'',activityId:a.id});

/** The activity library's entry for a catalog activity (activity-library.mjs lists it with the starters; Insert writes it as parts). */
export function libraryEntry(a){
  return Object.freeze({
    id:`catalog-${a.id}`,kind:a.phase,title:a.title,minutes:a.minutes,
    summary:`${a.grouping}. Makes: ${a.output}.`,method:a.method||'',mode:a.mode,notes:'',
    objective:'',caseStudy:'',...catalogParts(a)
  });
}
export const CATALOG_ENTRIES=Object.freeze(ACTIVITIES.map(libraryEntry));
