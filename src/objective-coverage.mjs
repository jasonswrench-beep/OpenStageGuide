import {validate,value} from './model.mjs';
import {objectiveText,unsupportedObjective} from './objective-text.mjs';
export const COVERAGE_ROLES=Object.freeze(['practice','assessment']);
const fail=()=>{throw new Error('Objective annotations are unsupported, duplicated or refer to another project/item. Preserve this guide; their edits are unavailable.');};
const object=x=>!!x&&typeof x==='object'&&!Array.isArray(x),note=x=>typeof x==='string'&&!!x.trim()&&x.length<=50000;
const keys=(x,allowed)=>object(x)&&Object.keys(x).length===allowed.length&&allowed.every(k=>Object.hasOwn(x,k));
const stable=x=>Array.isArray(x)?x.map(stable):object(x)?Object.fromEntries(Object.keys(x).sort().map(k=>[k,stable(x[k])])):x;
const same=(a,b)=>JSON.stringify(stable(a))===JSON.stringify(stable(b));
function read(project){
 validate(project);const absent=project.objectiveCoverage===undefined,data=absent?{version:1,projectId:project.id,links:[],exceptions:[]}:project.objectiveCoverage;
 if(!object(data)||data.version!==1||data.projectId!==project.id||!Array.isArray(data.links)||!Array.isArray(data.exceptions)||data.links.length>2000||data.exceptions.length>2000)fail();
 const modules=new Set(project.modules.map(m=>m.id)),activities=new Set(project.modules.flatMap(m=>m.activities.map(a=>a.id))),pairs=new Set(),gaps=new Set();
 for(const l of data.links){const key=JSON.stringify([l?.activityId,l?.objectiveModuleId]);if(!object(l)||!activities.has(l.activityId)||!modules.has(l.objectiveModuleId)||pairs.has(key)||!Array.isArray(l.roles)||!l.roles.length||l.roles.length>2||new Set(l.roles).size!==l.roles.length||l.roles.some(r=>!COVERAGE_ROLES.includes(r))||!note(l.reviewedObjective)||!note(l.rationale))fail();pairs.add(key);}
 for(const e of data.exceptions){const key=JSON.stringify([e?.objectiveModuleId,e?.role]);if(!object(e)||!modules.has(e.objectiveModuleId)||!COVERAGE_ROLES.includes(e.role)||gaps.has(key)||!note(e.reviewedObjective)||!note(e.note))fail();gaps.add(key);}
 return {absent,data};
}
export function objectiveCoverage(project){try{const {absent,data}=read(project);return {supported:true,present:!absent,data:structuredClone(data)};}catch{return {supported:false,present:project?.objectiveCoverage!==undefined,message:'Objective annotations are unsupported or reference missing/foreign items. They are preserved; no current coverage is counted. Other guide outputs remain available.'};}}
function currentText(project,id){const module=project.modules.find(m=>m.id===id);if(!module)fail();if(unsupportedObjective(module))throw new Error('This module objective has an unsupported value. Preserve this guide and repair the objective before editing its annotations.');return objectiveText(module);}
function snapshot(project,id,existing,fields){
 if(typeof fields.reviewCurrent!=='boolean'||typeof fields.reviewedObjective!=='string')fail();
 if(fields.reviewCurrent){const text=currentText(project,id);if(!text.trim()||fields.reviewedObjective!==text)throw new Error('Review the exact current nonblank module objective before approving this annotation.');return text;}
 if(!existing||fields.reviewedObjective!==existing.reviewedObjective)throw new Error('A new or changed objective snapshot requires explicit current-objective review.');return existing.reviewedObjective;
}
export function editObjectiveCoverage(project,change){
 const {absent,data}=read(project),next=structuredClone(project),copy=structuredClone(data);if(!object(change))fail();
 if(change.kind==='links'){
  if(!keys(change,['kind','activityId','links'])||!project.modules.some(m=>m.activities.some(a=>a.id===change.activityId))||!Array.isArray(change.links)||change.links.length>project.modules.length)fail();const proposals=new Map();
  for(const old of data.links.filter(l=>l.activityId===change.activityId))currentText(project,old.objectiveModuleId);
  for(const input of change.links){if(!keys(input,['objectiveModuleId','roles','reviewedObjective','rationale','reviewCurrent'])||typeof input.rationale!=='string'||proposals.has(input.objectiveModuleId)||!Array.isArray(input.roles)||!input.roles.length||input.roles.length>2||new Set(input.roles).size!==input.roles.length||input.roles.some(r=>!COVERAGE_ROLES.includes(r)))fail();currentText(project,input.objectiveModuleId);
   const old=data.links.find(l=>l.activityId===change.activityId&&l.objectiveModuleId===input.objectiveModuleId),rationale=value(input.rationale,'a private objective-link rationale',true);
   proposals.set(input.objectiveModuleId,{...structuredClone(old),activityId:change.activityId,objectiveModuleId:input.objectiveModuleId,roles:COVERAGE_ROLES.filter(r=>input.roles.includes(r)),reviewedObjective:snapshot(project,input.objectiveModuleId,old,input),rationale});
  }
  copy.links=data.links.flatMap(l=>{if(l.activityId!==change.activityId)return [structuredClone(l)];const proposal=proposals.get(l.objectiveModuleId);proposals.delete(l.objectiveModuleId);return proposal?[proposal]:[];});copy.links.push(...proposals.values());
 }else if(change.kind==='exception'){
  if(!keys(change,change.remove===true?['kind','objectiveModuleId','role','remove']:['kind','objectiveModuleId','role','remove','reviewedObjective','note','reviewCurrent']))fail();
  currentText(project,change.objectiveModuleId);if(!COVERAGE_ROLES.includes(change.role)||typeof change.remove!=='boolean')fail();const index=data.exceptions.findIndex(e=>e.objectiveModuleId===change.objectiveModuleId&&e.role===change.role),old=data.exceptions[index];
  if(change.remove){if(index>=0)copy.exceptions.splice(index,1);}
  else{if(typeof change.note!=='string')fail();const entry={...structuredClone(old),objectiveModuleId:change.objectiveModuleId,role:change.role,reviewedObjective:snapshot(project,change.objectiveModuleId,old,change),note:value(change.note,'a private gap exception note',true)};if(index<0)copy.exceptions.push(entry);else copy.exceptions[index]=entry;}
 }else fail();
 if(!same(data,copy)){next.objectiveCoverage=copy;next.revision++;read(next);}else if(!absent)next.objectiveCoverage=structuredClone(data);
 return next;
}
export function coverageReport(project){
 const parsed=objectiveCoverage(project);if(!parsed.supported)return parsed;
 const activities=new Map(project.modules.flatMap(m=>m.activities.map(a=>[a.id,{id:a.id,title:a.title,moduleTitle:m.title}]))),rows=project.modules.map(m=>{
  const objective=objectiveText(m),invalid=unsupportedObjective(m),blank=!invalid&&!objective.trim(),roles=Object.fromEntries(COVERAGE_ROLES.map(role=>{
   const links=parsed.data.links.filter(l=>l.objectiveModuleId===m.id&&l.roles.includes(role)),current=links.filter(l=>!blank&&!invalid&&l.reviewedObjective===objective),stale=links.filter(l=>!invalid&&(blank||l.reviewedObjective!==objective)),unavailable=links.filter(()=>invalid),exception=parsed.data.exceptions.find(e=>e.objectiveModuleId===m.id&&e.role===role),exceptionCurrent=!!exception&&!blank&&!invalid&&exception.reviewedObjective===objective;
   return [role,{current:current.map(l=>({...structuredClone(l),activity:activities.get(l.activityId)})),stale:structuredClone(stale),unavailable:structuredClone(unavailable),gap:invalid?null:!current.length,exception:exception?structuredClone(exception):null,exceptionCurrent,acknowledgedGap:!invalid&&!current.length&&exceptionCurrent,conflict:!!current.length&&exceptionCurrent}];
  }));return {moduleId:m.id,title:m.title,objective,blank,invalid,roles};
 });return {supported:true,present:parsed.present,rows,scope:'Author-declared design mapping, not proof of learning or measurement validity.'};
}
