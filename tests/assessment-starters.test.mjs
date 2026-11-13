import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import {ASSESSMENT_STARTERS,assessmentStarter} from '../src/assessment-starters.mjs';
import {ACTIVITY_DRAFT_FIELDS,rawControl,wouldReplaceRaw,sameRaw} from '../src/activity-draft.mjs';
import {libraryPlaceholders,libraryFields,emptyLibrary,saveLibraryCopy,serializeLibrary,readLibrary} from '../src/library.mjs';
import {course,addModule,addActivity} from '../src/model.mjs';
import {fixture,makeHost} from './helpers/mock-office.mjs';
import {frontHost} from './helpers/frontmatter-fixture.mjs';
import {wrenchCourse} from '../src/model.mjs';
import {dom,settle} from './helpers/app-dom.mjs';
import {worksheetBlob} from '../src/worksheet.mjs';
import {agendaBlob} from '../src/agenda.mjs';
import {materialsBlob} from '../src/materials-output.mjs';
import {facilitatorBlob} from '../src/facilitator-output.mjs';
import {auditExport} from './helpers/export-contract.mjs';
import {auditFacilitator} from './helpers/facilitator-contract.mjs';
const defaults={title:'',icon:'Learning activity',minutes:'10',facilitator:'',participant:'',answers:'',lines:'5'};
const raw=fields=>({fields:Object.fromEntries(ACTIVITY_DRAFT_FIELDS.map(k=>[k,rawControl({value:fields[k]})])),materials:null});
let seq=0;
async function app(project=fixture()){
 const h=makeHost(project),ui=dom(),n=id=>ui.nodes.get(id);Object.assign(globalThis,{Office:h.Office,Word:h.Word,document:ui.document,location:{search:''},JSZip});
 globalThis.fetch=async()=>({ok:true,blob:async()=>new Blob(['mock cue']),arrayBuffer:async()=>new Uint8Array([137,80,78,71]).buffer});globalThis.FileReader=class{readAsDataURL(){this.result='data:image/png;base64,AAAA';queueMicrotask(()=>this.onload());}};
 await import(`../src/app.mjs?assessment=${++seq}`);await settle();n.busy=()=>ui.document.querySelector('main').attributes['aria-busy']==='true';return {h,ui,n};
}
function activityValues(n){return Object.fromEntries(ACTIVITY_DRAFT_FIELDS.map(k=>[k,n('activity-'+k).value]));}
async function catalog(n){const fields=fixture().modules[0].activities[0],data=saveLibraryCopy(emptyLibrary(),'Saved copy',[],{...Object.fromEntries(ACTIVITY_DRAFT_FIELDS.map(k=>[k,fields[k]])),materials:[{item:'Paper',quantity:1,unit:'sets',use:'consumable',notes:'Private row one'},{item:'Pens',quantity:2,unit:'sets',use:'reusable',notes:'Private row two'}]});const text=serializeLibrary(data);n('library-file').files=[{size:text.length,text:async()=>text}];n('library-file').onchange();await n('library-import').onclick();return data;}
async function load(n,kind,id){if(kind==='starter')n('assessment-starter').value=id||ASSESSMENT_STARTERS[0].id;await n(kind==='starter'?'assessment-load':'library-use').onclick();}
async function replace(n,kind,id){await load(n,kind,id);if(!n('activity-replacement').hidden)await n('activity-replace-confirm').onclick();}
function author(n,canary='AUTHORED'){
 for(const key of ['title','facilitator','participant','answers']){n('activity-'+key).value=n('activity-'+key).value.replace(/\[(?:insert|add|replace)\b[^\]]*\]/gi,`${canary}_${key}`);n('activity-'+key).oninput();}
}
function review(n){n('assessment-reviewed').checked=true;n('assessment-reviewed').onchange();}
async function insert(n){n('activity-form').onsubmit({preventDefault(){}});await settle();for(let i=0;i<1000&&n.busy();i++)await new Promise(resolve=>setImmediate(resolve));assert.equal(n.busy(),false);}

test('three original independent editable starters separate learner prompts from private suitability administration and keys without validity/scoring/scheduling claims',()=>{
 assert.equal(ASSESSMENT_STARTERS.length,3);for(const {id} of ASSESSMENT_STARTERS){const a=assessmentStarter(id),b=assessmentStarter(id);assert.deepEqual(libraryFields(a.fields),a.fields);assert.ok(libraryPlaceholders(a.fields).includes('participant'));assert.ok(libraryPlaceholders(a.fields).includes('facilitator'));assert.ok(libraryPlaceholders(a.fields).includes('answers'));assert.match(a.fields.facilitator,/suitability rationale/i);assert.doesNotMatch(a.fields.participant,/Private administration|Private key|Author suitability rationale/);assert.equal(a.origin.version,1);a.fields.facilitator='Changed';assert.notEqual(a.fields.facilitator,b.fields.facilitator);assert.equal(assessmentStarter(id).fields.facilitator,b.fields.facilitator);}
 assert.match(assessmentStarter('level1-feedback').fields.participant,/not a test of what you learned/);assert.match(assessmentStarter('level2-learning').fields.participant,/Success criterion/);assert.match(assessmentStarter('level3-application').fields.participant,/follow-up interval/);assert.match(assessmentStarter('level3-application').fields.facilitator,/does not schedule/);assert.throws(()=>assessmentStarter('unknown'));
});

test('raw replacement decisions preserve whitespace invalid-number validity origin and material distinctions without normalizing defaults or treating destination as overwritten',()=>{
 const baseline=raw(defaults),candidate=raw({...defaults,title:'Target',participant:'Task'});baseline.fields.title.invalid.valueMissing=baseline.fields.participant.invalid.valueMissing=true;assert.equal(wouldReplaceRaw(raw(defaults),baseline,candidate,null,{id:'target'}),false);
 for(const change of [c=>c.fields.title.value=' ',c=>c.fields.minutes.value='1.',c=>{c.fields.minutes.value='';c.fields.minutes.invalid.badInput=true;},c=>c.materials=[{quantity:rawControl({value:'',validity:{badInput:true}})}]]){const current=raw(defaults);change(current);const saved=structuredClone(current);assert.equal(wouldReplaceRaw(current,baseline,candidate,null,{id:'target'}),true);assert.deepEqual(current,saved);}
 assert.equal(wouldReplaceRaw(candidate,baseline,candidate,{id:'one'},{id:'two'}),true);assert.equal(wouldReplaceRaw(candidate,baseline,candidate,{id:'one'},{id:'one'}),false);assert.equal(wouldReplaceRaw(candidate,baseline,candidate,null,{id:'one'}),false);assert.equal(sameRaw({raw:'10'},{raw:'10.0'}),false);
});

test('all manual library and starter replacement combinations keep raw work on cancel and preserve destination on explicit discard/load',async()=>{
 for(const from of ['manual','library','starter'])for(const to of ['library','starter']){
  let p=addModule(fixture(),'Second module','Different objective').next;const {h,n}=await app(p);await catalog(n);n('activity-module').value=p.modules[1].id;
  if(from!=='manual')await replace(n,from);n('activity-title').value=`KEEP ${from}`;n('activity-title').oninput();n('activity-minutes').value='';n('activity-minutes').validity={badInput:true};n('activity-minutes').oninput();const before=activityValues(n),origin=n('activity-library-note').textContent,starter=n('assessment-origin').textContent,body=structuredClone(h.records);
  await load(n,to,'level2-learning');assert.equal(n('activity-replacement').hidden,false);assert.deepEqual(activityValues(n),before);assert.equal(n('activity-minutes').validity.badInput,true);assert.equal(n('activity-library-note').textContent,origin);assert.equal(n('assessment-origin').textContent,starter);await n('activity-replace-cancel').onclick();assert.deepEqual(activityValues(n),before);assert.equal(n('activity-replacement').hidden,true);
  await load(n,to,'level2-learning');await n('activity-replace-confirm').onclick();assert.notEqual(n('activity-title').value,before.title);assert.equal(n('activity-module').value,p.modules[1].id);assert.equal(n('activity-replacement').hidden,true);assert.equal(h.saves,0);assert.deepEqual(h.records,body);
 }
});

test('pristine defaults and destination-only changes load directly while identical manual/library values do not need a discard prompt',async()=>{
 const p=addModule(fixture(),'Second','Objective').next,{h,n}=await app(p);await catalog(n);n('activity-module').value=p.modules[1].id;await load(n,'starter');assert.equal(n('activity-replacement').hidden,true);assert.match(n('assessment-origin').textContent,/Level 1/);assert.equal(n('activity-module').value,p.modules[1].id);
 const other=await app(p);await catalog(other.n);const fields=fixture().modules[0].activities[0];for(const key of ACTIVITY_DRAFT_FIELDS)other.n('activity-'+key).value=String(fields[key]);await load(other.n,'library');assert.equal(other.n('activity-replacement').hidden,false,'Candidate adds saved materials, so current manually authored work would change');await other.n('activity-replace-cancel').onclick();assert.equal(h.saves,0);
});

test('pending intents reject newer raw edits changed selection renamed or removed sources and unknown starters without overwriting current origin or review',async()=>{
 for(const change of ['raw','starter-selection','rename','remove','unknown']){const {h,n}=await app();await catalog(n);await replace(n,'starter','level2-learning');author(n);review(n);assert.equal(n('assessment-reviewed').checked,true);const before=activityValues(n),origin=n('assessment-origin').textContent;
  if(change==='unknown'){n('assessment-starter').value='unknown';await n('assessment-load').onclick();assert.deepEqual(activityValues(n),before);assert.equal(n('assessment-reviewed').checked,true);assert.equal(n('assessment-origin').textContent,origin);continue;}
  await load(n,change==='starter-selection'?'starter':'library','level1-feedback');assert.equal(n('activity-replacement').hidden,false);
  if(change==='raw'){n('activity-facilitator').value+=' New raw work';n('activity-facilitator').oninput();}else if(change==='starter-selection')n('assessment-starter').value='level3-application';else if(change==='rename'){await n('library-manage-load').onclick();n('library-edit-name').value='Renamed source';n('library-manage-form').onsubmit({preventDefault(){}});await settle();}else await n('library-remove').onclick();
  const newer=activityValues(n);await n('activity-replace-confirm').onclick();assert.deepEqual(activityValues(n),newer);assert.equal(n('assessment-origin').textContent,origin);assert.equal(h.saves,0);assert.match(n('error').textContent,/changed|Select a saved/);if(change!=='raw')assert.equal(n('assessment-reviewed').checked,true);
 }
});

test('live copied material controls retain unfinished numeric validity and raw text through refusal review rendering and removal of another row',async()=>{
 const {h,n}=await app();await catalog(n);await load(n,'library');const root=n('activity-copied-materials-list'),first=root.children[0],quantity=first.children.find(x=>x.id?.endsWith('-quantity')),notes=first.children.find(x=>x.id?.endsWith('-notes'));quantity.value='';quantity.validity={badInput:true};quantity.oninput();notes.value='UNFINISHED PRIVATE MATERIAL';notes.oninput();await load(n,'starter','level3-application');assert.equal(root.children[0],first);assert.equal(first.children.find(x=>x.id?.endsWith('-quantity')),quantity);assert.equal(quantity.value,'');assert.equal(quantity.validity.badInput,true);assert.equal(notes.value,'UNFINISHED PRIVATE MATERIAL');await n('activity-replace-cancel').onclick();assert.equal(root.children[0],first);
  await root.children[1].children.at(-1).onclick();assert.equal(root.children.length,1);assert.equal(root.children[0],first);assert.equal(quantity.validity.badInput,true);assert.equal(notes.value,'UNFINISHED PRIVATE MATERIAL');await load(n,'starter');notes.value+=' changed after intent';notes.oninput();await n('activity-replace-confirm').onclick();assert.equal(root.children[0],first);assert.equal(quantity.validity.badInput,true);assert.match(notes.value,/changed after intent/);assert.equal(h.saves,0);
});

test('starter prompts explicit review field/destination changes and eventless edits block bypass before ordinary fresh-ID insertion without false library provenance',async()=>{
 const p=addModule(fixture(),'Target second module','Objective').next,{h,n}=await app(p);const fresh=[];
 for(const {id} of ASSESSMENT_STARTERS){n('activity-module').value=p.modules[0].id;n('activity-module').onchange();await replace(n,'starter',id);assert.equal(n('activity-insert').disabled,true);n('assessment-reviewed').checked=true;await insert(n);assert.equal(h.saves,fresh.length*2);assert.match(n('error').textContent,/unfinished prompts/);author(n,'AUTHORED_'+id);assert.equal(n('activity-insert').disabled,true);n('assessment-reviewed').checked=true;await insert(n);assert.match(n('error').textContent,/Review this exact/);review(n);assert.equal(n('activity-insert').disabled,false);n('activity-module').value=p.modules[1].id;n('activity-module').onchange();assert.equal(n('assessment-reviewed').checked,false);review(n);n('activity-facilitator').value+=' EVENTLESS CHANGE';await insert(n);assert.match(n('error').textContent,/Review this exact/);review(n);await insert(n);const saved=h.durable['openstage.guide.project.v1'].modules[1].activities.at(-1);assert.ok(saved);assert.equal(saved.librarySource,undefined);assert.equal(saved.starterOrigin,undefined);assert.match(saved.facilitator,/AUTHORED/);assert.equal(n('activity-title').value,'');assert.equal(n('activity-minutes').value,'10');assert.equal(n('activity-lines').value,'5');assert.equal(n('activity-module').value,p.modules[1].id);assert.equal(n('assessment-review').hidden,true);fresh.push(saved.id);}
 assert.equal(new Set(fresh).size,3);assert.equal(h.durable['openstage.guide.project.v1'].modules[0].activities.length,1);
});

test('newly edited Activity work and same-value edit generations survive asynchronous success uncertain persistence and repeated recovery without replay',async()=>{
 for(const mode of ['success','same-value-generation','final-not-saved','final-saved']){const {h,n}=await app();await replace(n,'starter','level2-learning');author(n);review(n);const submitted=activityValues(n),reset=n('activity-form').resetCount||0;let changed=false;h.hooks.beforeRefresh=()=>{if(changed)return;changed=true;n('activity-title').value=mode==='same-value-generation'?submitted.title:'NEW RAW TITLE';n('activity-title').oninput();};if(mode.startsWith('final'))h.saveFailures.set(2,{persist:mode==='final-saved'});await insert(n);assert.equal(h.durable['openstage.guide.project.v1']?.modules[0].activities.at(-1)?.title===submitted.title||h.durable['openstage.guide.pending.v1']?.project.modules[0].activities.at(-1).title===submitted.title,true);assert.equal(n('activity-title').value,mode==='same-value-generation'?submitted.title:'NEW RAW TITLE');assert.equal(n('activity-form').resetCount||0,reset);assert.match(n('assessment-origin').textContent,/Level 2/);const count=h.records.length;h.hooks.beforeRefresh=undefined;h.saveFailures.clear();if(mode.startsWith('final')){await n('recover').onclick();await n('recover').onclick();assert.equal(h.records.length,count);assert.equal(n('activity-title').value,'NEW RAW TITLE');}assert.equal(h.durable['openstage.guide.project.v1'].modules[0].activities.at(-1).librarySource,undefined);}
});

test('starter insertion retains profile5 and uncertain full-field recovery on an isolated Wrench host',async()=>{
 const p=addModule(wrenchCourse('Wrench assessment'), 'Lesson','Objective').next,{h,n}=await app(p);frontHost(h);await n('recover').onclick();await replace(n,'starter','level1-feedback');author(n);review(n);h.saveFailures.set(h.saves+2,{persist:false});await insert(n);assert.equal(h.durable['openstage.guide.pending.v1'].schema,5);const count=h.records.length;h.saveFailures.clear();await n('recover').onclick();await n('recover').onclick();assert.equal(h.records.length,count);assert.equal(h.durable['openstage.guide.pending.v1'].schema,5);assert.equal(h.durable['openstage.guide.project.v1'].modules[0].activities.length,1);assert.equal(n('activity-title').value,assessmentStarter('level1-feedback').fields.title);
});

test('manual and genuine library insertions retain eventless newer raw fields and live invalid material buffers after success',async()=>{
 for(const kind of ['manual','library']){const {h,n}=await app();if(kind==='library'){await catalog(n);await load(n,'library');}else{n('activity-title').value='Submitted manual task';n('activity-participant').value='Submitted manual instructions';}const submitted=activityValues(n),root=n('activity-copied-materials-list'),quantity=root.children[0]?.children.find(x=>x.id?.endsWith('-quantity'));let once=false;
  h.hooks.beforeRefresh=()=>{if(once)return;once=true;if(quantity){quantity.value='';quantity.validity={badInput:true};}else n('activity-title').value='NEW EVENTLESS MANUAL TITLE';};
  await insert(n);const saved=h.durable['openstage.guide.project.v1'].modules[0].activities.at(-1);assert.equal(saved.title,submitted.title);assert.equal(h.paragraphs.length,9);assert.equal(n('activity-form').resetCount,undefined);if(quantity){assert.equal(root.children[0].children.find(x=>x.id?.endsWith('-quantity')),quantity);assert.equal(quantity.value,'');assert.equal(quantity.validity.badInput,true);assert.match(n('activity-library-note').textContent,/Independent copy/);assert.equal(saved.materials[0].quantity,1);assert.ok(saved.librarySource);}else{assert.equal(n('activity-title').value,'NEW EVENTLESS MANUAL TITLE');assert.equal(saved.librarySource,undefined);}
 }
});

test('pending replacement preserves a newly chosen destination because the candidate does not overwrite it',async()=>{
 const p=addModule(fixture(),'Destination second','Objective').next,{h,n}=await app(p);n('activity-title').value='Keep me';n('activity-module').value=p.modules[0].id;await load(n,'starter');assert.equal(n('activity-replacement').hidden,false);n('activity-module').value=p.modules[1].id;n('activity-module').onchange();await n('activity-replace-confirm').onclick();assert.equal(n('activity-replacement').hidden,true);assert.equal(n('activity-module').value,p.modules[1].id);assert.equal(n('activity-title').value,assessmentStarter('level1-feedback').fields.title);assert.equal(h.saves,0);
});

test('authored starter catalog copies round-trip independently and public/private output choices isolate suitability administration and keys',async()=>{
 let p=addModule(course('Assessment outputs'),'Actual learning','Apply an authored task').next;const privateCanaries=[];
 for(const {id} of ASSESSMENT_STARTERS){const s=assessmentStarter(id);s.fields.participant=s.fields.participant.replace(/\[insert[^\]]*\]/g,'Authored learner task');s.fields.facilitator='PRIVATE_ADMIN_'+id+' PRIVATE_RATIONALE_'+id;s.fields.answers='PRIVATE_KEY_'+id;privateCanaries.push(s.fields.facilitator,s.fields.answers);const library=saveLibraryCopy(emptyLibrary(),id,[],s.fields),round=readLibrary(serializeLibrary(library));assert.deepEqual(round.entries[0].fields,s.fields);const a=addActivity(p,p.modules[0].id,s.fields);p=a.next;const saved=structuredClone(round);p.modules[0].activities.at(-1).facilitator+=' Changed inserted copy';assert.deepEqual(round,saved);}
 for(const kind of ['workbook','agenda','materials']){const blob=kind==='workbook'?await worksheetBlob(p,JSZip):kind==='agenda'?await agendaBlob(p,'09:00',JSZip):await materialsBlob(p,undefined,JSZip);await auditExport(blob,{project:p,privateCanaries,kind});}
 for(const answers of [false,true])for(const participant of [false,true]){const {text}=await auditFacilitator(await facilitatorBlob(p,{answers,participant},'2026-10-26T07:30:00.000Z',JSZip));for(const {id} of ASSESSMENT_STARTERS){assert.ok(text.includes('PRIVATE_ADMIN_'+id));assert.ok(text.includes('PRIVATE_RATIONALE_'+id));assert.equal(text.includes('PRIVATE_KEY_'+id),answers);}assert.equal(text.includes('Authored learner task'),participant);}
});
