import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import {fixture,makeHost} from './helpers/mock-office.mjs';
import {dom,settle} from './helpers/app-dom.mjs';
import {addActivity} from '../src/model.mjs';
import {editDeliveryPlan} from '../src/delivery-plans.mjs';
import {KEY} from '../src/document-store.mjs';
import {WORKBOOK_FAMILIES} from '../src/workbook-presentation.mjs';
import {worksheetFiles} from '../src/worksheet.mjs';
import {planWorkbookFiles} from '../src/selected-plan-outputs.mjs';

let sequence=0;
function saved(){
 let p=fixture();p=addActivity(p,p.modules[0].id,{title:'Second selected activity',minutes:9,participant:'Learner second\nAnother line 😊',facilitator:'PRIVATE TEACHING',answers:'PRIVATE ANSWER',lines:2}).next;
 p=addActivity(p,p.modules[0].id,{title:'Omitted activity',minutes:5,participant:'OMITTED LEARNER',facilitator:'PRIVATE TEACHING',answers:'PRIVATE ANSWER',lines:0}).next;
 const groups=[{moduleId:p.modules[0].id,activityIds:[p.modules[0].activities[1].id]}];
 p=editDeliveryPlan(p,{kind:'create',name:'Short session',targetMinutes:null,groups}).next;
 return p;
}
async function app(p=saved()){
 const h=makeHost(p),ui=dom(),n=id=>ui.nodes.get(id);
 Object.assign(globalThis,{Office:h.Office,Word:h.Word,document:ui.document,location:{search:''},JSZip});
 await import(`../src/app.mjs?family-ui=${++sequence}`);await settle();await ui.routes.find(button=>button.dataset.view==='output').onclick();await settle();return {h,n,p,id:p.deliveryPlans.plans[0].id};
}
async function click(n,id){await n(id).onclick();await settle();}
async function reviewed(n,{family='clearline',planId=''}={}){
 n('workbook-plan').value=planId;n('workbook-plan').onchange();
 n('family-template').value=family;n('family-template').onchange();
 await click(n,'family-preview');
 assert.equal(n('error').hidden,true);
 assert.equal(n('family-reviewed').disabled,false);
 n('family-reviewed').checked=true;n('family-reviewed').onchange();
 assert.equal(n('family-apply').disabled,false);
}
async function downloads(action){
 const blobs=[],anchors=[],originalUrl=URL.createObjectURL,originalCreate=document.createElement,originalTimer=globalThis.setTimeout;
 URL.createObjectURL=blob=>{blobs.push(blob);return `blob:family-${blobs.length}`;};
 document.createElement=tag=>{const item=originalCreate(tag);if(tag==='a')anchors.push(item);return item;};
 globalThis.setTimeout=(fn,ms)=>ms===60000?0:originalTimer(fn,ms);
 try{await action();}finally{URL.createObjectURL=originalUrl;document.createElement=originalCreate;globalThis.setTimeout=originalTimer;}
 return {blobs,anchors};
}

test('separate design route shows all five families and preserves the legacy full/selected exports and raw drafts',async()=>{
 const {n,p,id,h}=await app(),legacyFull=worksheetFiles(p,'2026-10-26T10:00:00.000Z'),legacySelected=planWorkbookFiles(p,id,'2026-10-26T10:00:00.000Z');
 assert.deepEqual(n('family-template').children.map(x=>x.value),WORKBOOK_FAMILIES.map(x=>x.id));
 n('activity-title').value='RAW ACTIVITY DRAFT';n('plan-name').value='RAW PLAN DRAFT';
 await reviewed(n,{family:'clearline'});
 const view=n('family-preview-content').textContent;
 assert.match(view,/SOURCE PREVIEW ONLY/);assert.match(view,/Full course/);assert.match(view,/Applied text and response roles/);
 assert.doesNotMatch(view,/PRIVATE TEACHING|PRIVATE ANSWER/);
 assert.equal(n('download').disabled,true,'legacy route retains its own preview/review');
 n('workbook-plan').value=id;n('workbook-plan').onchange();assert.equal(n('family-apply').disabled,true);
 await reviewed(n,{family:'practice-lab',planId:id});
 assert.match(n('family-preview-content').textContent,/Short session/);
 assert.match(n('family-preview-content').textContent,/Second selected activity/);
 assert.doesNotMatch(n('family-preview-content').textContent,/OMITTED LEARNER/);
 assert.equal(n('activity-title').value,'RAW ACTIVITY DRAFT');assert.equal(n('plan-name').value,'RAW PLAN DRAFT');
 assert.equal(h.saves,0);assert.deepEqual(worksheetFiles(p,'2026-10-26T10:00:00.000Z'),legacyFull);
 assert.deepEqual(planWorkbookFiles(p,id,'2026-10-26T10:00:00.000Z'),legacySelected);
});

test('healthy repeated Apply creates unique new DOCX names with the same reviewed generation and exact selected scope',async()=>{
 const {n,id,h}=await app();await reviewed(n,{family:'reflection-journal',planId:id});
 const result=await downloads(async()=>{await click(n,'family-apply');await click(n,'family-apply');});
 assert.equal(result.blobs.length,2);assert.notEqual(result.anchors[0].download,result.anchors[1].download);
 assert.match(result.anchors[0].download,/Reflection Journal Workbook.*\.docx$/);
 const zips=await Promise.all(result.blobs.map(async blob=>JSZip.loadAsync(await blob.arrayBuffer())));
 const first=await zips[0].file('word/document.xml').async('string'),second=await zips[1].file('word/document.xml').async('string');
 assert.equal(first,second);assert.match(first,/Second selected activity/);assert.doesNotMatch(first,/OMITTED LEARNER|PRIVATE TEACHING|PRIVATE ANSWER/);
 assert.equal(await zips[0].file('docProps/core.xml').async('string'),await zips[1].file('docProps/core.xml').async('string'));
 assert.equal(h.saves,0);assert.equal(n('family-apply').disabled,false);
});

test('all 25 ordered design transitions keep one shared learner model and authored overrides',async()=>{
 const {n,h,id,p}=await app(),before=structuredClone(p);
 n('activity-title').value='RAW ACTIVITY';n('family-brand-organizationName').value='Example College';n('family-brand-bodyPt').value='13';
 for(const from of WORKBOOK_FAMILIES)for(const to of WORKBOOK_FAMILIES){
  n('family-template').value=from.id;n('family-template').onchange();await click(n,'family-preview');
  n('family-reviewed').checked=true;n('family-reviewed').onchange();
  n('family-template').value=to.id;n('family-template').onchange();assert.equal(n('family-reviewed').checked,false);
  await click(n,'family-preview');const view=n('family-preview-content').textContent;
  assert.ok(view.includes(`Design: ${to.name}`));assert.equal((view.match(/MODULE 1:/g)||[]).length,1);
  assert.equal(n('family-brand-organizationName').value,'Example College');assert.equal(n('family-brand-bodyPt').value,'13');
  assert.match(view,/bodyPt 13/);
 }
 n('workbook-plan').value=id;n('workbook-plan').onchange();assert.equal(n('family-brand-bodyPt').value,'13');
 assert.deepEqual(p,before);assert.equal(h.saves,0);assert.equal(n('activity-title').value,'RAW ACTIVITY');
});

test('invalid typed branding is local, keeps raw text, and safe alternatives require an explicit action',async()=>{
 const {n}=await app();n('activity-title').value='RAW AUTHOR DRAFT';
 await click(n,'preview');assert.equal(n('download').disabled,false);
 n('family-brand-bodyPt').value='12abc';n('family-brand-bodyPt').oninput();await click(n,'family-preview');
 assert.match(n('family-brand-status').textContent,/plain decimal/);assert.equal(n('family-brand-bodyPt').value,'12abc');
 assert.equal(n('error').hidden,true);assert.equal(n('download').disabled,false);assert.equal(n('family-apply').disabled,true);
 n('family-brand-bodyPt').value='';n('family-brand-bodyPt').oninput();
 n('family-brand-accent').value='#246A82';n('family-brand-accent').oninput();
 n('family-brand-text').value='#EEEEEE';n('family-brand-text').oninput();await click(n,'family-preview');
 assert.match(n('family-brand-status').textContent,/contrast|Suggested/i);assert.equal(n('family-brand-text').value,'#EEEEEE');
 assert.equal(n('family-use-safe-alternative').hidden,false);assert.equal(n('family-apply').disabled,true);
 await click(n,'family-use-safe-alternative');assert.equal(n('family-brand-text').value,'#000000');
 assert.equal(n('family-brand-accent').value,'#246A82');assert.equal(n('family-reviewed').checked,false);
 await click(n,'family-preview');assert.equal(n('family-reviewed').disabled,false);
 assert.equal(n('activity-title').value,'RAW AUTHOR DRAFT');
});

test('white profile, native invalid buffer, unsupported font and grayscale are reviewed without rewriting inputs',async()=>{
 const {n}=await app();
 n('family-brand-accent').value='#246A82';n('family-brand-accent').oninput();
 n('family-brand-background').value='#202020';n('family-brand-background').oninput();n('family-brand-text').value='#FFFFFF';n('family-brand-text').oninput();
 await click(n,'family-preview');assert.match(n('family-brand-status').textContent,/white page background/);assert.equal(n('family-brand-background').value,'#202020');
 assert.equal(n('family-use-safe-alternative').hidden,false);await click(n,'family-use-safe-alternative');
 assert.equal(n('family-brand-background').value,'#FFFFFF');assert.equal(n('family-brand-text').value,'#1A1A1A');
 assert.equal(n('family-brand-accent').value,'#246A82');assert.equal(n('family-reviewed').checked,false);
 n('family-brand-background').value='';n('family-brand-background').oninput();n('family-brand-text').value='';n('family-brand-text').oninput();
 n('family-brand-headingFont').value='Invalid Font';n('family-brand-headingFont').onchange();await click(n,'family-preview');
 assert.match(n('family-brand-status').textContent,/headingFont/);assert.equal(n('family-brand-headingFont').value,'Invalid Font');
 n('family-brand-headingFont').value='';n('family-brand-headingFont').onchange();
 n('family-brand-bodyPt').validity={badInput:true};await click(n,'family-preview');assert.match(n('family-brand-status').textContent,/unfinished or invalid/);
 n('family-brand-bodyPt').value='';n('family-brand-bodyPt').oninput();
 n('family-brand-printMode').value='grayscale';n('family-brand-printMode').onchange();await click(n,'family-preview');
 assert.match(n('family-preview-content').textContent,/grayscale/);assert.equal(n('family-reviewed').disabled,false);
});

test('grayscale contrast witness keeps requested blue until an explicit safe suggestion and fresh review',async()=>{
 const {n}=await app();n('family-brand-text').value='#0072F0';n('family-brand-text').oninput();
 n('family-brand-printMode').value='grayscale';n('family-brand-printMode').onchange();
 await click(n,'family-preview');assert.equal(n('family-apply').disabled,true);
 assert.match(n('family-brand-status').textContent,/grayscale conversion/);
 assert.equal(n('family-brand-text').value,'#0072F0');assert.equal(n('family-use-safe-alternative').hidden,false);
 await click(n,'family-use-safe-alternative');assert.equal(n('family-brand-text').value,'#000000');
 assert.equal(n('family-brand-printMode').value,'grayscale');assert.equal(n('family-reviewed').checked,false);
 await click(n,'family-preview');assert.equal(n('family-reviewed').disabled,false);
});

test('full-course problems and selected-plan changes stay local to the new-copy route',async()=>{
 const bad=saved();bad.modules[0].activities[0].participant='   ';
 const broken=await app(bad);broken.n('activity-title').value='RAW ACTIVITY';
 await click(broken.n,'family-preview');assert.equal(broken.n('family-apply').disabled,true);
 assert.match(broken.n('family-brand-status').textContent,/Needs attention|participant|resolve/i);
 assert.equal(broken.n('error').hidden,true);assert.equal(broken.n('activity-title').value,'RAW ACTIVITY');
 const {n,h,id}=await app();n('activity-title').value='RAW ACTIVITY';
 n('workbook-plan').value=id;n('workbook-plan').onchange();await click(n,'family-preview');assert.equal(n('family-reviewed').disabled,false);
 n('family-reviewed').checked=true;n('family-reviewed').onchange();
 class PlanChanges extends JSZip{async generateAsync(options){h.durable[KEY].deliveryPlans.plans[0].name='Changed selected plan';return super.generateAsync(options);}}
 globalThis.JSZip=PlanChanges;
 const result=await downloads(()=>click(n,'family-apply'));
 assert.equal(result.blobs.length,0);assert.match(n('error').textContent,/changed|refresh/i);
 assert.equal(n('activity-title').value,'RAW ACTIVITY');
});

test('stale source, controls, preview visibility, consent and ZIP failures never offer a new file',async()=>{
 for(const drift of ['source','brand','family','language','language-away-back','scope','scope-away-back','visibility','section-hidden','review-hidden','brand-hidden','consent','zip-error','zip-tamper']){
  const {n,h,id}=await app();n('activity-title').value='RAW DRAFT';await reviewed(n);
  class ChangingZip extends JSZip{async generateAsync(options){
    if(drift==='source')h.durable[KEY].title='Changed source during packaging';
    if(drift==='brand')n('family-brand-bodyPt').value='12abc';
    if(drift==='family')n('family-template').value='editorial-study';
    if(drift==='language')n('family-base-language').value='fr-FR';
    if(drift==='language-away-back'){n('family-base-language').value='fr-FR';n('family-base-language').onchange();n('family-base-language').value='en-US';n('family-base-language').onchange();}
    if(drift==='scope')n('workbook-plan').value=id;
    if(drift==='scope-away-back'){n('workbook-plan').value=id;n('workbook-plan').onchange();n('workbook-plan').value='';n('workbook-plan').onchange();}
    if(drift==='visibility')n('family-preview-content').hidden=true;
    if(drift==='section-hidden')n('workbook-family-section').hidden=true;
    if(drift==='review-hidden')n('family-review-label').hidden=true;
    if(drift==='brand-hidden')n('family-brand-controls').hidden=true;
    if(drift==='consent')n('family-reviewed').checked=false;
    if(drift==='zip-error')throw Error('Synthetic ZIP failure');
    if(drift==='zip-tamper')this.file('word/document.xml','changed',{createFolders:false});
    return super.generateAsync(options);
  }}
  globalThis.JSZip=ChangingZip;
  const result=await downloads(()=>click(n,'family-apply'));
  assert.equal(result.blobs.length,0,drift);assert.equal(result.anchors.length,0,drift);
  assert.equal(n('family-apply').disabled,true,drift);assert.equal(n('activity-title').value,'RAW DRAFT',drift);
  assert.match(n('error').textContent,/changed|refresh|failure|member/i,drift);
  assert.equal(h.paragraphs.length,0,drift);
 }
});

test('broken or empty selected plans refuse only the new design route and keep authoring drafts',async()=>{
 for(const defect of ['missing','empty']){
  const {n,h,id}=await app();n('activity-title').value='RAW AUTHORING';
  h.durable[KEY].deliveryPlans.plans[0].groups=defect==='missing'?[{moduleId:h.durable[KEY].modules[0].id,activityIds:['missing']}]:[];
  n('workbook-plan').value=id;n('workbook-plan').onchange();await click(n,'family-preview');
  assert.equal(n('family-apply').disabled,true);assert.match(n('family-brand-status').textContent,/missing|No activities|unique|repair|selected/i);
  assert.equal(n('activity-title').value,'RAW AUTHORING');assert.equal(n('error').hidden,true);
  n('workbook-plan').value='';n('workbook-plan').onchange();await click(n,'family-preview');assert.equal(n('family-reviewed').disabled,false);
 }
});

test('authored base language is reviewed in the UI and exact new DOCX while invalid tags stay local',async()=>{
 const {n}=await app();
 assert.equal(n('family-base-language').value,'en-US');
 n('family-base-language').value='fr-CA';n('family-base-language').onchange();
 await reviewed(n);assert.match(n('family-preview-content').textContent,/Base document language: fr-CA/);
 assert.match(n('family-preview-content').textContent,/no authored per-run language tagging/);
 const result=await downloads(()=>click(n,'family-apply'));assert.equal(result.blobs.length,1);
 const zip=await JSZip.loadAsync(await result.blobs[0].arrayBuffer());
 assert.match(await zip.file('word/styles.xml').async('string'),/w:lang w:val="fr-CA"/);
 assert.match(await zip.file('docProps/core.xml').async('string'),/<dc:language>fr-CA<\/dc:language>/);
 n('family-base-language').value='xx-XX';n('family-base-language').onchange();await click(n,'family-preview');
 assert.equal(n('family-apply').disabled,true);assert.match(n('family-brand-status').textContent,/base language is unsupported/);
 assert.equal(n('error').hidden,true);
});
