import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import {createHash} from 'node:crypto';
import {fixture,makeHost} from './helpers/mock-office.mjs';
import {makeDOM} from './helpers/workbook-design-dom.mjs';
import {syntheticLogo,logoChunk} from './helpers/logo-fixture.mjs';
import {editDeliveryPlan} from '../src/delivery-plans.mjs';
import {WORKBOOK_FAMILIES} from '../src/workbook-presentation.mjs';
import {settle} from './helpers/app-dom.mjs';
import {parseWorkbookDesignFile} from '../src/workbook-design-file.mjs';
let sequence=0;
async function app(project=fixture()){
 const h=makeHost(project),ui=makeDOM(fs.readFileSync(new URL('../public/taskpane.html',import.meta.url),'utf8')),n=id=>ui.document.getElementById(id);let fetches=0,reads=0,writes=0;
 Object.assign(globalThis,{Office:h.Office,Word:h.Word,document:ui.document,location:{search:''},fetch:async()=>{fetches++;throw Error('No design file network');},localStorage:{getItem(){reads++;return null;},setItem(){writes++;throw Error('No design file storage write');}}});
 await import(`../src/app.mjs?design-files=${++sequence}`);await settle();const go=async name=>{await ui.routes.find(b=>b.dataset.view===name).onclick();await settle();};await go('output');
 return {project,h,ui,n,go,f:id=>n('family-design-file-'+id),l:id=>n('family-logo-'+id),calls:()=>({fetches,reads,writes,saves:h.saves,runs:h.runs,refreshes:h.refreshes})};
}
const click=async(a,id)=>{await a.n(id).onclick();await settle();};
const set=(a,id,value,event='oninput')=>{a.n(id).value=value;a.n(id)[event]?.();};
async function offers(fn){const out=[],old=URL.createObjectURL,create=document.createElement,timer=globalThis.setTimeout;URL.createObjectURL=blob=>{out.push({blob,name:''});return 'blob:design-test';};document.createElement=tag=>{const e=create(tag);if(tag==='a')e.click=()=>{out.at(-1).name=e.download;};return e;};globalThis.setTimeout=(fn,ms,...args)=>ms===60000?0:timer(fn,ms,...args);try{await fn();}finally{URL.createObjectURL=old;document.createElement=create;globalThis.setTimeout=timer;}return out;}
async function saved(a){const files=await offers(()=>click(a,'family-design-file-save'));assert.equal(files.length,1,a.f('error').textContent);return {bytes:new Uint8Array(await files[0].blob.arrayBuffer()),name:files[0].name};}
async function selectLogo(a,bytes=syntheticLogo(),name='PRIVATE_LOGO_FILENAME.png'){a.l('file').files=[new File([bytes],name,{type:'image/png'})];a.l('file').value='C:\\fakepath\\'+name;await a.l('file').onchange();await settle();}
function author(a){set(a,'family-brand-organizationName','Synthetic organization');set(a,'family-front-cover-title','Public workbook title');a.n('family-front-include-cover').checked=true;a.n('family-front-include-cover').onchange();set(a,'family-front-copyright-noticeText','Copyright wording — authored ©');set(a,'family-research-rationale','Rationale\nSecond paragraph 😀');a.n('family-research-include').checked=true;a.n('family-research-include').onchange();a.n('family-toc-include').checked=true;a.n('family-toc-include').onchange();set(a,'family-logo-role','informative');set(a,'family-logo-alt','Geometric organization mark');set(a,'family-logo-width','1.5');a.l('include').checked=true;a.l('include').onchange();}
async function checked(a,bytes,{name='DESIGN.json',file}={}){const f=file||new File([bytes],name,{type:'application/json'});a.f('input').files=[f];a.f('input').value=name;a.f('input').onchange();await a.f('load').onclick();await settle();return f;}
function raw(a){return {family:a.n('family-template').value,language:a.n('family-base-language').value,brand:a.n('family-brand-organizationName').value,title:a.n('family-front-cover-title').value,research:a.n('family-research-rationale').value,width:a.l('width').value,alt:a.l('alt').value,include:a.l('include').checked,image:a.l('image').getAttribute('src'),input:a.l('file'),files:a.l('file').files,scope:a.n('workbook-plan').value,activity:a.n('activity-title').value,assessment:a.n('measure-item-prompt').value};}

test('explicit design files work without a course, are uniquely named and exclude course/assessment/response/preset/review state',async()=>{
 const a=await app(null),before=a.calls();set(a,'family-brand-organizationName','Org 😀');set(a,'family-brand-bodyPt','12.');set(a,'family-research-rationale','Unfinished research\n\tmeaning');a.n('activity-facilitator').value='PRIVATE_FACILITATOR_CANARY';a.n('measure-item-prompt').value='PRIVATE_ASSESSMENT_KEY_CANARY';a.n('family-preset-name').value='LOCAL_PRESET_NAME_CANARY';a.n('family-reviewed').checked=true;
 const one=await saved(a),two=await saved(a);assert.notEqual(one.name,two.name);assert.match(one.name,/PRIVATE.*\.json$/u);assert.deepEqual(one.bytes,two.bytes);const parsed=parseWorkbookDesignFile(one.bytes);assert.equal(parsed.payload.brand.bodyPt,'12.');const text=new TextDecoder().decode(one.bytes);assert.doesNotMatch(text,/PRIVATE_FACILITATOR|PRIVATE_ASSESSMENT|LOCAL_PRESET_NAME|reviewed|consent|projectId|planId/u);assert.deepEqual(a.calls(),before);
 set(a,'family-brand-organizationName','Current');const prior=raw(a);await checked(a,one.bytes);assert.equal(a.f('confirm').hidden,false,a.f('error').textContent);assert.equal(a.n('family-brand-organizationName').value,'Current');a.f('keep').onclick();assert.deepEqual(raw(a),prior);await checked(a,one.bytes);a.f('replace').onclick();assert.equal(a.f('error').hidden,true,a.f('error').textContent);assert.equal(a.n('family-brand-organizationName').value,'Org 😀');assert.equal(a.n('family-brand-bodyPt').value,'12.');assert.equal(a.n('family-reviewed').checked,false);assert.equal(a.n('family-apply').disabled,true);assert.deepEqual(a.calls(),before);
});

test('PNG bundle roundtrip creates a fresh owned image and swaps picker while preserving exact raw design and unrelated work',async()=>{
 const a=await app();await selectLogo(a);author(a);const original=await saved(a),originalData=parseWorkbookDesignFile(original.bytes),oldInput=a.l('file'),oldFiles=oldInput.files,oldImage=a.l('image').getAttribute('src');set(a,'family-brand-organizationName','NEW CURRENT');a.n('activity-title').value='UNRELATED ACTIVITY';a.n('measure-item-prompt').value='UNRELATED ASSESSMENT';a.n('family-preset-name').value='PRESET DRAFT KEPT';const before=a.calls(),source=structuredClone(a.project);
 await checked(a,original.bytes);assert.equal(a.f('confirm').hidden,false,a.f('error').textContent);assert.equal(a.l('file'),oldInput);assert.equal(a.l('file').files,oldFiles);assert.equal(a.l('image').getAttribute('src'),oldImage);assert.equal(a.f('image').hidden,false);assert.match(a.f('output').textContent,/SHA-256/u);a.f('replace').onclick();assert.equal(a.f('error').hidden,true,a.f('error').textContent);assert.notEqual(a.l('file'),oldInput);assert.equal(oldInput.files,oldFiles);assert.equal(a.l('file').files.length,0);assert.equal(a.l('image').getAttribute('src'),oldImage);assert.equal(a.n('family-brand-organizationName').value,'Synthetic organization');assert.equal(a.n('activity-title').value,'UNRELATED ACTIVITY');assert.equal(a.n('measure-item-prompt').value,'UNRELATED ASSESSMENT');assert.equal(a.n('family-preset-name').value,'PRESET DRAFT KEPT');assert.deepEqual(a.project,source);assert.deepEqual(a.calls(),before);assert.deepEqual(parseWorkbookDesignFile((await saved(a)).bytes),originalData);
});

test('Save refuses failed or pending logo choices without substituting lastValid; Load can explicitly recover an invalid current selection',async()=>{
 const a=await app();await selectLogo(a);author(a);const good=await saved(a);await selectLogo(a,new Uint8Array([1,2,3]),'BAD.png');const prior=raw(a);assert.equal((await offers(()=>click(a,'family-design-file-save'))).length,0);assert.match(a.f('error').textContent,/not ready/u);assert.deepEqual(raw(a),prior);await checked(a,good.bytes);assert.equal(a.f('confirm').hidden,false,a.f('error').textContent);a.f('keep').onclick();assert.deepEqual(raw(a),prior);await checked(a,good.bytes);a.f('replace').onclick();assert.equal(a.f('error').hidden,true,a.f('error').textContent);assert.equal(a.l('image').hidden,false);assert.equal(a.l('alt').value,'Geometric organization mark');
});

test('malformed Unicode/version/oversize files and canceled or repeated selections keep all current drafts',async()=>{
 const a=await app();set(a,'family-front-cover-title','KEEP TITLE');set(a,'family-brand-bodyPt','1.');const prior=raw(a),base=await saved(a),data=parseWorkbookDesignFile(base.bytes);
 for(const bytes of [new Uint8Array([0xff]),new TextEncoder().encode('{"format":"x"}'),new TextEncoder().encode(JSON.stringify({...data,version:99})),new Uint8Array(6*1024*1024+1)]){await checked(a,bytes);assert.equal(a.f('error').hidden,false);assert.deepEqual(raw(a),prior);assert.equal(a.f('confirm').hidden,true);}
 a.f('input').files=[];a.f('input').onchange();await a.f('load').onclick();a.f('input').oncancel();assert.deepEqual(raw(a),prior);
 await checked(a,base.bytes);a.f('keep').onclick();await a.f('load').onclick();assert.equal(a.f('confirm').hidden,false,a.f('error').textContent);a.f('keep').onclick();assert.deepEqual(raw(a),prior);
});

test('pending replacement binds actual file, raw draft, confirmation text and normalized image visibility',async()=>{
 for(const mutate of [a=>{a.f('input').files=[new File(['{}'],'other.json')];},a=>a.n('family-brand-organizationName').value='NEW EVENTLESS',a=>a.f('confirm-text').textContent+=' tampered',a=>a.f('image').hidden=true,a=>a.f('image').className='hidden-candidate',a=>a.f('image').setAttribute('src','data:image/png;base64,WRONG')]){
  const a=await app();await selectLogo(a);author(a);const file=await saved(a);await checked(a,file.bytes);const prior=raw(a);mutate(a);a.f('replace').onclick();assert.equal(a.f('error').hidden,false);assert.equal(a.l('file'),prior.input);assert.equal(a.l('file').files,prior.files);assert.equal(a.l('image').getAttribute('src'),prior.image);assert.equal(a.n('family-front-cover-title').value,prior.title);
 }
});

test('interrupted local reads across cancel, empty/same-file return, host busy, view away/back or raw edits cannot open a replacement',async()=>{
 for(const change of [a=>a.f('input').oncancel(),a=>{const file=a.f('input').files[0];a.f('input').files=[];a.f('input').onchange();a.f('input').files=[file];},async a=>{await a.go('course');await a.go('output');},async a=>{await click(a,'refresh');},a=>{a.n('family-front-cover-title').value='NEW RAW';}]){
  const a=await app(),file=await saved(a),prior=raw(a);let finish,reads=0;const selected={size:file.bytes.length,stream(){return {getReader(){return {read:()=>++reads===1?new Promise(resolve=>finish=()=>resolve({done:false,value:file.bytes})):Promise.resolve({done:true}),cancel(){return Promise.resolve();},releaseLock(){}};}};}};a.f('input').files=[selected];a.f('input').onchange();const pending=a.f('load').onclick();await settle();await change(a);finish();await pending;assert.equal(a.f('confirm').hidden,true);assert.equal(a.f('error').hidden,false);assert.equal(a.l('file'),prior.input);assert.equal(a.l('image').getAttribute('src'),prior.image);
 }
});

test('current PNG normalization can change supported stored bytes without changing the original bundle or refusing a safe import',async()=>{
 const a=await app();await selectLogo(a);author(a);const original=await saved(a),file=JSON.parse(new TextDecoder().decode(original.bytes));const png=syntheticLogo({before:[logoChunk('tEXt',Buffer.from('Credit\0PRIVATE_STRIPPABLE_METADATA'))]});file.payload.logo.image.base64=Buffer.from(png).toString('base64');file.payload.logo.image.sha256=createHash('sha256').update(png).digest('hex');const bytes=new TextEncoder().encode(JSON.stringify(file)),before=new Uint8Array(bytes);
 await checked(a,bytes);assert.equal(a.f('confirm').hidden,false,a.f('error').textContent);assert.match(a.f('output').textContent,/normalization changed/u);assert.doesNotMatch(a.f('output').textContent,/PRIVATE_STRIPPABLE_METADATA/u);a.f('replace').onclick();assert.equal(a.f('error').hidden,true,a.f('error').textContent);const after=parseWorkbookDesignFile((await saved(a)).bytes);assert.notEqual(after.payload.logo.image.sha256,file.payload.logo.image.sha256);assert.equal(after.payload.logo.image.base64,parseWorkbookDesignFile(original.bytes).payload.logo.image.base64);assert.deepEqual(bytes,before);
});

test('all five family bundles retain current full/selected scope, source data and local preset selection without writing them',async()=>{
 let p=fixture();const made=editDeliveryPlan(p,{kind:'create',name:'Scope stays current',targetMinutes:30,groups:p.modules.map(m=>({moduleId:m.id,activityIds:m.activities.map(a=>a.id)}))});p=made.next;
 for(const scope of ['',made.id]){const a=await app(structuredClone(p));a.n('workbook-plan').value=scope;a.n('workbook-plan').onchange();const source=structuredClone(a.project),before=a.calls();a.n('family-preset-name').value='UNRELATED PRESET NAME';
  for(const family of WORKBOOK_FAMILIES){set(a,'family-template',family.id,'onchange');set(a,'family-brand-organizationName','Organization '+family.id);const file=await saved(a);set(a,'family-template',WORKBOOK_FAMILIES.find(x=>x.id!==family.id).id,'onchange');set(a,'family-brand-organizationName','Other current');await checked(a,file.bytes);a.f('replace').onclick();assert.equal(a.f('error').hidden,true,a.f('error').textContent);assert.equal(a.n('family-template').value,family.id);assert.equal(a.n('family-brand-organizationName').value,'Organization '+family.id);assert.equal(a.n('workbook-plan').value,scope);assert.equal(a.n('family-preset-name').value,'UNRELATED PRESET NAME');assert.equal(a.n('family-reviewed').checked,false);}
  assert.deepEqual(a.project,source);assert.deepEqual(a.calls(),before);
 }
});

test('overlapping front/research/logo/preset choices refuse design file actions and keep their drafts',async()=>{
 const a=await app();await selectLogo(a);author(a);const file=await saved(a);
 for(const [request,keep]of [['family-front-reset','family-front-keep'],['family-research-reset','family-research-keep'],['family-logo-reset','family-logo-keep'],['family-preset-reset','family-preset-keep']]){await click(a,request);const prior=raw(a);assert.equal((await offers(()=>click(a,'family-design-file-save'))).length,0);await checked(a,file.bytes);assert.equal(a.f('error').hidden,false);assert.equal(a.f('confirm').hidden,true);assert.deepEqual(raw(a),prior);await click(a,keep);}
});

test('post-swap late failure rolls back original native file node/FileList, raw controls, thumbnail and usable source handlers',async()=>{
 const a=await app();await selectLogo(a);author(a);const file=await saved(a);set(a,'family-brand-organizationName','OLD CURRENT');const prior=raw(a);await checked(a,file.bytes);const target=a.n('family-brand-status'),descriptor=Object.getOwnPropertyDescriptor(Object.getPrototypeOf(target),'textContent');let fired=false;Object.defineProperty(target,'textContent',{configurable:true,get(){return descriptor.get.call(this);},set(v){if(!fired&&String(v)==='Choose a design and refresh its preview.'){fired=true;throw Error('Synthetic late render fault');}return descriptor.set.call(this,v);}});try{a.f('replace').onclick();}finally{delete target.textContent;}
 assert.equal(fired,true);assert.equal(a.f('error').hidden,false);assert.deepEqual(raw(a),prior);assert.equal(a.l('file').files,prior.files);a.l('alt').value='Still usable after rollback';a.l('alt').oninput();assert.equal(a.l('image').getAttribute('alt'),'Still usable after rollback');
});

test('stale original logo chooser after successful same-ID picker swap cannot modify the imported image',async()=>{
 const a=await app();await selectLogo(a);author(a);const file=await saved(a),old=a.l('file');await checked(a,file.bytes);a.f('replace').onclick();const before=raw(a);old.files=[new File([syntheticLogo({width:9,height:2})],'STALE.png')];await old.onchange();assert.equal(a.l('error').hidden,false);assert.equal(a.l('image').getAttribute('src'),before.image);assert.equal(a.l('file'),before.input);assert.equal(a.l('alt').value,before.alt);
});
