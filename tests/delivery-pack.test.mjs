import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import {DOMParser} from '@xmldom/xmldom';
import {fixture} from './helpers/mock-office.mjs';
import {addActivity} from '../src/model.mjs';
import {editDeliveryPlan} from '../src/delivery-plans.mjs';
import {buildDeliveryPack,deliveryPackPreview,verifyDeliveryPack,deliveryPackName} from '../src/delivery-pack.mjs';
import {planAgendaFiles} from '../src/plan-agenda.mjs';
import {planWorkbookFiles,planMaterialsFiles} from '../src/selected-plan-outputs.mjs';
import {planFacilitatorFiles} from '../src/selected-plan-facilitator.mjs';

const stamp='2026-10-26T10:00:00.000Z';
globalThis.DOMParser=DOMParser;
function saved(){
 let p=fixture();p=addActivity(p,p.modules[0].id,{title:'OMITTED_TASK',minutes:6,participant:'OMITTED_LEARNER',facilitator:'OMITTED_PRIVATE',answers:'OMITTED_ANSWER',lines:2,materials:[{item:'OMITTED_SUPPLY',quantity:1,unit:'each',use:'consumable',notes:'OMITTED_PREP'}]}).next;
 p=addActivity(p,p.modules[0].id,{title:'Selected task',minutes:12,participant:'SELECTED_LEARNER',facilitator:'SELECTED_PRIVATE',answers:'SELECTED_ANSWER',lines:3,materials:[{item:'Cards',quantity:2,unit:'sets',use:'consumable',notes:'SELECTED_PREP'}]}).next;
 p=editDeliveryPlan(p,{kind:'create',name:'Short plan',targetMinutes:20,groups:[{moduleId:p.modules[0].id,activityIds:[p.modules[0].activities[2].id]}]}).next;
 return {p,id:p.deliveryPlans.plans[0].id};
}
const options=(route,answers=false,participant=false)=>route==='private'?{route,start:'09:00',stamp,answers,participant}:{route,start:'09:00',stamp};
const parts=async blob=>{const z=await JSZip.loadAsync(blob);return Object.fromEntries(await Promise.all(Object.values(z.files).filter(f=>!f.dir).map(async f=>[f.name,await f.async('string')])));};

test('shareable pack uses one selected identity, exact public roster and actual member part hashes',async()=>{
 const {p,id}=saved(),before=structuredClone(p);p.opaque={private:'OPAQUE_SECRET'};
 const result=await buildDeliveryPack(p,id,options('shareable'),JSZip),manifest=await verifyDeliveryPack(result.data,JSZip),zip=await JSZip.loadAsync(result.data);
 assert.deepEqual(result.preview.contentPreviews.map(row=>row.name),['agenda.docx','participant-workbook.docx','materials-checklist.docx']);assert.match(result.preview.contentPreviews[1].text,/SELECTED_LEARNER/);assert.doesNotMatch(JSON.stringify(result.preview),/SELECTED_PRIVATE|SELECTED_ANSWER|SELECTED_PREP|OMITTED_PRIVATE|OPAQUE_SECRET/);
 assert.deepEqual(manifest.members.map(m=>m.name),['agenda.docx','participant-workbook.docx','materials-checklist.docx']);assert.equal(manifest.facilitatorOptions,null);assert.equal(manifest.label,'SHAREABLE LEARNER PACK');
 assert.deepEqual(manifest.identity.scope,[{moduleId:p.modules[0].id,activityIds:[p.modules[0].activities[2].id]}]);assert.equal(manifest.identity.minutes,12);assert.equal(manifest.identity.start,'09:00');assert.equal(manifest.identity.end,'09:12');
 const expected=[planAgendaFiles(p,id,'09:00',stamp),planWorkbookFiles(p,id,stamp),planMaterialsFiles(p,id,stamp)];
 for(let i=0;i<3;i++){const member=manifest.members[i],actual=await parts(await zip.file(member.name).async('uint8array'));assert.deepEqual(actual,expected[i]);assert.equal(Object.keys(member.partHashes).length,6);for(const forbidden of ['OMITTED_TASK','OMITTED_LEARNER','OMITTED_PRIVATE','OMITTED_ANSWER','OMITTED_SUPPLY','OMITTED_PREP','SELECTED_PRIVATE','SELECTED_ANSWER','SELECTED_PREP','OPAQUE_SECRET'])assert.doesNotMatch(Object.values(actual).join('\n'),new RegExp(forbidden));}
 assert.doesNotMatch(await zip.file('manifest.json').async('string'),/PRIVATE_|SELECTED_PREP|answers\/debrief|OPAQUE_SECRET/);assert.equal(Object.keys(zip.files).length,4);assert.deepEqual(p.opaque,{private:'OPAQUE_SECRET'});assert.deepEqual({...p,opaque:undefined},{...before,opaque:undefined});
});

test('private pack has fixed private roster, independent options and unchanged standalone source output parts',async()=>{
 const {p,id}=saved(),standalone=planFacilitatorFiles(p,id,{answers:false,participant:false},stamp);
 for(const answers of [false,true])for(const participant of [false,true]){
  const result=await buildDeliveryPack(p,id,options('private',answers,participant),JSZip),manifest=await verifyDeliveryPack(result.data,JSZip),zip=await JSZip.loadAsync(result.data);
  assert.match(result.preview.contentPreviews[3].text,/SELECTED_PRIVATE|SELECTED_PREP/);
  assert.equal(manifest.label,'PRIVATE FACILITATOR PACK');assert.deepEqual(manifest.facilitatorOptions,{answers,participant});assert.deepEqual(manifest.members.map(m=>m.name),['agenda.docx','participant-workbook.docx','materials-checklist.docx','PRIVATE-facilitator-guide.docx']);
  const inner=await parts(await zip.file('PRIVATE-facilitator-guide.docx').async('uint8array')),body=inner['word/document.xml'];
  const source=planFacilitatorFiles(p,id,{answers,participant},stamp);
  for(const [part,text] of Object.entries(source))assert.equal(inner[part],part==='word/document.xml'?text.replace('Coordinated packs and existing-output regeneration are unavailable.','Existing-output regeneration remains unavailable.'):text,`pack-only disclosure difference: ${part}`);
  assert.match(body,/PRIVATE|SELECTED_PRIVATE|SELECTED_PREP/);assert.equal(body.includes('SELECTED_ANSWER'),answers);assert.equal(body.includes('SELECTED_LEARNER'),participant);assert.doesNotMatch(body,/OMITTED_TASK|OMITTED_PRIVATE|OMITTED_ANSWER|OMITTED_PREP/);
  assert.ok(body.includes('Existing-output regeneration remains unavailable.'));assert.ok(!body.includes('Coordinated packs and existing-output regeneration are unavailable.'));
  for(const name of ['agenda.docx','participant-workbook.docx','materials-checklist.docx'])assert.doesNotMatch(Object.values(await parts(await zip.file(name).async('uint8array'))).join(''),/SELECTED_PRIVATE|SELECTED_PREP|SELECTED_ANSWER/);
 }
 assert.deepEqual(planFacilitatorFiles(p,id,{answers:false,participant:false},stamp),standalone);assert.match(deliveryPackName('private'),/PRIVATE Facilitator Pack/);
});

test('repeated healthy builds retain exact member and outer ZIP bytes while new download names differ',async()=>{
 const {p,id}=saved();for(const route of ['shareable','private']){
  const choice=options(route,true,true),first=await buildDeliveryPack(p,id,choice,JSZip),second=await buildDeliveryPack(p,id,choice,JSZip);
  assert.deepEqual(first.data,second.data);assert.deepEqual(first.manifest,second.manifest);
  const z=await JSZip.loadAsync(first.data);for(const member of first.manifest.members){const other=await JSZip.loadAsync(second.data);assert.deepEqual(await z.file(member.name).async('uint8array'),await other.file(member.name).async('uint8array'));}
  assert.notEqual(deliveryPackName(route),deliveryPackName(route));
 }
});

test('LF, CRLF, tabs and escaped text keep exact selected identity across all members',async()=>{
 for(const name of ['First\nSecond','First\r\nSecond','First\tSecond & <Third>']){
  const {p,id}=saved();p.deliveryPlans.plans[0].name=name;p.title='Course\r\nTitle & <Teaching>';p.modules[0].title='Module\nTwo';p.modules[0].activities[2].title='Selected\tTask & <Practice>';
  for(const route of ['shareable','private']){
   const built=await buildDeliveryPack(p,id,options(route),JSZip),manifest=await verifyDeliveryPack(built.data,JSZip);
   assert.equal(manifest.identity.planName,name.replace(/\r\n?/g,'\n'));assert.equal(built.preview.identity.planName,manifest.identity.planName);
   for(const row of built.preview.contentPreviews){assert.match(row.text,/Delivery plan: First[\n\t]Second/);assert.match(row.text,/Module\nTwo|Selected\tTask/);}
  }
 }
});

test('generator-owned inner and outer byte buffers cannot change the final validated download',async()=>{
 const {p,id}=saved();for(const route of ['shareable','private']){
  const retained=[];class AliasingZip extends JSZip{async generateAsync(config){const value=await super.generateAsync(config);retained.push(value);return value;}}
  let calls=0,mutatedInner=false,mutatedOuter=false;
  const guard=()=>{calls++;if(calls===3){retained[0].fill(0);mutatedInner=true;}if(calls===(route==='private'?20:16)){for(const buffer of retained)buffer.fill(0);mutatedOuter=true;}};
  const built=await buildDeliveryPack(p,id,options(route),AliasingZip,guard);
  assert.equal(mutatedInner,true);assert.equal(mutatedOuter,true);assert.equal(retained.at(-1).every(value=>value===0),true);
  assert.deepEqual((await verifyDeliveryPack(built.data,JSZip)).members.map(row=>row.name),route==='private'?['agenda.docx','participant-workbook.docx','materials-checklist.docx','PRIVATE-facilitator-guide.docx']:['agenda.docx','participant-workbook.docx','materials-checklist.docx']);
 }
});

test('explicit expected context requires exact own-data member and six-part maps',async()=>{
 const {p,id}=saved(),built=await buildDeliveryPack(p,id,options('shareable'),JSZip),z=await JSZip.loadAsync(built.data),expectedManifestBytes=await z.file('manifest.json').async('string');
 const maps={'agenda.docx':planAgendaFiles(p,id,'09:00',stamp),'participant-workbook.docx':planWorkbookFiles(p,id,stamp),'materials-checklist.docx':planMaterialsFiles(p,id,stamp)};
 assert.deepEqual((await verifyDeliveryPack(built.data,JSZip,{expectedManifestBytes,expectedDocFiles:maps})).members.map(row=>row.name),Object.keys(maps));
 let calls=0;const accessor={...maps};Object.defineProperty(accessor,'agenda.docx',{get(){calls++;return maps['agenda.docx'];}});
 const nested={...maps,'agenda.docx':{...maps['agenda.docx']}};Object.defineProperty(nested['agenda.docx'],'word/document.xml',{get(){calls++;return maps['agenda.docx']['word/document.xml'];}});
 const inherited=Object.assign(Object.create({'agenda.docx':maps['agenda.docx']}),{'participant-workbook.docx':maps['participant-workbook.docx'],'materials-checklist.docx':maps['materials-checklist.docx']});
 for(const context of [null,{expectedManifestBytes},{expectedDocFiles:maps},{expectedManifestBytes,expectedDocFiles:null},{expectedManifestBytes,expectedDocFiles:undefined},{expectedManifestBytes,expectedDocFiles:{}},{expectedManifestBytes,expectedDocFiles:{...maps,'agenda.docx':undefined}},{expectedManifestBytes,expectedDocFiles:{'participant-workbook.docx':maps['participant-workbook.docx'],'materials-checklist.docx':maps['materials-checklist.docx']}},{expectedManifestBytes,expectedDocFiles:accessor},{expectedManifestBytes,expectedDocFiles:nested},{expectedManifestBytes,expectedDocFiles:inherited},{expectedManifestBytes,expectedDocFiles:{...maps,extra:{}}}])await assert.rejects(verifyDeliveryPack(built.data,JSZip,context),/expected|context/i);
 assert.equal(calls,0);
});

test('malformed plans, private choices, partial assembly and changed source refuse without a pack',async()=>{
 const {p,id}=saved();
 for(const bad of ['missing','duplicate','wrong-parent','empty']){const q=structuredClone(p);if(bad==='duplicate')q.deliveryPlans.plans.push(structuredClone(q.deliveryPlans.plans[0]));if(bad==='wrong-parent')q.deliveryPlans.plans[0].groups[0].activityIds=['foreign'];if(bad==='empty')q.deliveryPlans.plans[0].groups[0].activityIds=[];await assert.rejects(buildDeliveryPack(q,bad==='missing'?'missing':id,options('shareable'),JSZip),/unique|repair|No activities|plan/i);}
 assert.throws(()=>deliveryPackPreview(p,id,{...options('shareable'),answers:true}),/Private choices/);
 const broken=structuredClone(p);broken.modules[0].activities[2].facilitator={secret:'PRIVATE_INVALID'};assert.doesNotThrow(()=>deliveryPackPreview(broken,id,options('shareable')));assert.throws(()=>deliveryPackPreview(broken,id,options('private')),/selected facilitator instructions/i);
 let stages=0;class FailingZip extends JSZip{async generateAsync(o){if(++stages===2)throw new Error('assembly stopped');return super.generateAsync(o);}}
 await assert.rejects(buildDeliveryPack(p,id,options('shareable'),FailingZip),/assembly stopped/);
 let guards=0;await assert.rejects(buildDeliveryPack(p,id,options('shareable'),JSZip,()=>{if(++guards===4)throw new Error('source changed');}),/source changed/);
});

test('pack configuration refuses inherited, accessor, symbol and cross-route private choices without invoking getters',()=>{
 const {p,id}=saved();let calls=0;
 const inherited=Object.create({get answers(){calls++;return true;}});Object.assign(inherited,{route:'private',start:'09:00',stamp});assert.throws(()=>deliveryPackPreview(p,id,inherited),/Inherited pack choices/);
 const inheritedPublic=Object.create({get participant(){calls++;return true;}});Object.assign(inheritedPublic,options('shareable'));assert.throws(()=>deliveryPackPreview(p,id,inheritedPublic),/Inherited pack choices/);
 const accessor=options('private');Object.defineProperty(accessor,'answers',{get(){calls++;return true;}});assert.throws(()=>deliveryPackPreview(p,id,accessor),/explicit data fields/);
 const hidden=options('private');Object.defineProperty(hidden,'participant',{get(){calls++;return true;},enumerable:false});assert.throws(()=>deliveryPackPreview(p,id,hidden),/explicit data fields/);
 const symbol=options('shareable');symbol[Symbol('secret')]=true;assert.throws(()=>deliveryPackPreview(p,id,symbol),/explicit data fields/);
 assert.throws(()=>deliveryPackPreview(p,id,{...options('shareable'),participant:false}),/Private choices/);
 assert.throws(()=>deliveryPackPreview(p,id,{...options('private'),answers:'true'}),/boolean|option/i);
 assert.throws(()=>deliveryPackPreview(p,id,{...options('shareable'),extra:'private'}),/explicit data fields/);
 for(const route of [null,undefined])assert.throws(()=>deliveryPackPreview(p,id,{...options('shareable'),route}),/supported pack route/);
 for(const start of [null,undefined])assert.throws(()=>deliveryPackPreview(p,id,{...options('shareable'),start}),/start time/);
 for(const stampValue of [null,undefined])assert.throws(()=>deliveryPackPreview(p,id,{...options('shareable'),stamp:stampValue}),/generation stamp/);
 for(const key of ['answers','participant'])for(const value of [null,undefined])assert.throws(()=>deliveryPackPreview(p,id,{...options('private'),[key]:value}),/boolean|option/i);
 assert.equal(calls,0);
});

test('outer roster, names, manifest and inner Word alterations are rejected',async()=>{
 const {p,id}=saved(),{data}=await buildDeliveryPack(p,id,options('shareable'),JSZip);
 const changed=async edit=>{const z=await JSZip.loadAsync(data);await edit(z);return z.generateAsync({type:'uint8array',compression:'DEFLATE'});};
 for(const edit of [
  z=>z.file('manifest.json','{}'),
  z=>z.remove('materials-checklist.docx'),
  z=>z.file('EXTRA.docx','x'),
  z=>z.file('AGENDA.docx','x'),
  z=>z.file('evil\\name.docx','x'),
  z=>z.file('agenda.docx','corrupt'),
  async z=>{const m=JSON.parse(await z.file('manifest.json').async('string'));m.identity.planId='wrong';z.file('manifest.json',JSON.stringify(m));},
  async z=>{const m=JSON.parse(await z.file('manifest.json').async('string'));m.members[0].partHashes['word/document.xml']='0'.repeat(64);z.file('manifest.json',JSON.stringify(m));},
  async z=>{const raw=await z.file('manifest.json').async('string');z.file('manifest.json',raw.replace('"route": "shareable"','"route": "private",\n  "route": "shareable"'));},
  z=>{z.comment='PRIVATE_HIDDEN_COMMENT';},
  async z=>{const body=await z.file('agenda.docx').async('uint8array');z.file('agenda.docx',body,{comment:'PRIVATE_HIDDEN_COMMENT'});},
  z=>z.file('HIDDEN_DIRECTORY/',null,{dir:true,createFolders:false}),
  z=>z.file('manifest.json',new Uint8Array([0xff,0xfe])),
  async z=>{const inner=await JSZip.loadAsync(await z.file('agenda.docx').async('uint8array'));inner.file('word/document.xml',new Uint8Array([0xff,0xfe]));z.file('agenda.docx',await inner.generateAsync({type:'uint8array',compression:'DEFLATE',createFolders:false}));},
 ])await assert.rejects(verifyDeliveryPack(await changed(edit),JSZip),/Delivery pack|Corrupted|zip/i);
});

test('generated pack refuses when a real XML parser is unavailable',async()=>{
 const {p,id}=saved(),parser=globalThis.DOMParser;delete globalThis.DOMParser;
 try{await assert.rejects(buildDeliveryPack(p,id,options('shareable'),JSZip),/XML parser is required/);}finally{globalThis.DOMParser=parser;}
});
