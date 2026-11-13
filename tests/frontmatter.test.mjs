import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import {course,wrenchCourse,validate,reconcile,managedTags,addModule,addActivity,total,participantView} from '../src/model.mjs';
import {frontMatterFields,frontMatterEvidence,validFrontJournal} from '../src/frontmatter.mjs';
import {KEY,PENDING} from '../src/document-store.mjs';
import {makeHost,adapter,records} from './helpers/mock-office.mjs';
import {frontHost,W,paragraph,control} from './helpers/frontmatter-fixture.mjs';
import {inspectCapturedFrontMatter,assertBlankMaster,inspectMaster,createCheckpoint,readCheckpoint} from '../src/checkpoint.mjs';
import {assertBlankBodyOoxml} from '../src/frontmatter-package.mjs';
import {assets} from './helpers/checkpoint-fixture.mjs';
import {worksheetBlob} from '../src/worksheet.mjs';
import {agendaBlob} from '../src/agenda.mjs';
import {materialsBlob} from '../src/materials-output.mjs';
import {auditExport} from './helpers/export-contract.mjs';
const selected=['getting-started','background','team','overview'];
const details={audience:'PRIVATE_WRENCH_AUDIENCE',purpose:'PRIVATE_WRENCH_PURPOSE',authors:'PRIVATE_WRENCH_AUTHOR',attribution:'PRIVATE_WRENCH_LICENSE'};
const starter=()=>wrenchCourse('Wrench sample',details,selected);
const bytes=async h=>(await h.package()).generateAsync({type:'uint8array'});
const plain=h=>h.records.map(({tag,text})=>({tag,text}));

test('versioned Wrench sections have canonical stable unique identities and cannot become learning modules or legacy adoption',()=>{
 const p=starter();assert.equal(p.schema,2);assert.equal(p.revision,0);assert.equal(total(p),0);assert.deepEqual(p.modules,[]);assert.equal(p.frontMatter.sections.length,5);assert.equal(frontMatterFields(p).length,10);assert.equal(managedTags(p).size,11);assert.deepEqual(participantView(p).modules,[]);assert.equal(wrenchCourse('Opening').frontMatter.sections.length,1);
 assert.deepEqual(wrenchCourse('Canonical',undefined,['overview','background']).frontMatter.sections.map(s=>s.kind),['opening','background','overview']);
 for(const mutate of [p=>p.frontMatter.version=2,p=>p.frontMatter.profile='future',p=>p.frontMatter.sections.pop()&&p.frontMatter.sections.reverse(),p=>p.frontMatter.sections[1].id=p.id,p=>p.frontMatter.sections[1].id=p.frontMatter.sections[0].id,p=>p.frontMatter.sections[0].body=null,p=>p.frontMatter.sections[0].heading='',p=>p.frontMatter.sections[0].body='x'.repeat(50001)]){const q=structuredClone(p);mutate(q);assert.throws(()=>validate(q));}
 for(const kinds of [['opening'],['background','background'],['future'],null])assert.throws(()=>wrenchCourse('Invalid',undefined,kinds));
 p.frontMatter.unknown={keep:true};p.frontMatter.sections[0].future='PRIVATE_WRENCH_FUTURE';const q=reconcile(p,records(p));assert.deepEqual(q,p);const m=addModule(p,'Actual learning','Objective').next;assert.deepEqual(m.frontMatter,p.frontMatter);assert.equal(m.modules.length,1);
 const legacy=course('Legacy');legacy.frontMatter={version:999,arbitrary:'opaque'};assert.deepEqual(validate(legacy),legacy);assert.equal(managedTags(legacy).size,1);
});

test('blank Wrench creation commits every protected heading and real empty body in a single schema5 transaction',async()=>{
 const p=starter(),h=frontHost(makeHost(null)),w=await adapter(h);assert.deepEqual(await w.insert(null,p,'course'),p);assert.equal(h.saves,2);assert.equal(h.records.length,11);assert.equal(h.paragraphs.length,11);assert.ok(h.records.every(r=>r.cannotDelete));assert.deepEqual(plain(h),records(p).map(r=>({...r,text:r.text||'\u200b'})));assert.deepEqual(h.paragraphs.map(p=>p.styleBuiltIn),['Title',...Array.from({length:5},()=>['Heading1','Normal']).flat()]);assert.equal(h.fileReads,h.fileCloses);assert.equal(h.fileReads,4);assert.equal(h.durable[PENDING].schema,5);assert.equal(validFrontJournal(h.durable[PENDING],p),true);assert.deepEqual(h.durable[PENDING].frontMatter,frontMatterEvidence(p));
 assert.deepEqual(await w.recover(),p);assert.deepEqual(await (await adapter(h.reopen())).loadProject(),p);assert.equal(h.saves,2);assert.equal(h.records.length,11);
});

test('Minimal retains unmanaged-content creation and never requests a blank or compressed-file guard',async()=>{
 const p=course('Minimal unchanged'),h=makeHost(null);h.extraBody=paragraph('USER MANUAL TEXT');h.Office.context.document.getFileAsync=()=>{throw new Error('Minimal must not capture a package');};h.bodyOoxml=()=>{throw new Error('Minimal must not ask for blank body');};const w=await adapter(h);assert.deepEqual(await w.insert(null,p,'course'),p);assert.equal(h.extraBody,paragraph('USER MANUAL TEXT'));assert.equal(h.saves,2);assert.equal(h.durable[PENDING].schema,2);
});

test('nonblank and unsupported stories assets controls and metadata are refused before settings or content mutation',async()=>{
 const cases=[h=>h.extraBody=paragraph('Manual text'),h=>h.extraBody='<w:tbl/>',h=>h.extraBody='<w:p><w:r><w:drawing/></w:r></w:p>',h=>h.extraBody=control({tag:'other:control',text:''}),h=>h.extraBody='<w:p/><w:p/>',h=>h.packageTransform=z=>z.file('relocated/STORY.XML',`<w:hdr xmlns:w="${W}">${paragraph('Hidden header')}</w:hdr>`),h=>h.packageTransform=z=>z.file('word/media/keep.png',new Uint8Array([1])),h=>h.packageTransform=z=>z.file('customXml/item1.xml','<private/>'),h=>h.packageTransform=z=>z.file('word/webextensions/other.xml','<we:properties xmlns:we="http://schemas.microsoft.com/office/webextensions/webextension/2010/11"><we:property name="openstage.guide.pending.v1" value="{}"/></we:properties>')];
 for(const setup of cases){const h=frontHost(makeHost(null));setup(h);const w=await adapter(h),before=structuredClone(h.durable),extra=h.extraBody;await assert.rejects(w.insert(null,starter(),'course'),e=>e.code==='OSG_INPUT');assert.equal(h.saves,0);assert.equal(h.paragraphs.length,0);assert.equal(h.records.length,0);assert.equal(h.extraBody,extra);assert.deepEqual(h.durable,before);assert.equal(w.isBlocked(),false);assert.equal(h.fileReads,h.fileCloses);}
});

test('late content at second package guard final body guard or completion stops without clearing user work or replaying insertion',async()=>{
 for(const boundary of ['package','body','completion']){const h=frontHost(makeHost(null));if(boundary==='package')h.hooks.beforeFile=n=>{if(n===2)h.extraBody=paragraph('LATE MANUAL');};if(boundary==='body')h.hooks.beforeBodyOoxml=()=>h.extraBody=paragraph('LATE MANUAL');if(boundary==='completion')h.hooks.beforeFile=n=>{if(n===3)h.extraBody=paragraph('LATE MANUAL');};const w=await adapter(h);await assert.rejects(w.insert(null,starter(),'course'));assert.match(h.extraBody,/LATE MANUAL/);assert.equal(h.saves,1);assert.equal(h.durable[KEY],undefined);assert.equal(h.durable[PENDING].phase,'prepared');const fields=h.records.length,saves=h.saves;await assert.rejects(w.recover());await assert.rejects((await adapter(h.reopen())).recover());assert.equal(h.saves,saves);assert.equal(h.records.length,fields);assert.match(h.extraBody,/LATE MANUAL/);}
});

test('every dropped heading or empty body makes initial recovery incomplete with no replay deletion or guessed rollback',async()=>{
 for(let drop=0;drop<11;drop++){const p=starter(),h=frontHost(makeHost(null));let field=0;h.dropField=()=>field++===drop;const w=await adapter(h);await assert.rejects(w.insert(null,p,'course'));assert.equal(h.records.length,10);assert.equal(h.saves,1);assert.equal(h.durable[KEY],undefined);const preserved=plain(h),durable=structuredClone(h.durable);for(let retry=0;retry<2;retry++)await assert.rejects(w.recover());await assert.rejects((await adapter(h.reopen())).recover());assert.deepEqual(plain(h),preserved);assert.deepEqual(h.durable,durable);assert.equal(h.paragraphs.length,11);assert.equal(h.saves,1);}
});

test('every physical insertion boundary uncertain sync preserves partial work and recovers only a complete exact starter',async()=>{
 // The isolated queue injects before/after each field synchronization, not a
 // mocked successful store result. Locate the first write through its queue.
 for(let boundary=1;boundary<=11;boundary++)for(const apply of [false,true]){const p=starter(),h=frontHost(makeHost(null)),w=await adapter(h);let started=false,writes=0;h.hooks.beforeSync=(n)=>{if(!started&&h.saves===1&&h.fileReads===2&&h.records.length===0){started=true;return;}if(started&&h.records.length<11&&++writes===boundary)h.syncFailures.set(n,{apply});};await assert.rejects(w.insert(null,p,'course'));const count=h.records.length,paras=h.paragraphs.length,saves=h.saves;h.syncFailures.clear();h.hooks.beforeSync=undefined;if(count===11){assert.deepEqual(await w.recover(),p);assert.equal(h.saves,saves+1);}else{await assert.rejects(w.recover());assert.equal(h.saves,saves);}assert.equal(h.records.length,count);assert.equal(h.paragraphs.length,paras);}
});

test('initial divergent title heading or body is not adopted as a manual edit during prepared recovery',async()=>{
 for(const index of [0,1,2,9,10]){const p=starter(),h=frontHost(makeHost(null)),w=await adapter(h);h.saveFailures.set(2,{persist:false});await assert.rejects(w.insert(null,p,'course'));h.records[index].text='TRUNCATED OR ADDED PRIVATE TEXT';const before=plain(h),saved=structuredClone(h.durable);await assert.rejects(w.recover());await assert.rejects((await adapter(h.reopen())).recover());assert.equal(h.saves,2);assert.deepEqual(plain(h),before);assert.deepEqual(h.durable,saved);}
});

test('uncertain prepared final and recovery saves retain exact intent and converge idempotently without reinsertion',async()=>{
 for(const save of [1,2,3])for(const persist of [false,true]){const p=starter(),h=frontHost(makeHost(null)),w=await adapter(h);if(save===3)h.saveFailures.set(2,{persist:false});h.saveFailures.set(save,{persist});await assert.rejects(w.insert(null,p,'course'));if(save===1){assert.equal(h.records.length,0);await assert.rejects(w.recover());assert.equal(h.saves,1);continue;}if(save===3)await assert.rejects(w.recover());const reopened=await adapter(h.reopen());assert.deepEqual(await reopened.recover(),p);const saves=h.saves;assert.deepEqual(await reopened.recover(),p);assert.equal(h.saves,saves);assert.equal(h.records.length,11);assert.equal(h.paragraphs.length,11);assert.deepEqual(h.durable[KEY],p);assert.equal(h.durable[PENDING].schema,5);}
});

test('committed manual section edits remain authoritative and all subsequent operations retain frontmatter and timing evidence',async()=>{
 const p=starter();p.unknown={private:'PRIVATE_WRENCH_UNKNOWN'};p.frontMatter.sections[0].unknown='PRIVATE_SECTION_UNKNOWN';const h=frontHost(makeHost(p)),w=await adapter(h),ids=p.frontMatter.sections.map(s=>s.id);h.records[1].text='Manually renamed opening';h.records[2].text='PRIVATE_WRENCH_BODY\nSecond paragraph';h.extraBody=paragraph('PRIVATE_MANUAL_NOTE');let fresh=await w.readProject();assert.equal(fresh.frontMatter.sections[0].heading,'Manually renamed opening');assert.equal(fresh.frontMatter.sections[0].body,'PRIVATE_WRENCH_BODY\nSecond paragraph');await w.persist(fresh);fresh=await w.updateCourseDetails(fresh,{...details,purpose:'Revised private purpose'});
 const m=addModule(fresh,'Real lesson','Objective');fresh=await w.insert(fresh,m.next,'module',m.item);const a=addActivity(fresh,m.item.id,{title:'Practice',minutes:5,lines:2,participant:'Public task',facilitator:'PRIVATE_FACILITATOR',answers:'PRIVATE_ANSWER'});fresh=await w.insert(fresh,a.next,'activity',a.item);fresh=await w.updateActivityMaterials(fresh,a.item.id,[{item:'Paper',quantity:1,unit:'sets',use:'consumable',notes:'PRIVATE_MATERIAL_NOTE'}]);fresh=await w.updateActivitySettings(fresh,a.item.id,{minutes:8,lines:3});assert.equal(h.durable[PENDING].schema,5);assert.deepEqual(h.durable[PENDING].cueChecks,[a.item.id]);assert.equal(validFrontJournal(h.durable[PENDING],fresh),true);assert.equal(h.cues[a.item.id],'Learning activity · 8 min');assert.deepEqual(fresh.frontMatter.sections.map(s=>s.id),ids);assert.equal(fresh.frontMatter.sections[0].body,'PRIVATE_WRENCH_BODY\nSecond paragraph');assert.deepEqual(fresh.unknown,p.unknown);assert.equal(fresh.frontMatter.sections[0].unknown,'PRIVATE_SECTION_UNKNOWN');assert.match(h.extraBody,/PRIVATE_MANUAL_NOTE/);assert.equal(h.records.filter(r=>r.tag.startsWith('osg:section-')).length,10);assert.equal(total(fresh),8);assert.equal(fresh.modules.length,1);assert.deepEqual(await (await adapter(h.reopen())).loadProject(),fresh);
});

test('lost misordered nested table or shared-heading profile inventory blocks reads metadata modules materials settings and backup',async()=>{
 let p=starter();p=addModule(p,'Lesson','Objective').next;p=addActivity(p,p.modules[0].id,{title:'Practice',minutes:5,participant:'Public task'}).next;const tag=frontMatterFields(p)[0].tag;
 const corruptions=[h=>h.records.splice(2,1),h=>h.records.push({...h.records[1]}),h=>h.records.push({tag:'osg:section-body:'+crypto.randomUUID(),text:''}),h=>h.physicalOrder=[h.records[1].tag,h.records[0].tag,...h.records.slice(2).map(r=>r.tag)],h=>h.nestedTags=new Set([tag]),h=>h.tableTags=new Set([tag]),h=>h.unrelatedTags=new Set([tag]),h=>h.paragraphText={[tag]:'Shared manual text Opening'}];
 for(const corrupt of corruptions)for(const operation of ['read','details','module','materials','settings','backup']){const h=frontHost(makeHost(p)),w=await adapter(h);corrupt(h);const saved=structuredClone(h.durable),before=plain(h),id=p.modules[0].activities[0].id;const actions={read:()=>w.readProject(),details:()=>w.updateCourseDetails(p,{...details,purpose:'Changed'}),module:()=>{const m=addModule(p,'New','Objective');return w.insert(p,m.next,'module',m.item);},materials:()=>w.updateActivityMaterials(p,id,[]),settings:()=>w.updateActivitySettings(p,id,{minutes:8,lines:2}),backup:()=>w.captureMaster()};await assert.rejects(actions[operation]());assert.equal(h.saves,0);assert.equal(h.paragraphs.length,0);assert.equal(h.cueWrites,0);assert.deepEqual(h.durable,saved);assert.deepEqual(plain(h),before);}
});

test('unknown downgraded missing or incomplete profile receipts cannot be adopted on reads recovery or backups',async()=>{
 for(const change of [j=>j.schema=2,j=>j.schema=3,j=>j.schema=4,j=>j.frontMatter.version=9,j=>j.frontMatter.sectionIds.pop(),j=>delete j.placement,j=>delete j.cueChecks,()=>null]){const p=starter(),h=frontHost(makeHost(p));const j=change(h.durable[PENDING]);if(j===null)delete h.durable[PENDING];h.memory=structuredClone(h.durable);const w=await adapter(h),saved=structuredClone(h.durable);await assert.rejects(w.loadProject());await assert.rejects(w.recover());await assert.rejects(inspectMaster(await bytes(h),JSZip));assert.deepEqual(h.durable,saved);assert.equal(h.saves,0);assert.equal(h.paragraphs.length,0);}
});

test('later uncertain operations recover the complete profile without downgrading or regenerating committed sections',async()=>{
 for(const kind of ['details','module','activity','materials','settings'])for(const persist of [false,true]){
  let p=starter();p.frontMatter.sections[0].body='PRIVATE_COMMITTED_BODY';p.frontMatter.extra={keep:true};p=addModule(p,'Lesson','Objective').next;p=addActivity(p,p.modules[0].id,{title:'Existing practice',minutes:5,participant:'Public task'}).next;
  const h=frontHost(makeHost(p)),w=await adapter(h);h.saveFailures.set(2,{persist});const id=p.modules[0].activities[0].id;
  const actions={details:()=>w.updateCourseDetails(p,{...details,purpose:'New purpose'}),module:()=>{const m=addModule(p,'New lesson','Objective');return w.insert(p,m.next,'module',m.item);},activity:()=>{const a=addActivity(p,p.modules[0].id,{title:'New practice',minutes:3,participant:'New task'});return w.insert(p,a.next,'activity',a.item);},materials:()=>w.updateActivityMaterials(p,id,[{item:'Paper',quantity:2,unit:'sets',use:'consumable',notes:'PRIVATE'}]),settings:()=>w.updateActivitySettings(p,id,{minutes:9,lines:2})};
  await assert.rejects(actions[kind]());assert.equal(h.durable[PENDING].schema,5);const fields=plain(h),paragraphs=h.paragraphs.length,physicalWrites=h.cueWrites;
  const reopened=await adapter(h.reopen()),result=await reopened.recover();assert.equal(validFrontJournal(h.durable[PENDING],result),true);assert.deepEqual(result.frontMatter,p.frontMatter);assert.deepEqual(plain(h),fields);assert.equal(h.paragraphs.length,paragraphs);assert.equal(h.cueWrites,physicalWrites);const saves=h.saves;await reopened.recover();assert.equal(h.saves,saves);
 }
});

test('all-story package faults remain blockers for later read write and backup operations even when native scan looks complete',async()=>{
 const p=starter();for(const action of ['read','details','module','backup']){const h=frontHost(makeHost(p)),w=await adapter(h);h.packageTransform=z=>z.file('external/STORY.XML',`<w:hdr xmlns:w="${W}">${control(h.records[2])}</w:hdr>`);const before=structuredClone(h.durable),body=plain(h);const actions={read:()=>w.readProject(),details:()=>w.updateCourseDetails(p,{...details,purpose:'New'}),module:()=>{const m=addModule(p,'New','Objective');return w.insert(p,m.next,'module',m.item);},backup:()=>w.captureMaster()};await assert.rejects(actions[action]());assert.deepEqual(h.durable,before);assert.deepEqual(plain(h),body);assert.equal(h.saves,0);assert.equal(h.paragraphs.length,0);}
});

test('captured all-story proof rejects moved orphan nested reordered or partial fields and validates actual blank body XML',async()=>{
 const p=starter(),h=frontHost(makeHost(p));await inspectCapturedFrontMatter(await bytes(h),p,JSZip);
 const transforms=[async z=>{const d=await z.file('word/document.xml').async('string');z.file('word/document.xml',d.replace(control(h.records[2]),''));},async z=>{const d=await z.file('word/document.xml').async('string');z.file('word/document.xml',d.replace(control(h.records[1]),`<w:tbl><w:tr><w:tc>${control(h.records[1])}</w:tc></w:tr></w:tbl>`));},z=>z.file('relocated/PART.XML',`<w:hdr xmlns:w="${W}">${control(h.records[2])}</w:hdr>`),async z=>{const d=await z.file('word/document.xml').async('string');z.file('word/document.xml',d.replace(control(h.records[1]),`<w:sdt><w:sdtPr/><w:sdtContent>${control(h.records[1])}</w:sdtContent></w:sdt>`));},async z=>{const d=await z.file('word/document.xml').async('string');z.file('word/document.xml',d.replace(control(h.records[1]),control({...h.records[1],text:'Partial'})));}];
 for(const transform of transforms){h.packageTransform=transform;await assert.rejects(inspectCapturedFrontMatter(await bytes(h),p,JSZip));}h.packageTransform=undefined;
 const blank=frontHost(makeHost(null));assertBlankBodyOoxml(blank.bodyOoxml());await assertBlankMaster(await bytes(blank),JSZip);for(const body of [paragraph('Manual'),'<w:tbl/>','<w:p/><w:p/>','<w:p><w:r><w:br/></w:r></w:p>']){blank.extraBody=body;assert.throws(()=>assertBlankBodyOoxml(blank.bodyOoxml()));}
});

test('full backup preserves exact governed master bytes unknown metadata manual formatting and private front matter while public packages exclude it',async()=>{
 let p=starter();p.frontMatter.sections.forEach((s,i)=>{s.body=`PRIVATE_FRONT_${i}`;});p.frontMatter.future='PRIVATE_FRONT_FUTURE';p=addModule(p,'Learning','Public objective').next;p=addActivity(p,p.modules[0].id,{title:'Practice',minutes:5,participant:'Public instructions',facilitator:'PRIVATE_FACILITATOR',answers:'PRIVATE_ANSWERS',materials:[{item:'Paper',quantity:2,unit:'sets',use:'consumable',notes:'PRIVATE_NOTES'}]}).next;const h=frontHost(makeHost(p));h.extraBody='<w:p><w:r><w:rPr><w:b/><w:color w:val="663399"/></w:rPr><w:t>PRIVATE_MANUAL_NOTE</w:t></w:r></w:p>';h.packageTransform=async z=>{let d=await z.file('word/document.xml').async('string');const a=p.modules[0].activities[0];d=d.replace(control({tag:`osg:activity:${a.id}`,text:a.title}),paragraph(`${a.icon} · ${a.minutes} min`)+control({tag:`osg:activity:${a.id}`,text:a.title}));z.file('word/document.xml',d);z.file('word/media/manual.png',new Uint8Array([137,80,78,71,1,2,3]));z.file('customXml/private.xml','<private>KEEP UNKNOWN PRIVATE XML</private>');};const master=await bytes(h);const saved=await createCheckpoint(master,p,JSZip,assets),restored=await readCheckpoint(saved.data,JSZip);assert.deepEqual(restored.master,master);assert.deepEqual(restored.project,p);assert.equal(restored.project.schema,2);assert.deepEqual(h.durable[PENDING].frontMatter,frontMatterEvidence(p));
 const canaries=[...Object.values(details),...p.frontMatter.sections.map(s=>s.body),'PRIVATE_FRONT_FUTURE','PRIVATE_MANUAL_NOTE','PRIVATE_FACILITATOR','PRIVATE_ANSWERS','PRIVATE_NOTES'];for(const kind of ['workbook','agenda','materials']){const blob=kind==='workbook'?await worksheetBlob(p,JSZip):kind==='agenda'?await agendaBlob(p,'09:00',JSZip):await materialsBlob(p,undefined,JSZip);await auditExport(blob,{project:p,privateCanaries:canaries,kind});}
});

test('explicit hidden required content refuses initial completion prepared recovery and later reads while supported off values and unrelated semiHidden remain valid',async()=>{
 const conceal=(tag,val,index=0)=>async z=>{const name='word/document.xml',d=await z.file(name).async('string'),attribute=val===undefined?'':` w:val="${val}"`;let at=0;z.file(name,d.replace(/<w:r><w:t>/g,match=>at++===index?'<w:r><w:rPr><w:'+tag+attribute+'/></w:rPr><w:t>':match));};
 for(const tag of ['vanish','webHidden'])for(const val of [undefined,'true','1','on','unsupported'])for(const index of [0,1,2]){
  const p=starter(),h=frontHost(makeHost(null)),w=await adapter(h);h.hooks.beforeFile=n=>{if(n===3)h.packageTransform=conceal(tag,val,index);};await assert.rejects(w.insert(null,p,'course'));assert.equal(h.saves,1);assert.equal(h.records.length,11);const body=plain(h),saved=structuredClone(h.durable);await assert.rejects(w.recover());await assert.rejects((await adapter(h.reopen())).recover());assert.equal(h.saves,1);assert.deepEqual(plain(h),body);assert.deepEqual(h.durable,saved);
  const complete=frontHost(makeHost(p));complete.packageTransform=conceal(tag,val,index);await assert.rejects((await adapter(complete)).readProject());await assert.rejects(inspectMaster(await bytes(complete),JSZip));assert.equal(complete.saves,0);
 }
 for(const tag of ['vanish','webHidden'])for(const val of ['false','0','off']){const p=starter(),h=frontHost(makeHost(p));h.packageTransform=conceal(tag,val);assert.deepEqual(await (await adapter(h)).readProject(),p);await inspectMaster(await bytes(h),JSZip);}
 const p=starter(),h=frontHost(makeHost(p));h.packageTransform=z=>z.file('word/styles.xml',`<w:styles xmlns:w="${W}"><w:style w:styleId="Unused"><w:semiHidden/><w:rPr><w:vanish/></w:rPr></w:style></w:styles>`);assert.deepEqual(await (await adapter(h)).readProject(),p);
});
