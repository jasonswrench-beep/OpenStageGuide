import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import {fixture,makeHost} from './helpers/mock-office.mjs';
import {addModule,addActivity} from '../src/model.mjs';
import JSZip from 'jszip';
import {masterFixture,installParser,assets} from './helpers/checkpoint-fixture.mjs';
import {createCheckpoint} from '../src/checkpoint.mjs';
import {emptyLibrary,saveLibraryCopy,serializeLibrary} from '../src/library.mjs';
import {pptxFixture} from './helpers/pptx-fixture.mjs';
import {readPptx,pptxDraft} from '../src/pptx.mjs';
import {serializeReview} from '../src/pptx-review.mjs';
import {serializeMaterialsDraft} from '../src/materials-draft.mjs';
import {exportCorpus,auditExport} from './helpers/export-contract.mjs';
import {frontHost,paragraph} from './helpers/frontmatter-fixture.mjs';
import {auditFacilitator} from './helpers/facilitator-contract.mjs';
import {facilitatorPreview,facilitatorParagraphs} from '../src/facilitator-output.mjs';

import {Element,dom} from './helpers/app-dom.mjs';

test('guided course setup starts blank and read-only guidance never populates entries changes outputs or calls Word',async()=>{
  const h=makeHost(null),ui=dom();Object.assign(globalThis,{Office:h.Office,Word:h.Word,document:ui.document,location:{search:''}});await import('../src/app.mjs?guided-course-setup');await settle();const n=id=>ui.nodes.get(id),keys=['audience','purpose','authors','attribution'];assert.equal(n('course-title').value,'');for(const k of keys)assert.equal(n('course-setup-'+k).value,'');
  const html=fs.readFileSync(new URL('../public/taskpane.html',import.meta.url),'utf8'),guide=html.match(/<details id="course-setup-guidance">([\s\S]*?)<\/details>/)[1];assert.match(guide,/Read-only example/);assert.doesNotMatch(guide,/<(?:input|textarea|button)\b|\bon\w+=/i);assert.equal(n('course-setup-guidance').ontoggle,undefined);
  n('course-title').value='  Guided new course  ';for(const k of keys)n('course-setup-'+k).value=`  PRIVATE_SETUP_${k}\r\nSecond line  `;const values=()=>Object.fromEntries(['course-title',...keys.map(k=>'course-setup-'+k)].map(id=>[id,n(id).value])),before=values(),outputs=['workbook-preview','agenda-summary','materials-checklist-output'].map(id=>n(id).textContent),calls={reads:h.refreshes,saves:h.saves,runs:h.runs};
  for(const open of [true,false,true]){n('course-setup-guidance').open=open;n('course-setup-guidance').ontoggle?.({target:n('course-setup-guidance')});}assert.deepEqual(values(),before);assert.deepEqual(['workbook-preview','agenda-summary','materials-checklist-output'].map(id=>n(id).textContent),outputs);assert.deepEqual({reads:h.refreshes,saves:h.saves,runs:h.runs},calls);
  n('course-form').onsubmit({preventDefault(){}});await settle();const p=h.durable['openstage.guide.project.v1'];assert.equal(p.title,'Guided new course');assert.equal(p.revision,0);assert.deepEqual(p.modules,[]);for(const k of keys)assert.equal(p.details[k],`PRIVATE_SETUP_${k}\nSecond line`);assert.equal(h.saves,2);assert.equal(h.durable['openstage.guide.pending.v1'].kind,'course');assert.equal(h.durable['openstage.guide.pending.v1'].phase,'committed');assert.equal(h.records.length,1);assert.equal(h.paragraphs.length,1);assert.deepEqual(values(),before);assert.equal(n('course-form').resetCount,undefined);
});

test('setup validation and stale or existing-course refusal retain entered fields without changing Word content',async()=>{
  const h=makeHost(null),ui=dom();Object.assign(globalThis,{Office:h.Office,Word:h.Word,document:ui.document,location:{search:''}});await import('../src/app.mjs?guided-course-invalid-stale');await settle();const n=id=>ui.nodes.get(id),keys=['audience','purpose','authors','attribution'];for(const k of keys)n('course-setup-'+k).value=`PRIVATE_UNSAVED_${k}`;
  n('course-form').onsubmit({preventDefault(){}});await settle();assert.match(n('error').textContent,/course title/);assert.equal(h.saves,0);n('course-title').value='Retained new title';n('course-setup-purpose').value='x'.repeat(50001);n('course-form').onsubmit({preventDefault(){}});await settle();assert.match(n('error').textContent,/too long/);assert.equal(n('course-setup-purpose').value.length,50001);assert.equal(h.records.length,0);assert.equal(h.saves,0);n('course-setup-purpose').value='PRIVATE_UNSAVED_purpose';
  const existing=fixture();h.durable['openstage.guide.project.v1']=structuredClone(existing);h.records=[{tag:`osg:course:${existing.id}`,text:existing.title},...existing.modules.flatMap(m=>[{tag:`osg:module:${m.id}`,text:m.title},{tag:`osg:objective:${m.id}`,text:m.objective},...m.activities.flatMap(a=>['activity','facilitator','participant','answers'].map(type=>({tag:`osg:${type}:${a.id}`,text:type==='activity'?a.title:a[type]})))])];const saved=structuredClone(h.durable),body=structuredClone(h.records);
  n('course-form').onsubmit({preventDefault(){}});await settle();assert.match(n('error').textContent,/changed/);assert.deepEqual(h.durable,saved);assert.deepEqual(h.records,body);assert.equal(h.saves,0);for(const k of keys)assert.equal(n('course-setup-'+k).value,`PRIVATE_UNSAVED_${k}`);assert.equal(n('course-title').value,'Retained new title');assert.equal(n('course-form').resetCount,undefined);
  await n('refresh').onclick();const counts={refreshes:h.refreshes,saves:h.saves,runs:h.runs,syncs:h.syncs};n('course-form').onsubmit({preventDefault(){}});await settle();assert.match(n('error').textContent,/already has a course/);assert.deepEqual({refreshes:h.refreshes,saves:h.saves,runs:h.runs,syncs:h.syncs},counts);assert.deepEqual(h.records,body);assert.deepEqual(h.durable,saved);
});

test('uncertain setup final save keeps every form entry and recovers complete revision-zero details without title replay',async()=>{
  for(const persisted of [false,true]){const h=makeHost(null),ui=dom();Object.assign(globalThis,{Office:h.Office,Word:h.Word,document:ui.document,location:{search:''}});await import(`../src/app.mjs?guided-course-uncertain-${persisted}`);await settle();const n=id=>ui.nodes.get(id),keys=['audience','purpose','authors','attribution'];n('course-title').value='Retained setup title';for(const k of keys)n('course-setup-'+k).value=`PRIVATE_RETAIN_${k}`;h.saveFailures.set(h.saves+2,{persist:persisted});n('course-form').onsubmit({preventDefault(){}});await settle();assert.equal(n('course-form').resetCount,undefined);assert.equal(n('course-title').value,'Retained setup title');for(const k of keys)assert.equal(n('course-setup-'+k).value,`PRIVATE_RETAIN_${k}`);assert.equal(h.records.length,1);assert.equal(h.paragraphs.length,1);assert.equal(n('course-form').querySelector('button').disabled,true);
    await n('recover').onclick();const p=h.durable['openstage.guide.project.v1'];assert.equal(p.revision,0);assert.deepEqual(p.modules,[]);for(const k of keys)assert.equal(p.details[k],`PRIVATE_RETAIN_${k}`);const saves=h.saves;await n('recover').onclick();assert.equal(h.saves,saves);assert.equal(h.records.length,1);assert.equal(h.paragraphs.length,1);for(const k of keys)assert.equal(n('course-setup-'+k).value,`PRIVATE_RETAIN_${k}`);
  }
});

test('Wrench selector creates only selected private sections and leaves learning destinations empty without populating private detail entries',async()=>{
 const h=frontHost(makeHost(null)),ui=dom();Object.assign(globalThis,{Office:h.Office,Word:h.Word,document:ui.document,location:{search:''}});await import('../src/app.mjs?wrench-selector');await settle();const n=id=>ui.nodes.get(id),html=fs.readFileSync(new URL('../public/taskpane.html',import.meta.url),'utf8');assert.match(html,/<option value="minimal" selected>/);assert.match(html,/<fieldset id="course-frontmatter" hidden>/);// Boolean HTML defaults are checked in the markup; this fixture models subsequent events.
 n('course-starter').value='wrench';n('course-starter').onchange();assert.equal(n('course-frontmatter').hidden,false);n('course-section-overview').checked=true;n('course-section-background').checked=true;n('course-title').value='Private Wrench starter';n('course-setup-audience').value='PRIVATE UI AUDIENCE';n('course-starter').value='minimal';n('course-starter').onchange();assert.equal(n('course-frontmatter').hidden,true);assert.equal(n('course-section-overview').checked,true);n('course-starter').value='wrench';n('course-starter').onchange();n('course-form').onsubmit({preventDefault(){}});await settleWrench(ui);const p=h.durable['openstage.guide.project.v1'];assert.equal(p.schema,2);assert.deepEqual(p.frontMatter.sections.map(s=>s.kind),['opening','background','overview']);assert.ok(p.frontMatter.sections.every(s=>s.body===''));assert.deepEqual(p.modules,[]);assert.equal(h.records.length,7);assert.equal(h.saves,2);assert.equal(n('activity-module').children.length,0);assert.equal(n('course-setup-audience').value,'PRIVATE UI AUDIENCE');assert.equal(n('course-section-overview').checked,true);assert.equal(n('course-form').resetCount,undefined);assert.equal(n('course-setup-authors').value,'');
});

test('Wrench blank refusal and uncertain final save retain starter selections and all draft fields for recovery',async()=>{
 const h=frontHost(makeHost(null)),ui=dom();Object.assign(globalThis,{Office:h.Office,Word:h.Word,document:ui.document,location:{search:''}});await import('../src/app.mjs?wrench-refusal-recovery');await settle();const n=id=>ui.nodes.get(id);n('course-starter').value='wrench';n('course-section-team').checked=true;n('course-title').value='Retain Wrench draft';n('course-setup-attribution').value='PRIVATE DRAFT ATTRIBUTION';h.extraBody=paragraph('MANUAL USER TEXT');n('course-form').onsubmit({preventDefault(){}});await settleWrench(ui);assert.match(n('error').textContent,/new blank/);assert.equal(h.saves,0);assert.equal(h.records.length,0);assert.match(h.extraBody,/MANUAL USER TEXT/);assert.equal(n('course-title').value,'Retain Wrench draft');assert.equal(n('course-section-team').checked,true);
 h.extraBody=undefined;h.saveFailures.set(2,{persist:false});n('course-form').onsubmit({preventDefault(){}});await settleWrench(ui);assert.equal(n('course-form').querySelector('button').disabled,true);assert.equal(n('course-starter').value,'wrench');assert.equal(n('course-setup-attribution').value,'PRIVATE DRAFT ATTRIBUTION');assert.equal(n('course-section-team').checked,true);assert.equal(h.records.length,5);await n('recover').onclick();assert.equal(h.durable['openstage.guide.project.v1'].schema,2);assert.equal(h.durable['openstage.guide.pending.v1'].schema,5);assert.equal(h.records.length,5);assert.equal(n('course-form').resetCount,undefined);assert.equal(n('course-section-team').checked,true);
});

test('private facilitator copy requires exact consent and binds all four content choices to preview stamp output and unique generic filenames',async()=>{
 const p=fixture(),a=p.modules[0].activities[0];a.materials=[{item:'Sheets',quantity:1,unit:'sets',use:'consumable',notes:'PRIVATE PREPARATION'}];p.details={authors:'Authored by Jason',attribution:'Authored attribution'};const h=makeHost(p),ui=dom();Object.assign(globalThis,{Office:h.Office,Word:h.Word,document:ui.document,location:{search:''},JSZip});const blobs=[],links=[],create=ui.document.createElement,oldURL=URL.createObjectURL,oldTimer=setTimeout;ui.document.createElement=tag=>{const e=create(tag);if(tag==='a')links.push(e);return e;};URL.createObjectURL=b=>{blobs.push(b);return 'blob:private-guide';};globalThis.setTimeout=()=>0;
 try{await import('../src/app.mjs?private-guide-choices');await settle();const n=id=>ui.nodes.get(id),html=fs.readFileSync(new URL('../public/taskpane.html',import.meta.url),'utf8');for(const id of ['facilitator-answers','facilitator-participant']){assert.doesNotMatch(html.match(new RegExp(`<input id="${id}"[^>]*>`))[0],/\bchecked\b/);assert.ok(!n(id).checked);}await n('facilitator-download').onclick();assert.equal(blobs.length,0);n('course-detail-purpose').value='PRIVATE UNSAVED COURSE';n('activity-facilitator').value='PRIVATE UNSAVED TEACHER';
  for(const answers of [false,true])for(const participant of [false,true]){n('facilitator-answers').checked=answers;n('facilitator-participant').checked=participant;n('facilitator-answers').onchange();await n('facilitator-preview').onclick();assert.equal(n('facilitator-download').disabled,true);const preview=n('facilitator-output').textContent,stamp=preview.match(/Generated: ([^\n]+)/)[1],expected=facilitatorParagraphs(facilitatorPreview(p,{answers,participant},stamp));assert.equal(preview,expected.map(([,s])=>s).join('\n'));n('facilitator-reviewed').checked=true;n('facilitator-reviewed').onchange();assert.equal(n('facilitator-download').disabled,false);await n('facilitator-download').onclick();const {text}=await auditFacilitator(blobs.at(-1),{expectedRows:expected,stamp,forbidden:['PRIVATE UNSAVED COURSE','PRIVATE UNSAVED TEACHER',...answers?[]:['SECRET ANSWER'],...participant?[]:['Learner task.']]});assert.equal(text,preview);assert.equal(text.includes('SECRET ANSWER'),answers);assert.equal(text.includes('Learner task.'),participant);assert.equal(n('course-detail-purpose').value,'PRIVATE UNSAVED COURSE');assert.equal(n('activity-facilitator').value,'PRIVATE UNSAVED TEACHER');}
  assert.equal(blobs.length,4);assert.equal(new Set(links.map(l=>l.download)).size,4);assert.ok(links.every(l=>/^OpenStage Private Facilitator Guide - [a-f0-9-]{36}\.docx$/.test(l.download)));assert.equal(h.saves,0);assert.equal(h.paragraphs.length,0);assert.equal(h.cueWrites,0);
 }finally{URL.createObjectURL=oldURL;globalThis.setTimeout=oldTimer;globalThis.JSZip=JSZip;}
});

test('private-copy async changes to choices consent preview or saved private fields withhold offered download and retain raw drafts',async()=>{
 const p=fixture(),h=makeHost(p),ui=dom();Object.assign(globalThis,{Office:h.Office,Word:h.Word,document:ui.document,location:{search:''},JSZip});const blobs=[],oldURL=URL.createObjectURL,oldTimer=setTimeout;URL.createObjectURL=b=>{blobs.push(b);return 'blob:private-async';};globalThis.setTimeout=()=>0;
 try{await import('../src/app.mjs?private-guide-async');await settle();const n=id=>ui.nodes.get(id),word=await import('../src/word.mjs');n('course-detail-purpose').value='PRIVATE UNFINISHED';n('activity-answers').value='PRIVATE RAW ANSWER';n('material-notes').value='PRIVATE RAW MATERIAL';const review=async()=>{globalThis.JSZip=JSZip;await n('facilitator-preview').onclick();n('facilitator-reviewed').checked=true;n('facilitator-reviewed').onchange();};
  for(const change of ['answers','participant','choice-without-event','consent','preview-text','teacher','unselected-answer','private-details','pack-failure']){await review();let release;globalThis.JSZip=class extends JSZip{generateAsync(options){return new Promise((resolve,reject)=>{release=()=>change==='pack-failure'?reject(new Error('Injected private-copy packaging failure')):super.generateAsync(options).then(resolve,reject);});}};const pending=n('facilitator-download').onclick();await settle();assert.ok(release);
   if(change==='answers'||change==='participant'){const id='facilitator-'+change;n(id).checked=!n(id).checked;n(id).onchange();}else if(change==='choice-without-event')n('facilitator-answers').checked=!n('facilitator-answers').checked;else if(change==='consent'){n('facilitator-reviewed').checked=false;n('facilitator-reviewed').onchange();}else if(change==='preview-text')n('facilitator-output').textContent='Changed displayed preview';else if(change==='teacher')h.records.find(r=>r.tag.startsWith('osg:facilitator:')).text='NEW PRIVATE TEACHING';else if(change==='unselected-answer')h.records.find(r=>r.tag.startsWith('osg:answers:')).text='NEW UNSELECTED PRIVATE ANSWER';else if(change==='private-details')await word.updateCourseDetails(await word.readProject(),{audience:'PRIVATE NEW AUDIENCE',purpose:'',authors:'',attribution:''});
   release();await pending;assert.equal(blobs.length,0);assert.equal(n('facilitator-download').disabled,true);assert.equal(n('facilitator-reviewed').checked,false);assert.equal(n('course-detail-purpose').value,'PRIVATE UNFINISHED');assert.equal(n('activity-answers').value,'PRIVATE RAW ANSWER');assert.equal(n('material-notes').value,'PRIVATE RAW MATERIAL');
  }
  await review();await n('facilitator-download').onclick();assert.equal(blobs.length,1);assert.equal(h.paragraphs.length,0);assert.equal(h.cueWrites,0);
 }finally{URL.createObjectURL=oldURL;globalThis.setTimeout=oldTimer;globalThis.JSZip=JSZip;}
});

test('invalid selected private text refuses preview without touching source or drafts while unselected answers remain excluded',async()=>{
 const p=fixture(),a=p.modules[0].activities[0];a.answers='EXCLUDED INVALID \ufffe';const h=makeHost(p),ui=dom();Object.assign(globalThis,{Office:h.Office,Word:h.Word,document:ui.document,location:{search:''},JSZip});await import('../src/app.mjs?private-guide-invalid');await settle();const n=id=>ui.nodes.get(id);n('activity-facilitator').value='UNSAVED PRIVATE TEACHER';await n('facilitator-preview').onclick();assert.match(n('facilitator-output').textContent,/PRIVATE DIRECTIONS/);assert.doesNotMatch(n('facilitator-output').textContent,/EXCLUDED INVALID/);const before=structuredClone(h.durable);n('facilitator-answers').checked=true;n('facilitator-answers').onchange();await n('facilitator-preview').onclick();assert.match(n('error').textContent,/selected answers\/debrief contains an unsupported character/);assert.equal(n('facilitator-download').disabled,true);assert.equal(n('activity-facilitator').value,'UNSAVED PRIVATE TEACHER');assert.deepEqual(h.durable,before);assert.equal(h.saves,0);assert.equal(h.paragraphs.length,0);assert.equal(n('facilitator-answers').checked,true);
});

test('integrated materials checklist requires reviewed saved preview and exports its exact stamp and supplies as new UUID files',async()=>{
  const p=fixture();p.modules[0].activities[0].materials=[{item:'Saved supply <&>',quantity:1.125,unit:'Sets',use:'reusable',notes:'PRIVATE PREPARATION'}];p.details={audience:'PRIVATE AUDIENCE'};const h=makeHost(p),ui=dom();Object.assign(globalThis,{Office:h.Office,Word:h.Word,document:ui.document,location:{search:''},JSZip});installParser();
  const blobs=[],links=[],create=ui.document.createElement,oldURL=URL.createObjectURL,oldTimer=setTimeout;ui.document.createElement=tag=>{const e=create(tag);if(tag==='a')links.push(e);return e;};URL.createObjectURL=b=>{blobs.push(b);return 'blob:materials-test';};globalThis.setTimeout=()=>0;
  try{
    await import('../src/app.mjs?materials-checklist-ui');await settle();const n=id=>ui.nodes.get(id);await n('materials-load').onclick();n('material-item').value='UNSAVED PRIVATE SUPPLY';n('material-notes').value='PRIVATE RAW ENTRY';
    await n('materials-checklist-download').onclick();assert.equal(blobs.length,0);await n('materials-checklist-preview').onclick();assert.equal(n('materials-checklist-download').disabled,true);assert.match(n('materials-checklist-output').textContent,/Saved supply <&>/);assert.doesNotMatch(n('materials-checklist-output').textContent,/PRIVATE|UNSAVED/);
    const preview=n('materials-checklist-output').textContent;n('materials-checklist-reviewed').checked=true;n('materials-checklist-reviewed').onchange();assert.equal(n('materials-checklist-download').disabled,false);await n('materials-checklist-download').onclick();assert.equal(blobs.length,1);
    const z=await JSZip.loadAsync(await blobs[0].arrayBuffer()),doc=new DOMParser().parseFromString(await z.file('word/document.xml').async('string'),'application/xml'),W='http://schemas.openxmlformats.org/wordprocessingml/2006/main',text=Array.from(doc.getElementsByTagNameNS(W,'p')).map(p=>Array.from(p.getElementsByTagNameNS(W,'t')).map(t=>t.textContent).join('\n')).join('\n');assert.equal(text,preview);await auditExport(blobs[0],{project:p,privateCanaries:['PRIVATE PREPARATION','PRIVATE AUDIENCE','PRIVATE RAW ENTRY','UNSAVED PRIVATE SUPPLY','PRIVATE DIRECTIONS','SECRET ANSWER','Learner task.'],kind:'materials'});
    await n('materials-checklist-download').onclick();assert.equal(blobs.length,2);assert.notEqual(links[0].download,links[1].download);assert.match(links[0].download,/^OpenStage Materials Checklist - [a-f0-9-]{36}\.docx$/);assert.equal(n('material-notes').value,'PRIVATE RAW ENTRY');assert.equal(n('material-item').value,'UNSAVED PRIVATE SUPPLY');assert.equal(h.paragraphs.length,0);assert.equal(h.cueWrites,0);
  }finally{URL.createObjectURL=oldURL;globalThis.setTimeout=oldTimer;globalThis.JSZip=JSZip;}
});

test('materials checklist review changes or saved supply edits during packaging withhold download and invalidate preview without losing drafts',async()=>{
  const p=fixture();p.modules[0].activities[0].materials=[{item:'Saved sheets',quantity:0.125,unit:'sets',use:'consumable',notes:'PRIVATE NOTES'}];const h=makeHost(p),ui=dom();Object.assign(globalThis,{Office:h.Office,Word:h.Word,document:ui.document,location:{search:''},JSZip});const blobs=[],oldURL=URL.createObjectURL,oldTimer=setTimeout;URL.createObjectURL=b=>{blobs.push(b);return 'blob:materials-async';};globalThis.setTimeout=()=>0;
  try{
    await import('../src/app.mjs?materials-checklist-async');await settle();const n=id=>ui.nodes.get(id);await n('materials-load').onclick();n('material-quantity').value='Unfinished quantity';n('material-notes').value='PRIVATE UNSAVED NOTES';const review=async()=>{await n('materials-checklist-preview').onclick();n('materials-checklist-reviewed').checked=true;n('materials-checklist-reviewed').onchange();};
    for(const change of ['review','saved-materials']){globalThis.JSZip=JSZip;await review();let release;globalThis.JSZip=class extends JSZip{generateAsync(o){return new Promise(resolve=>{release=()=>super.generateAsync(o).then(resolve);});}};const pending=n('materials-checklist-download').onclick();await settle();assert.ok(release);if(change==='review'){n('materials-checklist-reviewed').checked=false;n('materials-checklist-reviewed').onchange();}else{const word=await import('../src/word.mjs'),base=await word.readProject();await word.updateActivityMaterials(base,base.modules[0].activities[0].id,[{item:'Revised saved sheets',quantity:0.375,unit:'sets',use:'consumable',notes:'NEW PRIVATE NOTES'}]);}release();await pending;assert.equal(blobs.length,0);assert.equal(n('materials-checklist-download').disabled,true);assert.equal(n('materials-checklist-reviewed').checked,false);assert.equal(n('material-quantity').value,'Unfinished quantity');assert.equal(n('material-notes').value,'PRIVATE UNSAVED NOTES');}
    globalThis.JSZip=JSZip;await review();await n('materials-checklist-download').onclick();assert.equal(blobs.length,1);assert.match(n('materials-checklist-output').textContent,/Revised saved sheets/);assert.equal(h.paragraphs.length,0);assert.equal(h.cueWrites,0);
  }finally{URL.createObjectURL=oldURL;globalThis.setTimeout=oldTimer;globalThis.JSZip=JSZip;}
});

test('empty saved materials checklist is reviewable and explicitly uncertified',async()=>{
  const h=makeHost(fixture()),ui=dom();Object.assign(globalThis,{Office:h.Office,Word:h.Word,document:ui.document,location:{search:''},JSZip});await import('../src/app.mjs?materials-checklist-empty');await settle();const n=id=>ui.nodes.get(id);await n('materials-checklist-preview').onclick();assert.match(n('materials-checklist-output').textContent,/0 of 1 activities/);assert.match(n('materials-checklist-output').textContent,/No materials recorded/);assert.match(n('materials-checklist-output').textContent,/may still need materials/);assert.equal(n('materials-checklist-download').disabled,true);
});

test('invalid XML checklist text refuses preview and download while saved source and authoring drafts remain intact',async()=>{
  const blobs=[],oldURL=URL.createObjectURL,oldTimer=setTimeout;URL.createObjectURL=b=>{blobs.push(b);return 'blob:invalid-materials';};globalThis.setTimeout=()=>0;
  try{
    for(const invalid of ['\ufffe','\uffff','\ud800','\udfff'])for(const field of ['item','unit','activity']){
      const p=fixture(),a=p.modules[0].activities[0];a.materials=[{item:'Saved supply',quantity:1,unit:'sets',use:'consumable',notes:'PRIVATE NOTE'}];if(field==='activity')a.title='Title '+invalid+' inside';else a.materials[0][field]='Saved '+invalid+' inside';const h=makeHost(p),ui=dom();Object.assign(globalThis,{Office:h.Office,Word:h.Word,document:ui.document,location:{search:''},JSZip});await import(`../src/app.mjs?materials-invalid-${field}-${invalid.charCodeAt(0)}`);await settle();const n=id=>ui.nodes.get(id);await n('materials-load').onclick();n('material-item').value='UNSAVED SUPPLY';n('material-notes').value='PRIVATE UNSAVED NOTE';const source=structuredClone(h.durable),body=structuredClone(h.records),saves=h.saves;
      await n('materials-checklist-preview').onclick();assert.match(n('error').textContent,/unsupported character/);assert.equal(n('materials-checklist-reviewed').checked,false);assert.equal(n('materials-checklist-download').disabled,true);await n('materials-checklist-download').onclick();assert.equal(blobs.length,0);assert.equal(n('material-item').value,'UNSAVED SUPPLY');assert.equal(n('material-notes').value,'PRIVATE UNSAVED NOTE');assert.deepEqual(h.durable,source);assert.deepEqual(h.records,body);assert.equal(h.saves,saves);assert.equal(h.paragraphs.length,0);assert.equal(h.cueWrites,0);assert.equal(n('refresh').disabled,false);
    }
  }finally{URL.createObjectURL=oldURL;globalThis.setTimeout=oldTimer;globalThis.JSZip=JSZip;}
});

test('invalid selected XML text blocks workbook and agenda previews without clearing drafts or changing saved records',async()=>{
  const blobs=[],oldURL=URL.createObjectURL,oldTimer=setTimeout;URL.createObjectURL=b=>{blobs.push(b);return 'blob:invalid-public';};globalThis.setTimeout=()=>0;
  try{for(const bad of ['\ufffe','\uffff','\ud800','\udfff']){
    const p=fixture();p.modules[0].activities[0].title='Invalid '+bad+' title';const h=makeHost(p),ui=dom();Object.assign(globalThis,{Office:h.Office,Word:h.Word,document:ui.document,location:{search:''},JSZip});await import(`../src/app.mjs?public-invalid-${bad.charCodeAt(0)}`);await settle();const n=id=>ui.nodes.get(id);n('activity-participant').value='UNSAVED PUBLIC DRAFT';n('activity-facilitator').value='UNSAVED PRIVATE DRAFT';await n('materials-load').onclick();n('material-notes').value='UNSAVED MATERIAL NOTES';const saved=structuredClone(h.durable),records=structuredClone(h.records),saves=h.saves;
    await n('preview').onclick();assert.match(n('error').textContent,/unsupported character/);assert.equal(n('download').disabled,true);await n('download').onclick();n('agenda-form').onsubmit({preventDefault(){}});await settle();assert.match(n('error').textContent,/unsupported character/);assert.equal(n('agenda-download').disabled,true);await n('agenda-download').onclick();assert.equal(blobs.length,0);assert.equal(n('activity-participant').value,'UNSAVED PUBLIC DRAFT');assert.equal(n('activity-facilitator').value,'UNSAVED PRIVATE DRAFT');assert.equal(n('material-notes').value,'UNSAVED MATERIAL NOTES');assert.deepEqual(h.durable,saved);assert.deepEqual(h.records,records);assert.equal(h.saves,saves);assert.equal(h.paragraphs.length,0);assert.equal(n('refresh').disabled,false);
  }}finally{URL.createObjectURL=oldURL;globalThis.setTimeout=oldTimer;globalThis.JSZip=JSZip;}
});
test('activity destination survives refresh and uncertain commit; a new module becomes the default',async()=>{
  const p=addModule(fixture(),'Later module','').next,h=makeHost(p),ui=dom();
  globalThis.Office=h.Office;globalThis.Word=h.Word;globalThis.document=ui.document;globalThis.location={search:''};
  globalThis.fetch=async()=>({ok:true,blob:async()=>new Blob(['mock icon'])});
  globalThis.FileReader=class{readAsDataURL(){this.result='data:image/png;base64,AAAA';queueMicrotask(()=>this.onload());}};
  await import('../src/app.mjs?module-placement');await settle();const n=id=>ui.nodes.get(id);
  assert.equal(n('activity-module').value,p.modules[1].id);
  n('activity-module').value=p.modules[0].id;n('activity-module').onchange();
  n('activity-title').value='Earlier activity';n('activity-participant').value='Earlier learner task';
  await n('refresh').onclick();assert.equal(n('activity-module').value,p.modules[0].id);
  const commitSave=h.saves+2;h.saveFailures.set(commitSave,{persist:true});
  n('activity-form').onsubmit({preventDefault(){}});await settle();
  assert.equal(n('activity-module').value,p.modules[0].id);assert.equal(n('activity-form').resetCount,undefined);
  assert.equal(n('activity-module').disabled,true);assert.equal(h.paragraphs.length,9);
  await n('recover').onclick();assert.equal(n('activity-module').disabled,false);
  assert.equal(n('activity-module').value,p.modules[0].id);
  assert.equal(h.durable['openstage.guide.project.v1'].modules[0].activities.at(-1).title,'Earlier activity');
  n('module-title').value='Newest module';n('module-objective').value='';
  n('module-form').onsubmit({preventDefault(){}});await settle();
  assert.equal(n('activity-module').value,h.durable['openstage.guide.project.v1'].modules.at(-1).id);
});
test('outline filters stay local, preserve IDs/destination and refresh direct edits while guarding navigation',async()=>{
  const p=addModule(fixture(),'Later module','').next,h=makeHost(p),ui=dom();globalThis.Office=h.Office;globalThis.Word=h.Word;globalThis.document=ui.document;globalThis.location={search:''};
  await import('../src/app.mjs?outline-search-ui');await settle();const n=id=>ui.nodes.get(id),a=p.modules[0].activities[0];n('activity-module').value=p.modules[0].id;
  const counts={saves:h.saves,syncs:h.syncs,runs:h.runs,refreshes:h.refreshes};n('outline-query').value='PRIVATE DIRECTIONS';n('outline-query').oninput();assert.equal(n('outline').children.length,2);assert.equal(n('outline').children[1].dataset.outlineId,a.id);assert.ok(!n('outline').textContent.includes('PRIVATE DIRECTIONS'));assert.match(n('outline-results').textContent,/1 of 2 modules/);assert.equal(n('activity-module').value,p.modules[0].id);
  n('outline-icon').value='Discussion';n('outline-icon').onchange();assert.match(n('outline').textContent,/No items match/);n('outline-clear').onclick();assert.equal(n('outline-query').value,'');assert.equal(n('outline-status').value,'all');assert.equal(n('outline').children.length,3);
  assert.deepEqual({saves:h.saves,syncs:h.syncs,runs:h.runs,refreshes:h.refreshes},counts);
  const cachedButton=n('outline').children.find(b=>b.dataset.outlineId===a.id);await cachedButton.onclick();assert.equal(h.selection,`osg:activity:${a.id}`);
  n('outline-query').value='New direct learner text';n('outline-query').oninput();assert.match(n('outline').textContent,/No items match/);h.records.find(r=>r.tag===`osg:participant:${a.id}`).text='New direct learner text';await n('refresh').onclick();assert.equal(n('outline').children[1].dataset.outlineId,a.id);assert.equal(n('outline-query').value,'New direct learner text');assert.equal(n('activity-module').value,p.modules[0].id);
  h.refreshFailures.add(h.refreshes+1);await n('refresh').onclick();assert.equal(n('outline-query').disabled,true);assert.equal(n('outline').children.length,0);const paused={saves:h.saves,syncs:h.syncs,runs:h.runs,refreshes:h.refreshes};await cachedButton.onclick();assert.deepEqual({saves:h.saves,syncs:h.syncs,runs:h.runs,refreshes:h.refreshes},paused);n('outline-clear').onclick();assert.equal(n('outline-query').value,'New direct learner text');
  h.refreshFailures.clear();await n('recover').onclick();assert.equal(n('outline-query').disabled,false);assert.equal(n('outline-query').value,'New direct learner text');assert.equal(h.paragraphs.length,0);assert.equal(h.cueWrites,0);
});
test('no-recorded-materials outline filter switches and clears without touching drafts or navigating when paused',async()=>{
 let p=fixture();p=addActivity(p,p.modules[0].id,{title:'Saved supplies task',icon:'Discussion',minutes:5,participant:'Try task',facilitator:'Teach task',answers:'Debrief',lines:1,materials:[{item:'Pencils',quantity:2,unit:'boxes',use:'consumable',notes:'PRIVATE SUPPLY NOTE'}]}).next;
 const h=makeHost(p),ui=dom(),n=id=>ui.nodes.get(id),first=p.modules[0].activities[0];Object.assign(globalThis,{Office:h.Office,Word:h.Word,document:ui.document,location:{search:''}});await import('../src/app.mjs?outline-materials-ui');await settle();
 const fields=['activity-title','material-item','library-name','plan-name','coverage-rationale','ppt-notes'];for(const id of fields)n(id).value=`RAW ${id}`;const raw=Object.fromEntries(fields.map(id=>[id,n(id).value])),counts=()=>({saves:h.saves,syncs:h.syncs,runs:h.runs});const before=counts();
 n('outline-status').value='no-recorded-materials';n('outline-status').onchange();assert.equal(n('outline').children.length,2);assert.equal(n('outline').children[1].dataset.outlineId,first.id);assert.match(n('outline-results').textContent,/does not mean no supplies/);assert.doesNotMatch(n('outline').textContent,/PRIVATE SUPPLY NOTE/);
 n('outline-icon').value='Discussion';n('outline-icon').onchange();assert.match(n('outline').textContent,/No items match/);n('outline-icon').value='Learning activity';n('outline-icon').onchange();assert.equal(n('outline').children[1].dataset.outlineId,first.id);
 n('outline-query').value='Original';n('outline-query').oninput();assert.equal(n('outline').children[1].dataset.outlineId,first.id);n('outline-status').value='missing-objective';n('outline-status').onchange();assert.match(n('outline').textContent,/No items match/);n('outline-status').value='no-recorded-materials';n('outline-status').onchange();assert.equal(n('outline').children[1].dataset.outlineId,first.id);
 n('outline-clear').onclick();assert.equal(n('outline-status').value,'all');assert.equal(n('outline-query').value,'');assert.equal(n('outline-icon').value,'');assert.equal(n('outline').children.length,3);assert.deepEqual(Object.fromEntries(fields.map(id=>[id,n(id).value])),raw);assert.deepEqual(counts(),before);
 n('outline-status').value='no-recorded-materials';n('outline-status').onchange();const cached=n('outline').children[1];h.refreshFailures.add(h.refreshes+1);await n('refresh').onclick();assert.equal(n('outline-query').disabled,true);assert.equal(n('outline').children.length,0);const paused=counts();await cached.onclick();n('outline-clear').onclick();assert.deepEqual(counts(),paused);assert.equal(n('outline-status').value,'no-recorded-materials');assert.deepEqual(Object.fromEntries(fields.map(id=>[id,n(id).value])),raw);
 h.refreshFailures.clear();await n('recover').onclick();assert.equal(n('outline-status').value,'no-recorded-materials');assert.equal(n('outline').children[1].dataset.outlineId,first.id);assert.deepEqual(Object.fromEntries(fields.map(id=>[id,n(id).value])),raw);assert.deepEqual(h.durable['openstage.guide.project.v1'],p);
});
async function settleWrench(ui){for(let i=0;i<1000;i++){await new Promise(resolve=>setImmediate(resolve));if(ui.nodes.get('error').textContent||ui.nodes.get('course-form').hidden)return;}assert.fail('Synthetic Wrench operation did not settle');}
async function settle(){for(let i=0;i<12;i++)await new Promise(resolve=>setImmediate(resolve));}
test('actual library-content submit stays gated when nested material buttons precede it in markup',async()=>{
  const h=makeHost(null),ui=dom(),n=id=>ui.nodes.get(id),html=fs.readFileSync(new URL('../public/taskpane.html',import.meta.url),'utf8'),formMarkup=html.match(/<form id="library-content-form">([\s\S]+?)<\/form>/)[1];
  const buttons=[...formMarkup.matchAll(/<button\b([^>]*)>/g)].map(([,attrs])=>({id:attrs.match(/id="([^"]+)"/)?.[1],submit:/type="submit"/.test(attrs)}));assert.equal(buttons[0].id,'library-material-apply');assert.equal(buttons.find(b=>b.submit).id,'library-content-save');
  // Resolve the real first nested button instead of the generic synthetic submit.
  n('library-content-form').querySelector=selector=>selector==='button'?n(buttons[0].id):selector==='button[type="submit"]'?n(buttons.find(b=>b.submit).id):null;
  globalThis.Office=h.Office;globalThis.Word=h.Word;globalThis.document=ui.document;globalThis.location={search:''};await import('../src/app.mjs?library-submit-gating');await settle();assert.equal(n('library-content-save').disabled,true,'No draft');
  const {id,...fields}=fixture().modules[0].activities[0],catalog=saveLibraryCopy(emptyLibrary(),'First',[],fields),text=serializeLibrary(catalog);n('library-file').files=[{size:text.length,text:async()=>text}];n('library-file').onchange();await n('library-import').onclick();await n('library-content-load').onclick();assert.equal(n('library-material-apply').disabled,false);assert.equal(n('library-content-save').disabled,true,'Unreviewed');
  n('library-content-reviewed').checked=true;n('library-content-reviewed').onchange();assert.equal(n('library-content-save').disabled,false);n('library-material-notes').value='Unfinished raw entry';n('library-material-notes').oninput();assert.equal(n('library-content-save').disabled,true,'Unfinished material');
  await n('library-material-clear').onclick();n('library-content-reviewed').checked=true;n('library-content-reviewed').onchange();assert.equal(n('library-content-save').disabled,false);
  const incoming=serializeLibrary(saveLibraryCopy(emptyLibrary(),'Second',[],fields));let release;n('library-file').files=[{size:incoming.length,text:()=>new Promise(resolve=>{release=resolve;})}];n('library-file').onchange();const pending=n('library-import').onclick();await settle();assert.ok(release);assert.equal(n('library-content-save').disabled,true,'Busy');release(incoming);await pending;assert.equal(n('library-content-save').disabled,false);
  await n('library-content-discard').onclick();assert.equal(n('library-content-save').disabled,true,'Discarded');assert.equal(h.paragraphs.length,0);
});
test('Activity insert button is gated independently of preceding dynamic copied-material buttons',async()=>{
  const p=fixture(),h=makeHost(p),ui=dom(),n=id=>ui.nodes.get(id);globalThis.Office=h.Office;globalThis.Word=h.Word;globalThis.document=ui.document;globalThis.location={search:''};await import('../src/app.mjs?activity-real-submit-gating');await settle();
  const {id,...fields}=p.modules[0].activities[0];fields.materials=[{item:'Board',quantity:1,unit:'sets',use:'reusable',notes:'PRIVATE MATERIAL'}];const incoming=saveLibraryCopy(emptyLibrary(),'Copy',[],fields),text=serializeLibrary(incoming);n('library-file').files=[{size:text.length,text:async()=>text}];n('library-file').onchange();await n('library-import').onclick();await n('library-use').onclick();
  n('activity-form').querySelector=selector=>selector==='button'?n('activity-copied-materials-list').children[0].children.at(-1):selector==='button[type="submit"]'?n('activity-insert'):null;assert.equal(n('activity-form').querySelector('button').textContent,'Remove copied material');assert.equal(n('activity-insert').disabled,false);
  const next=serializeLibrary(saveLibraryCopy(emptyLibrary(),'Second',[],fields));let release;n('library-file').files=[{size:next.length,text:()=>new Promise(resolve=>{release=resolve;})}];n('library-file').onchange();const pending=n('library-import').onclick();await settle();assert.equal(n('activity-insert').disabled,true,'Busy actual insert');release(next);await pending;assert.equal(n('activity-insert').disabled,false);
  h.refreshFailures.add(h.refreshes+1);await n('refresh').onclick();assert.equal(n('activity-insert').disabled,true,'Paused actual insert');assert.equal(n('activity-participant').value,fields.participant);h.refreshFailures.clear();await n('recover').onclick();assert.equal(n('activity-insert').disabled,false);assert.equal(h.paragraphs.length,0);
});
test('stored-library editor works without a course, retains stale/raw drafts and is Word-free even while paused',async()=>{
  const h=makeHost(null),ui=dom();globalThis.Office=h.Office;globalThis.Word=h.Word;globalThis.document=ui.document;globalThis.location={search:''};
  const blobs=[],oldURL=URL.createObjectURL,oldTimer=setTimeout;URL.createObjectURL=b=>{blobs.push(b);return 'blob:stored-library';};globalThis.setTimeout=()=>0;
  try{
    await import('../src/app.mjs?stored-library-content');await settle();const n=id=>ui.nodes.get(id),{id,...fields}=fixture().modules[0].activities[0];fields.materials=[{item:'Private board',quantity:1,unit:'boards',use:'reusable',notes:'PRIVATE MATERIAL'}];
    const catalog=saveLibraryCopy(emptyLibrary(),'Stored',['tag'],fields),entry=catalog.entries[0];
    const importCatalog=async data=>{const text=serializeLibrary(data);n('library-file').files=[{size:text.length,text:async()=>text}];n('library-file').onchange();await n('library-import').onclick();};
    await importCatalog(catalog);const counts=()=>({saves:h.saves,syncs:h.syncs,runs:h.runs,refreshes:h.refreshes});const initial=counts(),insertion={title:'Unfinished insertion',participant:'UNSAVED INSERTION',facilitator:'PRIVATE UNFINISHED'};for(const [key,value] of Object.entries(insertion))n('activity-'+key).value=value;
    await n('library-content-load').onclick();assert.match(n('library-content-materials').textContent,/PRIVATE MATERIAL/);n('library-content-participant').value='[insert edited scenario]';n('library-content-facilitator').value='';n('library-content-answers').value='PRIVATE EDITED ANSWER';n('library-content-minutes').value='Not decided';
    const submit=async()=>{n('library-content-form').onsubmit({preventDefault(){}});await settle();};await submit();assert.match(n('error').textContent,/Review/);n('library-content-reviewed').checked=true;await submit();assert.equal(n('library-content-minutes').value,'Not decided');assert.match(n('error').textContent,/Unsupported/);
    await n('library-content-load').onclick();assert.match(n('error').textContent,/Preserve and discard/);assert.equal(n('library-content-participant').value,'[insert edited scenario]');
    // Unrelated imported entry must survive a loaded content draft.
    await importCatalog(saveLibraryCopy(emptyLibrary(),'Other',[],fields));n('library-item').value=entry.id;n('library-item').onchange();n('library-content-minutes').value='0';n('library-content-minutes').oninput();assert.equal(n('library-content-reviewed').checked,false);n('library-content-reviewed').checked=true;await submit();await n('library-export').onclick();let saved=JSON.parse(await blobs.at(-1).text());assert.equal(saved.entries.length,2);assert.equal(saved.entries[0].id,entry.id);assert.equal(saved.entries[0].fields.facilitator,'');assert.equal(saved.entries[0].fields.answers,'PRIVATE EDITED ANSWER');assert.deepEqual(saved.entries[0].fields.materials,fields.materials);assert.deepEqual(counts(),initial);
    for(const [key,value] of Object.entries(insertion))assert.equal(n('activity-'+key).value,value);
    n('library-content-minutes').value='Invalid retained draft';n('library-content-reviewed').checked=true;await submit();assert.match(n('error').textContent,/Unsupported/);assert.equal(n('library-content-minutes').value,'Invalid retained draft');assert.equal(n('library-content-undo').disabled,false);
    n('library-content-minutes').value='0';n('library-content-reviewed').checked=true;await submit();assert.equal(n('library-content-undo').disabled,false,'Unchanged save retains the last real edit undo');
    await n('library-content-undo').onclick();assert.match(n('error').textContent,/discard/);await n('library-content-discard').onclick();await n('library-content-undo').onclick();await n('library-export').onclick();saved=JSON.parse(await blobs.at(-1).text());assert.deepEqual(saved.entries[0],entry);assert.equal(saved.entries.length,2);
    await n('library-content-load').onclick();n('library-content-title').value='Retained stale draft';await n('library-manage-load').onclick();n('library-edit-name').value='Renamed while editing';n('library-manage-form').onsubmit({preventDefault(){}});await settle();n('library-content-reviewed').checked=true;await submit();assert.match(n('error').textContent,/changed/);assert.equal(n('library-content-title').value,'Retained stale draft');
    await n('library-remove').onclick();await submit();assert.equal(n('library-content-title').value,'Retained stale draft');await n('library-undo').onclick();assert.equal(n('library-content-title').value,'Retained stale draft');await n('library-content-discard').onclick();
    h.refreshFailures.add(h.refreshes+1);await n('refresh').onclick();const paused=counts(),journal=JSON.stringify(h.memory);assert.equal(n('library-content-load').disabled,false);await n('library-content-load').onclick();n('library-content-participant').value='Paused edit';n('library-content-reviewed').checked=true;await submit();await n('library-export').onclick();saved=JSON.parse(await blobs.at(-1).text());assert.equal(saved.entries[0].fields.participant,'Paused edit');await n('library-content-discard').onclick();await n('library-content-undo').onclick();assert.deepEqual(counts(),paused);assert.equal(JSON.stringify(h.memory),journal);
    h.refreshFailures.clear();await n('recover').onclick();assert.equal(h.paragraphs.length,0);assert.equal(h.cueWrites,0);
  }finally{URL.createObjectURL=oldURL;globalThis.setTimeout=oldTimer;}
});
test('stored-content edits do not propagate into loaded insertion drafts or existing Word copies',async()=>{
  const p=fixture(),h=makeHost(p),ui=dom();globalThis.Office=h.Office;globalThis.Word=h.Word;globalThis.document=ui.document;globalThis.location={search:''};globalThis.fetch=async()=>({ok:true,blob:async()=>new Blob(['mock icon'])});globalThis.FileReader=class{readAsDataURL(){this.result='data:image/png;base64,AAAA';queueMicrotask(()=>this.onload());}};
  await import('../src/app.mjs?stored-library-insertion-isolation');await settle();const n=id=>ui.nodes.get(id),{id,...fields}=p.modules[0].activities[0];fields.materials=[{item:'Board',quantity:1,unit:'sets',use:'reusable',notes:'PRIVATE SUPPLY'}];
  const catalog=saveLibraryCopy(emptyLibrary(),'Independent copy',['practice'],fields),entry=catalog.entries[0],text=serializeLibrary(catalog);n('library-file').files=[{size:text.length,text:async()=>text}];n('library-file').onchange();await n('library-import').onclick();await n('library-use').onclick();
  const insertion=Object.fromEntries(['title','participant','facilitator','answers','icon','minutes','lines'].map(k=>[k,n('activity-'+k).value])),note=n('activity-library-note').textContent,counts={saves:h.saves,syncs:h.syncs,runs:h.runs,refreshes:h.refreshes};
  await n('library-content-load').onclick();n('library-content-title').value='Edited reusable title';n('library-content-participant').value='[insert future scenario]';n('library-content-facilitator').value='PRIVATE NEW TEACHER';await n('library-content-materials').children[0].children[1].onclick();n('library-material-quantity').value='2';n('library-material-quantity').oninput();await n('library-material-apply').onclick();n('library-content-reviewed').checked=true;n('library-content-form').onsubmit({preventDefault(){}});await settle();
  for(const [key,value] of Object.entries(insertion))assert.equal(n('activity-'+key).value,value);assert.equal(n('activity-library-note').textContent,note);assert.deepEqual({saves:h.saves,syncs:h.syncs,runs:h.runs,refreshes:h.refreshes},counts);
  n('activity-form').onsubmit({preventDefault(){}});await settle();const saved=h.durable['openstage.guide.project.v1'],copy=saved.modules[0].activities.at(-1);assert.equal(saved.modules[0].activities.length,2);assert.notEqual(copy.id,id);assert.notEqual(copy.id,entry.id);assert.equal(copy.title,insertion.title);assert.equal(copy.librarySource.entryId,entry.id);assert.deepEqual(copy.materials,fields.materials);assert.equal(saved.modules[0].activities[0].title,p.modules[0].activities[0].title);
  await n('library-use').onclick();assert.equal(n('activity-title').value,'Edited reusable title');assert.equal(n('activity-participant').value,'[insert future scenario]');const saves=h.saves;n('activity-form').onsubmit({preventDefault(){}});await settle();assert.equal(h.saves,saves);assert.match(n('error').textContent,/unfinished prompts/);assert.equal(n('activity-participant').value,'[insert future scenario]');
  n('activity-participant').value='Reviewed future scenario';n('activity-form').onsubmit({preventDefault(){}});await settle();const final=h.durable['openstage.guide.project.v1'];assert.equal(final.modules[0].activities.at(-1).materials[0].quantity,2);assert.equal(final.modules[0].activities[1].materials[0].quantity,1);
});
test('stored-material editor preserves unfinished raw entries, rejects stale row controls and saves private rows with full undo while paused',async()=>{
  const h=makeHost(null),ui=dom();globalThis.Office=h.Office;globalThis.Word=h.Word;globalThis.document=ui.document;globalThis.location={search:''};const blobs=[],oldURL=URL.createObjectURL,oldTimer=setTimeout;URL.createObjectURL=b=>{blobs.push(b);return 'blob:stored-materials';};globalThis.setTimeout=()=>0;
  try{
    await import('../src/app.mjs?stored-material-editor');await settle();const n=id=>ui.nodes.get(id),{id,...fields}=fixture().modules[0].activities[0],library=saveLibraryCopy(emptyLibrary(),'Supply draft',[],fields),entry=library.entries[0],text=serializeLibrary(library);n('library-file').files=[{size:text.length,text:async()=>text}];n('library-file').onchange();await n('library-import').onclick();await n('library-content-load').onclick();
    const counts=()=>({saves:h.saves,syncs:h.syncs,runs:h.runs,refreshes:h.refreshes}),initial=counts(),submit=async()=>{n('library-content-form').onsubmit({preventDefault(){}});await settle();};
    for(const [key,value] of Object.entries({item:'Board',quantity:'undecided',unit:'sets',use:'reusable',notes:'PRIVATE RAW PREPARATION'})){n('library-material-'+key).value=value;n('library-material-'+key).oninput();}
    await n('library-material-apply').onclick();assert.equal(n('library-material-quantity').value,'undecided');assert.equal(n('library-content-materials').children.length,0);n('library-content-reviewed').checked=true;await submit();assert.match(n('error').textContent,/unfinished material/);await n('library-export').onclick();assert.equal(Object.hasOwn(JSON.parse(await blobs.at(-1).text()).entries[0].fields,'materials'),false);
    n('library-material-quantity').value='1';n('library-material-quantity').oninput();await n('library-material-apply').onclick();assert.equal(n('library-content-reviewed').checked,false);const staleRemove=n('library-content-materials').children[0].children[2];
    n('library-material-item').value='Handout';n('library-material-unit').value='copies';n('library-material-notes').value='PRIVATE ACCEPTED SUPPLY';n('library-material-item').oninput();await n('library-material-apply').onclick();await staleRemove.onclick();assert.match(n('error').textContent,/changed/);assert.equal(n('library-content-materials').children.length,2);
    await n('library-content-materials').children[0].children[1].onclick();n('library-material-quantity').value='bad';n('library-material-quantity').oninput();await n('library-material-apply').onclick();assert.equal(n('library-material-quantity').value,'bad');await n('library-content-materials').children[1].children[2].onclick();assert.equal(n('library-content-materials').children.length,2);assert.match(n('error').textContent,/unfinished/);await n('library-material-clear').onclick();assert.equal(n('library-material-quantity').value,'1');
    h.refreshFailures.add(h.refreshes+1);await n('refresh').onclick();const paused=counts(),journal=JSON.stringify(h.memory);n('library-content-reviewed').checked=true;await submit();await n('library-export').onclick();const saved=JSON.parse(await blobs.at(-1).text());assert.equal(saved.entries[0].id,entry.id);assert.equal(saved.entries[0].fields.materials.length,2);assert.match(JSON.stringify(saved),/PRIVATE ACCEPTED SUPPLY/);assert.deepEqual(counts(),paused);assert.equal(JSON.stringify(h.memory),journal);
    await n('library-content-materials').children[1].children[2].onclick();await n('library-content-materials').children[0].children[2].onclick();n('library-content-reviewed').checked=true;await submit();await n('library-export').onclick();assert.deepEqual(JSON.parse(await blobs.at(-1).text()).entries[0].fields.materials,[]);await n('library-content-discard').onclick();await n('library-content-undo').onclick();await n('library-export').onclick();assert.deepEqual(JSON.parse(await blobs.at(-1).text()).entries[0],saved.entries[0]);assert.deepEqual(counts(),paused);assert.equal(JSON.stringify(h.memory),journal);
    assert.ok(initial.refreshes<paused.refreshes);h.refreshFailures.clear();await n('recover').onclick();assert.equal(h.paragraphs.length,0);assert.equal(h.cueWrites,0);
  }finally{URL.createObjectURL=oldURL;globalThis.setTimeout=oldTimer;}
});
test('combined public/private exports isolate coexisting private PPTX/library/material drafts and retain them through stale, async and paused refusals',async()=>{
  const corpus=exportCorpus();corpus.project.modules.forEach(m=>m.activities.forEach((a,i)=>{a.materials[0].item='Public supply '+i;}));const h=makeHost(corpus.project),ui=dom();
  const globals=['Office','Word','document','location','JSZip','DOMParser'],prior=Object.fromEntries(globals.map(k=>[k,globalThis[k]]));
  Object.assign(globalThis,{Office:h.Office,Word:h.Word,document:ui.document,location:{search:''},JSZip});installParser();
  const blobs=[],oldURL=URL.createObjectURL,oldTimer=setTimeout;URL.createObjectURL=b=>{blobs.push(b);return 'blob:combined-contract';};globalThis.setTimeout=()=>0;
  try{
    await import('../src/app.mjs?combined-export-contract');await settle();const n=id=>ui.nodes.get(id);
    const original=await pptxFixture(async z=>{const f=z.file('ppt/slides/slide2.xml');z.file(f.name,(await f.async('string')).replace('Text &amp; task 2','PRIVATE_SLIDE_BASELINE'));});
    const incoming=await pptxFixture(async z=>{const f=z.file('ppt/slides/slide2.xml');z.file(f.name,(await f.async('string')).replace('Title 2','PRIVATE_SLIDE_COMPARISON'));});
    n('ppt-file').files=[{name:'Private teaching.pptx',size:original.length,arrayBuffer:async()=>original.buffer}];n('ppt-file').onchange();await n('ppt-read').onclick();assert.match(n('ppt-baseline').textContent,/PRIVATE_SLIDE_BASELINE/);
    n('ppt-text').value='PRIVATE_SLIDE_REVIEW';n('ppt-review-form').onsubmit({preventDefault(){}});await settle();
    n('ppt-notes').value='PRIVATE_SLIDE_UNSAVED';n('ppt-notes').oninput();
    n('ppt-update-file').files=[{name:'Private revised.pptx',size:incoming.length,arrayBuffer:async()=>incoming.buffer}];n('ppt-update-file').onchange();n('ppt-update-lineage').checked=true;n('ppt-update-lineage').onchange();await n('ppt-compare').onclick();assert.match(n('ppt-change-preview').textContent,/PRIVATE_SLIDE_COMPARISON/);
    n('activity-title').value='Private reusable draft';n('activity-participant').value='PRIVATE_LIBRARY_DRAFT_LEARNER';n('activity-facilitator').value='PRIVATE_LIBRARY_DRAFT_TEACHER';n('library-name').value='Private library draft';n('library-save-form').onsubmit({preventDefault(){}});await settle();assert.match(n('library-summary').textContent,/1 session copies/);
    await n('materials-load').onclick();n('material-notes').value='PRIVATE_MATERIAL_UNSAVED';n('course-detail-purpose').value='PRIVATE_COURSE_UNSAVED';
    const ids=['ppt-title','ppt-text','ppt-notes','activity-participant','activity-facilitator','library-name','material-notes','course-detail-purpose'];
    const drafts=()=>Object.fromEntries(ids.map(id=>[id,n(id).value]));const before=drafts(),baseline=n('ppt-baseline').textContent,comparison=n('ppt-change-preview').textContent;
    const secrets=[...corpus.privateCanaries,'PRIVATE SPEAKER NOTE',...Object.values(before).filter(s=>s.startsWith('PRIVATE_')),'Private library draft'];
    const privateForbidden=[...secrets.filter(x=>/PRIVATE_(?:SLIDE|LIBRARY|UNKNOWN|MODULE)/.test(x)),'PRIVATE SPEAKER NOTE','PRIVATE_COURSE_UNSAVED','PRIVATE_MATERIAL_UNSAVED','Private library draft'];
    const preview=async kind=>{if(kind==='facilitator'){await n('facilitator-preview').onclick();n('facilitator-reviewed').checked=true;n('facilitator-reviewed').onchange();}else if(kind==='materials'){await n('materials-checklist-preview').onclick();n('materials-checklist-reviewed').checked=true;n('materials-checklist-reviewed').onchange();}else if(kind==='workbook')await n('preview').onclick();else{n('agenda-form').onsubmit({preventDefault(){}});await settle();}};
    const download=kind=>n(kind==='facilitator'?'facilitator-download':kind==='materials'?'materials-checklist-download':kind==='workbook'?'download':'agenda-download').onclick();
    for(const kind of ['workbook','agenda','materials','facilitator']){
      await preview(kind);await download(kind);assert.equal(blobs.length,['workbook','agenda','materials','facilitator'].indexOf(kind)+1);
      if(kind==='facilitator')await auditFacilitator(blobs.at(-1),{forbidden:privateForbidden});else await auditExport(blobs.at(-1),{project:h.durable['openstage.guide.project.v1'],privateCanaries:secrets,kind});
      assert.deepEqual(drafts(),before);assert.equal(n('ppt-baseline').textContent,baseline);assert.equal(n('ppt-change-preview').textContent,comparison);
    }
    // Reuse the existing generation-boundary mock, now with all private drafts present.
    for(const kind of ['workbook','agenda','materials','facilitator']){
      await preview(kind);const count=blobs.length;let release;
      globalThis.JSZip=class extends JSZip{generateAsync(options){return new Promise(resolve=>{release=()=>super.generateAsync(options).then(resolve);});}};
      const pending=download(kind);await settle();assert.ok(release,'Packaging boundary reached');
      const word=await import('../src/word.mjs');await word.updateCourseDetails(await word.readProject(),{audience:`PRIVATE_ASYNC_${kind}`,purpose:'',authors:'',attribution:''});
      release();await pending;assert.equal(blobs.length,count);assert.match(n('error').textContent,/changed/);assert.deepEqual(drafts(),before);
      globalThis.JSZip=JSZip;await preview(kind);
      h.records.find(r=>r.tag.startsWith(kind==='facilitator'?'osg:facilitator:':kind==='workbook'?'osg:participant:':'osg:activity:')).text=kind==='workbook'?'Revised learner instruction':`Revised ${kind} title`;
      await download(kind);assert.equal(blobs.length,count);assert.match(n('error').textContent,/changed/);assert.deepEqual(drafts(),before);
    }
    const privateState=async()=>{const from=blobs.length;for(const id of ['ppt-export','library-export','materials-draft-download'])await n(id).onclick();assert.equal(blobs.length,from+3);return Promise.all(blobs.slice(from).map(async b=>JSON.parse(await b.text())));};
    const preserved=await privateState();assert.match(JSON.stringify(preserved),/PRIVATE_SLIDE_BASELINE/);assert.match(JSON.stringify(preserved),/PRIVATE_SLIDE_REVIEW/);assert.match(JSON.stringify(preserved),/PRIVATE_SLIDE_UNSAVED/);assert.match(JSON.stringify(preserved),/PRIVATE_LIBRARY_DRAFT_TEACHER/);assert.equal(preserved[2].unsavedEntry.notes,'PRIVATE_MATERIAL_UNSAVED');
    const word=await import('../src/word.mjs');await word.persist(await word.readProject());
    for(const kind of ['workbook','agenda','materials','facilitator'])for(const failure of ['refresh','pending-journal']){
      await preview(kind);let release;globalThis.JSZip=class extends JSZip{generateAsync(options){return new Promise(resolve=>{release=()=>super.generateAsync(options).then(resolve);});}};
      const count=blobs.length,pending=download(kind);await settle();assert.ok(release);const saves=h.saves;
      let external;
      if(failure==='refresh')h.refreshFailures.add(h.refreshes+1);
      else{const base=structuredClone(h.durable['openstage.guide.project.v1']),next=structuredClone(base);next.revision++;next.details.audience='PRIVATE_EXTERNAL_PENDING';external={schema:2,phase:'prepared',txId:crypto.randomUUID(),kind:'metadata',base,project:next};h.durable['openstage.guide.pending.v1']=structuredClone(external);}
      release();await pending;assert.equal(blobs.length,count,'Post-ZIP failure must not download');assert.equal(h.saves,saves,'Export must not replay/persist pending work');assert.deepEqual(drafts(),before);
      if(external){assert.deepEqual(h.durable['openstage.guide.pending.v1'],external);assert.deepEqual(h.memory['openstage.guide.pending.v1'],external);}
      const paused={saves:h.saves,syncs:h.syncs,runs:h.runs,refreshes:h.refreshes};assert.deepEqual(await privateState(),preserved);assert.deepEqual({saves:h.saves,syncs:h.syncs,runs:h.runs,refreshes:h.refreshes},paused);
      globalThis.JSZip=JSZip;h.refreshFailures.clear();await n('recover').onclick();assert.equal(n('refresh').disabled,false);assert.deepEqual(await privateState(),preserved);assert.equal(h.paragraphs.length,0);assert.equal(h.cueWrites,0);
    }
    h.refreshFailures.add(h.refreshes+1);await n('refresh').onclick();const counts={saves:h.saves,syncs:h.syncs,runs:h.runs,refreshes:h.refreshes},journal=JSON.stringify(h.memory),count=blobs.length;
    for(const id of ['download','agenda-download','materials-checklist-download','facilitator-download']){assert.equal(n(id).disabled,true);await n(id).onclick();}
    assert.equal(blobs.length,count);assert.deepEqual({saves:h.saves,syncs:h.syncs,runs:h.runs,refreshes:h.refreshes},counts);assert.equal(JSON.stringify(h.memory),journal);assert.deepEqual(drafts(),before);
    assert.equal(n('ppt-baseline').textContent,baseline);assert.equal(h.paragraphs.length,0);assert.equal(h.cueWrites,0);
    h.refreshFailures.clear();await n('recover').onclick();assert.equal(n('refresh').disabled,false);assert.deepEqual(drafts(),before);
  }finally{URL.createObjectURL=oldURL;globalThis.setTimeout=oldTimer;for(const k of globals)if(prior[k]===undefined)delete globalThis[k];else globalThis[k]=prior[k];}
});
test('same and different activity materials reloads require explicit discard and retain authored rows/raw editor',async()=>{
  const p=fixture(),first=p.modules[0].activities[0],second={...structuredClone(first),id:crypto.randomUUID(),title:'Other activity'};p.modules[0].activities.push(second);p.revision++;
  const h=makeHost(p),ui=dom();globalThis.Office=h.Office;globalThis.Word=h.Word;globalThis.document=ui.document;globalThis.location={search:''};await import('../src/app.mjs?materials-reload-protection');await settle();const n=id=>ui.nodes.get(id);await n('materials-load').onclick();
  n('material-item').value='Authored supply';n('material-unit').value='sets';n('material-quantity').value='2';n('material-notes').value='PRIVATE NEW NOTES';n('materials-add-form').onsubmit({preventDefault(){}});await settle();n('material-quantity').value='Still deciding';const counts={saves:h.saves,syncs:h.syncs,runs:h.runs,refreshes:h.refreshes};
  for(const id of [first.id,second.id]){n('materials-activity').value=id;n('materials-activity').onchange();await n('materials-load').onclick();assert.match(n('error').textContent,/Preserve and discard/);assert.match(n('materials-draft-list').textContent,/Authored supply/);assert.equal(n('material-quantity').value,'Still deciding');assert.equal(n('material-notes').value,'PRIVATE NEW NOTES');assert.deepEqual({saves:h.saves,syncs:h.syncs,runs:h.runs,refreshes:h.refreshes},counts);}
  await n('materials-discard').onclick();await n('materials-load').onclick();assert.equal(n('materials-draft-list').children.length,0);assert.equal(n('materials-activity').value,second.id);assert.equal(n('material-quantity').value,'1');assert.equal(h.paragraphs.length,0);
});

test('materials draft restore preserves active work, binds destination and requires review before replacing saved requirements',async()=>{
  const p=fixture(),id=p.modules[0].activities[0].id,entry={item:'Handout',quantity:2,unit:'copies',use:'consumable',notes:'PRIVATE RESTORED NOTES'},raw={item:'Unfinished',quantity:'not decided',unit:'',use:'consumable',notes:'PRIVATE RAW NOTES\nsecond line'},text=serializeMaterialsDraft(p,id,[entry],raw);
  p.modules[0].activities[0].materials=[{...entry,quantity:5}];p.revision++;const h=makeHost(p),ui=dom();globalThis.Office=h.Office;globalThis.Word=h.Word;globalThis.document=ui.document;globalThis.location={search:''};
  const blobs=[],oldURL=URL.createObjectURL,oldTimer=setTimeout;URL.createObjectURL=b=>{blobs.push(b);return 'blob:restored-materials';};globalThis.setTimeout=()=>0;
  try{
    await import('../src/app.mjs?materials-restore-ui');await settle();const n=id=>ui.nodes.get(id),body=structuredClone(h.records),initialSaves=h.saves;
    n('materials-restore-file').files=[{size:10,text:async()=>'{bad'}];n('materials-restore-file').onchange();await n('materials-restore').onclick();assert.equal(h.saves,initialSaves);assert.equal(n('materials-draft-list').children.length,0);
    let release;n('materials-restore-file').files=[{size:text.length,text:()=>new Promise(r=>{release=r;})}];n('materials-restore-file').onchange();const pending=n('materials-restore').onclick();await settle();n('materials-restore-file').onchange();release(text);await pending;assert.match(n('error').textContent,/file or destination changed/);assert.equal(n('materials-draft-list').children.length,0);
    n('materials-restore-file').files=[{size:text.length,text:async()=>text}];n('materials-restore-file').onchange();await n('materials-restore').onclick();assert.equal(n('materials-restore-review').hidden,false);assert.match(n('materials-current-list').textContent,/5 copies/);assert.match(n('materials-draft-list').textContent,/2 copies/);assert.equal(n('material-quantity').value,'not decided');assert.equal(n('material-notes').value,raw.notes);assert.equal(n('materials-save').disabled,true);assert.equal(h.saves,initialSaves);
    await n('materials-save').onclick();assert.match(n('error').textContent,/Review the restored list/);assert.equal(h.saves,initialSaves);await n('materials-restore').onclick();assert.match(n('error').textContent,/Preserve and discard/);assert.equal(n('material-quantity').value,'not decided');
    n('materials-review-approved').checked=true;n('materials-review-approved').onchange();assert.equal(n('materials-save').disabled,false);await n('materials-save').onclick();assert.equal(h.durable['openstage.guide.project.v1'].modules[0].activities[0].materials[0].quantity,2);assert.deepEqual(h.records,body);assert.equal(h.paragraphs.length,0);
    await n('materials-discard').onclick();h.records.find(r=>r.tag===`osg:participant:${id}`).text='New Word text';await n('materials-restore').onclick();assert.match(n('error').textContent,/course changed/);assert.equal(n('materials-draft-list').children.length,0);await n('materials-refresh').onclick();await n('materials-restore').onclick();assert.equal(n('materials-save').disabled,true);
    n('materials-review-approved').checked=true;n('materials-review-approved').onchange();await n('materials-draft-list').children[0].children[1].onclick();assert.equal(n('materials-review-approved').checked,false);assert.equal(n('materials-save').disabled,true);n('material-quantity').value='1';n('material-unit').value='sets';n('materials-add-form').onsubmit({preventDefault(){}});await settle();assert.equal(n('materials-save').disabled,true);
    h.refreshFailures.add(h.refreshes+1);await n('refresh').onclick();const paused={saves:h.saves,syncs:h.syncs,runs:h.runs,refreshes:h.refreshes};await n('materials-draft-download').onclick();const downloaded=JSON.parse(await blobs[0].text());assert.equal(downloaded.items[0].item,'Unfinished');assert.equal(downloaded.unsavedEntry.notes,raw.notes);await n('materials-restore').onclick();assert.deepEqual({saves:h.saves,syncs:h.syncs,runs:h.runs,refreshes:h.refreshes},paused);h.refreshFailures.clear();await n('recover').onclick();assert.equal(n('refresh').disabled,false);
  }finally{URL.createObjectURL=oldURL;globalThis.setTimeout=oldTimer;}
});

test('activity materials preserve stale/uncertain drafts, export privately while paused and follow independent library insertions',async()=>{
  const p=fixture(),h=makeHost(p),ui=dom();globalThis.Office=h.Office;globalThis.Word=h.Word;globalThis.document=ui.document;globalThis.location={search:''};globalThis.fetch=async()=>({ok:true,blob:async()=>new Blob(['mock icon'])});globalThis.FileReader=class{readAsDataURL(){this.result='data:image/png;base64,AAAA';queueMicrotask(()=>this.onload());}};
  const blobs=[],oldURL=URL.createObjectURL,oldTimer=setTimeout;URL.createObjectURL=b=>{blobs.push(b);return 'blob:materials';};globalThis.setTimeout=()=>0;
  try{
    await import('../src/app.mjs?materials-ui');await settle();const n=id=>ui.nodes.get(id),id=p.modules[0].activities[0].id;assert.equal(n('materials-save').disabled,true);await n('materials-load').onclick();
    const add=async()=>{n('materials-add-form').onsubmit({preventDefault(){}});await settle();};n('material-item').value='Handout';n('material-unit').value='copies';n('material-notes').value='PRIVATE PREPARATION';n('material-quantity').value='0';await add();assert.equal(n('materials-draft-list').children.length,0);n('material-quantity').value='2';await add();assert.equal(n('materials-draft-list').children.length,1);const oldRemove=n('materials-draft-list').children[0].children[1];
    n('material-item').value='Board';n('material-unit').value='pieces';n('material-use').value='reusable';n('material-quantity').value='1';await add();await oldRemove.onclick();assert.equal(n('materials-draft-list').children.length,2);assert.match(n('error').textContent,/draft item changed/);
    h.records.find(r=>r.tag===`osg:participant:${id}`).text='New direct learner text';await n('materials-save').onclick();assert.match(n('error').textContent,/changed/);assert.equal(n('materials-draft-list').children.length,2);assert.equal(n('material-item').value,'Board');
    await n('materials-discard').onclick();await n('materials-load').onclick();assert.equal(n('materials-draft-list').children.length,0);n('material-item').value='Board';n('material-unit').value='pieces';n('material-use').value='reusable';n('material-quantity').value='1';n('material-notes').value='PRIVATE PREPARATION';await add();const body=structuredClone(h.records);h.saveFailures.set(h.saves+2,{persist:true});await n('materials-save').onclick();assert.equal(n('materials-save').disabled,true);assert.equal(n('materials-draft-list').children.length,1);assert.equal(n('material-notes').value,'PRIVATE PREPARATION');
    const paused={saves:h.saves,syncs:h.syncs,runs:h.runs,refreshes:h.refreshes},journal=JSON.stringify(h.memory);await n('materials-draft-download').onclick();const record=JSON.parse(await blobs[0].text());assert.equal(record.items[0].item,'Board');assert.equal(record.unsavedEntry.notes,'PRIVATE PREPARATION');assert.ok(!JSON.stringify(record).includes('PRIVATE DIRECTIONS'));assert.ok(!JSON.stringify(record).includes('SECRET ANSWER'));assert.equal(JSON.stringify(h.memory),journal);await n('materials-save').onclick();assert.deepEqual({saves:h.saves,syncs:h.syncs,runs:h.runs,refreshes:h.refreshes},paused);
    await n('recover').onclick();assert.deepEqual(h.records,body);assert.equal(h.paragraphs.length,0);assert.equal(h.cueWrites,0);assert.match(n('materials-plan-list').textContent,/Board/);await n('materials-discard').onclick();await n('materials-load').onclick();const saves=h.saves;await n('materials-save').onclick();assert.equal(h.saves,saves);
    n('materials-library-name').value='Prepared board activity';n('materials-library-tags').value='supplies';n('materials-library-form').onsubmit({preventDefault(){}});await settle();await n('library-use').onclick();assert.match(n('activity-library-note').textContent,/1 material entries/);assert.equal(n('activity-participant').value,'New direct learner text');
    const note=()=>n('activity-copied-materials-list').children[0].children.find(e=>e.id.endsWith('-notes'));note().value='TODO supplies';note().oninput();const beforeInsert=h.saves;n('activity-form').onsubmit({preventDefault(){}});await settle();assert.match(n('error').textContent,/material 1 notes/);assert.equal(h.saves,beforeInsert);note().value='PRIVATE PREPARATION';note().oninput();
    n('activity-form').onsubmit({preventDefault(){}});await settle();const stored=h.durable['openstage.guide.project.v1'],copied=stored.modules[0].activities.at(-1);assert.notEqual(copied.id,id);assert.deepEqual(copied.materials,stored.modules[0].activities[0].materials);assert.equal(stored.modules[0].activities.length,2);assert.equal(copied.materials[0].notes,'PRIVATE PREPARATION');
  }finally{URL.createObjectURL=oldURL;globalThis.setTimeout=oldTimer;}
});

test('build diagnostic remains honestly unknown without pinned execution and exports no course data even during quarantine',async()=>{
  const p=fixture(),h=makeHost(p),ui=dom();p.details={audience:'PRIVATE AUDIENCE',purpose:'PRIVATE PURPOSE',authors:'PRIVATE USER',attribution:'PRIVATE RIGHTS'};
  globalThis.Office=h.Office;globalThis.Word=h.Word;globalThis.document=ui.document;globalThis.location={search:''};
  const blobs=[],oldURL=URL.createObjectURL,oldTimer=setTimeout,oldFetch=globalThis.fetch;URL.createObjectURL=b=>{blobs.push(b);return 'blob:diagnostic';};globalThis.setTimeout=()=>0;
  try{
    await import('../src/app.mjs?diagnostic-ui');await settle();const n=id=>ui.nodes.get(id);let requests=0;globalThis.fetch=async()=>{requests++;throw new Error('Unexpected metadata request');};
    const counts={saves:h.saves,syncs:h.syncs,runs:h.runs,refreshes:h.refreshes},originalDocument=Object.getOwnPropertyDescriptor(h.Office.context,'document');
    Object.defineProperty(h.Office.context,'document',{configurable:true,get(){throw new Error('Diagnostic must not inspect document');}});
    try{const pending=n('build-check').onclick();assert.equal(n('build-check').disabled,true);assert.equal(n('build-download').disabled,true);await pending;assert.equal(n('build-check').disabled,false);await n('build-download').onclick();}finally{Object.defineProperty(h.Office.context,'document',originalDocument);}
    assert.equal(requests,0);assert.match(n('build-diagnostic-summary').textContent,/unknown/);assert.match(n('build-loaded').textContent,/unavailable/);const report=JSON.parse(await blobs[0].text());assert.equal(report.loadedBuildId,null);assert.equal(report.nativeManifestVerified,false);assert.doesNotMatch(JSON.stringify(report),/PRIVATE|SECRET|Recovery sample|Jason|openstage.guide.project|facilitator|answers|C:\\/);assert.deepEqual({saves:h.saves,syncs:h.syncs,runs:h.runs,refreshes:h.refreshes},counts);
    h.refreshFailures.add(h.refreshes+1);await n('refresh').onclick();const paused={saves:h.saves,syncs:h.syncs,runs:h.runs,refreshes:h.refreshes},journal=JSON.stringify(h.memory);assert.equal(n('build-check').disabled,false);const recheck=n('build-check').onclick();assert.equal(n('build-check').disabled,true);assert.equal(n('build-download').disabled,true);assert.match(n('build-loaded').textContent,/Build not verified/);await recheck;assert.equal(n('build-check').disabled,false);await n('build-download').onclick();assert.equal(blobs.length,2);assert.equal(JSON.stringify(h.memory),journal);assert.deepEqual({saves:h.saves,syncs:h.syncs,runs:h.runs,refreshes:h.refreshes},paused);h.refreshFailures.clear();await n('recover').onclick();assert.equal(n('refresh').disabled,false);
  }finally{URL.createObjectURL=oldURL;globalThis.setTimeout=oldTimer;globalThis.fetch=oldFetch;}
});

test('library management selects newly restored entries, preserves drafts on conflicts and never calls Word',async()=>{
  const p=fixture(),h=makeHost(p),ui=dom();globalThis.Office=h.Office;globalThis.Word=h.Word;globalThis.document=ui.document;globalThis.location={search:''};
  // A select cannot select a value until its option exists, as in real HTML.
  const selector=ui.nodes.get('library-item');let selected='';Object.defineProperty(selector,'value',{get(){return this.children.some(o=>o.value===selected)?selected:'';},set(v){selected=this.children.some(o=>o.value===v)?v:'';}});
  const blobs=[],oldURL=URL.createObjectURL,oldTimer=setTimeout;URL.createObjectURL=b=>{blobs.push(b);return 'blob:managed-library';};globalThis.setTimeout=()=>0;
  try{
    await import('../src/app.mjs?library-management-ui');await settle();const n=id=>ui.nodes.get(id),counts={saves:h.saves,syncs:h.syncs,runs:h.runs,refreshes:h.refreshes};n('activity-title').value='Original activity draft';n('activity-participant').value='Learner task';
    for(const name of ['One','Two']){n('library-name').value=name;n('library-save-form').onsubmit({preventDefault(){}});await settle();assert.equal(n('library-item').children.find(o=>o.value===n('library-item').value).textContent,name);}
    const second=n('library-item').value;await n('library-manage-load').onclick();assert.equal(n('library-edit-name').value,'Two');n('library-edit-name').value='One';n('library-manage-form').onsubmit({preventDefault(){}});await settle();assert.equal(n('library-edit-name').value,'One');assert.match(n('library-summary').textContent,/2 session/);
    n('library-edit-name').value='Renamed';n('library-edit-tags').value='practice';n('library-manage-form').onsubmit({preventDefault(){}});await settle();assert.equal(n('library-item').value,second);assert.match(n('library-summary').textContent,/2 session/);await n('library-export').onclick();const catalog=JSON.parse(await blobs[0].text());assert.equal(catalog.entries[1].id,second);assert.equal(catalog.entries[1].name,'Renamed');assert.equal(catalog.entries[1].fields.title,'Original activity draft');
    await n('library-remove').onclick();assert.equal(n('library-manage-form').querySelector().disabled,true);assert.equal(n('library-edit-name').value,'Renamed');assert.match(n('library-manage-summary').textContent,/Undo available/);
    n('library-name').value='Renamed';n('library-save-form').onsubmit({preventDefault(){}});await settle();await n('library-undo').onclick();assert.match(n('error').textContent,/Unsupported/);assert.equal(n('library-undo').disabled,false);assert.match(n('library-summary').textContent,/2 session/);
    await n('library-manage-load').onclick();n('library-edit-name').value='Other new copy';n('library-manage-form').onsubmit({preventDefault(){}});await settle();await n('library-undo').onclick();assert.equal(n('library-item').value,second);assert.equal(n('library-undo').disabled,true);assert.match(n('library-summary').textContent,/3 session/);assert.equal(n('activity-title').value,'Original activity draft');
    for(let i=0;i<3;i++)await n('library-remove').onclick();await n('library-export').onclick();assert.deepEqual(JSON.parse(await blobs[1].text()).entries,[]);assert.equal(n('library-undo').disabled,false);assert.deepEqual({saves:h.saves,syncs:h.syncs,runs:h.runs,refreshes:h.refreshes},counts);
    h.refreshFailures.add(h.refreshes+1);await n('refresh').onclick();assert.equal(n('library-undo').disabled,false);const paused={saves:h.saves,syncs:h.syncs,runs:h.runs,refreshes:h.refreshes},journal=JSON.stringify(h.memory);await n('library-undo').onclick();assert.match(n('library-summary').textContent,/1 session/);assert.equal(n('library-remove').disabled,true);await n('library-remove').onclick();await n('library-export').onclick();assert.equal(JSON.parse(await blobs[2].text()).entries.length,1);assert.equal(JSON.stringify(h.memory),journal);assert.deepEqual({saves:h.saves,syncs:h.syncs,runs:h.runs,refreshes:h.refreshes},paused);h.refreshFailures.clear();await n('recover').onclick();assert.equal(h.paragraphs.length,0);assert.equal(h.cueWrites,0);
  }finally{URL.createObjectURL=oldURL;globalThis.setTimeout=oldTimer;}
});

test('revised PPTX comparison is read-only, protects authored reviews and invalidates results after edits or file races',async()=>{
  const p=fixture(),h=makeHost(p),ui=dom();globalThis.Office=h.Office;globalThis.Word=h.Word;globalThis.document=ui.document;globalThis.location={search:''};globalThis.JSZip=JSZip;installParser();
  const blobs=[],oldURL=URL.createObjectURL,oldTimer=setTimeout;URL.createObjectURL=b=>{blobs.push(b);return 'blob:comparison-review';};globalThis.setTimeout=()=>0;
  try{
    await import('../src/app.mjs?ppt-comparison-ui');await settle();const n=id=>ui.nodes.get(id),original=await pptxFixture(),incoming=await pptxFixture(async z=>{const path='ppt/slides/slide2.xml',text=await z.file(path).async('string');z.file(path,text.replace('Title 2','Incoming title'));}),before=original.slice(),counts={saves:h.saves,syncs:h.syncs,runs:h.runs,refreshes:h.refreshes};
    n('ppt-file').files=[{name:'Original.pptx',size:original.length,arrayBuffer:async()=>original.buffer}];n('ppt-file').onchange();await n('ppt-read').onclick();const baseline=n('ppt-baseline').textContent,summary=n('ppt-summary').textContent;
    n('ppt-title').value='Local reviewed title';n('ppt-review-form').onsubmit({preventDefault(){}});await settle();n('ppt-notes').value='PRIVATE UNSAVED COMPARISON NOTE';n('ppt-notes').oninput();
    let reads=0;n('ppt-update-file').files=[{name:'Revised.pptx',size:incoming.length,arrayBuffer:async()=>{reads++;return incoming.buffer;}}];n('ppt-update-file').onchange();await n('ppt-compare').onclick();assert.equal(reads,0);assert.match(n('error').textContent,/Confirm that you intend/);
    n('ppt-update-lineage').checked=true;n('ppt-update-lineage').onchange();await n('ppt-compare').onclick();assert.match(n('ppt-compare-summary').textContent,/1 changed/);assert.match(n('ppt-compare-summary').textContent,/1 with local conflicts/);assert.match(n('ppt-change-preview').textContent,/Local reviewed title/);assert.match(n('ppt-change-preview').textContent,/Incoming title/);assert.match(n('ppt-change-preview').textContent,/PRIVATE UNSAVED COMPARISON NOTE/);assert.equal(n('ppt-baseline').textContent,baseline);assert.equal(n('ppt-summary').textContent,summary);assert.equal(n('ppt-title').value,'Local reviewed title');
    n('ppt-notes').value='Later unsaved notes';n('ppt-notes').oninput();assert.equal(n('ppt-change-preview').textContent,'');await n('ppt-compare').onclick();assert.match(n('ppt-change-preview').textContent,/Later unsaved notes/);
    for(const kind of ['editor','file','intent']){let release;n('ppt-update-file').files=[{name:'Revised.pptx',size:incoming.length,arrayBuffer:()=>new Promise(r=>{release=r;})}];n('ppt-update-file').onchange();n('ppt-update-lineage').checked=true;const pending=n('ppt-compare').onclick();await settle();assert.ok(release);if(kind==='editor'){n('ppt-text').value='Edited while reading';n('ppt-text').oninput();}else if(kind==='file')n('ppt-update-file').onchange();else{n('ppt-update-lineage').checked=false;n('ppt-update-lineage').onchange();}release(incoming.buffer);await pending;assert.match(n('error').textContent,/file or author review changed/);assert.equal(n('ppt-change-preview').textContent,'');assert.equal(n('ppt-title').value,'Local reviewed title');}
    n('ppt-update-lineage').checked=true;n('ppt-update-file').files=[{name:'Bad.pptx',size:100,arrayBuffer:async()=>new Uint8Array(100).buffer}];n('ppt-update-file').onchange();await n('ppt-compare').onclick();assert.equal(n('ppt-baseline').textContent,baseline);assert.equal(n('ppt-text').value,'Edited while reading');
    await n('ppt-export').onclick();const exported=JSON.parse(await blobs[0].text());assert.equal(exported.source.slides[0].title,'Title 2');assert.equal(exported.review.slides[0].title,'Local reviewed title');assert.equal(exported.unsavedEditor.text,'Edited while reading');assert.ok(!Object.hasOwn(exported,'comparison'));assert.deepEqual(original,before);assert.deepEqual({saves:h.saves,syncs:h.syncs,runs:h.runs,refreshes:h.refreshes},counts);
    h.refreshFailures.add(h.refreshes+1);await n('refresh').onclick();assert.equal(n('ppt-compare').disabled,true);const paused={saves:h.saves,syncs:h.syncs,runs:h.runs,refreshes:h.refreshes};await n('ppt-compare').onclick();assert.deepEqual({saves:h.saves,syncs:h.syncs,runs:h.runs,refreshes:h.refreshes},paused);h.refreshFailures.clear();await n('recover').onclick();assert.equal(h.paragraphs.length,0);assert.equal(h.cueWrites,0);
  }finally{URL.createObjectURL=oldURL;globalThis.setTimeout=oldTimer;}
});

test('private review restores a matching new session without Word calls and retains work on failed or changed-file restores',async()=>{
  const p=fixture(),h=makeHost(p),ui=dom();globalThis.Office=h.Office;globalThis.Word=h.Word;globalThis.document=ui.document;globalThis.location={search:''};globalThis.JSZip=JSZip;installParser();
  await import('../src/app.mjs?ppt-restore-ui');await settle();const n=id=>ui.nodes.get(id),bytes=await pptxFixture(),source=await readPptx(bytes,'Source.pptx',JSZip),review=pptxDraft(source);
  review.slides[1].title='Saved reviewed title';review.slides[1].moduleId='old-course-module';review.slides[1].lesson='Saved proposed lesson';
  const editor={...review.slides[1],title:'Unsaved title',text:'Unsaved private text',notes:'PRIVATE UNSAVED NOTES',moduleId:'other-old-module'},text=serializeReview(source,review,editor),counts={saves:h.saves,syncs:h.syncs,runs:h.runs,refreshes:h.refreshes};
  n('ppt-file').files=[{name:'Renamed.pptx',size:bytes.length,arrayBuffer:async()=>bytes.buffer}];n('ppt-file').onchange();
  n('ppt-review-file').files=[{size:10,text:async()=>'{bad'}];n('ppt-review-file').onchange();await n('ppt-restore').onclick();assert.equal(n('ppt-title').value,'');
  for(const changed of ['ppt-file','ppt-review-file']){let release;n('ppt-review-file').files=[{size:text.length,text:()=>new Promise(r=>{release=r;})}];n('ppt-review-file').onchange();const pending=n('ppt-restore').onclick();await settle();n(changed).onchange();release(text);await pending;assert.match(n('error').textContent,/selected restore file changed/);assert.equal(n('ppt-title').value,'');}
  n('ppt-review-file').files=[{size:text.length,text:async()=>text}];n('ppt-review-file').onchange();await n('ppt-restore').onclick();assert.equal(n('ppt-title').value,'Unsaved title');assert.equal(n('ppt-text').value,'Unsaved private text');assert.equal(n('ppt-slide').value,'slide:256');assert.equal(n('ppt-module').value,'other-old-module');assert.match(n('ppt-baseline').textContent,/Title 1/);assert.ok(!n('ppt-baseline').textContent.includes('Saved reviewed title'));assert.match(n('ppt-summary').textContent,/Renamed.pptx/);
  n('ppt-review-form').onsubmit({preventDefault(){}});await settle();assert.match(n('error').textContent,/module no longer exists/);assert.equal(n('ppt-title').value,'Unsaved title');n('ppt-module').value=p.modules[0].id;n('ppt-review-form').onsubmit({preventDefault(){}});await settle();assert.match(n('status').textContent,/saved in this session/);
  let readCount=0;n('ppt-review-file').files=[{size:10,text:async()=>{readCount++;return '{bad';}}];await n('ppt-restore').onclick();assert.equal(readCount,0);assert.match(n('error').textContent,/Discard the current/);assert.equal(n('ppt-title').value,'Unsaved title');
  assert.deepEqual({saves:h.saves,syncs:h.syncs,runs:h.runs,refreshes:h.refreshes},counts);assert.equal(h.paragraphs.length,0);assert.equal(h.cueWrites,0);
});

test('PowerPoint review keeps edits separate, retains drafts after failed reads and never calls Word',async()=>{
  const p=fixture(),h=makeHost(p),ui=dom();globalThis.Office=h.Office;globalThis.Word=h.Word;globalThis.document=ui.document;globalThis.location={search:''};globalThis.JSZip=JSZip;installParser();
  const blobs=[],oldURL=URL.createObjectURL,oldTimer=setTimeout;URL.createObjectURL=b=>{blobs.push(b);return 'blob:ppt-review';};globalThis.setTimeout=()=>0;
  try{
    await import('../src/app.mjs?ppt-review-ui');await settle();const n=id=>ui.nodes.get(id),bytes=await pptxFixture(),before=bytes.slice(),counts={saves:h.saves,syncs:h.syncs,runs:h.runs,refreshes:h.refreshes};
    let release;n('ppt-file').files=[{name:'Other.pptx',size:bytes.length,arrayBuffer:()=>new Promise(r=>{release=r;})}];n('ppt-file').onchange();const pending=n('ppt-read').onclick();await settle();n('ppt-file').onchange();release(bytes.buffer);await pending;assert.match(n('error').textContent,/selected PowerPoint file changed/);assert.equal(n('ppt-title').value,'');
    n('ppt-file').files=[{name:'Source.pptx',size:bytes.length,arrayBuffer:async()=>bytes.buffer}];n('ppt-file').onchange();await n('ppt-read').onclick();assert.match(n('ppt-summary').textContent,/2 slides/);assert.equal(n('ppt-title').value,'Title 2');assert.match(n('ppt-baseline').textContent,/Title 2/);assert.match(n('ppt-warnings').textContent,/External/);
    n('ppt-title').value='Edited review title';n('ppt-notes').value='PRIVATE REVIEW NOTES';n('ppt-module').value=p.modules[0].id;n('ppt-lesson').value='Lesson One';
    n('ppt-slide').value='slide:256';await n('ppt-slide').onchange();assert.match(n('error').textContent,/Save this slide review/);assert.equal(n('ppt-title').value,'Edited review title');assert.equal(n('ppt-slide').value,'slide:258');
    n('ppt-review-form').onsubmit({preventDefault(){}});await settle();assert.match(n('ppt-baseline').textContent,/Title 2/);assert.ok(!n('ppt-baseline').textContent.includes('Edited review title'));
    n('ppt-slide').value='slide:256';await n('ppt-slide').onchange();assert.equal(n('ppt-title').value,'Title 1');assert.equal(n('ppt-notes').value,'PRIVATE SPEAKER NOTE');
    n('ppt-slide').value='slide:258';await n('ppt-slide').onchange();assert.equal(n('ppt-title').value,'Edited review title');assert.equal(n('ppt-module').value,p.modules[0].id);
    n('ppt-text').value='Retain this unsaved review';n('ppt-file').files=[{name:'Bad.pptx',size:100,arrayBuffer:async()=>new Uint8Array(100).buffer}];n('ppt-file').onchange();await n('ppt-read').onclick();assert.equal(n('ppt-text').value,'Retain this unsaved review');assert.match(n('ppt-summary').textContent,/Source.pptx/);
    const otherBytes=await pptxFixture(async z=>{const path='ppt/slides/slide2.xml',text=await z.file(path).async('string');z.file(path,text.replace('Title 2','Different deck title'));});
    for(const [name,data] of [['Source.pptx',bytes],['Different.pptx',otherBytes]]){n('ppt-file').files=[{name,size:data.length,arrayBuffer:async()=>data.buffer}];n('ppt-file').onchange();await n('ppt-read').onclick();assert.match(n('error').textContent,/Discard the current slide review/);assert.equal(n('ppt-text').value,'Retain this unsaved review');assert.equal(n('ppt-title').value,'Edited review title');assert.equal(n('ppt-module').value,p.modules[0].id);}
    await n('ppt-export').onclick();const saved=JSON.parse(await blobs[0].text());assert.equal(saved.source.slides[0].title,'Title 2');assert.equal(saved.review.slides[0].title,'Edited review title');assert.equal(saved.unsavedEditor.text,'Retain this unsaved review');assert.equal(saved.review.slides[0].notes,'PRIVATE REVIEW NOTES');
    assert.deepEqual({saves:h.saves,syncs:h.syncs,runs:h.runs,refreshes:h.refreshes},counts);assert.deepEqual(bytes,before);assert.equal(h.paragraphs.length,0);assert.equal(h.cueWrites,0);
    h.refreshFailures.add(h.refreshes+1);await n('refresh').onclick();assert.equal(n('ppt-read').disabled,true);assert.equal(n('ppt-export').disabled,false);const paused={saves:h.saves,syncs:h.syncs,runs:h.runs,refreshes:h.refreshes};await n('ppt-export').onclick();await n('ppt-read').onclick();assert.deepEqual({saves:h.saves,syncs:h.syncs,runs:h.runs,refreshes:h.refreshes},paused);assert.equal(blobs.length,2);
    h.refreshFailures.clear();await n('recover').onclick();assert.equal(n('refresh').disabled,false);
    await n('ppt-reset').onclick();assert.equal(n('ppt-text').value,'');await n('ppt-read').onclick();assert.equal(n('ppt-title').value,'Different deck title');
  }finally{URL.createObjectURL=oldURL;globalThis.setTimeout=oldTimer;}
});

test('session library preserves drafts on failure, checks imports and inserts independent reviewed copies into the selected module',async()=>{
  const p=addModule(fixture(),'Later module','Objective').next,h=makeHost(p),ui=dom();
  globalThis.Office=h.Office;globalThis.Word=h.Word;globalThis.document=ui.document;globalThis.location={search:''};
  globalThis.fetch=async()=>({ok:true,blob:async()=>new Blob(['mock icon'])});globalThis.FileReader=class{readAsDataURL(){this.result='data:image/png;base64,AAAA';queueMicrotask(()=>this.onload());}};
  const blobs=[],oldURL=URL.createObjectURL,oldTimer=setTimeout;URL.createObjectURL=b=>{blobs.push(b);return 'blob:library';};globalThis.setTimeout=()=>0;
  try{
    await import('../src/app.mjs?library-ui');await settle();const n=id=>ui.nodes.get(id);
    n('activity-title').value='Reusable scenario';n('activity-participant').value='[insert scenario]';n('activity-facilitator').value='PRIVATE LIBRARY NOTE';n('activity-answers').value='PRIVATE LIBRARY ANSWER';
    n('library-name').value='Role play';n('library-tags').value='Teamwork, Practice';n('library-save-form').onsubmit({preventDefault(){}});await settle();assert.match(n('library-summary').textContent,/1 session copies/);assert.match(n('library-summary').textContent,/participant/);assert.equal(h.paragraphs.length,0);
    n('library-save-form').onsubmit({preventDefault(){}});await settle();assert.match(n('error').textContent,/Unsupported/);assert.equal(n('library-name').value,'Role play');assert.equal(n('activity-participant').value,'[insert scenario]');assert.match(n('library-summary').textContent,/1 session copies/);
    n('library-filter').value='missing';n('library-filter').oninput();assert.equal(n('library-use').disabled,true);n('library-filter').value='teamwork';n('library-filter').oninput();assert.equal(n('library-use').disabled,false);
    n('activity-module').value=p.modules[0].id;await n('library-use').onclick();assert.equal(n('activity-module').value,p.modules[0].id);assert.match(n('activity-library-note').textContent,/Independent copy/);assert.match(n('library-preview').textContent,/PRIVATE LIBRARY ANSWER/);
    const saves=h.saves;n('activity-form').onsubmit({preventDefault(){}});await settle();assert.match(n('error').textContent,/participant/);assert.equal(h.saves,saves);
    n('activity-participant').value='Reviewed learner scenario';n('activity-lines').value='1.2';n('activity-form').onsubmit({preventDefault(){}});await settle();assert.equal(h.saves,saves);assert.match(n('error').textContent,/Unsupported/);n('activity-lines').value='5';
    n('activity-form').onsubmit({preventDefault(){}});await settle();assert.equal(n('activity-module').value,p.modules[0].id);
    const saved=h.durable['openstage.guide.project.v1'],copy=saved.modules[0].activities.at(-1);assert.equal(saved.modules[1].activities.length,0);assert.notEqual(copy.id,saved.modules[0].activities[0].id);assert.equal(copy.librarySource.entryName,'Role play');assert.notEqual(copy.id,copy.librarySource.entryId);assert.equal(copy.participant,'Reviewed learner scenario');assert.equal(n('activity-library-note').textContent,'');
    await n('library-export').onclick();const exported=JSON.parse(await blobs[0].text());assert.equal(exported.entries[0].fields.participant,'[insert scenario]');assert.ok(!JSON.stringify(exported).includes(p.id));assert.match(JSON.stringify(exported),/PRIVATE LIBRARY/);
    n('library-file').files=[{size:10,text:async()=>'{bad'}];n('library-file').onchange();await n('library-import').onclick();assert.match(n('library-summary').textContent,/1 session copies/);assert.equal(n('activity-participant').value,'');
    const incoming=saveLibraryCopy(emptyLibrary(),'Second copy',[],exported.entries[0].fields);let resolveFile;n('library-file').files=[{size:500,text:()=>new Promise(r=>{resolveFile=r;})}];n('library-file').onchange();const pending=n('library-import').onclick();await settle();n('library-file').onchange();resolveFile(JSON.stringify(incoming));await pending;assert.match(n('error').textContent,/selected file changed/);assert.match(n('library-summary').textContent,/1 session copies/);
    n('library-file').files=[{size:500,text:async()=>JSON.stringify(incoming)}];n('library-file').onchange();await n('library-import').onclick();assert.match(n('library-summary').textContent,/2 session copies/);
    await n('library-import').onclick();assert.match(n('library-summary').textContent,/2 session copies/);
    await n('library-use').onclick();n('activity-participant').value='Retained draft after uncertain save';h.saveFailures.set(h.saves+2,{persist:true});n('activity-form').onsubmit({preventDefault(){}});await settle();assert.equal(n('activity-participant').value,'Retained draft after uncertain save');assert.match(n('activity-library-note').textContent,/Independent copy/);assert.equal(n('library-use').disabled,true);
    const withheld=h.records.pop();await n('recover').onclick();assert.equal(n('library-use').disabled,true);
    const counts={saves:h.saves,syncs:h.syncs,runs:h.runs,refreshes:h.refreshes},journal=JSON.stringify(h.memory);await n('library-use').onclick();await n('library-import').onclick();assert.equal(n('library-export').disabled,false);await n('library-export').onclick();assert.deepEqual({saves:h.saves,syncs:h.syncs,runs:h.runs,refreshes:h.refreshes},counts);assert.equal(JSON.stringify(h.memory),journal);assert.equal(blobs.length,2);assert.equal(JSON.parse(await blobs[1].text()).entries.length,2);assert.equal(n('library-use').disabled,true);
    h.records.push(withheld);await n('recover').onclick();assert.equal(n('library-use').disabled,false);assert.match(n('library-summary').textContent,/2 session copies/);assert.equal(n('activity-participant').value,'Retained draft after uncertain save');
  }finally{URL.createObjectURL=oldURL;globalThis.setTimeout=oldTimer;}
});

test('workbook export requires reviewed current content and refuses edits or valid metadata commits during packaging',async()=>{
  const p=fixture(),h=makeHost(p),ui=dom();
  globalThis.Office=h.Office;globalThis.Word=h.Word;globalThis.document=ui.document;globalThis.location={search:''};globalThis.JSZip=JSZip;
  const blobs=[],links=[],create=ui.document.createElement,oldURL=URL.createObjectURL,oldTimer=setTimeout;
  ui.document.createElement=tag=>{const e=create(tag);if(tag==='a')links.push(e);return e;};URL.createObjectURL=b=>{blobs.push(b);return 'blob:workbook';};globalThis.setTimeout=()=>0;
  try{
    await import('../src/app.mjs?workbook-output-ui');await settle();const n=id=>ui.nodes.get(id);
    await n('download').onclick();assert.equal(blobs.length,0);assert.match(n('error').textContent,/Refresh and review/);
    await n('preview').onclick();assert.equal(n('download').disabled,false);assert.match(n('workbook-preview').textContent,/Project ID:.*Generation revision/);
    n('agenda-start').value='23:55';n('agenda-start').oninput();assert.equal(n('download').disabled,false);await n('download').onclick();assert.equal(blobs.length,1);assert.match(links[0].download,/ - Participant Workbook - .*\.docx$/);
    h.records.find(r=>r.tag.startsWith('osg:participant:')).text='New learner instructions';await n('download').onclick();assert.equal(blobs.length,1);assert.match(n('error').textContent,/course changed/);assert.equal(n('download').disabled,true);
    await n('preview').onclick();await n('download').onclick();assert.equal(blobs.length,2);
    const doc=await (await JSZip.loadAsync(await blobs[1].arrayBuffer())).file('word/document.xml').async('string');assert.match(doc,/New learner instructions/);assert.doesNotMatch(doc,/PRIVATE|SECRET/);
    let release;globalThis.JSZip=class extends JSZip{generateAsync(options){return new Promise(resolve=>{release=()=>super.generateAsync(options).then(resolve);});}};
    const during=n('download').onclick();await settle();assert.ok(release);assert.equal(n('download').disabled,true);h.records.find(r=>r.tag.startsWith('osg:participant:')).text='Changed during compression';release();await during;assert.equal(blobs.length,2);assert.match(n('error').textContent,/course changed/);
    await n('preview').onclick();release=null;const otherPane=n('download').onclick();await settle();assert.ok(release);
    const other=await import('../src/word.mjs');await other.updateCourseDetails(await other.readProject(),{audience:'PRIVATE OTHER PANE',purpose:'',authors:'',attribution:''});release();await otherPane;assert.equal(blobs.length,2);assert.match(n('error').textContent,/course changed/);
    globalThis.JSZip=JSZip;h.records.find(r=>r.tag.startsWith('osg:participant:')).text='';await n('preview').onclick();assert.equal(n('download').disabled,true);assert.match(n('workbook-preview').textContent,/1 items need attention/);await n('download').onclick();assert.equal(blobs.length,2);
    h.records.find(r=>r.tag.startsWith('osg:participant:')).text='Restored learner instructions';await n('preview').onclick();assert.equal(n('download').disabled,false);
    h.refreshFailures.add(h.refreshes+1);await n('download').onclick();assert.equal(blobs.length,2);assert.equal(n('download').disabled,true);const counts={saves:h.saves,syncs:h.syncs,runs:h.runs};await n('download').onclick();assert.deepEqual({saves:h.saves,syncs:h.syncs,runs:h.runs},counts);
    h.refreshFailures.clear();await n('recover').onclick();assert.equal(n('refresh').disabled,false);assert.equal(h.paragraphs.length,0);assert.equal(h.cueWrites,0);
  }finally{URL.createObjectURL=oldURL;globalThis.setTimeout=oldTimer;globalThis.JSZip=JSZip;}
});

test('agenda download requires a fresh reviewed schedule, follows start-time edits and blocks stale or interrupted exports',async()=>{
  const p=fixture(),h=makeHost(p),ui=dom();
  globalThis.Office=h.Office;globalThis.Word=h.Word;globalThis.document=ui.document;globalThis.location={search:''};globalThis.JSZip=JSZip;
  const blobs=[],links=[],create=ui.document.createElement,oldURL=URL.createObjectURL,oldTimer=setTimeout;
  ui.document.createElement=tag=>{const e=create(tag);if(tag==='a')links.push(e);return e;};
  URL.createObjectURL=b=>{blobs.push(b);return 'blob:agenda-test';};globalThis.setTimeout=()=>0;
  try{
    await import('../src/app.mjs?agenda-ui');await settle();const n=id=>ui.nodes.get(id);
    assert.equal(n('agenda-download').disabled,true);await n('agenda-download').onclick();assert.match(n('error').textContent,/Refresh and review/);assert.equal(blobs.length,0);
    n('agenda-form').onsubmit({preventDefault(){}});await settle();assert.equal(n('agenda-download').disabled,false);assert.match(n('agenda-summary').textContent,/Project ID:.*Generation revision/);
    n('agenda-start').value='23:55';n('agenda-start').oninput();assert.equal(n('agenda-download').disabled,true);
    n('agenda-form').onsubmit({preventDefault(){}});await settle();await n('agenda-download').onclick();
    assert.equal(blobs.length,1);assert.match(links[0].download,/ - Agenda - .*\.docx$/);
    const zip=await JSZip.loadAsync(await blobs[0].arrayBuffer()),doc=await zip.file('word/document.xml').async('string');
    assert.match(doc,/23:55–Day 2 · 00:05/);assert.doesNotMatch(doc,/PRIVATE DIRECTIONS|SECRET ANSWER|Learner task/);assert.equal(h.paragraphs.length,0);assert.equal(h.cueWrites,0);
    h.records.find(r=>r.tag.startsWith('osg:activity:')).text='Directly edited schedule item';
    await n('agenda-download').onclick();assert.match(n('error').textContent,/course changed/);assert.equal(blobs.length,1);assert.equal(n('agenda-download').disabled,true);
    n('agenda-form').onsubmit({preventDefault(){}});await settle();await n('agenda-download').onclick();assert.equal(blobs.length,2);
    assert.match(await (await JSZip.loadAsync(await blobs[1].arrayBuffer())).file('word/document.xml').async('string'),/Directly edited schedule item/);
    let release;globalThis.JSZip=class extends JSZip{generateAsync(options){return new Promise(resolve=>{release=()=>super.generateAsync(options).then(resolve);});}};
    const pending=n('agenda-download').onclick();await settle();assert.ok(release);assert.equal(n('agenda-download').disabled,true);
    n('agenda-start').value='10:00';n('agenda-start').oninput();release();await pending;assert.equal(blobs.length,2);assert.match(n('error').textContent,/agenda changed/);
    n('agenda-form').onsubmit({preventDefault(){}});await settle();
    release=null;const during=n('agenda-download').onclick();await settle();assert.ok(release);
    h.records.find(r=>r.tag.startsWith('osg:activity:')).text='Changed during compression';release();await during;
    assert.equal(blobs.length,2);assert.match(n('error').textContent,/agenda changed/);
    n('agenda-form').onsubmit({preventDefault(){}});await settle();
    release=null;const otherPane=n('agenda-download').onclick();await settle();assert.ok(release);
    const other=await import('../src/word.mjs');
    await other.updateCourseDetails(await other.readProject(),{audience:'Changed in another pane',purpose:'',authors:'',attribution:''});
    release();await otherPane;assert.equal(blobs.length,2);assert.match(n('error').textContent,/agenda changed/);
    globalThis.JSZip=JSZip;n('agenda-form').onsubmit({preventDefault(){}});await settle();
    h.refreshFailures.add(h.refreshes+1);await n('agenda-download').onclick();assert.equal(blobs.length,2);assert.equal(n('agenda-download').disabled,true);
    const counts={runs:h.runs,saves:h.saves,syncs:h.syncs};await n('agenda-download').onclick();assert.deepEqual({runs:h.runs,saves:h.saves,syncs:h.syncs},counts);
    h.refreshFailures.clear();await n('recover').onclick();assert.equal(n('refresh').disabled,false);
  }finally{URL.createObjectURL=oldURL;globalThis.setTimeout=oldTimer;globalThis.JSZip=JSZip;}
});

test('course details editor loads explicitly, preserves stale drafts, invalidates outputs and recovers uncertain saves',async()=>{
  const p=fixture(),h=makeHost(p),ui=dom();
  globalThis.Office=h.Office;globalThis.Word=h.Word;globalThis.document=ui.document;globalThis.location={search:''};
  await import('../src/app.mjs?course-details-ui');await settle();const n=id=>ui.nodes.get(id);
  assert.equal(n('course-details').hidden,false);assert.equal(n('course-details-form').querySelector().disabled,true);
  await n('course-details-load').onclick();assert.equal(n('course-details-form').querySelector().disabled,false);assert.equal(n('course-detail-audience').value,'');
  const fields={audience:'Private course audience',purpose:'Private course purpose',authors:'Author One\nAuthor Two',attribution:'Private rights notes'};
  for(const [key,value] of Object.entries(fields))n('course-detail-'+key).value=value;
  await n('preview').onclick();assert.equal(n('download').disabled,false);
  n('course-details-form').onsubmit({preventDefault(){}});await settle();
  assert.deepEqual(h.durable['openstage.guide.project.v1'].details,fields);assert.equal(n('download').disabled,true);assert.equal(n('workbook-preview').children.length,0);
  assert.equal(h.paragraphs.length,0);assert.equal(h.cueWrites,0);
  await n('preview').onclick();for(const value of Object.values(fields))assert.ok(!n('workbook-preview').textContent.includes(value));
  n('course-detail-purpose').value='Unsubmitted purpose draft';
  h.records.find(r=>r.tag.startsWith('osg:facilitator:')).text='Newer direct Word note';
  const saves=h.saves;n('course-details-form').onsubmit({preventDefault(){}});await settle();
  assert.match(n('error').textContent,/changed/);assert.equal(n('course-detail-purpose').value,'Unsubmitted purpose draft');assert.equal(h.saves,saves);
  await n('course-details-load').onclick();assert.equal(n('course-detail-purpose').value,fields.purpose);
  n('course-detail-purpose').value='Retain this draft after uncertain save';h.saveFailures.set(h.saves+2,{persist:false});
  n('course-details-form').onsubmit({preventDefault(){}});await settle();
  assert.equal(n('course-details-form').querySelector().disabled,true);assert.equal(n('course-details-load').disabled,true);assert.equal(n('course-detail-purpose').disabled,true);
  assert.equal(n('course-detail-purpose').value,'Retain this draft after uncertain save');assert.equal(n('download').disabled,true);
  const counts={saves:h.saves,syncs:h.syncs};n('course-details-form').onsubmit({preventDefault(){}});await n('course-details-load').onclick();await settle();assert.deepEqual({saves:h.saves,syncs:h.syncs},counts);
  await n('recover').onclick();assert.equal(n('course-details-load').disabled,false);
  assert.equal(n('course-detail-purpose').value,'Retain this draft after uncertain save');
  await n('course-details-load').onclick();assert.equal(n('course-detail-purpose').value,'Retain this draft after uncertain save');
  assert.equal(h.records.find(r=>r.tag.startsWith('osg:facilitator:')).text,'Newer direct Word note');assert.equal(h.paragraphs.length,0);
});

test('existing activity editor rejects stale drafts, reloads direct edits, refreshes exports and retains drafts after uncertain saves',async()=>{
  const p=fixture(),h=makeHost(p),ui=dom();
  globalThis.Office=h.Office;globalThis.Word=h.Word;globalThis.document=ui.document;globalThis.location={search:''};
  await import('../src/app.mjs?activity-settings-ui');await settle();const n=id=>ui.nodes.get(id),id=p.modules[0].activities[0].id;
  assert.equal(n('edit-activity-id').value,id);assert.equal(n('edit-activity-form').querySelector().disabled,true);
  await n('edit-activity-load').onclick();assert.equal(n('edit-activity-minutes').value,'10');assert.equal(n('edit-activity-form').querySelector().disabled,false);
  n('edit-activity-minutes').value='7.5';n('edit-activity-lines').value='8';
  h.records.find(r=>r.tag===`osg:participant:${id}`).text='New Word instructions after loading';
  n('edit-activity-form').onsubmit({preventDefault(){}});await settle();
  assert.match(n('error').textContent,/changed/);assert.equal(n('edit-activity-minutes').value,'7.5');assert.equal(h.cueWrites,0);
  await n('edit-activity-load').onclick();n('edit-activity-minutes').value='7.5';n('edit-activity-lines').value='8';
  n('edit-activity-form').onsubmit({preventDefault(){}});await settle();assert.equal(h.cueWrites,1);
  await n('preview').onclick();assert.equal(n('download').disabled,false);assert.match(n('workbook-preview').textContent,/New Word instructions/);
  assert.equal(n('workbook-preview').children.filter(e=>e.className==='writing-line').length,8);
  assert.ok(!n('workbook-preview').textContent.includes('SECRET ANSWER'));
  await ui.routes.find(b=>b.dataset.view==='agenda').onclick();assert.match(n('agenda-summary').textContent,/7.5 minutes/);
  await n('edit-activity-load').onclick();n('edit-activity-minutes').value='9';n('edit-activity-lines').value='4';
  h.saveFailures.set(h.saves+2,{persist:true});n('edit-activity-form').onsubmit({preventDefault(){}});await settle();
  assert.equal(n('edit-activity-minutes').value,'9');assert.equal(n('edit-activity-lines').value,'4');assert.equal(n('edit-activity-form').querySelector().disabled,true);
  assert.equal(n('download').disabled,true);assert.equal(n('agenda-body').children.length,0);
  await n('recover').onclick();assert.equal(h.cueWrites,2);assert.equal(n('edit-activity-minutes').value,'9');
  await n('edit-activity-load').onclick();assert.equal(n('edit-activity-lines').value,'4');
});

test('already-open pane clears cached results and disables operations after final save failure, then recovers safely',async()=>{
  const p=fixture();p.modules[0].objective='';const h=makeHost(p),ui=dom();
  globalThis.Office=h.Office;globalThis.Word=h.Word;globalThis.document=ui.document;globalThis.location={search:''};
  await import('../src/app.mjs');await settle();const n=id=>ui.nodes.get(id);
  assert.equal(n('course-name').textContent,p.title);
  await n('preview').onclick();assert.equal(n('download').disabled,false);assert.ok(n('workbook-preview').children.length>0);
  await ui.routes.find(x=>x.dataset.view==='review').onclick();assert.ok(n('review-list').children.length>0);
  await ui.routes.find(x=>x.dataset.view==='agenda').onclick();assert.ok(n('agenda-body').children.length>0);
  n('module-title').value='New module';n('module-objective').value='';h.saveFailures.set(2,{persist:false});
  n('module-form').onsubmit({preventDefault(){}});await settle();
  for(const id of ['preview','download','refresh','agenda-refresh','review-refresh'])assert.equal(n(id).disabled,true,id);
  for(const id of ['course-form','module-form'])assert.equal(n(id).querySelector().disabled,true,id);
  assert.equal(n('activity-insert').disabled,true,'Actual activity submit');
  assert.equal(n('recover').disabled,false);assert.equal(n('error').hidden,false);assert.match(n('error').textContent,/paused/);
  assert.equal(n('workbook-preview').children.length,0);assert.equal(n('agenda-body').children.length,0);assert.equal(n('review-list').children.length,0);assert.equal(n('agenda-table').hidden,true);
  assert.equal(n('module-form').resetCount,undefined);assert.equal(n('module-title').value,'New module');
  const counts={saves:h.saves,syncs:h.syncs};await n('preview').onclick();await n('download').onclick();await n('refresh').onclick();
  assert.deepEqual({saves:h.saves,syncs:h.syncs},counts);
  h.records.find(x=>x.tag.startsWith('osg:participant:')).text='Manual edit retained after failure';
  const count=h.paragraphs.length;await n('recover').onclick();assert.equal(h.paragraphs.length,count);
  assert.equal(n('preview').disabled,false);assert.equal(n('download').disabled,true);
  await n('preview').onclick();assert.equal(n('download').disabled,false);assert.match(n('workbook-preview').textContent,/Manual edit retained/);
  assert.equal(h.durable['openstage.guide.project.v1'].modules.length,2);
});

test('actual ribbon pane-opening errors immediately render invalidated outputs and keep integrity failures blocked',async()=>{
  const p=fixture();p.modules[0].objective='';const h=makeHost(p),ui=dom(),ribbon=new Map();let paneError=null,completed=0;
  h.Office.actions.associate=(name,handler)=>ribbon.set(name,handler);
  h.Office.addin.showAsTaskpane=async()=>{if(paneError)throw paneError;};
  globalThis.Office=h.Office;globalThis.Word=h.Word;globalThis.document=ui.document;globalThis.location={search:''};
  await import('../src/app.mjs?ribbon-opening-regression');await settle();const n=id=>ui.nodes.get(id);
  const populate=async()=>{
    await n('preview').onclick();await ui.routes.find(x=>x.dataset.view==='review').onclick();await ui.routes.find(x=>x.dataset.view==='agenda').onclick();
    assert.equal(n('download').disabled,false);for(const id of ['workbook-preview','agenda-body','review-list'])assert.ok(n(id).children.length>0,id);
  };
  const empty=()=>{
    for(const id of ['workbook-preview','agenda-body','review-list'])assert.equal(n(id).children.length,0,id);
    assert.equal(n('agenda-table').hidden,true);assert.equal(n('download').disabled,true);assert.equal(n('error').hidden,false);
  };
  assert.equal(ribbon.size,27,'the twenty-one ribbon commands and the six keyboard navigation actions');assert.equal(typeof ribbon.get('openDesigner'),'function');await populate();
  let counts={saves:h.saves,syncs:h.syncs};paneError=new Error('Injected pane-opening failure');
  await ribbon.get('openDesigner')({completed(){completed++;}});
  empty();assert.match(n('error').textContent,/Injected pane-opening failure/);assert.equal(n('preview').disabled,false);
  assert.deepEqual({saves:h.saves,syncs:h.syncs},counts);assert.equal(completed,1);
  paneError=null;await populate();
  counts={saves:h.saves,syncs:h.syncs};paneError=Object.assign(new Error('Injected integrity failure'),{code:'OSG_INTEGRITY'});
  await ribbon.get('participantWorkbook')({completed(){completed++;}});
  empty();assert.match(n('error').textContent,/Injected integrity failure/);assert.equal(completed,2);
  for(const id of ['preview','refresh','agenda-refresh','review-refresh'])assert.equal(n(id).disabled,true,id);
  for(const id of ['course-form','module-form'])assert.equal(n(id).querySelector().disabled,true,id);
  assert.equal(n('activity-insert').disabled,true,'Actual activity submit');
  await n('preview').onclick();await n('download').onclick();await n('refresh').onclick();
  assert.deepEqual({saves:h.saves,syncs:h.syncs},counts);assert.equal(h.paragraphs.length,0);
});

test('verified restore remains available while course operations are paused and never writes to the current guide',async()=>{
  installParser();const {project:p,master}=await masterFixture(),backup=await createCheckpoint(master,p,JSZip,assets),h=makeHost(p),ui=dom(),ribbon=new Map();
  h.Office.actions.associate=(id,fn)=>ribbon.set(id,fn);
  h.Office.addin.showAsTaskpane=async()=>{throw Object.assign(new Error('Paused course test'),{code:'OSG_INTEGRITY'});};
  globalThis.Office=h.Office;globalThis.Word=h.Word;globalThis.document=ui.document;globalThis.location={search:''};globalThis.JSZip=JSZip;
  const blobs=[],links=[],create=ui.document.createElement;
  ui.document.createElement=tag=>{const e=create(tag);if(tag==='a')links.push(e);return e;};
  const originalUrl=URL.createObjectURL,originalTimer=globalThis.setTimeout;
  URL.createObjectURL=blob=>{blobs.push(blob);return 'blob:checkpoint-test';};globalThis.setTimeout=()=>0;
  try{
    await import('../src/app.mjs?backup-ui');await settle();const n=id=>ui.nodes.get(id);
    await ribbon.get('saveBackup')({completed(){}});assert.equal(n('backup-save').disabled,true);assert.equal(n('backup-inspect').disabled,false);
    const counts={saves:h.saves,syncs:h.syncs},state=structuredClone(h.durable),body=structuredClone(h.records);
    n('backup-file').files=[{size:backup.data.length,arrayBuffer:async()=>backup.data.buffer}];n('backup-file').onchange();
    await n('backup-inspect').onclick();assert.equal(n('backup-restore').disabled,false);assert.match(n('backup-details').textContent,/SHA-256/);
    await n('backup-restore').onclick();assert.deepEqual(new Uint8Array(await blobs[0].arrayBuffer()),master);assert.match(links[0].download,/Restored .*\.docx$/);
    assert.deepEqual(h.durable,state);assert.deepEqual(h.records,body);assert.deepEqual({saves:h.saves,syncs:h.syncs},counts);
    n('backup-file').files=[{size:3,arrayBuffer:async()=>new Uint8Array([1,2,3]).buffer}];n('backup-file').onchange();assert.equal(n('backup-restore').disabled,true);
    await n('backup-inspect').onclick();assert.equal(n('backup-restore').disabled,true);assert.equal(blobs.length,1);
    assert.equal(n('preview').disabled,true);assert.equal(n('backup-save').disabled,true);
  }finally{URL.createObjectURL=originalUrl;globalThis.setTimeout=originalTimer;}
});
