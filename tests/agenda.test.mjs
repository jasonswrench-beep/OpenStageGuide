import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import {DOMParser} from '@xmldom/xmldom';
import {fixture} from './helpers/mock-office.mjs';
import {addActivity} from '../src/model.mjs';
import {agendaPreview,agendaFiles,agendaBlob,agendaName,AGENDA_SCOPE} from '../src/agenda.mjs';

test('agenda output is a six-part allowlisted package with course identity and semantic headings',async()=>{
  const p=fixture();p.details={audience:'PRIVATE AUDIENCE',purpose:'PRIVATE PURPOSE',authors:'PRIVATE AUTHORS',attribution:'PRIVATE RIGHTS'};
  p.modules[0].objective='PRIVATE OBJECTIVE';p.modules[0].activities[0].participant='PRIVATE LEARNER TEXT';p.source={note:'PRIVATE SOURCE'};
  const files=agendaFiles(p,'09:00','2026-10-25T12:00:00.000Z');assert.equal(Object.keys(files).length,6);
  for(const body of Object.values(files)){
    const errors=[];new DOMParser({errorHandler:{warning:m=>errors.push(m),error:m=>errors.push(m),fatalError:m=>errors.push(m)}}).parseFromString(body,'application/xml');assert.deepEqual(errors,[]);
    assert.doesNotMatch(body,/PRIVATE|SECRET|Learner task|openstage.guide.project|customXml|facilitator|answers/);
  }
  const doc=files['word/document.xml'];assert.ok(doc.includes(p.id));assert.match(doc,new RegExp(`Generation revision: ${p.revision}`));assert.ok(doc.includes(AGENDA_SCOPE));
  assert.match(doc,/w:pStyle w:val="Heading1"/);assert.match(doc,/w:pStyle w:val="Heading2"/);
  const blob=await agendaBlob(p,'09:00',JSZip),zip=await JSZip.loadAsync(await blob.arrayBuffer());
  assert.deepEqual(Object.values(zip.files).filter(f=>!f.dir).map(f=>f.name).sort(),Object.keys(files).sort());
});
test('agenda preserves fractional durations, seconds and day rollover',()=>{
  let p=fixture();p.modules[0].activities[0].minutes=12.5;
  p=addActivity(p,p.modules[0].id,{title:'Closing',minutes:3.5,participant:'Task',lines:0}).next;
  const a=agendaPreview(p,'23:50');assert.equal(a.end,'Day 2 · 00:06');assert.equal(a.minutes,16);assert.equal(a.rows[0].end,'Day 2 · 00:02:30');
  assert.match(agendaFiles(p,'23:50')['word/document.xml'],/12.5 minutes/);
});
test('agenda escapes every user-facing field and strips XML control characters',()=>{
  const p=fixture();p.title='Course <&> "\'\u0001';p.modules[0].title='Module <&>';p.modules[0].activities[0].title='Activity <&>\nNext line';
  const f=agendaFiles(p),doc=f['word/document.xml'];assert.match(doc,/Course &lt;&amp;&gt; &quot;&apos;/);assert.match(doc,/Module &lt;&amp;&gt;/);assert.match(doc,/Activity &lt;&amp;&gt;/);assert.ok(!doc.includes('\u0001'));assert.match(doc,/<w:br\/>/);
});
test('agenda rejects invalid timing, empty schedules and missing titles',()=>{
  const p=fixture();assert.throws(()=>agendaFiles(p,'24:00'),/valid start/);
  p.modules[0].activities[0].minutes=-1;assert.throws(()=>agendaFiles(p));
  p.modules[0].activities[0].minutes=10;p.modules[0].activities[0].title=' ';assert.throws(()=>agendaFiles(p),/titles/);
  p.modules[0].activities=[];assert.throws(()=>agendaFiles(p),/Add an activity/);
});
test('preview discloses unresolved review items without emitting teaching content into agenda files',()=>{
  const p=fixture();p.modules[0].activities[0].participant='';p.modules[0].activities[0].facilitator='TODO PRIVATE NOTE';
  const a=agendaPreview(p);assert.ok(a.issues.some(i=>i.severity==='error'));assert.ok(!JSON.stringify(a).includes('PRIVATE NOTE'));
  assert.match(agendaFiles(p)['word/document.xml'],/Original activity/);
});
test('agenda filenames are safe and unique for separate new outputs',()=>{
  const a=agendaName('../Course:*?'),b=agendaName('../Course:*?');assert.notEqual(a,b);assert.match(a,/^Course - Agenda - [a-f0-9-]+\.docx$/);assert.match(agendaName('...'),/^Course - Agenda/);
});
