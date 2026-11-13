import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import {DOMParser} from '@xmldom/xmldom';
import {fixture} from './helpers/mock-office.mjs';
import {worksheetFiles,worksheetBlob,worksheetName,WORKBOOK_SCOPE} from '../src/worksheet.mjs';

test('workbook identity and learner content use an isolated allowlisted package',async()=>{
  const p=fixture(),before=JSON.stringify(p);p.details={audience:'PRIVATE AUDIENCE',purpose:'PRIVATE PURPOSE',authors:'PRIVATE AUTHORS',attribution:'PRIVATE RIGHTS'};p.source={secret:'PRIVATE SOURCE'};
  const input=JSON.stringify(p),files=worksheetFiles(p,'2026-10-25T12:00:00Z');assert.equal(Object.keys(files).length,6);
  for(const body of Object.values(files)){
    assert.doesNotMatch(body,/PRIVATE|SECRET|customXml|openstage.guide.project/);
    const errors=[];new DOMParser({errorHandler:{warning:m=>errors.push(m),error:m=>errors.push(m),fatalError:m=>errors.push(m)}}).parseFromString(body,'application/xml');assert.deepEqual(errors,[]);
  }
  const doc=files['word/document.xml'];assert.ok(doc.includes(p.id));assert.ok(doc.includes(WORKBOOK_SCOPE));assert.match(doc,new RegExp(`Generation revision: ${p.revision}`));assert.match(doc,/Learner task\./);assert.match(doc,/Practice a response\./);
  assert.equal((doc.match(/<w:bottom /g)||[]).length,2);
  const blob=await worksheetBlob(p,JSZip),zip=await JSZip.loadAsync(await blob.arrayBuffer());assert.deepEqual(Object.values(zip.files).filter(f=>!f.dir).map(f=>f.name).sort(),Object.keys(files).sort());
  for(const file of Object.values(zip.files).filter(f=>!f.dir))assert.doesNotMatch(await file.async('string'),/PRIVATE|SECRET/);
  assert.equal(JSON.stringify(p),input);assert.notEqual(input,before);
});
test('workbook identity is independent of agenda start time and safely escapes course text',()=>{
  const p=fixture();p.title='Course <&> "\'\u0001';p.modules[0].objective='';p.modules[0].activities[0].lines=0;
  p.agendaStart='09:00';const a=worksheetFiles(p,'2026-10-25T12:00:00Z');p.agendaStart='23:55';assert.deepEqual(worksheetFiles(p,'2026-10-25T12:00:00Z'),a);
  const doc=a['word/document.xml'];assert.match(doc,/Course &lt;&amp;&gt; &quot;&apos;/);assert.ok(!doc.includes('\u0001'));assert.doesNotMatch(doc,/Learning objective|<w:bottom /);
});
test('workbook new-copy names are safe and unique and required learner issues still block generation',()=>{
  const a=worksheetName('../Course:*?'),b=worksheetName('../Course:*?');assert.notEqual(a,b);assert.match(a,/^Course - Participant Workbook - [a-f0-9-]+\.docx$/);assert.match(worksheetName('...'),/^Course - Participant Workbook/);
  const p=fixture();p.modules[0].activities[0].participant='';assert.throws(()=>worksheetFiles(p),/Review Guide/);
});
