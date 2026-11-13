import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import {fixture} from './helpers/mock-office.mjs';
import {auditExport} from './helpers/export-contract.mjs';
import {publicOutputText} from '../src/output-text.mjs';
import {worksheetFiles,worksheetBlob,workbookPreview} from '../src/worksheet.mjs';
import {agendaFiles,agendaBlob,agendaPreview} from '../src/agenda.mjs';
import {materialsBlob,materialsPreview} from '../src/materials-output.mjs';
const stamp='2026-10-26T05:00:00.000Z';
test('workbook and agenda refuse invalid selected public text without altering model-valid source',()=>{
  for(const bad of ['\ufffe','\uffff','\ud800','\udfff'])for(const field of ['course','module','activity','objective','participant']){
    const p=fixture();if(field==='course')p.title+=' '+bad;if(field==='module')p.modules[0].title+=' '+bad;if(field==='activity')p.modules[0].activities[0].title+=' '+bad;if(field==='objective')p.modules[0].objective+=' '+bad;if(field==='participant')p.modules[0].activities[0].participant+=' '+bad;const before=structuredClone(p);
    assert.throws(()=>worksheetFiles(p,stamp),/unsupported character/);assert.throws(()=>workbookPreview(p),/unsupported character/);if(['course','module','activity'].includes(field)){assert.throws(()=>agendaFiles(p,'09:00',stamp),/unsupported character/);assert.throws(()=>agendaPreview(p),/unsupported character/);}else assert.doesNotThrow(()=>agendaFiles(p,'09:00',stamp));assert.deepEqual(p,before);
  }
});
test('shared reviewed public projection keeps existing safe control stripping and normalizes line endings without changing source',()=>{
  const original='A\u0000\u0001\u000b\r\nB\rC\t☐🙂';assert.equal(publicOutputText(original),'A\nB\nC\t☐🙂');const p=fixture();p.title=original;p.modules[0].activities[0].materials=[{item:'Sup\u0000ply',quantity:1,unit:'s\u000bets',use:'consumable',notes:'PRIVATE'}];const before=structuredClone(p);assert.equal(workbookPreview(p).title,'A\nB\nC\t☐🙂');assert.equal(agendaPreview(p).title,'A\nB\nC\t☐🙂');assert.equal(materialsPreview(p,stamp).rows[0].item,'Supply');assert.equal(materialsPreview(p,stamp).rows[0].unit,'sets');assert.deepEqual(p,before);
});
test('all public DOCX outputs preserve valid multilingual text emoji newlines and checkbox characters',async()=>{
  const p=fixture();p.title='课程 Café 🙂';p.modules[0].title='العربية';p.modules[0].objective='目标 🚀\nSecond line';const a=p.modules[0].activities[0];a.title='☐ 合作 🙂\nتعلّم';a.participant='दिशा 🙂\nNext step';a.materials=[{item:'中文 🚀',quantity:1.125,unit:'组\tsets',use:'reusable',notes:'PRIVATE_NOTE'}];
  for(const kind of ['workbook','agenda','materials']){const blob=kind==='workbook'?await worksheetBlob(p,JSZip):kind==='agenda'?await agendaBlob(p,'09:00',JSZip):await materialsBlob(p,stamp,JSZip);await auditExport(blob,{project:p,privateCanaries:['PRIVATE_NOTE','PRIVATE DIRECTIONS','SECRET ANSWER'],kind});}
});
test('invalid characters in excluded private fields remain unexported without rejecting unrelated public content',async()=>{
  const p=fixture(),a=p.modules[0].activities[0];p.details={audience:'PRIVATE_DETAIL \ufffe',purpose:'',authors:'',attribution:''};a.facilitator='PRIVATE_FACILITATOR \ufffe';a.answers='PRIVATE_ANSWER \ud800';a.materials=[{item:'Public sheets',quantity:1,unit:'sets',use:'consumable',notes:'PRIVATE_NOTE \uffff'}];p.unknown='PRIVATE_METADATA \udfff';const before=structuredClone(p),privateCanaries=[p.details.audience,a.facilitator,a.answers,a.materials[0].notes,p.unknown];
  for(const kind of ['workbook','agenda','materials']){const blob=kind==='workbook'?await worksheetBlob(p,JSZip):kind==='agenda'?await agendaBlob(p,'09:00',JSZip):await materialsBlob(p,stamp,JSZip);await auditExport(blob,{project:p,privateCanaries,kind});}assert.deepEqual(p,before);
  const other=structuredClone(p);other.modules[0].activities[0].materials[0].item='PRIVATE_MATERIAL_FOR_LEARNER \ufffe';const saved=structuredClone(other);for(const kind of ['workbook','agenda'])await auditExport(kind==='workbook'?await worksheetBlob(other,JSZip):await agendaBlob(other,'09:00',JSZip),{project:other,privateCanaries:[...privateCanaries,other.modules[0].activities[0].materials[0].item],kind});await assert.rejects(materialsBlob(other,stamp,JSZip),/unsupported character/);assert.deepEqual(other,saved);
});
