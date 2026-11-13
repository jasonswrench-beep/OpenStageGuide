import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import {DOMParser} from '@xmldom/xmldom';
import {pptxFixture} from './helpers/pptx-fixture.mjs';
import {readPptx,pptxDraft} from '../src/pptx.mjs';
import {comparePptx} from '../src/pptx-compare.mjs';
import {fixture} from './helpers/mock-office.mjs';
import {worksheetFiles} from '../src/worksheet.mjs';
globalThis.DOMParser=DOMParser;
async function sample(){const original=await readPptx(await pptxFixture(),'Original.pptx',JSZip),review=pptxDraft(original),incoming=structuredClone(original);incoming.source.sha256='b'.repeat(64);return {original,review,incoming,editor:structuredClone(review.slides[0])};}
const renumber=s=>s.slides.forEach((slide,i)=>slide.order=i+1);
test('identical source is unchanged while local/unsaved review stays separate and inputs remain unchanged',async()=>{
  const x=await sample();x.editor.text='PRIVATE UNSAVED CHANGE';const before=JSON.stringify(x),result=comparePptx(x.original,x.original,x.review,x.editor);assert.equal(result.sameSnapshot,true);assert.equal(result.counts.unchanged,2);assert.equal(result.counts.conflicts,0);assert.equal(result.rows[0].authored.text,'PRIVATE UNSAVED CHANGE');assert.equal(result.rows[0].fields.find(f=>f.key==='text').localChanged,true);assert.equal(JSON.stringify(x),before);result.rows[0].authored.text='later';assert.equal(x.editor.text,'PRIVATE UNSAVED CHANGE');
});
test('insertions/deletions shift absolute positions without falsely declaring relative reorders',async()=>{
  const x=await sample(),added={...x.incoming.slides[0],id:'slide:999',sourceId:999,title:'New content',text:'New body',notes:'',sourcePart:'ppt/slides/new.xml'};x.incoming.slides.unshift(added);renumber(x.incoming);let result=comparePptx(x.original,x.incoming,x.review,x.editor);assert.equal(result.counts.added,1);assert.equal(result.counts.moved,0);assert.equal(result.rows.find(r=>r.id==='slide:258').incomingOrder,2);
  x.incoming.slides=x.incoming.slides.filter(s=>s.id!=='slide:258');renumber(x.incoming);result=comparePptx(x.original,x.incoming,x.review,x.editor);assert.equal(result.counts.missing,1);assert.equal(result.counts.moved,0);assert.equal(result.rows.at(-1).authored.id,'slide:258');
});
test('true relative reordering is proposed without promoting source or assignment baselines',async()=>{
  const x=await sample();x.incoming.slides.reverse();renumber(x.incoming);x.review.slides[0].moduleId='review-only-module';const result=comparePptx(x.original,x.incoming,x.review,x.editor);assert.equal(result.counts.moved,1);assert.equal(result.counts.changed,0);assert.equal(x.original.slides[0].sourceId,258);assert.equal(x.review.sourceHash,x.original.source.sha256);assert.equal(x.review.slides[0].moduleId,'review-only-module');
});
test('three-way field conflicts distinguish divergent local changes from identical edits',async()=>{
  const x=await sample();x.editor.title='Local title';x.incoming.slides[0].title='Incoming title';x.editor.notes='Same new note';x.incoming.slides[0].notes='Same new note';const result=comparePptx(x.original,x.incoming,x.review,x.editor),row=result.rows[0];assert.equal(row.conflict,true);assert.equal(row.fields.find(f=>f.key==='title').conflict,true);assert.equal(row.fields.find(f=>f.key==='notes').conflict,false);assert.equal(row.original.title,'Title 2');assert.equal(row.authored.title,'Local title');assert.equal(row.incoming.title,'Incoming title');
});
test('reused/replaced IDs and duplicate source content remain ambiguous candidates, never automatic mappings',async()=>{
  const x=await sample();const old=x.original.slides[0];x.incoming.slides[0]={...structuredClone(old),id:'slide:999',sourceId:999,sourcePart:'ppt/slides/new.xml'};let result=comparePptx(x.original,x.incoming,x.review,x.editor);assert.deepEqual(result.rows[0].contentCandidates,['slide:258']);assert.equal(result.rows[0].kind,'added-candidate');assert.equal(result.rows.at(-1).kind,'missing-candidate');assert.equal(result.rows.at(-1).incoming,null);
  x.incoming.slides[0]={...structuredClone(x.original.slides[1]),id:old.id,sourceId:old.sourceId,order:1};result=comparePptx(x.original,x.incoming,x.review,x.editor);assert.ok(result.rows[0].warnings.some(w=>/reused/.test(w)));assert.match(result.warning,/not proof of deck lineage/);
});
test('invalid inputs are refused and all private comparison fields remain excluded from learner files',async()=>{
  const x=await sample(),p=fixture();x.editor.notes='PRIVATE COMPARISON REVIEW';const result=comparePptx(x.original,x.incoming,x.review,x.editor);p.slideComparison=result;for(const text of Object.values(worksheetFiles(p)))assert.doesNotMatch(text,/PRIVATE COMPARISON REVIEW|PRIVATE SPEAKER NOTE/);
  const invalid=structuredClone(x.incoming);invalid.slides[1].id=invalid.slides[0].id;assert.throws(()=>comparePptx(x.original,invalid,x.review,x.editor));assert.throws(()=>comparePptx(x.original,x.incoming,x.review,{...x.editor,id:'slide:999'}));
});
