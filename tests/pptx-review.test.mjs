import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import {DOMParser} from '@xmldom/xmldom';
import {pptxFixture} from './helpers/pptx-fixture.mjs';
import {readPptx,pptxDraft} from '../src/pptx.mjs';
import {serializeReview,restoreReview,parseReviewRecord,MAX_REVIEW_BYTES,MAX_EDITOR_BYTES} from '../src/pptx-review.mjs';
import {fixture} from './helpers/mock-office.mjs';
import {worksheetFiles} from '../src/worksheet.mjs';
globalThis.DOMParser=DOMParser;
async function saved(){const bytes=await pptxFixture(),source=await readPptx(bytes,'Original.pptx',JSZip),review=pptxDraft(source);review.slides[0].text='Saved author review';review.slides[0].notes='PRIVATE REVIEW NOTE';review.slides[0].moduleId='former-module';review.slides[0].lesson='Lesson 1';const editor={...review.slides[0],text:'Unsaved author edit',moduleId:'different-former-module'};return {bytes,source,review,editor,text:serializeReview(source,review,editor)};}
test('new-session restore reparses matching bytes and preserves saved plus unsaved authored review independently',async()=>{
  const x=await saved(),before=x.bytes.slice(),restored=await restoreReview(x.text,x.bytes,'Renamed.pptx',JSZip);
  assert.equal(restored.source.source.name,'Original.pptx');assert.equal(restored.selectedName,'Renamed.pptx');assert.equal(serializeReview(restored.source,restored.review,restored.unsavedEditor),x.text);assert.equal(restored.source.source.sha256,x.source.source.sha256);assert.equal(restored.review.slides[0].text,'Saved author review');assert.equal(restored.unsavedEditor.text,'Unsaved author edit');assert.equal(restored.review.slides[0].moduleId,'former-module');assert.equal(restored.unsavedEditor.moduleId,'different-former-module');
  restored.review.slides[0].text='Later edit';assert.equal(x.review.slides[0].text,'Saved author review');assert.equal(restored.source.slides[0].text,x.source.slides[0].text);assert.deepEqual(x.bytes,before);assert.ok(Object.isFrozen(restored.source.slides[0]));
});
test('restore refuses different source snapshots and edited baselines even with plausible retained hashes',async()=>{
  const x=await saved(),other=await pptxFixture(z=>z.file('docProps/change.xml','<changed/>'));
  await assert.rejects(restoreReview(x.text,other,'Other.pptx',JSZip),/mismatched/);
  for(const change of [r=>r.source.slides[0].text='Forged baseline',r=>r.source.slides[0].id='slide:999',r=>r.source.warnings=[],r=>r.source.slides.reverse(),r=>r.review.sourceHash='0'.repeat(64),r=>r.review.slides.reverse(),r=>r.unsavedEditor.id='slide:999']){
    const r=JSON.parse(x.text);change(r);await assert.rejects(restoreReview(JSON.stringify(r),x.bytes,'Original.pptx',JSZip));
  }
});
test('restore rejects unsupported shape/version, duplicate IDs, oversized or malformed records',async()=>{
  const x=await saved();
  for(const change of [r=>r.version=2,r=>r.privateProject={secret:'no'},r=>r.review.slides[0].answers='no',r=>r.source.source.secret='no',r=>r.review.slides[0].title=8,r=>r.review.slides[0].notes='x'.repeat(50001),r=>r.source.slides[1].sourceId=r.source.slides[0].sourceId]){
    const r=JSON.parse(x.text);change(r);await assert.rejects(restoreReview(JSON.stringify(r),x.bytes,'Original.pptx',JSZip));
  }
  await assert.rejects(restoreReview('{bad',x.bytes,'Original.pptx',JSZip));await assert.rejects(restoreReview(' '.repeat(MAX_REVIEW_BYTES+1),x.bytes,'Original.pptx',JSZip));
});
test('failed restore does not mutate exported records, sources or previous authored drafts',async()=>{
  const x=await saved(),before=JSON.stringify(x.review);await assert.rejects(restoreReview(x.text,new Uint8Array(100),'Bad.pptx',JSZip));assert.equal(JSON.stringify(x.review),before);assert.equal(JSON.parse(x.text).unsavedEditor.text,'Unsaved author edit');
  assert.throws(()=>serializeReview(x.source,x.review,{...x.editor,notes:'x'.repeat(50001)}));assert.equal(JSON.stringify(x.review),before);
});
test('restored source notes and authored review remain excluded from every participant package part',async()=>{
  const x=await saved(),restored=await restoreReview(x.text,x.bytes,'Original.pptx',JSZip),p=fixture();p.slideReview=restored;
  for(const part of Object.values(worksheetFiles(p)))assert.doesNotMatch(part,/PRIVATE REVIEW NOTE|PRIVATE SPEAKER NOTE|Saved author review|Unsaved author edit|former-module/);
  assert.doesNotThrow(()=>serializeReview(restored.source,restored.review,restored.unsavedEditor));
});
test('maximum JSON-escaping unsaved editor exports, restores and reexports without changing saved review',async()=>{
  const x=await saved(),editor={id:x.editor.id,...Object.fromEntries(['title','text','notes','moduleId','lesson'].map(k=>[k,'\u0000'.repeat(50000)]))};
  assert.ok(new TextEncoder().encode(JSON.stringify(editor)).length>1024*1024);
  const text=serializeReview(x.source,x.review,editor),restored=await restoreReview(text,x.bytes,'Original.pptx',JSZip);
  assert.deepEqual(restored.review,x.review);assert.deepEqual(restored.unsavedEditor,editor);assert.equal(serializeReview(restored.source,restored.review,restored.unsavedEditor),text);
  assert.doesNotThrow(()=>serializeReview(restored.source,restored.review,{...restored.review.slides[1]}));
});
test('one consistent near-cap base-plus-editor budget accepts its exports and rejects unsafe proposals atomically',async()=>{
  const x=await saved(),source=structuredClone(x.source);source.slides=Array.from({length:310},(_,i)=>({...source.slides[0],id:`slide:${256+i}`,sourceId:256+i,sourcePart:`ppt/slides/slide${i+1}.xml`,order:i+1,title:'',text:'x'.repeat(50000),notes:'',warnings:[]}));
  const review=pptxDraft(source),editor={...review.slides[0]};
  const raw=()=>JSON.stringify({format:'openstage-pptx-review',version:1,source,review,unsavedEditor:editor}),enc=new TextEncoder();
  let remaining=MAX_REVIEW_BYTES-MAX_EDITOR_BYTES-(enc.encode(raw()).length-enc.encode(JSON.stringify(editor)).length);
  for(const s of source.slides){if(!remaining)break;const amount=Math.min(50000,remaining);s.notes='x'.repeat(amount);remaining-=amount;}
  assert.equal(remaining,0);const maxEditor={id:editor.id,...Object.fromEntries(['title','text','notes','moduleId','lesson'].map(k=>[k,'\u0000'.repeat(50000)]))};
  const text=serializeReview(source,review,maxEditor),parsed=parseReviewRecord(text);assert.ok(enc.encode(text).length<=MAX_REVIEW_BYTES);assert.equal(serializeReview(parsed.source,parsed.review,parsed.unsavedEditor),text);
  assert.doesNotThrow(()=>serializeReview(parsed.source,parsed.review,parsed.review.slides[0]));
  const before=serializeReview(source,review,editor),proposed=structuredClone(review);proposed.slides[0].notes='x';assert.throws(()=>serializeReview(source,proposed,editor),/editing space/);assert.equal(serializeReview(source,review,editor),before);
  const incoming=JSON.parse(before);incoming.review.slides[0].notes='x';assert.throws(()=>parseReviewRecord(JSON.stringify(incoming)),/editing space/);
});
