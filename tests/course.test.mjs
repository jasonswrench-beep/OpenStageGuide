import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import {course,addModule,addActivity,reconcile,participantView,total,validate} from '../src/model.mjs';
import {worksheetFiles} from '../src/worksheet.mjs';
function fixture(){
  let p=course('Conflict & communication');
  p=addModule(p,'Listen before responding','Demonstrate an active listening response.').next;
  return addActivity(p,p.modules[0].id,{title:'Reframe the conversation',icon:'Discussion',minutes:12,facilitator:'PRIVATE FACILITATOR DIRECTIONS',participant:'Write a response to the scenario.\nExplain your reasoning.',answers:'SECRET ANSWER KEY',lines:5}).next;
}
function records(p){return [{tag:`osg:course:${p.id}`,text:p.title},...p.modules.flatMap(m=>[{tag:`osg:module:${m.id}`,text:m.title},{tag:`osg:objective:${m.id}`,text:m.objective},...m.activities.flatMap(a=>['activity','facilitator','participant','answers'].map(type=>({tag:`osg:${type}:${a.id}`,text:type==='activity'?a.title:a[type]})))])];}
test('learner export never packages instructor fields, including hidden metadata',async()=>{
  const p=fixture(), files=worksheetFiles(p), zip=new JSZip();
  for(const [name,data] of Object.entries(files))zip.file(name,data);
  const loaded=await JSZip.loadAsync(await zip.generateAsync({type:'nodebuffer'}));
  for(const [name,file] of Object.entries(loaded.files)){
    if(file.dir)continue;
    const text=await file.async('string');assert.ok(!text.includes('SECRET ANSWER KEY'),name);assert.ok(!text.includes('PRIVATE FACILITATOR DIRECTIONS'),name);assert.ok(!name.includes('customXml'));
  }
  assert.match(files['word/document.xml'],/Conflict &amp; communication/);
  assert.match(files['word/document.xml'],/Write a response/);
  assert.equal(Object.hasOwn(participantView(p).modules[0].activities[0],'answers'),false);
});
test('direct Word edits flow into learner exports',()=>{
  const p=fixture(),r=records(p);r.find(x=>x.tag.startsWith('osg:participant:')).text='Revised learner task';
  const next=reconcile(p,r);assert.equal(next.revision,p.revision+1);assert.match(worksheetFiles(next)['word/document.xml'],/Revised learner task/);
  assert.equal(p.modules[0].activities[0].participant.startsWith('Write'),true);
});
test('a missing field or duplicate pasted field pauses generation',()=>{
  const p=fixture(),r=records(p);assert.throws(()=>reconcile(p,r.slice(1)),/deleted/);assert.throws(()=>reconcile(p,[...r,r[0]]),/copied/);
});
test('timing and IDs remain valid across independent additions',()=>{
  const p=fixture();assert.equal(total(p),12);const next=addModule(p,'Practice','').next;assert.notEqual(next.modules[0].id,next.modules[1].id);assert.equal(p.modules.length,1);
  next.modules[1].id=next.modules[0].id;assert.throws(()=>validate(next),/duplicate/);
});
test('reject invalid duration and missing learner instructions',()=>{
  const p=fixture();assert.throws(()=>addActivity(p,p.modules[0].id,{title:'Bad',minutes:-1}),/duration/);
  assert.throws(()=>addActivity(p,p.modules[0].id,{title:'Bad',minutes:10,participant:' '}),/participant instructions/);
});
test('XML injection cannot create extra document elements',()=>{
  const p=fixture();p.modules[0].activities[0].participant='<w:altChunk/> & "quotes"';
  const output=worksheetFiles(p)['word/document.xml'];assert.ok(!output.includes('<w:altChunk/>'));assert.ok(output.includes('&lt;w:altChunk/&gt;'));
});

test('consecutive writing paragraphs retain an interior rule in desktop Word',()=>{
  const p=fixture();
  const output=worksheetFiles(p)['word/document.xml'];
  const rules=[...output.matchAll(/<w:p>\s*<w:pPr><w:spacing[^]*?<\/w:pPr><w:r><w:t> <\/w:t><\/w:r><\/w:p>/g)].map(m=>m[0]);
  assert.equal(rules.length,5);
  // Word groups adjacent identical paragraph borders; a bottom border alone
  // draws only the last line. Between borders retain each interior boundary.
  for(const rule of rules){assert.match(rule,/<w:between\b/);assert.match(rule,/<w:bottom\b/);}
});
