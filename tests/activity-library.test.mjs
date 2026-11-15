// The activity library (requested 2026-11-07: "we should probably have at least a few examples"; the author clicked it and could not
// tell what to do): ready-made activities and the author's own, each put into the guide as an Activity written as labelled parts.
// Drives the real pane module and the real taskpane.html ids in memory. It is not native Word.
import test from 'node:test';
import assert from 'node:assert/strict';
import {STARTER_ACTIVITIES,STORE_KEY,MAX_MINE,blockFromEntry,entryFromBlock,cleanMine,addMine,removeMine,matches} from '../src/activity-library.mjs';
import {ACTIVITY_PARTS} from '../src/guide-cues.mjs';
import {readBlock} from '../src/guide-model.mjs';
import {started,click,addModule,addLesson,addBlock} from './helpers/guide-harness.mjs';

const memoryStorage=(start={})=>{const data={...start};return {data,getItem:k=>k in data?data[k]:null,setItem:(k,v)=>{data[k]=String(v);}};};
const lines=t=>String(t||'').split('\n').map(l=>l.trim()).filter(Boolean);
const withLibrary=async(storage=memoryStorage())=>{
  const a=await started('Reframing conflict',{storage});
  await addModule(a,'Module one');await addLesson(a,'Lesson one');
  a.p.useTool('activityLibrary');
  await a.p.openTool('activityLibrary');
  return {a,storage};
};
const itemTitles=(a,id)=>a.n(id).children.map(li=>li.children[0].textContent);
const buttonOf=(li,label)=>li.children[3].children.find(b=>b.textContent===label);
const activityAt=async a=>{const paras=await a.adapter.paragraphs();const i=paras.findIndex(p=>/^\s*Activity\b/.test(String(p.text)));return {paras,i};};

test('there are eight ready-made activities, each whole: a title, minutes, a summary, an objective and steps',()=>{
  assert.equal(STARTER_ACTIVITIES.length,8);
  assert.equal(new Set(STARTER_ACTIVITIES.map(e=>e.id)).size,8,'every id is different');
  assert.equal(new Set(STARTER_ACTIVITIES.map(e=>e.title)).size,8,'and every title');
  for(const e of STARTER_ACTIVITIES){
    assert.ok(e.title&&e.title.length<=60,`${e.id} title`);
    assert.ok(Number.isInteger(e.minutes)&&e.minutes>=1&&e.minutes<=60,`${e.id} minutes`);
    assert.ok(e.summary&&e.summary.length<=100,`${e.id} summary is one short line`);
    assert.ok(e.objective,`${e.id} objective`);
    assert.ok(lines(e.instructions).length>=3,`${e.id} has at least three steps`);
    for(const part of [...ACTIVITY_PARTS.map(p=>p.id),'notes'])for(const l of lines(e[part]))assert.ok(l.length<=600,`${e.id} ${part} line length`);
    assert.ok(!/[—–]/.test(JSON.stringify(e)),`${e.id} has no dashes standing in for words`);
  }
  assert.ok(STARTER_ACTIVITIES.some(e=>e.caseStudy),'one of them uses a case study');
  assert.ok(STARTER_ACTIVITIES.some(e=>e.materials&&e.preparation),'and some say what to have ready');
});

test('an entry becomes an Activity written as labelled parts, with its minutes and notes',()=>{
  const e=STARTER_ACTIVITIES.find(x=>x.id==='think-pair-share'),b=blockFromEntry(e);
  assert.equal(b.cue,'activity');assert.equal(b.text,'Think, pair, share');assert.equal(b.minutes,'10');assert.equal(b.audience,'everyone');
  assert.deepEqual(Object.keys(b.parts),ACTIVITY_PARTS.map(p=>p.id));
  assert.equal(b.parts.instructions,e.instructions);assert.equal(b.notes,e.notes);
});

test('the Activity Library opens on its own panel, with the ready-made activities listed and described',async()=>{
  const {a}=await withLibrary();
  assert.equal(a.n('guide-sec-library').hidden,false);
  assert.equal(a.n('guide-sec-module').hidden,true);assert.equal(a.n('guide-sec-block').hidden,true);
  assert.deepEqual(itemTitles(a,'guide-lib-list'),STARTER_ACTIVITIES.map(e=>`${e.title} (${e.minutes} min)${e.kind}`));
  assert.equal(a.n('guide-lib-count').textContent,'8 activities ready to use.');
  assert.match(a.n('guide-lib-mine-note').textContent,/You have not saved any yet/);
  const li=a.n('guide-lib-list').children[1];
  assert.equal(li.children[1].textContent,STARTER_ACTIVITIES[1].summary);
  const inside=li.children[2].children[1].textContent;
  assert.match(inside,/Objective: Learners form an answer on their own/);
  assert.match(inside,/Materials \(facilitator only\): The question on a slide/);
  assert.match(inside,/Instructions:/);
  assert.equal(a.err(),'');
});

test('every ready-made activity goes into the guide and reads back as the same parts, minutes and notes',async()=>{
  const {a}=await withLibrary();
  for(const e of STARTER_ACTIVITIES){
    const li=a.n('guide-lib-list').children.find(x=>x.children[0].textContent.startsWith(e.title));
    await buttonOf(li,'Insert').onclick();
    assert.equal(a.err(),'',e.id);
    assert.ok(a.n('guide-status').textContent.startsWith(`${e.title} was added after the block your cursor was in`),a.n('guide-status').textContent);
    const {paras,i}=await activityAt(a);
    assert.ok(i>=0,`${e.id} is in the guide`);
    const b=readBlock(paras,i);
    assert.equal(b.cue,'activity');assert.equal(b.text,e.title);assert.equal(b.minutes,String(e.minutes));
    assert.ok(b.parts,`${e.id} reads back as labelled parts`);
    for(const part of ACTIVITY_PARTS.map(p=>p.id))assert.equal(b.parts[part]??'',lines(e[part]).join('\n'),`${e.id} ${part}`);
    assert.equal(b.notes,lines(e.notes).join('\n'));
    // clear the way for the next one: take the activity out again
    await a.adapter.deleteRange({start:b.span.start,end:b.span.end,guard:paras[b.span.start].text});await a.p.refresh();
  }
});

test('a second activity goes after the first, and the guide says how to take it back',async()=>{
  const {a}=await withLibrary();
  await buttonOf(a.n('guide-lib-list').children[1],'Insert').onclick();
  await buttonOf(a.n('guide-lib-list').children[3],'Insert').onclick();
  const rows=a.rows(),first=rows.findIndex(r=>/Think, pair, share/.test(r)),second=rows.findIndex(r=>/Role play in threes/.test(r));
  assert.ok(first>0&&second>first,'in the order they were put in');
  assert.match(a.n('guide-status').textContent,/choose Edit, then Undo/);
});

test('finding an activity narrows both lists by title, kind or summary',async()=>{
  const {a}=await withLibrary();
  a.type('guide-lib-find','role');a.n('guide-lib-find').oninput();
  assert.deepEqual(itemTitles(a,'guide-lib-list').map(t=>t.replace(/ \(.*/,'')),['Role play in threes']);
  assert.equal(a.n('guide-lib-count').textContent,'1 activity found.');
  a.type('guide-lib-find','check');a.n('guide-lib-find').oninput();
  assert.deepEqual(itemTitles(a,'guide-lib-list').map(t=>t.replace(/ \(.*/,'')),['Pulse check','Teach-back'],'by kind');
  a.type('guide-lib-find','zzz');a.n('guide-lib-find').oninput();
  assert.equal(a.n('guide-lib-count').textContent,'0 activities found.');assert.equal(a.n('guide-lib-list').children.length,0);
  a.type('guide-lib-find','');a.n('guide-lib-find').oninput();
  assert.equal(a.n('guide-lib-list').children.length,8);
  assert.ok(matches({title:'A',kind:'B',summary:'C d'},'c D'));assert.ok(!matches({title:'A',kind:'B',summary:'C'},'x'));
});

test('an activity the author wrote can be saved, is kept on this computer, and is there in the next pane',async()=>{
  const storage=memoryStorage();
  const {a}=await withLibrary(storage);
  a.n('guide-lib-find').value='';
  await buttonOf(a.n('guide-lib-list').children[2],'Insert').onclick();
  const {i}=await activityAt(a);
  a.adapter.place(i);await a.p.quietRefresh();
  await click(a,'guide-lib-save');
  assert.equal(a.err(),'');
  assert.match(a.n('guide-status').textContent,/Small-group case discussion was saved to your activities\./);
  assert.deepEqual(itemTitles(a,'guide-lib-mine').map(t=>t.replace(/ \(.*/,'')),['Small-group case discussion']);
  const kept=JSON.parse(storage.data[STORE_KEY]);
  assert.equal(kept.length,1);assert.equal(kept[0].caseStudy.includes('Dana leads a team of six'),true);
  // saving it again updates it rather than adding a second
  await click(a,'guide-lib-save');
  assert.match(a.n('guide-status').textContent,/was updated in your activities/);
  assert.equal(JSON.parse(storage.data[STORE_KEY]).length,1);
  // the next pane (a new document, a new day) finds it
  const b=await started('Another guide',{storage});
  await addModule(b,'M');await addLesson(b,'L');
  assert.equal(b.n('guide-lib-mine').children.length,1);
  assert.match(b.n('guide-lib-mine-note').textContent,/kept on this computer/);
  // and puts it in
  await buttonOf(b.n('guide-lib-mine').children[0],'Insert').onclick();
  assert.equal(b.err(),'');assert.ok(b.rows().some(r=>/Small-group case discussion \(20 min\)/.test(r)));
});

test('an activity written as plain lines is saved with its first line as the title and the rest as the steps',async()=>{
  const storage=memoryStorage();
  const {a}=await withLibrary(storage);
  a.pick('activity');a.type('guide-block-text','Warm-up\nTalk to the person next to you.\nSay one thing you noticed.');a.type('guide-block-minutes','5');
  await a.n('guide-block-form').onsubmit({preventDefault(){}});
  assert.equal(a.err(),'');
  const {i}=await activityAt(a);a.adapter.place(i);await a.p.quietRefresh();
  await click(a,'guide-lib-save');
  assert.equal(a.err(),'');
  const kept=JSON.parse(storage.data[STORE_KEY])[0];
  assert.equal(kept.title,'Warm-up');assert.equal(kept.minutes,5);assert.equal(kept.instructions,'Talk to the person next to you.\nSay one thing you noticed.');
});

test('saving says what is wrong: not in an Activity, nothing to save, or storage refused',async()=>{
  const {a}=await withLibrary();
  await click(a,'guide-lib-save');
  assert.match(a.err(),/Put your cursor in an Activity block, then choose Save\./);
  a.pick('activity');a.type('guide-block-text','Only a title');
  await a.n('guide-block-form').onsubmit({preventDefault(){}});
  const {i}=await activityAt(a);a.adapter.place(i);await a.p.quietRefresh();
  await click(a,'guide-lib-save');
  assert.match(a.err(),/This activity has no steps to save/);
  // storage that refuses
  const refusing={getItem:()=>null,setItem:()=>{throw new Error('full');}};
  const b=await withLibrary(refusing);
  await buttonOf(b.a.n('guide-lib-list').children[0],'Insert').onclick();
  const at=await activityAt(b.a);b.a.adapter.place(at.i);await b.a.p.quietRefresh();
  await click(b.a,'guide-lib-save');
  assert.match(b.a.err(),/would not keep your activity.*Nothing was saved/);
  assert.equal(b.a.n('guide-lib-mine').children.length,0);
});

test('removing one of the author\'s own takes it out of the list and of storage, and leaves the guide alone',async()=>{
  const storage=memoryStorage();
  const {a}=await withLibrary(storage);
  await buttonOf(a.n('guide-lib-list').children[0],'Insert').onclick();
  const {i}=await activityAt(a);a.adapter.place(i);await a.p.quietRefresh();
  await click(a,'guide-lib-save');
  const rowsBefore=a.rows().length;
  await buttonOf(a.n('guide-lib-mine').children[0],'Remove').onclick();
  assert.equal(a.err(),'');
  assert.equal(a.n('guide-lib-mine').children.length,0);assert.deepEqual(JSON.parse(storage.data[STORE_KEY]),[]);
  assert.match(a.n('guide-status').textContent,/Opening round was removed from your activities\. The copy already in a guide is not touched\./);
  assert.equal(a.rows().length,rowsBefore);
});

test('what is read back from storage keeps only whole entries, and the list holds fifty at most',()=>{
  assert.deepEqual(cleanMine(null),[]);assert.deepEqual(cleanMine('x'),[]);
  const good={id:'a',title:'  Warm   up ',minutes:'7',instructions:'Do it.',objective:'Obj'};
  const out=cleanMine([good,{...good},{id:'b',title:'',instructions:'x'},{id:'c',title:'T',instructions:''},{title:'No id',instructions:'x'},null,'junk',{id:'d',title:'T',instructions:'x',minutes:-3}]);
  assert.deepEqual(out.map(e=>[e.id,e.title,e.minutes]),[['a','Warm up',7],['d','T',0]],'a repeat, a blank title, no steps, no id and junk are left out');
  assert.equal(out[0].kind,'Yours');
  let mine=[];
  for(let i=0;i<MAX_MINE;i++)mine=addMine(mine,{title:`Activity ${i}`,minutes:5,instructions:'Do it.',objective:'',materials:'',preparation:'',caseStudy:'',notes:''},{now:()=>1000+i}).list;
  assert.equal(mine.length,MAX_MINE);
  assert.throws(()=>addMine(mine,{title:'One too many',instructions:'x'}),/most.*Remove one first/);
  const again=addMine(mine,{title:'activity 3',minutes:9,instructions:'New steps.'});
  assert.equal(again.replaced,true);assert.equal(again.list.length,MAX_MINE);assert.equal(again.list[0].instructions,'New steps.','an update goes first');
  assert.equal(removeMine(mine,mine[0].id).length,MAX_MINE-1);
});

test('a block with a picture cannot be saved, and the reason is given',()=>{
  assert.throws(()=>entryFromBlock({cue:'activity',text:'T\nstep',hasPicture:true}),/has a picture/);
  assert.throws(()=>entryFromBlock({cue:'say',text:'T'}),/Put your cursor in an Activity block/);
  assert.throws(()=>entryFromBlock(null),/Put your cursor in an Activity block/);
  assert.throws(()=>entryFromBlock({cue:'activity',text:'T',parts:{instructions:'x'.repeat(4001)},notes:''}),/longer than 4000/);
});
