import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import {CUES,CUE_GROUPS,AUDIENCES,ICON_FILES,iconUrl,allCues,cueById,cuesInGroup,setCustomCues,customCueNames,checkCustomCueNames,CUSTOM_CUE_SLOTS} from '../src/guide-cues.mjs';
import {makeContext,cueBlock,buildPackage} from '../src/guide-ooxml.mjs';
import {kindOf,cueFromText} from '../src/guide-model.mjs';
import {paragraphsOf,TINY} from './helpers/guide-package.mjs';
import {started,click,submit,guideRows,addModule,addLesson} from './helpers/guide-harness.mjs';

const iconData=Object.fromEntries(ICON_FILES.map(i=>[i,TINY]));
const picture={base64:TINY,type:'png',mime:'image/png',size:{cx:914400,cy:914400},alt:'A test picture',decorative:false,name:'t.png'};
// A Picture block needs a picture; every other cue is written with words alone.
const written=(cue,extra={})=>{const ctx=makeContext();return paragraphsOf(buildPackage(cueBlock({cue,text:'Hello there',minutes:5,...(cue==='picture'?{picture}:{}),...extra},ctx),ctx,iconData));};

test('the catalogue has about seventy cues, each complete and in a known group',async()=>{
  assert.ok(CUES.length>=58&&CUES.length<=70,`${CUES.length} cues`);
  assert.equal(new Set(CUES.map(c=>c.id)).size,CUES.length,'ids are unique');
  assert.equal(new Set(CUES.map(c=>c.label.toLowerCase())).size,CUES.length,'labels are unique');
  const groups=new Set(CUE_GROUPS.map(g=>g.id)),audiences=new Set(AUDIENCES.map(a=>a.id));
  for(const c of CUES){
    assert.ok(groups.has(c.group),`${c.id} has a group`);
    assert.ok(audiences.has(c.audience),`${c.id} has an audience`);
    assert.ok(c.label.trim()&&c.help.trim(),`${c.id} has a label and help`);
    assert.match(c.label,/^[\p{L}\p{N}]/u,`${c.id}'s label starts with a letter`);
    assert.doesNotMatch(c.label,/[:()]/,`${c.id}'s label has no colon or bracket`);
    const bytes=await fs.readFile(new URL(`../public${iconUrl(c.icon)}`,import.meta.url));
    assert.equal(bytes.subarray(1,4).toString(),'PNG',`${c.id}'s picture exists`);
  }
  // The rows are written in group order, which is the order of the Graphic Cues page.
  const order=CUE_GROUPS.map(g=>g.id),seen=CUES.map(c=>order.indexOf(c.group));
  assert.deepEqual(seen,[...seen].sort((a,b)=>a-b),'rows are in group order');
  for(const g of CUE_GROUPS)assert.ok(cuesInGroup(g.id).length>0||g.id==='you','every group has cues');
});

test('the groups hold the cues the plan named',()=>{
  const ids=g=>cuesInGroup(g).map(c=>c.id);
  assert.deepEqual(ids('virtual'),['chat','breakout','polling','vote','raise-hand','microphone','webcam','website','check-computer','presenter-1','presenter-2','producer','qa']);
  for(const id of ['highlight','important','process','steps','tools'])assert.ok(ids('teaching').includes(id),id);
  for(const id of ['do-this','lab','role-play','game','case-study','break'])assert.ok(ids('activities').includes(id),id);
  for(const id of ['slide-notes','picture','video','audio','flipchart','data-tables'])assert.ok(ids('slides-media').includes(id),id);
  for(const id of ['files','evaluation','capture'])assert.ok(ids('materials').includes(id),id);
});

test('virtual-class, presenter and producer cues default to facilitator only; handouts and activities to everyone',()=>{
  for(const c of cuesInGroup('virtual'))if(c.id!=='qa')assert.equal(c.audience,'facilitator',c.id);
  for(const id of ['handout','workbook','files','case-study','activity','do-this','evaluation','picture'])assert.equal(cueById(id).audience,'everyone',id);
  for(const id of ['slide-notes','flipchart','trainer-notes','transition','say'])assert.equal(cueById(id).audience,'facilitator',id);
});

test('every cue reads back as itself from the cue line it writes, in the right style',()=>{
  for(const c of CUES){
    const paras=written(c.id),cue=paras.find(p=>kindOf(p)==='cue');
    assert.ok(cue,`${c.id} writes a cue line`);
    assert.equal(cueFromText(cue.text)?.id,c.id,`${c.id} reads back as ${c.id}`);
    assert.equal(cue.style,c.audience==='facilitator'?'OSG Cue Facilitator':'OSG Cue',`${c.id} uses the style for its audience`);
  }
});

test('a label that begins another label is not mistaken for it',()=>{
  assert.equal(cueFromText('Slide notes\tOpen with the second click.').id,'slide-notes');
  assert.equal(cueFromText('Slide\tShow slide 3.').id,'slide');
  assert.equal(cueFromText('Breakout\tPairs, ten minutes.').id,'breakout');
  assert.equal(cueFromText('Break\tTen minutes.').id,'break');
  assert.equal(cueFromText('Time on task\tWrite quietly.').id,'time-on-task');
});

test('your own cues: names are checked, kept apart from the built-in ones, and used like any cue',()=>{
  try{
    assert.deepEqual(checkCustomCueNames(['Recap talk','','Wrap up']).problems,[]);
    const bad=checkCustomCueNames(['Say','Say more','x'.repeat(30)]);
    assert.equal(bad.ok,false);
    assert.equal(bad.problems.length,3,bad.problems.join(' | '));
    assert.match(checkCustomCueNames(['Note: later']).problems[0],/colon or round brackets/);
    assert.match(checkCustomCueNames(['(x)']).problems[0],/start with a letter/);
    assert.match(checkCustomCueNames(['Recap talk','recap talk']).problems[0],/too like "Recap talk"/);
    assert.match(checkCustomCueNames(['Slide notes extra']).problems[0],/too like "Slide/);

    setCustomCues(['Recap talk','','Wrap up']);
    assert.deepEqual(customCueNames(),['Recap talk','','Wrap up']);
    assert.equal(allCues().length,CUES.length+2);
    assert.equal(cueById('custom-1').label,'Recap talk');
    assert.equal(cueById('custom-2'),null,'an empty slot is not a cue');
    assert.equal(cueById('custom-3').group,'you');
    const cue=written('custom-3').find(p=>kindOf(p)==='cue');
    assert.equal(cueFromText(cue.text).id,'custom-3');
    setCustomCues(['Say more']);
    assert.equal(allCues().length,CUES.length,'a name that is too like another is left out, not half-used');
  }finally{setCustomCues([]);}
  assert.equal(allCues().length,CUES.length);
});

test('the pane offers the kinds of cue, filters by kind and by name, and still counts matches',async()=>{
  const a=await started();
  const labels=()=>a.n('guide-cues').children;
  assert.equal(labels().length,CUES.length);
  const options=a.n('guide-cue-group').children.map(o=>o.value);
  assert.deepEqual(options,['',...CUE_GROUPS.map(g=>g.id)]);
  a.n('guide-cue-group').value='virtual';a.n('guide-cue-group').oninput();
  assert.equal(labels().filter(l=>!l.hidden).length,13);
  assert.equal(a.n('guide-cue-count').textContent,'13 cues match.');
  a.type('guide-cue-find','poll');a.n('guide-cue-find').oninput();
  assert.deepEqual(labels().filter(l=>!l.hidden).map(l=>l.dataset.label),['polling']);
  assert.equal(a.n('guide-cue-count').textContent,'1 cue match.');
  a.type('guide-cue-find','');a.n('guide-cue-find').oninput();a.n('guide-cue-group').value='';a.n('guide-cue-group').oninput();
  assert.equal(labels().filter(l=>!l.hidden).length,CUES.length);
  assert.equal(a.n('guide-cue-count').textContent,'');
});

test('a new cue can be inserted from the pane and is read back as that cue',async()=>{
  const a=await started();
  await addModule(a,'Module one');await addLesson(a,'Lesson one');
  a.pick('polling');a.type('guide-block-text','Launch the poll: Which tool do you use?');a.type('guide-block-notes','');a.type('guide-block-minutes','3');
  await submit(a,'guide-block-form');
  assert.equal(a.err(),'');
  assert.ok(guideRows(a).some(r=>/^OSG Cue Facilitator: .*Polling.*Launch the poll/.test(r)),guideRows(a).join(' / '));
});

test('your own cues are saved in the document, offered in the list, and cannot be renamed from under blocks that use them',async()=>{
  try{
    const a=await started();
    await addModule(a,'Module one');await addLesson(a,'Lesson one');
    a.type('guide-own-1','Recap talk');a.type('guide-own-2','');a.type('guide-own-3','');
    await click(a,'guide-own-save');
    assert.equal(a.err(),'');
    assert.deepEqual(await a.adapter.customCues(),['Recap talk','','']);
    assert.ok(a.radio('custom-1'),'the own cue is in the list');
    assert.match(a.n('guide-own-note').textContent,/Saved\. Your cue is in the list under "For you"/);
    a.pick('custom-1');a.type('guide-block-text','What went well?');a.type('guide-block-notes','');a.type('guide-block-minutes','');
    await submit(a,'guide-block-form');
    assert.equal(a.err(),'');
    // The block is in use, so its cue's name cannot change or go.
    a.type('guide-own-1','Wrap up');
    await click(a,'guide-own-save');
    assert.match(a.err(),/Blocks in your guide use the cue "Recap talk"/);
    assert.equal(a.n('guide-own-note').textContent,'','the earlier "Saved." line does not stay beside the refusal');
    assert.deepEqual(await a.adapter.customCues(),['Recap talk','',''],'nothing was changed');
    // A bad name is refused with the reason, and nothing is saved.
    a.type('guide-own-1','Recap talk');a.type('guide-own-2','Say');
    await click(a,'guide-own-save');
    assert.match(a.err(),/Custom cue 2 \("Say"\) is too like "Say"/);
    // A second pane on the same document reads the names back when it loads.
    a.type('guide-own-2','');
    await a.adapter.setCustomCues(['Recap talk','Summary','']);
    await a.p.refresh();
    assert.ok(a.radio('custom-2'),'a name added elsewhere appears after a refresh');
    assert.equal(a.n('guide-own-2').value,'Summary');
  }finally{setCustomCues([]);}
});
