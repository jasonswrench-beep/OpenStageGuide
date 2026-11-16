import test from 'node:test';
import assert from 'node:assert/strict';
import {DOMParser} from '@xmldom/xmldom';
import {blockChoices,choiceById,buildChoice,blockMinutes,ACTIVITY_PARTS_CHOICE} from '../src/guide-block-choice.mjs';
import {CUES,CUE_GROUPS} from '../src/guide-cues.mjs';
import {makeContext} from '../src/guide-ooxml.mjs';
import {STYLE_IDS,NS} from '../src/guide-styles.mjs';
import {emptyBlocks,readBlock,guideIssues,updatePlan,blockSpan} from '../src/guide-model.mjs';

const P=(text,style='Normal')=>({text,style});

// The paragraphs a builder makes, as the model reads them: text and style name, in order.
function parasOf(xml){
  const doc=new DOMParser().parseFromString(`<w:body ${NS}>${xml}</w:body>`,'application/xml');
  const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  return Array.from(doc.documentElement.childNodes).filter(n=>n.nodeType===1&&n.localName==='p').map(p=>{
    const st=p.getElementsByTagNameNS(W,'pStyle')[0]?.getAttributeNS(W,'val')||'';
    let text='';
    for(const r of Array.from(p.getElementsByTagNameNS(W,'r'))){
      if(r.getElementsByTagNameNS(W,'drawing').length)text+='\u0001';
      else if(r.getElementsByTagNameNS(W,'tab').length)text+='\t';
      else text+=Array.from(r.getElementsByTagNameNS(W,'t')).map(t=>t.textContent).join('');
    }
    return {text,style:st};
  });
}

test('the chooser lists every kind of block the guide has, in the order of the Graphic Cues page, except the one that needs a picture file', () => {
  const list=blockChoices();
  const ids=list.map(c=>c.id);
  assert.ok(!ids.includes('picture'));
  assert.equal(list.length,CUES.length-1+1,'every cue but Picture, and Activity with its parts');
  assert.deepEqual(ids.filter(id=>id!==ACTIVITY_PARTS_CHOICE),CUES.filter(c=>c.id!=='picture').map(c=>c.id));
  assert.equal(ids[ids.indexOf('activity')+1],ACTIVITY_PARTS_CHOICE,'the activity with its parts follows the activity');
  for(const c of list){assert.ok(c.label&&c.help&&c.groupLabel,c.id);assert.ok(CUE_GROUPS.some(g=>g.id===c.group));}
  assert.equal(choiceById('say').groupLabel,'Teaching');
  assert.equal(choiceById('nope'),null);
});

test('the minutes of a block are nothing, or a whole number from 1 to 999', () => {
  assert.equal(blockMinutes(''),0);
  assert.equal(blockMinutes(undefined),0);
  assert.equal(blockMinutes(' 12 '),12);
  for(const bad of ['0','abc','1.5','-3','1000','1 2'])assert.throws(()=>blockMinutes(bad),/whole number from 1 to 999/,bad);
});

test('a chosen block is built with its icon and its label and no words, with its minutes when they are given', () => {
  const plain=parasOf(buildChoice('say')(makeContext()));
  assert.equal(plain.length,1);
  assert.match(plain[0].text,/^\u0001 Say\t$/);
  const timed=parasOf(buildChoice('break',{minutes:'15'})(makeContext()));
  assert.match(timed[0].text,/^\u0001 Break\t \(15 min\)$|^\u0001 Break\t\(15 min\)$/);
  assert.throws(()=>buildChoice('picture'),/Choose a kind of block/);
  assert.throws(()=>buildChoice('say',{minutes:'x'}),/whole number/);
});

test('an activity with its parts has the labels of its parts in place, the facilitator parts in the notes style', () => {
  const xml=buildChoice(ACTIVITY_PARTS_CHOICE,{minutes:10})(makeContext());
  const paras=parasOf(xml);
  assert.deepEqual(paras.slice(1).map(p=>p.text.trim()),['Objective:','Materials:','Preparation:','Instructions:','Case study:']);
  assert.match(xml,/Activity/);
  const notes=paras.slice(1).map((p,i)=>[p,i]).filter(([p])=>p.style===STYLE_IDS.notes).map(([p])=>p.text.trim());
  assert.deepEqual(notes,['Materials:','Preparation:'],'materials and preparation are in the notes style');
});

// A small guide with blocks in it, written the way the page reads them.
const guide=[
  P('Cover','OSG Guide Title'),
  P('Module 1: One','OSG Module'),
  P('Lesson 1.1: First','OSG Lesson'),
  P('\u0001 Time: 10 minutes. Cumulative time: 0 minutes.','OSG Time'),
  P('\u0001 Say\t','OSG Cue Facilitator'),                                   // 4: nothing
  P('\u0001 Ask\tWhat makes it hard? (4 min)','OSG Cue'),                    // 5: words
  P('\u0001 Break\t(15 min)','OSG Cue'),                                     // 6: only its minutes
  P('\u0001 Discussion\t','OSG Cue'),                                        // 7: nothing, but a line under it
  P('Talk about the case.','OSG Cue Body'),
  P('\u0001 Activity\t','OSG Cue'),                                          // 9: its parts are empty
  P('Objective: ','OSG Cue Body'),
  P('Materials: ','OSG Facilitator Notes'),
  P('Instructions: ','OSG Cue Body'),
  P('\u0001 Key point\t','OSG Cue'),                                         // 13: only a note under it, and the note is empty
  P('Facilitator notes: ','OSG Facilitator Notes'),
  P('\u0001 Debrief\t','OSG Cue Facilitator'),                               // 15: a note with words
  P('Facilitator notes: Ask what changed.','OSG Facilitator Notes'),
];

test('a block with no words is found: nothing after its label, no lines under it, empty parts and empty notes', () => {
  const found=emptyBlocks(guide);
  assert.deepEqual(found.map(b=>b.index),[4,9,13]);
  assert.deepEqual(found.map(b=>b.cue),['say','activity','key-point']);
  assert.equal(found[0].label,'Say');
});

test('a block that is only its minutes, or has a line or a note with words, is not empty', () => {
  const found=emptyBlocks(guide).map(b=>b.index);
  for(const ok of [5,6,7,15])assert.ok(!found.includes(ok),`block at ${ok}`);
});

test('a picture under an empty block means it is not empty', () => {
  const withPicture=[...guide.slice(0,4),P('\u0001 Say\t','OSG Cue Facilitator'),P('￼','OSG Cue Body')];
  assert.deepEqual(emptyBlocks(withPicture),[]);
});

test('the Review says so for each: BLOCK-EMPTY, a warning that names the block and the lesson and points at it', () => {
  const issues=guideIssues(updatePlan(guide)).filter(i=>i.rule==='BLOCK-EMPTY');
  assert.equal(issues.length,3);
  assert.deepEqual(issues.map(i=>i.paragraph),[4,9,13]);
  assert.ok(issues.every(i=>i.level==='warn'));
  assert.match(issues[0].message,/^A Say block in Lesson 1\.1: First has no words yet\./);
});

test('what the author types into the block takes the warning away, and the block reads back as it was typed', () => {
  const typed=guide.map((p,i)=>i===4?P('\u0001 Say\tWelcome everyone.','OSG Cue Facilitator'):p);
  assert.ok(!emptyBlocks(typed).some(b=>b.index===4));
  assert.equal(readBlock(typed,4).text,'Welcome everyone.');
  assert.deepEqual(blockSpan(typed,4),{start:4,end:4});
});
