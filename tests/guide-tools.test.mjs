import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import {GROUPS,GUIDE_SECTIONS,TOOLS,TOOL_IDS,toolFor,isTool,groupOf,relatedTools} from '../src/guide-tools.mjs';

const root=new URL('..',import.meta.url);
const PARTS=['add','slides','elearning','look','organize'];
const LEGACY_VIEWS=['library','edit-activity','agenda','pack','checkpoints'];
const SCREENS=['assess','workbook','review','help'];

test('there are seven groups and twenty-one tools, each tool in exactly one group, in the planned order',()=>{
  assert.deepEqual(GROUPS.map(g=>g.label),['Start','Write','Bring In','E-Learning & Assessment','Check & Plan','Make','Protect']);
  assert.equal(TOOL_IDS.length,21);
  const listed=GROUPS.flatMap(g=>g.tools);
  assert.equal(listed.length,21);
  assert.deepEqual([...listed].sort(),[...TOOL_IDS].sort(),'every tool is in a group and every listed tool exists');
  assert.equal(new Set(listed).size,21,'no tool is in two groups');
  for(const id of TOOL_IDS)assert.equal(groupOf(id).id,TOOLS[id].group,`${id} names the group it is in`);
});

test('no group has more than three buttons: Word draws a bigger group with mixed icon sizes (seen 2026-11-07)',()=>{
  for(const g of GROUPS)assert.equal(g.tools.length,3,`${g.label} has exactly three, so every group is built alike`);
});

test('the nine buttons the ribbon has always had keep their function names',()=>{
  for(const id of ['openDesigner','addModule','addActivity','participantWorkbook','courseOutline','timingAndAgenda','reviewGuide','saveBackup','restoreBackup']){
    assert.ok(isTool(id),id);
    assert.equal(TOOLS[id].command,id,'an older manifest still names the function by this word');
  }
});

test('every tool has a label, a title, a note, a tip, an icon and a way to open its panel',()=>{
  for(const id of TOOL_IDS){
    const t=toolFor(id);
    for(const key of ['label','title','note','tip','icon','group','kind'])assert.ok(t[key]&&String(t[key]).trim(),`${id} has ${key}`);
    assert.match(t.command,/^[a-z][A-Za-z]+$/,`${id} command is a plain function name`);
    if(t.kind==='guide'){
      assert.ok(PARTS.includes(t.part),`${id} part`);
      for(const s of t.sections)assert.ok(GUIDE_SECTIONS.includes(s),`${id} section ${s}`);
    }else if(t.kind==='screen')assert.ok(SCREENS.includes(t.screen),`${id} screen`);
    else if(t.kind==='legacy')assert.ok(LEGACY_VIEWS.includes(t.view),`${id} older panel`);
    else assert.fail(`${id} has an unknown kind ${t.kind}`);
  }
});

test('More tools lists a few tools that belong with the open one, never all of them and never itself',()=>{
  for(const id of TOOL_IDS){
    const more=relatedTools(id);
    assert.ok(more.length>=1&&more.length<=5,`${id} offers ${more.length}`);
    assert.ok(!more.some(t=>t.id===id),`${id} does not list itself`);
    assert.equal(new Set(more.map(t=>t.id)).size,more.length,`${id} lists no tool twice`);
    for(const r of TOOLS[id].related)assert.ok(isTool(r),`${id} -> ${r} exists`);
  }
  // Working on a module, the other module-level and writing tools are close by.
  assert.ok(relatedTools('addModule').some(t=>t.id==='addActivity'));
  assert.ok(relatedTools('addActivity').some(t=>t.id==='addModule'),'the module part stays reachable from the writing tools');
  assert.ok(relatedTools('assessments').some(t=>t.id==='learningAlignment'));
  assert.ok(!relatedTools('assessments').some(t=>t.id==='addModule'),'Assessments does not offer the writing tools');
});

test('the tools that share a part of the Guide screen show different pieces of it',()=>{
  const seen=new Map();
  for(const t of Object.values(TOOLS).filter(t=>t.kind==='guide'))for(const s of t.sections){
    // A section belongs to the tools that show it; the two writing tools share the block form on purpose.
    seen.set(s,[...(seen.get(s)||[]),t.id]);
  }
  assert.deepEqual(seen.get('guide-sec-module'),['addModule']);
  assert.deepEqual(seen.get('guide-sec-block').sort(),['addActivity','instructionalCues']);
  assert.deepEqual(seen.get('guide-sec-look'),['appearance']);
  assert.deepEqual(seen.get('guide-sec-built'),['generateMaterials']);
  assert.deepEqual(seen.get('guide-sec-standard'),['courseSettings']);
});

test('the labels fit a crowded ribbon: none is longer than 20 characters',()=>{
  for(const id of TOOL_IDS)assert.ok(TOOLS[id].label.length<=20,`${id}: ${TOOLS[id].label}`);
});

test('every tool can be reached from Open Designer through More tools, for the author who hides the ribbon',()=>{
  const seen=new Set(['openDesigner']),queue=['openDesigner'];
  while(queue.length)for(const r of TOOLS[queue.shift()].related)if(!seen.has(r)){seen.add(r);queue.push(r);}
  assert.deepEqual(TOOL_IDS.filter(id=>!seen.has(id)),[],'tools no list leads to');
});
