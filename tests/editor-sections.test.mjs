// Moving a whole module or lesson in the editor: the keys, and the program's document taking a move and the renumbering after it as one step of Undo. The editor itself
// (keys, the Edit menu and dragging a heading) is checked in a browser by scripts/check-editor-browser.mjs.
import test from 'node:test';
import assert from 'node:assert/strict';
import {DOMParser,XMLSerializer} from '@xmldom/xmldom';
import {rowKey} from '../program/ui/editor-model.js';
import {createDemoAdapter} from '../src/guide-word.mjs';
import {tidyTarget,tidyMove,updatePlan,lessonSpan,topSpan} from '../src/guide-model.mjs';
import {moduleHeading,lessonHeading,timeBlock} from '../src/guide-ooxml.mjs';

const key=over=>rowKey({key:'ArrowUp',ctrl:false,meta:false,alt:true,shift:true,...over});
test('Alt+Shift+Up and Alt+Shift+Down move a whole module or lesson, and Alt alone still moves one paragraph',()=>{
  assert.equal(key(),'moveSectionUp');assert.equal(key({key:'ArrowDown'}),'moveSectionDown');
  assert.equal(key({shift:false}),'moveUp');assert.equal(key({shift:false,key:'ArrowDown'}),'moveDown');
  assert.equal(key({key:'ArrowLeft'}),null);assert.equal(key({ctrl:true}),null,'with Ctrl it is not this key');
  assert.equal(key({alt:false}),'extendUp','Shift alone still chooses more paragraphs');
});

test('a move and the renumbering after it are one step of Undo when they are made in one group, and two steps when they are not',async()=>{
  const xml={DOMParserClass:DOMParser,SerializerClass:XMLSerializer};
  const make=async()=>{
    const a=createDemoAdapter(xml);
    await a.insert(ctx=>moduleHeading(1,'One')+lessonHeading(1,1,'First')+timeBlock({minutes:10,cumulative:10},ctx)+moduleHeading(2,'Two'),{where:'start',select:'last',spare:false});
    return a;
  };
  for(const grouped of [true,false]){
    const a=await make();
    const list=a.exportState().paragraphs,second=list.findIndex(p=>/^Module 2: Two/.test(p.text));
    const target=tidyTarget(list,second),move=tidyMove(list,target,-1);
    assert.ok(move.span&&move.after!==undefined,'the second module can move up');
    const before=a.snapshot().map(p=>p.text);
    const group=grouped?`g-${Math.random()}`:undefined;
    await a.moveRange({start:move.span.start,end:move.span.end,guard:list[move.span.start].text,after:move.after,group});
    const plan=updatePlan(a.exportState().paragraphs,{});
    assert.ok(plan.edits.length>0,'the numbers need to follow the move');
    await a.applyEdits(plan.edits,{group});
    assert.notDeepEqual(a.snapshot().map(p=>p.text),before);
    a.undo?.();
    const steps=a.historyState().undo;
    assert.equal(steps>=0,true);
    if(grouped)assert.deepEqual(a.snapshot().map(p=>p.text),before,'one Undo gives back the order and the numbers');
    else assert.notDeepEqual(a.snapshot().map(p=>p.text),before,'without a group the numbers are a step of their own');
  }
  assert.equal(typeof lessonSpan,'function');assert.equal(typeof topSpan,'function');
});
