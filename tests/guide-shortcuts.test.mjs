// What the navigation actions do: next and previous module or lesson, Contents, Training at a Glance, run against a stand-in document.
// (The add-in's manifest and Word shortcut file tests stay in the add-in's repository.)
import test from 'node:test';
import assert from 'node:assert/strict';
import {DOMParser,XMLSerializer} from '@xmldom/xmldom';
import {NAVIGATION_ACTIONS,navigationTarget} from '../src/guide-model.mjs';
import {ROUTES,registerNavigationActions} from '../src/ribbon.mjs';
import {createDemoAdapter} from '../src/guide-word.mjs';

const P=(text,style='Normal')=>({text,style});

// ---- where each navigation shortcut goes
const GUIDE=[
  P('Title','OSG Guide Title'),P('Contents','Heading 1'),P('Update this table.'),
  P('Getting Started','Heading 1'),P('Training at a Glance','Heading 2'),P('Overview text.'),
  P('Module 1: One','OSG Module'),P('Lesson 1.1: A','OSG Lesson'),P('Time: 5 minutes. Cumulative time: 0 minutes.','OSG Time'),P('Lesson 1.2: B','OSG Lesson'),P('Time: 5 minutes. Cumulative time: 5 minutes.','OSG Time'),
  P('Module 2: Two','OSG Module'),P('Lesson 2.1: C','OSG Lesson'),P('Time: 5 minutes. Cumulative time: 10 minutes.','OSG Time')
];
const go=(at,kind)=>navigationTarget(GUIDE,at,kind);
test('next and previous go to the nearest module or lesson heading from the cursor, wherever it is',()=>{
  assert.deepEqual(go(0,'next-module'),{index:6,text:'Module 1: One'});
  assert.deepEqual(go(8,'next-module'),{index:11,text:'Module 2: Two'},'from inside lesson 1.1');
  assert.deepEqual(go(8,'previous-module'),{index:6,text:'Module 1: One'},'from inside module 1, previous is its own heading');
  assert.deepEqual(go(6,'previous-lesson'),{message:'There is no previous lesson in this guide.'});
  assert.deepEqual(go(7,'next-lesson'),{index:9,text:'Lesson 1.2: B'});
  assert.deepEqual(go(9,'previous-lesson'),{index:7,text:'Lesson 1.1: A'});
  assert.deepEqual(go(10,'next-lesson'),{index:12,text:'Lesson 2.1: C'},'across a module boundary');
});

test('past the ends the answer is a sentence, not a move',()=>{
  assert.deepEqual(go(13,'next-module'),{message:'There is no next module in this guide.'});
  assert.deepEqual(go(13,'next-lesson'),{message:'There is no next lesson in this guide.'});
  assert.deepEqual(go(0,'previous-module'),{message:'There is no previous module in this guide.'});
  assert.deepEqual(navigationTarget([P('Just text.')],0,'next-module'),{message:'There is no next module in this guide.'});
});

test('Contents and Training at a Glance go to their headings, and say what is missing when they are not there',()=>{
  assert.deepEqual(go(12,'contents'),{index:1,text:'Contents'});
  assert.deepEqual(go(12,'glance'),{index:4,text:'Training at a Glance'});
  assert.match(navigationTarget([P('Module 1: A','OSG Module')],0,'contents').message,/no Contents page/);
  assert.match(navigationTarget([P('Module 1: A','OSG Module')],0,'glance').message,/Choose Build Training at a Glance/);
  assert.match(go(0,'nonsense').message,/not a place this can go to/);
});

// ---- the actions, run against a document
function fixture(paras,cursor){
  const adapter=createDemoAdapter({DOMParserClass:DOMParser,SerializerClass:XMLSerializer});
  adapter.seed(paras);adapter.place(cursor);
  const handlers={},said=[],errors=[];
  registerNavigationActions({actions:{associate:(id,fn)=>handlers[id]=fn}},{adapter,say:t=>said.push(t),onError:e=>errors.push(e)});
  const press=async id=>{let done=0;await handlers[id]({completed(){done++;}});assert.equal(done,1,`${id} releases the Office command`);};
  return {adapter,handlers,said,errors,press};
}

test('each shortcut moves the cursor in the document, says where it went, and releases the command',async()=>{
  const f=fixture(GUIDE,0);
  assert.deepEqual(Object.keys(f.handlers).sort(),Object.keys(NAVIGATION_ACTIONS).sort());
  await f.press('nextModule');assert.equal(f.adapter.cursor(),6);
  await f.press('nextLesson');assert.equal(f.adapter.cursor(),7);
  await f.press('nextLesson');assert.equal(f.adapter.cursor(),9);
  await f.press('previousLesson');assert.equal(f.adapter.cursor(),7);
  await f.press('previousModule');assert.equal(f.adapter.cursor(),6);
  await f.press('goToGlance');assert.equal(f.adapter.cursor(),4);
  await f.press('goToContents');assert.equal(f.adapter.cursor(),1);
  assert.deepEqual(f.said.slice(0,2),['Moved to Module 1: One.','Moved to Lesson 1.1: A.']);
  assert.deepEqual(f.errors,[]);
});

test('where there is nowhere to go the cursor stays and the pane says why',async()=>{
  const f=fixture(GUIDE,13);
  await f.press('nextModule');
  assert.equal(f.adapter.cursor(),13);
  assert.deepEqual(f.said,['There is no next module in this guide.']);
});

test('a failure is reported through the app\'s error line and still releases the command',async()=>{
  const adapter={paragraphs:async()=>{throw new Error('Word is not ready yet.');},cursorIndex:async()=>0,go:async()=>{}};
  const handlers={},errors=[];
  registerNavigationActions({actions:{associate:(id,fn)=>handlers[id]=fn}},{adapter,say:()=>{},onError:e=>errors.push(e)});
  let done=0;await handlers.nextLesson({completed(){done++;}});
  assert.equal(done,1);assert.match(errors[0].message,/Word is not ready yet/);
});

test('a cold shortcut waits for the app to finish starting before it reads the document',async()=>{
  let release;const ready=new Promise(r=>release=r),calls=[];
  const adapter={paragraphs:async()=>{calls.push('read');return GUIDE;},cursorIndex:async()=>0,go:async i=>calls.push(`go ${i}`)};
  const handlers={};
  registerNavigationActions({actions:{associate:(id,fn)=>handlers[id]=fn}},{ready,adapter,say:()=>{},onError:()=>{}});
  const pending=handlers.nextModule({completed(){}});
  await Promise.resolve();assert.deepEqual(calls,[]);
  release();await pending;assert.deepEqual(calls,['read','go 6']);
});
