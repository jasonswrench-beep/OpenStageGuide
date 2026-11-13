// One guide document used from two pages (src/guide-remote.mjs): the frame's tools call it by message, the program's page answers.
import test from 'node:test';
import assert from 'node:assert/strict';
import {DOMParser,XMLSerializer} from '@xmldom/xmldom';
import {createDemoAdapter} from '../src/guide-word.mjs';
import {createProjectAdapter} from '../src/guide-project-adapter.mjs';
import {createRemoteAdapter,serveAdapter,REMOTE_METHODS} from '../src/guide-remote.mjs';
import {moduleHeading,lessonHeading,para,run} from '../src/guide-ooxml.mjs';

// Two ends joined the way two pages are: a message is copied (structured clone) and arrives a moment later, never in the same breath.
function wire(){
  const ends={frame:[],host:[]};
  const send=(to,message)=>{const copy=structuredClone(message);queueMicrotask(()=>ends[to].forEach(fn=>fn(copy)));};
  return {
    frame:{post:m=>send('host',m),listen:fn=>{ends.frame.push(fn);}},
    host:{post:m=>send('frame',m),listen:fn=>{ends.host.push(fn);return()=>{ends.host.length=0;};}},
    raw:{toHost:m=>send('host',m),fromHost:fn=>ends.frame.push(fn)}
  };
}
const demo=()=>createDemoAdapter({DOMParserClass:DOMParser,SerializerClass:XMLSerializer});
async function setup({onChange}={}){
  const saved=[];
  const guide=createProjectAdapter(demo(),{read:async()=>null,write:async t=>{saved.push(t);},delayMs:0,soon:fn=>fn(),cancel:()=>{},onChange});
  const w=wire(),statuses=[],downloads=[];
  const link=serveAdapter(guide,w.host);
  const remote=createRemoteAdapter({...w.frame,onStatus:s=>statuses.push(s),download:(...a)=>downloads.push(a)});
  await remote.ready;
  await new Promise(r=>setTimeout(r,10)); // the tools ask how the saving stands as soon as the document is open
  return {guide,remote,link,w,saved,statuses,downloads};
}
const texts=guide=>guide.snapshot().map(p=>p.text);

test('the tools use the program\'s document through the same interface: insert, read, go, delete',async()=>{
  const {guide,remote}=await setup();
  assert.deepEqual(await remote.paragraphs(),[]);
  await remote.insert(()=>moduleHeading(1,'Naming it')+lessonHeading(1,1,'First'),{where:'start'});
  assert.ok(texts(guide).includes('Module 1: Naming it')&&texts(guide).includes('Lesson 1.1: First'),'the document in the program page changed');
  const seen=await remote.paragraphs();
  assert.equal(seen[0].text,'Module 1: Naming it');
  assert.equal(seen[0].style,'OSG Module');
  await remote.go(0,'Module 1: Naming it');
  assert.equal(await remote.cursorIndex(),0);
  await remote.deleteRange({start:1,end:1,guard:'Lesson 1.1: First'});
  assert.ok(!texts(guide).includes('Lesson 1.1: First'));
});

test('the content is built in the frame and only the finished text is sent',async()=>{
  const w=wire(),sent=[];
  const guide=createProjectAdapter(demo(),{read:async()=>null,write:async()=>{},delayMs:0,soon:fn=>fn(),cancel:()=>{}});
  serveAdapter(guide,{post:w.host.post,listen:fn=>{w.host.listen(m=>{sent.push(m);fn(m);});}});
  const remote=createRemoteAdapter(w.frame);
  await remote.ready;
  let builderRanWith=null;
  await remote.insert(ctx=>{builderRanWith=ctx;return para('OSGCueBody',run('Hello'));},{where:'start',spare:false});
  assert.ok(builderRanWith&&builderRanWith.used instanceof Set,'the builder got a fresh context here');
  const insert=sent.find(m=>m.method==='insert');
  assert.equal(typeof insert.args[0],'string','what travelled was Word-format text');
  assert.match(insert.args[0],/Hello/);
  assert.ok(texts(guide).includes('Hello'));
});

test('an error in the document reaches the tool with the document\'s own words; a failing builder never reaches the document',async()=>{
  const {guide,remote}=await setup();
  await assert.rejects(remote.go(99),/no longer there/);
  await assert.rejects(remote.deleteRange({start:0,end:0,guard:'x'}),/changed since/i);
  const before=texts(guide).length;
  assert.throws(()=>remote.insert(()=>{throw new Error('Describe the picture in a few words.');}),/Describe the picture/);
  assert.equal(texts(guide).length,before);
});

test('only the document\'s interface can be called: nothing that imports, exports, moves the cursor or undoes',async()=>{
  const {guide,w}=await setup();
  const answers=[];
  w.raw.fromHost(m=>answers.push(m));
  const names=['exportState','importState','place','typeOver','undo','fireSelection','seed','snapshot','constructor','__proto__','toString','hasOwnProperty','then','kind'];
  names.forEach((method,i)=>w.raw.toHost({osg:'call',id:100+i,method,args:[]}));
  await new Promise(r=>setTimeout(r,20));
  const results=answers.filter(m=>m.osg==='result');
  assert.equal(results.length,names.length);
  assert.ok(results.every(r=>r.ok===false),'every one is refused');
  assert.ok(results.every(r=>r.message==='The guide does not offer that.'));
  for(const name of ['exportState','importState','place','typeOver','undo','fireSelection','seed','snapshot'])assert.ok(!REMOTE_METHODS.includes(name),name);
  assert.deepEqual(texts(guide),[]);
});

test('messages that are not calls are ignored: no answer, no change',async()=>{
  const {guide,w}=await setup();
  const answers=[];
  w.raw.fromHost(m=>answers.push(m));
  for(const bad of [null,undefined,'insert',42,{},{osg:'call'},{osg:'call',id:'1',method:'paragraphs'},{osg:'call',id:1.5,method:'paragraphs'},{osg:'call',id:2,method:7},{osg:'result',id:3,ok:true},{type:'osg-ready'}])w.raw.toHost(bad);
  await new Promise(r=>setTimeout(r,20));
  assert.deepEqual(answers,[]);
  assert.deepEqual(texts(guide),[]);
});

test('a block that is not text is refused, and the arguments of a call that is not a list are treated as none',async()=>{
  const {guide,w}=await setup();
  const answers=[];
  w.raw.fromHost(m=>answers.push(m));
  w.raw.toHost({osg:'call',id:1,method:'insert',args:[{evil:true},{where:'start'}]});
  w.raw.toHost({osg:'call',id:2,method:'paragraphs',args:'nonsense'});
  await new Promise(r=>setTimeout(r,20));
  assert.equal(answers.find(a=>a.id===1).ok,false);
  assert.match(answers.find(a=>a.id===1).message,/could not be made/);
  assert.deepEqual(answers.find(a=>a.id===2).value,[]);
  assert.deepEqual(texts(guide),[]);
});

test('events: how the saving stands, and where the cursor is, reach the tools',async()=>{
  const {guide,remote,link,statuses}=await setup();
  assert.equal(statuses.at(-1).state,'new','the tools learn how the saving stands as soon as the document is open');
  link.status({state:'saving'});
  link.status({state:'error',message:'The guide could not be saved: disk full.'});
  await new Promise(r=>setTimeout(r,10));
  assert.deepEqual(statuses.slice(-2).map(s=>s.state),['saving','error']);
  assert.equal(statuses.at(-1).message,'The guide could not be saved: disk full.');
  assert.equal(remote.status().state,'error');
  assert.equal(remote.blocked(),true);
  let moved=0;
  remote.onSelectionChanged(()=>{moved++;});
  guide.fireSelection();
  await new Promise(r=>setTimeout(r,10));
  assert.equal(moved,1);
});

test('the document\'s opening is the tools\' ready: a project that cannot be read is explained, and the tools wait for nothing else',async()=>{
  const w=wire();
  const guide=createProjectAdapter(demo(),{read:async()=>{throw new Error('the program answered 500');},write:async()=>{},delayMs:0,soon:fn=>fn(),cancel:()=>{}});
  serveAdapter(guide,w.host);
  const remote=createRemoteAdapter(w.frame);
  await assert.rejects(remote.ready,/the program answered 500/);
  await assert.rejects(remote.insert(()=>para('OSGCueBody',run('x')),{where:'start'}),/could not be read|cannot be changed/);
  assert.deepEqual(texts(guide),[]);
});

test('calls made together are each answered with their own answer',async()=>{
  const {remote}=await setup();
  await remote.insert(()=>moduleHeading(1,'A')+moduleHeading(2,'B'),{where:'start',spare:false});
  const [a,b,c]=await Promise.all([remote.title(),remote.paragraphs(),remote.cursorIndex()]);
  assert.equal(typeof a,'string');
  assert.deepEqual(b.slice(0,2).map(p=>p.text),['Module 1: A','Module 2: B'],'each call got its own answer (the document ends in one plain empty line)');
  assert.equal(typeof c,'number');
});

test('there is no Word window to open a file in: the tools fall back to a real download, which the frame does',async()=>{
  const {remote,downloads}=await setup();
  await assert.rejects(remote.openDocument('AAAA','x.docx'),/no Word window/);
  remote.offerDownload('AAAA','Workbook.docx','application/x');
  assert.deepEqual(downloads,[['AAAA','Workbook.docx','application/x']]);
});

test('saving still happens in the program\'s page: a change made by the tools is saved there, and the page is told of it',async()=>{
  const changes=[];
  const {guide,remote,saved}=await setup({onChange:()=>changes.push(1)});
  await remote.paragraphs();
  assert.equal(changes.length,0,'reading is not a change');
  await remote.insert(()=>moduleHeading(1,'Saved?'),{where:'start'});
  assert.ok(changes.length>=1,'the page that draws the guide is told it changed');
  await guide.flush();
  assert.ok(saved.length>=1&&saved.at(-1).includes('Saved?'));
});

test('a stopped link no longer answers',async()=>{
  const {link,w}=await setup();
  link.stop();
  const answers=[];
  w.raw.fromHost(m=>answers.push(m));
  w.raw.toHost({osg:'call',id:9,method:'paragraphs',args:[]});
  await new Promise(r=>setTimeout(r,20));
  assert.deepEqual(answers,[]);
});
