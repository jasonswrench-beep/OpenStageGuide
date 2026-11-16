import test from 'node:test';
import assert from 'node:assert/strict';
import {DOMParser,XMLSerializer} from '@xmldom/xmldom';
import {createDemoAdapter} from '../src/guide-word.mjs';
import {createProjectAdapter} from '../src/guide-project-adapter.mjs';
import {PROJECT_FORMAT,cleanGuide,serializeProject} from '../src/guide-project.mjs';
import {pane,submit,addModule,addLesson,addBlock} from './helpers/guide-harness.mjs';

const rawDocument=()=>createDemoAdapter({DOMParserClass:DOMParser,SerializerClass:XMLSerializer});

// A clock the test turns by hand, and a file that is a variable.
function rig({text=null,failWrites=0}={}){
  const timers=[],writes=[],statuses=[];
  let stored=text,failing=failWrites;
  const raw=rawDocument();
  const document=createProjectAdapter(raw,{
    read:async()=>stored,
    write:async t=>{if(failing>0){failing--;throw new Error('the disk said no');}writes.push(t);stored=t;},
    soon:fn=>{timers.push(fn);return timers.length-1;},
    cancel:id=>{timers[id]=null;},
    onStatus:s=>statuses.push(s.state)
  });
  const tick=async()=>{const due=timers.splice(0).filter(Boolean);for(const fn of due)fn();await document.flush();};
  return {raw,document,writes,statuses,tick,pending:()=>timers.filter(Boolean).length,file:()=>stored};
}

async function writeAGuide(r){
  const a=pane(undefined,{adapter:r.document});
  await a.p.refresh();
  a.type('guide-title','Reframing conflict');await submit(a,'guide-start-form');
  await addModule(a,'Naming the feeling');
  await addLesson(a,'Say it plainly','10');
  await addBlock(a,'ask','What makes it hard?',{notes:'Listen for blame.',minutes:'4'});
  await addBlock(a,'say','Name the feeling.');
  return a;
}

test('a guide written in the pane is saved to its file and comes back whole in a new session',async()=>{
  const r1=rig();await r1.document.ready;
  const a1=await writeAGuide(r1);
  await r1.document.setCustomCues(['Pause here']);
  await r1.document.setLook('{"name":"calm"}');
  await r1.document.setSlideMemory('{"slides":3}');
  await r1.document.setVersionMemory('{"plans":1}');
  await r1.document.setDesign('{"version":1,"objectives":[]}');
  await r1.tick();
  assert.equal(r1.document.status().state,'saved');
  assert.ok(r1.file(),'a file was written');

  const r2=rig({text:r1.file()});
  assert.deepEqual(await r2.document.ready,{opened:true});
  assert.deepEqual(r2.raw.snapshot(),r1.raw.snapshot());
  assert.equal(await r2.document.title(),'Reframing conflict');
  assert.deepEqual(await r2.document.customCues(),['Pause here']);
  assert.equal(await r2.document.getLook(),'{"name":"calm"}');
  assert.equal(await r2.document.slideMemory(),'{"slides":3}');
  assert.equal(await r2.document.versionMemory(),'{"plans":1}');
  assert.equal(await r2.document.design(),'{"version":1,"objectives":[]}','the notes of the design desk come back with the guide');
  assert.equal(r2.raw.cursor(),r1.raw.cursor());
  // The package the document would give to a Word export or a backup is the same, byte for byte.
  assert.equal(await r2.document.sourcePackage(),await r1.document.sourcePackage());

  // The pane reads the reopened guide as the same guide.
  const a2=pane(undefined,{adapter:r2.document});
  await a2.p.refresh();
  assert.equal(a2.n('guide-build').hidden,false);
  assert.equal(a2.n('guide-start').hidden,true);
  assert.equal(a2.n('guide-build-heading').textContent,'Reframing conflict');
  assert.deepEqual(a2.rows(),a1.rows());
});

test('a guide can be worked on after it is reopened, and the new work is saved too',async()=>{
  const r1=rig();await r1.document.ready;await writeAGuide(r1);await r1.tick();
  const r2=rig({text:r1.file()});await r2.document.ready;
  const a2=pane(undefined,{adapter:r2.document});await a2.p.refresh();
  await addModule(a2,'Hearing the other side');
  await r2.tick();
  assert.ok(a2.rows().some(x=>/Hearing the other side/.test(x)));
  const r3=rig({text:r2.file()});await r3.document.ready;
  assert.deepEqual(r3.raw.snapshot(),r2.raw.snapshot());
  assert.ok(r3.raw.snapshot().some(p=>/Hearing the other side/.test(p.text)));
});

test('several quick changes are one save, made a moment after the last one',async()=>{
  const r=rig();await r.document.ready;
  await r.document.setTitle('One');await r.document.setTitle('Two');await r.document.setTitle('Three');
  assert.equal(r.writes.length,0,'nothing is written while the author is still working');
  assert.equal(r.pending(),1,'only the last timer is waiting');
  assert.equal(r.document.status().state,'unsaved');
  await r.tick();
  assert.equal(r.writes.length,1);
  assert.equal(JSON.parse(r.writes[0]).guide.title,'Three');
  assert.equal(r.document.status().state,'saved');
});

test('flush saves at once, and a change made while a save is running is saved after it',async()=>{
  const r=rig();await r.document.ready;
  await r.document.setTitle('A');
  await r.document.flush();
  assert.equal(r.writes.length,1);
  await r.document.setTitle('B');
  const first=r.document.flush();
  await r.document.setTitle('C');
  await first;await r.tick();
  assert.equal(JSON.parse(r.file()).guide.title,'C');
});

test('looking at the guide, or setting up the session, does not save anything',async()=>{
  const r=rig();await r.document.ready;
  await r.document.paragraphs();await r.document.title();await r.document.cursorIndex();await r.document.reviewParts();await r.document.sourcePackage();
  r.document.setTheme({});r.document.setPicker(()=>[]);r.document.setAutoShow(false);
  assert.equal(r.pending(),0);
  assert.equal(r.document.status().state,'new');
  assert.equal(r.writes.length,0);
});

test('a save that fails says so, keeps the guide marked as not saved, and the next change saves everything',async()=>{
  const r=rig({failWrites:1});await r.document.ready;
  await r.document.setTitle('First');
  await r.tick();
  assert.equal(r.document.status().state,'error');
  assert.match(r.document.status().message,/could not be saved: the disk said no.*still here/);
  assert.equal(r.writes.length,0);
  await r.document.setTitle('Second');
  await r.tick();
  assert.equal(r.document.status().state,'saved');
  assert.equal(JSON.parse(r.file()).guide.title,'Second');
});

test('a file that cannot be read as a project is never written over, and the guide cannot be changed',async()=>{
  const r=rig({text:'this is not a project'});
  await assert.rejects(r.document.ready,/not an OpenStage Guide project|not readable/);
  assert.equal(r.document.status().state,'error');
  assert.match(r.document.status().message,/not changed.*saving is off/);
  assert.equal(r.document.blocked(),true);
  assert.throws(()=>r.document.setTitle('Oops'),/not changed/);
  await r.document.flush();
  assert.equal(r.writes.length,0);
  assert.equal(r.file(),'this is not a project');
});

test('a project from a newer program is left alone and says to update',async()=>{
  const text=JSON.stringify({format:PROJECT_FORMAT,version:99,guide:{paragraphs:[]}});
  const r=rig({text});
  await assert.rejects(r.document.ready,/newer OpenStage Guide/);
  assert.match(r.document.status().message,/newer OpenStage Guide/);
  await r.document.flush();
  assert.equal(r.writes.length,0);
  assert.equal(r.file(),text);
});

test('a file that cannot be read at all is not treated as a new guide',async()=>{
  const document=createProjectAdapter(rawDocument(),{read:async()=>{throw new Error('access denied');},write:async()=>{throw new Error('must not be called');}});
  await assert.rejects(document.ready,/access denied/);
  assert.match(document.status().message,/could not be read: access denied/);
  assert.throws(()=>document.setTitle('x'),/could not be read/);
});

test('changes before the guide has opened are refused, not lost',async()=>{
  let release;
  const document=createProjectAdapter(rawDocument(),{read:()=>new Promise(res=>{release=()=>res(null);}),write:async()=>{}});
  assert.throws(()=>document.setTitle('Too soon'),/still opening/);
  release();await document.ready;
  await document.setTitle('Now it works');
  assert.equal(await document.title(),'Now it works');
});

test('a project with a damaged paragraph is refused whole and the document is left as it was',async()=>{
  const W='xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
  for(const xml of ['this is not xml at all',`<w:r ${W}><w:t>a run, not a paragraph</w:t></w:r>`]){
    const bad=serializeProject({title:'Broken',paragraphs:[{xml:`<w:p ${W}><w:r><w:t>fine</w:t></w:r></w:p>`,text:'fine',style:'Normal'},{xml,text:'x',style:'Normal'}]});
    const r=rig({text:bad});
    r.raw.seed([{text:'Keep me',style:'Normal'}]);
    await assert.rejects(r.document.ready,/Paragraph 2 of the project is damaged/);
    assert.deepEqual(r.raw.snapshot(),[{text:'Keep me',style:'Normal'}]);
    assert.equal(r.document.blocked(),true);
    assert.equal(r.file(),bad,'the damaged file is not written over');
  }
});

test('the document gives its state out as plain data and takes it back, empty paragraphs of a table included',()=>{
  const raw=rawDocument();
  const guide=cleanGuide({title:'T',cursor:1,ownCues:['a'],look:'l',slideMemory:'s',versionMemory:'v',applied:[[{name:'X'}]],headerFooter:[{header:'h'}],paragraphs:[
    {xml:'<w:tbl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:tr><w:tc><w:p><w:r><w:t>a</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>b</w:t></w:r></w:p></w:tc></w:tr></w:tbl>',text:'a',style:'OSG Table Cell'},
    {xml:'',text:'b',style:'OSG Table Cell'}]});
  raw.importState(guide);
  assert.deepEqual(raw.exportState(),guide);
  assert.notEqual(raw.exportState().paragraphs,raw.exportState().paragraphs,'every export is its own copy');
  raw.importState(cleanGuide({paragraphs:[]}));
  assert.deepEqual(raw.snapshot(),[]);
  assert.equal(raw.undo(),false,'an opened project starts with no undo history');
});
