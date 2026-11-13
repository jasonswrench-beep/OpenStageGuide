import test from 'node:test';
import assert from 'node:assert/strict';
import {createReceiver,resultMessages,chunkBase64,fromBase64,pickFilesWithDialog,readChosen,pickWithBrowserInputs,CHUNK_BYTES,MAX_FILES,MAX_TOTAL_BYTES} from '../src/file-pick.mjs';
import {DOMParser,XMLSerializer} from '@xmldom/xmldom';
import {createWordAdapter,createDemoAdapter} from '../src/guide-word.mjs';

const bytesOf=n=>Uint8Array.from({length:n},(_,i)=>(i*31+7)&255);
const pic=(name,n)=>({name,type:'png',mime:'image/png',width:3,height:2,shrunk:false,bytes:bytesOf(n)});
const DECK={name:'d.pptx',slides:[{order:1,title:'Hello',text:'Hello',notes:'say hi',clicks:2,hidden:false}]};

test('bytes go into pieces of text and come back exactly, whatever their size',()=>{
  for(const n of [0,1,CHUNK_BYTES-1,CHUNK_BYTES,CHUNK_BYTES+1,3*CHUNK_BYTES+17]){
    const b=bytesOf(n),parts=chunkBase64(b);
    assert.ok(parts.length>=1);
    const back=new Uint8Array(parts.map(fromBase64).reduce((a,p)=>[...a,...p],[]));
    assert.deepEqual(back,b,`${n} bytes`);
  }
});

test('a result is sent as messages and put back together whole, in the same order',()=>{
  const files=[pic('a.png',CHUNK_BYTES*2+5),pic('b.png',10)];
  const messages=resultMessages({deck:DECK,files}),rx=createReceiver();
  assert.ok(messages.length>=6);
  assert.ok(messages.every(m=>typeof m==='string'&&m.length<CHUNK_BYTES*2),'each message is a small text');
  let state;for(const m of messages)state=rx.receive(m);
  assert.equal(state,'done');
  const out=rx.result();
  assert.deepEqual(out.deck,DECK);
  assert.deepEqual(out.files.map(f=>[f.name,f.type,f.width,f.height,f.bytes.length]),[['a.png','png',3,2,CHUNK_BYTES*2+5],['b.png','png',3,2,10]]);
  assert.deepEqual(out.files[0].bytes,files[0].bytes);
  assert.deepEqual(createReceiver().receive(JSON.stringify({t:'cancel'})),'cancel');
});

test('what did not arrive whole is refused with one plain message',()=>{
  const messages=resultMessages({files:[pic('a.png',CHUNK_BYTES+10)]});
  const lose=i=>{const rx=createReceiver();messages.forEach((m,k)=>{if(k!==i)rx.receive(m);});return rx;};
  assert.throws(()=>lose(2).result(),/did not arrive complete/);          // a piece missing
  assert.throws(()=>{const rx=createReceiver();rx.receive('not json');},/did not arrive complete/);
  assert.throws(()=>{const rx=createReceiver();rx.receive(JSON.stringify({t:'chunk',i:4,k:0,data:'x'}));},/did not arrive complete/);
  assert.throws(()=>{const rx=createReceiver();rx.receive(JSON.stringify({t:'nonsense'}));},/did not arrive complete/);
  const rx=createReceiver();messages.slice(0,-1).forEach(m=>rx.receive(m));
  assert.throws(()=>rx.result(),/did not arrive complete/,'no closing message yet');
  // a piece of the wrong size
  const bad=messages.map(m=>{const o=JSON.parse(m);if(o.t==='file')o.meta.size+=1;return JSON.stringify(o);});
  const rx2=createReceiver();bad.forEach(m=>rx2.receive(m));
  assert.throws(()=>rx2.result(),/did not arrive complete/);
});

test('too many files, or too many bytes, are refused before anything is sent',()=>{
  assert.throws(()=>resultMessages({files:Array.from({length:MAX_FILES+1},(_,i)=>pic(`${i}.png`,1))}),/Choose 300 files or fewer/);
  assert.throws(()=>resultMessages({files:[{...pic('big.png',1),bytes:{length:MAX_TOTAL_BYTES+1}}]}),/add up to 64 MB/);
});

// A stand-in for Office's dialog: the "window" sends the messages it is told to, or is closed.
function fakeOffice({send=[],closeAfter=false,openError=null}={}){
  const handlers={},log={opened:null,closed:false};
  const dialog={addEventHandler:(type,fn)=>{handlers[type]=fn;},close(){log.closed=true;}};
  const Office={EventType:{DialogMessageReceived:'msg',DialogEventReceived:'evt'},context:{ui:{displayDialogAsync:(url,opts,cb)=>{
    log.opened={url,opts};
    if(openError){cb({status:'failed',error:openError});return;}
    cb({status:'succeeded',value:dialog});
    queueMicrotask(()=>{for(const m of send)handlers.msg({message:m});if(closeAfter)handlers.evt({error:12006});});
  }}}};
  return {Office,log};
}

test('the pane opens the file window, collects what it sends and closes it',async()=>{
  const {Office,log}=fakeOffice({send:resultMessages({deck:DECK,files:[pic('a.png',100)]})});
  const out=await pickFilesWithDialog({Office,url:'https://localhost:3843/build/abc/picker.html',kind:'slides'});
  assert.equal(log.opened.url,'https://localhost:3843/build/abc/picker.html?kind=slides');
  assert.equal(log.opened.opts.displayInIframe,false);
  assert.equal(out.files.length,1);assert.equal(out.deck.name,'d.pptx');assert.equal(log.closed,true);
});

test('closing the window, pressing Cancel, a refusal to open it or a broken message each end in a plain answer',async()=>{
  const url='https://localhost:3843/picker.html';
  assert.deepEqual(await pickFilesWithDialog({...fakeOffice({closeAfter:true}),url,kind:'picture'}),{cancelled:true});
  assert.deepEqual(await pickFilesWithDialog({...fakeOffice({send:[JSON.stringify({t:'cancel'})]}),url,kind:'picture'}),{cancelled:true});
  await assert.rejects(pickFilesWithDialog({...fakeOffice({openError:{code:12007}}),url,kind:'picture'}),/A file window is already open/);
  await assert.rejects(pickFilesWithDialog({...fakeOffice({openError:{code:12009}}),url,kind:'picture'}),/choose Allow/);
  await assert.rejects(pickFilesWithDialog({...fakeOffice({openError:{code:1,message:'odd'}}),url,kind:'picture'}),/could not open the file window \(odd\)/);
  await assert.rejects(pickFilesWithDialog({...fakeOffice({send:['garbage']}),url,kind:'picture'}),/did not arrive complete/);
  await assert.rejects(pickFilesWithDialog({Office:{},url,kind:'picture'}),/cannot open the file window/);
  const slow={Office:{EventType:{},context:{ui:{displayDialogAsync:(u,o,cb)=>cb({status:'succeeded',value:{addEventHandler(){},close(){}}})}}}};
  await assert.rejects(pickFilesWithDialog({...slow,url,kind:'picture',timeoutMs:20}),/took too long/);
});

test('the real adapter opens picker.html next to the pane\'s own page, and looks Office up when it is used',async()=>{
  const {Office,log}=fakeOffice({send:resultMessages({files:[pic('a.png',5)]})});
  const had=Object.getOwnPropertyDescriptor(globalThis,'location');
  Object.defineProperty(globalThis,'location',{value:{href:'https://localhost:3843/build/abc123/taskpane.html'},configurable:true});
  try{
    const adapter=createWordAdapter({Office,fetchFn:async()=>({}),DOMParserClass:class{}});
    const out=await adapter.pickFiles({kind:'picture'});
    assert.equal(log.opened.url,'https://localhost:3843/build/abc123/picker.html?kind=picture');
    assert.equal(out.files[0].name,'a.png');
  }finally{if(had)Object.defineProperty(globalThis,'location',had);else delete globalThis.location;}
});

test('the file window\'s own work: pictures read and shrunk, the presentation read for its slides',async()=>{
  const prepared=[];
  const out=await readChosen({
    deckFile:{name:'d.pptx'},pictureFiles:[{name:'a.png'},{name:'b.png'}],
    prepare:async f=>{prepared.push(f.name);return pic(f.name,4);},
    readDeck:async()=>({slides:[{order:1,sourceId:256,title:'T',text:'T body',notes:'N',clicks:3,hidden:true},{order:2,title:'U',text:'',notes:''}]})
  });
  assert.deepEqual(prepared,['a.png','b.png']);
  assert.deepEqual(out.deck,{name:'d.pptx',slides:[{order:1,title:'T',text:'T body',notes:'N',clicks:3,hidden:true,titleGuessed:false,sourceId:256},{order:2,title:'U',text:'',notes:'',clicks:0,hidden:false,titleGuessed:false}]});
  assert.equal(out.files.length,2);
  assert.deepEqual(await readChosen({pictureFiles:[]}),{deck:null,files:[]});
  await assert.rejects(readChosen({deckFile:{name:'d.pptx'}}),/cannot read a presentation/);
});

test('the preview uses plain file inputs, one after another, and cancelling is not a problem',async()=>{
  const made=[];
  const doc={createElement:tag=>{const el={tag,click(){made.push(el);setImmediate(()=>el.onchange());}};return el;}};
  const answers=[[{name:'p.png'}]];
  // each input's files are set just before its change fires
  const document={createElement:()=>{const el={files:answers.shift()||[],click(){setImmediate(()=>el.onchange());}};return el;}};
  const out=await pickWithBrowserInputs('picture',{document,prepare:async f=>pic(f.name,3)});
  assert.equal(out.files[0].name,'p.png');
  const none=await pickWithBrowserInputs('picture',{document:{createElement:()=>({files:[],click(){setImmediate(()=>this.onchange());}})},prepare:async()=>{throw new Error('should not be asked');}});
  assert.deepEqual(none,{cancelled:true});
  void doc;void made;
});

test('the demo adapter lets a test stand in for the author\'s choice',async()=>{
  const adapter=createDemoAdapter({DOMParserClass:DOMParser,SerializerClass:XMLSerializer});
  adapter.setPicker(async opts=>({files:[pic('x.png',2)],kind:opts.kind}));
  assert.equal((await adapter.pickFiles({kind:'slides'})).kind,'slides');
});

// ---- any other file box in the pane (the OpenAssess bank, a backup, saved settings): the same file window, the file as it is
import {pickerAddress,readRawFiles} from '../src/file-pick.mjs';
import fs from 'node:fs';

test('the file window\'s address says what kind of choice it is, and for any other file what the box asked for',()=>{
  const u='https://localhost:3843/picker.html';
  assert.equal(pickerAddress(u,'picture'),`${u}?kind=picture`);
  assert.equal(pickerAddress(u,'slides',{accept:'.pptx'}),`${u}?kind=slides`,'only the any-file kind carries what the box asked for');
  assert.equal(pickerAddress(u,'files',{accept:'.openassess',label:'Bank file'}),`${u}?kind=files&accept=.openassess&label=Bank+file`);
  assert.equal(pickerAddress(u,'files',{accept:'.osgbackup,.zip',multiple:true}),`${u}?kind=files&accept=.osgbackup%2C.zip&multiple=1`);
  assert.match(pickerAddress(u,'files',{label:'x'.repeat(500)}),/label=x{120}$/);
});

test('any file goes to the pane as it is, through the same messages as the pictures',async()=>{
  const raw=await readRawFiles([{name:'bank.openassess',type:'application/zip',arrayBuffer:async()=>Uint8Array.from([80,75,3,4]).buffer}]);
  assert.deepEqual(raw.map(f=>[f.name,f.type,f.mime,Array.from(f.bytes)]),[['bank.openassess','','application/zip',[80,75,3,4]]]);
  const {Office,log}=fakeOffice({send:resultMessages({files:raw})});
  const out=await pickFilesWithDialog({Office,url:'https://localhost:3843/picker.html',kind:'files',accept:'.openassess',label:'Bank file'});
  assert.equal(log.opened.url,'https://localhost:3843/picker.html?kind=files&accept=.openassess&label=Bank+file');
  assert.equal(out.deck,null);
  assert.deepEqual(out.files.map(f=>[f.name,f.mime,Array.from(f.bytes)]),[['bank.openassess','application/zip',[80,75,3,4]]]);
  assert.throws(()=>resultMessages({files:[{name:'big.zip',type:'',mime:'',bytes:{length:MAX_TOTAL_BYTES+1}}]}),/Choose fewer or smaller files/);
});

test('the real adapter passes what a file box asked for on to the file window, and the preview offers a plain file box',async()=>{
  const {Office,log}=fakeOffice({send:resultMessages({files:[{name:'b.zip',type:'',mime:'',bytes:Uint8Array.from([1])}]})});
  const had=Object.getOwnPropertyDescriptor(globalThis,'location');
  Object.defineProperty(globalThis,'location',{value:{href:'https://localhost:3843/build/abc123/taskpane.html'},configurable:true});
  try{
    const adapter=createWordAdapter({Office,fetchFn:async()=>({}),DOMParserClass:class{}});
    await adapter.pickFiles({kind:'files',accept:'.osgbackup,.zip',multiple:false,label:'Backup file'});
    assert.equal(log.opened.url,'https://localhost:3843/build/abc123/picker.html?kind=files&accept=.osgbackup%2C.zip&label=Backup+file');
  }finally{if(had)Object.defineProperty(globalThis,'location',had);else delete globalThis.location;}
  const seen=[];
  const document={createElement:()=>{const el={files:[{name:'p.json',type:'application/json',arrayBuffer:async()=>Uint8Array.from([123,125]).buffer}],click(){seen.push([el.accept,el.multiple]);setImmediate(()=>el.onchange());}};return el;}};
  const out=await pickWithBrowserInputs('files',{document,accept:'.json',multiple:true});
  assert.deepEqual(seen,[['.json',true]]);
  assert.deepEqual(out.files.map(f=>[f.name,Array.from(f.bytes)]),[['p.json',[123,125]]]);
  assert.deepEqual(await pickWithBrowserInputs('files',{document:{createElement:()=>({files:[],click(){setImmediate(()=>this.onchange());}})}}),{cancelled:true});
});

test('the file window page has the part for any file, hidden until it is asked for',()=>{
  const html=fs.readFileSync(new URL('../public/picker.html',import.meta.url),'utf8');
  assert.match(html,/<div id="picker-files" hidden>/);
  assert.match(html,/<label id="pick-any-label" for="pick-any">/);
  assert.match(html,/<input id="pick-any" type="file">/);
  const code=fs.readFileSync(new URL('../src/picker.mjs',import.meta.url),'utf8');
  assert.match(code,/asked==='files'\?'files'/);
  assert.match(code,/readRawFiles\(rawFiles\)/);
});

test('the slide file window warns once when no PDF or pictures are chosen, then goes on without them if asked (author, 2026-11-07)',()=>{
  const code=fs.readFileSync(new URL('../src/picker.mjs',import.meta.url),'utf8');
  assert.match(code,/let warnedNoPictures=false;/);
  assert.match(code,/if\(!chosenPictures\.length&&!warnedNoPictures\)\{/,'the first press with no pictures does not send');
  assert.match(code,/the slides will go in as text only, with no pictures/);
  assert.match(code,/\$\('pick-go'\)\.textContent='Use without pictures';/);
  assert.match(code,/\$\('pick-slide-pictures'\)\.onchange=pictureChoiceChanged;/,'choosing pictures afterwards puts the window back');
  // only the slides kind warns: the picture and any-file kinds are untouched
  assert.equal(code.split('warnedNoPictures=true').length-1,1);
});
