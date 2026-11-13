// Every file box in the pane opens the pane's file window in Word (a file box does nothing in Word's pane: seen 2026-11-06 and, for the
// OpenAssess bank, 2026-11-07). Tested with stand-ins for the page, the file box and the file window. This is not native Word.
import test from 'node:test';
import assert from 'node:assert/strict';
import {installFileInputBridge,fileListFrom,labelOfInput} from '../src/file-input-bridge.mjs';

class FakeFile{constructor(parts,name,opts={}){this.parts=parts;this.name=name;this.type=opts.type||'';}}
class FakeTransfer{constructor(){this.list=[];this.items={add:f=>this.list.push(f)};}get files(){return this.list;}}
class FakeEvent{constructor(type,init={}){this.type=type;this.bubbles=!!init.bubbles;}}
const ctors={FileCtor:FakeFile,DataTransferCtor:FakeTransfer,EventCtor:FakeEvent};

function page(){
  const listeners=[],labels=new Map();
  const document={
    addEventListener:(type,fn,capture)=>listeners.push({type,fn,capture}),
    removeEventListener:(type,fn)=>{const i=listeners.findIndex(l=>l.type===type&&l.fn===fn);if(i>=0)listeners.splice(i,1);},
    querySelector:sel=>{const m=sel.match(/^label\[for="(.+)"\]$/);return m&&labels.has(m[1])?{textContent:labels.get(m[1])}:null;}
  };
  const box=(id,{accept='',multiple=false,label}={})=>{
    if(label!==undefined)labels.set(id,label);
    return {id,tagName:'INPUT',type:'file',accept,multiple,files:[],seen:[],dispatchEvent(e){this.seen.push(e.type);return true;},getAttribute:()=>null};
  };
  const click=async target=>{
    const ev={target,prevented:false,stopped:false,preventDefault(){this.prevented=true;},stopPropagation(){this.stopped=true;}};
    for(const l of listeners.filter(l=>l.type==='click'))await l.fn(ev);
    return ev;
  };
  return {document,listeners,box,click};
}
const entry=(name,mime='',n=3)=>({name,type:'',mime,bytes:Uint8Array.from({length:n},(_,i)=>i+1)});

test('pressing a file box opens the file window for what that box asked for, and the box gets what was chosen',async()=>{
  const p=page(),bank=p.box('guide-bank-file',{accept:'.openassess',label:'  Bank   file '}),asked=[];
  installFileInputBridge({document:p.document,pick:async o=>{asked.push(o);return {deck:null,files:[entry('my.openassess','application/octet-stream',5)]};},...ctors});
  const ev=await p.click(bank);
  assert.equal(ev.prevented,true,'the box\'s own (dead) window is not left to open');
  assert.deepEqual(asked,[{kind:'files',accept:'.openassess',multiple:false,label:'Bank file'}]);
  assert.equal(bank.files.length,1);
  assert.equal(bank.files[0].name,'my.openassess');
  assert.equal(bank.files[0].type,'application/octet-stream');
  assert.deepEqual(Array.from(bank.files[0].parts[0]),[1,2,3,4,5],'the bytes are the file\'s bytes');
  assert.deepEqual(bank.seen,['input','change'],'the box announces the choice, so the code that reads it runs as before');
});

test('a box that takes several files says so, and gets them all',async()=>{
  const p=page(),many=p.box('pick',{accept:'.json,application/json',multiple:true,label:'Settings'}),asked=[];
  installFileInputBridge({document:p.document,pick:async o=>{asked.push(o);return {files:[entry('a.json'),entry('b.json')]};},...ctors});
  await p.click(many);
  assert.equal(asked[0].multiple,true);assert.equal(asked[0].accept,'.json,application/json');
  assert.deepEqual(many.files.map(f=>f.name),['a.json','b.json']);
});

test('cancelling the file window leaves the box and what it held alone',async()=>{
  const p=page(),box=p.box('b');box.files=['kept'];
  installFileInputBridge({document:p.document,pick:async()=>({cancelled:true}),...ctors});
  const ev=await p.click(box);
  assert.equal(ev.prevented,true);assert.deepEqual(box.files,['kept']);assert.deepEqual(box.seen,[]);
  const p2=page(),b2=p2.box('b');
  installFileInputBridge({document:p2.document,pick:async()=>({files:[]}),...ctors});
  await p2.click(b2);assert.deepEqual(b2.seen,[],'a window that sent nothing is the same as cancelling');
});

test('a problem opening the window is told in words, and the next press works',async()=>{
  const p=page(),box=p.box('b'),told=[];let fail=true;
  installFileInputBridge({document:p.document,onError:e=>told.push(e.message),pick:async()=>{if(fail)throw new Error('Word does not trust the address of the file window.');return {files:[entry('x.bin')]};},...ctors});
  await p.click(box);
  assert.deepEqual(told,['Word does not trust the address of the file window.']);assert.deepEqual(box.seen,[]);
  fail=false;await p.click(box);
  assert.equal(box.files.length,1);
});

test('only file boxes are taken over, and a second press while the window is open does nothing',async()=>{
  const p=page(),box=p.box('b');let calls=0,release;
  installFileInputBridge({document:p.document,pick:()=>{calls++;return new Promise(r=>{release=()=>r({files:[entry('x.bin')]});});},...ctors});
  const other=await p.click({tagName:'BUTTON',type:'button'});
  assert.equal(other.prevented,false);
  const text=await p.click({tagName:'INPUT',type:'text'});
  assert.equal(text.prevented,false);
  const first=p.click(box);
  const second=await p.click(box);
  assert.equal(second.prevented,true,'still not left to the dead window');
  assert.equal(calls,1);
  release();await first;
  assert.equal(box.files.length,1);
  const third=p.click(box);release();await third;
  assert.equal(calls,2,'after the window is done a new press opens it again');
});

test('in the plain-browser preview nothing is taken over, and the bridge can be taken off',()=>{
  const p=page();
  const off=installFileInputBridge({document:p.document,pick:async()=>({}),enabled:false,...ctors});
  assert.equal(p.listeners.length,0);off();
  const on=installFileInputBridge({document:p.document,pick:async()=>({}),...ctors});
  assert.equal(p.listeners.length,1);assert.equal(p.listeners[0].capture,true,'it hears the press before the box\'s own handlers');
  on();assert.equal(p.listeners.length,0);
  assert.doesNotThrow(()=>installFileInputBridge({document:null,pick:async()=>({})})());
});

test('the file window\'s name for a box is its label, or its own name',()=>{
  const p=page(),a=p.box('a',{label:'Bank file'}),b=p.box('b');
  assert.equal(labelOfInput(a,p.document),'Bank file');
  assert.equal(labelOfInput(b,p.document),'');
  assert.equal(labelOfInput({id:'c',getAttribute:k=>k==='aria-label'?' Logo  picture ':null},p.document),'Logo picture');
  assert.equal(fileListFrom([entry('z.txt','text/plain')],ctors).length,1);
});
