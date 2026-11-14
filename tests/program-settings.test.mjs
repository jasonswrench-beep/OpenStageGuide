// What the program remembers besides the guides (src/program-settings.mjs), kept by its server in a file beside the guides, and the folder the guides
// are in when no folder is named (scripts/project-server.mjs).
import test from 'node:test';
import assert from 'node:assert/strict';
import http from 'node:http';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import {startProjectServer,documentsFolder,defaultGuidesFolder,PROJECT_EXTENSION} from '../scripts/project-server.mjs';
import {parseSettings,serializeSettings,createProgramStorage,SETTING_KEYS,SETTINGS_FILE,SETTING_MAX_CHARS,AGENDA_START_KEY,LOOK_KEY} from '../src/program-settings.mjs';
import {STORE_KEY} from '../src/activity-library.mjs';

const snapshot={version:'0-test',buildId:'test-build',record:{},read:()=>null};

async function withServer(fn,{folder:given}={}){
  const dir=fs.mkdtempSync(path.join(os.tmpdir(),'osg-settings-'));
  const folder=given??path.join(dir,'My guides');
  const start=async()=>{
    const s=await startProjectServer({folder,port:0,snapshot});
    const call=(method,urlPath,{headers={},body,host=`127.0.0.1:${s.port}`}={})=>new Promise((resolve,reject)=>{
      const req=http.request({host:'127.0.0.1',port:s.port,path:urlPath,method,headers:{Host:host,...headers}},res=>{let d='';res.on('data',c=>d+=c);res.on('end',()=>resolve({status:res.statusCode,text:d,json:()=>JSON.parse(d)}));});
      req.on('error',reject);if(body!==undefined)req.write(body);req.end();
    });
    const set=(key,value)=>call('POST','/api/settings',{headers:{'Content-Type':'application/json'},body:JSON.stringify({key,value})});
    return {s,call,set};
  };
  try{await fn({folder,dir,start});}finally{fs.rmSync(dir,{recursive:true,force:true});}
}

test('the names the program keeps are the three the pane uses, and the file reads back only those, as text',()=>{
  assert.deepEqual([...SETTING_KEYS],[STORE_KEY,AGENDA_START_KEY,LOOK_KEY]);
  const text=serializeSettings({[STORE_KEY]:'[1]',[AGENDA_START_KEY]:'13:30',other:'x',[LOOK_KEY]:5});
  assert.deepEqual(parseSettings(text),{[STORE_KEY]:'[1]',[AGENDA_START_KEY]:'13:30'},'an unknown name and a value that is not text are not kept');
  for(const bad of ['','not json','[]','{"format":"other","values":{}}',`{"format":"openstage-guide-settings","values":[]}`,'null'])assert.deepEqual(parseSettings(bad),{},`${bad}`);
  assert.deepEqual(parseSettings(JSON.stringify({format:'openstage-guide-settings',values:{[LOOK_KEY]:'x'.repeat(SETTING_MAX_CHARS+1),[AGENDA_START_KEY]:'09:00'}})),{[AGENDA_START_KEY]:'09:00'},'a value that is too large is left out');
});

test('a setting is kept in a file beside the guides, is not listed as a guide, and is there when the server starts again on another port',async()=>{
  await withServer(async({folder,start})=>{
    const first=await start();
    try{
      assert.deepEqual((await first.call('GET','/api/settings')).json(),{values:{}},'nothing kept yet');
      assert.equal(fs.existsSync(folder),false,'asking makes no folder');
      const saved=await first.set(STORE_KEY,'[{"id":"a"}]');
      assert.equal(saved.status,200);
      assert.equal((await first.set(AGENDA_START_KEY,'13:30')).status,200);
      assert.deepEqual(fs.readdirSync(folder),[SETTINGS_FILE],'one file, no half-way copy left behind');
      assert.deepEqual((await first.call('GET','/api/state')).json().guides,[],'it is not a guide');
      assert.ok(!SETTINGS_FILE.endsWith(PROJECT_EXTENSION));
      assert.deepEqual((await first.call('GET','/api/settings')).json().values,{[STORE_KEY]:'[{"id":"a"}]',[AGENDA_START_KEY]:'13:30'});
    }finally{await first.s.close();}
    const second=await start();
    try{
      assert.notEqual(second.s.port,first.s.port,'a new port, as the window gets every time');
      assert.deepEqual((await second.call('GET','/api/settings')).json().values,{[STORE_KEY]:'[{"id":"a"}]',[AGENDA_START_KEY]:'13:30'},'what was kept is still there');
      assert.equal((await second.set(AGENDA_START_KEY,null)).status,200);
      assert.deepEqual((await second.call('GET','/api/settings')).json().values,{[STORE_KEY]:'[{"id":"a"}]'},'null takes a setting away');
    }finally{await second.s.close();}
  });
});

test('only the three names are kept, only as text, and nothing is kept for a page from another site or in another form',async()=>{
  await withServer(async({folder,start})=>{
    const {s,call,set}=await start();
    try{
      for(const key of ['__proto__','constructor','../x','activity','',null,5,{}])assert.equal((await set(key,'x')).status,400,String(key));
      for(const value of [5,true,{},['x']])assert.equal((await set(STORE_KEY,value)).status,400,JSON.stringify(value));
      assert.equal((await set(STORE_KEY,'x'.repeat(SETTING_MAX_CHARS+1))).status,413,'one setting may not be larger than this');
      const json={'Content-Type':'application/json'},body=JSON.stringify({key:STORE_KEY,value:'x'});
      assert.equal((await call('POST','/api/settings',{headers:{'Content-Type':'text/plain'},body})).status,415,'not in the form the program sends');
      assert.equal((await call('POST','/api/settings',{headers:{...json,'Sec-Fetch-Site':'cross-site'},body})).status,403,'not from another site');
      assert.equal((await call('POST','/api/settings',{headers:{...json,Origin:'http://evil.example'},body})).status,403,'not from another origin');
      assert.equal((await call('POST','/api/settings',{headers:json,host:'evil.example',body})).status,403,'not under another host name');
      assert.equal((await call('PUT','/api/settings',{headers:json,body})).status,405);
      assert.equal(fs.existsSync(folder),false,'nothing was written for any of them');
      assert.equal((await set(STORE_KEY,'ok')).status,200);
      // Two values that are each allowed but together are too much for the file are refused, and the file is as it was.
      assert.equal((await set(LOOK_KEY,'y'.repeat(SETTING_MAX_CHARS))).status,200);
      assert.equal((await set(AGENDA_START_KEY,'z'.repeat(SETTING_MAX_CHARS))).status,413);
      assert.deepEqual(Object.keys((await call('GET','/api/settings')).json().values).sort(),[STORE_KEY,LOOK_KEY].sort());
    }finally{await s.close();}
  });
});

test('a settings file that is damaged, or is not ours, reads as no settings, and the next setting replaces it',async()=>{
  await withServer(async({folder,start})=>{
    fs.mkdirSync(folder);
    fs.writeFileSync(path.join(folder,SETTINGS_FILE),'{ not json');
    const {s,call,set}=await start();
    try{
      assert.deepEqual((await call('GET','/api/settings')).json(),{values:{}});
      assert.equal((await set(LOOK_KEY,'{"look":"clearline"}')).status,200);
      assert.deepEqual((await call('GET','/api/settings')).json().values,{[LOOK_KEY]:'{"look":"clearline"}'});
    }finally{await s.close();}
  });
});

test('many settings written at once all arrive, one at a time',async()=>{
  await withServer(async({start})=>{
    const {s,call,set}=await start();
    try{
      const keys=[STORE_KEY,AGENDA_START_KEY,LOOK_KEY];
      await Promise.all(keys.flatMap(k=>[1,2,3,4].map(n=>set(k,`${k}:${n}`))));
      const values=(await call('GET','/api/settings')).json().values;
      assert.deepEqual(Object.keys(values).sort(),[...keys].sort());
      for(const k of keys)assert.match(values[k],new RegExp(`^${k.replace(/[.]/g,'\\.')}:[1-4]$`),'a whole value, never a mixture');
    }finally{await s.close();}
  });
});

// ---- the page's side

function fakeRequest(server){
  return ()=>{
    const r={headers:{},setRequestHeader(k,v){r.headers[k]=v;},open(method,url,async){r.method=method;r.url=url;r.async=async;},send(body){
      const out=server(r.method,r.url,body);r.status=out.status;r.responseText=out.text;}};
    return r;
  };
}
const memory=()=>{const m=new Map();return {getItem:k=>m.has(k)?m.get(k):null,setItem:(k,v)=>m.set(k,String(v)),removeItem:k=>m.delete(k),map:m};};

test('the page\'s storage asks the program at once, and gives back what it kept',()=>{
  const kept={};
  const calls=[];
  const server=(method,url,body)=>{
    calls.push([method,url]);
    if(method==='GET')return {status:200,text:JSON.stringify({values:kept})};
    const {key,value}=JSON.parse(body);if(value===null)delete kept[key];else kept[key]=value;return {status:200,text:'{"ok":true}'};
  };
  const storage=createProgramStorage({xhr:fakeRequest(server),legacy:null});
  assert.equal(storage.getItem(LOOK_KEY),null);
  storage.setItem(LOOK_KEY,'{"look":"practice-lab"}');
  assert.equal(storage.getItem(LOOK_KEY),'{"look":"practice-lab"}');
  storage.removeItem(LOOK_KEY);
  assert.equal(storage.getItem(LOOK_KEY),null);
  assert.ok(calls.every(([,url])=>url==='/api/settings'),'it only ever asks the program\'s one address');
});

test('what the program refuses is thrown, so the pane can say it could not keep it; a program that does not answer reads as nothing kept',()=>{
  const refuse=createProgramStorage({xhr:fakeRequest(()=>({status:413,text:JSON.stringify({error:'That is larger than the program can keep.'})})),legacy:null});
  assert.throws(()=>refuse.setItem(STORE_KEY,'x'),/larger than the program can keep/);
  const dead=createProgramStorage({xhr:fakeRequest(()=>({status:500,text:'The program could not do that.'})),legacy:null});
  assert.equal(dead.getItem(STORE_KEY),null);
  assert.throws(()=>dead.setItem(STORE_KEY,'x'),/The program answered 500|could not do that/);
});

test('an activity kept in this browser\'s own local storage before the program had a file is carried over, once, and not brought back after it is removed',()=>{
  const kept={};
  const server=(method,url,body)=>{
    if(method==='GET')return {status:200,text:JSON.stringify({values:kept})};
    const {key,value}=JSON.parse(body);if(value===null)delete kept[key];else kept[key]=value;return {status:200,text:'{}'};
  };
  const legacy=memory();legacy.setItem(STORE_KEY,'[{"id":"old"}]');
  const storage=createProgramStorage({xhr:fakeRequest(server),legacy});
  assert.equal(storage.getItem(STORE_KEY),'[{"id":"old"}]');
  assert.deepEqual(kept,{[STORE_KEY]:'[{"id":"old"}]'},'now the program has it');
  assert.equal(legacy.getItem(STORE_KEY),'[{"id":"old"}]','the browser\'s copy is left alone');
  kept[STORE_KEY]='[{"id":"new"}]';
  assert.equal(storage.getItem(STORE_KEY),'[{"id":"new"}]','the program\'s copy wins once it has one');
  storage.removeItem(STORE_KEY);
  assert.equal(storage.getItem(STORE_KEY),null);
});

test('the pane and the program use these names, and the program gives the pane this storage',()=>{
  const pane=fs.readFileSync(new URL('../src/guide-pane.mjs',import.meta.url),'utf8');
  assert.match(pane,/import \{AGENDA_START_KEY,LOOK_KEY\} from '\.\/program-settings\.mjs';/);
  assert.doesNotMatch(pane,/const (AGENDA_START_KEY|LOOK_KEY)\s*=/,'the names are defined once');
  const app=fs.readFileSync(new URL('../src/app.mjs',import.meta.url),'utf8');
  assert.match(app,/programMode\?\{storage:createProgramStorage\(\)\}:\{\}/,'in the program the pane\'s storage is the program\'s file');
});

// ---- the guides folder

test('with no folder named, the guides go in the Documents folder Windows reports, the one in OneDrive when OneDrive holds Documents',()=>{
  const home='C:\\Users\\Jason';
  const asked=[];
  const reg=answer=>env=>{asked.push(env.SystemRoot);return answer;};
  assert.equal(documentsFolder({platform:'win32',home,env:{SystemRoot:'C:\\Windows'},readRegistry:reg('C:\\Users\\Jason\\OneDrive\\Documents')}),'C:\\Users\\Jason\\OneDrive\\Documents');
  assert.equal(documentsFolder({platform:'win32',home,env:{},readRegistry:reg('D:\\Work\\Documents\\')}),'D:\\Work\\Documents\\');
  assert.equal(path.win32.join(documentsFolder({platform:'win32',home,env:{},readRegistry:reg('C:\\Users\\Jason\\OneDrive\\Documents')}),'OpenStage Guides'),'C:\\Users\\Jason\\OneDrive\\Documents\\OpenStage Guides');
  const usual=path.join(home,'Documents');
  for(const odd of [null,undefined,'','Documents','%USERPROFILE%\\Documents','C:\\x\u0001y',7])assert.equal(documentsFolder({platform:'win32',home,env:{},readRegistry:reg(odd)}),usual,`${String(odd)}`);
  assert.equal(documentsFolder({platform:'win32',home,env:{},readRegistry:()=>{throw new Error('no reg.exe');}}),usual,'no answer from Windows');
  assert.equal(documentsFolder({platform:'linux',home:'/home/jo',env:{},readRegistry:()=>{throw new Error('must not be asked');}}),path.join('/home/jo','Documents'));
  assert.equal(defaultGuidesFolder({platform:'linux',home:'/home/jo',env:{}}),path.join('/home/jo','Documents','OpenStage Guides'));
});

test('on this computer the registry answers with an absolute folder that ends in Documents, or the usual one is used (reads the registry; asks nothing else)',{skip:process.platform!=='win32'},()=>{
  const folder=documentsFolder();
  assert.ok(path.isAbsolute(folder),folder);
  assert.match(folder,/Documents$/i);
});
