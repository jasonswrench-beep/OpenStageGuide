// The connection to Prism Reference Manager (scripts/prm-link.mjs): where its token is kept (the program's data folder, never the guides folder), what is done with a
// file left in the guides folder by earlier releases, and what is done with a stored port that is not a port. Nothing here reaches a network or the person's real folders.
import test from 'node:test';
import assert from 'node:assert/strict';
import http from 'node:http';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import {createPrmLink,PRM_FILE} from '../scripts/prm-link.mjs';
import {startProjectServer,defaultDataFolder,defaultGuidesFolder} from '../scripts/project-server.mjs';
import {SOURCES} from '../src/design-models.mjs';

const DATA=`/data/org.openstage.guide/${PRM_FILE}`,OLD=`/guides/${PRM_FILE}`;
const snapshot={version:'0-test',buildId:'test-build',record:{},read:()=>null};

// Files in memory, with a note of what was written and removed, in order.
function memory(initial={}){
  const files=new Map(Object.entries(initial));
  const log=[];
  const fs={
    files,log,failWrites:false,
    read:async f=>{if(!files.has(f))throw Object.assign(new Error('ENOENT'),{code:'ENOENT'});return files.get(f);},
    write:async(f,t)=>{if(fs.failWrites)throw new Error('no room');log.push(['write',f]);files.set(f,t);},
    remove:async f=>{log.push(['remove',f]);files.delete(f);},
  };
  return fs;
}
const reply=(data,status=200)=>new Response(JSON.stringify(data),{status,headers:{'Content-Type':'application/json'}});
// Prism Reference Manager as its documentation says: pairing, then calls that need the token it gave.
function standIn(){
  const calls=[];
  const fetch=async(url,init={})=>{
    const u=new URL(String(url));
    calls.push({path:u.pathname,port:u.port,method:init.method,auth:init.headers?.Authorization});
    if(u.pathname==='/v1/pair'&&init.method==='POST')return reply({id:'p1'});
    if(u.pathname==='/v1/pair/p1')return reply({state:'approved',token:'tok-123'});
    if(u.pathname==='/v1/call')return init.headers?.Authorization==='Bearer tok-123'?reply({result:{name:'Library'}}):reply({error:'Not allowed.'},401);
    return reply({error:'Not found.'},404);
  };
  return {fetch,calls};
}
const link=(m,prm,extra={})=>createPrmLink({file:DATA,legacyFile:OLD,fetch:prm.fetch,read:m.read,write:m.write,remove:m.remove,...extra});

test('the token is written under the data folder, and nothing is written anywhere else',async()=>{
  const m=memory(),prm=standIn(),l=link(m,prm);
  assert.equal((await l.handle({action:'connect',port:23119})).state,'pending');
  assert.equal((await l.handle({action:'status'})).state,'connected');
  assert.deepEqual([...m.files.keys()],[DATA],'one file, in the data folder');
  assert.deepEqual(JSON.parse(m.files.get(DATA)),{port:23119,token:'tok-123'});
  assert.ok(m.log.every(([,f])=>f===DATA),'no write or removal outside the data folder: '+JSON.stringify(m.log));
  // and the kept token is the one a later call uses
  const again=await l.handle({action:'status'});
  assert.equal(again.state,'connected');
  assert.ok(prm.calls.some(c=>c.path==='/v1/call'&&c.auth==='Bearer tok-123'));
  assert.equal((await l.handle({action:'forget'})).state,'not-connected');
  assert.equal(m.files.size,0);
});

test('a stored port that is not a port is not a connection: the token is not sent, and the file is dropped',async()=>{
  for(const port of [0,70000,1023,-1,23119.5,'23119',null]){
    const m=memory({[DATA]:JSON.stringify({port,token:'tok-123'})}),prm=standIn(),l=link(m,prm);
    const out=await l.handle({action:'status'});
    assert.equal(out.state,'not-connected',`port ${JSON.stringify(port)}`);
    assert.ok(prm.calls.every(c=>!c.auth),`the token was not sent to port ${JSON.stringify(port)}`);
    assert.equal(m.files.has(DATA),false,`port ${JSON.stringify(port)}: the file is dropped`);
    const cite=await createPrmLink({file:DATA,fetch:prm.fetch,read:memory({[DATA]:JSON.stringify({port,token:'tok-123'})}).read,write:m.write,remove:m.remove}).handle({action:'cite',source:Object.keys(SOURCES)[0]});
    assert.equal(cite.ok,false);assert.equal(cite.code,'unauthorized');
  }
  // a good port still works, and a file that is not whole yet is left where it is
  const good=memory({[DATA]:JSON.stringify({port:23119,token:'tok-123'})});
  assert.equal((await link(good,standIn()).handle({action:'status'})).state,'connected');
  const half=memory({[DATA]:'{"port":23'});
  assert.equal((await link(half,standIn()).handle({action:'status'})).state,'not-connected');
  assert.equal(half.files.has(DATA),true,'a file being written is not removed');
});

test('a connection kept in the guides folder is moved to the data folder once, and the old file is deleted',async()=>{
  const old=JSON.stringify({port:23119,token:'tok-123'});
  const m=memory({[OLD]:old}),prm=standIn(),l=link(m,prm);
  await l.migrate();
  assert.deepEqual([...m.files.keys()],[DATA]);
  assert.deepEqual(JSON.parse(m.files.get(DATA)),{port:23119,token:'tok-123'});
  assert.deepEqual(m.log,[['write',DATA],['remove',OLD]],'written first, then the old one removed');
  await l.migrate();
  assert.equal(m.log.length,2,'a second start does nothing more');
  assert.equal((await l.handle({action:'status'})).state,'connected');

  // used before the server asked for the move: the first action moves it
  const m2=memory({[OLD]:old}),l2=link(m2,standIn());
  assert.equal((await l2.handle({action:'status'})).state,'connected');
  assert.equal(m2.files.has(OLD),false);

  // already a connection in the data folder: it stays, and the old file goes
  const newer=JSON.stringify({port:23120,token:'tok-new'});
  const m3=memory({[OLD]:old,[DATA]:newer});
  await link(m3,standIn()).migrate();
  assert.equal(m3.files.get(DATA),newer);assert.equal(m3.files.has(OLD),false);

  // an old file with a port that is not a port, or no token, is deleted and nothing is carried over
  for(const bad of [{port:0,token:'t'},{port:70000,token:'t'},{port:23119},{token:'t'},'not json']){
    const m4=memory({[OLD]:typeof bad==='string'?bad:JSON.stringify(bad)});
    await link(m4,standIn()).migrate();
    assert.equal(m4.files.size,0,JSON.stringify(bad));
  }

  // nothing in the guides folder: nothing is done
  const m5=memory();
  await link(m5,standIn()).migrate();
  assert.deepEqual(m5.log,[]);

  // the data folder cannot be written: the old file is not lost, and the next start tries again
  const m6=memory({[OLD]:old});m6.failWrites=true;
  await link(m6,standIn()).migrate();
  assert.equal(m6.files.get(OLD),old);
  m6.failWrites=false;
  await link(m6,standIn()).migrate();
  assert.equal(m6.files.has(OLD),false);assert.ok(m6.files.has(DATA));
});

test('the data folder is the program\'s own, by platform, and is never the guides folder',()=>{
  const win={platform:'win32',home:'C:\\Users\\Ana',env:{APPDATA:'C:\\Users\\Ana\\AppData\\Roaming'}};
  assert.equal(defaultDataFolder(win),'C:\\Users\\Ana\\AppData\\Roaming\\org.openstage.guide');
  assert.equal(defaultDataFolder({...win,env:{}}),'C:\\Users\\Ana\\AppData\\Roaming\\org.openstage.guide','no answer from Windows: the usual place');
  assert.equal(defaultDataFolder({...win,env:{APPDATA:'..\\x'}}),'C:\\Users\\Ana\\AppData\\Roaming\\org.openstage.guide','a path that is not absolute is not used');
  assert.equal(defaultDataFolder({platform:'darwin',home:'/Users/ana',env:{}}),'/Users/ana/Library/Application Support/org.openstage.guide');
  assert.equal(defaultDataFolder({platform:'linux',home:'/home/ana',env:{}}),'/home/ana/.config/org.openstage.guide');
  assert.equal(defaultDataFolder({platform:'linux',home:'/home/ana',env:{XDG_CONFIG_HOME:'/home/ana/cfg'}}),'/home/ana/cfg/org.openstage.guide');
  assert.notEqual(defaultDataFolder(win),defaultGuidesFolder({platform:'win32',home:win.home,env:{},readRegistry:()=>null}));
  assert.ok(!/OpenStage Guides|Documents/.test(defaultDataFolder(win)));
});

function api(port,name,data){
  return new Promise((resolve,reject)=>{
    const body=JSON.stringify(data);
    const req=http.request({host:'127.0.0.1',port,path:`/api/${name}`,method:'POST',headers:{Host:`127.0.0.1:${port}`,'Content-Type':'application/json','Content-Length':Buffer.byteLength(body)}},res=>{let t='';res.on('data',c=>t+=c);res.on('end',()=>resolve({status:res.statusCode,json:()=>JSON.parse(t)}));});
    req.on('error',reject);req.end(body);
  });
}

test('the server keeps the token in its data folder, moves one left in the guides folder, and refuses a data folder that is the guides folder',async()=>{
  const dir=fs.mkdtempSync(path.join(os.tmpdir(),'osg-prm-link-'));
  const guides=path.join(dir,'Guides'),data=path.join(dir,'Data');
  fs.mkdirSync(guides);
  fs.writeFileSync(path.join(guides,PRM_FILE),JSON.stringify({port:23119,token:'tok-123'}));
  fs.writeFileSync(path.join(guides,'Notes.txt'),'not a guide');
  const prm=standIn();
  try{
    const s=await startProjectServer({folder:guides,port:0,snapshot,dataFolder:data,prmFetch:prm.fetch});
    try{
      assert.equal(s.dataFolder,data);
      assert.equal(fs.existsSync(path.join(guides,PRM_FILE)),false,'the old file is gone from the guides folder');
      assert.deepEqual(JSON.parse(fs.readFileSync(path.join(data,PRM_FILE),'utf8')),{port:23119,token:'tok-123'});
      assert.equal((await api(s.port,'prm',{action:'status'})).json().state,'connected');
      // forget, then connect again from nothing: what is written is written in the data folder
      assert.equal((await api(s.port,'prm',{action:'forget'})).json().state,'not-connected');
      assert.equal(fs.existsSync(path.join(data,PRM_FILE)),false);
      assert.equal((await api(s.port,'prm',{action:'connect',port:23119})).json().state,'pending');
      assert.equal((await api(s.port,'prm',{action:'status'})).json().state,'connected');
      assert.equal(JSON.parse(fs.readFileSync(path.join(data,PRM_FILE),'utf8')).token,'tok-123');
      assert.deepEqual(fs.readdirSync(guides),['Notes.txt'],'the guides folder holds no connection file');
    }finally{await s.close();}
    await assert.rejects(()=>startProjectServer({folder:guides,port:0,snapshot,dataFolder:guides,prmFetch:prm.fetch}),/cannot be the guides folder/);
  }finally{fs.rmSync(dir,{recursive:true,force:true});}
});
