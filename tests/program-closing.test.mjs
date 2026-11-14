// Closing the program's window waits for the open guide to be saved: the window asks the page, the page saves and tells the server (/api/closing), and the
// server tells the window on its output (scripts/project-server.mjs, program/ui/app.js, shell/src-tauri/src/main.rs).
import test from 'node:test';
import assert from 'node:assert/strict';
import http from 'node:http';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import {spawn} from 'node:child_process';
import {fileURLToPath} from 'node:url';
import {startProjectServer} from '../scripts/project-server.mjs';
import {serializeProject} from '../src/guide-project.mjs';

const snapshot={version:'0-test',buildId:'test-build',record:{},read:()=>null};
const SERVER=path.resolve(path.dirname(fileURLToPath(import.meta.url)),'..','scripts','project-server.mjs');
const JSON_HEADERS={'Content-Type':'application/json'};

const request=(port,method,urlPath,{headers={},body,host=`127.0.0.1:${port}`}={})=>new Promise((resolve,reject)=>{
  const req=http.request({host:'127.0.0.1',port,path:urlPath,method,headers:{Host:host,...headers}},res=>{let d='';res.on('data',c=>d+=c);res.on('end',()=>resolve({status:res.statusCode,text:d,json:()=>JSON.parse(d)}));});
  req.on('error',reject);if(body!==undefined)req.write(body);req.end();
});

test('the server tells the window the page has saved with the last change on the disk, and says when the page could not save',async()=>{
  const dir=fs.mkdtempSync(path.join(os.tmpdir(),'osg-closing-'));
  const folder=path.join(dir,'Guides');
  const told=[];
  const s=await startProjectServer({folder,port:0,snapshot,onClosing:saved=>told.push([saved,fs.readFileSync(path.join(folder,'G.guide-project.json'),'utf8')])});
  try{
    const made=(await request(s.port,'POST','/api/new',{headers:JSON_HEADERS,body:JSON.stringify({name:'G'})})).json();
    const put=text=>request(s.port,'PUT','/project',{headers:{'Content-Type':'text/plain; charset=utf-8','X-Guide-Token':made.token},body:serializeProject({title:text,paragraphs:[{xml:'',text,style:'Normal'}]})});
    // The page saves, waits for the answer, and then says it is done: what is told to the window is told with the last change on the disk.
    const saved=await put('the last words');
    const answer=await request(s.port,'POST','/api/closing',{headers:JSON_HEADERS,body:JSON.stringify({saved:true})});
    assert.equal(saved.status,200);
    assert.equal(answer.status,200);
    assert.deepEqual(told.map(t=>t[0]),[true]);
    assert.match(told[0][1],/the last words/,'when the window was told, the guide on the disk already held the last change');
    await request(s.port,'POST','/api/closing',{headers:JSON_HEADERS,body:JSON.stringify({saved:false})});
    await request(s.port,'POST','/api/closing',{headers:JSON_HEADERS,body:'{}'});
    assert.deepEqual(told.map(t=>t[0]),[true,false,true],'false only when the page says it could not save');
  }finally{await s.close();fs.rmSync(dir,{recursive:true,force:true});}
});

test('only the program\'s own page, in the form it sends, can say it has saved',async()=>{
  const dir=fs.mkdtempSync(path.join(os.tmpdir(),'osg-closing-'));
  const told=[];
  const s=await startProjectServer({folder:path.join(dir,'Guides'),port:0,snapshot,onClosing:saved=>told.push(saved)});
  try{
    const body=JSON.stringify({saved:true});
    assert.equal((await request(s.port,'POST','/api/closing',{headers:{'Content-Type':'text/plain'},body})).status,415);
    assert.equal((await request(s.port,'POST','/api/closing',{headers:{...JSON_HEADERS,'Sec-Fetch-Site':'cross-site'},body})).status,403);
    assert.equal((await request(s.port,'POST','/api/closing',{headers:{...JSON_HEADERS,Origin:'http://evil.example'},body})).status,403);
    assert.equal((await request(s.port,'POST','/api/closing',{headers:JSON_HEADERS,host:'evil.example',body})).status,403);
    assert.equal((await request(s.port,'GET','/api/closing')).status,405);
    assert.deepEqual(told,[],'none of those told the window anything');
    assert.equal((await request(s.port,'POST','/api/closing',{headers:JSON_HEADERS,body})).status,200);
    assert.deepEqual(told,[true]);
  }finally{await s.close();fs.rmSync(dir,{recursive:true,force:true});}
});

test('started by the window the server says on its output that the page saved, or could not; started any other way it says nothing',async()=>{
  const dir=fs.mkdtempSync(path.join(os.tmpdir(),'osg-closing-'));
  const run=(extra)=>new Promise((resolve,reject)=>{
    const child=spawn(process.execPath,[SERVER,'--folder',path.join(dir,'Guides'),'--port','0','--no-update-check',...extra],{stdio:['pipe','pipe','pipe']});
    let out='';
    child.stdout.on('data',d=>{out+=d;});
    child.stderr.on('data',d=>reject(new Error(String(d))));
    const start=Date.now();
    const wait=async()=>{
      while(!/OpenStage Guide at http:\/\/127\.0\.0\.1:(\d+)\//.test(out)){if(Date.now()-start>15000)throw new Error('the server did not start');await new Promise(r=>setTimeout(r,30));}
      return Number(/127\.0\.0\.1:(\d+)\//.exec(out)[1]);
    };
    resolve({child,wait,lines:()=>out.split(/\r?\n/).filter(Boolean)});
  });
  try{
    const window=await run(['--exit-when-stdin-closes']);
    const port=await window.wait();
    await request(port,'POST','/api/closing',{headers:JSON_HEADERS,body:JSON.stringify({saved:true})});
    await request(port,'POST','/api/closing',{headers:JSON_HEADERS,body:JSON.stringify({saved:false})});
    for(let i=0;i<100&&window.lines().length<4;i++)await new Promise(r=>setTimeout(r,30));
    assert.deepEqual(window.lines().slice(1),[`OPENSTAGE_READY http://127.0.0.1:${port}/`,'OPENSTAGE_CLOSE_READY','OPENSTAGE_CLOSE_HELD']);
    const exited=new Promise(r=>window.child.once('exit',code=>r(code)));
    window.child.stdin.end();
    assert.equal(await exited,0,'closing its input still stops it, with nothing in the middle of being written');

    const browser=await run([]);
    const other=await browser.wait();
    await request(other,'POST','/api/closing',{headers:JSON_HEADERS,body:JSON.stringify({saved:true})});
    await new Promise(r=>setTimeout(r,200));
    assert.equal(browser.lines().length,1,'no line about closing outside the window');
    browser.child.kill();
  }finally{fs.rmSync(dir,{recursive:true,force:true});}
});

test('the program\'s page saves before it tells the server, and the window holds the first request to close for it',()=>{
  const page=fs.readFileSync(new URL('../program/ui/app.js',import.meta.url),'utf8');
  assert.match(page,/Object\.defineProperty\(window, "osgClosing"/,'the page offers one function, which nothing can replace');
  assert.match(page,/saved = await settle\(\);[\s\S]*?post\("closing", \{ saved \}\)/,'it saves first, then says so');
  const rust=fs.readFileSync(new URL('../shell/src-tauri/src/main.rs',import.meta.url),'utf8');
  assert.match(rust,/window\.osgClosing\)/,'the window asks the page by name');
  assert.match(rust,/CloseRequested/);
  assert.match(rust,/api\.prevent_close\(\)/);
});
