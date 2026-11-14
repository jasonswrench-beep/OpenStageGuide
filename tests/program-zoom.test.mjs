// Zoom for the document space (src/zoom.mjs, program/ui, program/style.css, src/app.mjs): how big the guide's page is drawn on the screen, like the zoom of a word
// processor. It scales the guide's page and nothing else, and its level is kept in the program's settings.
import test from 'node:test';
import assert from 'node:assert/strict';
import http from 'node:http';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import {fileURLToPath} from 'node:url';
import {ZOOM_KEY,ZOOM_STEPS,MIN_ZOOM,MAX_ZOOM,cleanZoom,nextZoom,zoomWords,zoomKeyAction} from '../src/zoom.mjs';
import {SETTING_KEYS,SETTINGS_FILE} from '../src/program-settings.mjs';
import {startProjectServer} from '../scripts/project-server.mjs';
import {buildMenus} from '../program/ui/commands.js';
import {KEY_GROUPS} from '../program/ui/keys.js';

const ROOT=path.resolve(path.dirname(fileURLToPath(import.meta.url)),'..');
const read=p=>fs.readFileSync(path.join(ROOT,p),'utf8');
const snapshot={version:'0-test',buildId:'test-build',record:{},read:()=>null};
const request=(port,method,urlPath,{headers={},body}={})=>new Promise((resolve,reject)=>{
  const req=http.request({host:'127.0.0.1',port,path:urlPath,method,headers:{Host:`127.0.0.1:${port}`,...headers}},res=>{let d='';res.on('data',c=>d+=c);res.on('end',()=>resolve({status:res.statusCode,text:d,json:()=>JSON.parse(d)}));});
  req.on('error',reject);if(body!==undefined)req.write(body);req.end();
});

test('the sizes go from half to three times, through 100%, and the next size in either direction stays within them',()=>{
  assert.deepEqual([...ZOOM_STEPS],[...ZOOM_STEPS].sort((a,b)=>a-b),'in order');
  assert.ok(ZOOM_STEPS.includes(1));
  assert.deepEqual([MIN_ZOOM,MAX_ZOOM],[0.5,3]);
  assert.equal(nextZoom(1,1),1.1);
  assert.equal(nextZoom(1,-1),0.9);
  assert.equal(nextZoom(MAX_ZOOM,1),MAX_ZOOM,'no bigger than the biggest');
  assert.equal(nextZoom(MIN_ZOOM,-1),MIN_ZOOM,'no smaller than the smallest');
  assert.equal(nextZoom(1.2,1),1.25,'from a size between two steps, the next step up');
  assert.equal(nextZoom(1.2,-1),1.1,'and the next one down');
  assert.equal(nextZoom('nonsense',1),1.1,'a size that is not one is taken as 100%');
  let z=1;const seen=[z];for(let i=0;i<30;i++){z=nextZoom(z,1);seen.push(z);}
  assert.equal(z,MAX_ZOOM);
  let down=3;for(let i=0;i<30;i++)down=nextZoom(down,-1);
  assert.equal(down,MIN_ZOOM);
  assert.equal(zoomWords(1.25),'Zoom 125%');assert.equal(zoomWords(0.67),'Zoom 67%');
});

test('a size is kept only when it is a number within the limits',()=>{
  assert.equal(cleanZoom(1.25),1.25);assert.equal(cleanZoom('1.5'),1.5);assert.equal(cleanZoom(1.2345),1.23);
  for(const bad of [undefined,null,'',' ','x',NaN,Infinity,0,0.2,3.5,-1,{},[],true])assert.equal(cleanZoom(bad),1,String(bad));
});

test('zoom keys are Ctrl+Alt with =, + , -, or 0, by the character the key makes, and never Ctrl+= (subscript) or an AltGr character',()=>{
  const key=(k,extra={})=>zoomKeyAction({ctrlKey:true,altKey:true,key:k,...extra});
  assert.deepEqual(['=','+','-','_','0'].map(k=>key(k)),['more','more','less','less','actual']);
  assert.equal(zoomKeyAction({ctrlKey:true,key:'='}),null,'Ctrl+= is subscript');
  assert.equal(zoomKeyAction({ctrlKey:true,shiftKey:true,key:'+'}),null,'Ctrl+Shift+= is superscript');
  assert.equal(zoomKeyAction({altKey:true,key:'='}),null);
  assert.equal(key('=',{metaKey:true}),null);
  for(const k of ['}','{','@','€','e','a','1','Tab'])assert.equal(key(k),null,`AltGr or another key: ${k}`);
});

test('the View menu has Zoom in, Zoom out and Actual size, with their keys, and they work only while a guide is open',()=>{
  const calls=[];let can=false;
  const ctx={hasGuide:()=>false,openTool:()=>{},updateGuide:()=>{},newGuide:()=>{},openGuide:()=>{},closeGuide:()=>{},search:()=>{},keys:()=>{},about:()=>{},
    zoom:{more:()=>calls.push('more'),less:()=>calls.push('less'),actual:()=>calls.push('actual'),can:()=>can}};
  const {menus}=buildMenus(ctx);
  assert.deepEqual(menus.map(m=>m.label),['File','Edit','Insert','Format','View','Start','Sections','Activities','Slides','Assessment','eLearning','Check & Plan','Make','Help']);
  const view=menus.find(m=>m.id==='view');
  assert.deepEqual(view.items.map(i=>[i.label,i.keys]),[['Zoom in','Ctrl+Alt+='],['Zoom out','Ctrl+Alt+-'],['Actual size','Ctrl+Alt+0']]);
  assert.ok(view.items.every(i=>i.enabled()===false),'dimmed with no guide open');
  can=true;
  assert.ok(view.items.every(i=>i.enabled()===true),'with a guide open');
  view.items.forEach(i=>i.run());
  assert.deepEqual(calls,['more','less','actual']);
  const without=buildMenus({...ctx,zoom:undefined}).menus.find(m=>m.id==='view');
  assert.ok(without.items.every(i=>i.enabled()===false),'with no zoom given, nothing is offered');
  without.items.forEach(i=>i.run());
  assert.ok(view.items.every(i=>/guide's page/.test(i.title)),'each says it is the guide\'s page that changes');
  const keyRow=KEY_GROUPS.find(g=>g.id==='anywhere').rows.find(r=>/Ctrl\+Alt\+=/.test(r.keys));
  assert.ok(keyRow&&/Ctrl\+Alt\+-/.test(keyRow.keys)&&/Ctrl\+Alt\+0/.test(keyRow.keys)&&/wheel/.test(keyRow.keys),'the key list says them');
});

test('the size is one of the settings the program keeps, so it is there the next time the program starts, on another port',async()=>{
  assert.ok(SETTING_KEYS.includes(ZOOM_KEY));
  const dir=fs.mkdtempSync(path.join(os.tmpdir(),'osg-zoom-'));
  const folder=path.join(dir,'Guides');
  const set=(port,value)=>request(port,'POST','/api/settings',{headers:{'Content-Type':'application/json'},body:JSON.stringify({key:ZOOM_KEY,value})});
  try{
    const a=await startProjectServer({folder,port:0,snapshot});
    try{
      assert.equal((await set(a.port,'1.25')).status,200);
      assert.equal(JSON.parse(fs.readFileSync(path.join(folder,SETTINGS_FILE),'utf8')).values[ZOOM_KEY],'1.25');
    }finally{await a.close();}
    const b=await startProjectServer({folder,port:0,snapshot});
    try{
      assert.equal((await request(b.port,'GET','/api/settings')).json().values[ZOOM_KEY],'1.25');
      assert.equal((await set(b.port,null)).status,200,'100% takes it away');
      assert.deepEqual((await request(b.port,'GET','/api/settings')).json().values,{});
    }finally{await b.close();}
  }finally{fs.rmSync(dir,{recursive:true,force:true});}
});

test('the page zooms the guide\'s page only: a style variable the guide\'s page uses, set by the program\'s page and kept; the tools pass their keys on; the window and the server know nothing of it',()=>{
  const css=read('program/style.css'),page=read('program/ui/app.js'),frame=read('program/ui/frame.js'),tools=read('src/app.mjs');
  assert.match(css,/\.page \{[^}]*zoom: var\(--doc-zoom, 1\);/,'the guide\'s page takes the zoom');
  assert.doesNotMatch(css.replace(/\.page \{[^}]*\}/,''),/--doc-zoom/,'and nothing else does');
  assert.match(page,/setProperty\("--doc-zoom"/);
  assert.match(page,/post\("settings", \{ key: ZOOM_KEY/);
  assert.match(page,/data\.type === "osg-zoom"/);
  assert.match(frame,/zoomKeyAction\(e\)/);
  assert.match(frame,/closest\("#work-guide"\)/,'the wheel zooms only over the guide');
  assert.match(tools,/type:'osg-zoom'/);
  assert.doesNotMatch(tools,/'wheel'/,'the tools\' own wheel is left alone');
  assert.doesNotMatch(read('shell/src-tauri/src/main.rs'),/zoom/i,'the window has no part in it');
  assert.doesNotMatch(read('scripts/project-server.mjs'),/zoom/i,'nor the server, apart from keeping a setting');
});
