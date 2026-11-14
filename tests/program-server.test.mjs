// The program's server (scripts/project-server.mjs, folder mode): the guides folder, the program's own files, and the tools' page.
import test from 'node:test';
import assert from 'node:assert/strict';
import http from 'node:http';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import {startProjectServer,fileNameFor,PROJECT_EXTENSION} from '../scripts/project-server.mjs';
import {parseProject,serializeProject} from '../src/guide-project.mjs';

const PANE='<!doctype html><html><head><script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>\n<script type="module" src="/build/x/src/app.mjs"></script></head><body>pane</body></html>';
const snapshot={version:'0-test',buildId:'test-build',record:{},read:url=>url==='/taskpane.html'?Buffer.from(PANE):null};
const project=title=>serializeProject({title,paragraphs:[{xml:'',text:title,style:'OSG Guide Title'}]},{savedAt:'2026-11-13T00:00:00.000Z'});

async function withServer(fn,{existing={}}={}){
  const dir=fs.mkdtempSync(path.join(os.tmpdir(),'osg-program-'));
  const folder=path.join(dir,'My guides');
  if(Object.keys(existing).length){fs.mkdirSync(folder);for(const [n,t] of Object.entries(existing))fs.writeFileSync(path.join(folder,n),t);}
  const s=await startProjectServer({folder,port:0,snapshot});
  const call=(method,urlPath,{headers={},body,host=`127.0.0.1:${s.port}`}={})=>new Promise((resolve,reject)=>{
    const req=http.request({host:'127.0.0.1',port:s.port,path:urlPath,method,headers:{Host:host,...headers}},res=>{let d='';res.on('data',c=>d+=c);res.on('end',()=>resolve({status:res.statusCode,text:d,headers:res.headers,json:()=>JSON.parse(d)}));});
    req.on('error',reject);if(body!==undefined)req.write(body);req.end();
  });
  const post=(name,data)=>call('POST',`/api/${name}`,{headers:{'Content-Type':'application/json'},body:JSON.stringify(data??{})});
  const put=body=>call('PUT','/project',{headers:{'Content-Type':'text/plain; charset=utf-8'},body});
  try{await fn({folder,dir,call,post,put,port:s.port});}finally{await s.close();fs.rmSync(dir,{recursive:true,force:true});}
}
const EXT=PROJECT_EXTENSION;

test('a guide\'s name becomes a file name: what Windows does not allow is taken out, and nothing else is changed',()=>{
  assert.equal(fileNameFor('Reframing conflict'),`Reframing conflict${EXT}`);
  assert.equal(fileNameFor('  Team: the "next" step? *  '),`Team the next step${EXT}`);
  assert.equal(fileNameFor('a/b\\c|d'),`a b c d${EXT}`);
  assert.equal(fileNameFor('Ending with dots...'),`Ending with dots${EXT}`);
  assert.equal(fileNameFor('Guía de facilitación'),`Guía de facilitación${EXT}`);
  assert.throws(()=>fileNameFor(''),/Type a name/);
  assert.throws(()=>fileNameFor('  ?*  '),/Type a name/);
  assert.throws(()=>fileNameFor('x'.repeat(81)),/up to 80 letters/);
  for(const bad of ['CON','nul','Com1','LPT9'])assert.throws(()=>fileNameFor(bad),/Windows does not allow/);
  assert.equal(fileNameFor('Console'),`Console${EXT}`);
});

test('with no guides yet the program says so, the folder is made when the first guide is, and a new guide is open and valid',async()=>{
  await withServer(async({folder,call,post,put})=>{
    const state=(await call('GET','/api/state')).json();
    assert.deepEqual([state.current,state.guides,state.folder],[null,[],folder]);
    assert.equal(fs.existsSync(folder),false,'listing does not make the folder');
    assert.equal((await call('GET','/project')).status,409,'no guide open yet');
    assert.equal((await put(project('x'))).status,409);
    const made=await post('new',{name:'Reframing conflict'});
    assert.equal(made.status,200);
    const after=made.json();
    assert.equal(after.current,`Reframing conflict${EXT}`);
    assert.deepEqual(after.guides.map(g=>g.name),[`Reframing conflict${EXT}`]);
    assert.ok(after.guides[0].when.endsWith('Z'));
    assert.deepEqual(fs.readdirSync(folder),[`Reframing conflict${EXT}`]);
    const read=await call('GET','/project');
    assert.equal(read.status,200);
    assert.deepEqual(parseProject(read.text).guide.paragraphs,[],'an empty, valid project');
    assert.equal((await put(project('Reframing conflict'))).status,200);
    assert.equal(parseProject(fs.readFileSync(path.join(folder,`Reframing conflict${EXT}`),'utf8')).guide.title,'Reframing conflict');
  });
});

test('a name that is taken is refused and the guide that is there is untouched; bad names are refused; a name never reaches outside the folder',async()=>{
  await withServer(async({folder,post})=>{
    assert.equal((await post('new',{name:'Mine'})).status,200);
    fs.writeFileSync(path.join(folder,`Mine${EXT}`),project('Precious'));
    const again=await post('new',{name:'Mine'});
    assert.equal(again.status,409);assert.match(again.json().error,/already a guide with that name/);
    assert.equal(parseProject(fs.readFileSync(path.join(folder,`Mine${EXT}`),'utf8')).guide.title,'Precious');
    assert.equal((await post('new',{name:''})).status,400);
    assert.equal((await post('new',{name:'CON'})).status,400);
    assert.equal((await post('new',{})).status,400);
    const sneaky=await post('new',{name:'..\\..\\evil'});
    assert.equal(sneaky.status,200);
    assert.ok(fs.readdirSync(folder).every(n=>n.endsWith(EXT)),'only guide files in the folder');
    assert.deepEqual(fs.readdirSync(path.dirname(folder)).sort(),['My guides'],'nothing was made outside the folder');
  });
});

test('opening a guide reads that guide, closing leaves none open, and a guide that is not in the folder cannot be opened by any name',async()=>{
  await withServer(async({folder,call,post,put})=>{
    await post('new',{name:'One'});await put(project('First guide'));
    await post('new',{name:'Two'});await put(project('Second guide'));
    assert.equal(parseProject((await call('GET','/project')).text).guide.title,'Second guide');
    const opened=await post('open',{name:`One${EXT}`});
    assert.equal(opened.json().current,`One${EXT}`);
    assert.equal(parseProject((await call('GET','/project')).text).guide.title,'First guide','the project is the one that was opened, not the last one made');
    assert.equal((await post('open',{name:'Nope'+EXT})).status,404);
    assert.equal((await post('open',{name:`..${path.sep}x${EXT}`})).status,404);
    assert.equal((await post('open',{name:path.join(folder,`Two${EXT}`)})).status,404,'a full path is not a name');
    assert.equal((await post('open',{})).status,404);
    assert.equal((await post('close')).json().current,null);
    assert.equal((await call('GET','/project')).status,409);
    assert.equal((await post('nothing')).status,404);
  });
});

test('the first save after a guide is opened keeps what was there as .bak; opening it again starts that over',async()=>{
  const first=project('As it was');
  await withServer(async({folder,call,post,put})=>{
    await post('open',{name:`Old${EXT}`});
    await put(project('Changed once'));
    assert.equal(fs.readFileSync(path.join(folder,`Old${EXT}.bak`),'utf8'),first);
    await put(project('Changed twice'));
    assert.equal(fs.readFileSync(path.join(folder,`Old${EXT}.bak`),'utf8'),first);
    await post('close');await post('open',{name:`Old${EXT}`});
    await put(project('Changed again'));
    assert.equal(parseProject(fs.readFileSync(path.join(folder,`Old${EXT}.bak`),'utf8')).guide.title,'Changed twice','the copy taken when it was opened again is of the guide as it was then');
    assert.deepEqual((await call('GET','/api/state')).json().guides.map(g=>g.name),[`Old${EXT}`],'a backup is not listed as a guide');
  },{existing:{[`Old${EXT}`]:first}});
});

test('only this page on this address may make, open or close guides, and only in the form the program sends',async()=>{
  await withServer(async({folder,call,port,post})=>{
    const body=JSON.stringify({name:'Mine'});
    const send=(headers,host)=>call('POST','/api/new',{host,headers:{'Content-Type':'application/json',...headers},body});
    assert.equal((await send({},'evil.example')).status,403,'another host name');
    assert.equal((await send({Origin:'http://evil.example'})).status,403,'another web site');
    assert.equal((await send({'Sec-Fetch-Site':'cross-site'})).status,403,'a cross-site fetch');
    assert.equal((await call('POST','/api/new',{headers:{'Content-Type':'text/plain'},body})).status,415,'a form post');
    assert.equal((await call('POST','/api/new',{headers:{'Content-Type':'application/json'},body:'{not json'})).status,400);
    assert.equal((await call('PUT','/api/new',{headers:{'Content-Type':'application/json'},body})).status,405);
    assert.equal((await call('GET','/api/new')).status,405);
    assert.equal(fs.existsSync(folder),false,'nothing refused made anything');
    assert.equal((await send({Origin:`http://127.0.0.1:${port}`,'Sec-Fetch-Site':'same-origin'})).status,200);
    assert.equal((await post('close')).status,200);
  });
});

test('the program\'s own files are served from program/ only, with a policy that allows nothing from elsewhere',async()=>{
  await withServer(async({call})=>{
    const root=await call('GET','/');
    assert.equal(root.status,302);assert.equal(root.headers.location,'/program/');
    const page=await call('GET','/program/');
    assert.equal(page.status,200);
    assert.match(page.headers['content-type'],/text\/html/);
    assert.match(page.text,/<title>OpenStage Guide<\/title>/);
    const csp=page.headers['content-security-policy'];
    assert.match(csp,/script-src 'self'/);assert.match(csp,/default-src 'self'/);assert.match(csp,/frame-src 'self'/);assert.doesNotMatch(csp,/unsafe-/);
    assert.equal(page.headers['x-content-type-options'],'nosniff');
    for(const [url,type] of [['/program/style.css',/text\/css/],['/program/openstage-ui.css',/text\/css/],['/program/ui/app.js',/javascript/],['/program/ui/commands.js',/javascript/],['/program/pane.css',/text\/css/]]){
      const r=await call('GET',url);
      assert.equal(r.status,200,url);assert.match(r.headers['content-type'],type);assert.ok(r.headers['content-security-policy']);
    }
    for(const bad of ['/program/ui/../scripts/project-server.mjs','/program/..%2Fpackage.json','/program/%2e%2e/package.json','/program/ui/nope.js','/program/x.exe','/program/ui/ui/app.js','/program/ui/sub/app.js','/program/icons.json.bak','/program/ui/','/program/.hidden.js'])
      assert.equal((await call('GET',bad)).status,404,bad);
    assert.equal((await call('POST','/program/')).status,405);
    assert.equal((await call('GET','/program/index.html')).status,200);
  });
});

test('the tools\' page is the pane\'s page without Office.js, which sends a framed page away; everything else in it is the pane\'s',async()=>{
  await withServer(async({call})=>{
    const r=await call('GET','/program/tools?project=1&program=1');
    assert.equal(r.status,200);assert.match(r.headers['content-type'],/text\/html/);
    assert.ok(!/appsforoffice/.test(r.text),'no Office.js');
    assert.ok(r.text.includes('<script type="module" src="/build/x/src/app.mjs"></script>'),'the pane\'s own script is still there');
    assert.ok(!r.headers['content-security-policy'],'the pane is shown as the add-in\'s page always was');
    assert.equal((await call('POST','/program/tools')).status,405);
  });
});

import {spawn} from 'node:child_process';
import {fileURLToPath as toPath} from 'node:url';

test('started by the program\'s window: it names the address it listens on, on a line of its own, and stops when its input closes',async()=>{
  const root=path.resolve(path.dirname(toPath(import.meta.url)),'..');
  const dir=fs.mkdtempSync(path.join(os.tmpdir(),'osg-window-'));
  const child=spawn(process.execPath,['scripts/project-server.mjs','--folder',path.join(dir,'Guides'),'--port','0','--exit-when-stdin-closes'],{cwd:root,stdio:['pipe','pipe','pipe']});
  try{
    const url=await new Promise((resolve,reject)=>{
      let text='';
      child.stdout.on('data',d=>{text+=d;const m=/^OPENSTAGE_READY (http:\/\/127\.0\.0\.1:\d+\/)$/m.exec(text);if(m)resolve(m[1]);});
      child.on('exit',()=>reject(new Error('it stopped before it was ready: '+text)));
      setTimeout(()=>reject(new Error('it was not ready in time')),15000);
    });
    assert.notEqual(new URL(url).port,'0','a port that was chosen for it, not the 0 it was asked for');
    const page=await fetch(url+'program/');
    assert.equal(page.status,200,'and it answers there');
    const exited=new Promise(resolve=>child.on('exit',(code)=>resolve(code)));
    child.stdin.end();
    assert.equal(await Promise.race([exited,new Promise(r=>setTimeout(()=>r('still running'),5000))]),0,'closing its input stops it');
  }finally{child.kill();fs.rmSync(dir,{recursive:true,force:true});}
});

test('the state says which version the program is and the start of its build, for the About window',async()=>{
  await withServer(async({call,post})=>{
    const state=(await call('GET','/api/state')).json();
    assert.equal(state.version,'0-test');
    assert.equal(state.build,'test-bui','the first eight characters of the build identity');
    assert.equal((await post('new',{name:'A guide'})).json().version,'0-test','every answer that gives the state gives it');
  });
});
