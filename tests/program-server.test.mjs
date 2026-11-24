// The program's server (scripts/project-server.mjs, folder mode): the guides folder, the program's own files, and the tools' page.
import test from 'node:test';
import assert from 'node:assert/strict';
import http from 'node:http';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import {startProjectServer,fileNameFor,reservedName,PROJECT_EXTENSION} from '../scripts/project-server.mjs';
import {parseProject,serializeProject} from '../src/guide-project.mjs';

const PANE='<!doctype html><html><head><script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>\n<script type="module" src="/build/x/src/app.mjs"></script></head><body>pane</body></html>';
const HELP='<!doctype html><html><head><style>b{color:red}</style></head><body>help</body></html>';
const snapshot={version:'0-test',buildId:'test-build',record:{},read:url=>url==='/taskpane.html'?Buffer.from(PANE):url==='/help.html'?Buffer.from(HELP):null};
const project=title=>serializeProject({title,paragraphs:[{xml:'',text:title,style:'OSG Guide Title'}]},{savedAt:'2026-11-13T00:00:00.000Z'});

async function withServer(fn,{existing={},start={}}={}){
  const dir=fs.mkdtempSync(path.join(os.tmpdir(),'osg-program-'));
  const folder=path.join(dir,'My guides');
  if(Object.keys(existing).length){fs.mkdirSync(folder);for(const [n,t] of Object.entries(existing))fs.writeFileSync(path.join(folder,n),t);}
  // (Its own data folder: a test never touches the real one.)
  const s=await startProjectServer({folder,port:0,snapshot,dataFolder:path.join(dir,'Data'),...start});
  // The page of the program holds the token the program gave it when it opened or made a guide, and sends it with every read and save of the
  // guide. The helper does the same unless a test says otherwise (an explicit X-Guide-Token header, or none, wins).
  let token=null;
  const call=(method,urlPath,{headers={},body,host=`127.0.0.1:${s.port}`}={})=>new Promise((resolve,reject)=>{
    const own=token&&urlPath==='/project'?{'X-Guide-Token':token}:{};
    const req=http.request({host:'127.0.0.1',port:s.port,path:urlPath,method,headers:{Host:host,...own,...headers}},res=>{let d='';res.on('data',c=>d+=c);res.on('end',()=>resolve({status:res.statusCode,text:d,headers:res.headers,json:()=>JSON.parse(d)}));});
    req.on('error',reject);if(body!==undefined)req.write(body);req.end();
  });
  const post=async(name,data)=>{
    const r=await call('POST',`/api/${name}`,{headers:{'Content-Type':'application/json'},body:JSON.stringify(data??{})});
    if(r.status===200){try{const j=r.json();if('token' in j)token=j.token;}catch{/* not json */}}
    return r;
  };
  const put=(body,headers={})=>call('PUT','/project',{headers:{'Content-Type':'text/plain; charset=utf-8',...headers},body});
  try{await fn({folder,dir,call,post,put,port:s.port,token:()=>token,server:s.server});}finally{await s.close();fs.rmSync(dir,{recursive:true,force:true});}
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
    assert.deepEqual([state.current,state.guides,state.folder],[null,[],path.basename(folder)],'the folder is named, never given as a path');
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

// The dated copies kept of a guide: "<file>.<date>T<time>Z-<nn>.bak", oldest first.
const copiesOf=(folder,name)=>fs.readdirSync(folder).filter(n=>n.startsWith(name+'.')&&/\.\d{8}T\d{6}Z-\d{2}\.bak$/.test(n)).sort();

test('the first save after a guide is opened keeps what was there as a dated copy; opening it again takes another',async()=>{
  const first=project('As it was');
  await withServer(async({folder,call,post,put})=>{
    await post('open',{name:`Old${EXT}`});
    await put(project('Changed once'));
    let copies=copiesOf(folder,`Old${EXT}`);
    assert.equal(copies.length,1);
    assert.equal(fs.readFileSync(path.join(folder,copies[0]),'utf8'),first);
    await put(project('Changed twice'));
    assert.deepEqual(copiesOf(folder,`Old${EXT}`),copies,'later saves in the same session take no more');
    await post('close');await post('open',{name:`Old${EXT}`});
    await put(project('Changed again'));
    copies=copiesOf(folder,`Old${EXT}`);
    assert.equal(copies.length,2);
    assert.equal(parseProject(fs.readFileSync(path.join(folder,copies[1]),'utf8')).guide.title,'Changed twice','the copy taken when it was opened again is of the guide as it was then');
    assert.deepEqual((await call('GET','/api/state')).json().guides.map(g=>g.name),[`Old${EXT}`],'a backup is not listed as a guide');
  },{existing:{[`Old${EXT}`]:first}});
});

test('only the newest three dated copies are kept',async()=>{
  await withServer(async({folder,post,put})=>{
    for(let i=1;i<=6;i++){
      await post('open',{name:`Old${EXT}`});
      await put(project(`Version ${i}`));
    }
    const copies=copiesOf(folder,`Old${EXT}`);
    assert.equal(copies.length,3);
    assert.deepEqual(copies.map(n=>parseProject(fs.readFileSync(path.join(folder,n),'utf8')).guide.title),['Version 3','Version 4','Version 5'],'the three before the last save, oldest first');
    assert.deepEqual(fs.readdirSync(folder).filter(n=>/\.tmp$/.test(n)),[],'no temporary file is left');
  },{existing:{[`Old${EXT}`]:project('Version 0')}});
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
    assert.match(r.headers['content-security-policy'],/frame-ancestors 'self'/,'the pane may be framed only by a page of this server');
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

// ---- the independent security review of 2026-11-13

test('a save is for the guide its page was given the token for: a page left over from another tab is refused, and the guide that is open is not touched',async()=>{
  await withServer(async({folder,call,post,put})=>{
    const x=(await post('new',{name:'X'})).json();
    assert.match(x.token,/^[A-Za-z0-9_-]{20,}$/,'opening or making a guide gives a token');
    assert.equal((await put(project('X one'))).status,200);
    const y=(await post('new',{name:'Y'})).json();
    assert.notEqual(y.token,x.token);
    const stale=await put(project('X after Y was opened'),{'X-Guide-Token':x.token});
    assert.equal(stale.status,409);
    assert.equal(stale.json().code,'stale-guide');
    assert.match(stale.json().error,/closed or changed in another window/);
    assert.equal(parseProject(fs.readFileSync(path.join(folder,`Y${EXT}`),'utf8')).guide.paragraphs.length,0,'Y is as it was made');
    assert.equal(parseProject(fs.readFileSync(path.join(folder,`X${EXT}`),'utf8')).guide.title,'X one','and X was not written to either');
    assert.equal((await call('GET','/project',{headers:{'X-Guide-Token':x.token}})).status,409,'a page that is out of date cannot read the open guide either');
    assert.equal((await put(project('no token'),{'X-Guide-Token':''})).status,409,'a save with no token is refused');
    assert.equal((await put(project('Y one'))).status,200,'the page with the right token saves');
    assert.equal(parseProject(fs.readFileSync(path.join(folder,`Y${EXT}`),'utf8')).guide.title,'Y one');
    assert.equal((await post('close')).json().token,null,'closing the guide ends its token');
    assert.equal((await call('GET','/api/state')).json().token,null);
  });
});

test('a save that was waiting when another guide was opened is not written to that guide',async()=>{
  await withServer(async({folder,post,put})=>{
    const x=(await post('new',{name:'X'})).json();
    const slow=put(project('X late'),{'X-Guide-Token':x.token});
    const y=post('new',{name:'Y'});
    const [a]=await Promise.all([slow,y]);
    // Either the save went through before Y was made (X holds it) or it was refused; it never lands on Y.
    assert.ok([200,409].includes(a.status));
    assert.equal(parseProject(fs.readFileSync(path.join(folder,`Y${EXT}`),'utf8')).guide.paragraphs.length,0,'Y is untouched');
  });
});

test('saves at once are saved one at a time: the file ends as exactly one of them, whole, and no temporary file is left',async()=>{
  await withServer(async({folder,post,put})=>{
    await post('new',{name:'Race'});
    const bodies=Array.from({length:8},(_,i)=>project(`Version ${i}`));
    const answers=await Promise.all(bodies.map(b=>put(b)));
    assert.ok(answers.every(a=>a.status===200));
    const onDisk=fs.readFileSync(path.join(folder,`Race${EXT}`),'utf8');
    assert.ok(bodies.includes(onDisk),'the file is exactly one of the guides that were sent');
    assert.deepEqual(fs.readdirSync(folder).filter(n=>/\.tmp$/.test(n)),[],'no temporary file');
  });
});

test('a temporary name cannot be guessed, and a link left at the old name is not followed',async()=>{
  await withServer(async({folder,post,put})=>{
    await post('new',{name:'Safe'});
    const old=path.join(folder,`Safe${EXT}.${process.pid}.tmp`);
    fs.writeFileSync(old,'PLANTED');
    assert.equal((await put(project('Saved'))).status,200);
    assert.equal(fs.readFileSync(old,'utf8'),'PLANTED','a file at the old predictable temporary name is left alone, and not written through');
    assert.equal(parseProject(fs.readFileSync(path.join(folder,`Safe${EXT}`),'utf8')).guide.title,'Saved');
  });
});

test('an /api request over 16 KB is refused, a guide far larger than that is still saved, and slow requests are not waited for forever',async()=>{
  await withServer(async({folder,post,put,server})=>{
    const r=await post('new',{name:'x'.repeat(20*1024)});
    assert.equal(r.status,413);
    await post('new',{name:'Big'});
    const big=serializeProject({title:'Big',paragraphs:Array.from({length:120},(_,i)=>({xml:'',text:`Paragraph ${i} `+'word '.repeat(300),style:'Normal'}))});
    assert.ok(big.length>150*1024,'the guide is well over 16 KB');
    assert.equal((await put(big)).status,200);
    assert.equal(fs.readFileSync(path.join(folder,`Big${EXT}`),'utf8'),big);
    assert.equal(server.requestTimeout,60000);
    assert.equal(server.headersTimeout,15000);
  });
});

test('every Windows device name is refused for a guide, with or without an extension, and for a program file',async()=>{
  for(const bad of ['CON','con.txt','Nul.draft','COM1','lpt9.x','COM\u00b9','LPT\u00b2.doc','CONIN$','CONOUT$.log','Aux','prn','  con  '])assert.throws(()=>fileNameFor(bad),/Windows does not allow/,bad);
  for(const fine of ['Console','Common','conx','Aux1','Computer','My con','Lesson 1'])assert.doesNotThrow(()=>fileNameFor(fine),fine);
  assert.equal(reservedName('con.js'),true);assert.equal(reservedName('app.js'),false);
  await withServer(async({call})=>{
    for(const p of ['/program/con.js','/program/ui/nul.js','/program/aux.css','/program/CONIN$.js','/program/fonts/com1.ttf'])assert.equal((await call('GET',p)).status,404,p);
  });
});

test('every HTML page says what it may load and that only a page of this server may frame it, and none loads Office.js',async()=>{
  await withServer(async({call})=>{
    for(const p of ['/program/','/program/tools','/taskpane.html','/help.html','/build/test-build/taskpane.html']){
      const r=await call('GET',p);
      assert.equal(r.status,200,p);
      const csp=r.headers['content-security-policy']||'';
      assert.match(csp,/default-src 'self'/,p);
      assert.match(csp,/script-src 'self';/,`${p}: scripts only from this server, none inline`);
      assert.match(csp,/frame-ancestors 'self'/,p);
      assert.match(csp,/object-src 'none'/,p);
      assert.match(csp,/base-uri 'none'/,p);
      assert.equal(r.headers['x-frame-options'],'SAMEORIGIN',p);
      assert.doesNotMatch(r.text,/appsforoffice/,`${p} does not name Office.js`);
    }
    const style=(await call('GET','/program/')).headers['content-security-policy'];
    assert.match(style,/style-src 'self';/,'the program\'s own page allows no inline style');
    assert.equal((await call('GET','/program/ui/app.js')).headers['x-frame-options'],'SAMEORIGIN','and its files carry the guard too');
    assert.equal((await call('GET','/health')).headers['x-frame-options'],undefined,'JSON answers need no frame guard');
  });
});

test('the state names the guides folder and does not give its path',async()=>{
  await withServer(async({folder,call,post})=>{
    await post('new',{name:'A'});
    const state=(await call('GET','/api/state')).json();
    assert.equal(state.folder,path.basename(folder));
    assert.ok(!JSON.stringify(state).includes(path.dirname(folder)),'no folder above it is given');
  });
});

// ---- Check for updates

const RELEASES=JSON.stringify([{tag_name:'v0.1.0-alpha.1'},{tag_name:'v0.1.0-alpha.2'}]);

test('Check for updates asks GitHub only when asked, answers a POST in the form the program sends and nothing else, and compares with the release this copy is',async()=>{
  let asked=0;
  const releases=async()=>{asked++;return RELEASES;};
  await withServer(async({call,post})=>{
    assert.equal(asked,0,'nothing is asked at start-up');
    await call('GET','/api/state');await post('new',{name:'A'});
    assert.equal(asked,0,'nor by anything else the program does');
    const r=await post('updates',{});
    assert.equal(r.status,200);
    const answer=r.json();
    assert.deepEqual([answer.state,answer.latest,answer.url],['newer','0.1.0 alpha 2','https://github.com/jasonswrench-beep/OpenStageGuide/releases/tag/v0.1.0-alpha.2']);
    assert.match(answer.message,/This copy is 0\.1\.0 alpha 1/);
    assert.equal(asked,1);
    await post('updates',{});
    assert.equal(asked,2,'each press asks again, and nothing is kept');
    assert.equal((await call('GET','/api/updates')).status,405,'a GET is not an action');
    assert.equal((await call('POST','/api/updates',{headers:{'Content-Type':'text/plain'},body:'{}'})).status,415,'not in the form the program sends');
    assert.equal((await call('POST','/api/updates',{headers:{'Content-Type':'application/json','Sec-Fetch-Site':'cross-site'},body:'{}'})).status,403,'not from another site');
    assert.equal((await call('POST','/api/updates',{headers:{'Content-Type':'application/json'},host:'evil.example',body:'{}'})).status,403,'not under another host name');
    assert.equal(asked,2,'and a refused request asks nothing');
  },{start:{release:'v0.1.0-alpha.1',releases}});
});

test('a copy that is the newest release says so, and one with no release recorded (or a made-up one) is told the newest',async()=>{
  const releases=async()=>RELEASES;
  await withServer(async({post})=>{assert.equal((await post('updates',{})).json().state,'current');},{start:{release:'v0.1.0-alpha.2',releases}});
  await withServer(async({post})=>{assert.equal((await post('updates',{})).json().state,'unknown');},{start:{releases}});
  await withServer(async({post})=>{assert.equal((await post('updates',{})).json().state,'unknown','a release that is not a version is no release');},{start:{release:'v1; rm -rf',releases}});
});

test('Check for updates turned off asks nothing of the network and says so; a failure to reach GitHub is an answer, not an error',async()=>{
  let asked=0;
  await withServer(async({post})=>{
    const r=await post('updates',{});
    assert.equal(r.status,200);
    assert.deepEqual(r.json(),{state:'off',message:'Checking for updates is turned off on this computer.'});
    assert.equal(asked,0);
  },{start:{updateCheck:false,release:'v0.1.0-alpha.1',releases:async()=>{asked++;return RELEASES;}}});
  await withServer(async({post})=>{
    const r=await post('updates',{});
    assert.equal(r.status,200,'the server does not fail because GitHub could not be reached');
    assert.equal(r.json().state,'error');
    assert.match(r.json().message,/could not reach GitHub/);
  },{start:{releases:async()=>{throw new Error('offline');}}});
});

test('started from the command line: --no-update-check and OPENSTAGE_UPDATE_CHECK=off turn the check off, and --release is the release this copy is',async()=>{
  const root=path.resolve(path.dirname(toPath(import.meta.url)),'..');
  const run=async(args,env={})=>{
    const dir=fs.mkdtempSync(path.join(os.tmpdir(),'osg-updates-'));
    const child=spawn(process.execPath,['scripts/project-server.mjs','--folder',path.join(dir,'Guides'),'--port','0','--exit-when-stdin-closes',...args],{cwd:root,stdio:['pipe','pipe','pipe'],env:{...process.env,OPENSTAGE_UPDATE_CHECK:'',...env}});
    try{
      const url=await new Promise((resolve,reject)=>{let text='';child.stdout.on('data',d=>{text+=d;const m=/OPENSTAGE_READY (http:\/\/127\.0\.0\.1:\d+\/)/.exec(text);if(m)resolve(m[1]);});child.on('error',reject);setTimeout(()=>reject(new Error('the server did not say where it is')),15000);});
      const answer=await fetch(new URL('/api/updates',url),{method:'POST',headers:{'Content-Type':'application/json'},body:'{}'});
      return answer.json();
    }finally{child.stdin.end();child.kill();fs.rmSync(dir,{recursive:true,force:true});}
  };
  assert.equal((await run(['--no-update-check','--release','v0.1.0-alpha.1'])).state,'off');
  assert.equal((await run(['--release','v0.1.0-alpha.1'],{OPENSTAGE_UPDATE_CHECK:'off'})).state,'off');
  assert.equal((await run(['--release','v0.1.0-alpha.1'],{OPENSTAGE_UPDATE_CHECK:' OFF '})).state,'off');
});
