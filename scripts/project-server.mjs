// Serves OpenStage Guide as a program in a browser: the program's own page (program/), the pane of the add-in for the tools, and the guides
// kept in one folder on this computer.
//
//   node scripts/project-server.mjs --folder "C:\Users\me\Documents\OpenStage Guides" [--port 3850]    the program: http://127.0.0.1:3850/
//   node scripts/project-server.mjs --file "C:\path\My guide.guide-project.json" [--port 3850]          the pane on one project: /taskpane.html?project=1
//
// The page asks this server to list, make, open and close guides (/api/...) and to read and store the one that is open (GET and PUT /project).
// Nothing else on the disk is reachable, and a guide is never named by a path: the page says a file's name and the server finds it in the
// folder. The server listens on this computer only and answers only to its own address, so a web page from elsewhere cannot read or change a guide.
//
// Saving is careful: every project is checked before it is stored, it is written to a temporary file and moved into place (a crash leaves the
// old file whole), and the project as it was found when it was opened is kept as a dated copy ("<file>.<date>-<nn>.bak", the newest three).
import {createBuildSnapshot} from './build-snapshot.mjs';
import {createHandler} from './http-handler.mjs';
import {PROGRAM_CSP,PANE_CSP,pageHeaders,stripOfficeScript} from './security-headers.mjs';
import {checkForUpdates,fetchReleases,versionKey} from './updates.mjs';
import {parseProject,serializeProject,cleanGuide,LIMITS} from '../src/guide-project.mjs';
import http from 'node:http';
import crypto from 'node:crypto';
import fs from 'node:fs/promises';
import {constants as fsConstants} from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import {fileURLToPath} from 'node:url';

const root=path.resolve(path.dirname(fileURLToPath(import.meta.url)),'..');
// What a guide's file is called. The ending is a decision for the author (see docs/STANDALONE-APP-PLAN-2026-11-13.md); change it here only.
export const PROJECT_EXTENSION='.guide-project.json';
const TYPES={'.html':'text/html; charset=utf-8','.js':'text/javascript; charset=utf-8','.css':'text/css; charset=utf-8','.svg':'image/svg+xml','.png':'image/png','.json':'application/json','.ttf':'font/ttf','.txt':'text/plain; charset=utf-8','.md':'text/plain; charset=utf-8'};
const PLAIN=/^[A-Za-z0-9_-]+\.(html|css|js|svg|png|json)$/;
// The fonts that come with the program, and their licenses (program/fonts): plain names only, like the rest.
const FONT_FILE=/^[A-Za-z0-9_-]+\.(ttf|txt|md)$/;
// Names Windows keeps for devices, with or without an extension (CON, con.txt, COM1 and the superscript digits ...). Neither a guide nor a program file may have one.
const RESERVED=/^(con|prn|aux|nul|com[0-9\u00b9\u00b2\u00b3]|lpt[0-9\u00b9\u00b2\u00b3]|conin\$|conout\$)$/i;
export const reservedName=name=>RESERVED.test(String(name??'').split('.')[0].trim());
// What an /api request may weigh (a name and little else); only a guide's own text may be large.
const API_MAX_BYTES=16*1024;
// How many dated copies of a guide are kept from before its first save in a session.
const BACKUPS_KEPT=3;

// A guide's name made into a file name: what Windows does not allow is taken out, and a name that is nothing but that is refused.
export function fileNameFor(raw){
  const name=String(raw??'').replace(/[<>:"/\\|?*\u0000-\u001f]/g,' ').replace(/\s+/g,' ').trim().replace(/[. ]+$/,'');
  if(!name)throw new Error('Type a name for the guide.');
  if(name.length>80)throw new Error('A guide\'s name can be up to 80 letters.');
  if(reservedName(name))throw new Error('Windows does not allow that name for a file. Choose another.');
  return name+PROJECT_EXTENSION;
}

// `release` is the release this copy is (a tag such as v0.1.0-alpha.1; the window gives it from the installer's build), `updateCheck` false turns
// Check for updates off, and `releases` is how GitHub's list is fetched (replaced in tests, which never reach the network). See scripts/updates.mjs.
export async function startProjectServer({file,folder,port=3850,snapshot,maxBytes=LIMITS.fileCharacters*4,release,updateCheck=true,releases=fetchReleases}={}){
  if(!file&&!folder)throw new Error('Name the guides folder with --folder, or one project file with --file.');
  const dir=path.resolve(folder||path.dirname(file));
  let current=file?path.basename(file):null,backedUp=false,token=null;
  const snap=snapshot||await createBuildSnapshot(root);
  const handler=createHandler(snap);
  const target=()=>current?path.join(dir,current):null;
  const send=(res,code,body='',type='text/plain; charset=utf-8',extra={})=>{res.writeHead(code,{'Content-Type':type,'Cache-Control':'no-store','X-Content-Type-Options':'nosniff','Referrer-Policy':'no-referrer',...extra});res.end(body);};
  const json=(res,code,data)=>send(res,code,JSON.stringify(data),'application/json; charset=utf-8');
  const own=()=>new Set([`127.0.0.1:${server.address().port}`,`localhost:${server.address().port}`]);

  async function readBody(req,limit=maxBytes){
    const chunks=[];let size=0;
    for await(const c of req){size+=c.length;if(size>limit)return null;chunks.push(c);}
    return Buffer.concat(chunks).toString('utf8');
  }
  // What a request must be to change anything: this page, on this address, in the form the program sends.
  function allowed(req,res,types){
    const site=req.headers['sec-fetch-site'];
    if(site&&site!=='same-origin'&&site!=='none'){send(res,403,'Not allowed from another site.');return false;}
    if(req.headers.origin&&!own().has(new URL(req.headers.origin).host)){send(res,403,'Not allowed from another site.');return false;}
    if(!types.test(req.headers['content-type']||'')){send(res,415,'That is not the form the program sends.');return false;}
    return true;
  }
  async function guides(){
    let names=[];
    try{names=(await fs.readdir(dir,{withFileTypes:true})).filter(e=>e.isFile()&&e.name.endsWith(PROJECT_EXTENSION)).map(e=>e.name);}catch(e){if(e.code!=='ENOENT')throw e;}
    const out=[];
    for(const name of names){try{const st=await fs.stat(path.join(dir,name));out.push({name,when:st.mtime.toISOString()});}catch{/* gone while listing */}}
    return out.sort((a,b)=>b.when.localeCompare(a.when));
  }
  // `version` and `build` are what the About window shows: the program's own version and the start of the build's identity (build-snapshot.mjs).
  // `folder` is the guides folder's name only (the page says where guides are kept without learning the person's folders). `token` names the guide that
  // is open: see tokenOk.
  const stateNow=async()=>({current,folder:path.basename(dir),guides:await guides(),version:snap.version,build:String(snap.buildId||'').slice(0,8),token});

  // Which guide a page may read and store: the one it was given the token for when it opened or made it. A page left over from another tab or
  // window holds an older token, so it is refused rather than writing its guide over the one that is open now. (Before any guide has been opened
  // or made in this run, as with --file, there is no token to ask for.)
  const newToken=()=>crypto.randomBytes(18).toString('base64url');
  const tokenOk=req=>{
    if(token===null)return true;
    const sent=Buffer.from(String(req.headers['x-guide-token']||'')),want=Buffer.from(token);
    return sent.length===want.length&&crypto.timingSafeEqual(sent,want);
  };
  const STALE='This guide was closed or changed in another window, so this page no longer saves it. Reload the page to go on.';

  // Saves run one at a time, in the order they came, so two at once cannot write the same file.
  let saving=Promise.resolve();
  const inTurn=job=>{const run=saving.then(job);saving=run.then(()=>{},()=>{});return run;};

  // Before the first save since a guide was opened: keep a dated copy of what is there, and only the newest few. The guide counts as backed up
  // only when the copy was made (or there was nothing to copy).
  async function backUp(file){
    const stamp=new Date().toISOString().replace(/[-:]/g,'').replace(/\.\d+Z$/,'Z');
    const folder=path.dirname(file),base=path.basename(file);
    const COPY=/^\d{8}T\d{6}Z-\d{2}\.bak$/;
    const listing=async()=>(await fs.readdir(folder)).filter(name=>name.startsWith(base+'.')&&COPY.test(name.slice(base.length+1))).map(name=>name.slice(base.length+1));
    // (the two digits make copies taken in the same second sort in the order they were taken: the next one is one more than the highest there is)
    const same=(await listing()).filter(name=>name.startsWith(stamp+'-')).map(name=>Number(name.slice(-6,-4)));
    let made=false;
    for(let n=Math.max(-1,...same)+1;n<100&&!made;n++){
      try{await fs.copyFile(file,`${file}.${stamp}-${String(n).padStart(2,'0')}.bak`,fsConstants.COPYFILE_EXCL);made=true;}
      catch(e){
        if(e.code==='ENOENT'){backedUp=true;return;}
        if(e.code!=='EEXIST')throw e;
      }
    }
    if(!made)throw new Error('A copy of the guide could not be kept.');
    backedUp=true;
    const copies=(await listing()).sort().reverse();
    for(const old of copies.slice(BACKUPS_KEPT))await fs.rm(path.join(folder,`${base}.${old}`),{force:true});
  }

  async function store(file,body){
    await fs.mkdir(path.dirname(file),{recursive:true});
    if(!backedUp)await backUp(file);
    // A temporary name nobody can guess, opened only if new (a link left at the name is not followed), then moved into place.
    const temp=path.join(path.dirname(file),`.${path.basename(file)}.${crypto.randomBytes(6).toString('hex')}.tmp`);
    try{
      await fs.writeFile(temp,body,{encoding:'utf8',flag:'wx'});
      await fs.rename(temp,file);
    }catch(e){
      // A failed move leaves the old project whole; do not leave the half-way copy lying about.
      await fs.rm(temp,{force:true});
      throw e;
    }
  }

  async function project(req,res){
    if(req.method==='GET'||req.method==='HEAD'){
      if(!current){json(res,409,{error:'No guide is open.'});return;}
      if(!tokenOk(req)){json(res,409,{error:STALE,code:'stale-guide'});return;}
      let text;
      try{text=await fs.readFile(target(),'utf8');}catch(e){if(e.code==='ENOENT'){send(res,204);return;}throw e;}
      send(res,200,req.method==='HEAD'?'':text);
      return;
    }
    if(req.method!=='PUT'){send(res,405,'Method not allowed');return;}
    if(!allowed(req,res,/^(text\/plain|application\/json)\b/i))return;
    if(!current){json(res,409,{error:'No guide is open.'});return;}
    if(!tokenOk(req)){json(res,409,{error:STALE,code:'stale-guide'});return;}
    const body=await readBody(req);
    if(body===null){send(res,413,'That project is larger than this program can store.');return;}
    try{parseProject(body);}catch(e){send(res,400,e.message);return;}
    const file=target(),mine=token;
    // Checked again when its turn comes: another guide may have been opened while this save waited.
    const saved=await inTurn(async()=>{
      if(token!==mine||target()!==file)return false;
      await store(file,body);
      return true;
    });
    if(!saved){json(res,409,{error:STALE,code:'stale-guide'});return;}
    send(res,200,'saved');
  }

  async function api(req,res,name){
    if(req.method==='GET'&&name==='state'){json(res,200,await stateNow());return;}
    if(req.method!=='POST'){send(res,405,'Method not allowed');return;}
    if(!allowed(req,res,/^application\/json\b/i))return;
    const raw=await readBody(req,API_MAX_BYTES);
    if(raw===null){json(res,413,{error:'That request is larger than this program takes.'});return;}
    let body={};
    try{body=raw?JSON.parse(raw):{};}catch{json(res,400,{error:'That request was not readable.'});return;}
    if(name==='new'){
      let file;
      try{file=fileNameFor(body.name);}catch(e){json(res,400,{error:e.message});return;}
      await fs.mkdir(dir,{recursive:true});
      try{await fs.writeFile(path.join(dir,file),serializeProject(cleanGuide({paragraphs:[]})),{encoding:'utf8',flag:'wx'});}
      catch(e){if(e.code==='EEXIST'){json(res,409,{error:'There is already a guide with that name. Choose another name, or open it.'});return;}throw e;}
      current=file;backedUp=false;token=newToken();
      json(res,200,await stateNow());return;
    }
    if(name==='open'){
      // The name must be one the folder lists: a request never names a path.
      const known=(await guides()).find(g=>g.name===body.name);
      if(!known){json(res,404,{error:'That guide is not in the guides folder.'});return;}
      current=known.name;backedUp=false;token=newToken();
      json(res,200,await stateNow());return;
    }
    if(name==='close'){current=null;token=null;json(res,200,await stateNow());return;}
    // Check for updates: only when the page asks (a POST, in the form the program sends, like every action), and it asks GitHub for nothing about the person.
    if(name==='updates'){json(res,200,await checkForUpdates({current:versionKey(release||'')?release:undefined,enabled:updateCheck!==false,fetchReleases:releases}));return;}
    json(res,404,{error:'Not found.'});
  }

  // The tools of the program are the add-in's pane. Office.js is the add-in's way of talking to Word, and it sends a page that is shown in a
  // frame away to a blank one, so the program's copy of the pane page leaves its script out; the pane itself is the same page.
  function tools(req,res){
    const html=snap.read('/taskpane.html');
    if(!html){send(res,404,'Not found');return;}
    send(res,200,req.method==='HEAD'?'':stripOfficeScript(html),'text/html; charset=utf-8',pageHeaders(PANE_CSP));
  }

  async function program(req,res,pathname){
    if(req.method!=='GET'&&req.method!=='HEAD'){send(res,405,'Method not allowed');return;}
    if(pathname==='/program/tools'){tools(req,res);return;}
    const rest=pathname.slice('/program/'.length)||'index.html';
    const parts=rest.split('/');
    const plain=((parts.length===1||(parts.length===2&&parts[0]==='ui'))&&parts.every((p,i)=>i===parts.length-1?PLAIN.test(p):p==='ui'))||(parts.length===2&&parts[0]==='fonts'&&FONT_FILE.test(parts[1]));
    if(!plain||parts.some(reservedName)){send(res,404,'Not found');return;}
    let data;
    try{data=await fs.readFile(path.join(root,'program',...parts));}catch{send(res,404,'Not found');return;}
    send(res,200,req.method==='HEAD'?'':data,TYPES[path.extname(rest)]||'application/octet-stream',pageHeaders(PROGRAM_CSP));
  }

  const server=http.createServer(async(req,res)=>{
    try{
      if(!own().has(req.headers.host||'')){send(res,403,'This address is not allowed.');return;}
      const url=new URL(req.url,'http://localhost');
      if(url.pathname==='/project'){await project(req,res);return;}
      const a=/^\/api\/([a-z]+)$/.exec(url.pathname);
      if(a){await api(req,res,a[1]);return;}
      if(url.pathname==='/'||url.pathname==='/program'){send(res,302,'','text/plain',{Location:'/program/'});return;}
      if(url.pathname.startsWith('/program/')){await program(req,res,url.pathname);return;}
      await handler(req,res);
    }catch(e){
      send(res,500,'The program could not do that.');
    }
  });
  // A request that is slow to arrive is dropped, so a client cannot hold a connection open with a body half sent.
  server.requestTimeout=60000;
  server.headersTimeout=15000;
  await new Promise((resolve,reject)=>{server.once('error',reject);server.listen(port,'127.0.0.1',resolve);});
  return {server,port:server.address().port,folder:dir,file:target(),close:()=>new Promise(r=>server.close(r))};
}

// Run from the command line.
if(process.argv[1]&&path.resolve(process.argv[1])===fileURLToPath(import.meta.url)){
  const arg=name=>{const i=process.argv.indexOf(name);return i>=0?process.argv[i+1]:undefined;};
  try{
    const folder=arg('--folder'),file=arg('--file');
    const updateOff=process.argv.includes('--no-update-check')||String(process.env.OPENSTAGE_UPDATE_CHECK||'').trim().toLowerCase()==='off';
    const s=await startProjectServer({file,folder:folder||(file?undefined:path.join(os.homedir(),'Documents','OpenStage Guides')),port:Number(arg('--port')||3850),release:arg('--release'),updateCheck:!updateOff});
    console.log(`OpenStage Guide at http://127.0.0.1:${s.port}/  (guides folder: ${s.folder})`);
    // The program's window (shell/) starts this server, reads this line to learn where it is listening, and stops it by closing its input:
    // when the window goes, with it or without it, the server goes too.
    if(process.argv.includes('--exit-when-stdin-closes')){
      console.log(`OPENSTAGE_READY http://127.0.0.1:${s.port}/`);
      process.stdin.resume();
      process.stdin.on('end',()=>{s.close().finally(()=>process.exit(0));});
      process.stdin.on('error',()=>process.exit(0));
    }
  }catch(e){console.error(e.message);process.exitCode=1;}
}
