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
// old file whole), and the project as it was found when it was opened is kept once as "<file>.bak".
import {createBuildSnapshot} from './build-snapshot.mjs';
import {createHandler} from './http-handler.mjs';
import {parseProject,serializeProject,cleanGuide,LIMITS} from '../src/guide-project.mjs';
import http from 'node:http';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import {fileURLToPath} from 'node:url';

const root=path.resolve(path.dirname(fileURLToPath(import.meta.url)),'..');
// What a guide's file is called. The ending is a decision for the author (see docs/STANDALONE-APP-PLAN-2026-11-13.md); change it here only.
export const PROJECT_EXTENSION='.guide-project.json';
const TYPES={'.html':'text/html; charset=utf-8','.js':'text/javascript; charset=utf-8','.css':'text/css; charset=utf-8','.svg':'image/svg+xml','.png':'image/png','.json':'application/json','.ttf':'font/ttf','.txt':'text/plain; charset=utf-8','.md':'text/plain; charset=utf-8'};
const PROGRAM_CSP="default-src 'self'; script-src 'self'; style-src 'self'; img-src 'self' data:; connect-src 'self'; frame-src 'self'; base-uri 'none'; form-action 'none'";
const PLAIN=/^[A-Za-z0-9_-]+\.(html|css|js|svg|png|json)$/;
// The fonts that come with the program, and their licenses (program/fonts): plain names only, like the rest.
const FONT_FILE=/^[A-Za-z0-9_-]+\.(ttf|txt|md)$/;
const RESERVED=/^(con|prn|aux|nul|com[0-9]|lpt[0-9])$/i;

// A guide's name made into a file name: what Windows does not allow is taken out, and a name that is nothing but that is refused.
export function fileNameFor(raw){
  const name=String(raw??'').replace(/[<>:"/\\|?*\u0000-\u001f]/g,' ').replace(/\s+/g,' ').trim().replace(/[. ]+$/,'');
  if(!name)throw new Error('Type a name for the guide.');
  if(name.length>80)throw new Error('A guide\'s name can be up to 80 letters.');
  if(RESERVED.test(name))throw new Error('Windows does not allow that name for a file. Choose another.');
  return name+PROJECT_EXTENSION;
}

export async function startProjectServer({file,folder,port=3850,snapshot,maxBytes=LIMITS.fileCharacters*4}={}){
  if(!file&&!folder)throw new Error('Name the guides folder with --folder, or one project file with --file.');
  const dir=path.resolve(folder||path.dirname(file));
  let current=file?path.basename(file):null,backedUp=false;
  const snap=snapshot||await createBuildSnapshot(root);
  const handler=createHandler(snap);
  const target=()=>current?path.join(dir,current):null;
  const send=(res,code,body='',type='text/plain; charset=utf-8',extra={})=>{res.writeHead(code,{'Content-Type':type,'Cache-Control':'no-store','X-Content-Type-Options':'nosniff','Referrer-Policy':'no-referrer',...extra});res.end(body);};
  const json=(res,code,data)=>send(res,code,JSON.stringify(data),'application/json; charset=utf-8');
  const own=()=>new Set([`127.0.0.1:${server.address().port}`,`localhost:${server.address().port}`]);

  async function readBody(req){
    const chunks=[];let size=0;
    for await(const c of req){size+=c.length;if(size>maxBytes)return null;chunks.push(c);}
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
  const stateNow=async()=>({current,folder:dir,guides:await guides(),version:snap.version,build:String(snap.buildId||'').slice(0,8)});

  async function project(req,res){
    if(req.method==='GET'||req.method==='HEAD'){
      if(!current){json(res,409,{error:'No guide is open.'});return;}
      let text;
      try{text=await fs.readFile(target(),'utf8');}catch(e){if(e.code==='ENOENT'){send(res,204);return;}throw e;}
      send(res,200,req.method==='HEAD'?'':text);
      return;
    }
    if(req.method!=='PUT'){send(res,405,'Method not allowed');return;}
    if(!allowed(req,res,/^(text\/plain|application\/json)\b/i))return;
    if(!current){json(res,409,{error:'No guide is open.'});return;}
    const body=await readBody(req);
    if(body===null){send(res,413,'That project is larger than this program can store.');return;}
    try{parseProject(body);}catch(e){send(res,400,e.message);return;}
    // First save since the guide was opened: keep what was there, once.
    const file=target();
    if(!backedUp){
      backedUp=true;
      try{await fs.copyFile(file,file+'.bak');}catch(e){if(e.code!=='ENOENT')throw e;}
    }
    await fs.mkdir(path.dirname(file),{recursive:true});
    const temp=`${file}.${process.pid}.tmp`;
    try{
      await fs.writeFile(temp,body,'utf8');
      await fs.rename(temp,file);
    }catch(e){
      // A failed move leaves the old project whole; do not leave the half-way copy lying about.
      await fs.rm(temp,{force:true});
      throw e;
    }
    send(res,200,'saved');
  }

  async function api(req,res,name){
    if(req.method==='GET'&&name==='state'){json(res,200,await stateNow());return;}
    if(req.method!=='POST'){send(res,405,'Method not allowed');return;}
    if(!allowed(req,res,/^application\/json\b/i))return;
    const raw=await readBody(req);
    let body={};
    try{body=raw?JSON.parse(raw):{};}catch{json(res,400,{error:'That request was not readable.'});return;}
    if(name==='new'){
      let file;
      try{file=fileNameFor(body.name);}catch(e){json(res,400,{error:e.message});return;}
      await fs.mkdir(dir,{recursive:true});
      try{await fs.writeFile(path.join(dir,file),serializeProject(cleanGuide({paragraphs:[]})),{encoding:'utf8',flag:'wx'});}
      catch(e){if(e.code==='EEXIST'){json(res,409,{error:'There is already a guide with that name. Choose another name, or open it.'});return;}throw e;}
      current=file;backedUp=false;
      json(res,200,await stateNow());return;
    }
    if(name==='open'){
      // The name must be one the folder lists: a request never names a path.
      const known=(await guides()).find(g=>g.name===body.name);
      if(!known){json(res,404,{error:'That guide is not in the guides folder.'});return;}
      current=known.name;backedUp=false;
      json(res,200,await stateNow());return;
    }
    if(name==='close'){current=null;json(res,200,await stateNow());return;}
    json(res,404,{error:'Not found.'});
  }

  // The tools of the program are the add-in's pane. Office.js is the add-in's way of talking to Word, and it sends a page that is shown in a
  // frame away to a blank one, so the program's copy of the pane page leaves its script out; the pane itself is the same page.
  function tools(req,res){
    const html=snap.read('/taskpane.html');
    if(!html){send(res,404,'Not found');return;}
    const page=html.toString().replace(/<script src="https:\/\/appsforoffice\.microsoft\.com[^"]*"><\/script>\s*/,'');
    send(res,200,req.method==='HEAD'?'':page,'text/html; charset=utf-8');
  }

  async function program(req,res,pathname){
    if(req.method!=='GET'&&req.method!=='HEAD'){send(res,405,'Method not allowed');return;}
    if(pathname==='/program/tools'){tools(req,res);return;}
    const rest=pathname.slice('/program/'.length)||'index.html';
    const parts=rest.split('/');
    const plain=((parts.length===1||(parts.length===2&&parts[0]==='ui'))&&parts.every((p,i)=>i===parts.length-1?PLAIN.test(p):p==='ui'))||(parts.length===2&&parts[0]==='fonts'&&FONT_FILE.test(parts[1]));
    if(!plain){send(res,404,'Not found');return;}
    let data;
    try{data=await fs.readFile(path.join(root,'program',...parts));}catch{send(res,404,'Not found');return;}
    send(res,200,req.method==='HEAD'?'':data,TYPES[path.extname(rest)]||'application/octet-stream',{'Content-Security-Policy':PROGRAM_CSP});
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
  await new Promise((resolve,reject)=>{server.once('error',reject);server.listen(port,'127.0.0.1',resolve);});
  return {server,port:server.address().port,folder:dir,file:target(),close:()=>new Promise(r=>server.close(r))};
}

// Run from the command line.
if(process.argv[1]&&path.resolve(process.argv[1])===fileURLToPath(import.meta.url)){
  const arg=name=>{const i=process.argv.indexOf(name);return i>=0?process.argv[i+1]:undefined;};
  try{
    const folder=arg('--folder'),file=arg('--file');
    const s=await startProjectServer({file,folder:folder||(file?undefined:path.join(os.homedir(),'Documents','OpenStage Guides')),port:Number(arg('--port')||3850)});
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
