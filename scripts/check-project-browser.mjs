// Real-browser check of the project file: write a guide in the pane, reload, keep working, a broken file, and the server's guards.
// It needs Playwright, which this repository does not depend on: set PLAYWRIGHT_DIR to the playwright package folder
// (for example <your folder>/playwright-tools/node_modules/playwright), then run  node scripts/check-project-browser.mjs  from the
// repository root. It starts its own server on port 3851 and stops it again. Screenshots go to a temporary folder named in the last line.
import {createRequire} from 'node:module';
import {spawn} from 'node:child_process';
import fs from 'node:fs';
import http from 'node:http';
import os from 'node:os';
import path from 'node:path';
import {fileURLToPath} from 'node:url';
const require=createRequire(import.meta.url);
if(!process.env.PLAYWRIGHT_DIR){console.error('Set PLAYWRIGHT_DIR to the playwright package folder.');process.exit(2);}
const {chromium}=require(process.env.PLAYWRIGHT_DIR);
const APP=path.resolve(path.dirname(fileURLToPath(import.meta.url)),'..');
const RUN=fs.mkdtempSync(path.join(os.tmpdir(),'osg-project-check-'));
fs.mkdirSync(path.join(RUN,'shots'));
const file=path.join(RUN,'My guide.project.json');
const PORT=3851,BASE=`http://127.0.0.1:${PORT}`;
let pass=0,fail=0;
const ok=(cond,what)=>{if(cond){pass++;console.log('PASS',what);}else{fail++;console.log('FAIL',what);}};

const server=spawn(process.execPath,['scripts/project-server.mjs','--file',file,'--port',String(PORT)],{cwd:APP,stdio:['ignore','pipe','pipe']});
console.log('server pid',server.pid);
await new Promise((res,rej)=>{let t='';server.stdout.on('data',d=>{t+=d;if(/OpenStage Guide at/.test(t))res();});server.stderr.on('data',d=>rej(new Error(String(d))));setTimeout(()=>rej(new Error('server did not start')),15000);});

const raw=(method,urlPath,{headers={},body}={})=>new Promise((resolve,reject)=>{
  const req=http.request({host:'127.0.0.1',port:PORT,path:urlPath,method,headers},res=>{let d='';res.on('data',c=>d+=c);res.on('end',()=>resolve({status:res.statusCode,text:d}));});
  req.on('error',reject);if(body)req.write(body);req.end();
});

let browser;
try{
  browser=await chromium.launch();
  const page=await browser.newPage({viewport:{width:360,height:900}});
  const errors=[];page.on('pageerror',e=>errors.push(String(e)));page.on('console',m=>{if(m.type()==='error')errors.push(m.text());});
  const conn=()=>page.locator('#connection').innerText();
  // The browser has no ribbon: the pane's own "More tools" buttons lead from the guide screen to Add Module (through the Course Outline).
  const moreTool=async id=>{await page.locator('#tool-more > summary').click();await page.click('#tool-more-'+id);};
  const goAddModule=async()=>{await moreTool('courseOutline');await moreTool('addModule');};
  const waitSaved=async()=>{await page.waitForFunction(()=>/Project file · saved/.test(document.getElementById('connection').textContent),null,{timeout:8000});};

  await page.goto(`${BASE}/taskpane.html?project=1`);
  await page.locator('#guide-start').waitFor({state:'visible',timeout:10000});
  ok(true,'the pane opens in project mode and shows the Start screen for a project with no file');
  ok(!fs.existsSync(file),'no file is made just by opening an empty project');
  ok(/Project file/.test(await conn()),'the header says it is a project file: '+JSON.stringify(await conn()));

  await page.fill('#guide-title','Spike guide');
  await page.press('#guide-title','Enter');
  await page.locator('#guide-build').waitFor({state:'visible',timeout:10000});
  await goAddModule();
  await page.fill('#guide-new-title','Naming the feeling');
  await page.click('#guide-add-module');
  await page.waitForFunction(()=>document.querySelector('#guide-status')?.textContent.length>0);
  await waitSaved();
  await page.screenshot({path:path.join(RUN,'shots','1-written.png'),fullPage:true});
  ok(fs.existsSync(file),'the file exists after the first changes');
  let saved=JSON.parse(fs.readFileSync(file,'utf8'));
  ok(saved.format==='openstage-guide-project'&&saved.version===1,'it is a project file of the right format and version');
  ok(saved.guide.title==='Spike guide','the title is in the file');
  ok(saved.guide.paragraphs.some(p=>p.text.trim()==='Module 1: Naming the feeling'&&p.style==='OSG Module'),'the module is in the file with its Word style');
  ok(fs.readdirSync(RUN).filter(n=>/\.tmp$/.test(n)).length===0,'no temporary file is left behind');
  const first=fs.readFileSync(file,'utf8');

  // A new session: reload and see the same guide.
  await page.reload();
  await page.locator('#guide-build').waitFor({state:'visible',timeout:10000});
  ok((await page.locator('#guide-build-heading').innerText())==='Spike guide','after a reload the guide is back: '+await page.locator('#guide-build-heading').innerText());
  ok(await page.locator('#guide-start').isHidden(),'the Start screen is not shown for a guide that exists');
  ok(/saved/.test(await conn()),'the header says saved: '+JSON.stringify(await conn()));
  await page.screenshot({path:path.join(RUN,'shots','2-reopened.png'),fullPage:true});

  // Keep working on it.
  await goAddModule();
  await page.fill('#guide-new-title','Say it plainly');
  await page.fill('#guide-new-minutes','10');
  await page.click('#guide-add-lesson');
  await page.waitForFunction(()=>/Project file · (saving|changes not saved yet|saved)/.test(document.getElementById('connection').textContent));
  await waitSaved();
  saved=JSON.parse(fs.readFileSync(file,'utf8'));
  ok(saved.guide.paragraphs.some(p=>/^Lesson 1\.1: Say it plainly/.test(p.text.trim())),'a lesson added after reopening is saved');
  ok(fs.existsSync(file+'.bak')===false,'no backup yet: the file was made by this run');

  // A broken file is explained and never written over.
  fs.writeFileSync(file,'this is not a project');
  await page.reload();
  await page.locator('#error').waitFor({state:'visible',timeout:10000});
  const msg=await page.locator('#error').innerText();
  ok(/not an OpenStage Guide project|not readable/.test(msg)&&/saving is off/.test(msg),'a broken file is explained: '+JSON.stringify(msg));
  await page.screenshot({path:path.join(RUN,'shots','3-broken-file.png'),fullPage:true});
  await new Promise(r=>setTimeout(r,1200));
  ok(fs.readFileSync(file,'utf8')==='this is not a project','the broken file is left exactly as it was');
  ok(await page.locator('#guide-start').isHidden(),'the Start screen is not offered over a file that could not be read');

  // Put the good file back; a new session opens it.
  fs.writeFileSync(file,first);
  await page.reload();
  await page.locator('#guide-build').waitFor({state:'visible',timeout:10000});
  ok((await page.locator('#guide-build-heading').innerText())==='Spike guide','a repaired file opens again');
  ok(errors.filter(e=>!/Failed to load resource.*(404|204)/.test(e)).length===0,'no page errors: '+JSON.stringify(errors.slice(0,3)));

  // The server's guards.
  const good=first;
  let r=await raw('PUT','/project',{headers:{'Content-Type':'text/plain','Host':'evil.example'},body:good});ok(r.status===403,'a request for another host name is refused (403): '+r.status);
  r=await raw('PUT','/project',{headers:{'Content-Type':'text/plain','Host':`127.0.0.1:${PORT}`,'Origin':'http://evil.example'},body:good});ok(r.status===403,'a request from another web site is refused (403): '+r.status);
  r=await raw('PUT','/project',{headers:{'Content-Type':'text/plain','Host':`127.0.0.1:${PORT}`,'Sec-Fetch-Site':'cross-site'},body:good});ok(r.status===403,'a cross-site fetch is refused (403): '+r.status);
  r=await raw('PUT','/project',{headers:{'Content-Type':'application/x-www-form-urlencoded','Host':`127.0.0.1:${PORT}`},body:good});ok(r.status===415,'a form post is refused (415): '+r.status);
  r=await raw('PUT','/project',{headers:{'Content-Type':'text/plain','Host':`127.0.0.1:${PORT}`},body:'{"nope":1}'});ok(r.status===400,'something that is not a project is refused (400): '+r.status+' '+r.text);
  r=await raw('PUT','/project',{headers:{'Content-Type':'text/plain','Host':`127.0.0.1:${PORT}`},body:JSON.stringify({format:'openstage-guide-project',version:9,guide:{paragraphs:[]}})});ok(r.status===400&&/newer/.test(r.text),'a project from a newer program is refused: '+r.text);
  r=await raw('POST','/project',{headers:{'Host':`127.0.0.1:${PORT}`}});ok(r.status===405,'POST is not allowed (405): '+r.status);
  ok(fs.readFileSync(file,'utf8')===good,'none of the refused requests changed the file');
  r=await raw('GET','/project',{headers:{'Host':`127.0.0.1:${PORT}`}});ok(r.status===200&&r.text===good,'GET gives back the file as saved');
  r=await raw('GET','/../package.json',{headers:{'Host':`127.0.0.1:${PORT}`}});ok(r.status!==200,'a path outside the pane is not served: '+r.status);
  r=await raw('GET','/health',{headers:{'Host':`127.0.0.1:${PORT}`}});ok(r.status===200&&/OpenStage Guide/.test(r.text),'health answers');
}finally{
  if(browser)await browser.close();
  server.kill();
}
console.log(`\n${pass} passed, ${fail} failed (screenshots in ${path.join(RUN,"shots")})`);
process.exitCode=fail?1:0;
