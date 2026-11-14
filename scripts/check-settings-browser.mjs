// Real-browser check that what the program remembers besides the guides (the agenda's start time, the workbook's look) is still there after the
// program is started again on another port, as the window does every time it starts. (They were once kept in the page's local storage, which belongs
// to one address and port, so they were gone the next time.) The tools' page is the real one; the program's server is started twice on one scratch
// guides folder, on two ports.
// It needs Playwright, which this repository does not depend on: set PLAYWRIGHT_DIR to the playwright package folder, then run
//   node scripts/check-settings-browser.mjs   from the repository root. It stops its servers again.
import {createRequire} from 'node:module';
import {spawn} from 'node:child_process';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import {fileURLToPath} from 'node:url';
const require=createRequire(import.meta.url);
if(!process.env.PLAYWRIGHT_DIR){console.error('Set PLAYWRIGHT_DIR to the playwright package folder.');process.exit(2);}
const {chromium}=require(process.env.PLAYWRIGHT_DIR);
const APP=path.resolve(path.dirname(fileURLToPath(import.meta.url)),'..');
const RUN=fs.mkdtempSync(path.join(os.tmpdir(),'osg-settings-check-'));
const GUIDES=path.join(RUN,'guides');
let pass=0,fail=0;
const ok=(cond,what)=>{if(cond){pass++;console.log('PASS',what);}else{fail++;console.log('FAIL',what);}};

async function startServer(port){
  const server=spawn(process.execPath,['scripts/project-server.mjs','--folder',GUIDES,'--port',String(port)],{cwd:APP,stdio:['ignore','pipe','pipe'],env:{...process.env,OPENSTAGE_UPDATE_CHECK:'off'}});
  await new Promise((res,rej)=>{let t='';server.stdout.on('data',d=>{t+=d;if(/OpenStage Guide at/.test(t))res();});server.stderr.on('data',d=>rej(new Error(String(d))));setTimeout(()=>rej(new Error('server did not start')),15000);});
  return server;
}
async function openTool(page,menu,item){
  await page.getByRole('menuitem',{name:new RegExp(`^${menu}`)}).click();
  await page.getByRole('menuitem',{name:new RegExp(item)}).click();
}

let browser,server;
try{
  browser=await chromium.launch();

  // ---- the first run: a guide, a start time, a look
  server=await startServer(3856);
  let page=await browser.newPage({viewport:{width:1280,height:800}});
  await page.goto('http://127.0.0.1:3856/');
  await page.click('#welcome-new');
  await page.fill('#dlg-name','Settings check');
  await page.press('#dlg-name','Enter');
  await page.waitForSelector('#tools-frame');
  let fr=page.frameLocator('#tools-frame');
  await fr.locator('#guide-title').fill('Settings check');
  await fr.locator('#guide-title').press('Enter');
  await fr.locator('#guide-build').waitFor({state:'visible',timeout:15000});
  await page.waitForFunction(()=>document.querySelector('.bar-context-state')?.getAttribute('data-state')==='saved',null,{timeout:10000});
  await openTool(page,'Check','Timing and agenda');
  await fr.locator('#guide-agenda-start').waitFor({state:'visible',timeout:10000});
  ok(await fr.locator('#guide-agenda-start').inputValue()==='09:00','the start time begins at 09:00');
  await fr.locator('#guide-agenda-start').fill('13:30');
  await fr.locator('#guide-agenda-start').dispatchEvent('change');
  await openTool(page,'Make','Participant workbook');
  await fr.locator('#guide-wb-look-practice-lab').waitFor({state:'visible',timeout:10000});
  await fr.locator('#guide-wb-look-practice-lab').check();
  const settings=async()=>(await page.evaluate(async()=>(await fetch('/api/settings')).json())).values;
  let values={};
  for(let i=0;i<20;i++){values=await settings();if(values['osg.agendaStart.v1']&&values['openstage-guide.workbook-look'])break;await page.waitForTimeout(100);}
  ok(values['osg.agendaStart.v1']==='13:30','the start time is kept by the program');
  ok(/practice-lab/.test(values['openstage-guide.workbook-look']||''),'the look is kept by the program');
  ok(fs.existsSync(path.join(GUIDES,'OpenStage Guide settings.json')),'in the file beside the guides');
  const local=await page.frames().find(f=>f.url().includes('/program/tools')).evaluate(()=>({n:localStorage.length,keys:Object.keys(localStorage)}));
  ok(local.n===0,'and not in the page\'s own local storage: '+JSON.stringify(local.keys));
  await page.close();
  server.kill();
  await new Promise(r=>server.once('exit',r));

  // ---- the second run, on another port: the same guides folder
  server=await startServer(3857);
  page=await browser.newPage({viewport:{width:1280,height:800}});
  await page.goto('http://127.0.0.1:3857/');
  await page.getByRole('button',{name:/Settings check/}).click();
  await page.waitForSelector('#tools-frame');
  fr=page.frameLocator('#tools-frame');
  await openTool(page,'Check','Timing and agenda');
  await fr.locator('#guide-agenda-start').waitFor({state:'visible',timeout:10000});
  ok(await fr.locator('#guide-agenda-start').inputValue()==='13:30','on another port the start time is still 13:30');
  await openTool(page,'Make','Participant workbook');
  await fr.locator('#guide-wb-look-practice-lab').waitFor({state:'visible',timeout:10000});
  ok(await fr.locator('#guide-wb-look-practice-lab').isChecked(),'and the workbook look is still the one chosen');
  await page.close();
}finally{
  if(browser)await browser.close();
  if(server)server.kill();
  fs.rmSync(RUN,{recursive:true,force:true});
}
console.log(`\n${pass} passed, ${fail} failed`);
process.exitCode=fail?1:0;
