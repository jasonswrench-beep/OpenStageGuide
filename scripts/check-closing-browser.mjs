// Real-browser check that the program's page saves the open guide when the window asks it to (window.osgClosing, which the window calls when it
// is asked to close) and then tells its server, which says so on its output, as it does for the window. A change is made in the tools and the page is
// asked at once, before the short wait that saving usually has, and the guide on the disk then holds it.
// It needs Playwright, which this repository does not depend on: set PLAYWRIGHT_DIR to the playwright package folder, then run
//   node scripts/check-closing-browser.mjs   from the repository root. It stops its server again.
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
const RUN=fs.mkdtempSync(path.join(os.tmpdir(),'osg-closing-check-'));
const GUIDES=path.join(RUN,'guides');
const PORT=3858;
let pass=0,fail=0;
const ok=(cond,what)=>{if(cond){pass++;console.log('PASS',what);}else{fail++;console.log('FAIL',what);}};

// As the window starts it: its input stays open, and what it says on its output is what the window reads.
const server=spawn(process.execPath,['scripts/project-server.mjs','--folder',GUIDES,'--port',String(PORT),'--exit-when-stdin-closes','--no-update-check'],{cwd:APP,stdio:['pipe','pipe','pipe']});
let out='';
server.stdout.on('data',d=>{out+=d;});
await new Promise((res,rej)=>{server.stderr.on('data',d=>rej(new Error(String(d))));setTimeout(()=>rej(new Error('server did not start')),15000);const t=setInterval(()=>{if(/OPENSTAGE_READY/.test(out)){clearInterval(t);res();}},30);});

let browser;
try{
  browser=await chromium.launch();
  const page=await browser.newPage({viewport:{width:1280,height:800}});
  await page.goto(`http://127.0.0.1:${PORT}/`);
  await page.click('#welcome-new');
  await page.fill('#dlg-name','Closing check');
  await page.press('#dlg-name','Enter');
  await page.waitForSelector('#tools-frame');
  const fr=page.frameLocator('#tools-frame');
  await fr.locator('#guide-title').fill('Closing check');
  await fr.locator('#guide-title').press('Enter');
  await fr.locator('#guide-build').waitFor({state:'visible',timeout:15000});
  await page.waitForFunction(()=>document.querySelector('.bar-context-state')?.getAttribute('data-state')==='saved',null,{timeout:10000});
  ok(await page.evaluate(()=>typeof window.osgClosing)==='function','the page offers what the window asks for');
  ok(await page.evaluate(()=>{try{window.osgClosing=()=>1;}catch{/* strict mode would throw */}return window.osgClosing.toString().includes('post');}),'and nothing can replace it');

  // a change, and the page asked at once
  await page.getByRole('menuitem',{name:/^Write/}).click();
  await page.getByRole('menuitem',{name:/Add module/}).click();
  await fr.locator('#guide-new-title').waitFor({state:'visible',timeout:10000});
  await fr.locator('#guide-new-title').fill('The last change before closing');
  await fr.locator('#guide-add-module').click();
  await page.waitForFunction(()=>/The last change before closing/.test(document.querySelector('#work-guide')?.innerText||''),null,{timeout:10000});
  await page.evaluate(()=>window.osgClosing());
  const file=path.join(GUIDES,'Closing check.guide-project.json');
  ok(/The last change before closing/.test(fs.readFileSync(file,'utf8')),'when the page has answered, the guide on the disk holds the last change');
  ok(await page.locator('.bar-context-state').innerText()==='Saved','and the navy bar says Saved');
  for(let i=0;i<50&&!/OPENSTAGE_CLOSE_READY/.test(out);i++)await new Promise(r=>setTimeout(r,50));
  ok(/^OPENSTAGE_CLOSE_READY$/m.test(out)&&!/OPENSTAGE_CLOSE_HELD/.test(out),'the server told the window the page had saved');

  // with no guide open there is nothing to save, and the window may still go
  await page.getByRole('menuitem',{name:/^File/}).click();
  await page.getByRole('menuitem',{name:/Close guide/}).click();
  await page.waitForSelector('#welcome-new');
  out='';
  await page.evaluate(()=>window.osgClosing());
  for(let i=0;i<50&&!/OPENSTAGE_CLOSE_READY/.test(out);i++)await new Promise(r=>setTimeout(r,50));
  ok(/^OPENSTAGE_CLOSE_READY$/m.test(out),'with no guide open the window is told it may go at once');
}finally{
  if(browser)await browser.close();
  server.stdin.end();
  setTimeout(()=>server.kill(),3000).unref();
  fs.rmSync(RUN,{recursive:true,force:true});
}
console.log(`\n${pass} passed, ${fail} failed`);
process.exitCode=fail?1:0;
