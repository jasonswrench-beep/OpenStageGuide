// Real-browser check of the program (program/): the welcome page, a new guide, the tools beside it, a tool from a menu, Update guide, the guide
// drawn from the project, closing and opening again, the command search and the keys window, and what the page may not do.
// It needs Playwright, which this repository does not depend on: set PLAYWRIGHT_DIR to the playwright package folder
// (for example <your folder>/playwright-tools/node_modules/playwright), then run  node scripts/check-program-browser.mjs  from the repository
// root. It starts its own server on port 3852 with a scratch guides folder, and stops it again. Screenshots go to a temporary folder named
// in the last line.
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
const RUN=fs.mkdtempSync(path.join(os.tmpdir(),'osg-program-check-'));
const SHOTS=path.join(RUN,'shots'),GUIDES=path.join(RUN,'guides');
fs.mkdirSync(SHOTS);
const PORT=3852,BASE=`http://127.0.0.1:${PORT}`;
let pass=0,fail=0;
const ok=(cond,what)=>{if(cond){pass++;console.log('PASS',what);}else{fail++;console.log('FAIL',what);}};

const server=spawn(process.execPath,['scripts/project-server.mjs','--folder',GUIDES,'--port',String(PORT)],{cwd:APP,stdio:['ignore','pipe','pipe']});
console.log('server pid',server.pid);
await new Promise((res,rej)=>{let t='';server.stdout.on('data',d=>{t+=d;if(/OpenStage Guide at/.test(t))res();});server.stderr.on('data',d=>rej(new Error(String(d))));setTimeout(()=>rej(new Error('server did not start')),15000);});

let browser;
try{
  browser=await chromium.launch();
  const page=await browser.newPage({viewport:{width:1280,height:800}});
  const problems=[];
  page.on('pageerror',e=>problems.push(String(e)));
  page.on('console',m=>{if(m.type()==='error'||/Refused|violat/i.test(m.text()))problems.push(m.text());});
  const shot=name=>page.screenshot({path:path.join(SHOTS,name)});
  // Who saves the guide: every request that stores the project, and the page it came from.
  const stores=[];
  page.on('request',r=>{if(r.method()==='PUT'&&r.url().endsWith('/project'))stores.push(new URL(r.frame().url()).pathname);});

  // ---- the welcome page
  await page.goto(`${BASE}/`);
  await page.waitForSelector('#welcome-new');
  ok(page.url().endsWith('/program/'),'the address / goes to the program');
  ok((await page.locator('h1').innerText())==='OpenStage Guide','the welcome page has its heading');
  ok(await page.locator('main#main').count()===1&&await page.locator('a.skip').count()===1,'a main region and a skip link');
  ok((await page.locator('[role=menubar] > .menubar-slot').count())===8,'eight menus: File, Guide, Write, Bring in, E-learning, Check, Make, Help');
  ok(await page.locator('#bar-buttons button').count()===0,'no guide buttons on the navy bar before a guide is open');
  ok(/no guides yet/i.test(await page.locator('.recent').innerText()),'it says there are no guides yet');
  await shot('1-welcome.png');

  // ---- a tool cannot be used with no guide
  await page.getByRole('menuitem',{name:/^Write/}).click();
  const addModule=page.getByRole('menuitem',{name:/Add module/});
  ok((await addModule.getAttribute('aria-disabled'))==='true','Add module is dimmed with no guide open');
  await page.keyboard.press('Escape');

  // ---- a new guide
  await page.click('#welcome-new');
  ok(await page.locator('dialog[open] h2').innerText()==='New guide','the New guide dialog opens');
  ok(await page.evaluate(()=>document.activeElement?.id)==='dlg-name','focus starts in the name box');
  await page.getByRole('button',{name:'Create'}).click();
  ok(/Type a name first/.test(await page.locator('dialog[open] .problem').innerText()),'an empty name is refused in the dialog');
  await page.fill('#dlg-name','Reframing conflict');
  await shot('2-new-dialog.png');
  await page.press('#dlg-name','Enter');
  await page.waitForSelector('#tools-frame');
  const fr=page.frameLocator('#tools-frame');
  await fr.locator('#guide-start').waitFor({state:'visible',timeout:15000});
  ok((await page.locator('.bar-context-name').innerText())==='Reframing conflict','the navy bar names the open guide');
  ok(await page.locator('#bar-buttons button').count()===3,'three buttons on the bar: Review guide, Update guide, Make workbook');
  ok(await page.locator('#bar-buttons .primary').innerText().then(t=>/Make workbook/.test(t)),'Make workbook is the filled one');
  ok((await page.locator('#tools-frame').getAttribute('title'))==='Guide tools','the frame has a title');
  const frameDoc=page.frames().find(f=>f.url().includes('/program/tools'));
  ok(!!frameDoc&&!await frameDoc.evaluate(()=>!!document.querySelector('script[src*="appsforoffice"]')),'the tools page has no Office.js');
  ok(await frameDoc.evaluate(()=>getComputedStyle(document.querySelector('header')).display==='none'),'the pane\'s own header is not shown in the program');
  ok(files(GUIDES).includes('Reframing conflict.guide-project.json'),'the guide is a file in the guides folder');
  await shot('3-guide-open.png');

  // ---- write in the tools, see it in the guide
  await fr.locator('#guide-title').fill('Reframing conflict');
  await fr.locator('#guide-title').press('Enter');
  await fr.locator('#guide-build').waitFor({state:'visible',timeout:15000});
  await page.waitForFunction(()=>/Reframing conflict/.test(document.querySelector('#work-guide')?.innerText||''),null,{timeout:10000});
  ok(true,'the guide is drawn after the tools save: its title is on the page');
  await page.waitForFunction(()=>document.querySelector('.bar-context-state')?.getAttribute('data-state')==='saved',null,{timeout:10000});
  ok((await page.locator('.bar-context-state').innerText())==='Saved','the navy bar says Saved');
  ok(/Acknowledgements/.test(await page.locator('#work-guide').innerText()),'the standard pages are drawn');

  // ---- a tool from a menu
  await page.getByRole('menuitem',{name:/^Write/}).click();
  ok((await page.getByRole('menuitem',{name:/Add module/}).getAttribute('aria-disabled'))!=='true','Add module can be used now');
  await shot('4-write-menu.png');
  await page.getByRole('menuitem',{name:/Add module/}).click();
  await fr.locator('#guide-new-title').waitFor({state:'visible',timeout:10000});
  await fr.locator('#guide-new-title').fill('Naming the feeling');
  await fr.locator('#guide-add-module').click();
  await page.waitForFunction(()=>/Module 1: Naming the feeling/.test(document.querySelector('#work-guide')?.innerText||''),null,{timeout:10000});
  ok(true,'a module added in the tools appears in the guide');
  await shot('5-module-added.png');

  // ---- a button on the navy bar
  await page.getByRole('button',{name:'Update guide'}).click();
  await page.waitForTimeout(600);
  ok(!problems.length,'nothing went wrong in the page so far: '+JSON.stringify(problems.slice(0,2)));

  // ---- the workbook reaches the author: there is no Word to open it in, so it is a real download
  // (a workbook needs something for learners in the guide: a highlight block, added from the Write menu)
  await page.getByRole('menuitem',{name:/^Write/}).click();
  await page.getByRole('menuitem',{name:/Instructional cues/i}).click();
  await fr.locator('#guide-cue-highlight').evaluate(e=>{e.checked=true;e.dispatchEvent(new Event('change',{bubbles:true}));});
  await fr.locator('#guide-block-text').fill('Name the feeling first.');
  await fr.locator('#guide-block-form').evaluate(f=>f.requestSubmit());
  await page.waitForFunction(()=>/Name the feeling first\./.test(document.querySelector('#work-guide')?.innerText||''),null,{timeout:10000});
  ok(true,'a block added in the tools appears in the guide');
  await page.getByRole('button',{name:'Make workbook'}).click();
  await fr.locator('#guide-workbook-go').waitFor({state:'visible',timeout:10000});
  const [download]=await Promise.all([page.waitForEvent('download',{timeout:20000}),fr.locator('#guide-workbook-go').click()]);
  ok(/\.docx$/.test(download.suggestedFilename()),'Make workbook downloads a Word file: '+download.suggestedFilename());
  const downloaded=await download.path();
  ok(!!downloaded&&fs.statSync(downloaded).size>2000,'the downloaded workbook is a real file');
  ok(/downloaded/i.test(await fr.locator('#guide-workbook-ok').innerText().catch(()=>'')),'the tools say it was downloaded');
  await shot('5b-workbook.png');

  // ---- the command search and the keys window
  await page.keyboard.press('Control+k');
  await page.locator('dialog.command-search').waitFor({state:'visible'});
  await page.keyboard.type('forms');
  await page.keyboard.press('Escape');
  await page.keyboard.press('Control+k');
  await page.keyboard.type('workbook');
  ok(/Participant workbook/.test(await page.locator('.cs-list').innerText()),'Ctrl+K finds a command by its name');
  await shot('6-search.png');
  await page.keyboard.press('Escape');
  ok(await page.locator('dialog[open]').count()===0,'Escape closes the search');
  await page.keyboard.press('Control+/');
  await page.locator('dialog.keys-dialog').waitFor({state:'visible'});
  ok((await page.locator('dialog.keys-dialog h2').innerText())==='Keyboard keys','Ctrl+/ opens the key list');
  await shot('7-keys.png');
  await page.keyboard.press('Escape');

  // ---- the keyboard reaches the menus
  await page.keyboard.press('F10');
  // (when the keyboard is in the tools' frame, F10 reaches the menus by a message to this page, so it is waited for, not read at once)
  const menus=await page.waitForFunction(()=>document.activeElement?.classList.contains('menubar-item'),null,{timeout:3000}).then(()=>true,()=>false);
  ok(menus,'F10 goes to the menus');
  await page.keyboard.press('Escape');

  // ---- close, and the guide is on the welcome page; open it again
  await page.getByRole('menuitem',{name:/^File/}).click();
  await page.getByRole('menuitem',{name:/Close guide/}).click();
  await page.waitForSelector('#welcome-new');
  ok(await page.locator('.recent li button').count()===1,'the closed guide is in the list');
  ok(await page.locator('#bar-buttons button').count()===0,'the bar has no guide buttons again');
  await page.locator('.recent li button').click();
  await page.waitForSelector('#tools-frame');
  await page.waitForFunction(()=>/Module 1: Naming the feeling/.test(document.querySelector('#work-guide')?.innerText||''),null,{timeout:10000});
  ok(true,'opened again, the guide is as it was');
  await page.getByRole('menuitem',{name:/^File/}).click();
  await page.getByRole('menuitem',{name:/Open guide/}).click();
  await page.locator('dialog[open] input[type=radio]').first().waitFor();
  ok(await page.locator('dialog[open] input[type=radio]').count()===1,'Open guide lists the guide');
  await shot('8-open-dialog.png');
  await page.keyboard.press('Escape');

  // ---- what the page may not do
  const before=problems.length;
  const csp=await page.evaluate(async()=>(await fetch('/program/')).headers.get('content-security-policy'));
  ok(/script-src 'self'/.test(csp||''),'the program\'s page has a content security policy');
  const bad=await page.evaluate(async()=>{
    const r=await fetch('/api/new',{method:'POST',headers:{'Content-Type':'text/plain'},body:'{"name":"x"}'});
    return r.status;
  });
  ok(bad===415,'a request in the wrong form is refused (415): '+bad);
  ok(files(GUIDES).filter(f=>f.endsWith('.guide-project.json')).length===1,'nothing refused made a guide');
  // (the one request above that was refused is the page's own probe, and the browser logs it)
  ok(stores.length>0&&stores.every(f=>f==='/program/'),'the guide is saved by the program page only, never by the tools frame: '+JSON.stringify([...new Set(stores)]));
  ok(problems.slice(0,before).length===0&&problems.slice(before).every(p=>/415/.test(p)),'no page errors or policy violations in the whole run: '+JSON.stringify(problems.slice(0,3)));
}finally{
  if(browser)await browser.close();
  server.kill();
}
function files(dir){try{return fs.readdirSync(dir);}catch{return [];}}
console.log(`\n${pass} passed, ${fail} failed (screenshots in ${SHOTS})`);
process.exitCode=fail?1:0;
