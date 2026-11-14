// Real-browser check of the program (program/): the welcome page, a new guide, the tools beside it, a tool from a menu, Update guide, the guide
// drawn from the project, closing and opening again, the command search and the keys window, and what the page may not do.
// It needs Playwright, which this repository does not depend on: set PLAYWRIGHT_DIR to the playwright package folder
// (for example <your folder>/playwright-tools/node_modules/playwright), then run  node scripts/check-program-browser.mjs  from the repository
// root. It starts its own server on port 3852 with a scratch guides folder, and stops it again. Screenshots go to a temporary folder named
// in the last line.
import {createRequire} from 'node:module';
import {spawn} from 'node:child_process';
import fs from 'node:fs';
import JSZip from 'jszip';
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

// (Check for updates is turned off for the run, so that nothing here reaches the network; its answer is then the one that says so.)
const server=spawn(process.execPath,['scripts/project-server.mjs','--folder',GUIDES,'--port',String(PORT)],{cwd:APP,stdio:['ignore','pipe','pipe'],env:{...process.env,OPENSTAGE_UPDATE_CHECK:'off'}});
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
  ok((await page.locator('[role=menubar] > .menubar-slot').count())===11,'eleven menus: File, Edit, Insert, Format, Guide, Write, Bring in, E-learning, Check, Make, Help');
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
  // A name Windows does not allow is refused in the same dialog, which stays open with what was typed (it once closed and left a message behind).
  await page.fill('#dlg-name','Con');
  await page.press('#dlg-name','Enter');
  await page.waitForFunction(()=>/Windows does not allow that name/.test(document.querySelector('dialog[open] .problem')?.textContent||''),null,{timeout:10000});
  ok(await page.locator('dialog[open]').count()===1&&await page.inputValue('#dlg-name')==='Con'&&await page.locator('dialog[open] .problem').isVisible(),'a name Windows does not allow is refused in the dialog, which stays open with the name still in it');
  ok(!files(GUIDES).some(f=>/^con\./i.test(f)),'and no guide was made');
  // The server's refusal (400) is what the browser logs for it; that one message is expected here and is not a fault of the page.
  let refused=-1;
  for(let i=0;i<20&&refused<0;i++){refused=problems.findIndex(p=>/status of 400/.test(p));if(refused<0)await page.waitForTimeout(100);}
  ok(refused>=0,'the refusal was the server\'s 400, which the check expects');
  if(refused>=0)problems.splice(refused,1);
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
  // What the tools said stays pinned at the top of their pane while buttons are used, and stops being pinned while a box is typed in, so it never covers what is written.
  const pinned=()=>fr.locator('#guide-status').evaluate(e=>getComputedStyle(e).position);
  await fr.locator('#guide-status').waitFor({state:'visible',timeout:5000});
  ok(await pinned()==='sticky','what the tools said is pinned at the top of their pane');
  await fr.locator('#guide-new-title').focus();
  ok(await pinned()==='static','and is not pinned while a box is being typed in');
  await fr.locator('#guide-new-title').blur();
  ok(await pinned()==='sticky','and is pinned again when the box is left');
  await shot('5-module-added.png');

  // ---- a button on the navy bar
  await page.getByRole('button',{name:'Update guide'}).click();
  await page.waitForTimeout(600);
  ok(!problems.length,'nothing went wrong in the page so far: '+JSON.stringify(problems.slice(0,2)));
  // The contents table Update guide fills in is drawn under "Contents": an entry for each module, at its own level, with no page number (Word gives those).
  await page.waitForFunction(()=>[...document.querySelectorAll('#work-guide .g-toc')].some(e=>/Module 1: Naming the feeling/.test(e.textContent)),null,{timeout:10000});
  const toc=await page.locator('#work-guide .g-toc').evaluateAll(els=>els.map(e=>[e.className,e.textContent]));
  ok(toc.some(([c,t])=>c==='g-toc g-toc-1'&&t==='Module 1: Naming the feeling')&&toc.some(([c,t])=>c==='g-toc g-toc-2'&&t==='About This Guide')&&toc.every(([,t])=>!t.includes('	')),'Update guide fills the Contents list in, drawn in the guide at its levels, with no page numbers: '+JSON.stringify(toc.slice(0,3)));

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
  const zip=await JSZip.loadAsync(fs.readFileSync(downloaded));
  const media=Object.keys(zip.files).filter(n=>/^word\/media\//.test(n));
  const sizes=await Promise.all(media.map(async n=>(await zip.file(n).async('uint8array')).length));
  ok(media.length>=1&&sizes.every(n=>n>200),'the workbook\'s cue icons are the real pictures, not one-pixel stand-ins: '+JSON.stringify(sizes));
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
  // The About window says the version, the license and what the program is made with, in a dialog with no sideways scrollbar.
  await page.getByRole('menuitem',{name:/^Help/}).click();
  await page.getByRole('menuitem',{name:/About OpenStage Guide/}).click();
  await page.locator('dialog[open] h2').waitFor({timeout:5000});
  const about=await page.evaluate(()=>{const d=document.querySelector('dialog[open]'),b=d.querySelector('.dialog-body');return {text:d.innerText,sideways:b.scrollWidth>b.clientWidth+1,fits:d.getBoundingClientRect().bottom<=innerHeight+1};});
  ok(/Version \d+\.\d+\.\d+ \(prototype\), build [0-9a-f]{8}\./.test(about.text),'About says the version and the build');
  ok(/Apache License 2\.0/.test(about.text)&&/Made with/.test(about.text)&&/JSZip/.test(about.text)&&/SIL Open Font License/.test(about.text)&&/NOTICE file/.test(about.text),'and the license, what it is made with, and where the full notices are');
  ok(!about.sideways&&about.fits,'in a dialog with no sideways scrollbar that fits the window: '+JSON.stringify({sideways:about.sideways,fits:about.fits}));
  await page.keyboard.press('Escape');
  // Help > Check for updates: asks only when chosen, and says the answer in words (here: turned off, so nothing is asked of GitHub and no link is offered).
  await page.getByRole('menuitem',{name:/^Help/}).click();
  await page.getByRole('menuitem',{name:/Check for updates/}).click();
  await page.locator('dialog[open] h2').waitFor({timeout:5000});
  await page.waitForFunction(()=>/turned off on this computer/.test(document.querySelector('dialog[open]')?.textContent||''),null,{timeout:10000});
  const updates=await page.evaluate(()=>{const d=document.querySelector('dialog[open]'),b=d.querySelector('.dialog-body');return {title:d.querySelector('h2').textContent,text:d.innerText,links:d.querySelectorAll('a').length,sideways:b.scrollWidth>b.clientWidth+1};});
  ok(updates.title==='Check for updates'&&/Checking for updates is turned off on this computer\./.test(updates.text),'Check for updates says the answer in words: '+JSON.stringify(updates.text.slice(0,120)));
  ok(updates.links===0&&/asks GitHub only when you choose this/.test(updates.text)&&!updates.sideways,'with no link when it is off, a line saying what it does and does not do, and no sideways scrollbar');
  await page.keyboard.press('Escape');
  // And a search for "pdf" finds where a PDF comes from: Save as Word file, which Word can save as one.
  await page.keyboard.press('Control+k');
  await page.locator('dialog.command-search').waitFor({state:'visible'});
  await page.keyboard.type('pdf');
  ok(/Save as Word file/.test(await page.locator('.cs-list').innerText()),'a search for pdf finds Save as Word file');
  await page.keyboard.press('Escape');

  // ---- the keyboard reaches the menus
  await page.keyboard.press('F10');
  // (when the keyboard is in the tools' frame, F10 reaches the menus by a message to this page, so it is waited for, not read at once)
  const menus=await page.waitForFunction(()=>document.activeElement?.classList.contains('menubar-item'),null,{timeout:3000}).then(()=>true,()=>false);
  ok(menus,'F10 goes to the menus');
  await page.keyboard.press('Escape');

  // ---- a menu taller than the room under the bar (the Format menu in a short window) stays inside the window and scrolls, so its last command can be reached
  const normalSize=page.viewportSize();
  await page.setViewportSize({width:1000,height:600});
  await page.getByRole('menuitem',{name:/^Format/}).click();
  const tall=await page.evaluate(()=>{
    const m=document.querySelector('.menu:not([hidden])');
    const items=[...m.querySelectorAll('[role=menuitem]')];
    const box=m.getBoundingClientRect();
    m.scrollTop=m.scrollHeight;
    const last=items[items.length-1].getBoundingClientRect();
    return {bottom:Math.round(box.bottom),winH:innerHeight,scrolls:m.scrollHeight>m.clientHeight,lastShown:last.bottom<=innerHeight&&last.top>=0,lastText:items[items.length-1].textContent.trim()};
  });
  ok(tall.bottom<=tall.winH&&tall.scrolls&&tall.lastShown,'in a short window the Format menu stays inside it and scrolls to its last command ('+tall.lastText+'): '+JSON.stringify(tall));
  await page.keyboard.press('Escape');
  await page.setViewportSize(normalSize);

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

  // ---- a page left over from another tab: another guide is opened elsewhere, and this page's next save is refused and writes to neither guide
  const first=path.join(GUIDES,files(GUIDES).find(f=>f.endsWith('.guide-project.json')));
  const firstBefore=fs.readFileSync(first,'utf8');
  const other=await browser.newPage();
  await other.goto(`${BASE}/`);
  const made=await other.evaluate(async()=>(await fetch('/api/new',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({name:'Opened in another tab'})})).status);
  ok(made===200,'a second tab makes and opens another guide');
  const second=path.join(GUIDES,'Opened in another tab.guide-project.json');
  const secondBefore=fs.readFileSync(second,'utf8');
  await page.getByRole('menuitem',{name:/^Write/}).click();
  await page.getByRole('menuitem',{name:/Add module/}).click();
  await fr.locator('#guide-new-title').waitFor({state:'visible',timeout:10000});
  await fr.locator('#guide-new-title').fill('Written by a page that is out of date');
  await fr.locator('#guide-add-module').click();
  await page.waitForFunction(()=>/Nothing more is saved from this page/.test(document.querySelector('#status')?.textContent||''),null,{timeout:10000});
  ok(true,'the page left over from the first tab says it no longer saves');
  ok(fs.readFileSync(first,'utf8')===firstBefore&&fs.readFileSync(second,'utf8')===secondBefore,'and neither guide was written to');
  await other.close();
}finally{
  if(browser)await browser.close();
  server.kill();
}
function files(dir){try{return fs.readdirSync(dir);}catch{return [];}}
console.log(`\n${pass} passed, ${fail} failed (screenshots in ${SHOTS})`);
process.exitCode=fail?1:0;
