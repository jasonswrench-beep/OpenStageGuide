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
import {serializeProject} from '../src/guide-project.mjs';
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
  ok((await page.locator('[role=menubar] > .menubar-slot').count())===14,'fourteen menus: File, Edit, Insert, Format, View, Start, Sections, Activities, Slides, Assessment, eLearning, Check & Plan, Make, Help');
  ok(await page.locator('#bar-buttons button').count()===0,'no guide buttons on the navy bar before a guide is open');
  ok(/no guides yet/i.test(await page.locator('.recent').innerText()),'it says there are no guides yet');
  await shot('1-welcome.png');

  // ---- a tool cannot be used with no guide
  await page.getByRole('menuitem',{name:/^Sections/}).click();
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
  await page.getByRole('menuitem',{name:/^Sections/}).click();
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
  // (a workbook needs something for learners in the guide: a highlight block, added from the Activities menu)
  await page.getByRole('menuitem',{name:/^Activities/}).click();
  await page.getByRole('menuitem',{name:/Instructional cues/i}).click();
  await fr.locator('#guide-cue-highlight').evaluate(e=>{e.checked=true;e.dispatchEvent(new Event('change',{bubbles:true}));});
  await fr.locator('#guide-block-text').fill('Name the feeling first.');
  await fr.locator('#guide-block-form').evaluate(f=>f.requestSubmit());
  await page.waitForFunction(()=>/Name the feeling first\./.test(document.querySelector('#work-guide')?.innerText||''),null,{timeout:10000});
  ok(true,'a block added in the tools appears in the guide');
  // its kind's picture is drawn in front of its name, from the program's own icon file, and it stays while the block's words are changed
  const iconOf=()=>page.evaluate(()=>{const row=[...document.querySelectorAll('#work-guide .g-cue')].find(r=>/Name the feeling first/.test(r.textContent||''));const img=row?.querySelector('img.g-icon');return img?{src:img.getAttribute('src'),alt:img.getAttribute('alt'),loaded:img.complete&&img.naturalWidth>0,w:Math.round(img.getBoundingClientRect().width),first:row.firstElementChild===img}:null;});
  let blockIcon=await iconOf();
  ok(!!blockIcon&&/^\/assets\/icons-doc\/highlight\.png$/.test(blockIcon.src)&&blockIcon.first,'the block has its kind\'s picture at the front: '+JSON.stringify(blockIcon));
  ok(!!blockIcon&&blockIcon.loaded&&blockIcon.w>16&&blockIcon.alt==='','and the picture is really there (it loads), a little bigger than the words, with no description of its own because the name beside it says it');
  await page.locator('#work-guide .g-cue',{hasText:'Name the feeling first'}).locator('.g-label').click();
  await page.waitForFunction(()=>document.activeElement?.isContentEditable===true,null,{timeout:5000});
  ok(await page.evaluate(()=>!!document.querySelector('#work-guide .g-editing > img.g-icon')),'while its words are changed the picture stays at the front');
  await page.keyboard.press('Escape');
  await page.waitForTimeout(300);
  ok(!!(await iconOf()),'and is there again afterwards');

  // ---- the menus of the tools, as the author listed them (2026-11-14), and the tools that are new
  const MENUS={Start:['Guide overview','Course settings','Course outline'],Sections:['Add module','Learning outcomes','Add lesson','Lesson details'],Activities:['Add activity','Instructional cues','Activity library','Forms'],Slides:['Import slides','Update from slides'],Assessment:['Assessments','Learning alignment'],eLearning:['E-learning screens','E-learning cues'],'Check & Plan':['Review guide','Timing and agenda','Accessibility'],Make:['Delivery versions','Appearance','Generate materials','Reference pages','Participant workbook']};
  for(const [menu,names] of Object.entries(MENUS)){
    await page.getByRole('menuitem',{name:new RegExp('^'+menu.replace(/[&]/g,'\\$&'))}).click();
    const got=(await page.locator('[role=menu]:visible [role=menuitem]').allInnerTexts()).map(t=>t.replace(/\s+/g,' ').trim());
    ok(JSON.stringify(got)===JSON.stringify(names),menu+' menu: '+got.join(', '));
    await page.keyboard.press('Escape');
  }
  // Where each list lands, not only what it says: at the window's starting width (1360) and at 1140 and 1280 every list, the last menus' (Make, Help) included,
  // has to sit inside the page with its words whole. They used to hang from the left edge of their menu's name and ran off the right edge.
  const startSize=page.viewportSize();
  const MENU_NAMES=['File','Edit','Insert','Format','View','Start','Sections','Activities','Slides','Assessment','eLearning','Check','Make','Help'];
  for(const width of [1360,1280,1140]){
    await page.setViewportSize({width,height:startSize.height});await page.waitForTimeout(250);
    const out=[];
    for(const name of MENU_NAMES){
      await page.getByRole('menuitem',{name:new RegExp('^'+name)}).first().click();
      const box=await page.evaluate(()=>{const m=[...document.querySelectorAll('[role=menu]')].find(e=>e.getBoundingClientRect().width>0);if(!m)return null;const b=m.getBoundingClientRect();return {left:Math.round(b.left),right:Math.round(b.right),page:document.documentElement.clientWidth,cut:[...m.querySelectorAll('[role=menuitem],[role=menuitemcheckbox]')].filter(e=>e.scrollWidth>e.clientWidth+1).length};});
      if(!box||box.left<0||box.right>box.page||box.cut)out.push(name+' '+JSON.stringify(box));
      await page.keyboard.press('Escape');
    }
    ok(out.length===0,'at '+width+' px wide every menu list is inside the page with its words whole'+(out.length?': '+out.join('; '):''));
  }
  await page.setViewportSize(startSize);await page.waitForTimeout(250);
  const openFromMenu=async(menu,item)=>{await page.getByRole('menuitem',{name:new RegExp('^'+menu)}).click();await page.getByRole('menuitem',{name:item}).click();await page.waitForTimeout(300);};
  await openFromMenu('Sections',/Add lesson/);
  await fr.locator('#guide-lesson-title').waitFor({state:'visible',timeout:10000});
  ok(await fr.locator('#guide-new-title').isHidden(),'Add lesson shows the lesson\'s title and minutes and not the module\'s title');
  await openFromMenu('Sections',/Add module/);
  await fr.locator('#guide-new-title').waitFor({state:'visible',timeout:10000});
  ok(await fr.locator('#guide-lesson-title').isHidden(),'and Add module the module\'s title and not the lesson\'s');
  await openFromMenu('Sections',/Learning outcomes/);
  await fr.locator('#guide-cue-current',{hasText:/^Objective$/}).waitFor({timeout:10000});
  ok((await fr.locator('#guide-cue-current').innerText())==='Objective','Learning outcomes opens the block form on the Objective kind');
  await openFromMenu('eLearning',/E-learning cues/);
  await fr.locator('#guide-cue-current',{hasText:/^Screen title$/}).waitFor({timeout:10000});
  ok((await fr.locator('#guide-cue-current').innerText())==='Screen title','E-learning cues opens it on the e-learning kinds, at Screen title');
  await openFromMenu('Check',/Accessibility/);
  await fr.locator('#guide-review-heading',{hasText:/^Accessibility$/}).waitFor({timeout:15000});
  ok((await fr.locator('#guide-review-heading').innerText())==='Accessibility'&&await fr.locator('#guide-review-pdf-part').isHidden(),'Accessibility opens the document checks, without Save as PDF');
  await openFromMenu('Check',/Review guide/);
  await fr.locator('#guide-review-heading',{hasText:/^Review guide$/}).waitFor({timeout:15000});
  ok((await fr.locator('#guide-review-heading').innerText())==='Review guide','and Review guide the guide\'s own');
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
  ok(/Version \d+\.\d+\.\d+, build [0-9a-f]{8}\./.test(about.text)&&!/prototype/i.test(about.text),'About says the version and the build');
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
  await page.getByRole('menuitem',{name:/^Sections/}).click();
  await page.getByRole('menuitem',{name:/Add module/}).click();
  await fr.locator('#guide-new-title').waitFor({state:'visible',timeout:10000});
  await fr.locator('#guide-new-title').fill('Written by a page that is out of date');
  await fr.locator('#guide-add-module').click();
  await page.waitForFunction(()=>/Nothing more is saved from this page/.test(document.querySelector('#status')?.textContent||''),null,{timeout:10000});
  ok(true,'the page left over from the first tab says it no longer saves');
  ok(fs.readFileSync(first,'utf8')===firstBefore&&fs.readFileSync(second,'utf8')===secondBefore,'and neither guide was written to');
  await other.close();

  // ---- a long guide scrolls inside its own pane and never makes the whole page scroll (a paragraph for screen readers inside the guide, positioned
  // absolutely, once stretched the page by what the guide was longer than the window, so scrolling past the end of the guide moved the page and left
  // empty space below the guide and the tools)
  fs.writeFileSync(path.join(GUIDES,'Long guide.guide-project.json'),serializeProject({title:'Long guide',paragraphs:Array.from({length:150},(_,i)=>({xml:`<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:r><w:t xml:space="preserve">Paragraph ${i+1} of a long guide.</w:t></w:r></w:p>`,text:`Paragraph ${i+1} of a long guide.`,style:'Normal'}))}));
  const long=await browser.newPage({viewport:{width:1920,height:1009}});
  await long.goto(`${BASE}/`);
  await long.waitForSelector('#welcome-new,#work-guide');
  if(await long.locator('#work-guide').count()){
    await long.getByRole('menuitem',{name:/^File/}).click();
    await long.getByRole('menuitem',{name:/Close guide/}).click();
    await long.waitForSelector('#welcome-new');
  }
  await long.getByRole('button',{name:/Long guide/}).click();
  await long.waitForSelector('#work-guide .page');
  const layout=()=>long.evaluate(()=>{const g=document.querySelector('#work-guide'),w=document.querySelector('.work').getBoundingClientRect();return {inner:innerHeight,page:document.scrollingElement.scrollHeight,pageTop:document.scrollingElement.scrollTop,guideInner:g.scrollHeight,guideShown:g.clientHeight,workBottom:Math.round(w.bottom)};});
  for(const [w,h] of [[1920,1009],[1360,900],[1100,700]]){
    await long.setViewportSize({width:w,height:h});
    await long.waitForTimeout(300);
    const m=await layout();
    ok(m.guideInner>m.guideShown+500,`${w} x ${h}: the guide is longer than its pane, so it scrolls inside the pane`);
    ok(m.page<=m.inner,`${w} x ${h}: the page itself does not scroll (page ${m.page}, window ${m.inner})`);
    ok(m.workBottom===m.inner,`${w} x ${h}: the guide and the tools reach the bottom of the window (${m.workBottom} of ${m.inner})`);
  }
  await long.setViewportSize({width:1920,height:1009});
  await long.locator('#work-guide').evaluate(e=>{e.scrollTop=e.scrollHeight;});
  await long.mouse.move(700,500);
  await long.mouse.wheel(0,3000);
  await long.waitForTimeout(400);
  ok((await layout()).pageTop===0,'scrolling on past the end of the guide does not move the page');

  // ---- zoom: the guide's page is drawn bigger or smaller, and nothing else is (the menus and the tools keep their size)
  const zoomSizes=()=>long.evaluate(()=>{const p=document.querySelector('#work-guide .page'),q=e=>Math.round(document.querySelector(e).getBoundingClientRect().height*10)/10,w=e=>Math.round(document.querySelector(e).getBoundingClientRect().width*10)/10;
    return {zoom:getComputedStyle(p).zoom,pageWidth:Math.round(p.getBoundingClientRect().width),menus:q('#menu-row'),bar:q('.shell'),tools:w('.work-tools'),status:document.querySelector('#status').textContent,docScroll:document.scrollingElement.scrollHeight,inner:innerHeight};});
  const settingsNow=()=>long.evaluate(async()=>(await (await fetch('/api/settings')).json()).values['osg.zoom.v1']??null);
  const menuView=async name=>{await long.getByRole('menuitem',{name:/^View/}).click();await long.getByRole('menuitem',{name}).click();await long.waitForTimeout(250);};
  const s100=await zoomSizes();
  ok(s100.zoom==='1'&&s100.pageWidth>=780&&s100.pageWidth<=782,'the guide\'s page starts at 100% ('+JSON.stringify([s100.zoom,s100.pageWidth])+')');
  await menuView(/Zoom in/);
  let z=await zoomSizes();
  ok(z.zoom==='1.1'&&Math.abs(z.pageWidth-Math.round(s100.pageWidth*1.1))<=2,'View, Zoom in draws the guide\'s page at 110%: '+z.pageWidth+' wide');
  ok(z.menus===s100.menus&&z.bar===s100.bar&&z.tools===s100.tools,'the menus, the navy bar and the tools keep their size');
  ok(/Zoom 110%/.test(z.status),'it says so: '+z.status);
  await long.keyboard.press('Control+Alt+=');z=await zoomSizes();
  ok(z.zoom==='1.25','Ctrl+Alt+= goes on to 125%');
  await long.keyboard.press('Control+Alt+-');z=await zoomSizes();
  ok(z.zoom==='1.1','Ctrl+Alt+- goes back to 110%');
  await long.mouse.move(700,500);await long.keyboard.down('Control');await long.mouse.wheel(0,-120);await long.keyboard.up('Control');await long.waitForTimeout(250);
  ok((await zoomSizes()).zoom==='1.25','the mouse wheel with Ctrl held, over the guide, zooms in');
  await long.mouse.move(1700,500);await long.keyboard.down('Control');await long.mouse.wheel(0,-120);await long.keyboard.up('Control');await long.waitForTimeout(250);
  ok((await zoomSizes()).zoom==='1.25','but not over the tools');
  await long.locator('#tools-frame').waitFor({state:'visible'});
  const toolsFrame=long.frames().find(f=>f.url().includes('/program/tools'));
  const inTools=toolsFrame.locator('button:visible, input:visible, select:visible').first();
  await inTools.waitFor({state:'visible',timeout:10000});
  await inTools.focus();
  await long.keyboard.press('Control+Alt+=');await long.waitForTimeout(300);
  ok((await zoomSizes()).zoom==='1.5','Ctrl+Alt+= pressed in the tools zooms the guide\'s page');
  await long.keyboard.press('Control+=');await long.waitForTimeout(200);
  ok((await zoomSizes()).zoom==='1.5','Ctrl+= does not zoom (it is subscript in the guide\'s words)');
  ok(await settingsNow()==='1.5','the size is kept in the program\'s settings');
  ok((await zoomSizes()).docScroll<=(await zoomSizes()).inner,'a bigger page scrolls inside the guide\'s pane, not the window');
  // a paragraph can still be chosen and its words changed on a page that is drawn bigger
  await long.locator('#work-guide .page .g-para').nth(4).click();
  await long.keyboard.press('Home');
  await long.keyboard.type('ZOOMED ');
  await long.keyboard.press('Escape');
  await long.waitForTimeout(500);
  ok(/ZOOMED Paragraph 5 of a long guide\./.test(await long.locator('#work-guide .page').innerText()),'on a page drawn at 150% a paragraph is chosen by clicking it and its words are changed');
  await long.reload();
  await long.waitForSelector('#work-guide .page');
  ok((await zoomSizes()).zoom==='1.5','after the page is loaded again the size is still 150%');
  // the same zoom, as buttons at the right end of the formatting bar
  const zbar=long.getByRole('toolbar',{name:'Text formatting'});
  const zlevel=zbar.locator('.format-zoom-level');
  ok(await zbar.getByRole('group',{name:'Zoom'}).count()===1,'the formatting bar has a Zoom group');
  ok((await zlevel.innerText())==='150%','it shows the size the page is drawn at: 150%');
  const rightEdge=await long.evaluate(()=>{const b=document.querySelector('.format-bar').getBoundingClientRect(),g=document.querySelector('.format-zoom').getBoundingClientRect();return b.right-g.right;});
  ok(rightEdge>=0&&rightEdge<24,'at the right end of the bar ('+Math.round(rightEdge)+' px from its edge)');
  await zbar.getByRole('button',{name:'Zoom in'}).click();await long.waitForTimeout(250);
  ok((await zoomSizes()).zoom==='1.75'&&(await zlevel.innerText())==='175%','the Zoom in button goes to 175% and the bar says so');
  await zbar.getByRole('button',{name:'Zoom out'}).click();await long.waitForTimeout(250);
  ok((await zoomSizes()).zoom==='1.5'&&(await zlevel.innerText())==='150%','the Zoom out button goes back to 150%');
  await long.keyboard.press('Control+Alt+=');await long.waitForTimeout(250);
  ok((await zlevel.innerText())==='175%','the number follows the keys and the menu too');
  ok(/Zoom 175%/.test(await zlevel.getAttribute('aria-label')),'and the button says it to a screen reader: '+await zlevel.getAttribute('aria-label'));
  await zlevel.click();await long.waitForTimeout(250);
  ok((await zoomSizes()).zoom==='1'&&(await zlevel.innerText())==='100%','the number is a button: it goes back to 100%');
  await long.keyboard.press('Control+Alt+=');await long.waitForTimeout(150);
  ok((await zlevel.innerText())==='110%','(and on again)');
  await long.keyboard.press('Control+Alt+0');
  z=await zoomSizes();
  ok(z.zoom==='1'&&z.pageWidth===s100.pageWidth&&await settingsNow()===null,'Ctrl+Alt+0 is back to the actual size, which is not kept as a setting');
  await menuView(/Zoom out/);
  ok((await zoomSizes()).zoom==='0.9','View, Zoom out draws the page smaller');
  await long.getByRole('menuitem',{name:/^File/}).click();
  await long.getByRole('menuitem',{name:/Close guide/}).click();
  await long.waitForSelector('#welcome-new');
  await long.getByRole('menuitem',{name:/^View/}).click();
  ok((await long.getByRole('menuitem',{name:/Zoom in/}).getAttribute('aria-disabled'))==='true','with no guide open the zoom commands are dimmed');
  await long.keyboard.press('Escape');
  await long.close();

  // ---- the Forms tool, in a guide of its own so nothing above depends on it
  const fpage=await browser.newPage({viewport:{width:1360,height:900}});
  await fpage.goto(`http://127.0.0.1:${PORT}/`);
  await fpage.click('#welcome-new');
  await fpage.fill('#dlg-name','Forms check');
  await fpage.press('#dlg-name','Enter');
  await fpage.waitForSelector('#tools-frame');
  const ff=fpage.frameLocator('#tools-frame');
  await ff.locator('#guide-title').fill('Forms check');
  await ff.locator('#guide-title').press('Enter');
  await ff.locator('#guide-build').waitFor({state:'visible',timeout:15000});
  await fpage.getByRole('menuitem',{name:/^Activities/}).click();
  ok((await fpage.locator('[role=menu]:visible [role=menuitem]').allInnerTexts()).map(t=>t.replace(/\s+/g,' ').trim()).includes('Forms'),'Forms is in the Activities menu');
  await fpage.getByRole('menuitem',{name:/^Forms$/}).click();
  await ff.locator('#guide-forms-summary').waitFor({state:'visible',timeout:10000});
  await fpage.waitForTimeout(500);
  ok(/^There are no forms in this guide yet/.test(await ff.locator('#guide-forms-summary').innerText()),'the Forms tool opens from the menu and says the guide has no forms');
  ok((await ff.locator('#tool-title').innerText())==='Forms','its heading says what its menu says');
  await ff.locator('#guide-forms-new').click();
  const dlg=fpage.locator('dialog[open]');
  await dlg.waitFor({timeout:10000});
  ok(/New form/.test(await dlg.innerText()),'New form in the tool opens the New form dialog of the program');
  await dlg.getByLabel('Title of the form').fill('Goals');
  await dlg.getByRole('button',{name:'Insert'}).click();
  await ff.locator('#guide-forms-list li').first().waitFor({timeout:15000});
  ok(/Goals/.test(await ff.locator('#guide-forms-list li').first().innerText()),'the new form is listed: '+(await ff.locator('#guide-forms-list li').first().innerText()).replace(/\s+/g,' '));
  ok(/Print: 1 per person/.test(await ff.locator('#guide-forms-list li').first().innerText()),'with the print line the template gave it');
  ok(/cannot be added up yet/.test(await ff.locator('#guide-forms-copies-summary').innerText()),'which cannot be added up until the number of people is said');
  await ff.locator('#guide-forms-people').fill('18');
  await ff.locator('#guide-forms-people').press('Tab');
  await ff.locator('#guide-forms-copies-summary',{hasText:/18 copies to make in all/}).waitFor({timeout:10000});
  ok(true,'saying there will be 18 people adds the print line up to 18 copies');
  await fpage.screenshot({path:path.join(SHOTS,'forms-panel.png')});
  // a Handout block cites the form by choosing it
  await fpage.getByRole('menuitem',{name:/^Activities/}).click();
  await fpage.getByRole('menuitem',{name:/Instructional cues/}).click();
  await ff.locator('#guide-cue-toggle').click();
  await ff.locator('#guide-cue-group').selectOption('');
  await ff.locator('label.guide-cue:has(#guide-cue-handout)').click();
  await ff.locator('#guide-block-which-part').waitFor({state:'visible',timeout:10000});
  ok((await ff.locator('#guide-block-which option').allInnerTexts()).some(t=>/Goals/.test(t)),'the Handout block offers the form to cite');
  await ff.locator('#guide-block-which').selectOption({index:1});
  await ff.locator('#guide-block-text').fill('Give out');
  await ff.locator('#guide-block-go').click();
  await fpage.waitForFunction(()=>/Give out\s+Handout: Goals/.test(document.querySelector('#work-guide')?.innerText||''),null,{timeout:15000});
  ok(true,'and the block on the page says Give out, then the name of the form');
  await fpage.getByRole('menuitem',{name:/^Check/}).click();
  await fpage.getByRole('menuitem',{name:/Review guide/}).click();
  await ff.locator('#guide-review-heading',{hasText:/^Review guide$/}).waitFor({timeout:15000});
  ok(!/how many people there will be/.test(await ff.locator('#guide-review-look').innerText()),'and the Review no longer asks how many people there will be');
  // the handout pack: the forms for everyone, as a Word file, with no print line
  await fpage.getByRole('menuitem',{name:/^Activities/}).click();
  await fpage.getByRole('menuitem',{name:/^Forms$/}).click();
  await ff.locator('#guide-forms-pack-scope').waitFor({state:'visible',timeout:10000});
  await ff.locator('#guide-forms-pack-scope').selectOption('guide');
  const [packDownload]=await Promise.all([fpage.waitForEvent('download',{timeout:20000}),ff.locator('#guide-forms-pack').click()]);
  const packPath=path.join(RUN,'pack.docx');
  await packDownload.saveAs(packPath);
  ok(/ - handouts\.docx$/.test(packDownload.suggestedFilename()),'Make handout pack gives a Word file: '+packDownload.suggestedFilename());
  const packXml=await (await JSZip.loadAsync(fs.readFileSync(packPath))).file('word/document.xml').async('string');
  ok(/Handout[^<]*Goals/.test(packXml)&&!/Print:/.test(packXml),'it holds the form and not its print line');
  // Ctrl+Z pressed in the tools, on a button, is the guide's Undo: it takes back the block that cited the form
  ok(/Give out\s+Handout/.test(await fpage.locator('#work-guide').innerText()),'the block that cites the form is on the page');
  await ff.locator('#guide-forms-new').focus();
  await fpage.keyboard.press('Control+z');
  await fpage.waitForFunction(()=>!/Give out\s+Handout/.test(document.querySelector('#work-guide')?.innerText||''),null,{timeout:10000});
  ok(true,'Ctrl+Z pressed in the tools takes back the last change to the guide');
  // File > Save a backup is a copy of the guide's own file, and File > Restore a backup opens it as a new guide
  await fpage.getByRole('menuitem',{name:/^File/}).click();
  const [backupDownload]=await Promise.all([fpage.waitForEvent('download',{timeout:20000}),fpage.getByRole('menuitem',{name:/Save a backup/}).click()]);
  const backupPath=path.join(RUN,'backup.guide-project.json');
  await backupDownload.saveAs(backupPath);
  ok(/ backup \d{4}-\d\d-\d\d\.guide-project\.json$/.test(backupDownload.suggestedFilename()),'Save a backup gives a file named for the guide and the day: '+backupDownload.suggestedFilename());
  const backupJson=JSON.parse(fs.readFileSync(backupPath,'utf8'));
  ok(backupJson.format&&Array.isArray(backupJson.guide?.paragraphs)&&backupJson.guide.paragraphs.some(p=>/Goals/.test(p.text)),'it is the guide as a project file, with the form in it');
  await fpage.getByRole('menuitem',{name:/^File/}).click();
  await fpage.getByRole('menuitem',{name:/Restore a backup/}).click();
  await fpage.locator('dialog[open] input[type=file]').setInputFiles(backupPath);
  await fpage.locator('dialog[open]').getByRole('button',{name:'Read the backup'}).click();
  await fpage.locator('dialog[open] #dlg-name').waitFor({timeout:10000});
  ok(/restored/.test(await fpage.locator('dialog[open] #dlg-name').inputValue()),'the restored guide is offered a name of its own: '+await fpage.locator('dialog[open] #dlg-name').inputValue());
  await fpage.locator('dialog[open]').getByRole('button',{name:'Restore'}).click();
  await fpage.waitForFunction(()=>/Goals/.test(document.querySelector('#work-guide')?.innerText||''),null,{timeout:20000});
  ok(files(GUIDES).some(f=>/\(restored\)\.guide-project\.json$/.test(f))&&files(GUIDES).some(f=>f==='Forms check.guide-project.json'),'the backup is a new guide beside the first, which is not changed');
  const nowOpen=await fpage.evaluate(()=>fetch('/api/state').then(r=>r.json()));
  ok(/\(restored\)\.guide-project\.json$/.test(String(nowOpen.current)),'and it is the one that is open: '+nowOpen.current);
  await fpage.screenshot({path:path.join(SHOTS,'forms-tool.png')});
  await fpage.close();
}finally{
  if(browser)await browser.close();
  server.kill();
}
function files(dir){try{return fs.readdirSync(dir);}catch{return [];}}
console.log(`\n${pass} passed, ${fail} failed (screenshots in ${SHOTS})`);
process.exitCode=fail?1:0;
