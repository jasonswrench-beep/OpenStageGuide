// Real-browser check of the session templates, materials, decks and slides (the tools of milestones 1 to 4 of the product spec of 2026-11-22): the Start and Make menus, the panels in the
// tools frame, what is written into the guide, what is kept in the guide's file, and the Word and PowerPoint files that are made. It needs Playwright, which this repository does not
// depend on: set PLAYWRIGHT_DIR to the playwright package folder, then run  node scripts/check-sessions-browser.mjs  from the repository root. It starts its own server on port 3859
// with a scratch guides folder, and stops it again. Screenshots go to a temporary folder named in the last line.
import {createRequire} from 'node:module';
import {spawn} from 'node:child_process';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import {fileURLToPath} from 'node:url';
import JSZip from 'jszip';
import {parseProject} from '../src/guide-project.mjs';
import {readSessions} from '../src/session-record.mjs';
const require=createRequire(import.meta.url);
if(!process.env.PLAYWRIGHT_DIR){console.error('Set PLAYWRIGHT_DIR to the playwright package folder.');process.exit(2);}
const {chromium}=require(process.env.PLAYWRIGHT_DIR);
const APP=path.resolve(path.dirname(fileURLToPath(import.meta.url)),'..');
const RUN=fs.mkdtempSync(path.join(os.tmpdir(),'osg-sessions-check-'));
const SHOTS=path.join(RUN,'shots'),GUIDES=path.join(RUN,'guides');
fs.mkdirSync(SHOTS);
const PORT=3859,BASE=`http://127.0.0.1:${PORT}`;
let pass=0,fail=0;
const ok=(cond,what)=>{if(cond){pass++;console.log('PASS',what);}else{fail++;console.log('FAIL',what);}};

const server=spawn(process.execPath,['scripts/project-server.mjs','--folder',GUIDES,'--port',String(PORT)],{cwd:APP,stdio:['ignore','pipe','pipe']});
await new Promise((res,rej)=>{let t='';server.stdout.on('data',d=>{t+=d;if(/OpenStage Guide at/.test(t))res();});server.stderr.on('data',d=>rej(new Error(String(d))));setTimeout(()=>rej(new Error('server did not start')),20000);});
const FILE='Sessions check.guide-project.json';
const onDisk=()=>parseProject(fs.readFileSync(path.join(GUIDES,FILE),'utf8')).guide;

let browser;
try{
  browser=await chromium.launch();
  const context=await browser.newContext({acceptDownloads:true,viewport:{width:1360,height:900}});
  const page=await context.newPage();
  const problems=[];
  page.on('pageerror',e=>problems.push(String(e)));
  page.on('console',m=>{if(m.type()==='error'||/Refused|violat/i.test(m.text()))problems.push(m.text());});
  const shot=name=>page.screenshot({path:path.join(SHOTS,name)});
  const saved=async()=>{await page.waitForTimeout(900);await page.waitForFunction(()=>document.querySelector('.bar-context-state')?.getAttribute('data-state')==='saved',null,{timeout:15000});};
  const menu=async(top,item)=>{await page.getByRole('menuitem',{name:new RegExp(`^${top}`)}).click();await page.getByRole('menuitem',{name:item}).click();};
  const download=async(click)=>{const [d]=await Promise.all([page.waitForEvent('download',{timeout:20000}),click()]);const p=path.join(RUN,`dl-${Date.now()}-${d.suggestedFilename().replace(/[^\w.\- ]/g,'_')}`);await d.saveAs(p);return {name:d.suggestedFilename(),path:p};};
  const fr=page.frameLocator('#tools-frame');
  const status=()=>fr.locator('#guide-status').innerText();

  // ---- a guide
  await page.goto(`${BASE}/`);
  await page.click('#welcome-new');
  await page.fill('#dlg-name','Sessions check');
  await page.press('#dlg-name','Enter');
  await fr.locator('#guide-start').waitFor({state:'visible',timeout:15000});
  await fr.locator('#guide-title').fill('Sessions check');
  await fr.locator('#guide-title').press('Enter');
  await fr.locator('#guide-build').waitFor({state:'visible',timeout:15000});

  // ---- session templates
  await menu('Start','Session templates');
  await fr.locator('#guide-sec-templates').waitFor({state:'visible',timeout:10000});
  ok(await fr.locator('#guide-sec-library').isHidden()&&await fr.locator('#guide-sec-module').isHidden(),'the Session templates panel opens on its own');
  ok(/16 templates/.test(await fr.locator('#guide-tpl-count').innerText()),'it lists the 16 templates: '+await fr.locator('#guide-tpl-count').innerText());
  await fr.locator('#guide-tpl-find').fill('retrospective');
  ok(await fr.locator('#guide-tpl-list > li').count()===1,'a search narrows the list to the one template');
  await fr.locator('#guide-tpl-find').fill('');
  await fr.locator('#guide-tpl-family').selectOption('Improve instruction');
  ok(await fr.locator('#guide-tpl-list > li').count()===2,'and so does the choice of what it is for');
  await fr.locator('#guide-tpl-family').selectOption('');
  await fr.getByRole('button',{name:'Preview Skill Practice Workshop'}).click();
  await fr.locator('#guide-tpl-preview').waitFor({state:'visible'});
  ok(await fr.locator('#guide-tpl-preview-body > tr').count()===6,'the preview shows the six rows of the agenda');
  ok(/Revised performance and one transfer action/.test(await fr.locator('#guide-tpl-preview-facts').innerText()),'and what the group leaves with');
  await fr.locator('#guide-tpl-start').fill('9:00');
  await fr.locator('#guide-tpl-start').press('Tab');
  ok(/09:35 to 09:47/.test(await fr.locator('#guide-tpl-preview-body').innerText()),'a start time gives clock times');
  await shot('1-preview.png');
  await fr.getByLabel('As a new module').check();
  await fr.locator('#guide-tpl-title').fill('Feedback practice');
  await fr.getByRole('button',{name:'Apply template'}).click();
  await fr.locator('#guide-status').filter({hasText:'A copy of "Skill Practice Workshop" is in the guide'}).waitFor({timeout:15000});
  ok(true,'Apply puts a copy in the guide: '+(await status()).slice(0,100));
  ok(!/Choose Update guide[^.]*\.[^]*Choose Update guide/.test(await status()),'the message does not say the same thing twice');
  await saved();
  const words=onDisk().paragraphs.map(p=>p.text);
  ok(words.includes('Module 1: Feedback practice'),'the guide holds the new module');
  ok(words.some(w=>/Scenario rehearsal \(18 min\)/.test(w))&&words.includes('Phase: Practice'),'with its activities and their parts');
  ok(words.includes('Template: Skill Practice Workshop (T01, version 1)'),'and says where it came from');
  await shot('2-applied.png');

  // ---- the activity library holds the catalog's activities
  await menu('Activities','Activity library');
  await fr.locator('#guide-sec-library').waitFor({state:'visible',timeout:10000});
  ok(/59 activities ready to use/.test(await fr.locator('#guide-lib-count').innerText()),'the library lists 59 activities: '+await fr.locator('#guide-lib-count').innerText());
  await fr.locator('#guide-lib-phase').selectOption('Decide');
  ok(/7 activities found/.test(await fr.locator('#guide-lib-count').innerText()),'a phase narrows it to seven');
  await fr.locator('#guide-lib-phase').selectOption('');
  ok(/It goes after/.test(await fr.locator('#guide-lib-dest').innerText()),'the panel says where an activity will go: '+await fr.locator('#guide-lib-dest').innerText());

  // ---- session materials
  await menu('Make','Session materials');
  await fr.locator('#guide-sec-materials').waitFor({state:'visible',timeout:10000});
  const run=await download(()=>fr.getByRole('button',{name:'Make the run sheet'}).click());
  const runZip=await JSZip.loadAsync(fs.readFileSync(run.path));
  const runXml=await runZip.file('word/document.xml').async('string');
  ok(/\.docx$/.test(run.name)&&/Run sheet: Sessions check/.test(runXml)&&/Scenario rehearsal \(18 min\)/.test(runXml),'the run sheet is a Word file with the timeline and the activities: '+run.name);
  ok(/<w:tblHeader\/>/.test(runXml),'its tables have header rows');
  ok(await fr.locator('#guide-mat-outputs > li').count()===1&&/Ready:/.test(await fr.locator('#guide-mat-outputs').innerText()),'it is listed as Ready');
  const prep=await download(()=>fr.getByRole('button',{name:'Make the preparation checklist'}).click());
  const prepXml=await (await JSZip.loadAsync(fs.readFileSync(prep.path))).file('word/document.xml').async('string');
  ok(/To print/.test(prepXml)&&/Preparation checklist/.test(prepXml),'the preparation checklist is a Word file of what to print, bring and open');
  await shot('3-materials.png');

  // ---- decks and slides
  await menu('Make','Decks and slides');
  await fr.locator('#guide-sec-decks').waitFor({state:'visible',timeout:10000});
  ok(/^\d+ slides, in this order/.test(await fr.locator('#guide-slide-note').innerText()),'the slides are listed: '+await fr.locator('#guide-slide-note').innerText());
  await fr.locator('#guide-deck-title').fill('Pocket cards');
  await fr.locator('#guide-deck-type').selectOption('facilitator');
  await fr.getByRole('button',{name:'Start a deck'}).click();
  await fr.locator('#guide-deck-editor').waitFor({state:'visible',timeout:10000});
  ok(/Pocket cards \(6 cards\)/.test(await fr.locator('#guide-deck-editor-title').innerText()),'a deck of pocket cards is started with a card for each activity');
  await fr.locator('#guide-deck-front').fill('Extra card\nA second line');
  await fr.locator('#guide-deck-back').fill('Back of the extra card');
  await fr.getByRole('button',{name:'Add this as a new card'}).click();
  await fr.locator('#guide-deck-editor-title').filter({hasText:'7 cards'}).waitFor({timeout:10000});
  ok(true,'a card can be added');
  await fr.locator('#guide-deck-paper').selectOption('A4');
  await fr.locator('#guide-deck-marks').check();
  await page.waitForTimeout(600);
  const sheet=await download(()=>fr.getByRole('button',{name:'Make the print sheet as a Word file'}).click());
  const sheetXml=await (await JSZip.loadAsync(fs.readFileSync(sheet.path))).file('word/document.xml').async('string');
  ok(/w:pgSz w:w="11906"/.test(sheetXml)&&/Extra card/.test(sheetXml)&&/Crop mark/.test(sheetXml),'the print sheet is on A4, with the card and the crop marks: '+sheet.name);
  const slides=await download(()=>fr.getByRole('button',{name:'Make the slides as a PowerPoint file'}).click());
  const deckZip=await JSZip.loadAsync(fs.readFileSync(slides.path));
  ok(/\.pptx$/.test(slides.name)&&!!deckZip.file('ppt/slides/slide1.xml')&&!!deckZip.file('ppt/notesSlides/notesSlide1.xml'),'the slides are a PowerPoint file with notes: '+slides.name);
  await shot('4-decks.png');

  // ---- what is kept in the guide's file
  await saved();
  const kept=readSessions(onDisk().sessions);
  ok(kept.outputs.map(o=>o.kind).sort().join(',')==='cards,checklist,runsheet,slides','the guide\'s file remembers what was made: '+kept.outputs.map(o=>o.kind).sort().join(','));
  ok(kept.decks.length===1&&kept.decks[0].cards.length===7&&kept.decks[0].paper==='A4'&&kept.decks[0].cropMarks===true,'and the deck, with its seven cards on A4 with crop marks');
  ok(!/^(Phase|Source):/m.test(onDisk().title||''),'the guide itself holds only its paragraphs and the record');

  // ---- the guide changes: the materials say so
  await menu('Start','Session templates');
  await fr.locator('#guide-sec-templates').waitFor({state:'visible'});
  await fr.getByRole('button',{name:'Preview Case Discussion'}).click();
  await fr.locator('#guide-tpl-preview').waitFor({state:'visible'});
  await fr.getByLabel('As a new module').check();
  await fr.locator('#guide-tpl-title').fill('Case morning');
  await fr.getByRole('button',{name:'Apply template'}).click();
  await fr.locator('#guide-status').filter({hasText:'A copy of "Case Discussion"'}).waitFor({timeout:15000});
  await menu('Make','Session materials');
  await fr.locator('#guide-sec-materials').waitFor({state:'visible'});
  const text=await fr.locator('#guide-mat-outputs').innerText();
  ok(/Review needed/.test(text),'after the guide changes, what was made says Review needed');
  await fr.getByRole('button',{name:/^What changed .*Run sheet/}).first().click();
  ok(/is new in the guide/.test(await fr.locator('#guide-mat-changes').innerText()),'and names what changed: '+(await fr.locator('#guide-mat-changes').innerText()).split('\n')[0]);
  await shot('5-review-needed.png');

  ok(problems.length===0,'no page error and no refused script or style: '+problems.slice(0,3).join(' | '));
}finally{
  await browser?.close();
  server.kill();
}
console.log(`\n${pass} passed, ${fail} failed (screenshots in ${SHOTS})`);
process.exit(fail?1:0);
