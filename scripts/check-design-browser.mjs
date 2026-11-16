// Real-browser check of the design desk (program/ui/design-desk.js): the Design menu, the five screens, what is kept in the guide, and the files it makes. It needs Playwright,
// which this repository does not depend on: set PLAYWRIGHT_DIR to the playwright package folder, then run  node scripts/check-design-browser.mjs  from the repository root.
// It starts its own server on port 3856 with a scratch guides folder, and stops it again. Screenshots go to a temporary folder named in the last line.
import {createRequire} from 'node:module';
import {spawn} from 'node:child_process';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import {fileURLToPath} from 'node:url';
import JSZip from 'jszip';
import {checkPstat,readPstat} from '@openstage/pstat';
import {parseProject} from '../src/guide-project.mjs';
import {readDesign} from '../src/design-model.mjs';
const require=createRequire(import.meta.url);
if(!process.env.PLAYWRIGHT_DIR){console.error('Set PLAYWRIGHT_DIR to the playwright package folder.');process.exit(2);}
const {chromium}=require(process.env.PLAYWRIGHT_DIR);
const APP=path.resolve(path.dirname(fileURLToPath(import.meta.url)),'..');
const RUN=fs.mkdtempSync(path.join(os.tmpdir(),'osg-design-check-'));
const SHOTS=path.join(RUN,'shots'),GUIDES=path.join(RUN,'guides');
fs.mkdirSync(SHOTS);
const PORT=3856,BASE=`http://127.0.0.1:${PORT}`;
let pass=0,fail=0;
const ok=(cond,what)=>{if(cond){pass++;console.log('PASS',what);}else{fail++;console.log('FAIL',what);}};

const server=spawn(process.execPath,['scripts/project-server.mjs','--folder',GUIDES,'--port',String(PORT)],{cwd:APP,stdio:['ignore','pipe','pipe']});
await new Promise((res,rej)=>{let t='';server.stdout.on('data',d=>{t+=d;if(/OpenStage Guide at/.test(t))res();});server.stderr.on('data',d=>rej(new Error(String(d))));setTimeout(()=>rej(new Error('server did not start')),20000);});
const FILE='Design check.guide-project.json';
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
  const desk=page.locator('#design-desk');
  const download=async(click)=>{const [d]=await Promise.all([page.waitForEvent('download',{timeout:20000}),click()]);const p=path.join(RUN,`dl-${Date.now()}-${d.suggestedFilename().replace(/[^\w.\- ]/g,'_')}`);await d.saveAs(p);return {name:d.suggestedFilename(),path:p};};

  // ---- a guide, and the desk over it
  await page.goto(`${BASE}/`);
  await page.click('#welcome-new');
  await page.fill('#dlg-name','Design check');
  await page.press('#dlg-name','Enter');
  const fr=page.frameLocator('#tools-frame');
  await fr.locator('#guide-start').waitFor({state:'visible',timeout:15000});
  await fr.locator('#guide-title').fill('Design check');
  await fr.locator('#guide-title').press('Enter');
  await fr.locator('#guide-build').waitFor({state:'visible',timeout:15000});
  ok(await page.getByRole('menuitem',{name:/^Design/}).count()===1,'there is a Design menu');
  await page.getByRole('menuitem',{name:/^Design/}).click();
  ok((await page.getByRole('menuitem').filter({hasText:/\.\.\.$/}).allInnerTexts()).filter(t=>/Needs and objectives|Models and activities|Course map|Storyboard|Evaluation plan/.test(t)).length===5,'it has the five screens');
  await page.getByRole('menuitem',{name:/Needs and objectives/}).click();
  await desk.waitFor({timeout:10000});
  ok(await page.locator('#work-guide').isHidden()&&await page.locator('.work-tools').isHidden(),'the desk takes the place of the guide and its tools');
  ok((await desk.getByRole('tab').allInnerTexts()).join('|')==='Needs and objectives|Models and activities|Course map|Storyboard|Evaluation plan','five tabs, in order');
  ok(await desk.getByRole('tab',{name:'Needs and objectives'}).getAttribute('aria-selected')==='true','the screen asked for is the one showing');
  await shot('1-needs.png');

  // ---- needs and objectives
  await desk.getByLabel('What is the problem, or the chance?').fill('New supervisors avoid hard conversations.');
  await desk.getByLabel('Who are the learners?').fill('First-year supervisors');
  await desk.getByLabel('They know, and have had too little practice').check();
  ok(/Training can help with this/.test(await desk.locator('.dd-verdict').innerText()),'the cause chosen says whether training can help');
  await desk.getByLabel('The tools, process or information are in the way').check();
  ok(/Training can help with part of it/.test(await desk.locator('.dd-verdict').innerText()),'and says so when only part of it is training');
  await desk.getByRole('button',{name:'Add an objective'}).click();
  await desk.getByLabel(/^Objective 1: what learners will be able to do/).fill('Learners will understand the escalation policy.');
  ok(/"understand" names something nobody can see a learner do\. Try "explain"/.test(await desk.locator('.dd-issues').first().innerText()),'a weak verb is pointed out, with verbs to use');
  await desk.getByLabel(/^Objective 1: what learners will be able to do/).fill('Learners will be able to explain the three steps of a coaching conversation.');
  ok(/Level: Understand \(from the verb\)/.test(await desk.locator('.dd-level').first().innerText())&&await desk.locator('.dd-issues').first().innerText()==='','a good verb gives its level, and nothing to say');
  await desk.locator('summary',{hasText:'Write one from the verbs'}).click();
  await desk.getByLabel('Level',{exact:true}).last().selectOption('apply');
  await desk.getByLabel('Verb').selectOption('demonstrate');
  await desk.getByLabel('What',{exact:true}).fill('a calm opening to a hard conversation');
  ok(/^Learners will be able to demonstrate a calm opening to a hard conversation\.$/.test(await desk.locator('.dd-preview').innerText()),'an objective written from the verbs is shown as it will read');
  await desk.getByRole('button',{name:'Add this objective'}).click();
  ok(await desk.getByLabel(/^Objective 2: what learners will be able to do/).inputValue()==='Learners will be able to demonstrate a calm opening to a hard conversation.','and added');
  ok(/2 objectives: Understand 1, Apply 1/.test(await desk.locator('.dd-spread').innerText()),'the levels are counted');
  await saved();
  const kept=readDesign(onDisk().design);
  ok(kept.needs.problem==='New supervisors avoid hard conversations.'&&kept.needs.causes.join()==='practice,tools'&&kept.objectives.length===2,'the notes are in the guide\'s file');
  ok(kept.objectives[1].level==='apply','including the level chosen');

  // ---- back to the guide, and again
  await desk.getByRole('button',{name:'Back to the guide'}).click();
  ok(await page.locator('#design-desk').count()===0&&await page.locator('#work-guide').isVisible()&&await page.locator('.work-tools').isVisible(),'Back to the guide gives the guide and its tools back');
  await menu('Design',/Needs and objectives/);
  ok(await desk.getByLabel('Who are the learners?').inputValue()==='First-year supervisors','opened again, the notes are as they were');

  // ---- models and activities
  await desk.getByRole('tab',{name:'Models and activities'}).click();
  ok(await desk.locator('.dd-nav button',{hasText:'Nine events of instruction'}).count()===1&&await desk.locator('.dd-nav button').count()>=9,'the models are listed');
  ok(await desk.locator('.dd-nav button',{hasText:'Nine events of instruction'}).getAttribute('aria-pressed')==='true','the first is chosen');
  ok(await desk.locator('.dd-events > li').count()===9,'the nine events of the model chosen');
  ok(/Gagné, R\. M\. \(1985\)\. The conditions of learning and theory of instruction \(4th ed\.\)\. Holt, Rinehart and Winston\./.test(await desk.locator('.dd-ref').first().innerText()),'and the work it comes from, in APA 7');
  await desk.getByRole('button',{name:/^Add the activity: A real problem or a short story/}).click();
  ok(await desk.getByLabel('Activity 1: title').inputValue()==='A real problem or a short story','an activity is added to the plan');
  await desk.getByLabel('Minutes').first().fill('10');
  await desk.getByRole('group',{name:'Activity 1 serves these objectives'}).getByLabel(/Objective 1/).check();
  await desk.locator('.dd-nav button',{hasText:'First principles of instruction'}).click();
  ok(/Merrill, M\. D\. \(2002\)\. First principles of instruction\. Educational Technology Research and Development, 50\(3\), 43–59\. https:\/\/doi\.org\/10\.1007\/BF02505024/.test(await desk.locator('.dd-ref').first().innerText()),'an article has its DOI in the citation');
  await page.waitForFunction(()=>/Prism Reference Manager: /.test(document.querySelector('.dd-prm-status')?.textContent||'')&&!/not checked/.test(document.querySelector('.dd-prm-status')?.textContent||''),null,{timeout:15000});
  ok(/not connected|not allowed|Connect/.test(await desk.locator('.dd-prm-status').innerText()),'the Reference Manager is said to be not connected, and everything else here still works: '+await desk.locator('.dd-prm-status').innerText());
  await desk.getByRole('button',{name:/^Put the reference of First principles/}).click();
  await page.waitForFunction(()=>/The reference is in the guide's list of references/.test(document.querySelector('#status')?.textContent||''),null,{timeout:10000});
  ok(true,'the reference goes into the guide\'s list of references');
  await shot('2-models.png');

  // ---- course map
  await desk.getByRole('tab',{name:'Course map'}).click();
  await desk.getByRole('button',{name:'Add an assessment'}).click();
  await desk.getByLabel('Assessment 1: title').fill('Short quiz');
  await desk.getByLabel('Item slots for each objective').fill('2');
  await desk.getByRole('group',{name:'Assessment 1 checks these objectives'}).getByLabel(/Objective 1/).check();
  await desk.getByRole('group',{name:'Assessment 1 checks these objectives'}).getByLabel(/Objective 2/).check();
  const planTable=desk.locator('section[aria-label="The plan\'s map"] table');
  ok((await planTable.locator('tbody tr').count())===2,'the plan\'s map has a row for each objective');
  ok(/2 objectives: 1 with an activity, 2 with an assessment, 1 with both\./.test(await desk.locator('.dd-summary').first().innerText()),'and says how many have what');
  ok(/Objective 2: No activity gives practice at this\./.test(await desk.locator('section[aria-label="The plan\'s map"] .dd-issues').innerText()),'what is missing is said in words');
  const csv=await download(()=>desk.getByRole('button',{name:'Save the plan\'s course map as a CSV file'}).click());
  const csvText=fs.readFileSync(csv.path,'utf8');
  ok(/course map \(plan\)\.csv$/.test(csv.name)&&csvText.startsWith('﻿Module,Lesson,Objective,Level,Activities,Assessments'),'the map is a CSV a spreadsheet opens as UTF-8: '+csv.name);
  const slots=await download(()=>desk.getByRole('button',{name:'Save item slots (CSV)'}).click());
  ok(fs.readFileSync(slots.path,'utf8').split('\r\n').filter(Boolean).length===5&&/\[Slot\] Write a question that checks/.test(fs.readFileSync(slots.path,'utf8')),'item slots for OpenAssess: 4 rows (2 objectives, 2 slots each) and a heading');
  await desk.getByRole('button',{name:'Put the plan\'s course map in the guide'}).click();
  await page.waitForFunction(()=>/Course map is in the guide/.test(document.querySelector('#status')?.textContent||''),null,{timeout:10000});
  ok(true,'the map can be put in the guide as a table');
  await shot('3-map.png');

  // ---- storyboard
  await desk.getByRole('tab',{name:'Storyboard'}).click();
  await desk.getByRole('button',{name:'Add a module'}).click();
  await desk.getByLabel('Module 1: title').fill('Talk first');
  await desk.getByRole('button',{name:'Add a lesson'}).click();
  await desk.getByLabel('Lesson 1.1: title').fill('The coaching conversation');
  await desk.getByLabel('Minutes').fill('30');
  await desk.getByRole('group',{name:'Lesson 1.1 teaches'}).getByLabel(/Objective 1/).check();
  await desk.getByRole('group',{name:'Lesson 1.1 has these activities'}).getByLabel(/Activity 1/).check();
  await desk.getByRole('group',{name:'Lesson 1.1 is checked by'}).getByLabel(/Assessment 1/).check();
  ok(/1 module, 1 lesson, 30 minutes in all/.test(await desk.locator('.dd-summary').innerText()),'the minutes are added up');
  const author=await download(()=>desk.getByRole('button',{name:'Save the Author skeleton'}).click());
  const az=await JSZip.loadAsync(fs.readFileSync(author.path));
  ok(/OpenStage Author project\.zip$/.test(author.name)&&Object.keys(az.files).some(f=>/\.osproj\/project\.json$/.test(f)),'the Author skeleton is a zip with a project: '+author.name);
  const scroll=await download(()=>desk.getByRole('button',{name:'Save the Scroll skeleton'}).click());
  const sz=await JSZip.loadAsync(fs.readFileSync(scroll.path));
  ok(/\.openscroll$/.test(scroll.name)&&Object.keys(sz.files).join()==='course.json'&&JSON.parse(await sz.file('course.json').async('string')).course.sections[0].lessons[0].title==='The coaching conversation','the Scroll skeleton is an .openscroll with a course.json: '+scroll.name);
  await saved();

  // ---- evaluation
  await desk.getByRole('tab',{name:'Evaluation plan'}).click();
  ok(/There are no measures yet/.test(await desk.locator('.dd-issues').innerText()),'an empty plan says what to add');
  await desk.getByRole('button',{name:'Add a measure to Level 1: Reaction'}).click();
  await desk.getByLabel('Level 1: Reaction, measure 1: name').fill('End of day');
  await desk.getByLabel('Statements, one to a line').fill('The session met its goals.\nI can use what I learned.');
  await desk.getByLabel('When').first().fill('At the end of the second day');
  await desk.getByRole('button',{name:'Add a measure to Level 2: Learning'}).click();
  await desk.getByLabel('Level 2: Learning, measure 1: name').fill('Escalation quiz');
  await desk.getByLabel('When').nth(1).fill('First and last hour');
  ok(/Nothing is planned for level 3, Behavior\./.test(await desk.locator('.dd-issues').innerText()),'a level left out is said');
  await desk.getByLabel(/^People to enter answers for/).fill('12');
  const pstat=await download(()=>desk.getByRole('button',{name:'Save for PrismStatistics'}).click());
  const bytes=new Uint8Array(fs.readFileSync(pstat.path));
  ok(/evaluation plan\.pstat$/.test(pstat.name)&&checkPstat(bytes).length===0,'the evaluation plan is a .pstat the core\'s checker accepts: '+JSON.stringify(checkPstat(bytes).slice(0,2)));
  const project=readPstat(bytes);
  ok(project.dataset.variables.map(v=>v.name).join()==='respondent,l1_end_of_day_q01,l1_end_of_day_q02,l2_escalation_quiz_pre,l2_escalation_quiz_post'&&project.dataset.columns[0].values.length===12,'a variable for each measure, and twelve rows');
  await shot('5-evaluation.png');

  // ---- a new guide from the storyboard
  await desk.getByRole('tab',{name:'Storyboard'}).click();
  await desk.getByRole('button',{name:'Make a new guide from the storyboard'}).click();
  await page.locator('dialog[open] #dlg-name').waitFor({timeout:10000});
  await page.locator('dialog[open] #dlg-name').fill('Coaching plan');
  await page.locator('dialog[open]').getByRole('button',{name:'Make the guide'}).click();
  // (the guide is drawn row by row, so the check waits for all of what it looks for, not only the first heading)
  await page.waitForFunction(()=>{const t=document.querySelector('#work-guide')?.innerText||'';return /Module 1: Talk first/.test(t)&&/Lesson 1\.1: The coaching conversation/.test(t)&&/A real problem or a short story/.test(t)&&/Short quiz/.test(t);},null,{timeout:20000}).catch(()=>{});
  const text=await page.locator('#work-guide').innerText();
  ok(await page.locator('#work-guide h1, #work-guide [data-index]').filter({hasText:/^Coaching plan$/}).count()>=1,'the name chosen is the title of the new guide');
  ok(/Lesson 1\.1: The coaching conversation/.test(text)&&/Learners will be able to explain the three steps/.test(text)&&/A real problem or a short story/.test(text)&&/Short quiz/.test(text),'the new guide has the lesson, its objective, its activity and its assessment');
  ok(fs.readdirSync(GUIDES).includes('Design check.guide-project.json')&&fs.readdirSync(GUIDES).includes('Coaching plan.guide-project.json'),'it is a guide beside the first, which is still there');
  await saved();
  const plan=parseProject(fs.readFileSync(path.join(GUIDES,'Coaching plan.guide-project.json'),'utf8')).guide;
  ok(readDesign(plan.design).objectives.length===2,'and it carries the notes of the desk, so the design can go on');
  await shot('6-plan.png');

  ok(problems.length===0,'nothing was logged as an error: '+problems.slice(0,3).join(' | '));
}catch(e){
  console.log('ERROR',e.stack||e);
  fail++;
}finally{
  await browser?.close();
  server.kill();
}
console.log(`${pass} passed, ${fail} failed (screenshots in ${SHOTS})`);
process.exit(fail?1:0);
