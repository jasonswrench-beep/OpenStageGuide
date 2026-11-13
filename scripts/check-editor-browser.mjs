// Real-browser check of the editor (program/ui/editor.js): the guide as rows, changing a paragraph's words, Enter and Backspace, headings and cue
// lines, what cannot be changed, Undo and Redo shared with the tools, and that a change is on disk. It needs Playwright, which this repository
// does not depend on: set PLAYWRIGHT_DIR to the playwright package folder, then run  node scripts/check-editor-browser.mjs  from the repository
// root. It starts its own server on port 3854 with a scratch guides folder, and stops it again. Screenshots go to a temporary folder named in
// the last line.
import {createRequire} from 'node:module';
import {spawn} from 'node:child_process';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import {fileURLToPath} from 'node:url';
import {parseProject} from '../src/guide-project.mjs';
const require=createRequire(import.meta.url);
if(!process.env.PLAYWRIGHT_DIR){console.error('Set PLAYWRIGHT_DIR to the playwright package folder.');process.exit(2);}
const {chromium}=require(process.env.PLAYWRIGHT_DIR);
const APP=path.resolve(path.dirname(fileURLToPath(import.meta.url)),'..');
const RUN=fs.mkdtempSync(path.join(os.tmpdir(),'osg-editor-check-'));
const SHOTS=path.join(RUN,'shots'),GUIDES=path.join(RUN,'guides');
fs.mkdirSync(SHOTS);
const PORT=3854,BASE=`http://127.0.0.1:${PORT}`;
let pass=0,fail=0;
const ok=(cond,what)=>{if(cond){pass++;console.log('PASS',what);}else{fail++;console.log('FAIL',what);}};

const server=spawn(process.execPath,['scripts/project-server.mjs','--folder',GUIDES,'--port',String(PORT)],{cwd:APP,stdio:['ignore','pipe','pipe']});
await new Promise((res,rej)=>{let t='';server.stdout.on('data',d=>{t+=d;if(/OpenStage Guide at/.test(t))res();});server.stderr.on('data',d=>rej(new Error(String(d))));setTimeout(()=>rej(new Error('server did not start')),15000);});

// What is on disk: the guide's paragraphs, once the page says it is saved.
const onDisk=()=>{
  const f=fs.readdirSync(GUIDES).find(n=>n.endsWith('.guide-project.json'));
  return parseProject(fs.readFileSync(path.join(GUIDES,f),'utf8')).guide.paragraphs;
};

let browser;
try{
  browser=await chromium.launch();
  const page=await browser.newPage({viewport:{width:1280,height:900}});
  const problems=[];
  page.on('pageerror',e=>problems.push(String(e)));
  page.on('console',m=>{if(m.type()==='error'||/Refused|violat/i.test(m.text()))problems.push(m.text());});
  const stores=[];
  page.on('request',r=>{if(r.method()==='PUT'&&r.url().endsWith('/project'))stores.push(new URL(r.frame().url()).pathname);});
  const shot=name=>page.screenshot({path:path.join(SHOTS,name)});
  // (the page saves 0.4 seconds after the last change, so what is on disk is read only after that has passed and it says Saved)
  const saved=async()=>{await page.waitForTimeout(600);await page.waitForFunction(()=>document.querySelector('.bar-context-state')?.getAttribute('data-state')==='saved',null,{timeout:10000});};
  const rows=()=>page.locator('#work-guide [data-index]');
  const row=re=>page.locator('#work-guide [data-index]').filter({hasText:re}).first();
  const active=()=>page.evaluate(()=>{const a=document.activeElement;return a?{tag:a.tagName,index:a.dataset?.index??null,label:a.getAttribute('aria-label'),value:a.value??null,start:a.selectionStart??null}:null;});

  // ---- a guide with something in it, made with the tools
  await page.goto(`${BASE}/`);
  await page.click('#welcome-new');
  await page.fill('#dlg-name','Editor guide');
  await page.press('#dlg-name','Enter');
  const fr=page.frameLocator('#tools-frame');
  await fr.locator('#guide-start').waitFor({state:'visible',timeout:15000});
  await fr.locator('#guide-title').fill('Editor guide');
  await fr.locator('#guide-title').press('Enter');
  await fr.locator('#guide-build').waitFor({state:'visible',timeout:15000});
  await page.getByRole('menuitem',{name:/^Write/}).click();
  await page.getByRole('menuitem',{name:/Add module/}).click();
  await fr.locator('#guide-new-title').fill('Naming it');
  await fr.locator('#guide-add-module').click();
  await page.waitForFunction(()=>/Module 1: Naming it/.test(document.querySelector('#work-guide')?.innerText||''),null,{timeout:10000});
  await fr.locator('#guide-new-title').fill('First');
  await fr.locator('#guide-new-minutes').fill('10');
  await fr.locator('#guide-add-lesson').click();
  await page.getByRole('menuitem',{name:/^Write/}).click();
  await page.getByRole('menuitem',{name:/Instructional cues/i}).click();
  await fr.locator('#guide-cue-ask').evaluate(e=>{e.checked=true;e.dispatchEvent(new Event('change',{bubbles:true}));});
  await fr.locator('#guide-block-text').fill('What makes it hard?');
  await fr.locator('#guide-block-notes').fill('Listen for blame.');
  await fr.locator('#guide-block-form').evaluate(f=>f.requestSubmit());
  await page.waitForFunction(()=>/What makes it hard\?/.test(document.querySelector('#work-guide')?.innerText||''),null,{timeout:10000});
  await saved();
  await shot('1-guide.png');

  // ---- the guide is rows
  const n=await rows().count();
  ok(n>=15,'the guide is drawn as rows, one for each paragraph: '+n);
  ok(await row(/^Naming it|Module 1: Naming it/).getAttribute('data-editable')==='true','a module heading can be changed');
  ok(await row(/Time: 10 minutes/).getAttribute('data-editable')==='false','the time line, which the program writes, cannot');
  ok(/program writes/.test(await row(/Time: 10 minutes/).getAttribute('aria-description')||''),'and says why, for a screen reader');
  ok(await page.locator('#work-guide [data-index][tabindex="0"]').count()===1,'one row is the Tab stop');

  // ---- moving between rows
  await row(/^Editor guide$/).focus();
  const first=await active();
  await page.keyboard.press('ArrowDown');
  const second=await active();
  ok(first.index!==null&&second.index!==null&&Number(second.index)>Number(first.index),'ArrowDown moves to the next row');
  await page.keyboard.press('ArrowUp');
  ok((await active()).index===first.index,'ArrowUp moves back');
  await page.keyboard.press('End');
  const last=await active();
  ok(Number(last.index)>Number(first.index)+5,'End goes to the last row');
  await page.keyboard.press('Home');
  ok((await active()).index===first.index,'Home goes to the first');

  // ---- changing the words of a plain paragraph
  await row(/^Facilitator guide$/).click();
  let a=await active();
  ok(a.tag==='TEXTAREA'&&a.value==='Facilitator guide'&&a.label==='Paragraph words','a click starts changing the words: a text box with the words and a name');
  await page.keyboard.press('Control+a');
  await page.keyboard.type('Facilitator guide for reframing conflict');
  a=await active();
  ok(a.tag==='TEXTAREA'&&a.value==='Facilitator guide for reframing conflict','typing keeps the keyboard in the box');
  await saved();
  ok(onDisk().some(p=>p.text==='Facilitator guide for reframing conflict'),'the words are in the project file on disk');
  const long='Line '.repeat(60).trim();
  await page.keyboard.press('Control+a');
  await page.keyboard.type(long,{delay:0});
  a=await active();
  ok(a.tag==='TEXTAREA'&&a.value===long,'a long, fast run of typing arrives whole and the keyboard stays in the box');
  await saved();
  ok(onDisk().some(p=>p.text===long),'and it is all on disk');
  await page.keyboard.press('Control+a');
  await page.keyboard.type('Facilitator guide for reframing conflict');
  await page.keyboard.press('Escape');
  a=await active();
  ok(a.tag!=='TEXTAREA'&&a.index!==null,'Escape stops changing the words and the row has the keyboard');
  ok((await row(/^Facilitator guide for reframing/).innerText()).trim()==='Facilitator guide for reframing conflict','the row shows the words as they read');
  await shot('2-words.png');

  // ---- Enter makes the next paragraph; Backspace at the start joins it back
  const placeholder=/^Thank the people and organizations/;
  await row(placeholder).click();
  await page.keyboard.press('Home');
  for(let i=0;i<5;i++)await page.keyboard.press('ArrowRight');
  await page.keyboard.press('Enter');
  a=await active();
  ok(a.tag==='TEXTAREA'&&a.value.startsWith(' the people and organizations')&&a.start===0,'Enter in the middle: the rest is a new paragraph, with the caret at its start');
  await saved();
  let disk=onDisk().map(p=>p.text);
  ok(disk.includes('Thank')&&disk.some(t=>t.startsWith(' the people and organizations')),'both paragraphs are on disk');
  await page.keyboard.press('Backspace');
  a=await active();
  ok(a.tag==='TEXTAREA'&&a.value.startsWith('Thank the people')&&a.start===5,'Backspace at the start joins it back, the caret where they met');
  await saved();
  ok(!onDisk().map(p=>p.text).includes('Thank'),'and the project file has one paragraph again');
  await page.keyboard.press('End');
  await page.keyboard.press('Enter');
  a=await active();
  ok(a.tag==='TEXTAREA'&&a.value===''&&a.start===0,'Enter at the end makes an empty paragraph and goes into it');
  ok(await page.locator('#work-guide textarea').count()===1,'only one paragraph is being changed: the row Enter left is drawn as it reads');
  await page.keyboard.type('A new line of the author\'s own.');
  await page.keyboard.press('Escape');
  ok(await page.locator('#work-guide textarea').count()===0,'no text box is left when changing the words has stopped');
  ok(/A new line of the author's own\./.test(await page.locator('#work-guide').innerText()),'the new paragraph is in the guide');
  await shot('3-enter.png');

  // ---- Undo and Redo, one history with the tools
  await page.keyboard.press('Control+z');
  ok(!/A new line of the author's own\./.test(await page.locator('#work-guide').innerText()),'Ctrl+Z undoes the paragraph typed');
  await page.keyboard.press('Control+y');
  ok(/A new line of the author's own\./.test(await page.locator('#work-guide').innerText()),'Ctrl+Y brings it back');
  await page.getByRole('menuitem',{name:/^Write/}).click();
  await page.getByRole('menuitem',{name:/Instructional cues/i}).click();
  await fr.locator('#guide-cue-say').evaluate(e=>{e.checked=true;e.dispatchEvent(new Event('change',{bubbles:true}));});
  await fr.locator('#guide-block-text').fill('From the tools.');
  await fr.locator('#guide-block-form').evaluate(f=>f.requestSubmit());
  await page.waitForFunction(()=>/From the tools\./.test(document.querySelector('#work-guide')?.innerText||''),null,{timeout:10000});
  await row(/^Editor guide$/).focus();
  await page.keyboard.press('Control+z');
  ok(!/From the tools\./.test(await page.locator('#work-guide').innerText()),'Ctrl+Z undoes what the tools added: one history');
  ok(/A new line of the author's own\./.test(await page.locator('#work-guide').innerText()),'and the author\'s own paragraph is still there');

  // ---- a heading: the number is shown, the title is typed
  await row(/Module 1: Naming it/).click();
  a=await active();
  ok(a.tag==='TEXTAREA'&&a.value==='Naming it'&&a.label==='Module title','a module heading: the title is in the box');
  ok(/Module 1:/.test(await page.locator('#work-guide .g-editing .g-lead').first().innerText()),'and the number the program writes is beside it');
  await page.keyboard.press('Control+a');
  await page.keyboard.type('Naming feelings');
  await page.keyboard.press('Escape');
  ok((await row(/Module 1: Naming feelings/).count())===1,'the row reads "Module 1: Naming feelings"');
  await saved();
  ok(onDisk().some(p=>p.text==='Module 1: Naming feelings'),'and so does the project file');
  await row(/Module 1: Naming feelings/).click();
  await page.keyboard.press('Enter');
  a=await active();
  ok(a.tag==='TEXTAREA'&&a.label==='Lesson title','Enter in a heading goes on to the next row, not a new paragraph');
  await page.keyboard.press('Escape');

  // ---- a cue line keeps its picture and its name
  await row(/What makes it hard\?/).first().click();
  a=await active();
  ok(a.tag==='TEXTAREA'&&a.label==='Block words'&&a.value.startsWith('What makes it hard?'),'a cue line: the words are in the box, the cue\'s name beside it');
  await page.keyboard.press('End');
  await page.keyboard.type(' Really?');
  await page.keyboard.press('Escape');
  await saved();
  const cue=onDisk().find(p=>p.text.includes('Really?'));
  ok(!!cue&&cue.xml.includes('<w:drawing>')&&cue.xml.includes('OSGCueLabel')&&cue.xml.includes('<w:tab/>'),'on disk the cue line still has its picture, its name and its tab');
  await shot('4-cue.png');

  // ---- whole paragraphs: choose, move, delete, add, kind, copy, cut, paste
  const order=()=>onDisk().map(p=>p.text);
  const selectedCount=()=>page.locator('#work-guide [data-selected="true"]').count();
  const status=()=>page.locator('#status').innerText();
  await saved();
  const original=order();
  await row(/^About This Guide$/).focus();
  await page.keyboard.press('Shift+ArrowDown');
  await page.keyboard.press('Shift+ArrowDown');
  ok(await selectedCount()===3,'Shift+Down chooses rows together: three are chosen');
  ok(/3 paragraphs selected/.test(await status()),'and says so');
  await shot('5-selected.png');
  await page.keyboard.press('Alt+ArrowDown');
  await saved();
  let now=order();
  ok(now.indexOf('Say where this program fits: who it is for, what need it meets and how it connects to other training.')<now.indexOf('About This Guide'),'Alt+Down moves the chosen rows below the next one, and it is on disk');
  ok(await selectedCount()===3,'the rows stay chosen after the move');
  await page.keyboard.press('Control+z');
  await saved();
  ok(JSON.stringify(order())===JSON.stringify(original),'Ctrl+Z puts them back exactly as they were');
  await page.keyboard.press('Escape');
  ok(await selectedCount()===0,'Escape clears the choice');

  const prep=/^List what the facilitator needs before the session/;
  await row(prep).focus();
  await page.keyboard.press('Delete');
  await saved();
  ok(!order().some(t=>/^List what the facilitator needs/.test(t)),'Delete takes the paragraph away, on disk too');
  ok(/deleted/.test(await status()),'and says that Ctrl+Z brings it back');
  await page.keyboard.press('Control+z');
  await saved();
  ok(order().some(t=>/^List what the facilitator needs/.test(t)),'Ctrl+Z brings it back');

  await row(/^Say what this guide is for/).focus();
  await page.keyboard.press('Control+Enter');
  a=await active();
  ok(a.tag==='TEXTAREA'&&a.value==='','Ctrl+Enter adds a paragraph below and goes into it');
  await page.keyboard.type('Added by the author.');
  await page.keyboard.press('Escape');
  await saved();
  now=order();
  ok(now.indexOf('Added by the author.')===now.indexOf('Say what this guide is for and who should use it.')+1,'it is right after the paragraph it was added below, on disk');

  await page.getByRole('menuitem',{name:/^Edit/}).click();
  await page.getByRole('menuitem',{name:/Kind of paragraph/}).click();
  await page.locator('dialog[open] input[value="notes"]').waitFor();
  ok(await page.locator('dialog[open] input[type=radio]').count()===5,'Edit, Kind of paragraph offers the five kinds');
  await page.locator('dialog[open] input[value="notes"]').check();
  await page.getByRole('button',{name:'Change'}).click();
  await saved();
  ok(onDisk().find(p=>p.text==='Added by the author.').style==='OSG Facilitator Notes','the paragraph is now a facilitator note, on disk');
  ok(await row(/^Added by the author\.$/).evaluate(e=>e.classList.contains('g-notes')),'and drawn as one');

  await row(/^Added by the author\.$/).focus();
  await page.keyboard.press('Control+c');
  ok(/Copied/.test(await status()),'Ctrl+C copies');
  await row(/^Say where this program fits/).focus();
  await page.keyboard.press('Control+v');
  await saved();
  const added=onDisk().filter(p=>p.text==='Added by the author.');
  ok(added.length===2&&added.every(p=>p.style==='OSG Facilitator Notes'),'Ctrl+V pastes a second one, still a facilitator note');
  await page.keyboard.press('Control+z');
  await saved();
  ok(onDisk().filter(p=>p.text==='Added by the author.').length===1,'and Ctrl+Z takes it away again');
  await row(/^Added by the author\.$/).focus();
  await page.keyboard.press('Control+x');
  await saved();
  ok(onDisk().filter(p=>p.text==='Added by the author.').length===0,'Ctrl+X takes it out');
  await row(/^Say where this program fits/).focus();
  await page.keyboard.press('Control+v');
  await saved();
  ok(onDisk().filter(p=>p.text==='Added by the author.').length===1,'and Ctrl+V puts it back somewhere else');

  await row(/^Say what this guide is for/).click();
  await page.keyboard.press('End');
  await page.evaluate(()=>{const ta=document.querySelector('#work-guide textarea');const dt=new DataTransfer();dt.setData('text/plain','First pasted.\nSecond pasted.\n\nThird pasted.');ta.dispatchEvent(new ClipboardEvent('paste',{clipboardData:dt,bubbles:true,cancelable:true}));});
  await saved();
  now=order();
  const at=now.indexOf('Say what this guide is for and who should use it.First pasted.');
  ok(at>=0&&now[at+1]==='Second pasted.'&&now[at+2]==='Third pasted.','several lines pasted into the words become a paragraph each (the blank line is dropped)');
  a=await active();
  ok(a.tag==='TEXTAREA'&&a.value==='Third pasted.','the keyboard is at the end of what was pasted');
  await page.keyboard.press('Escape');
  await page.getByRole('menuitem',{name:/^Edit/}).click();
  await page.getByRole('menuitem',{name:/^Undo/}).click();
  await saved();
  ok(!order().includes('Third pasted.'),'Edit, Undo undoes the paste');
  await shot('6-edited.png');

  // ---- the Insert menu: elements, a table, and a form from a template
  const insertMenu=async label=>{await page.getByRole('menuitem',{name:/^Insert/}).click();await page.getByRole('menuitem',{name:label}).click();};
  const styles=()=>onDisk().map(p=>p.style);
  await row(/^Say where this program fits/).focus();
  const linesBefore=styles().filter(x=>x==='OSG Form Line').length;
  await insertMenu(/Lines to write on/);
  await page.locator('dialog[open] #dlg-f-count').fill('99');
  await page.getByRole('button',{name:'Insert'}).click();
  ok(/1 to 12/.test(await page.locator('dialog[open] .problem').innerText()),'a number that is too big is explained in the dialog, which stays open');
  await page.locator('dialog[open] #dlg-f-count').fill('5');
  await page.locator('dialog[open] #dlg-f-label').fill('Your answer');
  await page.getByRole('button',{name:'Insert'}).click();
  await saved();
  ok(styles().filter(x=>x==='OSG Form Line').length===linesBefore+5,'five lines to write on are in the guide, on disk');
  now=order();
  ok(now.indexOf('Your answer')===now.indexOf('Say where this program fits: who it is for, what need it meets and how it connects to other training.')+1,'with their label, right after the paragraph the author was on');
  ok(await selectedCount()>=6,'what was inserted is chosen, so it can be moved or deleted');
  ok(/inserted/.test(await status()),'and the program says Ctrl+Z takes it back');
  await page.keyboard.press('Control+z');
  await saved();
  ok(styles().filter(x=>x==='OSG Form Line').length===linesBefore,'one Ctrl+Z takes the whole insert back');

  await row(/^Say where this program fits/).focus();
  await insertMenu(/Table to fill in/);
  await page.locator('dialog[open] #dlg-f-columns').fill('Topic, Notes, Action');
  await page.locator('dialog[open] #dlg-f-rows').fill('3');
  await page.locator('dialog[open] #dlg-f-caption').fill('Session notes');
  await page.getByRole('button',{name:'Insert'}).click();
  await saved();
  const tables=onDisk().filter(p=>/^<w:tbl[\s>]/.test(p.xml));
  ok(tables.length===1&&/Topic/.test(tables[0].xml)&&/tblHeader/.test(tables[0].xml),'a table with a header row and the three columns is in the guide');
  ok(await page.locator('#work-guide table.g-table').count()===1,'and is drawn as one row of the guide');
  ok((await page.locator('#work-guide table.g-table th').allInnerTexts()).join('|')==='Topic|Notes|Action','with its column names as header cells');
  await shot('7-table.png');

  await row(/^Say where this program fits/).focus();
  await insertMenu(/^Form\.\.\./);
  await page.locator('dialog[open] #dlg-f-template').selectOption('checklist');
  await page.locator('dialog[open] #dlg-f-title').fill('Preparation list');
  await page.getByRole('button',{name:'Insert'}).click();
  await saved();
  ok(onDisk().some(p=>p.style==='OSG Form Title'&&p.text==='Checklist: Preparation list'),'a form from the Checklist template is in the guide, with its title');
  ok(onDisk().some(p=>p.style==='OSG Form Print'),'with its print line');
  await page.getByRole('button',{name:'Update guide'}).click();
  await page.waitForTimeout(1200);
  await saved();
  ok(onDisk().some(p=>/^Checklist [A-Z0-9.]+: Preparation list$/.test(p.text)),'Update guide numbers it');
  await shot('8-form.png');

  // ---- what cannot be changed is read, not rewritten
  await row(/Time: 10 minutes/).focus();
  await page.keyboard.press('Enter');
  a=await active();
  ok(a.tag!=='TEXTAREA','Enter on the time line does not open it');
  ok(/program writes/.test(await page.locator('#status').innerText()),'the program says why');
  const timeBefore=onDisk().find(p=>/Time: 10 minutes/.test(p.text)).xml;

  // ---- who saved, and what went wrong
  await saved();
  ok(onDisk().find(p=>/Time: 10 minutes/.test(p.text)).xml===timeBefore,'the time line on disk is exactly as it was');
  ok(stores.length>0&&stores.every(f=>f==='/program/'),'the guide is saved by the program page only: '+JSON.stringify([...new Set(stores)]));
  ok(problems.length===0,'no page errors or policy violations in the whole run: '+JSON.stringify(problems.slice(0,3)));
}finally{
  if(browser)await browser.close();
  server.kill();
}
console.log(`\n${pass} passed, ${fail} failed (screenshots in ${SHOTS})`);
process.exitCode=fail?1:0;
