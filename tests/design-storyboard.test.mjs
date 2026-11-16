import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import {spawnSync} from 'node:child_process';
import {createRequire} from 'node:module';
import JSZip from 'jszip';
import {DOMParser,XMLSerializer} from '@xmldom/xmldom';
import {createDemoAdapter} from '../src/guide-word.mjs';
import {readGuide,readBlock,guideIssues,updatePlan,blockSpan} from '../src/guide-model.mjs';
import {authorProject} from '../src/guide-author.mjs';
import {zipWorkbook} from '../src/guide-workbook.mjs';
import {planBody,authorScreens,scrollCourse,scrollBytes,scrollFileName,scrollId,itemSlots,itemSlotsCsv,SLOT_COLUMNS,needStoryboard} from '../src/design-storyboard.mjs';
import {design} from './helpers/design-fixture.mjs';

const demo=()=>createDemoAdapter({DOMParserClass:DOMParser,SerializerClass:XMLSerializer});

test('the guide\'s course plan has each module and lesson in order, with times, an Objective block for each objective, the activities and assessments, and a Transition at the end of each module', async () => {
  const a=demo();
  await a.insert(ctx=>planBody(design(),{title:'Coaching first'},ctx),{where:'start',spare:false});
  const paras=await a.paragraphs();
  const heads=paras.filter(p=>/^OSG (Module|Lesson)$/.test(p.style)).map(p=>p.text);
  assert.deepEqual(heads,['Module 1: Talk first','Lesson 1.1: The coaching conversation','Lesson 1.2: The calm opening','Module 2: Know the policy','Lesson 2.1: When to escalate']);
  const times=paras.filter(p=>p.style==='OSG Time').map(p=>p.text.replace(/^\W+/,''));
  assert.deepEqual(times,['Time: 30 minutes. Cumulative time: 0 minutes.','Time: 45 minutes. Cumulative time: 30 minutes.','Time: 20 minutes. Cumulative time: 75 minutes.']);
  const g=readGuide(paras);
  assert.equal(g.modules.length,2);
  assert.equal(g.modules[0].lessons.length,2);
  assert.ok(g.modules.every(m=>m.hasTransition),'every module ends with its Transition block');
  const cues=paras.filter(p=>/^OSG Cue/.test(p.style)).map(p=>p.text.replace(/^\W+/,'').replace(/\s+/g,' '));
  assert.ok(cues.some(t=>/^Objective Learners will be able to explain the three steps/.test(t)));
  assert.ok(cues.some(t=>/^Activity Story of a hard conversation \(10 min\)/.test(t)));
  assert.ok(cues.some(t=>/^Assessment Short quiz/.test(t)));
  const noted=paras.filter(p=>/Facilitator notes: Planned with Nine events of instruction, event: Gain attention\. Use a real case\./.test(p.text));
  assert.equal(noted.length,1,'the model and its event are in the facilitator\'s notes');
  const blockAt=paras.findIndex(p=>/Activity\s+Story of a hard/.test(p.text));
  assert.match(readBlock(paras,blockAt).notes,/For: Learners will be able to explain/);
});

test('the course plan is a guide the Review has no failure for', async () => {
  const a=demo();
  await a.insert(ctx=>planBody(design(),{title:'Coaching first'},ctx),{where:'start',spare:false});
  const issues=guideIssues(updatePlan(await a.paragraphs()));
  assert.deepEqual(issues.filter(i=>i.level==='fail'),[]);
  assert.ok(!issues.some(i=>i.rule==='MODULE-TRANSITION'||i.rule==='LESSON-TIME'));
});

test('a storyboard with no lesson cannot become anything, and says what to do', () => {
  const d=design();
  d.storyboard=[];
  assert.throws(()=>needStoryboard(d),/The storyboard has no lessons yet/);
  assert.throws(()=>authorScreens(d),/no lessons/);
  assert.throws(()=>scrollCourse(d,{title:'x'}),/no lessons/);
  d.storyboard=[{id:'m1',title:'Empty',lessons:[]}];
  assert.throws(()=>authorScreens(d),/no lessons/);
});

test('Author gets a screen for each lesson, activity and assessment, with the levels in the notes, and writes a project', () => {
  const screens=authorScreens(design());
  assert.deepEqual(screens.map(s=>`${s.lesson} | ${s.number} | ${s.title}`),[
    'Lesson 1.1: The coaching conversation | 1 | The coaching conversation','Lesson 1.1: The coaching conversation | 2 | Story of a hard conversation','Lesson 1.1: The coaching conversation | 3 | Short quiz',
    'Lesson 1.2: The calm opening | 1 | The calm opening','Lesson 1.2: The calm opening | 2 | Role play: the opening','Lesson 1.2: The calm opening | 3 | Skills check',
    'Lesson 2.1: When to escalate | 1 | When to escalate']);
  assert.equal(screens[0].onScreen,'Learners will be able to explain the three steps of a coaching conversation.');
  assert.ok(screens[0].blocks.some(b=>b.cue==='developer-note'&&/Objective \(Understand\)/.test(b.text)));
  assert.ok(screens[1].blocks.some(b=>/Planned with Nine events of instruction, event: Gain attention/.test(b.text)));
  assert.ok(screens[2].blocks.some(b=>b.cue==='knowledge-check'&&/2 questions for each objective/.test(b.text)));
  const files=authorProject(screens,{title:'Coaching first',today:'2026-11-16'});
  const project=JSON.parse(files.find(f=>/project\.json$/.test(f.name)).text);
  assert.deepEqual(project.scenes.map(s=>[s.name,s.slides.length]),[['Lesson 1.1: The coaching conversation',3],['Lesson 1.2: The calm opening',3],['Lesson 2.1: When to escalate',1]]);
  const notes=files.find(f=>/hand-off notes/.test(f.name)).text;
  assert.match(notes,/Objective \(Apply\): Learners will be able to demonstrate a calm opening/);
});

const VALIDATOR=process.env.OPENSTAGE_AUTHOR_VALIDATOR||['C:/Dev/openstage-author/tools/validate.py'].find(f=>fs.existsSync(f))||'';
test('the Author skeleton passes OpenStage Author\'s own validator with no errors',{skip:!VALIDATOR&&'set OPENSTAGE_AUTHOR_VALIDATOR to the OpenStage Author validate.py to run this'},async () => {
  const dir=fs.mkdtempSync(path.join(os.tmpdir(),'osg-skeleton-'));
  try{
    const bytes=await zipWorkbook(authorProject(authorScreens(design()),{title:'Coaching first'}),JSZip,'uint8array');
    const zip=await JSZip.loadAsync(bytes);
    for(const [name,entry] of Object.entries(zip.files))if(!entry.dir){fs.mkdirSync(path.dirname(path.join(dir,name)),{recursive:true});fs.writeFileSync(path.join(dir,name),await entry.async('nodebuffer'));}
    const run=spawnSync('python',['-B',VALIDATOR,path.join(dir,'Coaching first.osproj')],{encoding:'utf8',cwd:path.dirname(path.dirname(VALIDATOR))});
    assert.equal(run.status,0,`${run.stdout}\n${run.stderr}`);
    assert.match(run.stdout,/0 errors/);
  }finally{fs.rmSync(dir,{recursive:true,force:true});}
});

test('Scroll gets a section for each module and a lesson for each lesson, with ids in its own form, and objectives, activities and checks as blocks', () => {
  const course=scrollCourse(design(),{title:'Coaching first',author:'A. Writer',now:'2026-11-16T12:00:00Z'});
  assert.equal(course.format,'openscroll');
  assert.equal(course.schemaVersion,3);
  assert.deepEqual(course.assets,[]);
  assert.deepEqual(course.course.sections.map(s=>[s.title,s.lessons.map(l=>l.title)]),[['Talk first',['The coaching conversation','The calm opening']],['Know the policy',['When to escalate']]]);
  const ids=[course.course.id,...course.course.sections.flatMap(s=>[s.id,...s.lessons.flatMap(l=>[l.id,...l.blocks.map(b=>b.id)])])];
  assert.equal(new Set(ids).size,ids.length,'every id is unique through the whole course');
  for(const id of ids)assert.match(id,/^[a-z][a-z0-9-]{0,63}$/,id);
  const first=course.course.sections[0].lessons[0];
  assert.deepEqual(first.blocks.map(b=>[b.type,b.variant]),[['heading','heading'],['statement','note'],['list','bulleted'],['list','bulleted'],['list','checklist']]);
  assert.deepEqual(first.blocks[2].content.items,['Learners will be able to explain the three steps of a coaching conversation.']);
  assert.deepEqual(first.blocks[3].content.items,['Story of a hard conversation (10 min)']);
  assert.equal(course.course.createdAt,'2026-11-16T12:00:00Z');
  assert.ok(course.course.sections.flatMap(s=>s.lessons).every(l=>l.blocks.every(b=>b.appearance.width==='reading')));
});

test('a Scroll id is made from words, never starts with a digit, and a name that gives nothing falls back', () => {
  assert.equal(scrollId('The calm opening!','x'),'the-calm-opening');
  assert.equal(scrollId('Écoute active','x'),'ecoute-active');
  assert.equal(scrollId('1. First','x'),'x');
  assert.equal(scrollId('???','module-1'),'module-1');
  assert.ok(scrollId('a'.repeat(100),'x').length<=40);
  assert.equal(scrollFileName('Coaching: first'),'Coaching- first.openscroll');
});

test('the package is a zip with course.json and nothing else, that reads back as the course', async () => {
  const course=scrollCourse(design(),{title:'Coaching first',now:'2026-11-16T12:00:00Z'});
  const zip=await JSZip.loadAsync(await scrollBytes(course,JSZip));
  assert.deepEqual(Object.keys(zip.files),['course.json']);
  assert.deepEqual(JSON.parse(await zip.file('course.json').async('string')),course);
});

// OpenScroll's own schema (its schema/course-v3.schema.json) is the real test of the format; it lives in another project, so this runs only when
// OPENSCROLL_SCHEMA names that file and a JSON Schema validator (ajv) can be found.
const SCHEMA=process.env.OPENSCROLL_SCHEMA||'';
function ajvOrNull(){try{return createRequire(path.join(process.env.OPENSCROLL_AJV_DIR||process.cwd(),'x.js'))('ajv/dist/2020');}catch{return null;}}
test('the Scroll course passes OpenScroll\'s own JSON schema',{skip:!(SCHEMA&&fs.existsSync(SCHEMA)&&ajvOrNull())&&'set OPENSCROLL_SCHEMA to OpenScroll\'s course-v3.schema.json, and OPENSCROLL_AJV_DIR to a folder where ajv is installed'},() => {
  const Ajv=ajvOrNull().default||ajvOrNull();
  const validate=new Ajv({allErrors:true,strict:false}).compile(JSON.parse(fs.readFileSync(SCHEMA,'utf8')));
  const course=scrollCourse(design(),{title:'Coaching first',author:'A. Writer',now:'2026-11-16T12:00:00Z'});
  assert.equal(validate(course),true,JSON.stringify(validate.errors?.slice(0,3)));
});

test('OpenAssess gets a slot for each assessment, objective and number, tagged with the objective, its level, the module and the lesson', () => {
  const rows=itemSlots(design());
  assert.equal(rows.length,4,'the quiz has two slots for each of its two objectives; a check with none or no objective gives none');
  assert.deepEqual(rows.map(r=>r.ID),['guide-o1-s2-1','guide-o1-s2-2','guide-o4-s2-1','guide-o4-s2-2']);
  const [first]=rows;
  assert.equal(first.Type,'ES');
  assert.equal(first.Chapter,'Talk first');
  assert.equal(first.Section,'The coaching conversation');
  assert.equal(first.Bloom,'understand');
  assert.equal(first.Objective,'Learners will be able to explain the three steps of a coaching conversation.');
  assert.match(first.Question,/^\[Slot\] Write a question that checks: Learners will be able to explain/);
  assert.equal(first.Tags,'slot; quiz; Understand');
  assert.equal(rows[2].Bloom,'evaluate','the objective\'s own level, not the quiz\'s');
});

test('the slots are a CSV with the columns OpenAssess reads, UTF-8 with a mark, and a plan with no slots says what to do', () => {
  const csv=itemSlotsCsv(design());
  assert.ok(csv.startsWith('\uFEFFType,Question,Points,Chapter,Section,Objective,Bloom,Tags,ID\r\n'));
  assert.equal(csv.split('\r\n').length,6);
  assert.deepEqual(SLOT_COLUMNS,['Type','Question','Points','Chapter','Section','Objective','Bloom','Tags','ID']);
  const d=design();
  d.assessments.forEach(a=>{a.slots=0;});
  assert.throws(()=>itemSlotsCsv(d),/There are no item slots yet/);
});
