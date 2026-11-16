import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import {spawnSync} from 'node:child_process';
import JSZip from 'jszip';
import {screensOf} from '../src/guide-elearning.mjs';
import {authorProject,slideFor,handoffNotes,authorZipName,metaWords,AUTHOR_SCHEMA_VERSION} from '../src/guide-author.mjs';
import {zipWorkbook} from '../src/guide-workbook.mjs';
import {started,click,addModule,addLesson} from './helpers/guide-harness.mjs';

const P=(style,text)=>({style,text});
const FAC=new Set(['Narration','Visual direction','Branch','Animation note','Developer note']);
const cue=(label,text)=>P(FAC.has(label)?'OSG Cue Facilitator':'OSG Cue',`${label}\t${text}`);
const DOC=[
  P('OSG Module','Module 1: Basics'),P('OSG Lesson','Lesson 1.1: Welcome'),
  cue('Screen title','Hello & welcome'),cue('Visual direction','A sunrise.'),cue('On-screen text','Welcome to the course.'),P('OSG Cue Body','It takes ten minutes.'),
  cue('Narration','Hello, and welcome to the course.'),P('OSG Cue Body','Let us begin.'),
  cue('Interaction','Click the sun.'),cue('Feedback','Well done.'),
  cue('Screen title','Only a title'),cue('Narration','Just a voice.'),
  P('OSG Lesson','Lesson 1.2: Practice'),cue('Screen title','Try it'),cue('On-screen text','Your turn.'),cue('Knowledge check','Question 2')
];
const screens=screensOf(DOC);
const text=(files,name)=>files.find(f=>f.name===name)?.text;

test('the project has a folder, its manifest and theme, a slide for each screen and a notes file beside it',()=>{
  const files=authorProject(screens,{title:'Safety: the basics',today:'2026-11-06'});
  assert.deepEqual(files.map(f=>f.name),[
    'Safety - the basics.osproj/project.json','Safety - the basics.osproj/theme.json',
    'Safety - the basics.osproj/slides/slide-001.json','Safety - the basics.osproj/slides/slide-002.json','Safety - the basics.osproj/slides/slide-003.json',
    'Safety - the basics - hand-off notes.txt']);
  const project=JSON.parse(text(files,'Safety - the basics.osproj/project.json'));
  assert.equal(project.schemaVersion,AUTHOR_SCHEMA_VERSION);
  assert.deepEqual(project.meta,{title:'Safety: the basics',author:'',language:'en-US',created:'2026-11-06',modified:'2026-11-06'});
  assert.deepEqual(project.scenes,[{id:'scene-1',name:'Lesson 1.1: Welcome',slides:['slide-001','slide-002']},{id:'scene-2',name:'Lesson 1.2: Practice',slides:['slide-003']}]);
  assert.deepEqual(project.theme,{ref:'theme.json'});
  assert.equal(authorZipName('Safety: the basics'),'Safety - the basics - OpenStage Author project.zip');
});

test('the project is written in the current format of OpenStage Author, 0.49.0', () => {
  assert.equal(AUTHOR_SCHEMA_VERSION, '0.49.0');
  const project = JSON.parse(text(authorProject(screens, { title: 'Safety basics', today: '2026-11-16' }), 'Safety basics.osproj/project.json'));
  assert.equal(project.schemaVersion, '0.49.0');
});

test('the version and the course code are written only when they hold something, trimmed, and refused when too long', () => {
  const meta = (extra) => JSON.parse(text(authorProject(screens, { title: 'Safety basics', today: '2026-11-16', ...extra }), 'Safety basics.osproj/project.json')).meta;
  assert.deepEqual(Object.keys(meta({})), ['title', 'author', 'language', 'created', 'modified']);
  assert.deepEqual(Object.keys(meta({ version: '   ', courseCode: '' })), ['title', 'author', 'language', 'created', 'modified']);
  assert.equal(meta({ version: ' 1.0 ', courseCode: 'CMM  360' }).version, '1.0');
  assert.equal(meta({ version: '1.0', courseCode: 'CMM  360' }).courseCode, 'CMM 360');
  assert.throws(() => meta({ version: 'x'.repeat(21) }), /version can be 20 characters at most/);
  assert.throws(() => meta({ courseCode: 'x'.repeat(41) }), /course code can be 40 characters at most/);
  assert.deepEqual(metaWords({ version: 'x'.repeat(20), courseCode: 'y'.repeat(40) }), { version: 'x'.repeat(20), courseCode: 'y'.repeat(40) });
});

test('a screen becomes a slide: a heading, the on-screen text, and the narration as the voice-over script',()=>{
  const s=slideFor(screens[0],1,'en-US');
  assert.equal(s.id,'slide-001');assert.equal(s.name,'Hello & welcome');
  const [heading,body]=s.baseLayer.objects;
  assert.equal(heading.content.html,'<h1>Hello &amp; welcome</h1>');
  assert.deepEqual(heading.accessibility,{altText:'Hello & welcome',role:'heading',tabOrder:1,decorative:false});
  assert.equal(heading.narration,null);
  assert.equal(body.content.html,'<p>Welcome to the course.</p><p>It takes ten minutes.</p>');
  assert.equal(body.accessibility.altText,'Welcome to the course. It takes ten minutes.');
  assert.deepEqual(body.narration,{script:'Hello, and welcome to the course. Let us begin.',voice:'af_heart',speed:1,lang:'en-US',pronunciations:[],audioCache:null});
  // Narration with no on-screen text rides on the heading; on-screen text with no narration has none.
  assert.equal(slideFor(screens[1],2,'en-US').baseLayer.objects.length,1);
  assert.equal(slideFor(screens[1],2,'en-US').baseLayer.objects[0].narration.script,'Just a voice.');
  assert.equal(slideFor(screens[2],3,'en-US').baseLayer.objects[1].narration,null);
  assert.equal(new Set(s.baseLayer.objects.map(o=>o.id)).size,2);
});

test('the notes file lists what the project has no place for, screen by screen',()=>{
  const notes=handoffNotes(screens,'Safety');
  assert.match(notes,/^Safety - notes for building in OpenStage Author\n/);
  assert.match(notes,/Lesson 1\.1: Welcome\n=+\n\nScreen 1: Hello & welcome\n  Visual direction: A sunrise\.\n  Interaction: Click the sun\.\n  Feedback: Well done\.\n\nScreen 2: Only a title\n  \(nothing more\)\n\n/);
  assert.match(notes,/Screen 1: Try it\n  Knowledge check: Question 2\n$/);
  assert.doesNotMatch(notes,/Hello, and welcome to the course/,'the narration is in the project, not repeated here');
});

test('a project needs at least one screen',()=>{
  assert.throws(()=>authorProject([],{title:'x'}),/There are no screens to send yet/);
});

test('the pane sends the storyboard as a zip that opens as an OpenStage Author project',async()=>{
  const a=await started('Safety basics');
  await addModule(a,'Module one');await addLesson(a,'Lesson one');
  await click(a,'guide-el-author');
  assert.match(a.err(),/There are no screens yet/);
  a.type('guide-el-title','Welcome');await click(a,'guide-el-add');
  a.adapter.typeOver(a.rows().findIndex(r=>/^OSG Cue Facilitator: [^A-Za-z]*Narration/.test(r)),'Narration\tHello there.');
  await a.p.refresh();
  await click(a,'guide-el-author');
  assert.equal(a.err(),'');
  const d=a.adapter.downloads().at(-1);
  assert.equal(d.fileName,'Safety basics - OpenStage Author project.zip');assert.equal(d.type,'application/zip');
  const zip=await JSZip.loadAsync(Buffer.from(d.base64,'base64'));
  assert.deepEqual(Object.keys(zip.files).sort(),['Safety basics - hand-off notes.txt','Safety basics.osproj/project.json','Safety basics.osproj/slides/slide-001.json','Safety basics.osproj/theme.json']);
  const slide=JSON.parse(await zip.file('Safety basics.osproj/slides/slide-001.json').async('string'));
  assert.equal(slide.name,'Welcome');
  assert.match(a.n('guide-status').textContent,/1 screen sent: the project was saved as "Safety basics - OpenStage Author project\.zip"/);
});

// The author's own validator is the real test of the format. It lives in another project (OpenStage Author, tools/validate.py), so this runs
// only when OPENSTAGE_AUTHOR_VALIDATOR names that file, or when Author is in its usual place on the author's PC (C:\Dev\openstage-author).
const VALIDATOR=process.env.OPENSTAGE_AUTHOR_VALIDATOR||['C:/Dev/openstage-author/tools/validate.py'].find(f=>fs.existsSync(f))||'';
test('the project passes OpenStage Author\'s own validator with no errors',{skip:!(VALIDATOR&&fs.existsSync(VALIDATOR))&&'set OPENSTAGE_AUTHOR_VALIDATOR to the OpenStage Author validate.py to run this'},async()=>{
  const dir=fs.mkdtempSync(path.join(os.tmpdir(),'osg-author-'));
  try{
    const bytes=await zipWorkbook(authorProject(screens,{title:'Safety basics'}),JSZip,'uint8array');
    const zip=await JSZip.loadAsync(bytes);
    for(const [name,entry] of Object.entries(zip.files))if(!entry.dir){fs.mkdirSync(path.dirname(path.join(dir,name)),{recursive:true});fs.writeFileSync(path.join(dir,name),await entry.async('nodebuffer'));}
    const run=spawnSync('python',['-B',VALIDATOR,path.join(dir,'Safety basics.osproj')],{encoding:'utf8',cwd:path.dirname(path.dirname(VALIDATOR))});
    assert.equal(run.status,0,`${run.stdout}\n${run.stderr}`);
    assert.match(run.stdout,/0 errors/);
  }finally{fs.rmSync(dir,{recursive:true,force:true});}
});
