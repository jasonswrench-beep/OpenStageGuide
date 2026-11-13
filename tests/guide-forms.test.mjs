import test from 'node:test';
import assert from 'node:assert/strict';
import {DOMParser,XMLSerializer} from '@xmldom/xmldom';
import {createDemoAdapter} from '../src/guide-word.mjs';
import {createProjectAdapter} from '../src/guide-project-adapter.mjs';
import {updatePlan,guideIssues} from '../src/guide-model.mjs';
import {parseTitle,titleText,lettersFor,numberForms,readForms,formPlan,newFormId,readSettings,writeSettings,isFormId,resolveScheme} from '../src/guide-forms.mjs';
import {formTitleBlock,citationControl} from '../src/guide-forms-ooxml.mjs';
import {para,run,makeContext} from '../src/guide-ooxml.mjs';
import {pane,submit,addModule,addLesson,click} from './helpers/guide-harness.mjs';

// ---- paragraphs written by hand, the way an adapter reports them
const P=(style,text,marks)=>({style,text,...(marks?{marks}:{})});
const mod=(n,t)=>P('OSG Module',`Module ${n}: ${t}`);
const form=(id,text)=>P('OSG Form Title',text,{bookmarks:id?[`osg_form_${id}`]:[],refs:[]});
const cite=(id,text,lead='Give out ')=>P('Normal',`${lead}${text}`,{bookmarks:[],refs:[{id,text}]});
const titleOf=()=>P('OSG Guide Title','A guide');

test('a form title reads as a label, a number, a colon and a title, and is written back the same way',()=>{
  assert.deepEqual(parseTitle('Handout 5.4: Standards in depth'),{label:'Handout',number:'5.4',title:'Standards in depth'});
  assert.deepEqual(parseTitle('Worksheet B: Goals'),{label:'Worksheet',number:'B',title:'Goals'});
  assert.deepEqual(parseTitle('Handout: Goals'),{label:'Handout',number:'',title:'Goals'});
  assert.deepEqual(parseTitle('Time Management: goals'),{label:'Time Management',number:'',title:'goals'},'a label of several words is not mistaken for a label and a number');
  assert.deepEqual(parseTitle('  Form AB:   x  '),{label:'Form',number:'AB',title:'x'});
  assert.deepEqual(parseTitle('Handout 5.4: Why: a question'),{label:'Handout',number:'5.4',title:'Why: a question'},'only the first colon divides');
  assert.equal(parseTitle('Standards in depth'),null);
  assert.equal(parseTitle('Handout 5.4:'),null,'a title is needed');
  assert.equal(titleText({label:'Handout',number:'5.4',title:'Standards'}),'Handout 5.4: Standards');
  assert.equal(titleText({label:'Cards',number:'',title:'Roles'}),'Cards: Roles');
  for(const t of ['Handout 5.4: Standards','Worksheet B: Goals','Handout: Goals'])assert.equal(titleText(parseTitle(t)),t);
});

test('letters run A to Z and then AA, AB',()=>{
  assert.deepEqual([0,1,25,26,27,51,52,701,702].map(lettersFor),['A','B','Z','AA','AB','AZ','BA','ZZ','AAA']);
});

test('numbering by unit gives a form its module and its place in it; forms outside a module take letters; the Appendix is outside',()=>{
  const paras=[form('f6','Handout: Front'),mod(1,'One'),form('f1','Handout: A'),form('f2','Handout: B'),mod(2,'Two'),form('f3','Handout: C'),P('heading 2','Appendix'),form('f4','Handout: D'),form('f5','Handout: E')];
  const read=readForms(paras);
  assert.equal(read.hasModules,true);
  assert.deepEqual(read.forms.map(f=>[f.id,f.module,f.appendix]),[['f6',0,false],['f1',1,false],['f2',1,false],['f3',2,false],['f4',0,true],['f5',0,true]]);
  assert.deepEqual(numberForms(read.forms,'auto',true),['A','1.1','1.2','2.1','B','C']);
  assert.deepEqual(numberForms(read.forms,'unit',true),['A','1.1','1.2','2.1','B','C']);
  assert.deepEqual(numberForms(read.forms,'letter',true),['A','B','C','D','E','F']);
  assert.deepEqual(numberForms(read.forms,'sequence',true),['1','2','3','4','5','6']);
  assert.deepEqual(numberForms(read.forms,'none',true),['','','','','','']);
  assert.equal(resolveScheme('auto',false),'letter','with no modules the automatic scheme is letters');
  assert.equal(resolveScheme(undefined,true),'unit');
  assert.equal(resolveScheme('sequence',true),'sequence');
});

test('a module after the Appendix brings the forms back into modules',()=>{
  const read=readForms([mod(1,'One'),P('heading 2','Appendix'),form('f1','Handout: A'),mod(2,'Two'),form('f2','Handout: B')]);
  assert.deepEqual(read.forms.map(f=>f.module),[0,2]);
  assert.deepEqual(numberForms(read.forms,'unit',true),['A','2.1']);
});

test('Update guide renumbers a form, keeps its label and title, and leaves right ones alone',()=>{
  const paras=[titleOf(),mod(1,'One'),form('f1','Handout 9: Goals'),form('f2','Worksheet 1.2: Plans'),form('f3','Worksheet 7: Moves')];
  const plan=formPlan(paras,{});
  assert.deepEqual(plan.edits,[{kind:'form',index:2,text:'Handout 1.1: Goals'},{kind:'form',index:4,text:'Worksheet 1.3: Moves'}]);
  assert.equal(plan.scheme,'unit');
  const letters=formPlan(paras,{scheme:'letter'});
  assert.deepEqual(letters.edits.map(e=>e.text),['Handout A: Goals','Worksheet B: Plans','Worksheet C: Moves']);
  const none=formPlan([form('f1','Handout 4: Goals')],{scheme:'none'});
  assert.deepEqual(none.edits,[{kind:'form',index:0,text:'Handout: Goals'}]);
});

test('a citation follows its form: its words are rewritten when the number or the title changes, each one by its place in the paragraph',()=>{
  const twice=P('Normal','Use Handout 9: Goals, then Handout 9: Goals again.',{refs:[{id:'f1',text:'Handout 9: Goals'},{id:'f1',text:'Handout 9: Goals'}]});
  const paras=[titleOf(),mod(1,'One'),form('f1','Handout 1.1: Goals'),cite('f1','Handout 1.1: Goals'),cite('f1','Handout 4: Old name'),twice];
  const edits=formPlan(paras,{}).edits;
  assert.deepEqual(edits,[
    {kind:'citation',index:4,id:'f1',nth:0,from:'Handout 4: Old name',to:'Handout 1.1: Goals'},
    {kind:'citation',index:5,id:'f1',nth:0,from:'Handout 9: Goals',to:'Handout 1.1: Goals'},
    {kind:'citation',index:5,id:'f1',nth:1,from:'Handout 9: Goals',to:'Handout 1.1: Goals'}
  ]);
});

test('what is wrong with forms and citations is said in plain words, once',()=>{
  const paras=[titleOf(),mod(1,'One'),
    form('f1','Handout 1.1: Goals'),form('f1','Handout 1.2: Copy'),form(null,'Handout 1.3: No identity'),form('f4','Just words'),form('f5','Handout 1.5: Never cited'),
    cite('f9','Handout 3: Gone'),cite('f9','Handout 3: Gone'),cite('f1','Handout 1.1: Goals')];
  const plan=updatePlan(paras);
  const issues=guideIssues(plan).filter(i=>/^FORM-/.test(i.rule));
  const by=r=>issues.filter(i=>i.rule===r);
  assert.equal(by('FORM-MISSING').length,2,'each paragraph that cites a form that is gone is told once');
  assert.ok(by('FORM-MISSING').every(i=>i.level==='fail'));
  assert.equal(by('FORM-DUPLICATE').length,1);assert.match(by('FORM-DUPLICATE')[0].message,/Handout 1\.2: Copy/);
  assert.equal(by('FORM-NO-ID').length,1,'one form has no bookmark');assert.match(by('FORM-NO-ID')[0].message,/No identity/);
  assert.equal(by('FORM-NAME').length,1);assert.match(by('FORM-NAME')[0].message,/Just words/);
  assert.equal(by('FORM-UNCITED').length,2,'two forms are never cited');assert.ok(by('FORM-UNCITED').some(i=>/Never cited/.test(i.message)));
  assert.ok(issues.every(i=>Number.isInteger(i.paragraph)),'each issue knows its paragraph so the pane can go there');
});

test('a guide that lists forms without citing any is not nagged; one that cites some is told about the rest',()=>{
  const text=P('OSG Form Text','Something to read.');
  const quiet=updatePlan([titleOf(),mod(1,'One'),form('f1','Handout 1.1: A'),text,form('f2','Handout 1.2: B'),text]);
  assert.deepEqual(guideIssues(quiet).filter(i=>/^FORM-/.test(i.rule)),[]);
  const some=updatePlan([titleOf(),mod(1,'One'),form('f1','Handout 1.1: A'),text,form('f2','Handout 1.2: B'),text,cite('f1','Handout 1.1: A')]);
  assert.deepEqual(guideIssues(some).filter(i=>/^FORM-/.test(i.rule)).map(i=>i.rule),['FORM-PRINT-NONE','FORM-UNCITED'],'the cited form has no print line yet, and the other is not cited');
});

test('a document that reports no marks has no forms and no form issues',()=>{
  const plan=updatePlan([titleOf(),mod(1,'One'),P('OSG Form Title','Handout 1.1: A')]);
  assert.equal(plan.forms.read.forms.length,1,'the form is there, but cannot be cited');
  assert.deepEqual(plan.forms.edits,[]);
  const word=updatePlan([titleOf(),mod(1,'One'),P('Normal','Handout 1.1: A')]);
  assert.deepEqual(word.forms.read.forms,[]);
  assert.deepEqual(word.forms.issues,[]);
});

test('a new form identity is one that no form and no citation uses, so a deleted form\'s id is never given to another',()=>{
  assert.equal(newFormId([]),'f1');
  assert.equal(newFormId([form('f1','Handout: A'),form('f3','Handout: B')]),'f2');
  assert.equal(newFormId([form('f1','Handout: A'),form('f3','Handout: B'),cite('f2','Handout: Gone')]),'f4');
  assert.equal(isFormId('f12'),true);assert.equal(isFormId('F1'),false);assert.equal(isFormId('f 1'),false);assert.equal(isFormId('1f'),false);assert.equal(isFormId('f"x'),false);
});

test('the settings read as the automatic scheme when nothing sensible is stored',()=>{
  assert.deepEqual(readSettings(null),{scheme:'auto'});
  assert.deepEqual(readSettings(''),{scheme:'auto'});
  assert.deepEqual(readSettings('not json'),{scheme:'auto'});
  assert.deepEqual(readSettings('{"scheme":"nonsense"}'),{scheme:'auto'});
  assert.deepEqual(readSettings('{"scheme":"letter"}'),{scheme:'letter'});
  for(const scheme of ['auto','unit','letter','sequence','none'])assert.deepEqual(readSettings(writeSettings({scheme})),{scheme});
  assert.deepEqual(readSettings(writeSettings({scheme:'bogus'})),{scheme:'auto'});
});

// ---- the Word-format pieces, and a document that reads them back
test('a form title is a paragraph in the form title style with a bookmark, and a citation is a tagged content control holding words',()=>{
  const title=formTitleBlock({id:'f2',label:'Handout',number:'1.1',title:'Goals & plans'});
  assert.match(title,/<w:pStyle w:val="OSGFormTitle"\/>/);
  assert.match(title,/<w:bookmarkStart w:id="9002" w:name="osg_form_f2"\/>/);
  assert.match(title,/<w:bookmarkEnd w:id="9002"\/>/);
  assert.match(title,/Handout 1\.1: Goals &amp; plans/);
  const ctx=makeContext();
  const c=citationControl({id:'f2',text:'Handout 1.1: Goals'},ctx);
  assert.match(c,/<w:tag w:val="osg-ref:f2"\/>/);
  assert.match(c,/<w:sdtContent>.*Handout 1\.1: Goals.*<\/w:sdtContent>/);
  assert.equal(ctx.sdt,1);
  assert.throws(()=>formTitleBlock({id:'F 2',label:'Handout',number:'',title:'x'}),/needs an identity/);
  assert.throws(()=>formTitleBlock({id:'f2',label:'Handout',number:'',title:'   '}),/Type a title/);
  assert.throws(()=>citationControl({id:'"><x',text:'t'},makeContext()),/needs an identity/);
});

async function guideWithForms(){
  const a=pane();
  await a.p.refresh();
  a.type('guide-title','Forms guide');await submit(a,'guide-start-form');
  await addModule(a,'One');await addLesson(a,'First');
  await addModule(a,'Two');await addLesson(a,'Second');
  return a;
}
const index=(a,text)=>a.adapter.snapshot().findIndex(p=>p.text.includes(text));
async function insertAfter(a,after,build,opts={}){await a.adapter.insert(build,{where:'after',after,spare:false,...opts});}

test('forms put into a guide are read back with their marks, Update guide corrects numbers and citations, and a second Update changes nothing',async()=>{
  const a=await guideWithForms();
  const lesson1=index(a,'Lesson 1.1');
  await insertAfter(a,lesson1+1,ctx=>formTitleBlock({id:'f1',label:'Handout',number:'',title:'Goals'}));
  const lesson2=index(a,'Lesson 2.1');
  await insertAfter(a,lesson2+1,ctx=>formTitleBlock({id:'f2',label:'Worksheet',number:'7',title:'Plans'})+para(null,run('Hand out ')+citationControl({id:'f1',text:'Handout 9: Old'},ctx)+run(' first.')));
  const list=await a.adapter.paragraphs();
  const marked=list.filter(p=>p.marks);
  assert.deepEqual(marked.map(p=>p.marks.bookmarks),[['osg_form_f1'],['osg_form_f2'],[]]);
  assert.deepEqual(marked[2].marks.refs,[{id:'f1',text:'Handout 9: Old'}]);

  const plan=updatePlan(list);
  assert.deepEqual(plan.forms.edits.map(e=>e.kind),['form','form','citation']);
  await a.adapter.applyEdits(plan.edits);
  const after=a.adapter.snapshot().map(p=>p.text.trim());
  assert.ok(after.includes('Handout 1.1: Goals'),'the first form is in module 1');
  assert.ok(after.includes('Worksheet 2.1: Plans'),'the second is the first of module 2, and keeps its label');
  assert.ok(after.includes('Hand out Handout 1.1: Goals first.'),'the citation says what the form is now called, with the words around it kept');
  assert.deepEqual(updatePlan(await a.adapter.paragraphs()).edits,[],'a second Update has nothing left to do');
  // The Word-format package carries the bookmark and the tagged citation, so a Word export has them as they are.
  const pkg=await a.adapter.sourcePackage();
  assert.match(pkg,/osg_form_f1/);assert.match(pkg,/osg-ref:f1/);
  assert.match(pkg,/Handout 1\.1: Goals/);
  assert.doesNotMatch(pkg,/Handout 9: Old/);
});

test('moving a form to another module renumbers it, its neighbours and every citation of it',async()=>{
  const a=await guideWithForms();
  await insertAfter(a,index(a,'Lesson 1.1')+1,ctx=>formTitleBlock({id:'f1',label:'Handout',number:'1.1',title:'Goals'})+formTitleBlock({id:'f2',label:'Handout',number:'1.2',title:'Plans'})+para(null,run('See ')+citationControl({id:'f2',text:'Handout 1.2: Plans'},ctx)));
  const f2At=index(a,'Handout 1.2: Plans');
  const lesson2End=index(a,'Lesson 2.1')+1;
  await a.adapter.moveRange({start:f2At,end:f2At,guard:'Handout 1.2: Plans',after:lesson2End});
  await a.adapter.applyEdits(updatePlan(await a.adapter.paragraphs()).edits);
  const texts=a.adapter.snapshot().map(p=>p.text.trim());
  assert.ok(texts.includes('Handout 1.1: Goals'));
  assert.ok(texts.includes('Handout 2.1: Plans'),'the moved form is now the first of module 2');
  assert.ok(texts.includes('See Handout 2.1: Plans'),'and the citation says so');
});

test('the scheme chosen for the guide is kept with the document and used by Update guide in the pane',async()=>{
  const a=await guideWithForms();
  await insertAfter(a,index(a,'Lesson 1.1')+1,ctx=>formTitleBlock({id:'f1',label:'Handout',number:'1.1',title:'Goals'})+formTitleBlock({id:'f2',label:'Handout',number:'1.2',title:'Plans'}));
  await a.adapter.setFormSettings(writeSettings({scheme:'letter'}));
  await a.p.refresh();
  assert.match(a.n('guide-update-note').textContent,/out of date/,'the pane sees forms that do not follow the chosen scheme');
  await click(a,'guide-update');
  const texts=a.adapter.snapshot().map(p=>p.text.trim());
  assert.ok(texts.includes('Handout A: Goals'));assert.ok(texts.includes('Handout B: Plans'));
  assert.equal(a.err(),'');
});

test('a form that is cited but deleted is a Review problem the pane lists',async()=>{
  const a=await guideWithForms();
  await insertAfter(a,index(a,'Lesson 1.1')+1,ctx=>para(null,run('Give out ')+citationControl({id:'f9',text:'Handout 1.1: Gone'},ctx)));
  const issues=guideIssues(updatePlan(await a.adapter.paragraphs())).filter(i=>i.rule==='FORM-MISSING');
  assert.equal(issues.length,1);
  assert.equal(issues[0].level,'fail');
});

// ---- the project file keeps all of it
const textOfFile=async build=>{
  const files=[];
  const raw=createDemoAdapter({DOMParserClass:DOMParser,SerializerClass:XMLSerializer});
  const timers=[];
  const doc=createProjectAdapter(raw,{read:async()=>files.at(-1)??null,write:async t=>{files.push(t);},soon:fn=>{timers.push(fn);return timers.length-1;},cancel:id=>{timers[id]=null;}});
  await doc.ready;await build(doc);
  for(const fn of timers.splice(0))fn?.();
  await doc.flush();
  return files.at(-1);
};

test('a project file keeps the forms, the citations and the numbering scheme, and the marks come back from the paragraphs',async()=>{
  const file=await textOfFile(async doc=>{
    await doc.insert(ctx=>formTitleBlock({id:'f1',label:'Handout',number:'A',title:'Goals'})+para(null,run('See ')+citationControl({id:'f1',text:'Handout A: Goals'},ctx)),{where:'start',spare:false});
    await doc.setFormSettings(writeSettings({scheme:'letter'}));
  });
  const saved=JSON.parse(file).guide;
  assert.equal(saved.forms,writeSettings({scheme:'letter'}));
  assert.ok(saved.paragraphs.every(p=>!('marks' in p)),'marks are derived, never stored');
  const raw=createDemoAdapter({DOMParserClass:DOMParser,SerializerClass:XMLSerializer});
  const doc=createProjectAdapter(raw,{read:async()=>file,write:async()=>{}});
  await doc.ready;
  const list=await doc.paragraphs();
  assert.deepEqual(list[0].marks,{bookmarks:['osg_form_f1'],refs:[]});
  assert.deepEqual(list[1].marks,{bookmarks:[],refs:[{id:'f1',text:'Handout A: Goals'}]});
  assert.equal(await doc.formSettings(),writeSettings({scheme:'letter'}));
  assert.deepEqual(updatePlan(list,{forms:readSettings(await doc.formSettings())}).edits,[]);
});

test('Update guide still works on a document that has no forms at all, and on one that cannot keep form settings',async()=>{
  const a=await guideWithForms();
  assert.deepEqual(updatePlan(await a.adapter.paragraphs()).forms.read.forms,[]);
  delete a.adapter.formSettings;
  await a.p.refresh();
  assert.equal(a.err(),'','a document with no way to keep form settings is not an error');
});
