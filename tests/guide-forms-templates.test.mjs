// The forms the program starts an author with (src/guide-forms-templates.mjs): each builds, reads back as a form, is numbered by Update guide,
// and is what the element builders make.
import test from 'node:test';
import assert from 'node:assert/strict';
import {DOMParser,XMLSerializer} from '@xmldom/xmldom';
import {createDemoAdapter} from '../src/guide-word.mjs';
import {FORM_TEMPLATES,buildForm,templateById} from '../src/guide-forms-templates.mjs';
import {readForms,formPlan,newFormId,LABELS} from '../src/guide-forms.mjs';

const demo=()=>createDemoAdapter({DOMParserClass:DOMParser,SerializerClass:XMLSerializer});
async function add(a,spec){
  const id=newFormId(await a.paragraphs());
  await a.insert(ctx=>buildForm({id,...spec},ctx),{where:'start',spare:false});
  return id;
}

test('there are templates for the forms the author is likely to want, each with a name, a kind, a note and a body',()=>{
  assert.deepEqual(FORM_TEMPLATES.map(t=>t.id),['reading','worksheet','notes','checklist','rating','instrument','observer','evaluation','action','commitment','rolecards','matchup','twobytwo','forcefield','planning','timeline','techmaterials','roomsetup','roles','prework','followup']);
  for(const t of FORM_TEMPLATES){
    assert.ok(t.label&&t.note.length>10,t.id);
    assert.ok(LABELS.includes(t.kind),`${t.id}: ${t.kind} is a kind of form the guide knows`);
    assert.equal(typeof t.body,'function');
  }
  assert.equal(templateById('worksheet').label,'Blank worksheet');
  assert.equal(templateById('nope'),null);
});

test('every template makes a form that reads back as one, with its own id, the title it was given, and nothing the Review would stop it for',async()=>{
  const a=demo();
  const ids=[];
  for(const t of FORM_TEMPLATES)ids.push(await add(a,{template:t.id,title:`${t.label} for test`}));
  assert.equal(new Set(ids).size,FORM_TEMPLATES.length,'every id is its own');
  const paras=await a.paragraphs();
  const {forms}=readForms(paras);
  assert.equal(forms.length,FORM_TEMPLATES.length);
  for(const t of FORM_TEMPLATES)assert.ok(paras.some(p=>/^OSG Form Title/.test(p.style)&&p.text.endsWith(`: ${t.label} for test`)),t.id);
  const plan=formPlan(paras,{});
  assert.deepEqual(plan.issues.filter(i=>i.level==='fail'),[],'no issue that stops the guide');
  assert.ok(plan.issues.every(i=>i.rule==='FORM-PRINT-GROUP'),'only the ones about how many people there will be: '+JSON.stringify(plan.issues.map(i=>i.rule)));
});

test('Update guide numbers the forms the templates made, in the guide\'s order',async()=>{
  const a=demo();
  await add(a,{template:'worksheet',title:'First'});
  await add(a,{template:'checklist',title:'Second'});
  const plan=formPlan(await a.paragraphs(),{});
  assert.ok(plan.edits.length>=2,'there are titles to number');
  await a.applyEdits(plan.edits);
  const titles=(await a.paragraphs()).filter(p=>/^OSG Form Title/.test(p.style)).map(p=>p.text);
  assert.equal(titles.length,2);
  assert.ok(titles.every(t=>/^[A-Za-z]+ [A-Z0-9.]+: /.test(t)),'numbered: '+titles.join(' | '));
});

test('the kind and who it is for can be chosen; a kind that is not known falls back to the template\'s; a title is needed',async()=>{
  const a=demo();
  await add(a,{template:'worksheet',title:'Mine',label:'Cards',facilitator:true});
  await add(a,{template:'worksheet',title:'Theirs',label:'Not a kind'});
  const ps=(await a.paragraphs()).filter(p=>/^OSG Form Title/.test(p.style));
  const mine=ps.find(p=>p.text.endsWith(': Mine')),theirs=ps.find(p=>p.text.endsWith(': Theirs'));
  assert.match(mine.text,/^Cards/);
  assert.equal(mine.style,'OSG Form Title Facilitator','a form for the facilitator is marked so, and the workbook leaves it out');
  assert.match(theirs.text,/^Worksheet/);
  assert.equal(theirs.style,'OSG Form Title');
  const ctx={used:new Set(),docPr:1,lists:0,bullets:new Set(),sdt:0,pictures:[]};
  assert.throws(()=>buildForm({template:'worksheet',id:'f9',title:'   '},ctx),/Type a title/);
  assert.throws(()=>buildForm({template:'nope',id:'f9',title:'x'},ctx),/Choose a kind of form/);
});

test('a form made from a template has the print line, and its elements are the real thing: check boxes, a header row, lines and fields',async()=>{
  const a=demo();
  await add(a,{template:'observer',title:'Observed'});
  await add(a,{template:'worksheet',title:'Written'});
  const ps=await a.paragraphs();
  assert.equal(ps.filter(p=>p.style==='OSG Form Print').length,2);
  assert.match(ps.find(p=>p.style==='OSG Form Print').text,/^Print: 1 per person/);
  const xml=a.exportState().paragraphs.map(p=>p.xml).join('');
  assert.ok((xml.match(/<w14:checkbox/g)||[]).length>=6,'the observer form has check boxes in its Met and Not met columns');
  assert.match(xml,/<w:tblHeader/,'a header row');
  assert.equal(ps.filter(p=>p.style==='OSG Form Line').length,6,'six lines to write on in the worksheet');
  assert.ok(ps.some(p=>p.style==='OSG Form Field'),'a row of fields');
});

test('the planning kit templates are for the facilitator whatever was ticked, and pre-work and the follow-up roadmap are for everyone',()=>{
  const ctx={used:new Set(),docPr:1,lists:0,bullets:new Set(),sdt:0,pictures:[]};
  for(const id of ['timeline','techmaterials','roomsetup','roles'])assert.match(buildForm({template:id,id:'f9',title:'Mine',facilitator:false},ctx),/OSGFormTitleFacilitator/,id);
  for(const id of ['prework','followup','worksheet'])assert.doesNotMatch(buildForm({template:id,id:'f9',title:'Mine'},ctx),/OSGFormTitleFacilitator/,id);
  assert.ok(FORM_TEMPLATES.filter(t=>t.facilitator).length===4);
});
