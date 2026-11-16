// The rating grids and self-scored instruments of a guide, sent to PrismStatistics as a .pstat file (src/guide-pstat.mjs). What it writes is checked with the
// shared core's own checker (@openstage/pstat, installed from its release), which refuses everything PrismStatistics would refuse.
import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import {DOMParser,XMLSerializer} from '@xmldom/xmldom';
import {checkPstat,readPstat} from '@openstage/pstat';
import {createDemoAdapter} from '../src/guide-word.mjs';
import {buildForm} from '../src/guide-forms-templates.mjs';
import {newFormId} from '../src/guide-forms.mjs';
import {ratingGrid,scoreInstrument} from '../src/guide-forms-ooxml.mjs';
import {readInstruments,pstatProject,pstatBytes,pstatFileName,scoreScale,MAX_ROWS} from '../src/guide-pstat.mjs';

const io={parse:x=>new DOMParser().parseFromString(x,'application/xml'),serialize:e=>new XMLSerializer().serializeToString(e)};
const demo=()=>createDemoAdapter({DOMParserClass:DOMParser,SerializerClass:XMLSerializer});
async function addTemplate(a,template,title){
  const id=newFormId(await a.paragraphs());
  await a.insert(ctx=>buildForm({id,template,title},ctx),{where:'start',spare:false});
}
const entries=a=>a.exportState().paragraphs;

test('the two templates are read back: a rating form with evidence, and a self-scored instrument with its scales and what the totals mean',async()=>{
  const a=demo();
  await addTemplate(a,'instrument','Teamwork check');
  await addTemplate(a,'rating','Observer ratings');
  const found=readInstruments(entries(a),io);
  assert.equal(found.length,2);
  const rating=found.find(i=>i.kind==='rating'),inst=found.find(i=>i.kind==='instrument');
  assert.deepEqual(rating.statements,['First statement','Second statement','Third statement']);
  assert.deepEqual(rating.marks,['1','2','3','4','5']);
  assert.equal(rating.evidence,true);
  assert.match(rating.formTitle,/Observer ratings/);
  assert.equal(inst.statements.length,4);
  assert.deepEqual(inst.scales,[{name:'First scale',items:[1,2],divide:null},{name:'Second scale',items:[3,4],divide:null}]);
  assert.deepEqual(inst.bands.map(b=>[b.group,b.from,b.to]),[['',2,5],['',6,10]]);
  assert.deepEqual([inst.scale.low,inst.scale.high,inst.scale.lowLabel,inst.scale.highLabel],[1,5,'never','always']);
  assert.match(inst.formTitle,/Teamwork check/);
});

test('a table that is neither is left alone, and a guide with none has nothing to send',async()=>{
  const a=demo();
  await addTemplate(a,'action','Action plan');
  await addTemplate(a,'worksheet','Worksheet');
  assert.deepEqual(readInstruments(entries(a),io),[]);
  assert.throws(()=>pstatProject([]),/Choose at least one/);
});

test('the scale of an instrument is read from its score column, in the words the form had',()=>{
  assert.deepEqual(scoreScale('Score (1 (never) to 5 (always))'),{text:'1 (never) to 5 (always)',low:1,high:5,lowLabel:'never',highLabel:'always'});
  assert.deepEqual(scoreScale('Score (0 to 3)'),{text:'0 to 3',low:0,high:3,lowLabel:'',highLabel:''});
  assert.equal(scoreScale('Score').low,null);
  assert.equal(scoreScale('Score (agree)').low,null);
});

test('the file is a PrismStatistics project the core\'s checker accepts: one variable for each statement, its label, its points, and the empty rows',async()=>{
  const a=demo();
  await addTemplate(a,'rating','Observer ratings');
  const ins=readInstruments(entries(a),io);
  const project=pstatProject(ins,{guideTitle:'Safety basics',rows:3,appVersion:'OpenStage Guide 0.1.0-test'});
  assert.equal(project.format,'PrismStatistics project');
  assert.equal(project.version,1);
  assert.equal(project.app_version,'OpenStage Guide 0.1.0-test');
  assert.deepEqual(project.dataset.variables.map(v=>v.name),['respondent','q01','q02','q03']);
  assert.deepEqual(project.dataset.variables.map(v=>v.id),[1,2,3,4]);
  assert.equal(project.dataset.next_id,5);
  assert.deepEqual(project.dataset.variables.slice(1).map(v=>v.label),['First statement','Second statement','Third statement']);
  assert.ok(project.dataset.variables.slice(1).every(v=>v.measure==='ordinal'&&v.value_labels.length===0),'points that are numbers need no labels');
  assert.deepEqual(project.dataset.columns[0],{type:'numeric',values:[1,2,3]});
  assert.deepEqual(project.dataset.columns[1],{type:'numeric',values:[null,null,null]});
  assert.deepEqual(project.operations,[]);
  assert.deepEqual(project.results,[]);
  assert.equal('profile' in project,false,'no profile: the core has none for an instrument yet');
  const bytes=await pstatBytes(project,JSZip);
  assert.deepEqual(checkPstat(bytes),[],'the core\'s checker finds nothing wrong');
  const back=readPstat(bytes);
  assert.equal(back.dataset.name,'Safety basics');
  assert.equal(back.source.report.rows,3);
  assert.equal(back.source.report.variables,4);
  assert.ok(back.source.report.notes.some(n=>/also has a column to write the evidence in/.test(n)));
});

test('the scoring guide goes in the notes the file carries: how each scale is added up and what each range of totals means; the labels name the scale',async()=>{
  const a=demo();
  await addTemplate(a,'instrument','Teamwork check');
  const project=pstatProject(readInstruments(entries(a),io),{guideTitle:'Safety basics',rows:2});
  assert.deepEqual(project.dataset.variables.map(v=>v.label),['Respondent (a number for each person who answers)','First statement (scale: First scale)','Second statement (scale: First scale)','Third statement (scale: Second scale)','Fourth statement (scale: Second scale)']);
  assert.deepEqual(project.dataset.variables[1].value_labels,[{value:1,label:'never'},{value:5,label:'always'}]);
  const notes=project.source.report.notes;
  assert.ok(notes.some(n=>/^Scale "First scale" of "Self-scored instrument": add q01, q02\.$/.test(n)),notes.join(' | '));
  assert.ok(notes.some(n=>/^Scale "Second scale" of "Self-scored instrument": add q03, q04\.$/.test(n)));
  // the ranges of this form are for every scale, so each scale is told both
  for(const scale of ['First scale','Second scale']){
    assert.ok(notes.some(n=>n===`${scale}, a total of 2 to 5: Say here what a low total means.`),`${scale} low`);
    assert.ok(notes.some(n=>n===`${scale}, a total of 6 to 10: Say here what a high total means.`),`${scale} high`);
  }
  assert.ok(notes.some(n=>/in Form: Teamwork check/.test(n)),'and which form the instrument was in');
  assert.deepEqual(checkPstat(await pstatBytes(project,JSZip)),[]);
});

test('an instrument that averages says so, and ranges for one scale go with that scale only',async()=>{
  const a=demo();
  const id=newFormId(await a.paragraphs());
  await a.insert(ctx=>buildForm({id,template:'worksheet',title:'Scales'},ctx)+scoreInstrument({statements:['I plan','I check','I ask','I share'],groups:[{name:'Doing',items:[1,2]},{name:'Together',items:[3,4]}],
    bands:[{group:'Doing',from:1,to:2.5,meaning:'Starting out.'},{group:'Together',from:3,to:5,meaning:'Strong.'}],divide:true,scale:'1 to 5',caption:'Habits'}),{where:'start',spare:false});
  const [ins]=readInstruments(entries(a),io);
  assert.deepEqual(ins.scales,[{name:'Doing',items:[1,2],divide:2},{name:'Together',items:[3,4],divide:2}]);
  const notes=pstatProject([ins],{rows:0}).source.report.notes;
  assert.ok(notes.some(n=>/^Scale "Doing" of "Habits": add q01, q02, then divide the total by 2 for the average\.$/.test(n)),notes.join(' | '));
  assert.ok(notes.some(n=>/^Doing, a score of 1 to 2\.5: Starting out\.$/.test(n)));
  assert.ok(!notes.some(n=>/^Together, a score of 1 to 2\.5/.test(n)),'a range for one scale is not given to the other');
  assert.ok(notes.some(n=>/^Together, a score of 3 to 5: Strong\.$/.test(n)));
});

test('a grid whose points are words gets them as value labels, 1 to the last; several instruments get a prefix each and names that never meet',()=>{
  const ctx={used:new Set(),docPr:1,lists:0,bullets:new Set(),sdt:0,pictures:[]};
  const paras=[
    {style:'OSG Form Title',text:'Form A: One',xml:'<w:p/>'},
    {style:'',text:'',xml:ratingGrid({statements:['I feel ready.','I can do it.'],labels:['Disagree','Neutral','Agree'],ctx,caption:'Readiness'})},
    {style:'OSG Form Title',text:'Form B: Two',xml:'<w:p/>'},
    {style:'',text:'',xml:ratingGrid({statements:['It was useful.'],labels:['1','2','3','4'],ctx,caption:'Use'})},
  ].map(p=>({...p,xml:p.xml.startsWith('<w:tbl')?withNs(p.xml):p.xml}));
  const found=readInstruments(paras,io);
  assert.equal(found.length,2);
  assert.deepEqual(found[0].marks,['Disagree','Neutral','Agree']);
  const project=pstatProject(found,{rows:1});
  assert.deepEqual(project.dataset.variables.map(v=>v.name),['respondent','f1_q01','f1_q02','f2_q01']);
  assert.deepEqual(project.dataset.variables[1].value_labels,[{value:1,label:'Disagree'},{value:2,label:'Neutral'},{value:3,label:'Agree'}]);
  assert.deepEqual(project.dataset.variables[3].value_labels,[]);
  assert.equal(project.dataset.variables[1].label,'Readiness: I feel ready.');
  assert.equal(new Set(project.dataset.variables.map(v=>v.name.toLowerCase())).size,4);
});

function withNs(xml){
  return xml.replace(/^<w:tbl/,'<w:tbl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"');
}

test('the number of people is a whole number from 0, and statements are limited, with a message that says what to do',async()=>{
  const a=demo();
  await addTemplate(a,'rating','Ratings');
  const ins=readInstruments(entries(a),io);
  for(const bad of [-1,1.5,MAX_ROWS+1,'x'])assert.throws(()=>pstatProject(ins,{rows:bad}),/whole number from 0 to 999/,String(bad));
  assert.equal(pstatProject(ins,{rows:0}).dataset.columns[0].values.length,0);
  assert.deepEqual(checkPstat(await pstatBytes(pstatProject(ins,{rows:0}),JSZip)),[],'a file with no rows yet is still a file PrismStatistics opens');
  assert.equal(pstatProject(ins,{rows:MAX_ROWS}).dataset.columns[1].values.length,MAX_ROWS);
  const long={...ins[0],statements:Array.from({length:201},(_,i)=>`Statement ${i+1}`)};
  assert.throws(()=>pstatProject([long]),/up to 200 statements; these have 201/);
});

test('the file is named for the guide and says what is in it',()=>{
  assert.equal(pstatFileName('Safety: basics'),'Safety- basics - instruments.pstat');
  assert.equal(pstatFileName(''),'Guide - instruments.pstat');
  assert.equal(pstatFileName('x'.repeat(200)).length,80+' - instruments.pstat'.length);
});
