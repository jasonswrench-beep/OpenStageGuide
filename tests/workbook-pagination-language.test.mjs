import test from 'node:test';
import assert from 'node:assert/strict';
import {DOMParser} from '@xmldom/xmldom';
import {course,wrenchCourse,addModule,addActivity} from '../src/model.mjs';
import {editDeliveryPlan} from '../src/delivery-plans.mjs';
import {WORKBOOK_FAMILIES,resolveWorkbookBrand,semanticWorkbookFromCourse,semanticWorkbookFromSelected} from '../src/workbook-presentation.mjs';
import {WORKBOOK_BASE_LANGUAGES,WORKBOOK_DEFAULT_BASE_LANGUAGE,renderWorkbookFamily} from '../src/workbook-family-output.mjs';
import {FAMILY_BRAND_FIELDS,prepareFamilyWorkbook,familyReviewText} from '../src/workbook-family-flow.mjs';

const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const DC='http://purl.org/dc/elements/1.1/';
const stamp='2026-10-27T10:00:00.000Z';
const xml=value=>new DOMParser().parseFromString(value,'application/xml');
const parts=doc=>Array.from(doc.getElementsByTagNameNS(W,'p'));
const style=p=>p.getElementsByTagNameNS(W,'pStyle')[0]?.getAttribute('w:val')||'';
const text=p=>Array.from(p.getElementsByTagNameNS(W,'t')).map(x=>x.textContent).join('');
const keep=p=>p.getElementsByTagNameNS(W,'pPr')[0]?.getElementsByTagNameNS(W,'keepNext').length||0;
const raw=()=>Object.fromEntries(FAMILY_BRAND_FIELDS.map(key=>[key,{value:'',invalid:Object.fromEntries(
  ['badInput','customError','patternMismatch','rangeOverflow','rangeUnderflow','stepMismatch','tooLong','tooShort','typeMismatch','valueMissing'].map(flag=>[flag,false]))}]));
function fixture(wrench=false){
 let p=wrench?wrenchCourse('Language and pages 🌱',{audience:'PRIVATE_AUDIENCE',purpose:'PRIVATE_PURPOSE',authors:'PRIVATE_AUTHOR',attribution:'PRIVATE_RIGHTS'},['background']):course('Language and pages 🌱');
 p=addModule(p,'First module','Practice carefully.').next;
 for(const [title,lines,instructions] of [
   ['Zero',0,'Zero response prompt'],['One',1,'Short one'],['Three',3,'Short three'],
   ['Four',4,'Short four'],['Twenty',20,'A long prompt '.repeat(12)+'é 😊']
 ])p=addActivity(p,p.modules[0].id,{title,minutes:5,participant:instructions,facilitator:'PRIVATE FACILITATOR',answers:'PRIVATE ANSWER',lines}).next;
 p=addModule(p,'Second module','Reflect on the work.').next;
 p=addActivity(p,p.modules[1].id,{title:'After lines',minutes:5,participant:'Two lines\nwith language',facilitator:'PRIVATE FACILITATOR',answers:'PRIVATE ANSWER',lines:2}).next;
 p=addActivity(p,p.modules[1].id,{title:'Final reflection',minutes:5,participant:'Write a brief reflection: what did the response show, what remains uncertain, and what would you change?\nA learner may answer in English, español, or 日本語.',facilitator:'PRIVATE FACILITATOR',answers:'PRIVATE ANSWER',lines:3}).next;
 p=addActivity(p,p.modules[1].id,{title:'Long prompt escape',minutes:5,participant:'A deliberately long prompt '.repeat(15),facilitator:'PRIVATE FACILITATOR',answers:'PRIVATE ANSWER',lines:3}).next;
 const groups=p.modules.map(m=>({moduleId:m.id,activityIds:m.activities.map(a=>a.id)}));
 const saved=editDeliveryPlan(p,{kind:'create',name:'Selected pages',targetMinutes:null,groups});
 return {p:saved.next,planId:saved.id};
}

test('bounded short response spaces stay together while 4–20 lines remain breakable in every family and scope',()=>{
 assert.deepEqual(WORKBOOK_BASE_LANGUAGES,['en-US','en-GB','en-CA','es-ES','es-MX','fr-FR','fr-CA','de-DE']);
 for(const wrench of [false,true]){
  const f=fixture(wrench),semantics=[semanticWorkbookFromCourse(f.p),semanticWorkbookFromSelected(f.p,f.planId,stamp)];
  for(const semantic of semantics)for(const family of WORKBOOK_FAMILIES){
   const out=renderWorkbookFamily(semantic,resolveWorkbookBrand(family.id),stamp,{baseLanguage:'fr-FR'});
   const ps=parts(xml(out.files['word/document.xml']));
   for(const [prompt,count,short] of [['Zero response prompt',0,false],['Short one',1,true],['Short three',3,true],['Short four',4,false]]){
    const at=ps.findIndex(p=>style(p)==='Prompt'&&text(p)===prompt);assert.ok(at>=0);
    assert.equal(keep(ps[at])>0,short,`${family.id}/${prompt} prompt association`);
    if(count){assert.equal(style(ps[at+1]),'ResponseLabel');assert.equal(style(ps[at+count+1]),'ResponseLine');
      for(let i=0;i<count;i++)assert.equal(keep(ps[at+2+i])>0,short&&i<count-1,`${family.id}/${prompt} line ${i}`);}
   }
   const long=ps.findIndex(p=>style(p)==='Prompt'&&text(p).startsWith('A long prompt'));
   assert.ok(long>=0);assert.equal(keep(ps[long]),0);
   assert.equal(ps.slice(long+2,long+22).filter(p=>style(p)==='ResponseLine').length,20);
   assert.ok(ps.slice(long+2,long+22).every(p=>!keep(p)));
   const second=ps.find(p=>style(p)==='Heading1'&&text(p).includes('Second module'));
   assert.ok(second);assert.equal(second.getElementsByTagNameNS(W,'pageBreakBefore').length,0);
   const finalPrompt=ps.findIndex(p=>style(p)==='Prompt'&&text(p).startsWith('Write a brief reflection:'));
   assert.ok(finalPrompt>=0);assert.equal(keep(ps[finalPrompt]),1,`${family.id} final multiline prompt stays with its short response`);
   assert.equal(style(ps[finalPrompt+1]),'ResponseLabel');
   assert.equal(ps.slice(finalPrompt+2,finalPrompt+5).filter(p=>style(p)==='ResponseLine').length,3);
   assert.deepEqual(ps.slice(finalPrompt+2,finalPrompt+5).map(keep),[1,1,0],`${family.id} final line ends keep chain`);
   const longPrompt=ps.findIndex(p=>style(p)==='Prompt'&&text(p).startsWith('A deliberately long prompt'));
   assert.ok(longPrompt>=0);assert.equal(keep(ps[longPrompt]),0,`${family.id} long prompt escapes keep chain`);
   assert.deepEqual(ps.slice(longPrompt+2,longPrompt+5).map(keep),[1,1,0]);
   assert.equal(ps.filter(p=>style(p)==='ResponseLine').length,36);
   assert.deepEqual(out.preview.paginationPolicy,{shortResponseMaxLines:3,shortPromptMaxCharacters:240,shortPromptMaxExplicitLines:3,shortPromptMaxEstimatedLines:3,shortActivityHeadingMaxEstimatedLines:2,shortPromptMaxHeightFraction:'1/3',headingToResponseMaxHeightFraction:'1/2',shortPromptWidthEstimateFactor:0.75,shortPromptAutomaticLineSafetyFactor:1.2,emptyObjectiveFirstActivityPromptLink:false,longResponseSplittable:true,forcedModuleBreaks:false});
   assert.ok(!Object.values(out.files).join('').includes('PRIVATE FACILITATOR'));
   assert.equal(out.profileVersion,3);assert.equal(out.familyVersion,3);
  }
 }
});

test('family core title, roman Reflection body and explicit document-default language match preview',()=>{
 const semantic=semanticWorkbookFromCourse(fixture().p);
 for(const family of WORKBOOK_FAMILIES){
  const out=renderWorkbookFamily(semantic,resolveWorkbookBrand(family.id),stamp,{baseLanguage:'es-MX'});
  const styles=xml(out.files['word/styles.xml']),core=xml(out.files['docProps/core.xml']);
  assert.equal(out.preview.baseLanguage,'es-MX');
  const lang=styles.getElementsByTagNameNS(W,'lang');assert.equal(lang.length,1);assert.equal(lang[0].getAttribute('w:val'),'es-MX');
  assert.equal(core.getElementsByTagNameNS(DC,'language')[0].textContent,'es-MX');
  assert.equal(core.getElementsByTagNameNS(DC,'title')[0].textContent,`${semantic.title} ${out.preview.subtitle}`);
  if(family.id==='reflection-journal')for(const name of ['Objective','Prompt']){
   const role=Array.from(styles.getElementsByTagNameNS(W,'style')).find(x=>x.getAttribute('w:styleId')===name);
   assert.equal(role.getElementsByTagNameNS(W,'i').length,0);
  }
 }
});

test('explicit lines, estimated wrapping and large response geometry escape prompt association',()=>{
 const semantic=semanticWorkbookFromCourse(fixture().p);
 const finalNode=semantic.nodes.find(node=>node.kind==='activity'&&node.title==='Final reflection');
 assert.ok(finalNode);
 for(const family of WORKBOOK_FAMILIES){
  for(const instructions of ['one\ntwo\nthree\nfour','word '.repeat(43)]){
   const changed=structuredClone(semantic);
   changed.nodes.find(node=>node.id===finalNode.id).instructions=instructions;
   const out=renderWorkbookFamily(changed,resolveWorkbookBrand(family.id),stamp);
   const prompt=parts(xml(out.files['word/document.xml'])).find(p=>style(p)==='Prompt'&&text(p)===instructions.replaceAll('\n',''));
   assert.ok(prompt);assert.equal(keep(prompt),0,`${family.id} oversized prompt can split`);
  }
  const shortForHeight=structuredClone(semantic);
  shortForHeight.nodes.find(node=>node.id===finalNode.id).instructions='Write one clear next step.';
  const ordinary=renderWorkbookFamily(shortForHeight,resolveWorkbookBrand(family.id),stamp);
  const ordinaryPrompt=parts(xml(ordinary.files['word/document.xml'])).find(p=>style(p)==='Prompt'&&text(p)==='Write one clear next step.');
  assert.ok(ordinaryPrompt);assert.equal(keep(ordinaryPrompt),1,`${family.id} short ordinary group can stay together`);
  const tall=renderWorkbookFamily(shortForHeight,resolveWorkbookBrand(family.id,{responseLineIn:0.7,bodyPt:18,marginIn:1.5,lineSpacing:2}),stamp);
  const prompt=parts(xml(tall.files['word/document.xml'])).find(p=>style(p)==='Prompt'&&text(p)==='Write one clear next step.');
  assert.ok(prompt);assert.equal(keep(prompt),0,`${family.id} one-line prompt with oversized geometry can split`);
 }
});

test('long heading and empty-objective module chain do not acquire another Prompt link',()=>{
 const semantic=semanticWorkbookFromCourse(fixture().p);
 const finalNode=semantic.nodes.find(node=>node.kind==='activity'&&node.title==='Final reflection');
 assert.ok(finalNode);
 for(const family of WORKBOOK_FAMILIES){
  const longHeading=structuredClone(semantic);
  longHeading.nodes.find(node=>node.id===finalNode.id).title='A long activity heading '.repeat(9);
  const longOut=renderWorkbookFamily(longHeading,resolveWorkbookBrand(family.id),stamp);
  const longPrompt=parts(xml(longOut.files['word/document.xml'])).find(p=>style(p)==='Prompt'&&text(p).startsWith('Write a brief reflection:'));
  assert.ok(longPrompt);assert.equal(keep(longPrompt),0,`${family.id} long heading stays breakable`);
  const carried=structuredClone(semantic);
  const module=carried.nodes.find(node=>node.kind==='module'&&node.title==='Second module');
  module.objective='';
  carried.nodes=carried.nodes.filter(node=>node.title!=='After lines');
  const carriedOut=renderWorkbookFamily(carried,resolveWorkbookBrand(family.id),stamp);
  const carriedPrompt=parts(xml(carriedOut.files['word/document.xml'])).find(p=>style(p)==='Prompt'&&text(p).startsWith('Write a brief reflection:'));
  assert.ok(carriedPrompt);assert.equal(keep(carriedPrompt),0,`${family.id} empty-objective module chain stays bounded`);
 }
});

test('auto line spacing uses applied font size when a three-line prompt exceeds the height budget',()=>{
 const semantic=structuredClone(semanticWorkbookFromCourse(fixture().p));
 const final=semantic.nodes.find(node=>node.kind==='activity'&&node.title==='Final reflection');
 assert.ok(final);final.title='Reflect';final.instructions='One.\nTwo.\nThree.';
 const brand=resolveWorkbookBrand('reflection-journal',{pageSize:'Letter',marginIn:0.8,bodyPt:18,lineSpacing:2,responseLineIn:0.25});
 const out=renderWorkbookFamily(semantic,brand,stamp);
 const prompt=parts(xml(out.files['word/document.xml'])).find(p=>style(p)==='Prompt'&&text(p)==='One.Two.Three.');
 assert.ok(prompt);assert.equal(keep(prompt),0);
 assert.equal(out.preview.appliedStyles.Prompt.sizePt,18);
 assert.equal(out.preview.appliedStyles.Prompt.lineTwips,480);
 assert.equal(out.preview.appliedStyles.Prompt.lineRule,'auto');
});

test('base language defaults only by omission; invalid, inherited and accessor options refuse without reads',()=>{
 const f=fixture(),semantic=semanticWorkbookFromCourse(f.p),brand=resolveWorkbookBrand('clearline');
 assert.equal(WORKBOOK_DEFAULT_BASE_LANGUAGE,'en-US');
 assert.equal(renderWorkbookFamily(semantic,brand,stamp).preview.baseLanguage,'en-US');
 for(const options of [undefined,null,{baseLanguage:null},{baseLanguage:'xx-XX'},{baseLanguage:'en-US',extra:true},
   Object.assign(Object.create({baseLanguage:'en-US'}),{})])
  assert.throws(()=>renderWorkbookFamily(semantic,brand,stamp,options),/Unsupported styled workbook/);
 let reads=0;const access={};Object.defineProperty(access,'baseLanguage',{enumerable:true,get(){reads++;return 'en-US';}});
 assert.throws(()=>renderWorkbookFamily(semantic,brand,stamp,access),/Unsupported styled workbook/);assert.equal(reads,0);
 const symbol={baseLanguage:'en-US'};symbol[Symbol('secret')]=1;
 assert.throws(()=>renderWorkbookFamily(semantic,brand,stamp,symbol),/Unsupported styled workbook/);
});

test('family flow binds language to output and review while leaving sixteen preset brand fields unchanged',()=>{
 const f=fixture(),config={planId:'',familyId:'clearline',rawBrand:raw(),stamp};
 const omitted=prepareFamilyWorkbook(f.p,config);
 assert.equal(omitted.output.preview.baseLanguage,'en-US');
 assert.equal(FAMILY_BRAND_FIELDS.length,16);
 const selected=prepareFamilyWorkbook(f.p,{...config,baseLanguage:'de-DE'});
 assert.equal(selected.output.preview.baseLanguage,'de-DE');
 assert.match(familyReviewText(selected),/Base document language: de-DE/);
 assert.match(familyReviewText(selected),/no authored per-run language tagging/);
 assert.match(familyReviewText(selected),/Page-flow policy \(generated profile 3\)/);
 assert.notDeepEqual(selected.output.files,omitted.output.files);
 for(const baseLanguage of [undefined,null,'xx-XX'])assert.throws(()=>prepareFamilyWorkbook(f.p,{...config,baseLanguage}),/base language/);
 let reads=0;const getter={...config};Object.defineProperty(getter,'baseLanguage',{enumerable:true,get(){reads++;return 'en-US';}});
 assert.throws(()=>prepareFamilyWorkbook(f.p,getter),/visible own data/);assert.equal(reads,0);
});
