import test from 'node:test';
import assert from 'node:assert/strict';
import {createRequire} from 'node:module';
import {course,wrenchCourse,addModule,addActivity} from '../src/model.mjs';
import {editDeliveryPlan} from '../src/delivery-plans.mjs';
import {WORKBOOK_FAMILIES,resolveWorkbookBrand} from '../src/workbook-presentation.mjs';
import {prepareFamilyWorkbook,familyReviewText,FAMILY_BRAND_FIELDS,familyWorkbookBlob} from '../src/workbook-family-flow.mjs';
import {renderWorkbookFamily} from '../src/workbook-family-output.mjs';
import {rawControl} from '../src/activity-draft.mjs';
import {EMPTY_WORKBOOK_FRONT_MATTER} from '../src/workbook-public-frontmatter.mjs';
import {WORKBOOK_CONTENTS_LIMITS} from '../src/workbook-contents.mjs';
import {WORKBOOK_RESEARCH_FIELDS,EMPTY_WORKBOOK_RESEARCH,validateWorkbookResearch,workbookResearchRequest,workbookResearchRows} from '../src/workbook-research.mjs';
const require=createRequire(import.meta.url),Zip=require('jszip'),{DOMParser}=require('@xmldom/xmldom');
const NS='http://schemas.openxmlformats.org/wordprocessingml/2006/main',stamp='2026-10-28T15:40:00.000Z';
const research=()=>({schemaVersion:1,include:true,rationale:'Authored educational rationale < & >\r\nSecond line\t\u4e2d\u6587 \u{1f642}',evidence:'Authored evidence with (Example, 2026).',application:'Apply the authored rationale to this workshop.',limitations:'Authored limitations, not measured validation.',references:'Example, A. (2026). Authored reference.\nLiteral \\r\\n stays literal.'});
const raw=r=>({schemaVersion:1,include:{checked:r.include,...rawControl({value:'on'})},...Object.fromEntries(WORKBOOK_RESEARCH_FIELDS.map(k=>[k,rawControl({value:r[k]})]))});
const brand=()=>Object.fromEntries(FAMILY_BRAND_FIELDS.map(k=>[k,rawControl({value:''})]));
function fixture(wrench=false){
 let p=wrench?wrenchCourse('Learner course',{audience:'PRIVATE_AUDIENCE',purpose:'PRIVATE_PURPOSE',authors:'PRIVATE_AUTHORS',attribution:'PRIVATE_RIGHTS'},['background']):course('Learner course');
 p=addModule(p,'References','Learning objective').next;p=addActivity(p,p.modules[0].id,{title:'Research evidence',minutes:5,participant:'Public selected learner',facilitator:'PRIVATE_FACILITATOR',answers:'PRIVATE_ANSWER',lines:3}).next;
 p=addModule(p,'Other module','Other objective').next;p=addActivity(p,p.modules[1].id,{title:'Omitted activity',minutes:7,participant:'OMITTED_LEARNER',facilitator:'PRIVATE_OTHER',answers:'PRIVATE_OTHER_ANSWER',lines:1}).next;
 const plan=editDeliveryPlan(p,{kind:'create',name:'Selected plan',targetMinutes:5,groups:[{moduleId:p.modules[0].id,activityIds:[p.modules[0].activities[0].id]}]});return {p:plan.next,id:plan.id};
}
const front=(cover,copyright)=>({...structuredClone(EMPTY_WORKBOOK_FRONT_MATTER),includeCover:cover,includeCopyright:copyright,cover:{...EMPTY_WORKBOOK_FRONT_MATTER.cover,title:'Explicit cover',author:'Explicit author'},copyright:{...EMPTY_WORKBOOK_FRONT_MATTER.copyright,rightsHolder:'Explicit rights holder',year:'2026',noticeChoice:'custom-text',noticeText:'Authored copyright wording'}});
const cfg=(familyId,planId='',extra={})=>({familyId,planId,rawBrand:brand(),stamp,baseLanguage:'en-US',...extra});
const parse=files=>Object.fromEntries(Object.entries(files).map(([name,xml])=>{const errors=[],d=new DOMParser({errorHandler:{warning:x=>errors.push(x),error:x=>errors.push(x),fatalError:x=>errors.push(x)}}).parseFromString(xml,'application/xml');assert.deepEqual(errors,[]);return [name,d];}));
const text=node=>Array.from(node.childNodes||[]).map(n=>n.nodeType===3?n.nodeValue:n.localName==='br'?'\n':n.localName==='tab'?'\t':text(n)).join('');

test('research is explicit defensive own-data text, active nonblank and public projection without generated claims',()=>{
 const r=research(),before=JSON.stringify(r),v=validateWorkbookResearch(r);assert.equal(JSON.stringify(r),before);assert.equal(Object.isFrozen(r),false);assert.equal(Object.isFrozen(v),true);assert.equal(v.rationale,'Authored educational rationale < & >\nSecond line\t\u4e2d\u6587 \u{1f642}');assert.match(v.references,/Literal \\r\\n/u);
 const rows=workbookResearchRows(v);assert.equal(rows.sections.length,5);assert.equal(rows.heading,'Educational research and rationale');assert.deepEqual(rows.sections.map(x=>x.key),WORKBOOK_RESEARCH_FIELDS);assert.equal(Object.isFrozen(rows.sections[0]),true);r.rationale='CHANGED';assert.notEqual(v.rationale,r.rationale);
 assert.deepEqual(workbookResearchRows({...research(),include:false}).sections,[]);assert.throws(()=>validateWorkbookResearch({...EMPTY_WORKBOOK_RESEARCH,include:true}));
 const blanks={...research(),evidence:'  ',references:'\t'};assert.deepEqual(workbookResearchRows(blanks).sections.map(x=>x.key),['rationale','application','limitations']);
 assert.throws(()=>validateWorkbookResearch({...research(),rationale:'x'.repeat(16001)}));assert.throws(()=>validateWorkbookResearch({...research(),...Object.fromEntries(WORKBOOK_RESEARCH_FIELDS.map(k=>[k,'x'.repeat(16000)]))}));
});

test('research versions descriptors malformed XML and raw validity safely refuse with zero getter reads',()=>{
 let hits=0;const good=research(),getter={...good};Object.defineProperty(getter,'evidence',{enumerable:true,get(){hits++;return 'unexpected';}});
 const hidden={...good};Object.defineProperty(hidden,'rationale',{value:'hidden',enumerable:false});
 for(const bad of [null,undefined,{},[],{...good,schemaVersion:2},{...good,include:1},{...good,private:'extra'},Object.assign(Object.create(good),good),getter,hidden,{...good,[Symbol('extra')]:1},{...good,rationale:'bad\u0001'},{...good,references:'bad\ud800'}])assert.throws(()=>validateWorkbookResearch(bad));assert.equal(hits,0);
 assert.throws(()=>validateWorkbookResearch(good,{}));assert.throws(()=>workbookResearchRows(good,{}));
 const active=raw(good);active.evidence.invalid.badInput=true;assert.throws(()=>workbookResearchRequest(active));active.include.checked=false;assert.equal(workbookResearchRequest(active).evidence,good.evidence);
 active.include.invalid.customError=true;assert.throws(()=>workbookResearchRequest(active));const rawGetter=raw(good);Object.defineProperty(rawGetter.references,'value',{enumerable:true,get(){hits++;return 'bad';}});assert.throws(()=>workbookResearchRequest(rawGetter));assert.equal(hits,0);
});

test('all families full-selected Minimal-Wrench publication and contents combinations preserve off bytes and scope',async()=>{
 for(const wrench of [false,true]){const {p,id}=fixture(wrench),before=JSON.stringify(p);
  for(const family of WORKBOOK_FAMILIES)for(const planId of ['',id])for(const cover of [false,true])for(const copyright of [false,true])for(const toc of [false,true]){
   const base=prepareFamilyWorkbook(p,cfg(family.id,planId,{frontMatter:front(cover,copyright),toc:{schemaVersion:1,include:toc}}));
   const off=prepareFamilyWorkbook(p,cfg(family.id,planId,{frontMatter:front(cover,copyright),toc:{schemaVersion:1,include:toc},research:{...research(),include:false}}));assert.deepEqual(off.output,base.output);
   const prepared=prepareFamilyWorkbook(p,cfg(family.id,planId,{frontMatter:front(cover,copyright),toc:{schemaVersion:1,include:toc},research:research()})),{files,preview}=prepared.output;
   for(const part of Object.keys(files).filter(k=>k!=='word/document.xml'))assert.equal(files[part],base.output.files[part],part);assert.equal(preview.research.sections.length,5);
   const docs=parse(files),body=docs['word/document.xml'],paras=Array.from(body.getElementsByTagNameNS(NS,'p')),researchHeading=paras.find(x=>text(x)==='Educational research and rationale'&&x.getElementsByTagNameNS(NS,'pStyle')[0]?.getAttributeNS(NS,'val')==='Heading1');assert.ok(researchHeading);assert.equal(researchHeading.getElementsByTagNameNS(NS,'pStyle')[0].getAttributeNS(NS,'val'),'Heading1');
   assert.ok(files['word/document.xml'].indexOf('Educational research and rationale')<files['word/document.xml'].indexOf('Project ID:'));
   const all=Object.values(files).join('');assert.doesNotMatch(all,/PRIVATE_/u);if(planId)assert.doesNotMatch(all,/OMITTED_LEARNER/u);assert.doesNotMatch(all,/Claims, references and applicability are authored|have not been verified by OpenStage/u);
   assert.match(files['word/document.xml'],/w:tab/u);assert.match(files['word/document.xml'],/w:br/u);assert.ok(text(body).includes('Literal \\r\\n stays literal.'));assert.ok(text(body).includes('\u4e2d\u6587 \u{1f642}'));
   if(toc){const entries=preview.toc.entries;assert.equal(entries.filter(e=>e.text==='Educational research and rationale').length,1);assert.equal(entries.filter(e=>e.text==='References'||e.text.endsWith(' \u00b7 References')).length,2);assert.equal(entries.filter(e=>e.text==='Research evidence'||e.text.endsWith(' \u00b7 Research evidence')).length,2);const starts=Array.from(body.getElementsByTagNameNS(NS,'bookmarkStart')),links=Array.from(body.getElementsByTagNameNS(NS,'hyperlink'));assert.equal(starts.length,entries.length);assert.deepEqual(links.map(text),entries.map(e=>e.text));assert.equal(new Set(entries.map(e=>e.anchor)).size,entries.length);}
   const blob=await familyWorkbookBlob(files,Zip),zip=await Zip.loadAsync(await blob.arrayBuffer(),{checkCRC32:true});assert.equal(Object.keys(zip.files).length,7);assert.equal(await zip.file('word/document.xml').async('string'),files['word/document.xml']);assert.match(familyReviewText(prepared),/authored content|have not been verified/u);
  }
  assert.equal(JSON.stringify(p),before);
 }
});

test('one authored subsection is sufficient and empty/unsupported research never invents a source or citation',()=>{
 const {p}=fixture(),r={...EMPTY_WORKBOOK_RESEARCH,include:true,references:'Authored reference only.'},a=prepareFamilyWorkbook(p,cfg('clearline','',{research:r})),b=prepareFamilyWorkbook(p,cfg('clearline','',{research:r}));assert.deepEqual(a,b);assert.deepEqual(a.output.preview.research.sections.map(s=>s.key),['references']);assert.match(a.output.files['word/document.xml'],/Authored reference only/u);assert.doesNotMatch(a.output.files['word/document.xml'].slice(0,a.output.files['word/document.xml'].indexOf('<w:pStyle w:val="Scope"')),/Educational rationale<|Research evidence<|Limitations</u);
 for(const bad of [null,undefined,{...r,schemaVersion:2},{...r,evidence:undefined},{...r,include:'true'},{...r,autoCitations:true}])assert.throws(()=>prepareFamilyWorkbook(p,cfg('clearline','',{research:bad})));
});

test('maximum semantic heading roster plus copyright and research fits4007 real linked headings',()=>{
 const semantic={version:1,projectId:'P',revision:0,scope:{kind:'full',planId:null,label:'Full'},title:'C',identityLines:[],nodes:Array.from({length:2000},(_,i)=>({kind:'module',id:'m'+i,title:'x'.repeat(119),objective:'o'}))};
 const output=renderWorkbookFamily(semantic,resolveWorkbookBrand('practice-lab'),stamp,{baseLanguage:'en-US',frontMatter:front(false,true),toc:{schemaVersion:1,include:true},research:research()});assert.equal(WORKBOOK_CONTENTS_LIMITS.headings,4007);assert.equal(output.preview.toc.entries.length,4007);assert.equal(new Set(output.preview.toc.entries.map(e=>e.anchor)).size,4007);assert.equal(output.preview.toc.entries.filter(e=>e.text==='Educational research and rationale').length,1);
});
