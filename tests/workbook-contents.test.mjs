import test from 'node:test';
import assert from 'node:assert/strict';
import {createRequire} from 'node:module';
import {course,wrenchCourse,addModule,addActivity} from '../src/model.mjs';
import {editDeliveryPlan} from '../src/delivery-plans.mjs';
import {WORKBOOK_FAMILIES,resolveWorkbookBrand} from '../src/workbook-presentation.mjs';
import {prepareFamilyWorkbook,FAMILY_BRAND_FIELDS,familyWorkbookBlob,familyReviewText} from '../src/workbook-family-flow.mjs';
import {renderWorkbookFamily} from '../src/workbook-family-output.mjs';
import {EMPTY_WORKBOOK_FRONT_MATTER} from '../src/workbook-public-frontmatter.mjs';
import {validateWorkbookContents,workbookContentsRequest,buildWorkbookContents,WORKBOOK_CONTENTS_LIMITS} from '../src/workbook-contents.mjs';
import {rawControl} from '../src/activity-draft.mjs';
const require=createRequire(import.meta.url),Zip=require('jszip'),{DOMParser}=require('@xmldom/xmldom');
const NS='http://schemas.openxmlformats.org/wordprocessingml/2006/main',stamp='2026-10-28T15:00:00.000Z';
const raw=()=>Object.fromEntries(FAMILY_BRAND_FIELDS.map(k=>[k,rawControl({value:k==='organizationName'?'Public Organization':''})]));
function fixture(wrench=false){
 let p=wrench?wrenchCourse('Public course',{audience:'PRIVATE_AUDIENCE',purpose:'PRIVATE_PURPOSE',authors:'PRIVATE_AUTHOR',attribution:'PRIVATE_RIGHTS'},['background']):course('Public course');
 p=addModule(p,'Contents','A learning objective').next;
 p=addActivity(p,p.modules[0].id,{title:'Duplicate activity \u4e2d\u6587 \u{1f642}',minutes:5,participant:'Selected learner instructions',facilitator:'PRIVATE_FACILITATOR',answers:'PRIVATE_ANSWERS',lines:3}).next;
 p=addActivity(p,p.modules[0].id,{title:'Duplicate activity \u4e2d\u6587 \u{1f642}',minutes:5,participant:'OMITTED_LEARNER_CANARY',facilitator:'PRIVATE_OTHER',answers:'PRIVATE_ANSWERS',lines:1}).next;
 p=addModule(p,'Contents',' ').next;p=addActivity(p,p.modules[1].id,{title:'Long activity '+('heading '.repeat(80)),minutes:2,participant:'Other learner text',facilitator:'PRIVATE_TEACHER',lines:0}).next;
 const plan=editDeliveryPlan(p,{kind:'create',name:'Public plan',targetMinutes:5,groups:[{moduleId:p.modules[0].id,activityIds:[p.modules[0].activities[0].id]}]});return {p:plan.next,id:plan.id};
}
const front=enabled=>({...structuredClone(EMPTY_WORKBOOK_FRONT_MATTER),includeCopyright:enabled,copyright:{...structuredClone(EMPTY_WORKBOOK_FRONT_MATTER.copyright),rightsHolder:'Public holder',year:'2026',noticeChoice:'custom-text',noticeText:'Public authored notice'}});
const cfg=(familyId,planId='',extra={})=>({familyId,planId,rawBrand:raw(),stamp,baseLanguage:'en-US',...extra});
function parsed(files){const errors=[],parser=new DOMParser({errorHandler:{warning:e=>errors.push(e),error:e=>errors.push(e),fatalError:e=>errors.push(e)}});const docs=Object.fromEntries(Object.entries(files).map(([name,xml])=>[name,parser.parseFromString(xml,'application/xml')]));assert.deepEqual(errors,[]);return docs;}
function text(node){let out='';const walk=n=>{if(n.nodeType===1&&n.namespaceURI===NS&&n.localName==='t')out+=n.textContent;else if(n.nodeType===1&&n.namespaceURI===NS&&n.localName==='br')out+='\n';else if(n.nodeType===1&&n.namespaceURI===NS&&n.localName==='tab')out+='\t';else for(const c of Array.from(n.childNodes||[]))walk(c);};walk(node);return out;}

test('contents config/raw/heading inputs are strict own data and defensive, with no inferred pages or identity leaks',()=>{
 const source={schemaVersion:1,include:true};assert.deepEqual(validateWorkbookContents(source),source);assert.equal(Object.isFrozen(source),false);
 const c=workbookContentsRequest({checked:true,...rawControl({value:'on'})});assert.equal(c.include,true);
 let hits=0;const getter={schemaVersion:1};Object.defineProperty(getter,'include',{enumerable:true,get(){hits++;return true;}});
 for(const bad of [null,undefined,{},getter,{...source,schemaVersion:2},{...source,include:1},{...source,extra:true},Object.assign(Object.create(source),source),{...source,[Symbol('x')]:1}])assert.throws(()=>validateWorkbookContents(bad));assert.equal(hits,0);
 const rowGetter={key:'stable',text:'Label'};Object.defineProperty(rowGetter,'level',{enumerable:true,get(){hits++;return 1;}});assert.throws(()=>buildWorkbookContents([rowGetter]));const arrayGetter=[];Object.defineProperty(arrayGetter,'0',{enumerable:true,get(){hits++;return {key:'x',level:1,text:'X'};}});assert.throws(()=>buildWorkbookContents(arrayGetter));const customIterator=[{key:'x',level:1,text:'X'}];customIterator[Symbol.iterator]=function(){hits++;return [][Symbol.iterator]();};assert.throws(()=>buildWorkbookContents(customIterator));assert.equal(hits,0);
 const badRaw={checked:false,...rawControl({value:'on'})};badRaw.invalid.customError=true;assert.throws(()=>workbookContentsRequest(badRaw));
 for(const bad of [null,{},[{key:'x',level:1,text:'ok',anchor:'forged'}],[{key:'x',level:3,text:'ok'}],[{key:'x',level:1,text:'\u0001'}],[{key:'x',level:1,text:'bad\ud800'}],[{key:'x',level:1,text:'one'},{key:'x',level:2,text:'two'}]])assert.throws(()=>buildWorkbookContents(bad));
 const rows=[{key:'module:one',level:1,text:'Duplicate'},{key:'module:two',level:1,text:'Duplicate'},{key:'activity:one',level:2,text:'Long\r\nlabel\t\u4e2d\u6587 \u{1f642}'}],before=JSON.stringify(rows),a=buildWorkbookContents(rows),b=buildWorkbookContents(rows);
 assert.equal(JSON.stringify(rows),before);assert.equal(Object.isFrozen(rows),false);assert.deepEqual(a,b);assert.equal(a.pageNumbers,false);assert.equal(Object.isFrozen(a.entries[0]),true);assert.equal(a.entries[2].text,'Long\nlabel\t\u4e2d\u6587 \u{1f642}');assert.ok(!JSON.stringify(a).includes('module:one'));
 assert.equal(new Set(a.entries.map(e=>e.anchor)).size,3);assert.deepEqual(a.entries.map(e=>e.bookmarkId),[1,2,3]);for(const e of a.entries)assert.match(e.anchor,/^[A-Za-z][A-Za-z0-9_]{0,39}$/u);
 assert.deepEqual(buildWorkbookContents([]).entries,[]);assert.throws(()=>buildWorkbookContents(rows,{}));
});

test('heading budget supports maximum semantic heading count without counting repeated identity against published labels',()=>{
 const count=WORKBOOK_CONTENTS_LIMITS.headings,rows=Array.from({length:count},(_,i)=>({key:('K'.repeat(180))+i,level:i%2?2:1,text:i%2?'Learning objective':'Module heading '+i}));
 const result=buildWorkbookContents(rows);assert.equal(result.entries.length,count);assert.equal(new Set(result.entries.map(e=>e.anchor)).size,count);assert.ok(WORKBOOK_CONTENTS_LIMITS.identityTextUnits>rows.reduce((n,r)=>n+r.key.length,0));
 assert.throws(()=>buildWorkbookContents([...rows,{key:'extra',level:1,text:'extra'}]));
});

test('valid near-limit2000-module semantic source with family prefixes produces all4000 contents entries',()=>{
 const semantic={version:1,projectId:'P',revision:0,scope:{kind:'full',planId:null,label:'Full'},title:'C',identityLines:[],nodes:Array.from({length:2000},(_,i)=>({kind:'module',id:'m'+i,title:'x'.repeat(119),objective:'o'}))};
 const units=semantic.projectId.length+semantic.scope.label.length+semantic.title.length+semantic.nodes.reduce((n,m)=>n+m.id.length+m.title.length+m.objective.length,0);assert.equal(units,248896);
 const output=renderWorkbookFamily(semantic,resolveWorkbookBrand('practice-lab'),stamp,{baseLanguage:'en-US',toc:{schemaVersion:1,include:true}});assert.equal(output.preview.toc.entries.length,4000);assert.ok(output.preview.toc.entries.reduce((n,e)=>n+e.text.length,0)>300000);assert.ok(output.preview.toc.entries.reduce((n,e)=>n+e.text.length,0)<=WORKBOOK_CONTENTS_LIMITS.labelTextUnits);assert.equal(new Set(output.preview.toc.entries.map(e=>e.anchor)).size,4000);
});

test('all five families full/selected Minimal/Wrench optional copyright create exact heading links and preserve disabled bytes',async()=>{
 for(const wrench of [false,true]){
  const {p,id}=fixture(wrench),before=JSON.stringify(p);let fullAnchors=null,selectedAnchors=null;
  for(const family of WORKBOOK_FAMILIES)for(const planId of ['',id])for(const copyright of [false,true]){
   const base=prepareFamilyWorkbook(p,cfg(family.id,planId,{frontMatter:front(copyright)})),off=prepareFamilyWorkbook(p,cfg(family.id,planId,{frontMatter:front(copyright),toc:{schemaVersion:1,include:false}}));
   assert.deepEqual(off.output,base.output);
   const prepared=prepareFamilyWorkbook(p,cfg(family.id,planId,{frontMatter:front(copyright),toc:{schemaVersion:1,include:true}}));assert.equal(prepared.eligible,true);const {files,preview}=prepared.output;
   for(const name of Object.keys(files).filter(n=>!['word/document.xml','word/styles.xml'].includes(n)))assert.equal(files[name],base.output.files[name],name);
   const docs=parsed(files),body=docs['word/document.xml'],paras=Array.from(body.getElementsByTagNameNS(NS,'p'));
   const headingParas=paras.filter(n=>{const st=n.getElementsByTagNameNS(NS,'pStyle')[0]?.getAttributeNS(NS,'val');return st==='Heading1'||st==='Heading2';});
   const links=Array.from(body.getElementsByTagNameNS(NS,'hyperlink')),starts=Array.from(body.getElementsByTagNameNS(NS,'bookmarkStart')),ends=Array.from(body.getElementsByTagNameNS(NS,'bookmarkEnd'));
   assert.deepEqual(links.map(text),headingParas.map(text));assert.deepEqual(preview.toc.entries.map(e=>e.text),headingParas.map(text));assert.equal(links.length,starts.length);assert.equal(starts.length,ends.length);
   assert.equal(new Set(starts.map(n=>n.getAttributeNS(NS,'name'))).size,starts.length);assert.equal(new Set(starts.map(n=>n.getAttributeNS(NS,'id'))).size,starts.length);
   for(const [i,entry]of preview.toc.entries.entries()){
    assert.equal(links[i].getAttributeNS(NS,'anchor'),entry.anchor);assert.equal(starts[i].getAttributeNS(NS,'name'),entry.anchor);assert.equal(starts[i].getAttributeNS(NS,'id'),ends[i].getAttributeNS(NS,'id'));assert.equal(text(starts[i].parentNode),entry.text);assert.match(entry.anchor,/^[A-Za-z][A-Za-z0-9_]{0,39}$/u);
   }
   assert.equal(preview.toc.entries.filter(e=>e.text==='Learning objective').length,prepared.semantic.nodes.filter(n=>n.kind==='module'&&n.objective).length);assert.equal(preview.toc.entries.filter(e=>e.text==='Copyright').length,copyright?1:0);assert.ok(preview.toc.entries.some(e=>e.text==='Contents'||e.text.endsWith(' · Contents')));
   const all=Object.values(files).join('');assert.doesNotMatch(all,/PRIVATE_/u);if(planId)assert.doesNotMatch(all,/OMITTED_LEARNER_CANARY/u);assert.doesNotMatch(files['word/document.xml'],/w:(?:fldChar|instrText|fldSimple)|PAGEREF/u);assert.doesNotMatch(files['word/settings.xml'],/updateFields/u);assert.doesNotMatch(files['word/_rels/document.xml.rels'],/hyperlink|TargetMode/u);
   const docStyles=docs['word/styles.xml'];for(const styleId of ['ContentsTitle','ContentsLevel1','ContentsLevel2']){const style=Array.from(docStyles.getElementsByTagNameNS(NS,'style')).find(n=>n.getAttributeNS(NS,'styleId')===styleId);assert.ok(style);const applied=preview.appliedStyles[styleId];assert.ok(applied);assert.equal(style.getElementsByTagNameNS(NS,'rFonts')[0].getAttributeNS(NS,'ascii'),applied.font);assert.equal(Number(style.getElementsByTagNameNS(NS,'sz')[0].getAttributeNS(NS,'val'))/2,applied.sizePt);if(styleId!=='ContentsTitle')assert.equal(style.getElementsByTagNameNS(NS,'outlineLvl').length,0);}
   for(const link of links)for(const run of Array.from(link.getElementsByTagNameNS(NS,'r')).filter(r=>r.getElementsByTagNameNS(NS,'t').length)){assert.equal(run.getElementsByTagNameNS(NS,'color')[0].getAttributeNS(NS,'val'),preview.renderedColors.text.slice(1).toUpperCase());assert.equal(run.getElementsByTagNameNS(NS,'u')[0].getAttributeNS(NS,'val'),'single');}
   const currentAnchors=preview.toc.entries.map(e=>e.anchor);if(!copyright){if(planId){if(selectedAnchors)assert.deepEqual(currentAnchors,selectedAnchors);else selectedAnchors=currentAnchors;}else{if(fullAnchors)assert.deepEqual(currentAnchors,fullAnchors);else fullAnchors=currentAnchors;}}
   const blob=await familyWorkbookBlob(files,Zip),zip=await Zip.loadAsync(await blob.arrayBuffer(),{checkCRC32:true});assert.equal(Object.keys(zip.files).length,7);assert.equal(await zip.file('word/document.xml').async('string'),files['word/document.xml']);assert.match(familyReviewText(prepared),/snapshot|manual edits/u);assert.doesNotMatch(familyReviewText(prepared),/internal target OSG_/u);
  }
  assert.equal(JSON.stringify(p),before);
 }
});

test('contents cover order, empty roster, repeated generation and malformed options preserve controls and metadata contracts',()=>{
 const {p}=fixture(),fm=front(true);fm.includeCover=true;fm.cover.title='Explicit cover';fm.cover.author='Explicit author';const c=cfg('clearline','',{frontMatter:fm,toc:{schemaVersion:1,include:true}}),a=prepareFamilyWorkbook(p,c),b=prepareFamilyWorkbook(p,c);assert.deepEqual(a,b);
 const xml=a.output.files['word/document.xml'];assert.ok(xml.indexOf('Explicit cover')<xml.indexOf('Copyright'));assert.ok(xml.indexOf('Copyright')<xml.indexOf('ContentsTitle'));assert.match(a.output.files['docProps/core.xml'],/<dc:creator>Explicit author/u);
 const brand=resolveWorkbookBrand('clearline'),semantic={version:1,projectId:'P',revision:0,scope:{kind:'full',planId:null,label:'Full course'},title:'Empty course',identityLines:[],nodes:[]};
 const empty=renderWorkbookFamily(semantic,brand,stamp,{baseLanguage:'en-US',toc:{schemaVersion:1,include:true}});assert.deepEqual(empty.preview.toc.entries,[]);assert.match(empty.files['word/document.xml'],/No generated content headings/u);assert.doesNotMatch(empty.files['word/document.xml'],/w:hyperlink|w:bookmarkStart/u);
 const whitespace=renderWorkbookFamily({...semantic,nodes:[{kind:'module',id:'M',title:'Module',objective:' '}]},brand,stamp,{baseLanguage:'en-US',toc:{schemaVersion:1,include:true}});assert.equal(whitespace.preview.toc.entries.filter(e=>e.text==='Learning objective').length,1);
 for(const toc of [null,undefined,{schemaVersion:2,include:true},{schemaVersion:1,include:'true'},{schemaVersion:1,include:true,pageNumbers:true}])assert.throws(()=>prepareFamilyWorkbook(p,{...cfg('clearline'),toc}));
});
