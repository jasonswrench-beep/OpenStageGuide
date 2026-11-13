import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import {DOMParser} from '@xmldom/xmldom';
import {course,wrenchCourse,addModule,addActivity} from '../src/model.mjs';
import {editDeliveryPlan} from '../src/delivery-plans.mjs';
import {worksheetFiles} from '../src/worksheet.mjs';
import {planWorkbookFiles} from '../src/selected-plan-outputs.mjs';
import {WORKBOOK_FAMILIES,resolveWorkbookBrand,semanticWorkbookFromCourse,semanticWorkbookFromSelected} from '../src/workbook-presentation.mjs';
import {FAMILY_WORKBOOK_PARTS,FAMILY_WORKBOOK_PROFILE_VERSION,FAMILY_WORKBOOK_SETTINGS_XML,
 FAMILY_WORKBOOK_CONTENT_TYPES_XML,FAMILY_WORKBOOK_DOCUMENT_RELS_XML,renderWorkbookFamily} from '../src/workbook-family-output.mjs';

const stamp='2026-10-26T10:00:00.000Z';
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
function fixture(wrench=false){
 let p=wrench?wrenchCourse('Learning 🌱',{audience:'PRIVATE_AUDIENCE',purpose:'PRIVATE_PURPOSE',authors:'PRIVATE_AUTHOR',attribution:'PRIVATE_RIGHTS'},['background']):course('Learning 🌱');
 for(let i=0;i<3;i++)p=addModule(p,`Module ${i}`,`Objective ${i}`).next;
 const add=(mi,title,lines)=>{p=addActivity(p,p.modules[mi].id,{title,minutes:10,participant:`Learner ${title}\nSecond line 😊`,facilitator:`PRIVATE_TEACH_${title}`,answers:`PRIVATE_KEY_${title}`,lines}).next;};
 add(0,'OMITTED',0);add(0,'A',1);add(1,'B',20);add(1,'C',0);
 const groups=[{moduleId:p.modules[2].id,activityIds:[]},{moduleId:p.modules[1].id,activityIds:[p.modules[1].activities[1].id,p.modules[1].activities[0].id]},{moduleId:p.modules[0].id,activityIds:[p.modules[0].activities[1].id]}];
 const created=editDeliveryPlan(p,{kind:'create',name:'Reviewed',targetMinutes:30,groups});p=created.next;
 p.opaque={secret:'PRIVATE_OPAQUE'};return {p,planId:created.id};
}
const docs=p=>({full:semanticWorkbookFromCourse(p.p),selected:semanticWorkbookFromSelected(p.p,p.planId,stamp)});
const xmlDoc=value=>{
 const errors=[];const parsed=new DOMParser({errorHandler:{warning:e=>errors.push(e),error:e=>errors.push(e),fatalError:e=>errors.push(e)}}).parseFromString(value,'application/xml');
 assert.deepEqual(errors,[]);assert.ok(parsed.documentElement);return parsed;
};
const paragraphs=doc=>Array.from(doc.getElementsByTagNameNS(W,'p'));
const pStyle=p=>p.getElementsByTagNameNS(W,'pStyle')[0]?.getAttribute('w:val')||'';
const pText=p=>Array.from(p.getElementsByTagNameNS(W,'t')).map(x=>x.textContent).join('');
const text=doc=>paragraphs(doc).map(pText).join('\n');
const pack=async files=>{const zip=new JSZip();for(const [name,data] of Object.entries(files))zip.file(name,data,{createFolders:false});return JSZip.loadAsync(await zip.generateAsync({type:'nodebuffer',compression:'DEFLATE'}));};

test('five original layouts generate strict seven-part DOCX in full and selected Minimal/Wrench scopes',async()=>{
 assert.equal(FAMILY_WORKBOOK_PROFILE_VERSION,3);
 const signatures=new Map(),all=[];
 for(const wrench of [false,true]){
  const f=fixture(wrench),before=structuredClone(f.p),semantics=docs(f);
  for(const [scope,semantic] of Object.entries(semantics))for(const family of WORKBOOK_FAMILIES){
   const brand=resolveWorkbookBrand(family.id,{organizationName:'Example College'}),out=renderWorkbookFamily(semantic,brand,stamp);
   assert.deepEqual(Object.keys(out.files),FAMILY_WORKBOOK_PARTS);
   assert.equal(out.profileVersion,3);assert.equal(out.familyId,family.id);assert.equal(out.familyVersion,3);assert.equal(out.preview.scope.kind,scope);
   const doc=xmlDoc(out.files['word/document.xml']),styles=xmlDoc(out.files['word/styles.xml']);
   assert.equal(out.files['word/settings.xml'],FAMILY_WORKBOOK_SETTINGS_XML);
   assert.equal(out.files['[Content_Types].xml'],FAMILY_WORKBOOK_CONTENT_TYPES_XML);
   assert.equal(out.files['word/_rels/document.xml.rels'],FAMILY_WORKBOOK_DOCUMENT_RELS_XML);
   const settings=xmlDoc(out.files['word/settings.xml']);
   const compat=settings.getElementsByTagNameNS(W,'compatSetting');assert.equal(compat.length,1);
   assert.equal(compat[0].getAttribute('w:name'),'compatibilityMode');
   assert.equal(compat[0].getAttribute('w:uri'),'http://schemas.microsoft.com/office/word');
   assert.equal(compat[0].getAttribute('w:val'),'15');
   for(const part of FAMILY_WORKBOOK_PARTS)if(part.endsWith('.xml')||part.endsWith('.rels'))xmlDoc(out.files[part]);
   const zip=await pack(out.files);assert.deepEqual(Object.keys(zip.files).sort(),[...FAMILY_WORKBOOK_PARTS].sort());
   assert.equal(doc.documentElement.localName,'document');assert.equal(styles.documentElement.localName,'styles');
   const body=text(doc),entire=Object.values(out.files).join('\n');
   for(const excluded of ['PRIVATE_TEACH','PRIVATE_KEY','PRIVATE_AUDIENCE','PRIVATE_PURPOSE','PRIVATE_AUTHOR','PRIVATE_RIGHTS','PRIVATE_OPAQUE'])assert.ok(!entire.includes(excluded),`${family.id}/${scope}/${excluded}`);
   if(scope==='selected'){assert.ok(!entire.includes('OMITTED'));assert.ok(!entire.includes('Objective 2'));assert.ok(body.includes(f.planId));assert.match(body,/does not confirm objective coverage/);}
   const activity=out.preview.sections.flatMap(s=>s.activities);
   assert.deepEqual(activity.map(a=>a.responseLines),scope==='selected'?[0,20,1]:[0,1,20,0]);
   assert.equal(paragraphs(doc).filter(p=>pStyle(p)==='ResponseLine').length,activity.reduce((n,a)=>n+a.responseLines,0));
   assert.equal(paragraphs(doc).filter(p=>pStyle(p)==='Title').length,1);
   assert.equal(paragraphs(doc).filter(p=>pStyle(p)==='Heading1').length,scope==='selected'?2:3);
   assert.equal(paragraphs(doc).filter(p=>pStyle(p)==='ResponseLabel').length,activity.filter(a=>a.responseLines>0).length);
   for(const section of out.preview.sections){assert.ok(body.includes(section.title));if(section.objective)assert.ok(body.includes(section.objective));for(const a of section.activities){assert.ok(body.includes(a.title));assert.ok(body.includes(a.instructions.split('\n')[0]));}}
   if(!wrench&&scope==='full')signatures.set(family.id,out.files['word/document.xml']+'\n'+out.files['word/styles.xml']);
   all.push(out);
  }
  assert.deepEqual(f.p,before);assert.equal(Object.isFrozen(f.p),false);
 }
 assert.equal(all.length,20);assert.equal(new Set(signatures.values()).size,5,'families must differ beyond palette');
});

test('all 25 ordered family switches and same-family reapply are deterministic with no accumulated cover or response lines',()=>{
 const f=fixture(),semantic=docs(f).selected,outputs=new Map();
 for(const family of WORKBOOK_FAMILIES){
  const brand=resolveWorkbookBrand(family.id,{organizationName:'College',accent:'#246A82'});
  outputs.set(family.id,renderWorkbookFamily(semantic,brand,stamp));
 }
 for(const from of WORKBOOK_FAMILIES)for(const to of WORKBOOK_FAMILIES){
  const first=outputs.get(from.id),middle=outputs.get(to.id),last=renderWorkbookFamily(semantic,resolveWorkbookBrand(from.id,{organizationName:'College',accent:'#246A82'}),stamp);
  assert.deepEqual(last.files,first.files);assert.deepEqual(last.preview,first.preview);
  assert.equal(xmlDoc(middle.files['word/document.xml']).getElementsByTagNameNS(W,'body').length,1);
  assert.equal(paragraphs(xmlDoc(last.files['word/document.xml'])).filter(p=>pStyle(p)==='Title').length,1);
  assert.equal(paragraphs(xmlDoc(last.files['word/document.xml'])).filter(p=>pStyle(p)==='ResponseLine').length,21);
 }
});

test('styles carry named heading levels, body typography, actual contrast pairs, page geometry and family-specific response treatment',()=>{
 const semantic=docs(fixture()).full;
 for(const family of WORKBOOK_FAMILIES){
  const brand=resolveWorkbookBrand(family.id,{organizationName:'École 日本語',text:'#202020',background:'#FFFFFF',accent:'#4C5678',accentText:'#FFFFFF',accentBackground:'#202020',bodyFont:'Arial',headingFont:'Georgia',pageSize:'A4',marginIn:0.9,responseLineIn:0.4});
  const out=renderWorkbookFamily(semantic,brand,stamp),styles=xmlDoc(out.files['word/styles.xml']),doc=xmlDoc(out.files['word/document.xml']);
  const styleNodes=Array.from(styles.getElementsByTagNameNS(W,'style')),ids=styleNodes.map(x=>x.getAttribute('w:styleId'));
  for(const id of ['Normal','Title','Subtitle','Heading1','Heading2','Objective','Prompt','ResponseLabel','ResponseLine'])assert.ok(ids.includes(id),`${family.id}/${id}`);
  const h1=styleNodes.find(x=>x.getAttribute('w:styleId')==='Heading1'),h2=styleNodes.find(x=>x.getAttribute('w:styleId')==='Heading2');
  assert.equal(h1.getElementsByTagNameNS(W,'outlineLvl')[0].getAttribute('w:val'),'0');
  assert.equal(h2.getElementsByTagNameNS(W,'outlineLvl')[0].getAttribute('w:val'),'1');
  assert.match(out.files['word/styles.xml'],/w:ascii="Georgia"/);assert.match(out.files['word/styles.xml'],/w:ascii="Arial"/);
  const pg=doc.getElementsByTagNameNS(W,'pgSz')[0],mar=doc.getElementsByTagNameNS(W,'pgMar')[0];
  assert.equal(pg.getAttribute('w:w'),'11906');assert.equal(pg.getAttribute('w:h'),'16838');assert.equal(mar.getAttribute('w:left'),'1296');
  for(const key of ['header','footer','gutter'])assert.equal(mar.getAttribute(`w:${key}`),'0');
  assert.deepEqual(out.preview.sectionGeometryTwips,{width:11906,height:16838,top:1296,bottom:1296,left:1296,right:1296,header:0,footer:0,gutter:0});
  assert.match(out.files['word/styles.xml'],/w:fill="202020"/);assert.match(out.files['word/styles.xml'],/w:color w:val="FFFFFF"/);
  assert.ok(out.preview.responseLineTwips>=576);assert.ok(text(doc).includes('École 日本語'));
 }
});

test('grayscale converts every rendered color and body-role spacing follows the reviewed brand',()=>{
 const semantic=docs(fixture()).selected;
 for(const family of WORKBOOK_FAMILIES){
  const brand=resolveWorkbookBrand(family.id,{text:'#123456',background:'#FFFFFF',accent:'#765432',accentText:'#FFFFFF',accentBackground:'#203850',lineSpacing:1.6,printMode:'grayscale'});
  assert.equal(brand.eligible,true);
  const out=renderWorkbookFamily(semantic,brand,stamp),styles=xmlDoc(out.files['word/styles.xml']);
  assert.equal(out.preview.printMode,'grayscale');assert.equal(out.preview.lineSpacing,1.6);
  for(const value of Object.values(out.preview.renderedColors))assert.match(value,/^#([0-9A-F]{2})\1\1$/u);
  const styleNodes=Array.from(styles.getElementsByTagNameNS(W,'style'));
  for(const name of ['Normal','Scope','Meta','Objective','Prompt','ResponseLabel']){
   const item=styleNodes.find(x=>x.getAttribute('w:styleId')===name);
   assert.equal(item.getElementsByTagNameNS(W,'spacing')[0].getAttribute('w:line'),'384',`${family.id}/${name}`);
  }
  for(const color of out.files['word/styles.xml'].matchAll(/w:(?:color|fill)="([0-9A-F]{6})"/gu))assert.match(color[1],/^([0-9A-F]{2})\1\1$/u);
 }
});

test('malformed semantic, forged brand, XML boundaries, unsupported profile and nonwhite backgrounds refuse the new route',()=>{
 const f=fixture(),semantic=docs(f).selected,brand=resolveWorkbookBrand('clearline');
 const refuse=(s=semantic,b=brand,t=stamp)=>assert.throws(()=>renderWorkbookFamily(s,b,t),/Unsupported styled workbook/);
 refuse({...semantic,version:2});refuse({...semantic,scope:{...semantic.scope,planId:null}});
 refuse({...semantic,nodes:[{kind:'unknown',id:'x'}]});
 refuse({...semantic,nodes:semantic.nodes.map((n,i)=>i===1?{...n,moduleId:'wrong'}:n)});
 refuse({...semantic,title:'bad\ud800'});refuse({...semantic,title:'bad\ufffe'});
 refuse({...semantic,nodes:semantic.nodes.map((n,i)=>i===1?{...n,responseLines:21}:n)});
 const accessor={...semantic};Object.defineProperty(accessor,'title',{get(){throw Error('getter ran');},enumerable:true});refuse(accessor);
 const node={...semantic.nodes[0]};Object.defineProperty(node,'kind',{get(){throw Error('node getter ran');},enumerable:true});refuse({...semantic,nodes:[node,...semantic.nodes.slice(1)]});
 const forged={...brand,effective:{...brand.effective,text:'#EEEEEE'}};refuse(semantic,forged);
 const hidden={...brand,effective:{...brand.effective}};Object.defineProperty(hidden.effective,'text',{get(){throw Error('hidden getter ran');},enumerable:false});refuse(semantic,hidden);
 refuse(semantic,resolveWorkbookBrand('clearline',{background:'#202020',text:'#FFFFFF'}));
 refuse(semantic,resolveWorkbookBrand('clearline',{text:'#EEEEEE'}));
 refuse(semantic,brand,'yesterday');
});

test('legacy full and selected six-part maps remain unchanged after every family output',()=>{
 for(const wrench of [false,true]){
  const f=fixture(wrench),legacyFull=worksheetFiles(f.p,stamp),legacySelected=planWorkbookFiles(f.p,f.planId,stamp);
  const sem=docs(f);
  for(const family of WORKBOOK_FAMILIES)for(const semantic of Object.values(sem))renderWorkbookFamily(semantic,resolveWorkbookBrand(family.id),stamp);
  assert.deepEqual(worksheetFiles(f.p,stamp),legacyFull);assert.deepEqual(planWorkbookFiles(f.p,f.planId,stamp),legacySelected);
 }
});

test('consecutive response paragraphs encode bottom and between rules with visible fallback ink',()=>{
 const semantic=docs(fixture()).full;
 for(const family of WORKBOOK_FAMILIES)for(const accent of ['#FFFFFF','#F8F8F8','#246A82']){
  const out=renderWorkbookFamily(semantic,resolveWorkbookBrand(family.id,{accent}),stamp);
  const styles=xmlDoc(out.files['word/styles.xml']);
  const response=Array.from(styles.getElementsByTagNameNS(W,'style')).find(x=>x.getAttribute('w:styleId')==='ResponseLine');
  const border=response.getElementsByTagNameNS(W,'pBdr')[0],bottom=border.getElementsByTagNameNS(W,'bottom')[0],between=border.getElementsByTagNameNS(W,'between')[0];
  assert.ok(bottom&&between,`${family.id}/${accent} needs both border encodings`);
  assert.equal(bottom.getAttribute('w:color'),between.getAttribute('w:color'));
  assert.equal(bottom.getAttribute('w:color'),out.preview.appliedStyles.ResponseLine.border.color);
  if(accent==='#FFFFFF'||accent==='#F8F8F8')assert.equal(bottom.getAttribute('w:color'),'767676');
  assert.equal(paragraphs(xmlDoc(out.files['word/document.xml'])).filter(p=>pStyle(p)==='ResponseLine').length,21);
 }
});

test('selected copy requires exact identity and policy and publishes typed plan ID',()=>{
 const semantic=docs(fixture()).selected,brand=resolveWorkbookBrand('clearline');
 const out=renderWorkbookFamily(semantic,brand,stamp),body=text(xmlDoc(out.files['word/document.xml']));
 assert.ok(body.includes(`Plan ID: ${semantic.scope.planId}`));
 assert.equal(out.preview.scope.planId,semantic.scope.planId);
 const bad=lines=>assert.throws(()=>renderWorkbookFamily({...semantic,identityLines:lines},brand,stamp),/Unsupported styled workbook/);
 bad(['Arbitrary identity']);bad(semantic.identityLines.slice(0,4));
 bad(semantic.identityLines.map((line,i)=>i===0?line.replace(semantic.scope.planId,'wrong'):line));
 bad(semantic.identityLines.map((line,i)=>i===1?'Exact selected scope: wrong: item':line));
 bad(semantic.identityLines.map((line,i)=>i===3?'Unreviewed objective claim':line));
 bad(semantic.identityLines.map((line,i)=>i===4?'Unreviewed policy claim':line));
});

test('selected identity preserves canonical fractional planned time across agenda rollover',()=>{
 let p=course('Fractional workshop');p=addModule(p,'One','Learn').next;
 p=addActivity(p,p.modules[0].id,{title:'Short practice',minutes:0.9,participant:'Try it',facilitator:'private',answers:'private',lines:1}).next;
 const groups=[{moduleId:p.modules[0].id,activityIds:[p.modules[0].activities[0].id]}];
 const created=editDeliveryPlan(p,{kind:'create',name:'Late plan',targetMinutes:1,groups});p=created.next;
 const semantic=semanticWorkbookFromSelected(p,created.id,stamp,'23:59');
 assert.ok(semantic.identityLines[2].includes('0.9 minutes'));
 const out=renderWorkbookFamily(semantic,resolveWorkbookBrand('clearline'),stamp);
 assert.ok(text(xmlDoc(out.files['word/document.xml'])).includes('0.9 minutes'));
 assert.equal(out.preview.identityLines[2],semantic.identityLines[2]);
});

test('low-level selected renderer refuses an empty or empty-only plan roster',()=>{
 const semantic=docs(fixture()).selected,brand=resolveWorkbookBrand('clearline');
 const identityLines=semantic.identityLines.map((line,i)=>i===1?'Exact selected scope: ':i===2?'Selected planned time: 0 minutes in saved plan order; absolute start times appear only in the agenda.':line);
 assert.throws(()=>renderWorkbookFamily({...semantic,nodes:[],identityLines},brand,stamp),/selected scope needs at least one activity/);
 const module=semantic.nodes.find(node=>node.kind==='module');
 const onlyEmpty=identityLines.map((line,i)=>i===1?`Exact selected scope: ${module.id}: `:line);
 assert.throws(()=>renderWorkbookFamily({...semantic,nodes:[module],identityLines:onlyEmpty},brand,stamp),/selected scope needs at least one activity/);
});

test('applied preview matches XML typography, spacing, and normalized organization text',()=>{
 const semantic=docs(fixture()).full;
 for(const family of WORKBOOK_FAMILIES){
  const raw={organizationName:'Line A\r\nLine B\rLine C\nLine D\tend',background:'#ffffff',bodyPt:18,headingPt:14,lineSpacing:1.75};
  const resolved=resolveWorkbookBrand(family.id,raw);assert.equal(resolved.eligible,true);
  const out=renderWorkbookFamily(semantic,resolved,stamp),styles=xmlDoc(out.files['word/styles.xml']);
  assert.equal(resolved.rawRequest.organizationName,raw.organizationName);
  assert.equal(out.preview.organizationName,'Line A\nLine B\nLine C\nLine D\tend');
  assert.ok(text(xmlDoc(out.files['word/document.xml'])).includes('Organization: Line ALine BLine CLine D\tend'));
  const named=Array.from(styles.getElementsByTagNameNS(W,'style'));
  for(const [name,applied] of Object.entries(out.preview.appliedStyles)){
   const item=named.find(x=>x.getAttribute('w:styleId')===name);assert.ok(item,name);
   assert.equal(Number(item.getElementsByTagNameNS(W,'sz')[0].getAttribute('w:val'))/2,applied.sizePt);
   assert.equal(item.getElementsByTagNameNS(W,'spacing')[0].getAttribute('w:line'),String(applied.lineTwips));
   assert.equal(item.getElementsByTagNameNS(W,'rFonts')[0].getAttribute('w:ascii'),applied.font);
   assert.equal(item.getElementsByTagNameNS(W,'rFonts')[0].getAttribute('w:eastAsia'),applied.fallbackFont);
  }
  assert.ok(out.preview.appliedStyles.Title.sizePt>out.preview.appliedStyles.Heading1.sizePt);
  assert.ok(out.preview.appliedStyles.Heading1.sizePt>out.preview.appliedStyles.Heading2.sizePt);
  assert.ok(out.preview.appliedStyles.Heading2.sizePt>out.preview.appliedStyles.Normal.sizePt);
 }
});

test('preview separates fractional base size tokens from actual half-point Word sizes',()=>{
 const semantic=docs(fixture()).full;
 for(const family of WORKBOOK_FAMILIES){
  const brand=resolveWorkbookBrand(family.id,{bodyPt:12.24,headingPt:16.24});
  const out=renderWorkbookFamily(semantic,brand,stamp),styles=xmlDoc(out.files['word/styles.xml']);
  assert.deepEqual(out.preview.effectiveBaseTypography,{bodyPt:12.24,headingPt:16.24});
  assert.equal(out.preview.appliedStyles.Normal.sizePt,12);
  for(const [name,applied] of Object.entries(out.preview.appliedStyles)){
   const item=Array.from(styles.getElementsByTagNameNS(W,'style')).find(x=>x.getAttribute('w:styleId')===name);
   assert.equal(Number(item.getElementsByTagNameNS(W,'sz')[0].getAttribute('w:val'))/2,applied.sizePt);
  }
 }
});

test('paragraph and run properties follow the supported Word profile sequence',()=>{
 const semantic=docs(fixture()).full;
 const pOrder=['pStyle','keepNext','pageBreakBefore','pBdr','shd','spacing','ind','jc','outlineLvl'];
 const rOrder=['rFonts','b','i','color','sz'];
 const check=(element,order)=>{
  const names=Array.from(element.childNodes).filter(x=>x.nodeType===1).map(x=>x.localName);
  assert.deepEqual(names,[...names].sort((a,b)=>order.indexOf(a)-order.indexOf(b)));
  assert.ok(names.every(name=>order.includes(name)));
 };
 for(const family of WORKBOOK_FAMILIES){
  const out=renderWorkbookFamily(semantic,resolveWorkbookBrand(family.id),stamp);
  for(const item of Array.from(xmlDoc(out.files['word/styles.xml']).getElementsByTagNameNS(W,'style'))){
   check(item.getElementsByTagNameNS(W,'pPr')[0],pOrder);
   check(item.getElementsByTagNameNS(W,'rPr')[0],rOrder);
  }
  for(const item of paragraphs(xmlDoc(out.files['word/document.xml'])))check(item.getElementsByTagNameNS(W,'pPr')[0],pOrder);
 }
});
