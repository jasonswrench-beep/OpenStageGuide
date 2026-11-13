import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import {course,wrenchCourse,addModule,addActivity} from '../src/model.mjs';
import {editDeliveryPlan} from '../src/delivery-plans.mjs';
import {WORKBOOK_FAMILIES} from '../src/workbook-presentation.mjs';
import {FAMILY_BRAND_FIELDS,familyBrandRequest,prepareFamilyWorkbook,familyReviewText,familyWorkbookBlob,familyWorkbookName} from '../src/workbook-family-flow.mjs';
import {worksheetFiles} from '../src/worksheet.mjs';
import {planWorkbookFiles} from '../src/selected-plan-outputs.mjs';
import {rawControl} from '../src/activity-draft.mjs';

const stamp='2026-10-26T10:00:00.000Z';
const raw=overrides=>Object.fromEntries(FAMILY_BRAND_FIELDS.map(key=>[key,rawControl({value:String(overrides?.[key]??'')})]));
function fixture(wrench=false){
 let p=wrench?wrenchCourse('Learning 🌱',{audience:'PRIVATE_AUDIENCE',purpose:'PRIVATE_PURPOSE',authors:'PRIVATE_AUTHOR',attribution:'PRIVATE_RIGHTS'},['background']):course('Learning 🌱');
 p=addModule(p,'First','Objective first').next;p=addModule(p,'Second','Objective second').next;
 p=addActivity(p,p.modules[0].id,{title:'Omitted',minutes:5,participant:'OMITTED_CONTENT',facilitator:'PRIVATE_TEACH',answers:'PRIVATE_ANSWER',lines:0}).next;
 p=addActivity(p,p.modules[0].id,{title:'Practice',minutes:5,participant:'Learner practice\nSecond line 日本語',facilitator:'PRIVATE_TEACH',answers:'PRIVATE_ANSWER',lines:20}).next;
 p=addActivity(p,p.modules[1].id,{title:'Reflection',minutes:5,participant:'Learner reflection',facilitator:'PRIVATE_TEACH',answers:'PRIVATE_ANSWER',lines:1}).next;
 const groups=[{moduleId:p.modules[1].id,activityIds:[p.modules[1].activities[0].id]},{moduleId:p.modules[0].id,activityIds:[p.modules[0].activities[1].id]}];
 const made=editDeliveryPlan(p,{kind:'create',name:'Reviewed',targetMinutes:10,groups});return {p:made.next,id:made.id};
}

test('family flow binds all five renderers to full and selected public content without touching source or legacy maps',async()=>{
 for(const wrench of [false,true]){
  const {p,id}=fixture(wrench),before=structuredClone(p),legacyFull=worksheetFiles(p,stamp),legacySelected=planWorkbookFiles(p,id,stamp);
  for(const planId of ['',id])for(const family of WORKBOOK_FAMILIES){
   const prepared=prepareFamilyWorkbook(p,{planId,familyId:family.id,rawBrand:raw({organizationName:'Example College'}),stamp});
   assert.equal(prepared.eligible,true);
   assert.equal(prepared.output.preview.familyId,family.id);
   assert.equal(prepared.output.preview.scope.kind,planId?'selected':'full');
   assert.deepEqual(prepared.output.preview.sections.flatMap(s=>s.activities.map(a=>a.responseLines)),planId?[1,20]:[0,20,1]);
   const review=familyReviewText(prepared);
   assert.ok(review.includes('Example College')&&review.includes('Applied text and response roles'));
   assert.ok(review.includes('Learner practice')&&review.includes('Writing lines: 20'));
   if(planId){assert.ok(review.includes(id));assert.ok(!review.includes('OMITTED_CONTENT'));}
   assert.doesNotMatch(review,/PRIVATE_TEACH|PRIVATE_ANSWER|PRIVATE_AUDIENCE|PRIVATE_PURPOSE|PRIVATE_AUTHOR|PRIVATE_RIGHTS/);
   const blob=await familyWorkbookBlob(prepared.output.files,JSZip),zip=await JSZip.loadAsync(await blob.arrayBuffer());
   assert.deepEqual(Object.keys(zip.files).sort(),Object.keys(prepared.output.files).sort());
  }
  assert.deepEqual(p,before);assert.deepEqual(worksheetFiles(p,stamp),legacyFull);assert.deepEqual(planWorkbookFiles(p,id,stamp),legacySelected);
 }
});

test('raw brand parsing refuses invalid buffers and uses family defaults only for blank overrides',()=>{
 const clean=raw({organizationName:'My college',bodyPt:'12.25',headingFont:'Georgia'}),request=familyBrandRequest(clean);
 assert.deepEqual(request,{organizationName:'My college',headingFont:'Georgia',bodyPt:12.25});
 for(const value of ['12abc','1e2','NaN','Infinity','-1',' 12','12.'])assert.throws(()=>familyBrandRequest(raw({bodyPt:value})),/plain decimal/);
 const bad=raw();bad.bodyPt.invalid.badInput=true;assert.throws(()=>familyBrandRequest(bad),/unfinished or invalid/);
 const {p}=fixture();
 const clear=prepareFamilyWorkbook(p,{planId:'',familyId:'clearline',rawBrand:raw(),stamp});
 const editorial=prepareFamilyWorkbook(p,{planId:'',familyId:'editorial-study',rawBrand:raw(),stamp});
 assert.notEqual(clear.brand.effective.headingFont,editorial.brand.effective.headingFont);
 assert.equal(clear.brand.rawRequest.headingFont,undefined);
});

test('exported raw-control and scope boundaries refuse hidden, inherited, getter and nontext values without reading getters',()=>{
 const source=raw(),{p}=fixture();let reads=0;
 const badGetter={...source};Object.defineProperty(badGetter,'text',{enumerable:true,get(){reads++;throw Error('getter');}});
 assert.throws(()=>familyBrandRequest(badGetter),/own data/);assert.equal(reads,0);
 const hidden={...source};Object.defineProperty(hidden,'text',{value:source.text,enumerable:false});assert.throws(()=>familyBrandRequest(hidden),/own data/);
 const extra={...source,unknown:rawControl({value:'private'})};assert.throws(()=>familyBrandRequest(extra),/fields changed/);
 const inherited=Object.assign(Object.create({unknown:'private'}),source);assert.throws(()=>familyBrandRequest(inherited),/plain data record/);
 const badEntry=raw();Object.defineProperty(badEntry.text,'value',{enumerable:true,get(){reads++;throw Error('getter');}});
 assert.throws(()=>familyBrandRequest(badEntry),/own data/);assert.equal(reads,0);
 const badValidity=raw();Object.defineProperty(badValidity.text.invalid,'badInput',{value:true,enumerable:false});assert.throws(()=>familyBrandRequest(badValidity),/own data/);
 const wrong=raw();wrong.text.value=17;assert.throws(()=>familyBrandRequest(wrong),/unfinished or invalid/);
 for(const planId of [null,false,0])assert.throws(()=>prepareFamilyWorkbook(p,{planId,familyId:'clearline',rawBrand:raw(),stamp}),/scope, design or generation/);
});

test('invalid contrast and nonwhite page refuse locally with explicit alternatives and retain raw request',()=>{
 const {p}=fixture();
 const contrast=prepareFamilyWorkbook(p,{planId:'',familyId:'clearline',rawBrand:raw({text:'#EEEEEE',organizationName:'Keep me'}),stamp});
 assert.equal(contrast.eligible,false);assert.equal(contrast.brand.rawRequest.text,'#EEEEEE');assert.equal(contrast.alternative.text,'#000000');
 const page=prepareFamilyWorkbook(p,{planId:'',familyId:'clearline',rawBrand:raw({background:'#202020',text:'#FFFFFF'}),stamp});
 assert.equal(page.eligible,false);assert.deepEqual(page.alternative,{background:'#FFFFFF',text:'#1A1A1A'});
 const font=prepareFamilyWorkbook(p,{planId:'',familyId:'clearline',rawBrand:raw({headingFont:'Invalid Font'}),stamp});
 assert.equal(font.eligible,false);assert.ok(font.issues.some(issue=>issue.includes('headingFont')));
});

test('grayscale print contrast witness is refused with an explicit safe text proposal',()=>{
 const {p}=fixture(),request=raw({text:'#0072F0',printMode:'grayscale'});
 const original=structuredClone(request),result=prepareFamilyWorkbook(p,{planId:'',familyId:'clearline',rawBrand:request,stamp});
 assert.equal(result.eligible,false);assert.match(result.issues.join(' '),/grayscale conversion/);
 assert.equal(result.brand.rawRequest.text,'#0072F0');assert.equal(result.alternative.text,'#000000');
 assert.deepEqual(request,original);
 const accepted=prepareFamilyWorkbook(p,{planId:'',familyId:'clearline',rawBrand:raw({text:'#000000',printMode:'grayscale'}),stamp});
 assert.equal(accepted.eligible,true);
});

test('family DOCX package and unique filenames are verified after serialization',async()=>{
 const {p}=fixture(),prepared=prepareFamilyWorkbook(p,{planId:'',familyId:'field-guide',rawBrand:raw(),stamp});
 const first=familyWorkbookName(p.title,'field-guide'),second=familyWorkbookName(p.title,'field-guide');
 assert.notEqual(first,second);assert.match(first,/Field Guide Workbook.*\.docx$/);
 assert.match(familyWorkbookName(p.title,'field-guide','Reviewed'),/Reviewed/);
 class ChangingZip extends JSZip{async generateAsync(options){this.file('word/document.xml','Changed',{createFolders:false});return super.generateAsync(options);}}
 await assert.rejects(familyWorkbookBlob(prepared.output.files,ChangingZip),/member word\/document.xml changed/);
 const blob=await familyWorkbookBlob(prepared.output.files,JSZip);assert.ok(blob.size>0);
});

test('ZIP offer uses private verified bytes when producer retains and changes its buffer during validation',async()=>{
 const {p}=fixture(),prepared=prepareFamilyWorkbook(p,{planId:'',familyId:'clearline',rawBrand:raw(),stamp});
 class RetainingZip extends JSZip{
  static retained;
  async generateAsync(options){const result=await super.generateAsync(options);RetainingZip.retained=result;return result;}
  static async loadAsync(bytes){RetainingZip.retained.fill(0);return super.loadAsync(bytes);}
 }
 const blob=await familyWorkbookBlob(prepared.output.files,RetainingZip),zip=await JSZip.loadAsync(await blob.arrayBuffer());
 assert.equal(await zip.file('word/document.xml').async('string'),prepared.output.files['word/document.xml']);
});

test('ZIP offer rejects malformed UTF-8 even when replacement decoding could match expected text',async()=>{
 const {p}=fixture();p.title='Replacement \uFFFD title';
 const prepared=prepareFamilyWorkbook(p,{planId:'',familyId:'clearline',rawBrand:raw(),stamp});
 class InvalidUtf8Zip extends JSZip{async generateAsync(options){
  const name='docProps/core.xml',bytes=new TextEncoder().encode(prepared.output.files[name]);
  const at=bytes.findIndex((byte,i)=>byte===0xef&&bytes[i+1]===0xbf&&bytes[i+2]===0xbd);
  assert.ok(at>0);const changed=Uint8Array.from([...bytes.slice(0,at),0xff,...bytes.slice(at+3)]);
  this.file(name,changed,{createFolders:false});return super.generateAsync(options);
 }}
 await assert.rejects(familyWorkbookBlob(prepared.output.files,InvalidUtf8Zip),/member docProps\/core.xml changed/);
});

test('new family profile refuses absent, foreign, malformed or changed compatibility parts before packaging',async()=>{
 const {p}=fixture(),prepared=prepareFamilyWorkbook(p,{planId:'',familyId:'clearline',rawBrand:raw(),stamp});
 const files=prepared.output.files;
 const altered=[
  Object.fromEntries(Object.entries(files).filter(([name])=>name!=='word/settings.xml')),
  {...files,'word/foreign.xml':'<x/>'},
  {...files,'word/settings.xml':files['word/settings.xml'].replace('w:val="15"','w:val="12"')},
  {...files,'word/settings.xml':files['word/settings.xml'].replace('</w:settings>','<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:settings>')},
  {...files,'word/settings.xml':'<w:settings>'},
  {...files,'word/_rels/document.xml.rels':files['word/_rels/document.xml.rels'].replace('Target="settings.xml"','Target="foreign.xml"')},
  {...files,'word/_rels/document.xml.rels':files['word/_rels/document.xml.rels'].replace('rId2','rId1')},
  {...files,'[Content_Types].xml':files['[Content_Types].xml'].replace('wordprocessingml.settings+xml','xml')}
 ];
 for(const candidate of altered)await assert.rejects(familyWorkbookBlob(candidate,JSZip),/generated DOCX/);
});

test('new family ZIP rejects duplicate central-directory member names before JSZip map can hide them',async()=>{
 const {p}=fixture(),prepared=prepareFamilyWorkbook(p,{planId:'',familyId:'clearline',rawBrand:raw(),stamp});
 class DuplicateZip extends JSZip{async generateAsync(options){
  const bytes=Uint8Array.from(await super.generateAsync(options)),view=new DataView(bytes.buffer),end=bytes.length-22;
  let at=view.getUint32(end+16,true),changed=false;
  for(let i=0;i<7;i++){
   const length=view.getUint16(at+28,true),extra=view.getUint16(at+30,true),comment=view.getUint16(at+32,true);
   const name=new TextDecoder().decode(bytes.subarray(at+46,at+46+length));
   if(name==='word/settings.xml'){
    const replacement=new TextEncoder().encode('word/document.xml');assert.equal(replacement.length,length);
    bytes.set(replacement,at+46);changed=true;break;
   }
   at+=46+length+extra+comment;
  }
  assert.equal(changed,true);return bytes;
 }}
 await assert.rejects(familyWorkbookBlob(prepared.output.files,DuplicateZip),/ZIP envelope changed/);
});

test('new family ZIP refuses ghost local settings and unknown records absent from central inventory',async()=>{
 const {p}=fixture(),prepared=prepareFamilyWorkbook(p,{planId:'',familyId:'clearline',rawBrand:raw(),stamp});
 const insertGhost=(bytes,name)=>{
  const old=Uint8Array.from(bytes),view=new DataView(old.buffer),end=old.length-22,offset=view.getUint32(end+16,true);
  const label=new TextEncoder().encode(name),body=new TextEncoder().encode('GHOST_DUPLICATE_LOCAL_SETTINGS');
  const ghost=new Uint8Array(30+label.length+body.length),g=new DataView(ghost.buffer);
  g.setUint32(0,0x04034b50,true);g.setUint32(18,body.length,true);g.setUint32(22,body.length,true);
  g.setUint16(26,label.length,true);ghost.set(label,30);ghost.set(body,30+label.length);
  const result=new Uint8Array(old.length+ghost.length);
  result.set(old.subarray(0,offset));result.set(ghost,offset);result.set(old.subarray(offset),offset+ghost.length);
  new DataView(result.buffer).setUint32(result.length-22+16,offset+ghost.length,true);
  return result;
 };
 for(const name of ['word/settings.xml','word/unknown.xml']){
  class GhostZip extends JSZip{async generateAsync(options){return insertGhost(await super.generateAsync(options),name);}}
  await assert.rejects(familyWorkbookBlob(prepared.output.files,GhostZip),/ZIP envelope changed/);
 }
});

test('new family ZIP refuses local-header name conflicts even when central names are correct',async()=>{
 const {p}=fixture(),prepared=prepareFamilyWorkbook(p,{planId:'',familyId:'clearline',rawBrand:raw(),stamp});
 class ConflictZip extends JSZip{async generateAsync(options){
  const bytes=Uint8Array.from(await super.generateAsync(options)),view=new DataView(bytes.buffer),end=bytes.length-22;
  let at=view.getUint32(end+16,true),changed=false;
  for(let i=0;i<7;i++){
   const n=view.getUint16(at+28,true),extra=view.getUint16(at+30,true),comment=view.getUint16(at+32,true);
   const name=new TextDecoder().decode(bytes.subarray(at+46,at+46+n));
   if(name==='word/settings.xml'){
    const local=view.getUint32(at+42,true),replacement=new TextEncoder().encode('word/document.xml');
    assert.equal(view.getUint16(local+26,true),replacement.length);
    bytes.set(replacement,local+30);changed=true;break;
   }
   at+=46+n+extra+comment;
  }
  assert.equal(changed,true);return bytes;
 }}
 await assert.rejects(familyWorkbookBlob(prepared.output.files,ConflictZip),/ZIP envelope changed/);
});

test('new family ZIP refuses unreviewed central file comments',async()=>{
 const {p}=fixture(),prepared=prepareFamilyWorkbook(p,{planId:'',familyId:'clearline',rawBrand:raw(),stamp});
 class CommentZip extends JSZip{async generateAsync(options){
  this.file('word/settings.xml').comment='UNREVIEWED_PROFILE3_METADATA';
  return super.generateAsync(options);
 }}
 await assert.rejects(familyWorkbookBlob(prepared.output.files,CommentZip),/ZIP entry metadata changed/);
});

test('new family ZIP refuses unreviewed central extra fields',async()=>{
 const {p}=fixture(),prepared=prepareFamilyWorkbook(p,{planId:'',familyId:'clearline',rawBrand:raw(),stamp});
 class ExtraZip extends JSZip{async generateAsync(options){
  const old=Uint8Array.from(await super.generateAsync(options)),v=new DataView(old.buffer),end=old.length-22;
  const at=v.getUint32(end+16,true),nameLength=v.getUint16(at+28,true),insertion=at+46+nameLength;
  const marker=new TextEncoder().encode('UNREVIEWED_PROFILE3_METADATA');
  const extra=new Uint8Array(marker.length+4),e=new DataView(extra.buffer);
  e.setUint16(0,0xcafe,true);e.setUint16(2,marker.length,true);extra.set(marker,4);
  const bytes=new Uint8Array(old.length+extra.length);
  bytes.set(old.subarray(0,insertion));bytes.set(extra,insertion);bytes.set(old.subarray(insertion),insertion+extra.length);
  const out=new DataView(bytes.buffer);
  out.setUint16(at+30,extra.length,true);
  out.setUint32(bytes.length-22+12,v.getUint32(end+12,true)+extra.length,true);
  return bytes;
 }}
 await assert.rejects(familyWorkbookBlob(prepared.output.files,ExtraZip),/ZIP entry metadata changed/);
});

test('new family ZIP verifies CRC against actual bytes even when local and central headers agree on a wrong value',async()=>{
 const {p}=fixture(),prepared=prepareFamilyWorkbook(p,{planId:'',familyId:'clearline',rawBrand:raw(),stamp});
 class BadCrcZip extends JSZip{async generateAsync(options){
  const bytes=Uint8Array.from(await super.generateAsync(options)),view=new DataView(bytes.buffer),end=bytes.length-22;
  const central=view.getUint32(end+16,true),local=view.getUint32(central+42,true);
  const wrong=(view.getUint32(central+16,true)^0x01010101)>>>0;
  view.setUint32(central+16,wrong,true);view.setUint32(local+14,wrong,true);
  return bytes;
 }}
 await assert.rejects(familyWorkbookBlob(prepared.output.files,BadCrcZip),/CRC32 mismatch/);
});
