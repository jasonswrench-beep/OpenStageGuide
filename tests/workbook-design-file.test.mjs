import test from 'node:test';
import assert from 'node:assert/strict';
import {createHash} from 'node:crypto';
import {WORKBOOK_DESIGN_FILE_FORMAT,WORKBOOK_DESIGN_FILE_VERSION,WORKBOOK_DESIGN_FILE_LIMITS,validateWorkbookDesignDraft,createWorkbookDesignFile,serializeWorkbookDesignFile,parseWorkbookDesignFile,prepareWorkbookDesignFile} from '../src/workbook-design-file.mjs';
import {prepareWorkbookLogo,workbookLogoInfo,workbookLogoBytes} from '../src/workbook-logo.mjs';
import {WORKBOOK_FAMILIES} from '../src/workbook-presentation.mjs';
import {prepareFamilyWorkbook,FAMILY_BRAND_FIELDS} from '../src/workbook-family-flow.mjs';
import {rawControl} from '../src/activity-draft.mjs';
import {editDeliveryPlan} from '../src/delivery-plans.mjs';
import {fixture} from './helpers/mock-office.mjs';
import {syntheticLogo,logoChunk} from './helpers/logo-fixture.mjs';
import {workbookDesignDraft} from './helpers/workbook-design-file-fixture.mjs';
const copy=x=>structuredClone(x),enc=s=>new TextEncoder().encode(s),sha=b=>createHash('sha256').update(b).digest('hex');
const round=file=>parseWorkbookDesignFile(serializeWorkbookDesignFile(file));
function freezeCheck(x){if(x&&typeof x==='object'){assert.ok(Object.isFrozen(x));for(const v of Object.values(x))freezeCheck(v);}}
const fileWithImage=(draft,bytes)=>{const file=copy(createWorkbookDesignFile(draft,null));file.payload.logo.image={mediaType:'image/png',normalizerVersion:1,sha256:sha(bytes),base64:Buffer.from(bytes).toString('base64')};return file;};

test('raw PRIVATE design files preserve five families, authored Unicode and incomplete/inactive work without authority or source content',async()=>{
 assert.equal(WORKBOOK_DESIGN_FILE_FORMAT,'openstage-workbook-design');assert.equal(WORKBOOK_DESIGN_FILE_VERSION,1);
 for(const {id} of WORKBOOK_FAMILIES){const d=workbookDesignDraft(id),before=copy(d),file=createWorkbookDesignFile(d,null),raw=serializeWorkbookDesignFile(file),read=parseWorkbookDesignFile(raw),prepared=await prepareWorkbookDesignFile(read);assert.deepEqual(prepared.draft,before);assert.equal(prepared.logoAsset,null);assert.equal(prepared.normalization,null);assert.deepEqual(d,before);assert.equal(Object.isFrozen(d),false);freezeCheck(prepared);assert.deepEqual(Object.keys(read),['format','version','payload']);assert.deepEqual(Object.keys(read.payload),['familyId','baseLanguage','brand','frontMatter','contents','research','logo']);assert.equal(Object.hasOwn(read.payload.logo,'image'),true);assert.equal(read.payload.logo.image,null);
  d.frontMatter.cover.title='Changed after capture';raw.fill(0);assert.equal(prepared.draft.frontMatter.cover.title,before.frontMatter.cover.title);assert.deepEqual(round(file),file);
 }
 const raw=workbookDesignDraft();raw.brand.text='#x';raw.brand.bodyPt='1.';raw.frontMatter.cover.title='';raw.frontMatter.copyright.year='20-';raw.frontMatter.copyright.licenseName='Inactive named-license text';raw.research.rationale='Unfinished \u0000 JSON-only draft\r\n\t\u{1f642}';raw.logo={include:true,role:'decorative',alt:'',widthIn:'1.'};const p=await prepareWorkbookDesignFile(round(createWorkbookDesignFile(raw,null)));assert.deepEqual(p.draft,raw);
 for(const key of ['course','project','planId','scope','responses','assessment','consent','reviewed','stamp','effectiveTokens','presets']){const x=copy(raw);x[key]=null;assert.throws(()=>validateWorkbookDesignDraft(x));}
});
test('factory saves only privately owned verified PNG bytes, revalidation creates a fresh handle and originals remain intact',async()=>{
 const source=syntheticLogo({before:[logoChunk('tEXt',enc('PRIVATE_SYNTHETIC_PATH\0C:/private/original.png'))]}),before=new Uint8Array(source),handle=await prepareWorkbookLogo(source),d=workbookDesignDraft();
 const file=createWorkbookDesignFile(d,handle),info=workbookLogoInfo(handle),read=round(file);assert.equal(read.payload.logo.image.sha256,info.sha256);assert.equal(read.payload.logo.include,false);assert.equal(JSON.stringify(file).includes(handle.id),false);const stored=Buffer.from(read.payload.logo.image.base64,'base64');assert.equal(new TextDecoder().decode(stored).includes('PRIVATE_SYNTHETIC_PATH'),false);assert.deepEqual(new Uint8Array(stored),workbookLogoBytes(handle));
 const a=await prepareWorkbookDesignFile(read),b=await prepareWorkbookDesignFile(read);assert.notEqual(a.logoAsset,handle);assert.notEqual(a.logoAsset,b.logoAsset);assert.notEqual(a.logoInfo.id,b.logoInfo.id);assert.deepEqual(workbookLogoBytes(a.logoAsset),workbookLogoBytes(handle));assert.deepEqual(a.draft,d);assert.deepEqual(source,before);assert.deepEqual(a.normalization,{storedSha256:info.sha256,currentSha256:a.logoInfo.sha256,bytesChanged:false});
 assert.throws(()=>createWorkbookDesignFile(d,{...handle}));assert.throws(()=>createWorkbookDesignFile(d,undefined));
 const altered=workbookLogoBytes(a.logoAsset);altered.fill(0);assert.deepEqual(workbookLogoBytes(a.logoAsset),workbookLogoBytes(handle));
 const metadataPng=syntheticLogo({before:[logoChunk('tEXt',enc('SAFE_DESCRIPTION\0metadata stripped'))]}),changed=await prepareWorkbookDesignFile(fileWithImage(d,metadataPng));assert.equal(changed.normalization.bytesChanged,true);assert.notEqual(changed.normalization.storedSha256,changed.normalization.currentSha256);assert.equal(new TextDecoder().decode(workbookLogoBytes(changed.logoAsset)).includes('SAFE_DESCRIPTION'),false);
});
test('current PNG/hash checks refuse corrupt and unsupported images before adoption while checking bounded canonical base64',async()=>{
 const d=workbookDesignDraft(),valid=syntheticLogo();
 for(const mutate of [f=>f.payload.logo.image.sha256='0'.repeat(64),f=>f.payload.logo.image.mediaType='image/jpeg',f=>f.payload.logo.image.normalizerVersion=2,f=>f.payload.logo.image.base64='data:image/png;base64,'+f.payload.logo.image.base64,f=>f.payload.logo.image.base64='AB==',f=>f.payload.logo.image.base64='AAA',f=>f.payload.logo.image.base64='====',f=>f.payload.logo.image.width=4]){const f=fileWithImage(d,valid);mutate(f);await assert.rejects(async()=>prepareWorkbookDesignFile(round(f)));}
 const corrupt=new Uint8Array(valid);corrupt[29]^=1;
 for(const bytes of [corrupt,enc('<svg xmlns="http://www.w3.org/2000/svg"/>'),new Uint8Array([255,216,255,217]),syntheticLogo({width:2049}),syntheticLogo({before:[logoChunk('acTL',new Uint8Array(8))]})])await assert.rejects(prepareWorkbookDesignFile(round(fileWithImage(d,bytes))));
 const boundary=await prepareWorkbookDesignFile(round(fileWithImage(d,syntheticLogo({width:2048,height:1}))));assert.equal(boundary.logoInfo.width,2048);assert.equal(boundary.logoInfo.height,1);
 const max=new Uint8Array(WORKBOOK_DESIGN_FILE_LIMITS.maxPngBytes),exact=fileWithImage(d,max);assert.equal(round(exact).payload.logo.image.base64.length,WORKBOOK_DESIGN_FILE_LIMITS.maxBase64Units);await assert.rejects(prepareWorkbookDesignFile(exact));
 const over=fileWithImage(d,new Uint8Array(WORKBOOK_DESIGN_FILE_LIMITS.maxPngBytes+1));assert.throws(()=>serializeWorkbookDesignFile(over),/2 MiB|large/);
});
test('strict JSON/UTF8/own-data caps refuse ambiguous or unsupported data without invoking getters or truncating raw drafts',()=>{
 const file=createWorkbookDesignFile(workbookDesignDraft(),null),json=new TextDecoder().decode(serializeWorkbookDesignFile(file));
 for(const s of [json.replace('"version": 1','"version": 1, "\\u0076ersion": 1'),json+' false','', '{"x":'+'['.repeat(10000)+'0'+']'.repeat(10000)+'}',JSON.stringify({...file,version:2}),JSON.stringify({...file,payload:null}),json.replace('"format":','"consent": true, "format":')])assert.throws(()=>parseWorkbookDesignFile(enc(s)));
 for(const data of [new Uint8Array([0xc0,0xaf]),new Uint8Array([0xed,0xa0,0xbd]),Uint8Array.from([0xef,0xbb,0xbf,...enc(json)])])assert.throws(()=>parseWorkbookDesignFile(data));
 const d=workbookDesignDraft();d.frontMatter.copyright.noticeText='x'.repeat(16000);assert.equal(validateWorkbookDesignDraft(d).frontMatter.copyright.noticeText.length,16000);d.frontMatter.copyright.noticeText+='x';assert.throws(()=>validateWorkbookDesignDraft(d));
 const width=workbookDesignDraft();width.logo.widthIn='1'.repeat(32);assert.equal(validateWorkbookDesignDraft(width).logo.widthIn,width.logo.widthIn);
 for(const modify of [x=>x.familyId='unknown',x=>x.baseLanguage='xx-XX',x=>x.brand.headingFont='Unknown font',x=>x.logo.widthIn='1'.repeat(33),x=>x.brand.bodyPt='1'.repeat(13),x=>x.frontMatter.cover.title='\ud800',x=>x.logo.alt='\udfff']){const x=workbookDesignDraft();modify(x);assert.throws(()=>validateWorkbookDesignDraft(x));}
 let getters=0;for(const modify of [x=>Object.defineProperty(x,'familyId',{enumerable:true,get(){getters++;throw Error('getter');}}),x=>Object.defineProperty(x.brand,'text',{enumerable:true,get(){getters++;throw Error('getter');}}),x=>Object.defineProperty(x.logo,'alt',{enumerable:true,get(){getters++;throw Error('getter');}}),x=>Object.setPrototypeOf(x,{foreign:true}),x=>x[Symbol('hidden')]=true,x=>Object.defineProperty(x,'hidden',{value:true})]){const x=workbookDesignDraft();modify(x);assert.throws(()=>validateWorkbookDesignDraft(x));}assert.equal(getters,0);
 const canonical=serializeWorkbookDesignFile(file),padded=new Uint8Array(WORKBOOK_DESIGN_FILE_LIMITS.maxFileBytes);padded.fill(32);padded.set(canonical);assert.deepEqual(parseWorkbookDesignFile(padded),file);assert.throws(()=>parseWorkbookDesignFile(new Uint8Array(padded.length+1)),/6 MiB/);assert.throws(()=>parseWorkbookDesignFile(enc('['+'0,'.repeat(100001)+'0]')),/node budget/);
});
test('private byte snapshots and final async input guard prevent buffer aliases and stale source during PNG checking',async()=>{
 const handle=await prepareWorkbookLogo(syntheticLogo()),file=createWorkbookDesignFile(workbookDesignDraft(),handle),source=serializeWorkbookDesignFile(file),buffer=Buffer.from(source);Object.defineProperty(buffer,'byteLength',{get(){throw Error('must use native bounds');}});const read=parseWorkbookDesignFile(buffer);buffer.fill(0);assert.deepEqual(read,file);
 const original=crypto.subtle.digest;
 try{const mutable=copy(file);let calls=0;crypto.subtle.digest=async function(...args){const result=await Reflect.apply(original,this,args);if(++calls===1)mutable.payload.brand.organizationName='Changed while awaiting';return result;};await assert.rejects(prepareWorkbookDesignFile(mutable),/source changed/);
  crypto.subtle.digest=async function(...args){const result=await Reflect.apply(original,this,args);new Uint8Array(args[1].buffer,args[1].byteOffset,args[1].byteLength).fill(0);return result;};const intact=await prepareWorkbookDesignFile(file);assert.deepEqual(workbookLogoBytes(intact.logoAsset),workbookLogoBytes(handle));
 }finally{crypto.subtle.digest=original;}
 for(const value of [source.buffer,new DataView(source.buffer),null,[],undefined,'{}'])assert.throws(()=>parseWorkbookDesignFile(value));
 for(const call of [()=>validateWorkbookDesignDraft(workbookDesignDraft(),1),()=>createWorkbookDesignFile(workbookDesignDraft()),()=>createWorkbookDesignFile(workbookDesignDraft(),null,1),()=>serializeWorkbookDesignFile(file,1),()=>parseWorkbookDesignFile(source,1)])assert.throws(call,/argument count/);
 await assert.rejects(prepareWorkbookDesignFile(file,1),/argument count/);
});
test('all-five full/selected generated part maps stay exact across design/logo file roundtrips without course or assessment material',async()=>{
 let project=fixture();project=editDeliveryPlan(project,{kind:'create',name:'Selected synthetic',targetMinutes:10,groups:[{moduleId:project.modules[0].id,activityIds:[project.modules[0].activities[0].id]}]}).next;
 const projectBefore=copy(project),handle=await prepareWorkbookLogo(syntheticLogo()),stamp='2026-10-28T20:00:00.000Z';
 const config=(d,asset,planId)=>({planId,familyId:d.familyId,baseLanguage:d.baseLanguage,rawBrand:Object.fromEntries(FAMILY_BRAND_FIELDS.map(k=>[k,rawControl({value:d.brand[k]})])),stamp,frontMatter:{schemaVersion:1,...d.frontMatter},toc:{schemaVersion:1,...d.contents},research:{schemaVersion:1,...d.research},logo:{schemaVersion:1,include:d.logo.include,role:d.logo.role,alt:d.logo.alt,widthIn:Number(d.logo.widthIn),asset}});
 for(const family of WORKBOOK_FAMILIES)for(const planId of ['',project.deliveryPlans.plans[0].id]){const d=workbookDesignDraft(family.id);d.logo.include=true;const file=createWorkbookDesignFile(d,handle),prepared=await prepareWorkbookDesignFile(round(file)),before=prepareFamilyWorkbook(project,config(d,handle,planId)),after=prepareFamilyWorkbook(project,config(prepared.draft,prepared.logoAsset,planId));assert.equal(before.eligible,true);assert.equal(after.eligible,true);assert.deepEqual(after.output.files,before.output.files);assert.equal(JSON.stringify(file).includes(project.id),false);assert.equal(JSON.stringify(file).includes(planId||'not-an-id'),false);assert.equal(JSON.stringify(file).includes('SECRET ANSWER'),false);assert.equal(JSON.stringify(file).includes('assessment-scoring'),false);}
 assert.deepEqual(project,projectBefore);
});
