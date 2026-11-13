import test from 'node:test';
import assert from 'node:assert/strict';
import {createRequire} from 'node:module';
import {deflateSync} from 'node:zlib';
import {syntheticLogo,logoChunk} from './helpers/logo-fixture.mjs';
import {prepareWorkbookLogo,workbookLogoInfo,workbookLogoBytes,workbookLogoBase64,copyWorkbookLogoBytes,validateWorkbookLogo,workbookLogoRequest,workbookLogoPresentation,EMPTY_WORKBOOK_LOGO} from '../src/workbook-logo.mjs';
import {FAMILY_BRAND_FIELDS,prepareFamilyWorkbook,familyWorkbookOutputBlob,familyReviewText} from '../src/workbook-family-flow.mjs';
import {FAMILY_LOGO_PARTS,captureWorkbookLogoOutput} from '../src/workbook-family-output.mjs';
import {WORKBOOK_FAMILIES} from '../src/workbook-presentation.mjs';
import {EMPTY_WORKBOOK_FRONT_MATTER} from '../src/workbook-public-frontmatter.mjs';
import {rawControl} from '../src/activity-draft.mjs';
import {fixture} from './helpers/mock-office.mjs';
import {editDeliveryPlan} from '../src/delivery-plans.mjs';
const require=createRequire(import.meta.url),Zip=require('jszip'),{DOMParser}=require('@xmldom/xmldom');
const stamp='2026-10-28T17:00:00.000Z';
const brand=()=>Object.fromEntries(FAMILY_BRAND_FIELDS.map(k=>[k,rawControl({value:''})]));
const front=()=>({...structuredClone(EMPTY_WORKBOOK_FRONT_MATTER),includeCover:true,cover:{...EMPTY_WORKBOOK_FRONT_MATTER.cover,title:'Synthetic public cover'}});
const cfg=(familyId,logo,extra={})=>({planId:'',familyId,rawBrand:brand(),stamp,frontMatter:front(),...(logo===undefined?{}:{logo}),...extra});
const logo=async(extra={})=>({schemaVersion:1,include:true,asset:await prepareWorkbookLogo(syntheticLogo()),role:'informative',alt:'Original geometric organization mark & "name"\r\nSecond\tline \u4e2d\u6587',widthIn:1.5,...extra});
const raw=v=>({schemaVersion:1,include:{checked:v.include,...rawControl({value:'on'})},role:rawControl({value:v.role}),alt:rawControl({value:v.alt}),widthIn:rawControl({value:String(v.widthIn)}),selection:{state:'ready',generation:1,asset:v.asset,current:true}});

test('logo PNG normalization copies input before await, strips descriptive metadata and privately owns verified pixels',async()=>{
 const original=syntheticLogo({before:[logoChunk('tEXt',new TextEncoder().encode('PRIVATE_FILE_NAME\0/private/path.png'))]}),before=new Uint8Array(original),pending=prepareWorkbookLogo(original);original.fill(0);const handle=await pending,info=workbookLogoInfo(handle),bytes=workbookLogoBytes(handle);
 assert.equal(info.width,4);assert.equal(info.height,2);assert.equal(Object.isFrozen(handle),true);assert.equal(bytes.includes(80),true);assert.equal(new TextDecoder().decode(bytes).includes('PRIVATE_FILE_NAME'),false);assert.notDeepEqual(bytes,before);
 bytes.fill(0);assert.notEqual(workbookLogoBytes(handle)[0],0);assert.equal(workbookLogoBase64(handle),Buffer.from(workbookLogoBytes(handle)).toString('base64'));assert.equal(workbookLogoInfo(handle).sha256,info.sha256);assert.throws(()=>workbookLogoInfo({...handle}));
});

test('logo strict PNG rejects unsupported rendering profiles, malformed ordering, CRC, animation and resource bombs',async()=>{
 const phys=()=>{const x=new Uint8Array(9),v=new DataView(x.buffer);v.setUint32(0,1000);v.setUint32(4,1000);x[8]=1;return logoChunk('pHYs',x);};
 for(const type of ['tRNS','gAMA','cHRM','iCCP','bKGD','acTL','PLTE','uNkN'])await assert.rejects(prepareWorkbookLogo(syntheticLogo({before:[logoChunk(type,new Uint8Array([1]))]})),/Cover logo:/u);
 for(const bytes of [syntheticLogo({width:2049}),syntheticLogo({depth:16}),syntheticLogo({interlace:1}),syntheticLogo({after:[phys()]}),syntheticLogo({before:[phys(),phys()]}),syntheticLogo({scanlines:new Uint8Array(34).fill(5)}),syntheticLogo({idat:new Uint8Array(deflateSync(new Uint8Array(500)))})])await assert.rejects(prepareWorkbookLogo(bytes),/Cover logo:/u);
 const crc=syntheticLogo();crc[29]^=1;await assert.rejects(prepareWorkbookLogo(crc));const trailing=new Uint8Array(syntheticLogo().length+1);trailing.set(syntheticLogo());await assert.rejects(prepareWorkbookLogo(trailing));
 assert.equal(workbookLogoInfo(await prepareWorkbookLogo(syntheticLogo({before:[phys(),logoChunk('sRGB',new Uint8Array([0]))]}))).width,4);
});

test('logo byte boundary rejects proxies shared and shadowed buffers without invoking custom accessors or iterators',async()=>{
 let reads=0;const source=syntheticLogo();Object.defineProperty(source,Symbol.iterator,{get(){reads++;throw Error('iterator read');}});assert.throws(()=>copyWorkbookLogoBytes(source));assert.equal(reads,0);
 const shadow=syntheticLogo();Object.defineProperty(shadow,'buffer',{get(){reads++;throw Error('buffer read');}});assert.throws(()=>copyWorkbookLogoBytes(shadow));assert.equal(reads,0);
 assert.throws(()=>copyWorkbookLogoBytes(new Proxy(syntheticLogo(),{get(){reads++;throw Error('proxy read');}})));assert.equal(reads,0);assert.throws(()=>copyWorkbookLogoBytes(Buffer.from(syntheticLogo())));assert.throws(()=>copyWorkbookLogoBytes(new Uint8Array(new SharedArrayBuffer(8))));
 assert.throws(()=>copyWorkbookLogoBytes(source,1));assert.throws(()=>workbookLogoInfo(null));await assert.rejects(prepareWorkbookLogo(new Uint8Array(2*1024*1024+1)));
});

test('logo active/inactive own-data policies preserve drafts and require explicit informative description and public cover',async()=>{
 const v=await logo();assert.equal(validateWorkbookLogo(v).alt,'Original geometric organization mark & "name"\nSecond\tline \u4e2d\u6587');assert.equal(workbookLogoRequest(raw(v)).asset,v.asset);
 for(const change of [{role:'decorative'},{role:''},{alt:'  '},{widthIn:null},{widthIn:3.1},{asset:{...v.asset}},{extra:true},{schemaVersion:2}])assert.throws(()=>validateWorkbookLogo({...v,...change}));
 const off={...v,include:false,role:'decorative',widthIn:null};assert.equal(validateWorkbookLogo(off).alt,v.alt.replace(/\r\n/gu,'\n'));assert.throws(()=>validateWorkbookLogo({...off,asset:{}}));
 const r=raw(v);r.widthIn.value='1e0';assert.throws(()=>workbookLogoRequest(r));r.include.checked=false;assert.equal(workbookLogoRequest(r).widthIn,null);r.selection.asset={};assert.throws(()=>workbookLogoRequest(r));
 let reads=0;const getter={...v};Object.defineProperty(getter,'alt',{enumerable:true,get(){reads++;return 'bad';}});assert.throws(()=>validateWorkbookLogo(getter));assert.equal(reads,0);
 assert.throws(()=>workbookLogoPresentation(v,{coverIncluded:false,pageWidthTwips:12240,marginTwips:1440}));const tall=await logo({asset:await prepareWorkbookLogo(syntheticLogo({width:2,height:8})),widthIn:3});const fit=workbookLogoPresentation(tall,{coverIncluded:true,pageWidthTwips:12240,marginTwips:1440});assert.equal(fit.heightIn,2);assert.equal(fit.widthIn,.5);
});

test('all five full/selected logo outputs preserve learner parts and use one informative inline image with exact geometry and descriptions',async()=>{
 let p=fixture();p=editDeliveryPlan(p,{kind:'create',name:'Selected',targetMinutes:10,groups:[{moduleId:p.modules[0].id,activityIds:[p.modules[0].activities[0].id]}]}).next;const v=await logo();
 for(const family of WORKBOOK_FAMILIES)for(const planId of ['',p.deliveryPlans.plans[0].id]){
  const base=prepareFamilyWorkbook(p,cfg(family.id,undefined,{planId})),off=prepareFamilyWorkbook(p,cfg(family.id,{...EMPTY_WORKBOOK_LOGO,asset:v.asset},{planId})),active=prepareFamilyWorkbook(p,cfg(family.id,v,{planId}));
  assert.deepEqual(off,base);assert.equal(active.output.profileVersion,4);assert.equal(active.output.assetProfileVersion,1);assert.deepEqual(Object.keys(active.output.files).sort(),[...FAMILY_LOGO_PARTS].sort());
  const xml=active.output.files['word/document.xml'],d=new DOMParser().parseFromString(xml,'application/xml');const props=d.getElementsByTagName('wp:docPr');assert.equal(props.length,1);assert.equal(props[0].getAttribute('descr'),v.alt.replace(/\r\n/gu,'\n'));assert.equal(d.getElementsByTagName('wp:anchor').length,0);assert.equal(d.getElementsByTagName('a:blip')[0].getAttribute('r:embed'),'rIdLogo');assert.equal(d.getElementsByTagName('wp:extent')[0].getAttribute('cx'),String(active.output.preview.logo.cx));
  for(const name of ['word/styles.xml','word/settings.xml','docProps/core.xml','_rels/.rels'])assert.equal(active.output.files[name],base.output.files[name]);assert.match(familyReviewText(active),/profile 4/u);assert.match(familyReviewText(active),/pixels retain their original colors/u);
  const b=await familyWorkbookOutputBlob(active.output,Zip),zip=await Zip.loadAsync(await b.arrayBuffer(),{checkCRC32:true});assert.deepEqual(Object.keys(zip.files).sort(),[...FAMILY_LOGO_PARTS].sort());assert.deepEqual(new Uint8Array(await zip.file('word/media/organization-logo.png').async('uint8array')),workbookLogoBytes(v.asset));assert.equal(xml.includes('facilitator'),false);
 }
});

test('sealed output refuses clones forged previews and maps while captures return independent binary copies',async()=>{
 const v=await logo(),out=prepareFamilyWorkbook(fixture(),cfg('clearline',v)).output,capture=captureWorkbookLogoOutput(out);capture.media.fill(0);assert.notEqual(captureWorkbookLogoOutput(out).media[0],0);
 for(const fake of [{...out},{...out,preview:{...out.preview,logo:{...out.preview.logo,alt:'WRONG'}}},{...out,files:{...out.files,'word/document.xml':'<changed/>'}},Object.create(out),null])await assert.rejects(familyWorkbookOutputBlob(fake,Zip));
 assert.throws(()=>captureWorkbookLogoOutput({...out}));await assert.rejects(familyWorkbookOutputBlob({...out,profileVersion:3},Zip));
});

test('logo ZIP async snapshots defeat retained generator-buffer mutation and refuse altered member bytes',async()=>{
 const out=prepareFamilyWorkbook(fixture(),cfg('clearline',await logo())).output,generate=Zip.prototype.generateAsync,load=Zip.loadAsync;let retained;
 Zip.prototype.generateAsync=async function(...args){retained=await generate.apply(this,args);return retained;};Zip.loadAsync=async function(bytes,...args){retained.fill(0);return load.call(this,bytes,...args);};
 try{const blob=await familyWorkbookOutputBlob(out,Zip),z=await load.call(Zip,await blob.arrayBuffer(),{checkCRC32:true});assert.equal(Object.keys(z.files).length,8);}finally{Zip.prototype.generateAsync=generate;Zip.loadAsync=load;}
 Zip.prototype.generateAsync=async function(...args){const bytes=await generate.apply(this,args),z=await load.call(Zip,bytes);z.file('word/media/organization-logo.png',new Uint8Array([1,2,3]));return generate.apply(z,args);};try{await assert.rejects(familyWorkbookOutputBlob(out,Zip));}finally{Zip.prototype.generateAsync=generate;}
});
