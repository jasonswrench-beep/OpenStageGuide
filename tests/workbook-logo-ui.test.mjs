import test from 'node:test';
import assert from 'node:assert/strict';
import {createRequire} from 'node:module';
import {fixture,makeHost} from './helpers/mock-office.mjs';
import {dom,settle} from './helpers/app-dom.mjs';
import {syntheticLogo} from './helpers/logo-fixture.mjs';
import {WORKBOOK_FAMILIES} from '../src/workbook-presentation.mjs';
const require=createRequire(import.meta.url),Zip=require('jszip');let sequence=0;
async function app(){
 const p=fixture(),h=makeHost(p),ui=dom(),n=id=>ui.nodes.get(id);let fetches=0,reads=0,writes=0;
 Object.assign(globalThis,{Office:h.Office,Word:h.Word,document:ui.document,location:{search:''},JSZip:Zip,fetch:async()=>{fetches++;throw Error('No logo network');},localStorage:{getItem(){reads++;return null;},setItem(){writes++;throw Error('No logo persistence');}}});
 await import(`../src/app.mjs?logo-ui=${++sequence}`);await settle();const go=async name=>{await ui.routes.find(x=>x.dataset.view===name).onclick();await settle();};await go('output');
 return {p,h,ui,n,go,l:id=>n('family-logo-'+id),f:id=>n('family-front-'+id),calls:()=>({fetches,reads,writes,saves:h.saves})};
}
const click=async(a,id)=>{await a.n(id).onclick();await settle();};
const set=(a,k,v)=>{a.l(k).value=v;a.l(k).oninput?.();};
async function select(a,bytes=syntheticLogo(),name='PRIVATE_LOCAL_FILENAME.png'){a.l('file').files=[new File([bytes],name,{type:'image/png'})];a.l('file').value='C:\\fakepath\\'+name;await a.l('file').onchange();await settle();}
function fields(a){a.f('include-cover').checked=true;a.f('include-cover').onchange();a.f('cover-title').value='Public synthetic workbook';a.f('cover-title').oninput();set(a,'role','informative');set(a,'alt','Original geometric organization mark');a.l('include').checked=true;a.l('include').onchange();}
async function review(a){await click(a,'family-preview');assert.equal(a.n('error').hidden,true,a.n('error').textContent);assert.equal(a.n('family-reviewed').disabled,false,a.n('family-brand-status').textContent);a.n('family-reviewed').checked=true;a.n('family-reviewed').onchange();assert.equal(a.n('family-apply').disabled,false,a.n('family-brand-status').textContent);}
async function offers(fn){const blobs=[],names=[],oldURL=URL.createObjectURL,oldCreate=document.createElement,oldTimer=globalThis.setTimeout;URL.createObjectURL=b=>{blobs.push(b);return 'blob:synthetic';};document.createElement=tag=>{const e=oldCreate(tag);if(tag==='a')e.click=()=>names.push(e.download);return e;};globalThis.setTimeout=(fn,ms,...args)=>ms===60000?0:oldTimer(fn,ms,...args);try{await fn();}finally{URL.createObjectURL=oldURL;document.createElement=oldCreate;globalThis.setTimeout=oldTimer;}return {blobs,names};}

test('session logo defaults off and reviewed local inclusion produces independent private eight-part copies without course/preset writes',async()=>{
 const a=await app(),before=a.calls(),source=structuredClone(a.p);assert.equal(a.l('include').checked,false);await select(a);fields(a);await review(a);assert.match(a.n('family-preview-content').textContent,/informative PNG/u);assert.doesNotMatch(a.n('family-preview-content').textContent,/PRIVATE_LOCAL_FILENAME|fakepath/u);
 const result=await offers(async()=>{await click(a,'family-apply');await click(a,'family-apply');});assert.equal(result.blobs.length,2,a.n('error').textContent);assert.notEqual(result.names[0],result.names[1]);for(const blob of result.blobs){const z=await Zip.loadAsync(await blob.arrayBuffer(),{checkCRC32:true});assert.equal(Object.keys(z.files).length,8);const xml=await z.file('word/document.xml').async('string');assert.match(xml,/Original geometric organization mark/u);assert.doesNotMatch(xml,/PRIVATE_LOCAL_FILENAME|fakepath/u);}
 assert.deepEqual(a.p,source);assert.deepEqual(a.calls(),before);assert.equal(a.n('family-reviewed').checked,true);
});

test('all25 family switches preserve local asset, authored description, branding and unrelated drafts',async()=>{
 const a=await app();await select(a);fields(a);a.n('activity-title').value='UNRELATED RAW';a.n('family-brand-organizationName').value='Public organization';a.n('family-brand-organizationName').oninput();
 for(const from of WORKBOOK_FAMILIES)for(const to of WORKBOOK_FAMILIES){a.n('family-template').value=from.id;a.n('family-template').onchange();await review(a);a.n('family-template').value=to.id;a.n('family-template').onchange();assert.equal(a.n('family-reviewed').checked,false);await review(a);assert.equal(a.l('alt').value,'Original geometric organization mark');assert.equal(a.l('image').hidden,false);}
 assert.equal(a.n('activity-title').value,'UNRELATED RAW');assert.equal(a.n('family-brand-organizationName').value,'Public organization');assert.equal(a.h.saves,0);
});

test('logo reset Keep and canceled file picker retain current review; explicit remove/reset protect other drafts',async()=>{
 const a=await app();await select(a);fields(a);await review(a);const text=a.n('family-preview-content').textContent;
 a.l('reset').onclick();assert.equal(a.l('confirm').hidden,false);assert.equal(a.n('family-apply').disabled,true);a.l('keep').onclick();assert.equal(a.n('family-reviewed').checked,true);assert.equal(a.n('family-apply').disabled,false);assert.equal(a.n('family-preview-content').textContent,text);
 a.l('file').files=[];a.l('file').value='';await a.l('file').onchange();assert.equal(a.n('family-reviewed').checked,true);assert.equal((await offers(()=>click(a,'family-apply'))).blobs.length,1);
 a.l('remove').onclick();a.l('replace').onclick();assert.equal(a.l('include').checked,false);assert.equal(a.l('image').hidden,true);assert.equal(a.l('alt').value,'Original geometric organization mark');assert.equal(a.f('cover-title').value,'Public synthetic workbook');
 a.l('reset').onclick();a.l('replace').onclick();assert.equal(a.l('alt').value,'');assert.equal(a.l('width').value,'1.5');assert.equal(a.f('include-cover').checked,true);
});

test('failed replacement makes old image unavailable until explicit Restore and valid replacement needs affirmative review',async()=>{
 const a=await app();await select(a);fields(a);await review(a);const first=a.l('image').attributes.src;
 await select(a,new Uint8Array([1,2,3]),'BAD.png');assert.equal(a.l('image').hidden,true);assert.equal(a.l('error').hidden,false);await click(a,'family-preview');assert.match(a.n('family-brand-status').textContent,/not ready/u);assert.equal(a.l('alt').value,'Original geometric organization mark');assert.equal((await offers(()=>click(a,'family-apply'))).blobs.length,0);
 a.l('restore').onclick();a.l('replace').onclick();assert.equal(a.l('image').attributes.src,first);await review(a);
 await select(a,syntheticLogo({width:8,height:2}),'NEW.png');assert.equal(a.l('confirm').hidden,false);assert.equal(a.l('image').hidden,true);a.l('replace').onclick();assert.notEqual(a.l('image').attributes.src,first);assert.equal(a.l('alt').value,'Original geometric organization mark');await review(a);
});

test('active logo requires public cover informative role complete width and authored alt; inactive invalid width draft is preserved',async()=>{
 const a=await app();await select(a);fields(a);
 for(const mutate of [()=>{a.f('include-cover').checked=false;a.f('include-cover').onchange();},()=>set(a,'role','decorative'),()=>set(a,'alt',' '),()=>set(a,'width','1e0')]){fields(a);set(a,'width','1.5');mutate();await click(a,'family-preview');assert.equal(a.n('family-reviewed').disabled,true);}
 a.l('include').checked=false;a.l('include').onchange();await review(a);assert.equal(a.l('width').value,'1e0');
});

test('image source description dimensions visibility and async consent/source races offer no stale logo file',async()=>{
 for(const mutate of [a=>a.l('image').setAttribute('src','data:image/png;base64,WRONG'),a=>a.l('image').setAttribute('alt','WRONG'),a=>a.l('image').setAttribute('width','999'),a=>a.l('image').hidden=true,a=>{a.l('image').style={cssText:'display:none'};},a=>{a.l('include').checked=false;a.l('include').onchange();a.l('include').checked=true;a.l('include').onchange();},a=>{a.n('family-reviewed').checked=false;a.n('family-reviewed').onchange();a.n('family-reviewed').checked=true;a.n('family-reviewed').onchange();}]){
  const a=await app();await select(a);fields(a);await review(a);const generate=Zip.prototype.generateAsync;Zip.prototype.generateAsync=async function(...args){const bytes=await generate.apply(this,args);mutate(a);return bytes;};try{assert.equal((await offers(()=>click(a,'family-apply'))).blobs.length,0);}finally{Zip.prototype.generateAsync=generate;}assert.equal(a.l('alt').value,'Original geometric organization mark');assert.equal(a.h.saves,0);
 }
});

test('hidden controls and stale asynchronous local reads never install or export an image and preserve drafts',async()=>{
 const a=await app();fields(a);await a.go('course');await select(a);assert.equal(a.l('image').hidden,true);assert.equal(a.l('error').hidden,false);await a.go('output');
 let finish;const bytes=syntheticLogo();a.l('file').files=[{size:bytes.length,stream(){let read=0;return {getReader(){return {read:()=>++read===1?new Promise(resolve=>finish=()=>resolve({done:false,value:bytes})):Promise.resolve({done:true}),cancel(){return Promise.resolve();},releaseLock(){}};}};}}];a.l('file').value='LOCAL.png';const pending=a.l('file').onchange();await settle();assert.equal(a.n('family-preview').disabled,true);set(a,'alt','NEW RAW DURING READ');finish();await pending;assert.equal(a.l('image').hidden,true);assert.equal(a.l('error').hidden,false);assert.equal(a.l('alt').value,'NEW RAW DURING READ');assert.equal(a.f('cover-title').value,'Public synthetic workbook');
});

test('bounded File streams copy retained chunks immediately and cancel declared-size overruns without fallback',async()=>{
 const a=await app(),bytes=syntheticLogo(),original=new Uint8Array(bytes),half=Math.floor(bytes.length/2),first=bytes.slice(0,half),last=bytes.slice(half);let read=0,canceled=0;
 a.l('file').files=[{size:bytes.length,stream(){return {getReader(){return {read(){read++;if(read===1)return Promise.resolve({done:false,value:first});if(read===2){first.fill(0);return Promise.resolve({done:false,value:last});}last.fill(0);return Promise.resolve({done:true});},cancel(){canceled++;return Promise.resolve();},releaseLock(){}};}};}}];a.l('file').value='PRIVATE.png';await a.l('file').onchange();assert.equal(a.l('error').hidden,true,a.l('error').textContent);assert.equal(a.l('image').hidden,false);assert.equal(canceled,0);fields(a);await review(a);assert.equal((await offers(()=>click(a,'family-apply'))).blobs.length,1);
 a.l('file').files=[{size:1,stream(){return {getReader(){return {read:()=>Promise.resolve({done:false,value:original}),cancel(){canceled++;return Promise.resolve();},releaseLock(){}};}};}}];a.l('file').value='LIES.png';await a.l('file').onchange();assert.equal(canceled,1);assert.equal(a.l('image').hidden,true);assert.match(a.l('error').textContent,/exceeded/u);assert.equal(a.l('alt').value,'Original geometric organization mark');
});

test('replacement confirmation tampering and hidden controls cannot destroy a current logo draft or overlap publication dialogs',async()=>{
 const a=await app();await select(a);fields(a);await review(a);const src=a.l('image').getAttribute('src');
 a.l('reset').onclick();a.l('confirm-text').textContent='CHANGED';a.l('replace').onclick();assert.equal(a.l('image').getAttribute('src'),src);assert.equal(a.l('alt').value,'Original geometric organization mark');assert.equal(a.l('confirm').hidden,true);
 a.l('reset').onclick();a.l('alt').value='EVENTLESS DRAFT';a.l('replace').onclick();assert.equal(a.l('alt').value,'EVENTLESS DRAFT');assert.equal(a.l('image').getAttribute('src'),src);
 a.f('reset').onclick();a.l('reset').onclick();assert.equal(a.l('confirm').hidden,true);assert.equal(a.f('confirm').hidden,false);a.f('keep').onclick();a.l('reset').onclick();a.f('reset').onclick();assert.equal(a.f('confirm').hidden,true);assert.equal(a.l('confirm').hidden,false);a.l('confirm').onkeydown({key:'Escape',preventDefault(){}});assert.equal(a.l('confirm').hidden,true);
 a.l('controls').hidden=true;a.l('remove').onclick();assert.equal(a.l('confirm').hidden,true);assert.equal(a.l('alt').value,'EVENTLESS DRAFT');assert.equal(a.h.saves,0);
});
