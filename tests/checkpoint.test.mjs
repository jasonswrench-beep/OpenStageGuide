import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import {createCheckpoint,readCheckpoint,inspectMaster,sha256,checkpointName,bindSavedCheckpoint,verifyPreparedCheckpointCapture} from '../src/checkpoint.mjs';
import {readCompressedDocument} from '../src/office-file.mjs';
import {createDocumentStore} from '../src/document-store.mjs';
import {makeHost,fixture,records} from './helpers/mock-office.mjs';
import {masterFixture,installParser,assets} from './helpers/checkpoint-fixture.mjs';
import {xml} from '../src/worksheet.mjs';
installParser();
test('saved-checkpoint binding verifies exact captured master bytes without claiming native document identity',async()=>{
  const {master,project}=await masterFixture(),backup=await createCheckpoint(master,project,JSZip,assets),before=await sha256(master),backupBefore=await sha256(backup.data);
  const proof=await bindSavedCheckpoint(backup.data,master,project,JSZip);assert.equal(proof.checkpointId,backup.manifest.id);assert.equal(proof.projectId,project.id);assert.equal(proof.revision,project.revision);assert.equal(proof.masterHash,before);assert.equal(proof.backupHash,backupBefore);assert.equal(proof.nativeDocumentVerified,false);assert.ok(Object.isFrozen(proof));assert.equal(await sha256(master),before);assert.equal(await sha256(backup.data),backupBefore);
});
test('checkpoint binding refuses another guide/revision and changes outside managed course text',async()=>{
  const {master,project}=await masterFixture(),backup=await createCheckpoint(master,project,JSZip,assets);
  for(const changed of [{...project,id:crypto.randomUUID()},{...project,revision:project.revision+1}])await assert.rejects(bindSavedCheckpoint(backup.data,master,changed,JSZip),e=>e.code==='OSG_CHECKPOINT');
  for(const part of ['word/document.xml','word/media/image1.png','customXml/item1.xml','word/settings.xml']){
    const zip=await JSZip.loadAsync(master),file=zip.file(part);if(part.endsWith('.png'))zip.file(part,new Uint8Array([137,80,78,71,9,8,7]));else zip.file(part,(await file.async('string')).replace('UNTAGGED MANUAL NOTE','NEW MANUAL NOTE')+' ');
    const different=await zip.generateAsync({type:'uint8array'});await assert.rejects(bindSavedCheckpoint(backup.data,different,project,JSZip),/current master differs/);
  }
});
test('checkpoint binding never trusts an earlier verification flag or a tampered saved package',async()=>{
  const {master,project}=await masterFixture(),backup=await createCheckpoint(master,project,JSZip,assets),zip=await JSZip.loadAsync(backup.data);zip.file('master.docx',master.slice(0,-1));
  await assert.rejects(bindSavedCheckpoint(await zip.generateAsync({type:'uint8array'}),master,project,JSZip),e=>e.code==='OSG_CHECKPOINT');
  await assert.rejects(bindSavedCheckpoint(new Uint8Array(),master,project,JSZip),e=>e.code==='OSG_CHECKPOINT');
});
test('checkpoint binding snapshots inputs and refuses asynchronous byte or project changes',async()=>{
  for(const target of ['master','backup','project']){
    const {master,project}=await masterFixture(),backup=await createCheckpoint(master,project,JSZip,assets);let changed=false;
    class DelayedZip extends JSZip{static async loadAsync(...args){const zip=await JSZip.loadAsync(...args);if(!changed){changed=true;if(target==='project')project.revision++;else if(target==='master')master[master.length-1]^=1;else backup.data[backup.data.length-1]^=1;}return zip;}}
    await assert.rejects(bindSavedCheckpoint(backup.data,master,project,DelayedZip),/inputs changed/);
  }
});
test('checkpoint binding uses plain non-aliasing copies for Buffer and Uint8Array subclasses',async()=>{
  const {master,project}=await masterFixture(),backup=await createCheckpoint(master,project,JSZip,assets);let slices=0;
  class SpecialBytes extends Uint8Array{slice(){slices++;throw new Error('Caller slice must not be invoked');}}
  for(const convert of [data=>Buffer.from(data),data=>new SpecialBytes(data)]){
    const b=convert(backup.data),m=convert(master),bHash=await sha256(b),mHash=await sha256(m);b.slice=()=>{slices++;return b;};m.slice=()=>{slices++;return m;};
    const proof=await bindSavedCheckpoint(b,m,project,JSZip);assert.equal(proof.masterHash,mHash);assert.equal(proof.backupHash,bHash);assert.equal(await sha256(b),bHash);assert.equal(await sha256(m),mHash);
  }
  assert.equal(slices,0);
});
test('checkpoint binding rejects project/backup/master mutations during the final digest resolutions',async()=>{
  const original=crypto.subtle.digest;let calls=0;
  try{
    crypto.subtle.digest=async function(...args){calls++;return Reflect.apply(original,this,args);};
    const fixture=await masterFixture(),saved=await createCheckpoint(fixture.master,fixture.project,JSZip,assets);calls=0;await bindSavedCheckpoint(saved.data,fixture.master,fixture.project,JSZip);const total=calls;assert.ok(total>2);
    for(const boundary of [total-1,total])for(const target of ['project','backup','master']){
      crypto.subtle.digest=original;const {master,project}=await masterFixture(),backup=await createCheckpoint(master,project,JSZip,assets);calls=0;
      crypto.subtle.digest=async function(...args){const index=++calls,result=await Reflect.apply(original,this,args);if(index===boundary){if(target==='project')project.unknownLateField={private:'Changed during final digest'};else if(target==='backup')backup.data[backup.data.length-1]^=1;else master[master.length-1]^=1;}return result;};
      await assert.rejects(bindSavedCheckpoint(backup.data,master,project,JSZip),/inputs changed/);
    }
  }finally{crypto.subtle.digest=original;}
});

test('checkpoint accepts a completed settings receipt and rejects an incomplete edit or mismatched timing cue',async()=>{
  const {project,master}=await masterFixture(),id=project.modules[0].activities[0].id;
  const zip=await JSZip.loadAsync(master),path='word/webextensions/webextension1.xml';
  const receipt={schema:3,txId:'settings-check',phase:'committed',kind:'activity-settings',project,settings:{id,after:{minutes:10,lines:2}}};
  const baseExt=await zip.file(path).async('string'),baseDoc=await zip.file('word/document.xml').async('string');
  zip.file(path,baseExt.replace('</we:properties>',`<we:property name="openstage.guide.pending.v1" value="${xml(JSON.stringify(receipt))}"/></we:properties>`));
  const anchor=`<w:sdt><w:sdtPr><w:tag w:val="osg:activity:${id}"`;
  const withCue=baseDoc.replace(anchor,`<w:p><w:r><w:t>Learning activity · 10 min</w:t></w:r></w:p>${anchor}`);zip.file('word/document.xml',withCue);
  const good=await zip.generateAsync({type:'uint8array'}),backup=await createCheckpoint(good,project,JSZip,assets);
  assert.deepEqual((await readCheckpoint(backup.data,JSZip)).master,good);
  zip.file('word/document.xml',withCue.replace('Learning activity · 10 min','Learning activity · 99 min'));
  await assert.rejects(inspectMaster(await zip.generateAsync({type:'uint8array'}),JSZip),/timing cue/);
  receipt.phase='prepared';zip.file(path,baseExt.replace('</we:properties>',`<we:property name="openstage.guide.pending.v1" value="${xml(JSON.stringify(receipt))}"/></we:properties>`));
  await assert.rejects(inspectMaster(await zip.generateAsync({type:'uint8array'}),JSZip),/unresolved/);
});
test('full-master backup restores identical bytes, including formatting, unknown metadata, settings and image parts',async()=>{
  const {master,project}=await masterFixture();project.modules[0].activities[0].participant+='\nSecond learner paragraph';
  const multiline=await masterFixture(project),backup=await createCheckpoint(multiline.master,project,JSZip,assets),checked=await readCheckpoint(backup.data,JSZip);
  assert.deepEqual(checked.master,multiline.master);assert.deepEqual(checked.project,project);assert.equal(checked.assetCount,15);
  assert.equal(checked.masterHash,await sha256(multiline.master));assert.equal(backup.manifest.files.length,16);
  const restored=await JSZip.loadAsync(checked.master);
  assert.equal(await restored.file('customXml/item1.xml').async('string'),'<unknown>KEEP UNKNOWN METADATA</unknown>');
  assert.match(await restored.file('word/document.xml').async('string'),/UNTAGGED MANUAL NOTE/);
  assert.match(await restored.file('word/webextensions/webextension1.xml').async('string'),/KEEP UNKNOWN DATA/);
  assert.match(await restored.file('word/document.xml').async('string'),/SECRET ANSWER/);
  assert.notEqual(await sha256(master),checked.masterHash);
});
for(const defect of ['master','asset','missing','extra','version','course','duplicate-list','unsafe-path'])test(`backup verification rejects ${defect} before restore`,async()=>{
  const {master,project}=await masterFixture(),backup=await createCheckpoint(master,project,JSZip,assets),zip=await JSZip.loadAsync(backup.data),manifest=JSON.parse(await zip.file('manifest.json').async('string'));
  if(defect==='master')zip.file('master.docx',new Uint8Array(master.length).fill(1));
  if(defect==='asset')zip.file('assets/icons/Assessment.png',new Uint8Array([0,0,0,0,0,0]));
  if(defect==='missing')zip.remove('assets/icons/Assessment.png');
  if(defect==='extra')zip.file('unexpected.txt','x');
  if(defect==='version')manifest.version=99;
  if(defect==='course')manifest.course.revision++;
  if(defect==='duplicate-list')manifest.files[1]=manifest.files[0];
  if(defect==='unsafe-path')zip.file('../escape.txt','x');
  zip.file('manifest.json',JSON.stringify(manifest));
  await assert.rejects(readCheckpoint(await zip.generateAsync({type:'uint8array'}),JSZip));
});
test('backup refuses missing assets, unsupported master and unsynchronized course text',async()=>{
  const {master,project}=await masterFixture();
  await assert.rejects(createCheckpoint(master,project,JSZip,async()=>({ok:false})),/asset/);
  await assert.rejects(inspectMaster(new Uint8Array([1,2,3]),JSZip));
  const changed=structuredClone(project);changed.modules[0].activities[0].participant='UNSAVED METADATA';
  await assert.rejects(createCheckpoint(master,changed,JSZip,assets),/differs/);
  const zip=await JSZip.loadAsync(master);zip.file('word/document.xml',(await zip.file('word/document.xml').async('string')).replace('Learner task.','Unreconciled manual edit'));
  await assert.rejects(inspectMaster(await zip.generateAsync({type:'uint8array'}),JSZip),/text and course records differ/);
});
function fileHost(data,{failSlice=-1,short=false,closeFail=false}={}){
  let closed=0;const file={size:data.length,sliceCount:Math.ceil(data.length/65536),getSliceAsync(i,cb){if(i===failSlice)return cb({status:'failed',error:{message:'Slice failure'}});cb({status:'succeeded',value:{index:i,data:Array.from(data.slice(i*65536,(i+1)*65536-(short?1:0)))}});},closeAsync(cb){closed++;cb({status:closeFail?'failed':'succeeded'});}};
  return {document:{getFileAsync(type,options,cb){assert.equal(type,'compressed');assert.equal(options.sliceSize,65536);cb({status:'succeeded',value:file});}},closed:()=>closed,file};
}
test('Office file reader joins all slices exactly and releases the file handle',async()=>{
  const data=Uint8Array.from({length:150000},(_,i)=>i%256),h=fileHost(data);
  assert.deepEqual(await readCompressedDocument(h.document),data);assert.equal(h.closed(),1);
});
for(const options of [{failSlice:1},{short:true},{closeFail:true}])test(`Office file read failure never offers partial data and closes its handle ${JSON.stringify(options)}`,async()=>{
  const h=fileHost(new Uint8Array(150000),options);await assert.rejects(readCompressedDocument(h.document));assert.equal(h.closed(),1);
});
test('capture serializes against mutations, reconciles direct edits, and rejects text changed during capture',async()=>{
  const p=fixture(),h=makeHost(p),store=createDocumentStore({settings:()=>h.settings,scan:async()=>h.records,write:async()=>{throw new Error('must not write');}});
  h.records.find(r=>r.tag.startsWith('osg:participant:')).text='Manual edit before checkpoint';
  let release,entered;const gate=new Promise(ok=>release=ok),start=new Promise(ok=>entered=ok);
  const capture=store.captureMaster(async()=>{entered();await gate;return new Uint8Array([1]);});await start;
  await assert.rejects(store.readProject(),/already running/);release();const result=await capture;
  assert.equal(result.project.revision,p.revision+1);assert.equal(result.project.modules[0].activities[0].participant,'Manual edit before checkpoint');
  await assert.rejects(store.captureMaster(async()=>{h.records[0].text='Changed during capture';return new Uint8Array([1]);}),/changed/);
});
test('a file-read failure does not quarantine an otherwise healthy course',async()=>{
  const p=fixture(),h=makeHost(p),store=createDocumentStore({settings:()=>h.settings,scan:async()=>h.records,write:async()=>{}});
  await assert.rejects(store.captureMaster(async()=>{throw Object.assign(new Error('File unavailable'),{code:'OSG_CHECKPOINT'});}));
  assert.equal(store.isBlocked(),false);assert.deepEqual(await store.readProject(),p);
  assert.deepEqual(h.records,records(p));
  const first=checkpointName(p.title,true),second=checkpointName(p.title,true);assert.notEqual(first,second);assert.match(first,/Restored .*\.docx$/);
});

const extensionPath='word/webextensions/webextension1.xml';
const pendingProperty=journal=>`<we:property name="openstage.guide.pending.v1" value="${xml(JSON.stringify(journal))}"/>`;
async function preparedFixture(receipt=false){
  let {master,project}=await masterFixture();
  if(receipt){const zip=await JSZip.loadAsync(master);zip.file(extensionPath,(await zip.file(extensionPath).async('string')).replace('</we:properties>',pendingProperty({schema:2,phase:'committed',project})+'</we:properties>'));master=await zip.generateAsync({type:'uint8array'});}
  const backup=await createCheckpoint(master,project,JSZip,assets),binding=await bindSavedCheckpoint(backup.data,master,project,JSZip),next=structuredClone(project);next.revision++;
  const journal={schema:4,kind:'activity-icon',phase:'prepared',txId:crypto.randomUUID(),base:project,project:next,checkpoint:{id:binding.checkpointId,masterHash:binding.masterHash,backupHash:binding.backupHash}};
  const zip=await JSZip.loadAsync(master),doc=new DOMParser().parseFromString(await zip.file(extensionPath).async('string'),'application/xml'),props=doc.getElementsByTagNameNS('http://schemas.microsoft.com/office/webextensions/webextension/2010/11','properties')[0];
  for(const p of Array.from(props.childNodes))if(p.nodeType===1&&p.getAttribute('name')==='openstage.guide.pending.v1')props.removeChild(p);
  const added=new DOMParser().parseFromString(`<we:webextension xmlns:we="http://schemas.microsoft.com/office/webextensions/webextension/2010/11">${pendingProperty(journal)}</we:webextension>`,'application/xml').documentElement.firstChild;
  props.appendChild(doc.importNode(added,true));zip.file(extensionPath,doc.toString());const prepared=await zip.generateAsync({type:'uint8array'});
  return {master,project,backup,binding,journal,prepared};
}
test('prepared-capture bridge accepts only a session-bound exact storage delta and makes no native or transition claim',async()=>{
  for(const receipt of [false,true]){const f=await preparedFixture(receipt),before=[await sha256(f.master),await sha256(f.prepared),await sha256(f.backup.data)],proof=await verifyPreparedCheckpointCapture(f.binding,f.prepared,f.journal,JSZip);assert.equal(proof.nativeDocumentVerified,false);assert.equal(proof.txId,f.journal.txId);assert.ok(Object.isFrozen(proof));assert.match(proof.scope,/structurally/);assert.deepEqual([await sha256(f.master),await sha256(f.prepared),await sha256(f.backup.data)],before);
    for(const fake of [{...f.binding},JSON.parse(JSON.stringify(f.binding)),null])await assert.rejects(verifyPreparedCheckpointCapture(fake,f.prepared,f.journal,JSZip),/in this session/);
  }
});
test('prepared-capture bridge refuses manual, asset, custom XML, unrelated storage and part-roster changes',async()=>{
  const f=await preparedFixture();
  for(const part of ['word/document.xml','word/media/image1.png','customXml/item1.xml','word/settings.xml','word/styles.xml',extensionPath,'new-part.bin','remove']){
    const zip=await JSZip.loadAsync(f.prepared);
    if(part==='remove')zip.remove('customXml/item1.xml');else if(part==='new-part.bin')zip.file(part,new Uint8Array([1]));else if(part.endsWith('.png'))zip.file(part,new Uint8Array([9]));else {let text=await zip.file(part).async('string');text=part===extensionPath?text.replace('KEEP UNKNOWN DATA','CHANGED UNKNOWN DATA'):text+' ';zip.file(part,text);}
    await assert.rejects(verifyPreparedCheckpointCapture(f.binding,await zip.generateAsync({type:'uint8array'}),f.journal,JSZip),e=>e.code==='OSG_CHECKPOINT',part);
  }
});
test('prepared-capture bridge requires exact pending data and an unambiguous supported storage location',async()=>{
  const f=await preparedFixture();
  for(const defect of ['missing','wrong','duplicate-pending','duplicate-course','moved','attributes','children','prolog','comment']){
    const zip=await JSZip.loadAsync(f.prepared),doc=new DOMParser().parseFromString(await zip.file(extensionPath).async('string'),'application/xml'),properties=Array.from(doc.getElementsByTagNameNS('*','property')),pending=properties.find(p=>p.getAttribute('name')==='openstage.guide.pending.v1'),parent=pending.parentNode;
    if(defect==='missing')parent.removeChild(pending);
    if(defect==='wrong')pending.setAttribute('value',JSON.stringify({...f.journal,txId:crypto.randomUUID()}));
    if(defect==='duplicate-pending')parent.appendChild(pending.cloneNode(true));
    if(defect==='duplicate-course')parent.appendChild(properties.find(p=>p.getAttribute('name')==='openstage.guide.project.v1').cloneNode(true));
    if(defect==='moved')doc.documentElement.appendChild(pending);
    if(defect==='attributes')pending.setAttribute('unknown','must preserve');
    if(defect==='children')pending.appendChild(doc.createComment('unknown pending content'));
    let text=doc.toString();if(defect==='prolog')text='<?xml version="1.0"?>'+text;if(defect==='comment')text='<!--unrelated-->'+text;zip.file(extensionPath,text);
    await assert.rejects(verifyPreparedCheckpointCapture(f.binding,await zip.generateAsync({type:'uint8array'}),f.journal,JSZip),e=>e.code==='OSG_CHECKPOINT',defect);
  }
  for(const change of [j=>j.schema=3,j=>j.kind='activity-settings',j=>j.phase='committed',j=>j.txId='-'.repeat(36),j=>j.base.title='different',j=>j.project.revision++,j=>j.checkpoint.id=crypto.randomUUID()]){const journal=structuredClone(f.journal);change(journal);await assert.rejects(verifyPreparedCheckpointCapture(f.binding,f.prepared,journal,JSZip),e=>e.code==='OSG_CHECKPOINT');}
});
test('prepared-capture bridge seals caller bytes and journal after the final asynchronous digest',async()=>{
  const original=crypto.subtle.digest;
  try{
    const f=await preparedFixture();let total=0;crypto.subtle.digest=async function(...args){total++;return Reflect.apply(original,this,args);};await verifyPreparedCheckpointCapture(f.binding,f.prepared,f.journal,JSZip);assert.ok(total>1);
    for(const target of ['bytes','journal']){crypto.subtle.digest=original;const g=await preparedFixture(),data=Buffer.from(g.prepared);let calls=0;data.slice=()=>{throw new Error('Must not call caller slice');};crypto.subtle.digest=async function(...args){const result=await Reflect.apply(original,this,args);if(++calls===total){if(target==='bytes')data[data.length-1]^=1;else g.journal.unknown='late input change';}return result;};await assert.rejects(verifyPreparedCheckpointCapture(g.binding,data,g.journal,JSZip),/inputs changed/);}
  }finally{crypto.subtle.digest=original;}
});

test('prepared-capture bridge rejects duplicate raw ZIP directory entries before JSZip can merge them',async()=>{
  const f=await preparedFixture(),data=f.prepared,v=new DataView(data.buffer,data.byteOffset,data.byteLength),end=data.length-22,offset=v.getUint32(end+16,true),n=v.getUint16(offset+28,true),extra=v.getUint16(offset+30,true),comment=v.getUint16(offset+32,true),entry=data.slice(offset,offset+46+n+extra+comment),result=new Uint8Array(data.length+entry.length);result.set(data.subarray(0,end));result.set(entry,end);result.set(data.subarray(end),end+entry.length);const e=new DataView(result.buffer),nextEnd=end+entry.length,count=v.getUint16(end+10,true)+1;e.setUint16(nextEnd+8,count,true);e.setUint16(nextEnd+10,count,true);e.setUint32(nextEnd+12,v.getUint32(end+12,true)+entry.length,true);
  // JSZip ordinarily collapses this ambiguous roster to a single path.
  assert.equal(Object.keys((await JSZip.loadAsync(result)).files).length,Object.keys((await JSZip.loadAsync(data)).files).length);
  await assert.rejects(verifyPreparedCheckpointCapture(f.binding,result,f.journal,JSZip),/duplicate captured ZIP paths/);
});

function duplicateDirectoryEntry(data){
  const v=new DataView(data.buffer,data.byteOffset,data.byteLength),end=data.length-22,offset=v.getUint32(end+16,true),length=46+v.getUint16(offset+28,true)+v.getUint16(offset+30,true)+v.getUint16(offset+32,true),entry=data.slice(offset,offset+length),result=new Uint8Array(data.length+length);result.set(data.subarray(0,end));result.set(entry,end);result.set(data.subarray(end),end+length);const e=new DataView(result.buffer),nextEnd=end+length,count=v.getUint16(end+10,true)+1;e.setUint16(nextEnd+8,count,true);e.setUint16(nextEnd+10,count,true);e.setUint32(nextEnd+12,v.getUint32(end+12,true)+length,true);return result;
}
test('master and saved-backup readers refuse raw duplicates, sanitized collisions and local/central filename disagreements',async()=>{
  const f=await preparedFixture();
  await assert.rejects(inspectMaster(duplicateDirectoryEntry(f.master),JSZip),/duplicate captured ZIP paths/);
  await assert.rejects(readCheckpoint(duplicateDirectoryEntry(f.backup.data),JSZip),/duplicate captured ZIP paths/);
  for(const [data,read] of [[f.master,b=>inspectMaster(b,JSZip)],[f.backup.data,b=>readCheckpoint(b,JSZip)],[f.prepared,b=>verifyPreparedCheckpointCapture(f.binding,b,f.journal,JSZip)]]){
    const zip=await JSZip.loadAsync(data);zip.file('shadow/../customXml/item1.xml','HIDDEN CHANGED PAYLOAD');const collision=await zip.generateAsync({type:'uint8array'});assert.ok((await JSZip.loadAsync(collision)).file('customXml/item1.xml'));await assert.rejects(read(collision),/Unsafe or duplicate captured ZIP paths/);
    const mismatch=new Uint8Array(data),v=new DataView(mismatch.buffer);assert.equal(v.getUint32(0,true),0x04034b50);mismatch[30]^=1;await assert.rejects(read(mismatch),/Inconsistent captured ZIP bounds/);
  }
});

test('strict master envelope accepts checked streaming descriptors but refuses unlisted prefixes and inconsistent local CRC',async()=>{
  const {master,project}=await masterFixture(),zip=await JSZip.loadAsync(master),streamed=await zip.generateAsync({type:'uint8array',streamFiles:true,compression:'DEFLATE'});assert.deepEqual((await inspectMaster(streamed,JSZip)).project,project);
  const prefixed=new Uint8Array(master.length+1);prefixed[0]=42;prefixed.set(master,1);const v=new DataView(prefixed.buffer),end=prefixed.length-22,offset=v.getUint32(end+16,true)+1,count=v.getUint16(end+10,true);v.setUint32(end+16,offset,true);let at=offset;for(let i=0;i<count;i++){v.setUint32(at+42,v.getUint32(at+42,true)+1,true);at+=46+v.getUint16(at+28,true)+v.getUint16(at+30,true)+v.getUint16(at+32,true);}await JSZip.loadAsync(prefixed);await assert.rejects(inspectMaster(prefixed,JSZip),/unlisted captured ZIP payloads/);
  const crc=new Uint8Array(master);crc[14]^=1;await assert.rejects(inspectMaster(crc,JSZip),/sizes or CRC/);
});

function firstZipRecord(data,method){
  const v=new DataView(data.buffer,data.byteOffset,data.byteLength),end=data.length-22,count=v.getUint16(end+10,true);let at=v.getUint32(end+16,true);
  for(let i=0;i<count;i++){if(v.getUint16(at+10,true)===method&&v.getUint32(at+24,true)>0)return {central:at,local:v.getUint32(at+42,true)};at+=46+v.getUint16(at+28,true)+v.getUint16(at+30,true)+v.getUint16(at+32,true);}throw new Error('Test fixture needs a matching file');
}
test('descriptor-mode local CRC and sizes require zero placeholders even when nonzero values match central records',async()=>{
  const f=await masterFixture(),zip=await JSZip.loadAsync(f.master),stream=await zip.generateAsync({type:'uint8array',streamFiles:true,compression:'DEFLATE'}),{central,local}=firstZipRecord(stream,8);
  await inspectMaster(stream,JSZip);
  for(const [lo,co] of [[14,16],[18,20],[22,24]])for(const matching of [true,false]){const bad=new Uint8Array(stream),v=new DataView(bad.buffer),actual=v.getUint32(central+co,true);assert.notEqual(actual,0);v.setUint32(local+lo,matching?actual:actual^1,true);await assert.rejects(inspectMaster(bad,JSZip),/local CRC and sizes must be zero/);}
});
test('raw ZIP profile accepts only method-appropriate descriptor UTF-8 and DEFLATE option flags',async()=>{
  const f=await masterFixture();
  for(const method of [0,8]){
    const z=await JSZip.loadAsync(f.master),data=await z.generateAsync({type:'uint8array',compression:method===8?'DEFLATE':'STORE'}),{central,local}=firstZipRecord(data,method),allowed=method===8?0x0806:0x0800;
    for(let bit=0;bit<16;bit++){if(bit===3)continue;const mask=1<<bit,changed=new Uint8Array(data),v=new DataView(changed.buffer);v.setUint16(central+8,mask,true);v.setUint16(local+6,mask,true);if(allowed&mask)await inspectMaster(changed,JSZip);else await assert.rejects(inspectMaster(changed,JSZip),/Unsupported or oversized captured ZIP parts/);}
    if(method===8){const changed=new Uint8Array(data),v=new DataView(changed.buffer);v.setUint16(central+8,0x0806,true);v.setUint16(local+6,0x0806,true);await inspectMaster(changed,JSZip);}
  }
});
