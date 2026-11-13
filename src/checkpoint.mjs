import {ICONS,validate,reconcile,managedTags} from './model.mjs';
import {KEY,PENDING} from './document-store.mjs';
import {MAX_MASTER_BYTES} from './office-file.mjs';
import {moduleAsset} from './diagnostics.mjs';
import {inspectOwnedCuePackage} from './cue-ownership.mjs';
import {cloneIconData} from './icon-transition.mjs';
import {hasFrontMatter,validFrontJournal} from './frontmatter.mjs';
import {inspectFrontMatterPackage,inspectBlankPackage} from './frontmatter-package.mjs';
export const FORMAT='openstage-guide-checkpoint',VERSION=1,MAX_PACKAGE_BYTES=80*1024*1024;
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main',WE='http://schemas.microsoft.com/office/webextensions/webextension/2010/11';
const equal=(a,b)=>JSON.stringify(a)===JSON.stringify(b);
const bytes=x=>x instanceof Uint8Array?x:new Uint8Array(x);
const bindings=new WeakMap(),preparedBindings=new WeakMap();
const exactInput=(input,snapshot)=>{const copy=new Uint8Array(input);if(copy.length!==snapshot.length)return false;for(let i=0;i<copy.length;i++)if(copy[i]!==snapshot[i])return false;return true;};
export async function sha256(data){return Array.from(new Uint8Array(await crypto.subtle.digest('SHA-256',bytes(data))),n=>n.toString(16).padStart(2,'0')).join('');}
function xml(text){
  if(/<!DOCTYPE|<!ENTITY/i.test(text))throw new Error('Unsupported XML declarations in the backup.');
  const doc=new DOMParser().parseFromString(text,'application/xml');
  if(doc.getElementsByTagName('parsererror').length||!doc.documentElement)throw new Error('The backup contains invalid document XML.');
  return doc;
}
function bounded(zip,maxTotal=128*1024*1024){
  let total=0;const files=Object.values(zip.files).filter(f=>!f.dir);
  if(files.length>2048)throw new Error('The backup has too many file parts.');
  for(const f of files){
    if(f.unsafeOriginalName&&f.unsafeOriginalName!==f.name||f.name.startsWith('/')||f.name.includes('\\')||f.name.split('/').some(x=>x==='..'))throw new Error('The backup contains an unsafe file path.');
    const size=f._data?.uncompressedSize;
    if(!Number.isSafeInteger(size)||size<0||size>MAX_MASTER_BYTES||(total+=size)>maxTotal)throw new Error('The expanded backup exceeds this prototype’s limits.');
  }
  return files;
}
function controlText(content){
  const walk=node=>node.nodeType===1?(node.namespaceURI===W&&node.localName==='t'?node.textContent:node.namespaceURI===W&&node.localName==='br'?'\n':node.namespaceURI===W&&node.localName==='tab'?'\t':Array.from(node.childNodes).map(walk).join('')):'';
  const paragraphs=Array.from(content.getElementsByTagNameNS(W,'p'));
  return paragraphs.length?paragraphs.map(walk).join('\n'):walk(content);
}
export async function inspectMaster(master,JSZip){
  if(!master?.byteLength||master.byteLength>MAX_MASTER_BYTES)throw new Error('The Word master is empty or exceeds 64 MB.');
  const zip=await strictCaptureZip(master,JSZip);bounded(zip);
  if(!zip.file('[Content_Types].xml')||!zip.file('word/document.xml')||!zip.file('_rels/.rels'))throw new Error('The backup does not contain a complete Word document.');
  const types=xml(await zip.file('[Content_Types].xml').async('string'));
  if(!Array.from(types.getElementsByTagNameNS('*','Override')).some(n=>n.getAttribute('PartName')==='/word/document.xml'&&n.getAttribute('ContentType')==='application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'))throw new Error('Only a standard DOCX master can be restored.');
  const found=[],journals=[];
  for(const f of Object.values(zip.files).filter(f=>/^word\/webextensions\/webextension[^/]*\.xml$/.test(f.name))){
    const doc=xml(await f.async('string'));
    for(const prop of Array.from(doc.getElementsByTagNameNS(WE,'property'))){
      if(prop.getAttribute('name')===KEY)found.push(JSON.parse(prop.getAttribute('value')));
      if(prop.getAttribute('name')===PENDING)journals.push(JSON.parse(prop.getAttribute('value')));
    }
  }
  if(found.length!==1||journals.length>1)throw new Error('The Word master has missing or ambiguous OpenStage course records.');
  const project=validate(found[0]),pending=journals[0];
  if(hasFrontMatter(project)&&!validFrontJournal(pending,project))throw new Error('The governed master is missing its supported front-matter receipt.');
  if(pending&&!((hasFrontMatter(project)?validFrontJournal(pending,project):(pending.schema===2||pending.schema===3&&['course','module','activity','metadata','activity-settings'].includes(pending.kind)))&&pending.phase==='committed'&&equal(pending.project,project)))throw new Error('This master contains an unresolved operation. Use a verified checkpoint from before the interruption.');
  const doc=xml(await zip.file('word/document.xml').async('string'));
  await inspectOwnedCuePackage(project,zip);
  await inspectFrontMatterPackage(project,zip);
  const records=Array.from(doc.getElementsByTagNameNS(W,'sdt')).map(c=>{
    const pr=Array.from(c.childNodes).find(n=>n.namespaceURI===W&&n.localName==='sdtPr');
    const tag=pr?.getElementsByTagNameNS(W,'tag')[0]?.getAttributeNS(W,'val')||'';
    const content=Array.from(c.childNodes).find(n=>n.namespaceURI===W&&n.localName==='sdtContent');
    return {tag,text:content?controlText(content):''};
  });
  if(!equal(reconcile(project,records),project))throw new Error('The saved master’s text and course records differ. Refresh and try saving the backup again.');
  if(pending?.schema===3||pending?.schema===5&&pending.cueChecks?.length){
    const checks=pending.cueChecks||(pending.settings?.id?[pending.settings.id]:[]);
    if(!Array.isArray(checks)||!checks.length||new Set(checks).size!==checks.length)throw new Error('The timing validation evidence is incomplete.');
    if(pending.kind==='activity-settings'){
      const s=pending.settings,item=project.modules.flatMap(m=>m.activities).find(a=>a.id===s?.id);
      if(!item||!checks.includes(item.id)||item.minutes!==s.after?.minutes||item.lines!==s.after?.lines)throw new Error('The activity settings receipt is inconsistent.');
    }
    for(const id of checks){
    const item=project.modules.flatMap(m=>m.activities).find(a=>a.id===id);
    if(!item)throw new Error('An edited activity is missing from the timing evidence.');
    const tag=`osg:activity:${item.id}`;
    const control=Array.from(doc.getElementsByTagNameNS(W,'sdt')).find(c=>Array.from(c.getElementsByTagNameNS(W,'tag')).some(t=>t.getAttributeNS(W,'val')===tag));
    let anchor=control;
    while(anchor?.parentNode&&!(anchor.parentNode.namespaceURI===W&&anchor.parentNode.localName==='body'))anchor=anchor.parentNode;
    let cue=anchor?.previousSibling;while(cue&&cue.nodeType!==1)cue=cue.previousSibling;
    if(!cue||cue.namespaceURI!==W||cue.localName!=='p'||controlText(cue).trim()!==`${item.icon} · ${item.minutes} min`)throw new Error('The saved activity timing cue does not match its settings.');
    }
  }
  if(pending?.placement===true){
    const managed=Array.from(doc.getElementsByTagNameNS(W,'sdt')).filter(c=>{
      const pr=Array.from(c.childNodes).find(n=>n.namespaceURI===W&&n.localName==='sdtPr');
      return pr?.getElementsByTagNameNS(W,'tag')[0]?.getAttributeNS(W,'val')?.startsWith('osg:');
    });
    const tags=records.filter(r=>r.tag.startsWith('osg:')).map(r=>r.tag);
    if(!equal(tags,[...managedTags(project)]))throw new Error('The preserved master has an invalid module order.');
    for(const c of managed){
      for(let n=c.parentNode;n&&!(n.namespaceURI===W&&n.localName==='body');n=n.parentNode){
        if(n.namespaceURI===W&&['sdt','tc'].includes(n.localName))throw new Error('The preserved master has nested managed fields.');
      }
    }
  }
  return {project,parts:Object.values(zip.files).filter(f=>!f.dir).length};
}
export async function inspectCapturedFrontMatter(master,project,Zip,options){return inspectFrontMatterPackage(project,await strictCaptureZip(master,Zip),options);}
export async function assertBlankMaster(master,Zip,expectedPending){try{return await inspectBlankPackage(await strictCaptureZip(master,Zip),expectedPending);}catch{throw Object.assign(new Error('The Wrench starter requires a new blank supported DOCX, without existing text, pictures, tables, controls or story content and with only its exact own prepared metadata. Nothing has been cleared.'),{code:'OSG_INPUT'});}}
export async function createCheckpoint(master,project,JSZip,fetchAsset=fetch){
  if(!master?.byteLength||master.byteLength>MAX_MASTER_BYTES)throw new Error('The Word master is empty or exceeds 64 MB.');
  const masterInput=bytes(master),capturedMaster=new Uint8Array(masterInput);
  const inspected=await inspectMaster(capturedMaster,JSZip);
  if(!equal(inspected.project,project))throw new Error('The captured Word master differs from the checked course. No backup was prepared.');
  const files=new Map([['master.docx',capturedMaster]]);
  for(const [,filename] of ICONS){
    const r=await fetchAsset(moduleAsset('assets/icons/'+encodeURIComponent(filename),import.meta.url));
    if(!r.ok)throw new Error('A teaching asset could not be captured. No backup was prepared.');
    files.set('assets/icons/'+filename,new Uint8Array(new Uint8Array(await r.arrayBuffer())));
  }
  files.set('assets/NOTICE.txt',new TextEncoder().encode('Training icons by Jason Wrench. CC BY-NC-SA 4.0. Original artwork used without modification.\nhttps://creativecommons.org/licenses/by-nc-sa/4.0/\nThe textbook icon license does not apply to the OpenStage brand or toolbar artwork.\n'));
  const manifest={format:FORMAT,version:VERSION,id:crypto.randomUUID(),created:new Date().toISOString(),course:{id:inspected.project.id,title:inspected.project.title,revision:inspected.project.revision},master:'master.docx',files:[]};
  const zip=new JSZip();
  for(const [path,data] of files){if(data.length>MAX_MASTER_BYTES)throw new Error('A backup asset exceeds the size limit.');manifest.files.push({path,bytes:data.length,sha256:await sha256(data)});zip.file(path,new Uint8Array(data));}
  // Complete-string encoding avoids splitting a surrogate pair in JSZip's
  // streaming string worker. Never hand the canonical expected bytes to ZIP.
  const manifestText=JSON.stringify(manifest,null,2),manifestBytes=new TextEncoder().encode(manifestText);
  zip.file('manifest.json',new Uint8Array(manifestBytes),{binary:true});
  const generated=await zip.generateAsync({type:'uint8array',compression:'DEFLATE'});
  if(!(generated instanceof Uint8Array))throw new Error('The backup exporter returned unsupported bytes.');
  const data=new Uint8Array(generated),checked=await strictCaptureZip(new Uint8Array(data),JSZip,MAX_PACKAGE_BYTES);
  const actual=await checked.file('manifest.json')?.async('uint8array');
  if(!(actual instanceof Uint8Array)||!exactInput(actual,manifestBytes))throw new Error('The generated backup manifest bytes changed. No backup was prepared.');
  if(new TextDecoder('utf-8',{fatal:true}).decode(actual)!==manifestText)throw new Error('The generated backup manifest encoding changed. No backup was prepared.');
  await readCheckpoint(new Uint8Array(data),JSZip); // Keep legacy reader compatibility; stricter bytes apply to this new write only.
  if(!equal(project,inspected.project)||!exactInput(masterInput,capturedMaster))throw new Error('Checkpoint inputs changed during packaging. No backup was prepared.');
  return {data:new Uint8Array(data),manifest};
}
export async function readCheckpoint(data,JSZip){
  data=bytes(data);if(!data.length||data.length>MAX_PACKAGE_BYTES)throw new Error('Choose an OpenStage backup up to 80 MB.');
  const zip=await strictCaptureZip(data,JSZip,MAX_PACKAGE_BYTES);const files=bounded(zip);
  if(!zip.file('manifest.json')||zip.file('manifest.json')._data.uncompressedSize>256*1024)throw new Error('The backup manifest is missing or too large.');
  const manifest=JSON.parse(await zip.file('manifest.json').async('string'));
  if(manifest.format!==FORMAT||manifest.version!==VERSION||manifest.master!=='master.docx'||typeof manifest.id!=='string'||!Number.isFinite(Date.parse(manifest.created))||!Array.isArray(manifest.files))throw new Error('This backup version or format is unsupported.');
  const required=new Set(['master.docx','assets/NOTICE.txt',...ICONS.map(([,f])=>'assets/icons/'+f)]);
  if(manifest.files.length!==required.size||files.length!==required.size+1)throw new Error('The backup is missing files or contains unexpected files.');
  const verified=new Map();
  for(const item of manifest.files){
    if(!required.delete(item.path)||!Number.isSafeInteger(item.bytes)||item.bytes<=0||!/^[a-f0-9]{64}$/.test(item.sha256)||!zip.file(item.path))throw new Error('The backup file list is invalid.');
    const file=zip.file(item.path);if(file._data.uncompressedSize!==item.bytes)throw new Error('A backup file has the wrong size.');
    const content=await file.async('uint8array');
    if(await sha256(content)!==item.sha256)throw new Error('A backup file failed its integrity check. Restore is unavailable.');
    verified.set(item.path,content);
  }
  if(required.size)throw new Error('The backup is incomplete.');
  const master=verified.get('master.docx'),inspection=await inspectMaster(master,JSZip);
  if(!equal(manifest.course,{id:inspection.project.id,title:inspection.project.title,revision:inspection.project.revision}))throw new Error('The manifest does not match the preserved master.');
  return {manifest,master,project:inspection.project,masterHash:await sha256(master),assetCount:verified.size-1};
}
// Pure prerequisite for future destructive edits. This binds supplied captured
// bytes only; a controller must obtain a fresh live capture and recheck freshness
// at its write boundary. This receipt alone never authorizes a Word mutation.
export async function bindSavedCheckpoint(savedBackupBytes,currentMasterBytes,project,JSZip){
 try{
  const fail=message=>{throw Object.assign(new Error(message),{code:'OSG_CHECKPOINT'});};
  const expected=structuredClone(validate(project)),backupInput=bytes(savedBackupBytes),masterInput=bytes(currentMasterBytes);
  if(!backupInput.length||backupInput.length>MAX_PACKAGE_BYTES||!masterInput.length||masterInput.length>MAX_MASTER_BYTES)fail('Choose a complete saved backup and a supported current master capture.');
  const backup=new Uint8Array(backupInput),master=new Uint8Array(masterInput),checked=await readCheckpoint(backup,JSZip);
  if(!equal(checked.project,expected))fail('The saved backup does not match this checked course and revision. Prepare and verify a fresh backup.');
  const current=await inspectMaster(master,JSZip);
  if(!equal(current.project,expected)||await sha256(master)!==checked.masterHash)fail('The current master differs from the saved backup, including manual content or file packaging. Preserve the latest work and prepare a fresh backup.');
  const backupHash=await sha256(backup);
  // No await after this final comparison. Caller-provided Buffer/subclass
  // slice/iteration methods cannot turn the private snapshots into aliases.
  if(!equal(project,expected)||!exactInput(backupInput,backup)||!exactInput(masterInput,master))fail('Checkpoint inputs changed during verification. No edit is authorized.');
  const proof=Object.freeze({format:'openstage-captured-master-checkpoint-binding',version:1,checkpointId:checked.manifest.id,projectId:expected.id,revision:expected.revision,backupHash,masterHash:checked.masterHash,scope:'Exact supplied captured bytes only',nativeDocumentVerified:false});
  bindings.set(proof,{master,project:expected});return proof;
 }catch(cause){const error=cause instanceof Error?cause:new Error(String(cause));error.code='OSG_CHECKPOINT';throw error;}
}
const stableValue=value=>Array.isArray(value)?value.map(stableValue):value&&typeof value==='object'?Object.fromEntries(Object.keys(value).sort().map(k=>[k,stableValue(value[k])])):value;
const sameValue=(a,b)=>equal(stableValue(a),stableValue(b));
function canonicalNode(node){
  if(node.nodeType===1)return [1,node.namespaceURI||'',node.nodeName,Array.from(node.attributes).map(a=>[a.namespaceURI||'',a.nodeName,a.value]).sort((a,b)=>JSON.stringify(a).localeCompare(JSON.stringify(b))),Array.from(node.childNodes).map(canonicalNode)];
  if(node.nodeType===9)return [9,Array.from(node.childNodes).map(canonicalNode)];
  return [node.nodeType,node.nodeName,node.nodeValue];
}
// Strict envelope for master/backup readers and the preparation proof. Validate before
// JSZip can merge duplicate entries or expand a malformed archive.
export function captureEnvelope(data,maxBytes=MAX_MASTER_BYTES){
  const v=new DataView(data.buffer,data.byteOffset,data.byteLength),len=data.length;
  if(len<22||len>maxBytes)throw new Error('Unsupported captured ZIP.');
  let end=-1;for(let i=len-22;i>=Math.max(0,len-65557);i--)if(v.getUint32(i,true)===0x06054b50&&i+22+v.getUint16(i+20,true)===len){end=i;break;}
  if(end<0)throw new Error('Unsupported captured ZIP directory.');
  const count=v.getUint16(end+10,true),size=v.getUint32(end+12,true),offset=v.getUint32(end+16,true);
  if(v.getUint16(end+4,true)||v.getUint16(end+6,true)||v.getUint16(end+8,true)!==count||count>2048||count===65535||offset+size!==end)throw new Error('Unsupported multipart or ZIP64 capture.');
  const names=new Set(),payloadNames=[],ranges=[];let at=offset,total=0;
  for(let i=0;i<count;i++){
    if(at+46>end||v.getUint32(at,true)!==0x02014b50)throw new Error('Invalid captured ZIP directory.');
    const flags=v.getUint16(at+8,true),method=v.getUint16(at+10,true),compressed=v.getUint32(at+20,true),expanded=v.getUint32(at+24,true),n=v.getUint16(at+28,true),extra=v.getUint16(at+30,true),comment=v.getUint16(at+32,true),local=v.getUint32(at+42,true);
    // PKWARE APPNOTE6.3.10 sections4.4.4/4.4.7-9: supported
    // STORE/DEFLATE, descriptor and UTF-8; DEFLATE option bits1/2 only.
    const allowedFlags=method===8?0x080e:0x0808;
    if(at+46+n+extra+comment>end||(flags&~allowedFlags)!==0||![0,8].includes(method)||v.getUint16(at+34,true)||expanded>MAX_MASTER_BYTES||(total+=expanded)>128*1024*1024)throw new Error('Unsupported or oversized captured ZIP parts.');
    const name=new TextDecoder('utf-8',{fatal:true}).decode(data.subarray(at+46,at+46+n)),path=name.endsWith('/')?name.slice(0,-1):name;
    if(!path||path.startsWith('/')||/[\\:%?#\x00-\x1f]/.test(path)||path.split('/').some(p=>!p||p==='.'||p==='..')||names.has(name.toLowerCase()))throw new Error('Unsafe or duplicate captured ZIP paths.');names.add(name.toLowerCase());if(!name.endsWith('/'))payloadNames.push(name);
    if(local+30>offset||v.getUint32(local,true)!==0x04034b50||v.getUint16(local+6,true)!==flags||v.getUint16(local+8,true)!==method)throw new Error('Inconsistent captured ZIP headers.');
    const ln=v.getUint16(local+26,true),le=v.getUint16(local+28,true);let finish=local+30+ln+le+compressed;
    if(finish>offset||ln!==n||!data.subarray(local+30,local+30+ln).every((b,j)=>b===data[at+46+j]))throw new Error('Inconsistent captured ZIP bounds.');
    if(flags&8){if(v.getUint32(local+14,true)||v.getUint32(local+18,true)||v.getUint32(local+22,true))throw new Error('Descriptor-mode local CRC and sizes must be zero.');const descriptor=finish+(finish+4<=offset&&v.getUint32(finish,true)===0x08074b50?4:0);if(descriptor+12>offset||v.getUint32(descriptor,true)!==v.getUint32(at+16,true)||v.getUint32(descriptor+4,true)!==compressed||v.getUint32(descriptor+8,true)!==expanded)throw new Error('Unsupported captured ZIP data descriptor.');finish=descriptor+12;}
    else if(v.getUint32(local+14,true)!==v.getUint32(at+16,true)||v.getUint32(local+18,true)!==compressed||v.getUint32(local+22,true)!==expanded)throw new Error('Inconsistent captured ZIP sizes or CRC.');
    ranges.push([local,finish]);at+=46+n+extra+comment;
  }
  if(at!==end)throw new Error('Invalid captured directory size.');ranges.sort((a,b)=>a[0]-b[0]);let boundary=0;for(const [start,finish] of ranges){if(start!==boundary)throw new Error('Overlapping or unlisted captured ZIP payloads.');boundary=finish;}if(boundary!==offset)throw new Error('Unlisted captured ZIP payloads.');return payloadNames.sort();
}
async function strictCaptureZip(data,Zip,maxBytes){
  const names=captureEnvelope(data,maxBytes),zip=await Zip.loadAsync(data,{checkCRC32:true,createFolders:false}),loaded=Object.values(zip.files).filter(f=>!f.dir).map(f=>f.name).sort();
  if(!equal(names,loaded))throw new Error('Captured ZIP names changed during decoding.');return zip;
}
async function journalCapture(master,project,expectedPending,Zip){
  const zip=await strictCaptureZip(master,Zip),files=bounded(zip),extensions=new Map(),course=[],pending=[];
  for(const f of files.filter(f=>/^word\/webextensions\/webextension[^/]*\.xml$/.test(f.name))){
    const raw=await f.async('string');if(/<!DOCTYPE|<!ENTITY/i.test(raw))throw new Error('Unsupported checkpoint storage XML.');
    const errors=[],doc=new DOMParser({errorHandler:{warning:e=>errors.push(e),error:e=>errors.push(e),fatalError:e=>errors.push(e)}}).parseFromString(raw,'application/xml');if(errors.length||!doc.documentElement)throw new Error('Malformed checkpoint storage XML.');
    extensions.set(f.name,{doc,raw});
    for(const property of Array.from(doc.getElementsByTagNameNS(WE,'property'))){const name=property.getAttribute('name');if(name===KEY)course.push({part:f.name,node:property,value:JSON.parse(property.getAttribute('value'))});if(name===PENDING)pending.push({part:f.name,node:property,value:JSON.parse(property.getAttribute('value'))});}
  }
  if(course.length!==1||pending.length>1||!equal(course[0].value,project))throw new Error('Ambiguous or changed checkpoint course storage.');
  const owner=course[0],parent=owner.node.parentNode;
  if(parent.namespaceURI!==WE||parent.localName!=='properties'||parent.parentNode!==extensions.get(owner.part).doc.documentElement||parent.parentNode.namespaceURI!==WE||parent.parentNode.localName!=='webextension')throw new Error('Unsupported checkpoint storage location.');
  if(pending.length&&(pending[0].part!==owner.part||pending[0].node.parentNode!==parent))throw new Error('The pending record is outside its expected storage location.');
  if(pending.length){const property=pending[0].node;if(property.childNodes.length||property.attributes.length!==2||!property.hasAttribute('name')||!property.hasAttribute('value'))throw new Error('Unsupported pending property metadata.');}
  if(expectedPending===undefined&&pending.length){const receipt=pending[0].value;if(!((receipt?.schema===2||receipt?.schema===3&&['course','module','activity','metadata','activity-settings'].includes(receipt.kind))&&receipt.phase==='committed'&&equal(receipt.project,project)))throw new Error('The baseline does not contain a supported completed receipt.');}
  if(expectedPending!==undefined&&(pending.length!==1||!sameValue(pending[0].value,expectedPending)))throw new Error('The pending record differs from the exact intended preparation.');
  if(pending.length)parent.removeChild(pending[0].node);
  const parts=[];
  for(const f of files.sort((a,b)=>a.name.localeCompare(b.name))){
    if(f.name===owner.part){const {doc,raw}=extensions.get(owner.part);parts.push([f.name,'storage-xml',raw.match(/^\uFEFF?<\?xml[^?]*\?>/)?.[0]||'',canonicalNode(doc)]);}
    else parts.push([f.name,'exact-payload',await sha256(await f.async('uint8array'))]);
  }
  return {ownerPart:owner.part,parts};
}
// Session-only, supplied-capture preparation proof. All unknown payloads remain
// covered. Only the exact pending property is removed from a structural XML
// comparison of its owning extension; this is not whole-file byte equality.
// This checks a storage delta, not the proposed icon transition, and is not
// connected to journal acceptance, a live document, or any mutation adapter.
export async function verifyPreparedCheckpointCapture(binding,preparedMasterBytes,journal,Zip){
 try{
  const saved=bindings.get(binding);if(!saved)throw new Error('Reverify the saved backup and current master in this session first.');
  const expected=structuredClone(journal),input=bytes(preparedMasterBytes);if(!input.length||input.length>MAX_MASTER_BYTES)throw new Error('The prepared capture is unsupported.');const captured=new Uint8Array(input);
  if(expected?.schema!==4||expected.kind!=='activity-icon'||expected.phase!=='prepared'||typeof expected.txId!=='string'||!/^[a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12}$/i.test(expected.txId)||!equal(expected.base,saved.project)||expected.project?.id!==saved.project.id||expected.project.revision!==saved.project.revision+1||!sameValue(expected.checkpoint,{id:binding.checkpointId,masterHash:binding.masterHash,backupHash:binding.backupHash}))throw new Error('The intended icon preparation does not match the saved checkpoint.');
  validate(expected.project);
  const before=await journalCapture(saved.master,saved.project,undefined,Zip),after=await journalCapture(captured,saved.project,expected,Zip);
  if(!equal(before,after))throw new Error('Master content or unknown storage changed during preparation. No edit is authorized.');
  // Complete every asynchronous read/hash above, then seal original inputs.
  if(!equal(journal,expected)||!exactInput(input,captured))throw new Error('Preparation inputs changed during verification. No edit is authorized.');
  const proof=Object.freeze({format:'openstage-prepared-capture-proof',version:1,checkpointId:binding.checkpointId,projectId:binding.projectId,revision:binding.revision,txId:expected.txId,kind:expected.kind,scope:'Only exact pending property differs; other storage XML compared structurally, all other payloads byte-hashed',nativeDocumentVerified:false});
  preparedBindings.set(proof,{binding,journal:expected});return proof;
 }catch(cause){const error=cause instanceof Error?cause:new Error(String(cause));error.code='OSG_CHECKPOINT';throw error;}
}
// Identity correlation of previously checked supplied captures only. No live
// identity, freshness, physical state or mutation authorization is implied.
export function matchesCheckpointPreparation(binding,proof,journal){
  try{
    const checked=cloneIconData(journal),prepared=preparedBindings.get(proof);
    return !!bindings.get(binding)&&!!prepared&&prepared.binding===binding&&sameValue(prepared.journal,checked);
  }catch{return false;}
}
export function matchesCheckpointBase(binding,project){
  try{return !!bindings.get(binding)&&sameValue(bindings.get(binding).project,cloneIconData(project));}catch{return false;}
}
export function checkpointName(title,restored=false){return (title.replace(/[^\p{L}\p{N} _-]/gu,'').slice(0,65)||'OpenStage Guide')+' - '+(restored?'Restored ':'Backup ')+new Date().toISOString().replace(/[:.]/g,'-')+'-'+crypto.randomUUID().slice(0,8)+(restored?'.docx':'.osgbackup');}
