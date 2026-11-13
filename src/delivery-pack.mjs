import {resolveSelectedPlanScope,SELECTED_PLAN_LIMITS,SELECTED_WORKBOOK_OBJECTIVES} from './selected-plan-scope.mjs';
import {planAgendaPreview,planAgendaFiles} from './plan-agenda.mjs';
import {planWorkbookPreview,planWorkbookFiles,planMaterialsPreview,planMaterialsFiles} from './selected-plan-outputs.mjs';
import {planFacilitatorPreview,SELECTED_FACILITATOR_LIMITS} from './selected-plan-facilitator.mjs';
import {facilitatorOptions,facilitatorPackage} from './facilitator-output.mjs';
import {captureEnvelope} from './checkpoint.mjs';

export const PACK_PUBLIC='shareable',PACK_PRIVATE='private';
const DOCX_PARTS=['[Content_Types].xml','_rels/.rels','docProps/core.xml','word/_rels/document.xml.rels','word/document.xml','word/styles.xml'];
const PUBLIC_MEMBERS=['agenda.docx','participant-workbook.docx','materials-checklist.docx'];
const PRIVATE_MEMBER='PRIVATE-facilitator-guide.docx';
const MANIFEST='manifest.json';
// Only bundled ZIP metadata is fixed. Standalone DOCX generators keep their
// accepted defaults; the reviewed generation stamp remains in file content.
const PACK_ZIP_DATE=new Date('2000-01-01T00:00:00.000Z');
const PACK_FACILITATOR_LIMITS=SELECTED_FACILITATOR_LIMITS.replace('Coordinated packs and existing-output regeneration are unavailable.','Existing-output regeneration remains unavailable.');
const MAX_PACK=32*1024*1024,MAX_MEMBER=12*1024*1024;
const same=(a,b)=>JSON.stringify(a)===JSON.stringify(b);
const reject=message=>{throw new Error(`Delivery pack: ${message}`);};
const safeName=name=>typeof name==='string'&&/^[A-Za-z0-9][A-Za-z0-9._-]*$/.test(name)&&!name.includes('..')&&!name.includes(':')&&!name.includes('\\')&&!name.includes('/')&&!/^(?:https?|file)$/i.test(name);
const validStamp=stamp=>typeof stamp==='string'&&/^\d{4}-\d\d-\d\dT\d\d:\d\d:\d\d\.\d{3}Z$/.test(stamp)&&Number.isFinite(Date.parse(stamp))&&new Date(stamp).toISOString()===stamp;
function packOptions(input){
 if(input===null||typeof input!=='object'||Array.isArray(input))reject('Use supported pack choices.');
 const allowed=['route','start','stamp','answers','participant'],descriptors=Object.getOwnPropertyDescriptors(input);
 if(Reflect.ownKeys(input).some(key=>typeof key!=='string'||!allowed.includes(key))||Reflect.ownKeys(descriptors).some(key=>!('value' in descriptors[key])))reject('Pack choices must be explicit data fields.');
 for(const key of allowed)if(key in input&&!Object.prototype.hasOwnProperty.call(input,key))reject('Inherited pack choices are unsupported.');
 const route=Object.hasOwn(descriptors,'route')?descriptors.route.value:PACK_PUBLIC,start=Object.hasOwn(descriptors,'start')?descriptors.start.value:'09:00',stamp=descriptors.stamp?.value;
 if(route!==PACK_PUBLIC&&route!==PACK_PRIVATE)reject('Choose a supported pack route.');
 if(typeof start!=='string')reject('Use an explicit workshop start time.');
 if(route===PACK_PUBLIC){if(descriptors.answers||descriptors.participant)reject('Private choices are unavailable in a shareable pack.');return {route,start,stamp,answers:false,participant:false};}
 const choice=facilitatorOptions({answers:Object.hasOwn(descriptors,'answers')?descriptors.answers.value:false,participant:Object.hasOwn(descriptors,'participant')?descriptors.participant.value:false});
 return {route,start,stamp,...choice};
}
const sha=async data=>Array.from(new Uint8Array(await crypto.subtle.digest('SHA-256',data)),x=>x.toString(16).padStart(2,'0')).join('');
const bytes=input=>input instanceof Uint8Array?input:new Uint8Array(input);
// Uint8Array constructor copies Buffer and Uint8Array elements; unlike slice,
// it never returns a view into a generator-owned backing store.
const ownBytes=value=>{if(!(value instanceof Uint8Array))reject('ZIP generation returned unsupported bytes.');const copy=new Uint8Array(value.length);copy.set(value);return copy;};
const utf8=data=>{try{return new TextDecoder('utf-8',{fatal:true}).decode(data);}catch{reject('A generated text part is not valid UTF-8.');}};
const sameBytes=(left,right)=>left.length===right.length&&left.every((value,index)=>value===right[index]);
const namesFor=route=>route===PACK_PUBLIC?[...PUBLIC_MEMBERS,MANIFEST]:route===PACK_PRIVATE?[...PUBLIC_MEMBERS,PRIVATE_MEMBER,MANIFEST]:reject('Choose a supported pack route.');

// Reuse the accepted master ZIP envelope checker before JSZip can merge names.
// Packs impose a narrower generated-only profile over that proven envelope.
function rawNames(input,max=MAX_PACK,inner=false){
 const data=bytes(input),names=captureEnvelope(data,max),v=new DataView(data.buffer,data.byteOffset,data.byteLength),len=data.length;
 const end=len-22;if(v.getUint32(end,true)!==0x06054b50||v.getUint16(end+20,true)!==0)reject('ZIP comments are unavailable in a generated pack.');
 if(names.length<1||names.length>6||v.getUint16(end+10,true)!==names.length)reject('The generated pack has an unsupported member count or directory entry.');
 let at=v.getUint32(end+16,true),total=0;
 for(let i=0;i<names.length;i++){
  const n=v.getUint16(at+28,true),extra=v.getUint16(at+30,true),comment=v.getUint16(at+32,true),local=v.getUint32(at+42,true),expanded=v.getUint32(at+24,true),attribute=v.getUint32(at+38,true),internal=v.getUint16(at+36,true),madeBy=v.getUint16(at+4,true),localExtra=v.getUint16(local+28,true);
  const name=new TextDecoder('utf-8',{fatal:true}).decode(data.subarray(at+46,at+46+n));
  if(!(inner?DOCX_PARTS.includes(name):safeName(name))||extra||comment||localExtra||expanded>MAX_MEMBER||(total+=expanded)>MAX_PACK||v.getUint16(at+34,true)||internal||attribute||madeBy!==0x14)reject('Unsupported generated ZIP member metadata.');
  at+=46+n+extra+comment;
 }
 return names;
}
const roster=(actual,expected)=>{if(!same([...actual].sort(),[...expected].sort()))reject('The member roster differs from the reviewed route.');};
const keys=(value,expected)=>value&&typeof value==='object'&&!Array.isArray(value)&&same(Object.keys(value).sort(),[...expected].sort());
const plain=value=>value!==null&&typeof value==='object'&&!Array.isArray(value)&&[Object.prototype,null].includes(Object.getPrototypeOf(value));
function expectedContext(input){
 if(!plain(input))reject('Expected pack context must be an explicit object.');
 const fields=Object.getOwnPropertyDescriptors(input),names=Reflect.ownKeys(fields);
 if(names.some(key=>typeof key!=='string'||!['expectedManifestBytes','expectedDocFiles'].includes(key)||!('value' in fields[key]))||['expectedManifestBytes','expectedDocFiles'].some(key=>key in input&&!Object.hasOwn(fields,key)))reject('Expected pack context contains unsupported or inherited fields.');
 const hasManifest=Object.hasOwn(fields,'expectedManifestBytes'),hasFiles=Object.hasOwn(fields,'expectedDocFiles');
 if(hasManifest!==hasFiles||hasManifest&&typeof fields.expectedManifestBytes.value!=='string')reject('Expected pack context must include the complete manifest and file maps together.');
 return {manifestBytes:hasManifest?fields.expectedManifestBytes.value:undefined,rawFiles:hasFiles?fields.expectedDocFiles.value:undefined,hasFiles};
}
function expectedPartMaps(raw,memberNames){
 if(!plain(raw))reject('Expected Word files must be complete own-data maps.');
 const fields=Object.getOwnPropertyDescriptors(raw),names=Reflect.ownKeys(fields);
 if(names.some(key=>typeof key!=='string')||!same(names.sort(),[...memberNames].sort()))reject('Expected Word member maps differ from the pack roster.');
 const result={};for(const name of memberNames){
  const field=fields[name],parts=field?.value;if(!field||!('value' in field)||!plain(parts))reject('Expected Word member maps must be own data.');
  const entries=Object.getOwnPropertyDescriptors(parts),paths=Reflect.ownKeys(entries);
  if(paths.some(key=>typeof key!=='string')||!same(paths.sort(),[...DOCX_PARTS].sort())||paths.some(path=>!('value' in entries[path])||typeof entries[path].value!=='string'))reject('Expected Word parts must be exact six-part text maps.');
  result[name]=Object.fromEntries(DOCX_PARTS.map(path=>[path,entries[path].value]));
 }return result;
}
function xmlValid(body){
 if(typeof body!=='string'||/<!DOCTYPE|<!ENTITY/i.test(body))reject('A Word part contains unsupported XML.');
 if(typeof DOMParser==='undefined')reject('An XML parser is required to validate generated Word files.');
 let doc;try{doc=new DOMParser({errorHandler:{error:message=>{throw new Error(message);},fatalError:message=>{throw new Error(message);}}}).parseFromString(body,'application/xml');}catch{reject('A Word part contains invalid XML.');}
 if(doc.getElementsByTagName('parsererror').length||!doc.documentElement)reject('A Word part contains invalid XML.');
 return doc;
}
function docxParagraphs(files){
 const doc=xmlValid(files['word/document.xml']),W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
 const read=node=>node.nodeType!==1?'':node.namespaceURI===W&&node.localName==='t'?node.textContent:node.namespaceURI===W&&node.localName==='br'?'\n':node.namespaceURI===W&&node.localName==='tab'?'\t':Array.from(node.childNodes).map(read).join('');
 return Array.from(doc.getElementsByTagNameNS(W,'p')).map(read).join('\n');
}
async function validateDocx(data,Zip,expected){
 const raw=rawNames(data,MAX_MEMBER,true);roster(raw,DOCX_PARTS);
 const z=await Zip.loadAsync(data,{checkCRC32:true,createFolders:false});roster(Object.values(z.files).filter(f=>!f.dir).map(f=>f.name),DOCX_PARTS);
 const partHashes={};let types='';for(const path of DOCX_PARTS){const rawPart=await z.file(path).async('uint8array'),actual=utf8(rawPart);xmlValid(actual);if(expected&&!sameBytes(rawPart,new TextEncoder().encode(expected[path])))reject(`The ${path} Word part changed during packaging.`);partHashes[path]=await sha(rawPart);if(path==='[Content_Types].xml')types=actual;}
 if(!types.includes('PartName="/word/document.xml"')||!types.includes('wordprocessingml.document.main+xml'))reject('The Word content type is missing.');
 return {partCount:raw.length,partHashes};
}
const identity=(resolved,start,stamp)=>({projectId:resolved.projectId,revision:resolved.revision,planId:resolved.planId,planName:resolved.planName,scope:resolved.scope,minutes:resolved.minutes,start,end:resolved.end,generation:stamp});
function checkPreviews(base,agenda,workbook,materials,facilitator){
 for(const [label,p] of [['agenda',agenda],['workbook',workbook],['materials',materials],...(facilitator?[['facilitator',facilitator]]:[])]){
  if(p.projectId!==base.projectId||p.revision!==base.revision||p.planId!==base.planId||p.planName!==base.planName||p.minutes!==base.minutes||!same(p.selectedScope??p.scope,base.scope))reject(`${label} has a different selected plan identity.`);
 }
 if(agenda.start!==base.start||agenda.end!==base.end||!same(agenda.rows.map(r=>r.activityId),base.scope.flatMap(g=>g.activityIds))||workbook.stamp!==base.generation||materials.stamp!==base.generation||facilitator&&facilitator.stamp!==base.generation)reject('Generated files do not share the reviewed time and identity.');
}
export function deliveryPackPreview(project,planId,config={}){
 const {route,start,stamp,answers,participant}=packOptions(config);
 if(!validStamp(stamp))reject('Use an exact valid generation stamp.');
 const options={answers,participant};
 const resolved=resolveSelectedPlanScope(project,planId,start),base=identity(resolved,start,stamp);
 const agenda=planAgendaPreview(project,planId,start),workbook=planWorkbookPreview(project,planId,stamp,start),materials=planMaterialsPreview(project,planId,stamp,start);
 const facilitator=route===PACK_PRIVATE?planFacilitatorPreview(project,planId,options,stamp,start):null;
 checkPreviews(base,agenda,workbook,materials,facilitator);
 const contents=[['agenda.docx',planAgendaFiles(project,planId,start,stamp)],['participant-workbook.docx',planWorkbookFiles(project,planId,stamp,start)],['materials-checklist.docx',planMaterialsFiles(project,planId,stamp,start)]];
 if(facilitator)contents.push([PRIVATE_MEMBER,facilitatorPackage({...facilitator,limits:PACK_FACILITATOR_LIMITS})]);
 return {route,identity:base,options:route===PACK_PRIVATE?options:null,members:namesFor(route),assumptions:[SELECTED_PLAN_LIMITS,SELECTED_WORKBOOK_OBJECTIVES,materials.warning,materials.assumption,'All outputs are structured copies. Manual Word additions, images, layout, attachments and full-master fidelity are not included. Edits do not flow back to the master.',...(route===PACK_PRIVATE?[PACK_FACILITATOR_LIMITS]:[])],privateIssues:facilitator?.issues||[],contentPreviews:contents.map(([name,files])=>({name,text:docxParagraphs(files)}))};
}
function manifestFor(preview,parts){
 const manifest={format:'openstage-guide-selected-delivery-pack',version:1,route:preview.route,label:preview.route===PACK_PRIVATE?'PRIVATE FACILITATOR PACK':'SHAREABLE LEARNER PACK',identity:preview.identity,
  facilitatorOptions:preview.options,assumptions:preview.assumptions,
  members:parts.map(({name,sha256,bytes,partCount,partHashes})=>({name,role:name===PRIVATE_MEMBER?'private-facilitator':name.replace('.docx',''),sha256,bytes,partCount,partHashes}))};
 return manifest;
}
// Structural inspection of a supplied pack is not proof of authorship. The
// generator also passes its captured expected manifest and six-part maps.
export async function verifyDeliveryPack(input,Zip,context={}){
 const expectedContextValue=expectedContext(context);
 const data=bytes(input),raw=rawNames(data),zip=await Zip.loadAsync(data,{checkCRC32:true,createFolders:false});
 roster(Object.values(zip.files).filter(f=>!f.dir).map(f=>f.name),raw);
 const file=zip.file(MANIFEST);if(!file)reject('The manifest is missing.');
 const rawManifestBytes=await file.async('uint8array'),rawManifest=utf8(rawManifestBytes);let manifest;try{manifest=JSON.parse(rawManifest);}catch{reject('The manifest is invalid JSON.');}
 if(!sameBytes(rawManifestBytes,new TextEncoder().encode(JSON.stringify(manifest,null,2)+'\n'))||expectedContextValue.manifestBytes!==undefined&&!sameBytes(rawManifestBytes,new TextEncoder().encode(expectedContextValue.manifestBytes)))reject('The manifest bytes differ from the canonical reviewed manifest.');
 if(manifest?.format!=='openstage-guide-selected-delivery-pack'||manifest.version!==1||![PACK_PUBLIC,PACK_PRIVATE].includes(manifest.route)||!validStamp(manifest.identity?.generation))reject('The manifest identity is unsupported.');
 if(!keys(manifest,['format','version','route','label','identity','facilitatorOptions','assumptions','members'])||!keys(manifest.identity,['projectId','revision','planId','planName','scope','minutes','start','end','generation']))reject('The manifest has unsupported fields.');
 const identity=manifest.identity;
 if(![identity.projectId,identity.planId,identity.planName,identity.start,identity.end].every(x=>typeof x==='string'&&x.length>0)||!Number.isSafeInteger(identity.revision)||identity.revision<0||!Number.isFinite(identity.minutes)||identity.minutes<0||!Array.isArray(identity.scope)||!identity.scope.length||identity.scope.some(group=>!keys(group,['moduleId','activityIds'])||typeof group.moduleId!=='string'||!Array.isArray(group.activityIds)||group.activityIds.some(id=>typeof id!=='string'||!id)))reject('The manifest has a malformed selected scope.');
 const expected=namesFor(manifest.route);roster(raw,expected);
 const expectedDocFiles=expectedContextValue.hasFiles?expectedPartMaps(expectedContextValue.rawFiles,expected.filter(name=>name!==MANIFEST)):null;
 if(manifest.label!==(manifest.route===PACK_PRIVATE?'PRIVATE FACILITATOR PACK':'SHAREABLE LEARNER PACK')||!Array.isArray(manifest.members)||manifest.members.length!==expected.length-1||!same(manifest.members.map(m=>m.name),expected.filter(n=>n!==MANIFEST)))reject('The manifest member list differs from the reviewed route.');
 if(manifest.route===PACK_PUBLIC&&manifest.facilitatorOptions!==null||manifest.route===PACK_PRIVATE&&(!keys(manifest.facilitatorOptions,['answers','participant'])||typeof manifest.facilitatorOptions.answers!=='boolean'||typeof manifest.facilitatorOptions.participant!=='boolean'))reject('The manifest options differ from its route.');
 if(!Array.isArray(manifest.assumptions)||manifest.assumptions.length!==(manifest.route===PACK_PRIVATE?6:5)||manifest.assumptions.some(x=>typeof x!=='string'||!x)||manifest.assumptions[0]!==SELECTED_PLAN_LIMITS||manifest.assumptions[1]!==SELECTED_WORKBOOK_OBJECTIVES||manifest.route===PACK_PRIVATE&&manifest.assumptions[5]!==PACK_FACILITATOR_LIMITS)reject('The manifest assumptions are unsupported.');
 const scopeLine=`Exact selected scope: ${identity.scope.map(g=>`${g.moduleId}: ${g.activityIds.join(', ')}`).join(' | ')}`;
 for(const member of manifest.members){
  if(!keys(member,['name','role','sha256','bytes','partCount','partHashes'])||!keys(member.partHashes,DOCX_PARTS))reject('The manifest member shape is unsupported.');
  const body=await zip.file(member.name).async('uint8array');
  const checked=await validateDocx(body,Zip,expectedDocFiles?.[member.name]);
  if(member.bytes!==body.length||member.sha256!==await sha(body)||member.partCount!==checked.partCount||!same(member.partHashes,checked.partHashes)||member.role!==(member.name===PRIVATE_MEMBER?'private-facilitator':member.name.replace('.docx','')))reject('A Word member differs from the manifest.');
  const doc=await Zip.loadAsync(body,{checkCRC32:true,createFolders:false}),visible=docxParagraphs({'word/document.xml':utf8(await doc.file('word/document.xml').async('uint8array'))});
  for(const phrase of [`Project ID: ${identity.projectId} · Generation revision: ${identity.revision}`,`Delivery plan: ${identity.planName} · Plan ID: ${identity.planId}`,scopeLine,`Generated: ${identity.generation}`])if(!visible.includes(phrase))reject('A Word member has a different project, plan, scope or generation.');
  if(member.name==='agenda.docx'&&!visible.includes(`${identity.minutes} minutes · ${identity.start}–${identity.end}`))reject('The agenda timing differs from the manifest.');
  if(member.name===PRIVATE_MEMBER&&(!visible.includes('PRIVATE: Includes teaching directions')||!visible.includes(`Include answers/debrief: ${manifest.facilitatorOptions.answers?'Yes':'No'} · Include participant instructions: ${manifest.facilitatorOptions.participant?'Yes':'No'}`)))reject('The PRIVATE guide choices differ from the manifest.');
  if(member.name!==PRIVATE_MEMBER&&/PRIVATE-facilitator-guide/i.test(member.name))reject('A private file appeared in a public role.');
 }
 return manifest;
}
export async function buildDeliveryPack(project,planId,options,Zip,assertFresh=()=>{}){
 await assertFresh();const preview=deliveryPackPreview(project,planId,options),{route,identity:base}=preview;
 const files=[['agenda.docx',planAgendaFiles(project,planId,base.start,base.generation)],['participant-workbook.docx',planWorkbookFiles(project,planId,base.generation,base.start)],['materials-checklist.docx',planMaterialsFiles(project,planId,base.generation,base.start)]];
 if(route===PACK_PRIVATE){const privatePreview=planFacilitatorPreview(project,planId,preview.options,base.generation,base.start);files.push([PRIVATE_MEMBER,facilitatorPackage({...privatePreview,limits:PACK_FACILITATOR_LIMITS})]);}
 const parts=[],outer=new Zip(),expectedDocFiles={};
 for(const [name,docFiles] of files){
  await assertFresh();roster(Object.keys(docFiles),DOCX_PARTS);for(const body of Object.values(docFiles))xmlValid(body);
  if(docxParagraphs(docFiles)!==preview.contentPreviews.find(item=>item.name===name)?.text)reject('A Word member differs from its reviewed content.');
  const z=new Zip();for(const [part,body] of Object.entries(docFiles))z.file(part,new Uint8Array(new TextEncoder().encode(body)),{createFolders:false,date:PACK_ZIP_DATE});
  const data=ownBytes(await z.generateAsync({type:'uint8array',compression:'DEFLATE'}));await assertFresh();
  const checked=await validateDocx(data,Zip,docFiles);await assertFresh();expectedDocFiles[name]=docFiles;
  parts.push({name,sha256:await sha(data),bytes:data.length,...checked});await assertFresh();outer.file(name,data,{createFolders:false,date:PACK_ZIP_DATE});
 }
 const manifest=manifestFor(preview,parts),manifestBytes=JSON.stringify(manifest,null,2)+'\n';outer.file(MANIFEST,new Uint8Array(new TextEncoder().encode(manifestBytes)),{createFolders:false,date:PACK_ZIP_DATE});
 await assertFresh();const data=ownBytes(await outer.generateAsync({type:'uint8array',compression:'DEFLATE'}));await assertFresh();
 const checked=await verifyDeliveryPack(data,Zip,{expectedManifestBytes:manifestBytes,expectedDocFiles});const offered=ownBytes(data);await assertFresh();if(!same(checked,manifest))reject('The final pack differs from the reviewed manifest.');
 return {data:offered,manifest,preview};
}
export const deliveryPackName=route=>{if(route!==PACK_PRIVATE&&route!==PACK_PUBLIC)reject('Choose a supported pack route.');return `OpenStage ${route===PACK_PRIVATE?'PRIVATE Facilitator':'Shareable Learner'} Pack - ${crypto.randomUUID()}.zip`;};
