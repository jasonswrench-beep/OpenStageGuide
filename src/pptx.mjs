// Read-only subset: no rendering, external fetching, execution or master writes.
const P='http://schemas.openxmlformats.org/presentationml/2006/main',A='http://schemas.openxmlformats.org/drawingml/2006/main',R='http://schemas.openxmlformats.org/officeDocument/2006/relationships',PKG='http://schemas.openxmlformats.org/package/2006/relationships';
export const MAX_PPTX_BYTES=16*1024*1024;
// The slide chooser (details mode) only reads the slides, their notes and the small relationship files, never the pictures, charts and
// video, so it accepts a much bigger file than the older review does (found with the author's real decks, 2026-11-06: a 16 MB limit
// refused ordinary decks). A part it never reads is not counted against the limits on what it does read.
export const MAX_DECK_BYTES=250*1024*1024;
const LIMIT=64*1024*1024,PART=2*1024*1024;
const READ_PARTS=/^(\[Content_Types\]\.xml|_rels\/|ppt\/presentation\.xml|ppt\/_rels\/|ppt\/slides\/|ppt\/notesSlides\/)/;
function reject(message){throw new Error('PowerPoint preview: '+message);}
function safePath(s){if(!s||s.startsWith('/')||/[\\:%?#\x00-\x1f]/.test(s)||s.split('/').some(x=>x==='..'||x==='.'||x===''))reject('unsafe package path.');return s;}
// lenient: a deck with charts, equations or embedded workbooks is still read (those parts are never opened or followed); macros and ActiveX still refuse.
function envelope(data,lenient=false){
  const v=new DataView(data.buffer,data.byteOffset,data.byteLength),len=data.length;
  if(lenient&&len>MAX_DECK_BYTES)reject(`that presentation is ${(len/1048576).toFixed(0)} MB and the slide chooser reads up to ${MAX_DECK_BYTES/1048576} MB. In PowerPoint, choose File > Info > Compress Media (or save a copy without the videos), then try again.`);
  if(len<22||(!lenient&&len>MAX_PPTX_BYTES))reject('choose a standard PPTX file up to 16 MB. Encrypted presentations are unsupported.');
  let end=-1;for(let i=len-22;i>=Math.max(0,len-65557);i--)if(v.getUint32(i,true)===0x06054b50&&i+22+v.getUint16(i+20,true)===len){end=i;break;}
  if(end<0)reject('not a supported ZIP presentation; encryption and legacy PPT are unsupported.');
  const count=v.getUint16(end+10,true),size=v.getUint32(end+12,true),offset=v.getUint32(end+16,true);
  if(v.getUint16(end+4,true)||v.getUint16(end+6,true)||v.getUint16(end+8,true)!==count||count===65535||count>(lenient?20000:4096)||offset+size!==end)reject('unsupported multipart or ZIP64 package.');
  const names=new Set(),ranges=[],parts=new Map();let at=offset,total=0;
  for(let i=0;i<count;i++){
    if(at+46>end||v.getUint32(at,true)!==0x02014b50)reject('invalid ZIP directory.');
    const flags=v.getUint16(at+8,true),method=v.getUint16(at+10,true),compressed=v.getUint32(at+20,true),expanded=v.getUint32(at+24,true),n=v.getUint16(at+28,true),extra=v.getUint16(at+30,true),comment=v.getUint16(at+32,true),local=v.getUint32(at+42,true);
    if(at+46+n+extra+comment>end||flags&65||![0,8].includes(method)||v.getUint16(at+34,true))reject('encrypted, unsupported or oversized ZIP parts.');
    let name;try{name=new TextDecoder('utf-8',{fatal:true}).decode(data.subarray(at+46,at+46+n));}catch{reject('unsupported package filename encoding.');}
    const directory=name.endsWith('/'),path=directory?name.slice(0,-1):name;safePath(path);
    // Strict (the older review): every part counts. Lenient (the slide chooser): only the parts it reads count.
    if(!lenient||READ_PARTS.test(name)){if(expanded>16*1024*1024||(total+=expanded)>LIMIT)reject('encrypted, unsupported or oversized ZIP parts.');}
    if(names.has(name.toLowerCase()))reject('duplicate package paths.');names.add(name.toLowerCase());
    if(!directory)parts.set(name,{size:expanded,crc:v.getUint32(at+16,true)});
    if(/\.(?:xml|rels)$/i.test(name)&&expanded>PART&&(!lenient||READ_PARTS.test(name)))reject('XML part exceeds 2 MB.');
    if(/(?:vbaProject|activeX)/i.test(name)||(!lenient&&/(?:embeddings|oleObject)/i.test(name)))reject('macros, active content and embedded objects are unsupported.');
    if(local+30>offset||v.getUint32(local,true)!==0x04034b50||v.getUint16(local+6,true)!==flags||v.getUint16(local+8,true)!==method)reject('inconsistent ZIP headers.');
    const ln=v.getUint16(local+26,true),le=v.getUint16(local+28,true),finish=local+30+ln+le+compressed;
    if(finish>offset||ln!==n||!data.subarray(local+30,local+30+ln).every((x,j)=>x===data[at+46+j]))reject('inconsistent ZIP part bounds.');
    ranges.push([local,finish]);at+=46+n+extra+comment;
  }
  if(at!==end)reject('invalid ZIP directory size.');ranges.sort((a,b)=>a[0]-b[0]);for(let i=1;i<ranges.length;i++)if(ranges[i][0]<ranges[i-1][1])reject('overlapping ZIP parts.');
  return parts;
}
const crcTable=Array.from({length:256},(_,n)=>{for(let i=0;i<8;i++)n=n&1?0xedb88320^(n>>>1):n>>>1;return n>>>0;});
function readXmlPart(file,expected){return new Promise((resolve,rejectPromise)=>{
  const stream=file.internalStream('uint8array'),chunks=[];let length=0,crc=0xffffffff,stopped=false;
  stream.on('data',chunk=>{
    if(stopped)return;length+=chunk.length;
    if(length>PART||length>expected.size){stopped=true;stream.pause();rejectPromise(new Error('PowerPoint preview: expanded XML exceeds its size limit.'));return;}
    for(const b of chunk)crc=crcTable[(crc^b)&255]^(crc>>>8);chunks.push(chunk);
  }).on('error',e=>{if(!stopped){stopped=true;rejectPromise(e);}}).on('end',()=>{
    if(stopped)return;
    if(length!==expected.size||((crc^0xffffffff)>>>0)!==expected.crc){rejectPromise(new Error('PowerPoint preview: damaged XML part.'));return;}
    const result=new Uint8Array(length);let offset=0;for(const c of chunks){result.set(c,offset);offset+=c.length;}
    try{resolve(new TextDecoder('utf-8',{fatal:true}).decode(result));}catch{rejectPromise(new Error('PowerPoint preview: XML must use UTF-8 encoding.'));}
  }).resume();
});}
function parse(text,namespace,root){
  if(/<!DOCTYPE|<!ENTITY/i.test(text))reject('XML declarations with entities are unsupported.');
  const errors=[],doc=new DOMParser({errorHandler:{warning:m=>errors.push(m),error:m=>errors.push(m),fatalError:m=>errors.push(m)}}).parseFromString(text,'application/xml');
  if(errors.length||doc.getElementsByTagName('parsererror').length||doc.documentElement?.namespaceURI!==namespace||doc.documentElement?.localName!==root)reject('malformed or unsupported XML.');return doc;
}
function resolve(part,target){
  if(!target||/[\\:%?#\x00-\x1f]/.test(target))reject('unsupported relationship target.');
  const result=target.startsWith('/')?[]:part.split('/').slice(0,-1);
  for(const segment of target.replace(/^\//,'').split('/')){if(segment==='..'){if(!result.length)reject('relationship escapes the package.');result.pop();}else if(segment!=='.')result.push(segment);}
  return safePath(result.join('/'));
}
const nodes=(node,ns,name)=>Array.from(node.getElementsByTagNameNS(ns,name));
const direct=(node,ns,name)=>Array.from(node.childNodes).filter(n=>n.namespaceURI===ns&&n.localName===name);
function textOf(node){return nodes(node,A,'p').map(p=>{let s='';const walk=n=>{if(n.namespaceURI===A&&n.localName==='t')s+=n.textContent;else if(n.namespaceURI===A&&n.localName==='br')s+='\n';else for(const child of Array.from(n.childNodes||[]))walk(child);};walk(p);return s;}).join('\n');}
function boundedText(s){if(s.length>50000)reject('extracted text exceeds the 50,000 character field limit.');return s;}
// The first line of the largest text on the slide (the highest one when sizes are equal), leaving out page numbers, dates, footers and
// long paragraphs. '' when the slide has no such text.
function guessTitle(doc){
  const skip=['sldNum','dt','hdr','ftr','sldImg'],found=[];
  nodes(doc,P,'sp').forEach((sp,order)=>{
    if(nodes(sp,P,'ph').some(ph=>skip.includes(ph.getAttribute('type'))))return;
    const text=textOf(sp).split('\n').map(s=>s.trim()).find(Boolean)||'';
    if(!text||text.length>120||/^[\d\s.,%:/-]+$/.test(text))return;
    const sizes=nodes(sp,A,'rPr').map(r=>Number(r.getAttribute('sz'))||0).filter(Boolean),off=nodes(sp,A,'off')[0];
    found.push({text,order,size:sizes.length?Math.max(...sizes):0,y:off?Number(off.getAttribute('y'))||0:Infinity});
  });
  found.sort((a,b)=>b.size-a.size||a.y-b.y||a.order-b.order);
  return found.length?boundedText(found[0].text):'';
}
function freeze(o){Object.freeze(o);for(const value of Object.values(o))if(value&&typeof value==='object')freeze(value);return o;}
export async function readPptx(input,name,Zip,{details=false}={}){
  const data=new Uint8Array(input instanceof Uint8Array?input:new Uint8Array(input)),parts=envelope(data,details);
  if(typeof name!=='string'||name.length>255||!name.toLowerCase().endsWith('.pptx'))reject('choose a standard .pptx presentation.');
  const zip=await Zip.loadAsync(data,{checkCRC32:false,createFolders:false}),warnings=['Text preview only. Layout, images, charts, diagrams, media, animation and inherited master/layout text are not imported.'];
  const get=async path=>{const f=zip.file(path),expected=parts.get(path);if(!f||!expected)reject('missing required part: '+path);return readXmlPart(f,expected);};
  const types=parse(await get('[Content_Types].xml'),'http://schemas.openxmlformats.org/package/2006/content-types','Types');
  const overrides=nodes(types,types.documentElement.namespaceURI,'Override'),seenParts=new Set();
  for(const t of [...overrides,...nodes(types,types.documentElement.namespaceURI,'Default')])if((details?/macroEnabled|vba|activeX/i:/macroEnabled|vba|oleObject|activeX/i).test(t.getAttribute('ContentType')))reject('macro-enabled or embedded content is unsupported.');
  for(const t of overrides){const path=t.getAttribute('PartName');if(seenParts.has(path))reject('duplicate content type declarations.');seenParts.add(path);if((details?/macroEnabled|vba|activeX/i:/macroEnabled|vba|oleObject|activeX/i).test(t.getAttribute('ContentType')))reject('macro-enabled or embedded content is unsupported.');}
  if(!overrides.some(n=>n.getAttribute('PartName')==='/ppt/presentation.xml'&&n.getAttribute('ContentType')==='application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml'))reject('not a supported standard PPTX presentation.');
  const relationships=new Map();
  for(const f of Object.values(zip.files).filter(f=>!f.dir&&f.name.endsWith('.rels'))){
    const rels=parse(await get(f.name),PKG,'Relationships'),map=new Map(),part=f.name==='_rels/.rels'?'':f.name.replace(/(^|\/)\_rels\//,'$1').replace(/\.rels$/,'');
    for(const r of direct(rels.documentElement,PKG,'Relationship')){
      const id=r.getAttribute('Id'),type=r.getAttribute('Type'),target=r.getAttribute('Target'),mode=r.getAttribute('TargetMode');
      if(!id||!type||!target||map.has(id)||mode&&!['Internal','External'].includes(mode))reject('missing or duplicate relationship identity.');
      const embeddedKind=/oleObject/i.test(type)||type.endsWith('/package')||type.endsWith('/control');
      if(/vba|activeX/i.test(type)||(embeddedKind&&!details))reject('active or embedded-object relationships are unsupported.');
      // An embedded workbook or object (a chart's data, an equation) is not followed or read: the slide's own text is all that is wanted.
      if(embeddedKind){map.set(id,{type,external:true});continue;}
      if(mode==='External'){warnings.push(`External relationship ignored in ${part||'package root'}.`);map.set(id,{type,external:true});continue;}
      const path=resolve(part,target);if(!zip.file(path))reject('missing relationship target: '+path);
      map.set(id,{type,path});
    }relationships.set(part,map);
  }
  const root=Array.from(relationships.get('')?.values()||[]).filter(r=>r.type===R+'/officeDocument');if(root.length!==1||root[0].path!=='ppt/presentation.xml')reject('missing or ambiguous presentation relationship.');
  const presentation=parse(await get('ppt/presentation.xml'),P,'presentation'),lists=direct(presentation.documentElement,P,'sldIdLst');
  if(lists.length!==1)reject('missing or duplicate slide list.');const ids=direct(lists[0],P,'sldId');if(!ids.length||ids.length>500)reject('choose a deck with 1 to 500 slides.');
  const usedIds=new Set(),usedTargets=new Set(),usedRels=new Set(),slides=[],deckRels=relationships.get('ppt/presentation.xml');
  for(const [index,idNode] of ids.entries()){
    const id=idNode.getAttribute('id'),rid=idNode.getAttributeNS(R,'id'),rel=deckRels?.get(rid);
    if(!/^\d+$/.test(id)||Number(id)<256||Number(id)>2147483647||usedIds.has(Number(id))||usedRels.has(rid)||!rel||rel.type!==R+'/slide'||rel.external||usedTargets.has(rel.path))reject('missing, duplicate or unsupported slide identity.');
    usedIds.add(Number(id));usedRels.add(rid);usedTargets.add(rel.path);
    const doc=parse(await get(rel.path),P,'sld'),slideWarnings=[];
    const titles=nodes(doc,P,'sp').filter(sp=>nodes(sp,P,'ph').some(ph=>['title','ctrTitle'].includes(ph.getAttribute('type'))));
    let title=boundedText(titles.map(textOf).filter(Boolean).join('\n'));const text=boundedText(textOf(doc.documentElement));
    // Many decks (those exported from other tools especially) have no title placeholder at all. The slide chooser then names a slide
    // by its largest, highest line of text, and says it guessed.
    let titleGuessed=false;
    if(!title&&details){title=guessTitle(doc);titleGuessed=!!title;}
    if(!title)slideWarnings.push('No explicit slide title found.');if(!text)slideWarnings.push('No editable slide text found.');if(doc.documentElement.getAttribute('show')==='0')slideWarnings.push('Hidden slide in source.');
    if(nodes(doc,P,'pic').length||nodes(doc,A,'graphic').length||nodes(doc,P,'graphicFrame').length)slideWarnings.push('Visual content is present but not imported.');
    const noteRels=Array.from(relationships.get(rel.path)?.values()||[]).filter(r=>r.type===R+'/notesSlide');if(noteRels.length>1||noteRels.some(r=>r.external))reject('ambiguous or external speaker notes.');
    let notes='';if(noteRels.length){const noteDoc=parse(await get(noteRels[0].path),P,'notes');notes=boundedText(nodes(noteDoc,P,'sp').filter(sp=>!nodes(sp,P,'ph').some(ph=>['sldNum','dt','hdr','ftr','sldImg'].includes(ph.getAttribute('type')))).map(textOf).filter(Boolean).join('\n'));}
    if(!notes)slideWarnings.push('No speaker notes text found.');
    // Each effect that starts on a click is one click for the presenter; effects that go with or after it share that click.
    const clicks=nodes(doc,P,'cTn').filter(n=>n.getAttribute('nodeType')==='clickEffect').length,hidden=doc.documentElement.getAttribute('show')==='0';
    // The extra facts are only added when asked for, so the shape the older preview and review code compares does not change.
    slides.push({id:`slide:${Number(id)}`,sourceId:Number(id),sourcePart:rel.path,order:index+1,title,text,notes,warnings:slideWarnings,...(details?{clicks,hidden,titleGuessed}:{})});
  }
  const hash=Array.from(new Uint8Array(await crypto.subtle.digest('SHA-256',data)),n=>n.toString(16).padStart(2,'0')).join('');
  return freeze({format:'openstage-pptx-preview',version:1,source:{name,sha256:hash,bytes:data.length},warnings:[...new Set(warnings)],slides});
}
export function pptxDraft(source){return {sourceHash:source.source.sha256,slides:source.slides.map(s=>({id:s.id,title:s.title,text:s.text,notes:s.notes,moduleId:'',lesson:''}))};}
export function editSlideDraft(draft,id,fields,moduleIds){
  const next=structuredClone(draft),slide=next.slides.find(s=>s.id===id);if(!slide)reject('select a slide.');
  for(const key of ['title','text','notes','moduleId','lesson'])if(typeof fields[key]!=='string'||fields[key].length>50000)reject('invalid slide review field.');
  if(fields.moduleId&&!moduleIds.includes(fields.moduleId))reject('the selected course module no longer exists.');
  for(const key of ['title','text','notes','moduleId','lesson'])slide[key]=fields[key];return next;
}
