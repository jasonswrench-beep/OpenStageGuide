import {CUE_ASSETS} from './cue-assets.mjs';
import {managedTags} from './model.mjs';
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main',WP='http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',A='http://schemas.openxmlformats.org/drawingml/2006/main',R='http://schemas.openxmlformats.org/officeDocument/2006/relationships',REL='http://schemas.openxmlformats.org/package/2006/relationships';
export const OWNED_CUE_PREFIX='osg-cue-image:';
const fail=()=>{throw new Error('Owned instructional image proof is missing, ambiguous or changed. Preserve this master; image editing is unavailable.');};
const elements=(n,ns,tag)=>Array.from(n.getElementsByTagNameNS(ns,tag));
const kids=n=>Array.from(n.childNodes||[]).filter(n=>n.nodeType===1);
const clean=s=>(s||'').replace(/[\r\n\u200b]/g,'').trim();
const hash=async bytes=>Array.from(new Uint8Array(await crypto.subtle.digest('SHA-256',bytes)),v=>v.toString(16).padStart(2,'0')).join('');
function xml(text){
  if(/<!DOCTYPE|<!ENTITY/i.test(text))fail();const errors=[];
  const doc=new DOMParser({errorHandler:{warning:e=>errors.push(e),error:e=>errors.push(e),fatalError:e=>errors.push(e)}}).parseFromString(text,'application/xml');
  if(errors.length||!doc.documentElement)fail();return doc;
}
function tag(control){const pr=kids(control).filter(n=>n.namespaceURI===W&&n.localName==='sdtPr');if(pr.length!==1)return '';const tags=kids(pr[0]).filter(n=>n.namespaceURI===W&&n.localName==='tag');return tags.length===1?tags[0].getAttributeNS(W,'val'):'';}
export function cueOwnerTag(owner){if(owner?.version!==1||!/^[a-f0-9]{32}$/.test(owner.nonce||''))fail();return OWNED_CUE_PREFIX+'1:'+owner.nonce;}
export function createCueOwnership(activityId,icon){
  const asset=CUE_ASSETS[icon];if(!asset||typeof activityId!=='string'||!/^[a-f0-9]{8}-(?:[a-f0-9]{4}-){3}[a-f0-9]{12}$/.test(activityId))fail();
  return {version:1,nonce:crypto.randomUUID().replaceAll('-',''),activityId,assetHash:asset.sha256};
}
function owner(item){
  const o=item.cueImage,keys=['version','nonce','activityId','assetHash'];
  if(!o||typeof o!=='object'||Array.isArray(o)||Object.keys(o).length!==keys.length||keys.some(k=>!Object.hasOwn(o,k))||o.activityId!==item.id||o.assetHash!==CUE_ASSETS[item.icon]?.sha256)fail();
  cueOwnerTag(o);return o;
}
function directContainer(anchor){let parent=anchor.parentNode;while(parent&&!(parent.namespaceURI===W&&['body','tc'].includes(parent.localName))){if(parent.namespaceURI===W&&['sdt','tbl'].includes(parent.localName))fail();anchor=parent;parent=parent.parentNode;}if(!parent)fail();return {anchor,parent};}
function nextElement(node){for(let n=node.nextSibling;n;n=n.nextSibling)if(n.nodeType===1)return n;return null;}
function plainText(node){return elements(node,W,'t').map(n=>n.textContent).join('');}
function supportedParagraph(node,ownedControl=null){
  const forbidden=new Set(['pict','object','altChunk','subDoc','txbxContent','instrText','fldSimple','del','ins','moveFrom','moveTo','vanish','webHidden']);
  for(const element of [node,...elements(node,'*','*')])if(element.namespaceURI==='urn:schemas-microsoft-com:vml'||element.namespaceURI==='urn:schemas-microsoft-com:office:office'&&element.localName==='OLEObject'||element.namespaceURI===W&&forbidden.has(element.localName))fail();
  if(node.namespaceURI!==W||node.localName!=='p')return;
  const content=kids(node).filter(n=>!(n.namespaceURI===W&&n.localName==='pPr'));
  if(ownedControl){if(content.length!==1||content[0]!==ownedControl)fail();return;}
  for(const run of content){if(run.namespaceURI!==W||run.localName!=='r')fail();for(const child of kids(run))if(child.namespaceURI!==W||!['rPr','t'].includes(child.localName))fail();}
}
export async function packageXml(zip){
  const CT='http://schemas.openxmlformats.org/package/2006/content-types',files=Object.values(zip.files).filter(f=>!f.dir),byName=new Map(files.map(f=>[f.name,f])),defaults=new Map(),overrides=new Map(),docs=new Map();
  const ctFile=byName.get('[Content_Types].xml');if(!ctFile)fail();const types=xml(await ctFile.async('string'));if(types.documentElement.namespaceURI!==CT||types.documentElement.localName!=='Types')fail();
  for(const node of kids(types.documentElement)){
    if(node.namespaceURI!==CT)fail();
    if(node.localName==='Default'){const ext=node.getAttribute('Extension').toLowerCase();if(!ext||defaults.has(ext))fail();defaults.set(ext,node.getAttribute('ContentType').toLowerCase());}
    else if(node.localName==='Override'){const name=node.getAttribute('PartName');if(!name.startsWith('/')||overrides.has(name))fail();overrides.set(name,node.getAttribute('ContentType').toLowerCase());}
    else fail();
  }
  const contentType=name=>overrides.get('/'+name)||defaults.get(name.split('.').at(-1).toLowerCase())||'';
  const xmlParts=new Set(files.filter(f=>/\.(?:xml|rels)$/i.test(f.name)||/(?:\+xml|\/xml)$/.test(contentType(f.name))).map(f=>f.name));
  const relationships=files.filter(f=>/\.rels$/i.test(f.name)||contentType(f.name)==='application/vnd.openxmlformats-package.relationships+xml');
  const read=async name=>{if(!byName.has(name))fail();if(!docs.has(name))docs.set(name,xml(await byName.get(name).async('string')));return docs.get(name);};
  function sourceFor(name){if(name.toLowerCase()==='_rels/.rels')return '';const match=name.match(/^(.*\/)?_rels\/([^/]+)\.rels$/i);if(!match)fail();return (match[1]||'')+match[2];}
  function resolve(source,target){
    if(!target||/[\\:?#\x00-\x1f]/.test(target))fail();let decoded;try{decoded=decodeURIComponent(target);}catch{fail();}if(/[\\:?#\x00-\x1f]/.test(decoded))fail();
    const parts=decoded.startsWith('/')?[]:source.split('/').slice(0,-1);for(const part of decoded.replace(/^\//,'').split('/')){if(part==='..'){if(!parts.length)fail();parts.pop();}else if(part==='.'||part==='')fail();else parts.push(part);}return parts.join('/');
  }
  const storyNames=new Set(['officeDocument','header','footer','footnotes','endnotes','comments','glossaryDocument','subDocument']),main=[];
  for(const file of relationships){
    const doc=await read(file.name);if(doc.documentElement.namespaceURI!==REL||doc.documentElement.localName!=='Relationships')fail();const source=sourceFor(file.name),ids=new Set();
    for(const relationship of kids(doc.documentElement)){
      if(relationship.namespaceURI!==REL||relationship.localName!=='Relationship')fail();const id=relationship.getAttribute('Id'),type=relationship.getAttribute('Type'),mode=relationship.getAttribute('TargetMode');if(!id||ids.has(id))fail();ids.add(id);
      const namespace=type.slice(0,type.lastIndexOf('/')),role=type.slice(type.lastIndexOf('/')+1);
      if(![R,'http://purl.oclc.org/ooxml/officeDocument/relationships'].includes(namespace)||!storyNames.has(role))continue;
      if(mode&&mode!=='Internal')fail();const target=resolve(source,relationship.getAttribute('Target'));if(!byName.has(target))fail();xmlParts.add(target);if(!source&&role==='officeDocument')main.push(target);
    }
  }
  for(const part of xmlParts)await read(part);
  return {docs,main,contentType};
}
// Read-only package proof; no automatic adoption, metadata/body writes or live
// Word identity assertion. New insertion/journal integration is a later gate.
export async function inspectOwnedCuePackage(project,zip){
  const inputProject=project;project=structuredClone(project);
  const discovered=await packageXml(zip),docs=discovered.docs,markers=[];
  for(const [part,doc] of docs){
    for(const marker of elements(doc,W,'tag')){const value=marker.getAttributeNS(W,'val');if(!value.startsWith(OWNED_CUE_PREFIX))continue;const pr=marker.parentNode,control=pr?.parentNode;if(pr?.namespaceURI!==W||pr.localName!=='sdtPr'||control?.namespaceURI!==W||control.localName!=='sdt'||tag(control)!==value)fail();markers.push({part,control,tag:value});}
  }
  // Legacy activity metadata is not adopted merely because it happens to use
  // the same property name. A new course-level format marker is explicit.
  if(project.cueOwnership===undefined){if(markers.length)fail();return {version:1,owned:[],legacy:true,nativeVerified:false};}
  if(!project.cueOwnership||project.cueOwnership.version!==1||Object.keys(project.cueOwnership).length!==1)fail();
  if(discovered.main.length!==1||discovered.main[0]!=='word/document.xml'||discovered.contentType('word/document.xml')!=='application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml')fail();
  const items=project.modules.flatMap(m=>m.activities).filter(a=>Object.hasOwn(a,'cueImage'));
  const tags=new Set();for(const item of items){const value=cueOwnerTag(owner(item));if(tags.has(value))fail();tags.add(value);}
  if(markers.length!==items.length||markers.some(m=>m.part!=='word/document.xml'||!tags.has(m.tag)))fail();
  const doc=docs.get('word/document.xml');if(!doc)fail();const results=[];
  const relFile=zip.file('word/_rels/document.xml.rels'),rels=relFile?xml(await relFile.async('string')):null;
  for(const item of items){
    const binding=owner(item),value=cueOwnerTag(binding),found=markers.filter(m=>m.tag===value);if(found.length!==1)fail();const control=found[0].control;
    const content=kids(control).filter(n=>n.namespaceURI===W&&n.localName==='sdtContent');if(content.length!==1)fail();
    let visible=kids(content[0]);if(visible.length!==1)fail();
    if(visible[0].namespaceURI===W&&visible[0].localName==='p')visible=kids(visible[0]).filter(n=>!(n.namespaceURI===W&&n.localName==='pPr'));
    if(visible.length!==1||visible[0].namespaceURI!==W||visible[0].localName!=='r')fail();
    const runContent=kids(visible[0]).filter(n=>!(n.namespaceURI===W&&n.localName==='rPr'));if(runContent.length!==1||runContent[0].namespaceURI!==W||runContent[0].localName!=='drawing')fail();
    if(elements(control,W,'sdt').length||elements(control,W,'tbl').length||elements(control,WP,'anchor').length||elements(control,W,'drawing').length!==1||clean(plainText(control)))fail();
    const inlines=elements(control,WP,'inline'),blips=elements(control,A,'blip'),properties=elements(control,WP,'docPr');if(inlines.length!==1||blips.length!==1||properties.length!==1)fail();
    const drawingChildren=kids(runContent[0]),graphicData=elements(control,A,'graphicData'),PIC='http://schemas.openxmlformats.org/drawingml/2006/picture';
    if(drawingChildren.length!==1||drawingChildren[0]!==inlines[0]||graphicData.length!==1||graphicData[0].getAttribute('uri')!==PIC||kids(graphicData[0]).length!==1||kids(graphicData[0])[0].namespaceURI!==PIC||kids(graphicData[0])[0].localName!=='pic')fail();
    const title=properties[0].getAttribute('title'),description=properties[0].getAttribute('descr');
    if(title!==`OpenStage Guide ${item.icon} cue`||description!==`${item.icon} icon`)fail();
    const extents=kids(inlines[0]).filter(n=>n.namespaceURI===WP&&n.localName==='extent');if(extents.length!==1)fail();
    const cx=Number(extents[0].getAttribute('cx')),cy=Number(extents[0].getAttribute('cy'));
    if(!Number.isSafeInteger(cx)||!Number.isSafeInteger(cy)||cx<=0||cy<=0||cx>9144000||cy>9144000||elements(control,A,'srcRect').length)fail();
    for(const transform of elements(control,A,'xfrm')){if(['rot','flipH','flipV'].some(k=>transform.hasAttribute(k)&&!['0','false'].includes(transform.getAttribute(k))))fail();}
    const embed=blips[0].getAttributeNS(R,'embed');if(!embed||blips[0].hasAttributeNS(R,'link')||!rels)fail();
    const references=elements(rels,REL,'Relationship').filter(r=>r.getAttribute('Id')===embed);if(references.length!==1)fail();const rel=references[0],target=rel.getAttribute('Target');
    if(rel.getAttribute('Type')!==R+'/image'||rel.hasAttribute('TargetMode')||!/^media\/[\w .-]+\.png$/i.test(target))fail();
    const picture=zip.file('word/'+target);if(!picture||picture._data?.uncompressedSize>2*1024*1024)fail();const image=await picture.async('uint8array');if(await hash(image)!==binding.assetHash)fail();
    const pictureLocation=directContainer(control),titles=elements(doc,W,'sdt').filter(c=>tag(c)===`osg:activity:${item.id}`);if(titles.length!==1)fail();
    const titleLocation=directContainer(titles[0]),cue=nextElement(pictureLocation.anchor);
    if(pictureLocation.parent!==titleLocation.parent||!cue||cue.namespaceURI!==W||cue.localName!=='p'||nextElement(cue)!==titleLocation.anchor||clean(plainText(cue))!==`${item.icon} · ${item.minutes} min`||elements(cue,W,'sdt').length||elements(cue,W,'drawing').length||elements(cue,W,'tbl').length)fail();
    if(elements(pictureLocation.anchor,W,'drawing').length!==1||clean(plainText(pictureLocation.anchor))||elements(pictureLocation.anchor,W,'sdt').some(c=>c!==control))fail();
    supportedParagraph(pictureLocation.anchor,pictureLocation.anchor===control?null:control);supportedParagraph(cue);
    results.push({activityId:item.id,tag:value,assetHash:binding.assetHash,widthPoints:cx/12700,heightPoints:cy/12700,container:pictureLocation.parent.localName==='tc'?'cell':'body'});
  }
  const actual=elements(doc,W,'sdt').map(tag).filter(value=>value.startsWith('osg:'));
  if(JSON.stringify(actual)!==JSON.stringify([...managedTags(project)]))fail();
  if(JSON.stringify(inputProject)!==JSON.stringify(project))fail();
  return {version:1,owned:results,legacy:false,nativeVerified:false};
}
