import {managedTags,reconcile} from './model.mjs';
import {hasFrontMatter} from './frontmatter.mjs';
import {packageXml} from './cue-ownership.mjs';
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const kids=n=>Array.from(n.childNodes||[]).filter(c=>c.nodeType===1),all=(n,tag)=>Array.from(n.getElementsByTagNameNS(W,tag));
const clean=s=>s.replace(/\r/g,'\n').replace(/\u200b/g,'').trim();
const fail=()=>{throw new Error('The front-matter inventory, captured text or main-body placement is incomplete or unsupported. Preserve this guide; no automatic repair is available.');};
// Direct run/paragraph-mark hiding is not proof of readable required content.
// CT_OnOff absent/true/1/on enables hiding; only explicit supported off values
// are accepted here. Unused style-gallery semiHidden is unrelated.
const hidden=node=>['vanish','webHidden'].some(name=>all(node,name).some(n=>!['false','0','off'].includes(n.getAttributeNS(W,'val'))));
function contentText(node){
 const walk=n=>n.nodeType===1?(n.namespaceURI===W&&n.localName==='t'?n.textContent:n.namespaceURI===W&&n.localName==='br'?'\n':n.namespaceURI===W&&n.localName==='tab'?'\t':Array.from(n.childNodes).map(walk).join('')):'';
 const paragraphs=all(node,'p');return paragraphs.length?paragraphs.map(walk).join('\n'):walk(node);
}
export async function inspectFrontMatterPackage(project,zip,{initial=false}={}){
 if(!hasFrontMatter(project))return;
 const {docs,main,contentType}=await packageXml(zip),expected=[...managedTags(project)],records=[],controls=[];
 if(main.length!==1||main[0]!=='word/document.xml'||contentType(main[0])!=='application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml')fail();
 for(const [part,doc] of docs)for(const marker of all(doc,'tag')){
   const tag=marker.getAttributeNS(W,'val');if(!tag.startsWith('osg:'))continue;
   const pr=marker.parentNode,c=pr?.parentNode;if(pr?.namespaceURI!==W||pr.localName!=='sdtPr'||c?.namespaceURI!==W||c.localName!=='sdt'||all(pr,'tag').length!==1||part!==main[0])fail();
   const content=kids(c).filter(n=>n.namespaceURI===W&&n.localName==='sdtContent');if(content.length!==1)fail();
   let anchor=c,parent=c.parentNode;
   while(parent&&!(parent.namespaceURI===W&&parent.localName==='body')){if(parent.namespaceURI!==W||parent.localName!=='p')fail();anchor=parent;parent=parent.parentNode;}
   if(!parent||hidden(anchor))fail();
   if(tag.startsWith('osg:section-heading:')||tag.startsWith('osg:module:')||tag.startsWith('osg:course:')){
    if(all(anchor,'p').length+(anchor.namespaceURI===W&&anchor.localName==='p'?1:0)!==1||clean(contentText(anchor))!==clean(contentText(content[0]))||['drawing','pict','object','fldSimple','instrText','tbl'].some(name=>all(anchor,name).length))fail();
   }
   records.push({tag,text:contentText(content[0])});controls.push(c);
 }
 if(JSON.stringify(records.map(r=>r.tag))!==JSON.stringify(expected)||JSON.stringify(reconcile(project,records))!==JSON.stringify(project))fail();
 if(initial){
  // Initial completion cannot mistake partial insertion plus unrelated content
  // for a clean starter. Later committed manual additions are authoritative.
  for(const doc of docs.values()){
   if(['drawing','pict','object','tbl','fldChar','fldSimple','instrText','altChunk','subDoc','ins','del'].some(name=>all(doc,name).length))fail();
   if(all(doc,'sdt').some(c=>!controls.includes(c)))fail();
   for(const text of all(doc,'t')){let n=text;while(n&&!controls.includes(n))n=n.parentNode;if(!n&&text.textContent.trim())fail();}
  }
 }
 return {profile:'wrench-starter',version:1,fields:controls.length,nativeVerified:false};
}
function blankBody(doc){
 const bodies=all(doc,'body');if(bodies.length!==1)fail();const children=kids(bodies[0]);
 if(children.some(n=>n.namespaceURI!==W||!['p','sectPr'].includes(n.localName))||children.filter(n=>n.localName==='p').length>1||children.filter(n=>n.localName==='sectPr').length>1)fail();
 for(const p of children.filter(n=>n.localName==='p'))for(const n of kids(p)){
  if(n.namespaceURI!==W||!['pPr','r','bookmarkStart','bookmarkEnd'].includes(n.localName))fail();
  if(n.localName==='r'&&kids(n).some(c=>c.namespaceURI!==W||!['rPr','t'].includes(c.localName)))fail();
 }
 if(all(bodies[0],'t').some(n=>n.textContent.trim())||all(bodies[0],'sectPr').length>1||['sdt','drawing','pict','object','tbl','fldChar','fldSimple','instrText','altChunk','subDoc','footnoteReference','endnoteReference','commentReference','br','tab','ins','del'].some(name=>all(bodies[0],name).length))fail();
}
export function assertBlankBodyOoxml(text){
 if(typeof text!=='string'||/<!DOCTYPE|<!ENTITY/i.test(text))fail();const errors=[];
 const doc=new DOMParser({errorHandler:{warning:e=>errors.push(e),error:e=>errors.push(e),fatalError:e=>errors.push(e)}}).parseFromString(text,'application/xml');
 if(errors.length||!doc.documentElement)fail();blankBody(doc);
}
export async function inspectBlankPackage(zip,expectedPending){
 const {docs,main,contentType}=await packageXml(zip);
 if(main.length!==1||main[0]!=='word/document.xml'||contentType(main[0])!=='application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml')fail();
 const doc=docs.get(main[0]);blankBody(doc);
 const WE='http://schemas.microsoft.com/office/webextensions/webextension/2010/11',projects=[],pending=[];
 for(const xml of docs.values())for(const p of Array.from(xml.getElementsByTagNameNS(WE,'property'))){if(p.getAttribute('name')==='openstage.guide.project.v1')projects.push(JSON.parse(p.getAttribute('value')));if(p.getAttribute('name')==='openstage.guide.pending.v1')pending.push(JSON.parse(p.getAttribute('value')));}
 if(projects.length>1||projects.some(p=>p!==null)||pending.length>1)fail();
 const stable=x=>Array.isArray(x)?x.map(stable):x&&typeof x==='object'?Object.fromEntries(Object.keys(x).sort().map(k=>[k,stable(x[k])])):x;
 if(expectedPending===undefined?pending.some(p=>p!==null):pending.length!==1||JSON.stringify(stable(pending[0]))!==JSON.stringify(stable(expectedPending)))fail();
 for(const [part,xml] of docs){
  if(part!==main[0]&&xml.documentElement.namespaceURI===W&&['hdr','ftr','comments','footnotes','endnotes','glossaryDocument'].includes(xml.documentElement.localName))fail();
  if(all(xml,'t').some(n=>n.textContent.trim())||['sdt','drawing','pict','object','tbl','fldChar','fldSimple','instrText','altChunk','subDoc','footnoteReference','endnoteReference','commentReference','br','tab','ins','del'].some(name=>all(xml,name).length))fail();
  if(part===main[0]&&all(xml,'sectPr').length>1)fail();
 }
 if(Object.values(zip.files).some(f=>!f.dir&&/^(?:word\/(?:media|embeddings|activeX)\/|customXml\/)|vbaProject/i.test(f.name)))fail();
 return {blank:true,nativeVerified:false};
}
