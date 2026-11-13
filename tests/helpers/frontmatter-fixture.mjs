import JSZip from 'jszip';
import {installParser} from './checkpoint-fixture.mjs';
import {KEY,PENDING} from '../../src/document-store.mjs';
import {frontMatterEvidence} from '../../src/frontmatter.mjs';
import {xml} from '../../src/worksheet.mjs';
export const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
export const paragraph=text=>`<w:p><w:r><w:t>${xml(text).replace(/\n/g,'</w:t><w:br/><w:t>')}</w:t></w:r></w:p>`;
export const control=r=>`<w:sdt><w:sdtPr><w:tag w:val="${r.tag}"/></w:sdtPr><w:sdtContent>${paragraph(r.text||'\u200b')}</w:sdtContent></w:sdt>`;
export function committed(project){return {schema:5,txId:crypto.randomUUID(),phase:'committed',kind:'metadata',placement:true,frontMatter:frontMatterEvidence(project),cueChecks:project.modules.flatMap(m=>m.activities.map(a=>a.id)),project:structuredClone(project)};}
// Synthetic compressed file API only. Never connects to Word or user files.
export function frontHost(h,{receipt=true}={}){
 installParser();globalThis.JSZip=JSZip;h.fileReads=0;h.fileCloses=0;
 if(receipt&&h.durable[KEY]){h.durable[PENDING]=committed(h.durable[KEY]);h.memory=structuredClone(h.durable);}
 h.bodyOoxml=()=>`<w:document xmlns:w="${W}"><w:body>${(h.physicalOrder?h.physicalOrder.map(tag=>h.records.find(r=>r.tag===tag)):h.records).map(control).join('')}${h.extraBody||'<w:p/>'}<w:sectPr/></w:body></w:document>`;
 h.package=async()=>{
  const zip=new JSZip();zip.file('[Content_Types].xml','<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  zip.file('_rels/.rels','<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="main" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  zip.file('word/document.xml',h.bodyOoxml());
  zip.file('word/webextensions/webextension1.xml',`<we:webextension xmlns:we="http://schemas.microsoft.com/office/webextensions/webextension/2010/11"><we:properties>${Object.entries(h.memory).map(([name,value])=>`<we:property name="${xml(name)}" value="${xml(JSON.stringify(value))}"/>`).join('')}</we:properties></we:webextension>`);
  zip.file('word/styles.xml','<w:styles xmlns:w="'+W+'"><w:style w:styleId="CustomPrivateStyle"/></w:styles>');
  await h.packageTransform?.(zip,h);return zip;
 };
 h.Office.context.document.getFileAsync=(_type,_options,callback)=>{
  const n=++h.fileReads;Promise.resolve().then(()=>h.hooks.beforeFile?.(n,h)).then(()=>h.package()).then(zip=>zip.generateAsync({type:'uint8array'})).then(bytes=>callback({status:'succeeded',value:{size:bytes.length,sliceCount:Math.ceil(bytes.length/65536),getSliceAsync(i,cb){cb({status:'succeeded',value:{index:i,data:Array.from(bytes.slice(i*65536,(i+1)*65536))}});},closeAsync(cb){h.fileCloses++;cb({status:'succeeded'});}}}),e=>callback({status:'failed',error:{message:e.message}}));
 };
 return h;
}
